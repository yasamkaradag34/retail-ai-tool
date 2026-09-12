"""Run a ClickUp task through isolated Codex Executor and Reviewer stages.

The orchestrator owns ClickUp and GitHub credentials. Agent subprocesses receive
a filtered environment and run with sandboxed command access and no shell
network access. Approved work becomes a pull request; this module never merges.
"""
from __future__ import annotations

import argparse
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime, timezone
import fcntl
import json
import os
from pathlib import Path
import re
import signal
import shutil
import subprocess
import sys
import time
import traceback
from typing import Any, Callable, Iterable
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen


REPO_ROOT = Path(__file__).resolve().parents[2]
PACKAGE_ROOT = Path(__file__).resolve().parent
CLICKUP_API = "https://api.clickup.com/api/v2"
SAFE_PROCESS_ENV = {
    "PATH", "HOME", "USER", "LOGNAME", "SHELL", "TMPDIR", "TMP", "TEMP",
    "LANG", "LC_ALL", "LC_CTYPE", "TERM", "COLORTERM", "NO_COLOR",
    "SSL_CERT_FILE", "SSL_CERT_DIR", "CODEX_HOME",
}


class PipelineError(RuntimeError):
    """A safe, user-readable pipeline error."""


def env_bool(name: str, default: bool) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


def env_int(name: str, default: int, minimum: int = 1) -> int:
    try:
        return max(minimum, int(os.getenv(name, str(default))))
    except ValueError:
        return default


def safe_slug(value: str, maximum: int = 42) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", value.lower()).strip("-")
    return (slug or "task")[:maximum].rstrip("-")


def compact(value: str, maximum: int) -> str:
    return re.sub(r"\s+", " ", value or "").strip()[:maximum]


def clickup_path_id(value: str, label: str, digits_only: bool = False) -> str:
    pattern = r"[0-9]+" if digits_only else r"[A-Za-z0-9_-]+"
    if not re.fullmatch(pattern, str(value or "")):
        raise PipelineError(f"Invalid ClickUp {label}.")
    return quote(str(value), safe="")


@dataclass(frozen=True)
class Settings:
    repo: Path
    runs_dir: Path
    worktrees_dir: Path
    clickup_token: str
    clickup_list_id: str
    queue_status: str
    working_status: str
    review_status: str
    blocked_status: str
    base_branch: str
    test_command: str
    codex_bin: str
    gh_bin: str
    executor_model: str
    reviewer_model: str
    codex_timeout: int
    test_timeout: int
    push_approved_branch: bool
    keep_success_worktree: bool
    isolated_runner: bool

    @classmethod
    def from_env(cls, repo: Path | None = None) -> "Settings":
        root = (repo or REPO_ROOT).resolve()
        return cls(
            repo=root,
            runs_dir=root / ".agent-runs",
            worktrees_dir=root / ".agent-worktrees",
            clickup_token=os.getenv("CLICKUP_API_TOKEN", "").strip(),
            clickup_list_id=os.getenv("CLICKUP_LIST_ID", "").strip(),
            queue_status=os.getenv("CLICKUP_QUEUE_STATUS", "to do").strip(),
            working_status=os.getenv("CLICKUP_WORKING_STATUS", "in progress").strip(),
            review_status=os.getenv("CLICKUP_REVIEW_STATUS", "review").strip(),
            blocked_status=os.getenv("CLICKUP_BLOCKED_STATUS", "blocked").strip(),
            base_branch=os.getenv("AGENT_BASE_BRANCH", "main").strip(),
            test_command=os.getenv(
                "AGENT_TEST_COMMAND",
                "PYTHONDONTWRITEBYTECODE=1 python3 -m unittest discover -s tests -p 'test_*.py'",
            ).strip(),
            codex_bin=os.getenv("CODEX_BIN", shutil.which("codex") or "codex"),
            gh_bin=os.getenv("GH_BIN", shutil.which("gh") or "gh"),
            executor_model=os.getenv("AGENT_EXECUTOR_MODEL", "").strip(),
            reviewer_model=os.getenv("AGENT_REVIEWER_MODEL", "").strip(),
            codex_timeout=env_int("AGENT_CODEX_TIMEOUT_SECONDS", 3600, 60),
            test_timeout=env_int("AGENT_TEST_TIMEOUT_SECONDS", 900, 30),
            push_approved_branch=env_bool("AGENT_PUSH_APPROVED_BRANCH", True),
            keep_success_worktree=env_bool("AGENT_KEEP_SUCCESS_WORKTREE", False),
            isolated_runner=env_bool("AGENT_ISOLATED_RUNNER", False),
        )

    def require_clickup(self) -> None:
        missing = [name for name, value in (("CLICKUP_API_TOKEN", self.clickup_token), ("CLICKUP_LIST_ID", self.clickup_list_id)) if not value]
        if missing:
            raise PipelineError("Missing required setting(s): " + ", ".join(missing))


@dataclass(frozen=True)
class Task:
    id: str
    name: str
    description: str
    url: str
    status: str
    priority: int | None = None

    def is_queued(self, status: str) -> bool:
        """Only ClickUp Urgent tasks in the configured queue may be claimed."""
        return self.status.casefold() == status.casefold() and self.priority == 1

    @classmethod
    def from_payload(cls, payload: dict[str, Any]) -> "Task":
        task_id = compact(str(payload.get("id") or ""), 100)
        name = compact(str(payload.get("name") or ""), 300)
        if not task_id or not name:
            raise PipelineError("ClickUp returned a task without an id or name.")
        description = str(
            payload.get("markdown_description")
            or payload.get("description")
            or payload.get("text_content")
            or ""
        ).strip()
        status = payload.get("status") or {}
        if isinstance(status, dict):
            status = status.get("status") or ""
        priority = payload.get("priority")
        if isinstance(priority, dict):
            priority = priority.get("id")
        priority = int(priority) if str(priority) in {"1", "2", "3", "4"} else None
        return cls(task_id, name, description[:50000], compact(str(payload.get("url") or ""), 1000), compact(str(status), 100), priority)


@dataclass(frozen=True)
class CommandResult:
    returncode: int
    stdout: str
    stderr: str
    seconds: float

    @property
    def combined(self) -> str:
        return "\n".join(part for part in (self.stdout.strip(), self.stderr.strip()) if part)


def filtered_agent_environment(source: dict[str, str] | None = None, include_codex_auth: bool = True) -> dict[str, str]:
    source = os.environ if source is None else source
    allowed = SAFE_PROCESS_ENV | ({"CODEX_API_KEY"} if include_codex_auth else set())
    env = {key: value for key, value in source.items() if key in allowed or key.startswith("LC_")}
    env["NO_COLOR"] = "1"
    return env


def redact(text: str, secrets: Iterable[str]) -> str:
    result = text
    for secret in sorted((item for item in secrets if item), key=len, reverse=True):
        result = result.replace(secret, "[REDACTED]")
    return result


def run_command(
    args: list[str],
    cwd: Path,
    timeout: int,
    env: dict[str, str] | None = None,
    input_text: str | None = None,
) -> CommandResult:
    started = time.monotonic()
    process: subprocess.Popen[str] | None = None
    try:
        process = subprocess.Popen(
            args,
            cwd=str(cwd),
            env=env,
            text=True,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            start_new_session=os.name == "posix",
        )
        stdout, stderr = process.communicate(input=input_text, timeout=timeout)
        return CommandResult(process.returncode, stdout, stderr, time.monotonic() - started)
    except subprocess.TimeoutExpired as error:
        stdout = error.stdout.decode() if isinstance(error.stdout, bytes) else (error.stdout or "")
        stderr = error.stderr.decode() if isinstance(error.stderr, bytes) else (error.stderr or "")
        if process is not None:
            _kill_process_group(process)
            tail_stdout, tail_stderr = process.communicate()
            stdout += tail_stdout or ""
            stderr += tail_stderr or ""
        return CommandResult(124, stdout, stderr + f"\nTimed out after {timeout} seconds.", time.monotonic() - started)
    finally:
        if process is not None:
            _kill_process_group(process)


def _kill_process_group(process: subprocess.Popen[str]) -> None:
    """Stop descendants that could mutate a reviewed worktree after a command exits."""
    if os.name != "posix":
        if process.poll() is None:
            process.kill()
        return
    for sig in (signal.SIGTERM, signal.SIGKILL):
        try:
            os.killpg(process.pid, sig)
        except ProcessLookupError:
            return
        except PermissionError:
            return


def run_test_command(settings: Settings, cwd: Path, env: dict[str, str]) -> CommandResult:
    command = settings.test_command
    if not command:
        raise PipelineError("AGENT_TEST_COMMAND must not be empty.")
    args = [
        settings.codex_bin, "sandbox", "-P", ":workspace", "-C", str(cwd),
        "--sandbox-state-disable-network", "/bin/sh", "-lc", command,
    ]
    return run_command(args, cwd, settings.test_timeout, env)


class ClickUpClient:
    def __init__(self, token: str, opener: Callable[..., Any] = urlopen):
        self.token = token
        self.opener = opener

    def _request(self, method: str, path: str, payload: dict[str, Any] | None = None, query: list[tuple[str, str]] | None = None) -> dict[str, Any]:
        url = CLICKUP_API + path
        if query:
            url += "?" + urlencode(query)
        data = json.dumps(payload).encode("utf-8") if payload is not None else None
        request = Request(url, data=data, method=method, headers={"Authorization": self.token, "Content-Type": "application/json", "Accept": "application/json"})
        try:
            with self.opener(request, timeout=25) as response:
                body = response.read().decode("utf-8")
        except HTTPError as error:
            detail = error.read(2000).decode("utf-8", errors="replace")
            raise PipelineError(f"ClickUp API returned HTTP {error.code}: {compact(detail, 500)}") from error
        except URLError as error:
            raise PipelineError("ClickUp API could not be reached.") from error
        if not body:
            return {}
        try:
            result = json.loads(body)
        except json.JSONDecodeError as error:
            raise PipelineError("ClickUp API returned an unreadable response.") from error
        if not isinstance(result, dict):
            raise PipelineError("ClickUp API returned an unexpected response.")
        return result

    def list_tasks(self, list_id: str, status: str) -> list[Task]:
        list_id = clickup_path_id(list_id, "list id", digits_only=True)
        tasks: list[Task] = []
        seen: set[str] = set()
        page = 0
        # Get Tasks has no documented priority filter and returns at most 100
        # tasks per page. An older non-Urgent page must not hide eligible work.
        while True:
            payload = self._request(
                "GET",
                f"/list/{list_id}/task",
                query=[("archived", "false"), ("include_markdown_description", "true"), ("order_by", "created"), ("reverse", "false"), ("statuses[]", status), ("page", str(page))],
            )
            items = payload.get("tasks", [])
            for item in items:
                task = Task.from_payload(item)
                if task.is_queued(status) and task.id not in seen:
                    tasks.append(task)
                    seen.add(task.id)
            if not items or payload.get("last_page") is True or (payload.get("last_page") is not False and len(items) < 100):
                return tasks
            page += 1

    def get_task(self, task_id: str) -> Task:
        task_id = clickup_path_id(task_id, "task id")
        return Task.from_payload(self._request("GET", f"/task/{task_id}", query=[("include_markdown_description", "true")]))

    def set_status(self, task_id: str, status: str) -> None:
        task_id = clickup_path_id(task_id, "task id")
        self._request("PUT", f"/task/{task_id}", {"status": status})

    def comment(self, task_id: str, message: str) -> None:
        task_id = clickup_path_id(task_id, "task id")
        self._request("POST", f"/task/{task_id}/comment", {"comment_text": message[:10000], "notify_all": False})


def executor_prompt(task: Task) -> str:
    role = (PACKAGE_ROOT / "prompts" / "executor.md").read_text()
    return f"""{role}

## ClickUp task

Task ID: {task.id}
Title: {task.name}
URL: {task.url or "not provided"}

<clickup-task-description>
{task.description or "No description was provided. Infer only what is necessary from the title and repository context."}
</clickup-task-description>
"""


def reviewer_prompt(task: Task, tests: CommandResult, test_command: str) -> str:
    role = (PACKAGE_ROOT / "prompts" / "reviewer.md").read_text()
    test_output = tests.combined[-12000:] or "The configured test command produced no output."
    return f"""{role}

## ClickUp task

Task ID: {task.id}
Title: {task.name}
URL: {task.url or "not provided"}

<clickup-task-description>
{task.description or "No description was provided."}
</clickup-task-description>

## Orchestrator test result

Command: {test_command}
Exit code: {tests.returncode}
Duration: {tests.seconds:.2f} seconds

<test-output>
{test_output}
</test-output>
"""


def codex_executor_command(settings: Settings, worktree: Path, result_path: Path) -> list[str]:
    command = [
        settings.codex_bin, "exec", "--ephemeral", "--ignore-user-config", "--strict-config",
        "--sandbox", "workspace-write", "-C", str(worktree),
        "-c", 'approval_policy="never"',
        "-c", "sandbox_workspace_write.network_access=false",
        "-c", 'web_search="disabled"',
        "-c", 'shell_environment_policy.inherit="core"',
        "-c", "shell_environment_policy.ignore_default_excludes=false",
        "--output-schema", str(PACKAGE_ROOT / "schemas" / "executor-result.schema.json"),
        "--output-last-message", str(result_path), "--json",
    ]
    if settings.executor_model:
        command.extend(["--model", settings.executor_model])
    command.append("-")
    return command


def codex_reviewer_command(settings: Settings, worktree: Path, result_path: Path) -> list[str]:
    command = [
        settings.codex_bin, "exec", "--ephemeral", "--ignore-user-config", "--strict-config",
        "--sandbox", "read-only", "-C", str(worktree),
        "-c", 'approval_policy="never"',
        "-c", 'web_search="disabled"',
        "-c", 'shell_environment_policy.inherit="core"',
        "-c", "shell_environment_policy.ignore_default_excludes=false",
        "--output-schema", str(PACKAGE_ROOT / "schemas" / "reviewer-result.schema.json"),
        "--output-last-message", str(result_path), "--json",
    ]
    if settings.reviewer_model:
        command.extend(["--model", settings.reviewer_model])
    command.append("-")
    return command


def parse_result(path: Path, expected: str) -> dict[str, Any]:
    try:
        payload = json.loads(path.read_text())
    except (OSError, json.JSONDecodeError) as error:
        raise PipelineError(f"{expected} did not produce valid structured output.") from error
    if not isinstance(payload, dict):
        raise PipelineError(f"{expected} returned an unexpected result.")
    return payload


def checked(result: CommandResult, label: str) -> CommandResult:
    if result.returncode:
        detail = compact(result.combined[-1500:], 1500)
        raise PipelineError(f"{label} failed with exit code {result.returncode}. {detail}")
    return result


def git(settings: Settings, *args: str, cwd: Path | None = None, timeout: int = 180) -> CommandResult:
    return run_command(["git", *args], cwd or settings.repo, timeout, os.environ.copy())


def prepare_worktree(settings: Settings, task: Task) -> tuple[Path, str]:
    if git(settings, "status", "--porcelain").stdout.strip():
        raise PipelineError("The orchestration checkout must be clean before starting a task.")
    checked(git(settings, "fetch", "origin", settings.base_branch, timeout=300), "Git fetch")
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d%H%M%S")
    task_key = safe_slug(task.id, 20)
    branch = f"agent/clickup-{task_key}-{safe_slug(task.name)}-{stamp}"
    worktree = settings.worktrees_dir / f"{task_key}-{stamp}"
    settings.worktrees_dir.mkdir(parents=True, exist_ok=True)
    checked(git(settings, "worktree", "add", "-b", branch, str(worktree), f"origin/{settings.base_branch}", timeout=300), "Git worktree creation")
    return worktree, branch


def task_run_dir(settings: Settings, task: Task) -> Path:
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    result = settings.runs_dir / f"{stamp}-{safe_slug(task.id, 30)}"
    result.mkdir(parents=True, exist_ok=False)
    return result


def save_log(path: Path, result: CommandResult, secrets: Iterable[str]) -> None:
    payload = f"exit_code={result.returncode}\nduration_seconds={result.seconds:.3f}\n\nSTDOUT\n{result.stdout}\n\nSTDERR\n{result.stderr}"
    path.write_text(redact(payload, secrets))


def review_comment(review: dict[str, Any]) -> str:
    lines = [f"Reviewer verdict: {review.get('verdict', 'request_changes')}", compact(str(review.get("summary") or ""), 1500)]
    findings = review.get("findings") if isinstance(review.get("findings"), list) else []
    for finding in findings[:10]:
        if not isinstance(finding, dict):
            continue
        location = compact(str(finding.get("file") or ""), 300)
        if finding.get("line"):
            location += f":{finding['line']}"
        lines.append(f"- [{finding.get('severity', 'medium')}] {compact(str(finding.get('title') or ''), 200)}" + (f" ({location})" if location else ""))
    return "\n".join(line for line in lines if line)


def create_pull_request(settings: Settings, task: Task, worktree: Path, branch: str, candidate_tree: str, tests: CommandResult, review: dict[str, Any], run_dir: Path, secrets: Iterable[str]) -> tuple[str, str]:
    title = compact(f"[ClickUp {task.id}] {task.name}", 120)
    task_reference = f"[{task.id}]({task.url})" if task.url.startswith("https://app.clickup.com/") else task.id
    body = f"""Implements ClickUp task {task_reference} through the Codex Executor and Reviewer pipeline.

The configured verification command passed with exit code {tests.returncode} in {tests.seconds:.2f} seconds.

Reviewer: {redact(compact(str(review.get("summary") or "Approved."), 2000), secrets)}

This pull request requires manual merge approval.
"""
    body_path = run_dir / "pull-request.md"
    body_path.write_text(body)
    checked(git(settings, "commit", "-m", title, cwd=worktree), "Git commit")
    committed_tree = checked(git(settings, "rev-parse", "HEAD^{tree}", cwd=worktree), "Committed tree verification").stdout.strip()
    if committed_tree != candidate_tree:
        raise PipelineError("The commit tree differs from the reviewed candidate; refusing to push.")
    commit_sha = checked(git(settings, "rev-parse", "HEAD", cwd=worktree), "Git commit id").stdout.strip()
    checked(git(settings, "push", "--set-upstream", "origin", branch, cwd=worktree, timeout=300), "Git push")
    command = [settings.gh_bin, "pr", "create", "--base", settings.base_branch, "--head", branch, "--title", title, "--body-file", str(body_path)]
    result = checked(run_command(command, worktree, 300, os.environ.copy()), "Pull request creation")
    url = next((line.strip() for line in reversed(result.stdout.splitlines()) if line.strip().startswith("http")), "")
    if not url:
        raise PipelineError("GitHub created no readable pull request URL.")
    return url, commit_sha


def safe_clickup_failure(client: ClickUpClient, settings: Settings, task: Task, message: str) -> dict[str, Any]:
    result: dict[str, Any] = {"comment_saved": False, "blocked_status_saved": False, "errors": []}
    actions = (
        ("comment_saved", lambda: client.comment(task.id, "Codex pipeline stopped.\n\n" + compact(message, 8000))),
        ("blocked_status_saved", lambda: client.set_status(task.id, settings.blocked_status)),
    )
    for label, action in actions:
        try:
            action()
        except Exception as error:
            result["errors"].append(f"{type(error).__name__}: {compact(str(error), 300)}")
        else:
            result[label] = True
    if result["errors"]:
        print("Warning: ClickUp failure update was incomplete: " + "; ".join(result["errors"]), file=sys.stderr)
    return result


def stage_candidate(settings: Settings, worktree: Path) -> str:
    checked(git(settings, "add", "-A", cwd=worktree), "Git candidate staging")
    tree = checked(git(settings, "write-tree", cwd=worktree), "Git candidate hash").stdout.strip()
    if not re.fullmatch(r"[0-9a-f]{40,64}", tree):
        raise PipelineError("Git did not produce a valid candidate tree hash.")
    return tree


def assert_candidate_unchanged(settings: Settings, worktree: Path, expected_tree: str) -> None:
    unstaged = git(settings, "diff", "--quiet", "--no-ext-diff", cwd=worktree)
    if unstaged.returncode not in (0, 1):
        checked(unstaged, "Git worktree verification")
    untracked = checked(git(settings, "ls-files", "--others", "--exclude-standard", cwd=worktree), "Git untracked-file verification")
    current_tree = checked(git(settings, "write-tree", cwd=worktree), "Git candidate recheck").stdout.strip()
    if unstaged.returncode or untracked.stdout.strip() or current_tree != expected_tree:
        raise PipelineError("The candidate changed after staging; refusing to commit unreviewed bytes.")


def reject_control_plane_changes(settings: Settings, worktree: Path) -> None:
    changed = checked(git(settings, "diff", "--cached", "--no-renames", "--name-only", "--diff-filter=ACMRD", cwd=worktree), "Git control-plane check")
    forbidden = []
    for raw in changed.stdout.splitlines():
        path = raw.strip().replace("\\", "/")
        lowered = path.lower()
        basename = lowered.rsplit("/", 1)[-1]
        if basename in {"agents.md", "agents.override.md"} or "/.codex/" in "/" + lowered or lowered.startswith(".codex/"):
            forbidden.append(path)
    if forbidden:
        raise PipelineError("Executor attempted to change reviewer control files: " + ", ".join(forbidden[:10]))


def secret_values(settings: Settings) -> tuple[str, ...]:
    values = {settings.clickup_token}
    for key, value in os.environ.items():
        if re.search(r"(?:KEY|TOKEN|SECRET|PASSWORD|CREDENTIAL)", key, re.I) and len(value) >= 8:
            values.add(value)
    codex_home = Path(os.getenv("CODEX_HOME", Path.home() / ".codex"))
    auth_file = codex_home / "auth.json"
    try:
        payload = json.loads(auth_file.read_text())
    except (OSError, json.JSONDecodeError):
        payload = None
    stack = [("", payload)]
    while stack:
        key, item = stack.pop()
        if isinstance(item, dict):
            stack.extend((str(child_key), child) for child_key, child in item.items())
        elif isinstance(item, list):
            stack.extend((key, child) for child in item)
        elif isinstance(item, str) and len(item) >= 16 and re.search(r"(?:KEY|TOKEN|SECRET|PASSWORD|CREDENTIAL)", key, re.I):
            values.add(item)
    return tuple(value for value in values if value)


def reject_staged_secrets(settings: Settings, worktree: Path, secrets: Iterable[str]) -> None:
    files = checked(git(settings, "diff", "--cached", "--name-only", "--diff-filter=ACMR", cwd=worktree), "Git secret-scan file list")
    encoded = [(secret, secret.encode()) for secret in secrets if len(secret) >= 8]
    for raw in files.stdout.splitlines():
        path = raw.strip()
        try:
            data = (worktree / path).read_bytes()
        except OSError as error:
            raise PipelineError(f"Could not inspect staged file for secrets: {path}") from error
        if any(value in data for _, value in encoded):
            raise PipelineError(f"A staged file contains credential material: {path}")


def save_outcome(run_dir: Path, payload: dict[str, Any], secrets: Iterable[str]) -> None:
    serialized = json.dumps(payload, indent=2, ensure_ascii=False, default=str)
    (run_dir / "outcome.json").write_text(redact(serialized, secrets))


def execute_task(settings: Settings, client: ClickUpClient | None, task: Task, dry_run: bool = False) -> dict[str, Any]:
    if dry_run:
        return {"dry_run": True, "task_id": task.id, "branch_prefix": f"agent/clickup-{safe_slug(task.id, 20)}-", "prompt": executor_prompt(task)}
    if client is None:
        raise PipelineError("A ClickUp client is required for a live run.")
    if not settings.isolated_runner:
        raise PipelineError("Live runs require AGENT_ISOLATED_RUNNER=true on a dedicated disposable runner with no user credential stores.")
    # Keep this outside failure recovery: an ineligible task must not be moved
    # to blocked or receive pipeline comments.
    task = client.get_task(task.id)
    if not task.is_queued(settings.queue_status):
        raise PipelineError("Task status or Urgent priority changed before it could be claimed.")
    run_dir = task_run_dir(settings, task)
    (run_dir / "task.json").write_text(json.dumps(task.__dict__, indent=2, ensure_ascii=False))
    worktree: Path | None = None
    branch = ""
    candidate_tree = ""
    pr_url = ""
    commit_sha = ""
    secrets = secret_values(settings)
    try:
        client.set_status(task.id, settings.working_status)
        client.comment(task.id, "Codex pipeline started. Executor is preparing an isolated worktree.")
        worktree, branch = prepare_worktree(settings, task)

        executor_path = run_dir / "executor-result.json"
        executor = run_command(codex_executor_command(settings, worktree, executor_path), worktree, settings.codex_timeout, filtered_agent_environment(), executor_prompt(task))
        save_log(run_dir / "executor.jsonl.log", executor, secrets)
        checked(executor, "Codex Executor")
        executor_result = parse_result(executor_path, "Codex Executor")
        if executor_result.get("status") != "completed":
            raise PipelineError("Executor reported a blocker: " + compact(str(executor_result.get("blocker") or executor_result.get("summary") or "unspecified blocker"), 1500))

        changes = checked(git(settings, "status", "--porcelain", cwd=worktree), "Git status").stdout.strip()
        if not changes:
            raise PipelineError("Executor completed without changing the repository.")

        candidate_tree = stage_candidate(settings, worktree)
        reject_control_plane_changes(settings, worktree)
        reject_staged_secrets(settings, worktree, secrets)

        tests = run_test_command(settings, worktree, filtered_agent_environment(include_codex_auth=False))
        save_log(run_dir / "tests.log", tests, secrets)
        assert_candidate_unchanged(settings, worktree, candidate_tree)

        reviewer_path = run_dir / "reviewer-result.json"
        reviewer = run_command(codex_reviewer_command(settings, worktree, reviewer_path), worktree, settings.codex_timeout, filtered_agent_environment(), reviewer_prompt(task, tests, settings.test_command))
        save_log(run_dir / "reviewer.jsonl.log", reviewer, secrets)
        checked(reviewer, "Codex Reviewer")
        review = parse_result(reviewer_path, "Codex Reviewer")
        assert_candidate_unchanged(settings, worktree, candidate_tree)

        if tests.returncode != 0 or review.get("verdict") != "approve":
            reason = (f"Configured tests failed with exit code {tests.returncode}.\n" if tests.returncode else "") + review_comment(review)
            clickup_transition = safe_clickup_failure(client, settings, task, redact(reason, secrets))
            result = {"status": "blocked", "task_id": task.id, "branch": branch, "worktree": str(worktree), "candidate_tree": candidate_tree, "clickup_transition": clickup_transition, "review": review, "tests_exit_code": tests.returncode}
            save_outcome(run_dir, result, secrets)
            return result

        if not settings.push_approved_branch:
            client.comment(task.id, redact("Tests passed and Reviewer approved the change. Branch push is disabled by AGENT_PUSH_APPROVED_BRANCH=false.\n\n" + review_comment(review), secrets))
            client.set_status(task.id, settings.review_status)
            result = {"status": "approved_local", "task_id": task.id, "branch": branch, "worktree": str(worktree), "candidate_tree": candidate_tree, "review": review}
            save_outcome(run_dir, result, secrets)
            return result

        assert_candidate_unchanged(settings, worktree, candidate_tree)
        pr_url, commit_sha = create_pull_request(settings, task, worktree, branch, candidate_tree, tests, review, run_dir, secrets)
        client.comment(task.id, redact(f"Tests passed and Reviewer approved the change.\n\nPull request: {pr_url}\n\n{review_comment(review)}", secrets))
        client.set_status(task.id, settings.review_status)
        result = {"status": "review", "task_id": task.id, "branch": branch, "commit": commit_sha, "candidate_tree": candidate_tree, "pull_request": pr_url, "clickup_status": settings.review_status, "review": review, "tests_exit_code": tests.returncode}
        save_outcome(run_dir, result, secrets)
        if not settings.keep_success_worktree:
            checked(git(settings, "worktree", "remove", "--force", str(worktree), timeout=300), "Worktree cleanup")
        return result
    except Exception as error:
        message = str(error) if isinstance(error, PipelineError) else f"{type(error).__name__}: {compact(str(error), 1000)}"
        (run_dir / "error.log").write_text(redact(traceback.format_exc(), secrets))
        clickup_transition = safe_clickup_failure(client, settings, task, redact(message, secrets))
        save_outcome(run_dir, {"status": "error", "task_id": task.id, "branch": branch, "worktree": str(worktree or ""), "candidate_tree": candidate_tree, "commit": commit_sha, "pull_request": pr_url, "error": message, "clickup_transition": clickup_transition}, secrets)
        if isinstance(error, PipelineError):
            raise
        raise PipelineError("Unexpected pipeline failure. See the per-task error log.") from error


@contextmanager
def pipeline_lock(settings: Settings):
    settings.runs_dir.mkdir(parents=True, exist_ok=True)
    lock_path = settings.runs_dir / "pipeline.lock"
    with lock_path.open("w") as lock:
        try:
            fcntl.flock(lock.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError as error:
            raise PipelineError("Another pipeline run is already active.") from error
        lock.write(str(os.getpid()))
        lock.flush()
        yield


def doctor(settings: Settings) -> tuple[bool, dict[str, Any]]:
    codex_cli = bool(shutil.which(settings.codex_bin) or Path(settings.codex_bin).exists())
    github_cli = bool(shutil.which(settings.gh_bin) or Path(settings.gh_bin).exists())
    codex_auth = codex_cli and run_command(
        [settings.codex_bin, "login", "status"], settings.repo, 30, filtered_agent_environment()
    ).returncode == 0
    github_auth = github_cli and run_command(
        [settings.gh_bin, "auth", "status"], settings.repo, 30, os.environ.copy()
    ).returncode == 0
    checks = {
        "repository": (settings.repo / ".git").exists(),
        "codex_cli": codex_cli,
        "codex_authenticated": codex_auth,
        "github_cli": github_cli,
        "github_authenticated": github_auth,
        "clickup_token": bool(settings.clickup_token),
        "clickup_list_id": bool(settings.clickup_list_id),
        "test_command": bool(settings.test_command),
        "isolated_runner_attested": settings.isolated_runner,
        "manual_merge_only": True,
    }
    return all(checks.values()), checks


def load_task_file(path: Path) -> Task:
    try:
        payload = json.loads(path.read_text())
    except (OSError, json.JSONDecodeError) as error:
        raise PipelineError("Task fixture must be a readable JSON object.") from error
    if not isinstance(payload, dict):
        raise PipelineError("Task fixture must contain a JSON object.")
    return Task.from_payload(payload)


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description="ClickUp → Codex Executor → tests → Codex Reviewer → PR")
    result.add_argument("--repo", type=Path, default=REPO_ROOT)
    commands = result.add_subparsers(dest="command", required=True)
    commands.add_parser("doctor", help="Check local commands and required settings.")
    commands.add_parser("list", help="List Urgent queued ClickUp tasks without changing them.")
    run = commands.add_parser("run", help="Run one Urgent task, or the oldest Urgent queued task.")
    run.add_argument("--task-id")
    run.add_argument("--task-file", type=Path)
    run.add_argument("--dry-run", action="store_true")
    return result


def main(argv: list[str] | None = None) -> int:
    args = parser().parse_args(argv)
    settings = Settings.from_env(args.repo)
    try:
        if args.command == "doctor":
            ok, checks = doctor(settings)
            print(json.dumps({"ok": ok, "checks": checks}, indent=2))
            return 0 if ok else 1

        client: ClickUpClient | None = None
        if args.command == "list" or not (args.command == "run" and args.task_file and args.dry_run):
            settings.require_clickup()
            client = ClickUpClient(settings.clickup_token)

        if args.command == "list":
            tasks = [task for task in client.list_tasks(settings.clickup_list_id, settings.queue_status) if task.is_queued(settings.queue_status)] if client else []
            print(json.dumps([task.__dict__ for task in tasks], indent=2, ensure_ascii=False))
            return 0

        if args.task_file:
            if not args.dry_run:
                raise PipelineError("Task files are only supported with --dry-run; live tasks must belong to the Urgent ClickUp queue.")
            task = load_task_file(args.task_file)
        elif args.task_id:
            tasks = [task for task in client.list_tasks(settings.clickup_list_id, settings.queue_status) if task.is_queued(settings.queue_status)] if client else []
            task = next((item for item in tasks if item.id == args.task_id), None)
            if task is None:
                raise PipelineError("The requested task is not Urgent in the configured ClickUp queue/status.")
        else:
            tasks = [task for task in client.list_tasks(settings.clickup_list_id, settings.queue_status) if task.is_queued(settings.queue_status)] if client else []
            task = tasks[0] if tasks else None
        if task is None:
            print(json.dumps({"status": "idle", "message": "No Urgent queued ClickUp task was found."}))
            return 0
        with pipeline_lock(settings):
            result = execute_task(settings, client, task, args.dry_run)
        print(json.dumps(result, indent=2, ensure_ascii=False))
        return 0 if result.get("status") not in {"blocked"} else 2
    except PipelineError as error:
        print(json.dumps({"status": "error", "message": str(error)}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
