"""Trusted commands for the ClickUp pipeline's separate GitHub Actions jobs.

Always invoke this file from the immutable control checkout with Python -I.
Candidate files are data in orchestration jobs; only the disposable test
container and Codex jobs execute candidate code.
"""
from __future__ import annotations

import argparse
from dataclasses import asdict
import hashlib
import json
import os
from pathlib import Path, PurePosixPath
import re
import subprocess
import sys
import tempfile
import time

# -I excludes both the candidate cwd and PYTHONPATH. Only this trusted checkout
# supplies imports, including the queue adapter and role prompts.
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from automation.codex_clickup.pipeline import (
    ClickUpClient, CommandResult, PipelineError, Settings, Task,
    executor_prompt, reviewer_prompt, clickup_path_id, redact, safe_slug,
)

MAX_PATCH_BYTES = 10 * 1024 * 1024
TEST_COMMAND = ["python3", "-m", "unittest", "discover", "-s", "tests", "-p", "test_*.py"]
PROTECTED_PREFIXES = (".github/", ".codex/", "automation/", "scripts/agent-pipeline")


def write_json(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n")


def read_json(path: Path) -> dict:
    if path.is_symlink() or not path.is_file() or path.stat().st_size > MAX_PATCH_BYTES:
        raise PipelineError("Missing or invalid pipeline artifact: " + path.name)
    try:
        value = json.loads(path.read_text())
    except (OSError, ValueError) as error:
        raise PipelineError("Invalid JSON artifact: " + path.name) from error
    if not isinstance(value, dict):
        raise PipelineError("Expected an object in " + path.name)
    return value


def digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha(value: str) -> str:
    if not re.fullmatch(r"[a-f0-9]{40}", value or ""):
        raise PipelineError("Invalid Git revision in pipeline artifact.")
    return value


def output(name: str, value: str) -> None:
    if not re.fullmatch(r"[a-z_]+", name) or "\n" in value or "\r" in value:
        raise PipelineError("Invalid workflow output.")
    if os.getenv("GITHUB_OUTPUT"):
        with Path(os.environ["GITHUB_OUTPUT"]).open("a") as stream:
            stream.write(f"{name}={value}\n")


def git(repo: Path, *args: str, input_data: bytes | None = None) -> bytes:
    command = ["git", "-c", "core.hooksPath=/dev/null", "-c", "core.fsmonitor=false", *args]
    result = subprocess.run(command, cwd=repo, input=input_data, capture_output=True, timeout=180)
    if result.returncode:
        # Git output can contain candidate text. Do not echo it to public logs.
        raise PipelineError("Git validation failed: " + args[0])
    return result.stdout


def current_sha(repo: Path) -> str:
    return sha(git(repo, "rev-parse", "HEAD").decode().strip())


def current_tree(repo: Path) -> str:
    return sha(git(repo, "write-tree").decode().strip())


def protected(path: str) -> bool:
    p = PurePosixPath(path)
    return (p.is_absolute() or ".." in p.parts or any(x in {".git", ".codex"} for x in p.parts)
            or p.name in {"AGENTS.md", "AGENTS.override.md", ".gitattributes", ".gitmodules"}
            or path.startswith(PROTECTED_PREFIXES)
            or p.name.startswith(".env")
            or p.name in {"auth.json", "credentials.json"})


def validate_index(repo: Path) -> None:
    names = git(repo, "diff", "--cached", "--name-only", "--no-renames", "-z", "HEAD").decode().split("\0")
    if any(protected(name) for name in names if name):
        raise PipelineError("Candidate changes protected control or credential files.")
    for entry in git(repo, "ls-files", "--stage", "-z").decode().split("\0"):
        if entry and entry.split(" ", 1)[0] not in {"100644", "100755"}:
            raise PipelineError("Symlinks and submodules are not accepted in automatic candidates.")


def preflight(settings: Settings) -> dict:
    checks = {
        "clickup_token": bool(settings.clickup_token),
        "clickup_list_id": bool(settings.clickup_list_id),
        "openai_api_key": bool(os.getenv("OPENAI_API_KEY")),
        "queue_enabled": os.getenv("CLICKUP_AGENT_ENABLED", "false").lower() == "true",
    }
    missing_statuses: list[str] = []
    if checks["clickup_token"] and checks["clickup_list_id"]:
        list_id = clickup_path_id(settings.clickup_list_id, "list id", digits_only=True)
        payload = ClickUpClient(settings.clickup_token)._request("GET", f"/list/{list_id}")
        actual = {str(item.get("status", "")).casefold() for item in payload.get("statuses", [])}
        expected = [settings.queue_status, settings.working_status, settings.review_status, settings.blocked_status]
        missing_statuses = [name for name in expected if name.casefold() not in actual]
        checks["clickup_statuses"] = not missing_statuses
    else:
        checks["clickup_statuses"] = False
    ready = all(checks.values())
    output("ready", str(ready).lower())
    output("base_sha", current_sha(settings.repo))
    result = {"ready": ready, "checks": checks, "missing_statuses": missing_statuses}
    if os.getenv("GITHUB_STEP_SUMMARY"):
        with Path(os.environ["GITHUB_STEP_SUMMARY"]).open("a") as stream:
            stream.write("### ClickUp runner preflight\n\n")
            for name, passed in checks.items():
                stream.write(f"- {'Ready' if passed else 'Missing'}: `{name}`\n")
            if missing_statuses:
                stream.write("\nAdd list statuses: " + ", ".join(missing_statuses) + "\n")
            stream.write("\nPreflight is read-only; it does not claim a task.\n")
    return result


def claim(settings: Settings, args: argparse.Namespace) -> dict:
    settings.require_clickup()
    if not settings.isolated_runner or os.getenv("CLICKUP_AGENT_ENABLED") != "true":
        raise PipelineError("Queue is disabled or the runner is not isolated.")
    if os.getenv("CODEX_AUTH_CONFIGURED") != "true":
        raise PipelineError("Codex authentication preflight has not passed.")
    client = ClickUpClient(settings.clickup_token)
    tasks = client.list_tasks(settings.clickup_list_id, settings.queue_status)
    task = next((item for item in tasks if item.id == args.task_id), None) if args.task_id else next(iter(tasks), None)
    if args.task_id and task is None:
        raise PipelineError("Requested task is not in the configured queue.")
    if task is None:
        output("has_task", "false")
        return {"status": "idle"}
    clickup_path_id(task.id, "task id")
    # Verify the current status again just before claiming it.
    if client.get_task(task.id).status.casefold() != settings.queue_status.casefold():
        raise PipelineError("Task status changed before it could be claimed.")
    base_sha = current_sha(settings.repo)
    run_id = args.run_id
    if not re.fullmatch(r"[0-9]+", run_id):
        raise PipelineError("Invalid workflow run ID.")
    args.output_dir.mkdir(parents=True, exist_ok=True)
    write_json(args.output_dir / "task.json", asdict(task))
    write_json(args.output_dir / "base.json", {"base_sha": base_sha, "run_id": run_id, "list_id": settings.clickup_list_id})
    (args.output_dir / "executor-prompt.md").write_text(executor_prompt(task))
    client.set_status(task.id, settings.working_status)
    output("has_task", "true")
    output("base_sha", base_sha)
    output("task_id", task.id)
    client.comment(task.id, "Codex Executor + Reviewer started. Changes will be tested and submitted as a pull request for review.")
    return {"status": "claimed", "task_id": task.id}


def snapshot(repo: Path, destination: Path) -> dict:
    destination.mkdir(parents=True, exist_ok=True)
    result = read_json(destination / "executor-result.json")
    if result.get("status") != "completed":
        raise PipelineError("Executor did not complete the task.")
    git(repo, "add", "--all")
    validate_index(repo)
    patch = git(repo, "diff", "--cached", "--binary", "--no-ext-diff", "--no-textconv", "--full-index", "HEAD")
    if not patch or len(patch) > MAX_PATCH_BYTES:
        raise PipelineError("Candidate is empty or exceeds the automatic patch limit.")
    metadata = {"base_sha": current_sha(repo), "candidate_tree": current_tree(repo), "patch_sha256": digest(patch)}
    (destination / "candidate.patch").write_bytes(patch)
    write_json(destination / "metadata.json", metadata)
    return metadata


def candidate_metadata(folder: Path) -> dict:
    metadata = read_json(folder / "metadata.json")
    sha(metadata.get("base_sha"))
    sha(metadata.get("candidate_tree"))
    patch = folder / "candidate.patch"
    if patch.is_symlink() or not patch.is_file() or patch.stat().st_size > MAX_PATCH_BYTES:
        raise PipelineError("Invalid candidate patch artifact.")
    if digest(patch.read_bytes()) != metadata.get("patch_sha256"):
        raise PipelineError("Candidate patch hash does not match its metadata.")
    return metadata


def apply_candidate(repo: Path, folder: Path, base_sha: str) -> dict:
    metadata = candidate_metadata(folder)
    if metadata["base_sha"] != sha(base_sha) or current_sha(repo) != base_sha:
        raise PipelineError("Candidate base does not match the trusted base revision.")
    if git(repo, "status", "--porcelain").strip():
        raise PipelineError("Candidate application requires a clean checkout.")
    git(repo, "apply", "--index", "--binary", "--whitespace=nowarn", str((folder / "candidate.patch").resolve()))
    validate_index(repo)
    assert_candidate(repo, metadata)
    return metadata


def assert_candidate(repo: Path, metadata: dict) -> None:
    if current_tree(repo) != metadata["candidate_tree"]:
        raise PipelineError("Candidate tree differs from the recorded tree.")
    if git(repo, "diff", "--name-only", "--no-ext-diff").strip() or git(repo, "ls-files", "--others", "--exclude-standard").strip():
        raise PipelineError("Candidate contains unreviewed file changes.")


def run_tests(repo: Path, args: argparse.Namespace) -> dict:
    metadata = candidate_metadata(args.candidate_dir)
    assert_candidate(repo, metadata)
    if not re.fullmatch(r"clickup-agent-tests:[a-zA-Z0-9_.-]+", args.container_image):
        raise PipelineError("Unexpected test container image.")
    args.output_dir.mkdir(parents=True, exist_ok=True)
    name = "clickup-tests-" + str(os.getpid())
    # Only tracked candidate files enter the container. No .git directory,
    # orchestrator artifacts, Docker socket, host home, or credentials are mounted.
    with tempfile.TemporaryDirectory(prefix="clickup-tests-") as temp:
        temp_path = Path(temp)
        archive = temp_path / "candidate.tar"
        archive.write_bytes(git(repo, "archive", metadata["candidate_tree"]))
        command = ["docker", "run", "--rm", "--name", name, "--network=none", "--cap-drop=ALL", "--security-opt=no-new-privileges", "--pids-limit=256", "--memory=3g", "--cpus=2", "--read-only", "--tmpfs", "/tmp:rw,nosuid,nodev,size=512m", "--tmpfs", "/work:rw,nosuid,nodev,size=512m", "--mount", f"type=bind,source={archive},target=/input/candidate.tar,readonly", args.container_image, "sh", "-c", "tar -xf /input/candidate.tar -C /work && cd /work && exec " + " ".join(TEST_COMMAND[:-1]) + " 'test_*.py'"]
        start = time.monotonic()
        try:
            run = subprocess.run(command, capture_output=True, text=True, timeout=900)
            code, log = run.returncode, run.stdout + run.stderr
        except subprocess.TimeoutExpired:
            code, log = 124, "Test container exceeded 900 seconds."
        finally:
            subprocess.run(["docker", "rm", "-f", name], capture_output=True, timeout=30)
        result = {**metadata, "exit_code": code, "seconds": round(time.monotonic() - start, 3)}
        (args.output_dir / "test-output.log").write_text(log[-100000:])
        write_json(args.output_dir / "test-result.json", result)
    assert_candidate(repo, metadata)
    return result


def matched_result(path: Path, metadata: dict) -> dict:
    result = read_json(path)
    if any(result.get(key) != metadata[key] for key in ("base_sha", "candidate_tree", "patch_sha256")):
        raise PipelineError("Verification result belongs to a different candidate.")
    return result


def load_claim(folder: Path) -> tuple[Task, dict]:
    payload = read_json(folder / "task.json")
    task = Task.from_payload(payload)
    clickup_path_id(task.id, "task id")
    base = read_json(folder / "base.json")
    sha(base.get("base_sha"))
    return task, base


def prepare_review(args: argparse.Namespace) -> dict:
    task, base = load_claim(args.claim_dir)
    metadata = candidate_metadata(args.candidate_dir)
    if metadata["base_sha"] != base["base_sha"]:
        raise PipelineError("Claim and candidate bases differ.")
    tests = matched_result(args.test_dir / "test-result.json", metadata)
    log_path = args.test_dir / "test-output.log"
    log = log_path.read_text()[-12000:]
    prompt = reviewer_prompt(task, CommandResult(tests["exit_code"], log, "", tests["seconds"]), " ".join(TEST_COMMAND))
    args.output_dir.mkdir(parents=True, exist_ok=True)
    (args.output_dir / "reviewer-prompt.md").write_text(prompt)
    return {"prepared": True}


def record_review(repo: Path, args: argparse.Namespace) -> dict:
    metadata = candidate_metadata(args.candidate_dir)
    assert_candidate(repo, metadata)
    result = read_json(args.output_dir / "reviewer-result.json")
    if result.get("verdict") not in {"approve", "request_changes"} or not isinstance(result.get("findings"), list):
        raise PipelineError("Reviewer did not produce a valid structured verdict.")
    write_json(args.output_dir / "review-metadata.json", {**metadata, "result_sha256": digest((args.output_dir / "reviewer-result.json").read_bytes())})
    return {"verdict": result["verdict"]}


def publish(settings: Settings, args: argparse.Namespace) -> dict:
    task, base = load_claim(args.claim_dir)
    metadata = candidate_metadata(args.candidate_dir)
    tests = matched_result(args.test_dir / "test-result.json", metadata)
    review_metadata = matched_result(args.review_dir / "review-metadata.json", metadata)
    review = read_json(args.review_dir / "reviewer-result.json")
    if digest((args.review_dir / "reviewer-result.json").read_bytes()) != review_metadata.get("result_sha256"):
        raise PipelineError("Reviewer result changed after verification.")
    if tests.get("exit_code") != 0 or review.get("verdict") != "approve" or review.get("findings"):
        raise PipelineError("Tests and Reviewer must approve the candidate without unresolved findings.")
    if base.get("list_id") != settings.clickup_list_id:
        raise PipelineError("Claim belongs to another ClickUp list.")
    client = ClickUpClient(settings.clickup_token)
    current = client.get_task(task.id)
    if current.status.casefold() != settings.working_status.casefold():
        raise PipelineError("Task is no longer in the claimed status.")
    apply_candidate(settings.repo, args.candidate_dir, base["base_sha"])
    branch = f"agent/clickup-{safe_slug(task.id, 30)}-{base['run_id']}"
    git(settings.repo, "checkout", "-b", branch)
    git(settings.repo, "-c", "user.name=DataProvido Agent", "-c", "user.email=agent@dataprovido.com", "commit", "-m", f"Implement ClickUp task {task.id}")
    if git(settings.repo, "rev-parse", "HEAD^{tree}").decode().strip() != metadata["candidate_tree"]:
        raise PipelineError("Committed tree differs from reviewed candidate.")
    git(settings.repo, "push", "origin", f"HEAD:refs/heads/{branch}")
    # Do not publish the private ClickUp task description or model transcript.
    with tempfile.TemporaryDirectory(prefix="clickup-pr-") as temp:
        body = Path(temp) / "body.md"
        body.write_text(f"Implements ClickUp task `{task.id}`.\n\nThe isolated test suite passed and the independent Reviewer approved this exact candidate.\n\nManual merge required.\n")
        command = ["gh", "pr", "create", "--head", branch, "--base", settings.base_branch, "--title", f"[ClickUp {task.id}] {task.name[:180]}", "--body-file", str(body)]
        result = subprocess.run(command, cwd=settings.repo, capture_output=True, text=True, timeout=180)
    if result.returncode:
        raise PipelineError("Branch was pushed but pull request creation failed.")
    url = result.stdout.strip()
    if not url.startswith("https://github.com/"):
        raise PipelineError("GitHub returned an unexpected pull request URL.")
    client.comment(task.id, "Tests passed and Reviewer approved the change.\n\nPull request: " + url)
    client.set_status(task.id, settings.review_status)
    return {"status": "review", "task_id": task.id, "pull_request": url, **metadata}


def fail(settings: Settings, args: argparse.Namespace) -> dict:
    task, base = load_claim(args.claim_dir)
    if base.get("list_id") != settings.clickup_list_id:
        raise PipelineError("Failure claim belongs to another list.")
    client = ClickUpClient(settings.clickup_token)
    if client.get_task(task.id).status.casefold() != settings.working_status.casefold():
        return {"status": "unchanged", "task_id": task.id}
    client.set_status(task.id, settings.blocked_status)
    client.comment(task.id, "Agent run stopped before completing review. " + args.reason[:500] + " The task needs attention before retrying.")
    return {"status": "blocked", "task_id": task.id}


def parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--repo", type=Path, default=Path.cwd())
    commands = p.add_subparsers(dest="command", required=True)
    pre = commands.add_parser("preflight")
    pre.add_argument("--require-enabled", action="store_true")
    c = commands.add_parser("claim")
    c.add_argument("--output-dir", type=Path, required=True)
    c.add_argument("--run-id", required=True)
    c.add_argument("--task-id", default="")
    s = commands.add_parser("snapshot")
    s.add_argument("--output-dir", type=Path, required=True)
    a = commands.add_parser("apply")
    a.add_argument("--candidate-dir", type=Path, required=True)
    a.add_argument("--base-sha", required=True)
    t = commands.add_parser("test")
    t.add_argument("--candidate-dir", type=Path, required=True)
    t.add_argument("--output-dir", type=Path, required=True)
    t.add_argument("--container-image", required=True)
    for name in ("prepare-review", "publish"):
        cmd = commands.add_parser(name)
        for folder in ("claim", "candidate", "test"):
            cmd.add_argument(f"--{folder}-dir", type=Path, required=True)
        cmd.add_argument("--output-dir" if name == "prepare-review" else "--review-dir", type=Path, required=True)
    r = commands.add_parser("record-review")
    r.add_argument("--candidate-dir", type=Path, required=True)
    r.add_argument("--output-dir", type=Path, required=True)
    f = commands.add_parser("fail")
    f.add_argument("--claim-dir", type=Path, required=True)
    f.add_argument("--reason", required=True)
    return p


def main(argv: list[str] | None = None) -> int:
    args = parser().parse_args(argv)
    settings = Settings.from_env(args.repo)
    try:
        if args.command == "preflight":
            result = preflight(settings)
            print(json.dumps(result))
            return 1 if args.require_enabled and not result["ready"] else 0
        handlers = {
            "claim": lambda: claim(settings, args),
            "snapshot": lambda: snapshot(settings.repo, args.output_dir),
            "apply": lambda: apply_candidate(settings.repo, args.candidate_dir, args.base_sha),
            "test": lambda: run_tests(settings.repo, args),
            "prepare-review": lambda: prepare_review(args),
            "record-review": lambda: record_review(settings.repo, args),
            "publish": lambda: publish(settings, args),
            "fail": lambda: fail(settings, args),
        }
        result = handlers[args.command]()
        print(json.dumps(result))
        return 1 if args.command == "test" and result["exit_code"] != 0 else 0
    except (PipelineError, OSError, ValueError, subprocess.SubprocessError) as error:
        # Avoid printing raw transport responses, candidate content, or credentials.
        message = str(error) if isinstance(error, PipelineError) else type(error).__name__
        secrets = [os.getenv(k, "") for k in ("CLICKUP_API_TOKEN", "OPENAI_API_KEY", "GH_TOKEN", "GITHUB_TOKEN")]
        print(json.dumps({"status": "error", "message": redact(message, secrets)}), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
