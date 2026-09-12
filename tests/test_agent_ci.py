import argparse
from dataclasses import asdict, replace
import io
import json
import os
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import Mock, patch

from automation.codex_clickup import ci
from automation.codex_clickup.pipeline import PipelineError, Settings, Task


class AgentCiTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.repo = self.root / "repo"
        self.repo.mkdir()
        self.candidate = self.root / "candidate"
        self.claim_dir = self.root / "claim"
        self.test_dir = self.root / "tests"
        self.review_dir = self.root / "review"
        self.env = patch.dict(os.environ, {}, clear=True)
        self.env.start()
        self.addCleanup(self.env.stop)
        self.settings = replace(
            Settings.from_env(self.repo), clickup_token="test-token", clickup_list_id="123",
            isolated_runner=True,
        )
        self.task = Task("task1", "Change example", "Test acceptance criteria", "", "to do", 1)
        self.git("init", "-q")
        self.git("config", "user.name", "Test")
        self.git("config", "user.email", "test@example.com")
        (self.repo / "app.py").write_text("value = 1\n")
        self.git("add", "--all")
        self.git("commit", "-qm", "Base")
        self.base_sha = ci.current_sha(self.repo)

    def git(self, *args):
        return subprocess.run(["git", *args], cwd=self.repo, capture_output=True, check=True).stdout

    def make_candidate(self, filename="app.py", content="value = 2\n"):
        target = self.repo / filename
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(content)
        ci.write_json(self.candidate / "executor-result.json", {"status": "completed"})
        return ci.snapshot(self.repo, self.candidate)

    def make_claim(self):
        ci.write_json(self.claim_dir / "task.json", asdict(self.task))
        ci.write_json(self.claim_dir / "base.json", {
            "base_sha": self.base_sha, "run_id": "123", "list_id": self.settings.clickup_list_id,
        })

    def make_verification(self, metadata, exit_code=0, verdict="approve", findings=None):
        ci.write_json(self.test_dir / "test-result.json", {**metadata, "exit_code": exit_code, "seconds": 0.1})
        (self.test_dir / "test-output.log").write_text("Ran 1 test\nOK\n")
        ci.write_json(self.review_dir / "reviewer-result.json", {"verdict": verdict, "findings": findings or []})
        ci.record_review(self.repo, argparse.Namespace(candidate_dir=self.candidate, output_dir=self.review_dir))

    def publish_args(self):
        return argparse.Namespace(claim_dir=self.claim_dir, candidate_dir=self.candidate,
                                  test_dir=self.test_dir, review_dir=self.review_dir)

    def test_snapshot_and_apply_bind_the_exact_tree(self):
        metadata = self.make_candidate()
        self.git("reset", "--hard", "HEAD")
        applied = ci.apply_candidate(self.repo, self.candidate, self.base_sha)
        self.assertEqual(applied, metadata)
        self.assertEqual(ci.current_tree(self.repo), metadata["candidate_tree"])
        self.assertEqual((self.repo / "app.py").read_text(), "value = 2\n")
        (self.repo / "app.py").write_text("value = 3\n")
        with self.assertRaisesRegex(PipelineError, "unreviewed"):
            ci.assert_candidate(self.repo, metadata)

    def test_altered_patch_is_rejected_before_application(self):
        self.make_candidate()
        self.git("reset", "--hard", "HEAD")
        path = self.candidate / "candidate.patch"
        path.write_bytes(path.read_bytes() + b"\nmodified\n")
        with self.assertRaisesRegex(PipelineError, "hash"):
            ci.apply_candidate(self.repo, self.candidate, self.base_sha)
        self.assertFalse(self.git("status", "--porcelain").strip())

    def test_wrong_base_and_wrong_tree_are_rejected(self):
        metadata = self.make_candidate()
        self.git("reset", "--hard", "HEAD")
        with self.assertRaisesRegex(PipelineError, "base"):
            ci.apply_candidate(self.repo, self.candidate, "a" * 40)
        ci.write_json(self.candidate / "metadata.json", {**metadata, "candidate_tree": "a" * 40})
        with self.assertRaisesRegex(PipelineError, "tree differs"):
            ci.apply_candidate(self.repo, self.candidate, self.base_sha)

    def test_malformed_hash_valid_patch_is_rejected(self):
        metadata = self.make_candidate()
        self.git("reset", "--hard", "HEAD")
        invalid = b"this is not a git patch\n"
        (self.candidate / "candidate.patch").write_bytes(invalid)
        ci.write_json(self.candidate / "metadata.json", {**metadata, "patch_sha256": ci.digest(invalid)})
        with self.assertRaises(PipelineError):
            ci.apply_candidate(self.repo, self.candidate, self.base_sha)
        self.assertFalse(self.git("status", "--porcelain").strip())

    def test_control_file_additions_are_rejected(self):
        for name in ("automation/new.py", ".github/workflows/attack.yml", "nested/AGENTS.md",
                     ".gitattributes", ".env.production", "nested/auth.json"):
            with self.subTest(name=name):
                self.git("reset", "--hard", "HEAD")
                self.git("clean", "-fd")
                with self.assertRaisesRegex(PipelineError, "protected"):
                    self.make_candidate(name, "candidate controlled\n")

    def test_control_file_deletion_and_rename_are_rejected(self):
        control = self.repo / "AGENTS.md"
        control.write_text("Trusted instructions\n")
        self.git("add", "--all")
        self.git("commit", "-qm", "Add control")
        for action in ("delete", "rename"):
            with self.subTest(action=action):
                self.git("reset", "--hard", "HEAD")
                self.git("clean", "-fd")
                control.unlink() if action == "delete" else control.rename(self.repo / "notes.md")
                ci.write_json(self.candidate / "executor-result.json", {"status": "completed"})
                with self.assertRaisesRegex(PipelineError, "protected"):
                    ci.snapshot(self.repo, self.candidate)

    def test_forged_control_patch_is_rejected_by_publisher_application(self):
        target = self.repo / ".github" / "workflows" / "attack.yml"
        target.parent.mkdir(parents=True)
        target.write_text("name: attack\n")
        self.git("add", "--all")
        raw = self.git("diff", "--cached", "--binary", "--full-index", "HEAD")
        metadata = {"base_sha": self.base_sha, "candidate_tree": ci.current_tree(self.repo),
                    "patch_sha256": ci.digest(raw)}
        ci.write_json(self.candidate / "metadata.json", metadata)
        (self.candidate / "candidate.patch").write_bytes(raw)
        self.git("reset", "--hard", "HEAD")
        with self.assertRaisesRegex(PipelineError, "protected"):
            ci.apply_candidate(self.repo, self.candidate, self.base_sha)

    def test_candidate_symlinks_are_rejected(self):
        (self.repo / "escape").symlink_to("/tmp")
        ci.write_json(self.candidate / "executor-result.json", {"status": "completed"})
        with self.assertRaisesRegex(PipelineError, "Symlinks"):
            ci.snapshot(self.repo, self.candidate)

    def test_symlink_artifacts_are_rejected(self):
        self.make_candidate()
        path = self.candidate / "candidate.patch"
        saved = self.root / "saved.patch"
        path.rename(saved)
        path.symlink_to(saved)
        with self.assertRaisesRegex(PipelineError, "patch artifact"):
            ci.candidate_metadata(self.candidate)

    def test_verification_results_cannot_be_reused_for_another_tree(self):
        metadata = self.make_candidate()
        ci.write_json(self.test_dir / "test-result.json", {**metadata, "candidate_tree": "b" * 40, "exit_code": 0})
        with self.assertRaisesRegex(PipelineError, "different candidate"):
            ci.matched_result(self.test_dir / "test-result.json", metadata)

    def test_review_cannot_be_recorded_after_candidate_changes(self):
        self.make_candidate()
        ci.write_json(self.review_dir / "reviewer-result.json", {"verdict": "approve", "findings": []})
        (self.repo / "app.py").write_text("value = 99\n")
        with self.assertRaisesRegex(PipelineError, "unreviewed"):
            ci.record_review(self.repo, argparse.Namespace(candidate_dir=self.candidate, output_dir=self.review_dir))
        self.assertFalse((self.review_dir / "review-metadata.json").exists())

    def test_preflight_without_credentials_does_not_access_queue(self):
        config = replace(self.settings, clickup_token="", clickup_list_id="")
        with patch.object(ci, "ClickUpClient") as client:
            result = ci.preflight(config)
        self.assertFalse(result["ready"])
        client.assert_not_called()

    def test_claim_without_auth_or_isolation_cannot_mutate_queue(self):
        args = argparse.Namespace(task_id="", run_id="123", output_dir=self.claim_dir)
        configurations = [({}, self.settings),
                          ({"CLICKUP_AGENT_ENABLED": "true"}, self.settings),
                          ({"CLICKUP_AGENT_ENABLED": "true", "CODEX_AUTH_CONFIGURED": "true"},
                           replace(self.settings, isolated_runner=False)),
                          ({"CLICKUP_AGENT_ENABLED": "true", "CODEX_AUTH_CONFIGURED": "true"},
                           replace(self.settings, clickup_token=""))]
        for env, config in configurations:
            with self.subTest(env=env, isolated=config.isolated_runner), patch.dict(os.environ, env, clear=True), patch.object(ci, "ClickUpClient") as client:
                with self.assertRaises(PipelineError):
                    ci.claim(config, args)
                client.assert_not_called()
        self.assertFalse(self.claim_dir.exists())

    def test_failed_claim_comment_preserves_recovery_artifacts(self):
        client = Mock()
        client.list_tasks.return_value = [self.task]
        client.get_task.return_value = self.task
        client.comment.side_effect = PipelineError("Comment failed")
        args = argparse.Namespace(task_id="", run_id="123", output_dir=self.claim_dir)
        with patch.dict(os.environ, {"CLICKUP_AGENT_ENABLED": "true", "CODEX_AUTH_CONFIGURED": "true"}), patch.object(ci, "ClickUpClient", return_value=client):
            with self.assertRaisesRegex(PipelineError, "Comment failed"):
                ci.claim(self.settings, args)
        self.assertEqual(ci.load_claim(self.claim_dir)[0].id, self.task.id)
        client.set_status.assert_called_once_with(self.task.id, "in progress")

    def test_claim_status_race_does_not_mutate_task(self):
        client = Mock()
        client.list_tasks.return_value = [self.task]
        client.get_task.return_value = replace(self.task, status="review")
        args = argparse.Namespace(task_id="", run_id="123", output_dir=self.claim_dir)
        with patch.dict(os.environ, {"CLICKUP_AGENT_ENABLED": "true", "CODEX_AUTH_CONFIGURED": "true"}), patch.object(ci, "ClickUpClient", return_value=client):
            with self.assertRaisesRegex(PipelineError, "changed"):
                ci.claim(self.settings, args)
        client.set_status.assert_not_called()
        client.comment.assert_not_called()

    def test_claim_priority_race_does_not_mutate_task(self):
        client = Mock()
        client.list_tasks.return_value = [self.task]
        client.get_task.return_value = replace(self.task, priority=2)
        args = argparse.Namespace(task_id="", run_id="123", output_dir=self.claim_dir)
        with patch.dict(os.environ, {"CLICKUP_AGENT_ENABLED": "true", "CODEX_AUTH_CONFIGURED": "true"}), patch.object(ci, "ClickUpClient", return_value=client):
            with self.assertRaisesRegex(PipelineError, "Urgent priority changed"):
                ci.claim(self.settings, args)
        client.set_status.assert_not_called()
        client.comment.assert_not_called()
        self.assertFalse(self.claim_dir.exists())

    def test_claim_only_urgent_and_manual_id_cannot_bypass_filter(self):
        for task_id in ("", self.task.id):
            with self.subTest(task_id=task_id):
                client = Mock()
                client.list_tasks.return_value = [replace(self.task, priority=3)]
                args = argparse.Namespace(task_id=task_id, run_id="123", output_dir=self.claim_dir)
                with patch.dict(os.environ, {"CLICKUP_AGENT_ENABLED": "true", "CODEX_AUTH_CONFIGURED": "true"}), patch.object(ci, "ClickUpClient", return_value=client):
                    if task_id:
                        with self.assertRaisesRegex(PipelineError, "not Urgent"):
                            ci.claim(self.settings, args)
                    else:
                        self.assertEqual(ci.claim(self.settings, args), {"status": "idle"})
                client.get_task.assert_not_called()
                client.set_status.assert_not_called()
                client.comment.assert_not_called()

    def test_failure_recovery_preserves_a_task_moved_by_the_user(self):
        self.make_claim()
        client = Mock()
        client.get_task.return_value = replace(self.task, status="done")
        with patch.object(ci, "ClickUpClient", return_value=client):
            result = ci.fail(self.settings, argparse.Namespace(claim_dir=self.claim_dir, reason="Executor failed"))
        self.assertEqual(result["status"], "unchanged")
        client.set_status.assert_not_called()
        client.comment.assert_not_called()

    def test_test_container_has_no_network_or_host_credentials_and_records_failure(self):
        metadata = self.make_candidate()
        args = argparse.Namespace(candidate_dir=self.candidate, output_dir=self.test_dir,
                                  container_image="clickup-agent-tests:123")
        real_run = subprocess.run
        docker_commands = []

        def run(command, **kwargs):
            if command[0] == "docker":
                docker_commands.append(command)
                return subprocess.CompletedProcess(command, 7 if command[1] == "run" else 0, "", "test failed")
            return real_run(command, **kwargs)

        with patch.object(ci.subprocess, "run", side_effect=run):
            result = ci.run_tests(self.repo, args)
        self.assertEqual(result["exit_code"], 7)
        self.assertEqual(ci.matched_result(self.test_dir / "test-result.json", metadata)["exit_code"], 7)
        command = docker_commands[0]
        self.assertIn("--network=none", command)
        self.assertIn("--read-only", command)
        self.assertIn("--cap-drop=ALL", command)
        self.assertNotIn("-e", command)
        self.assertNotIn("--env", command)
        self.assertEqual(command.count("--mount"), 1)
        self.assertIn("target=/input/candidate.tar,readonly", command[command.index("--mount") + 1])
        self.assertEqual(docker_commands[-1][:3], ["docker", "rm", "-f"])

    def test_test_command_exits_nonzero_when_suite_fails(self):
        with patch.object(ci, "run_tests", return_value={"exit_code": 7}), patch("sys.stdout", new_callable=io.StringIO):
            code = ci.main(["--repo", str(self.repo), "test", "--candidate-dir", str(self.candidate),
                            "--output-dir", str(self.test_dir), "--container-image", "clickup-agent-tests:123"])
        self.assertNotEqual(code, 0, "Failing candidate tests must fail the GitHub Actions test job")

    def test_publish_rejects_failed_tests_or_changed_review_before_remote_calls(self):
        metadata = self.make_candidate()
        self.make_claim()
        self.make_verification(metadata, exit_code=1)
        with patch.object(ci, "ClickUpClient") as client:
            with self.assertRaisesRegex(PipelineError, "must approve"):
                ci.publish(self.settings, self.publish_args())
            client.assert_not_called()
        self.make_verification(metadata)
        (self.review_dir / "reviewer-result.json").write_text('{"verdict":"approve","findings":[],"changed":true}')
        with patch.object(ci, "ClickUpClient") as client:
            with self.assertRaisesRegex(PipelineError, "changed after verification"):
                ci.publish(self.settings, self.publish_args())
            client.assert_not_called()

    def test_publisher_commits_only_reviewed_tree_and_never_executes_candidate(self):
        marker = self.root / "candidate-executed"
        metadata = self.make_candidate("sitecustomize.py", "from pathlib import Path\nPath(" + repr(str(marker)) + ").write_text('unsafe')\n")
        self.make_claim()
        self.make_verification(metadata)
        self.git("reset", "--hard", "HEAD")
        hook = self.repo / ".git" / "hooks" / "pre-commit"
        hook.write_text("#!/bin/sh\ntouch " + str(marker) + "\n")
        hook.chmod(0o755)
        client = Mock()
        client.get_task.return_value = replace(self.task, status="in progress")
        commands = []
        real_run = subprocess.run

        def run(command, **kwargs):
            commands.append(command)
            if command[0] == "gh":
                return subprocess.CompletedProcess(command, 0, "https://github.com/example/repo/pull/1\n", "")
            if command[0] == "git" and "push" in command:
                return subprocess.CompletedProcess(command, 0, b"", b"")
            return real_run(command, **kwargs)

        with patch.object(ci, "ClickUpClient", return_value=client), patch.object(ci.subprocess, "run", side_effect=run):
            result = ci.publish(self.settings, self.publish_args())
        self.assertEqual(result["candidate_tree"], metadata["candidate_tree"])
        self.assertEqual(self.git("rev-parse", "HEAD^{tree}").decode().strip(), metadata["candidate_tree"])
        self.assertFalse(marker.exists())
        self.assertTrue(all(command[0] in {"git", "gh"} for command in commands))
        self.assertTrue(all("core.hooksPath=/dev/null" in command for command in commands if command[0] == "git"))
        client.set_status.assert_called_once_with(self.task.id, "review")


if __name__ == "__main__":
    unittest.main()
