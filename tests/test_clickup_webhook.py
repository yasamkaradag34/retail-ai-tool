import hashlib
import hmac
import json
import os
import time
import unittest
from unittest.mock import Mock, patch

from fastapi import FastAPI
from fastapi.testclient import TestClient
import requests

from automation.codex_clickup import webhook


class ClickUpWebhookTests(unittest.TestCase):
    def setUp(self):
        self.environment = patch.dict(os.environ, {
            "CLICKUP_WEBHOOK_ENABLED": "true",
            "CLICKUP_WEBHOOK_SECRET": "test-signing-secret",
            "CLICKUP_API_TOKEN": "test-clickup-token",
            "CLICKUP_LIST_ID": "1234",
            "GITHUB_AGENT_DISPATCH_TOKEN": "test-github-token",
            "CLICKUP_WEBHOOK_ID": "hook-123",
        }, clear=True)
        self.environment.start()
        self.addCleanup(self.environment.stop)
        app = FastAPI()
        app.include_router(webhook.router)
        self.client = TestClient(app)
        self.addCleanup(self.client.close)
        self.payload = {"event": "taskPriorityUpdated", "task_id": "task123",
                        "webhook_id": "hook-123", "history_items": [
                            {"id": "history1", "date": str(int(time.time() * 1000)),
                             "field": "priority", "after": {"id": "1"}}]}
        self.task = {"id": "task123", "list": {"id": "1234"},
                     "priority": {"id": "1"}, "status": {"status": "to do"}}

    def send(self, payload=None, raw=None, signature=None):
        if raw is None:
            raw = json.dumps(self.payload if payload is None else payload).encode()
        if signature is None:
            signature = hmac.new(b"test-signing-secret", raw, hashlib.sha256).hexdigest()
        return self.client.post("/integrations/clickup/webhook", content=raw,
                                headers={"X-Signature": signature, "Content-Type": "application/json"})

    def test_disabled_by_default_and_missing_configuration_never_dispatch(self):
        with patch.object(webhook, "deliver") as delivery:
            for name in ("CLICKUP_WEBHOOK_ENABLED", "CLICKUP_WEBHOOK_SECRET", "CLICKUP_API_TOKEN",
                         "CLICKUP_LIST_ID", "GITHUB_AGENT_DISPATCH_TOKEN"):
                with self.subTest(name=name), patch.dict(os.environ, {name: ""}):
                    self.assertEqual(self.send().status_code, 503)
            delivery.assert_not_called()

    def test_invalid_signatures_and_changed_raw_bytes_never_fetch_tasks(self):
        original = json.dumps(self.payload).encode()
        signature = hmac.new(b"test-signing-secret", original, hashlib.sha256).hexdigest()
        with patch.object(webhook, "fetch_task") as fetch:
            for supplied in ("", "wrong", "0" * 64):
                self.assertEqual(self.send(signature=supplied).status_code, 401)
            self.assertEqual(self.send(raw=original + b" ", signature=signature).status_code, 401)
            fetch.assert_not_called()

    def test_malformed_signed_json_and_task_paths_are_rejected(self):
        with patch.object(webhook, "deliver") as delivery:
            for raw in (b"not-json", b"[]", b"null", b"\xff"):
                self.assertEqual(self.send(raw=raw).status_code, 400)
            for task_id in ("../other", "task?include=true", "task#secret", "", None, 1):
                self.assertEqual(self.send({**self.payload, "task_id": task_id}).status_code, 400)
            delivery.assert_not_called()

    def test_payload_limit_and_wrong_webhook_are_rejected(self):
        self.assertEqual(self.send(raw=b"x" * (webhook.MAX_BODY_BYTES + 1)).status_code, 413)
        self.assertEqual(self.send({**self.payload, "webhook_id": "other-hook"}).status_code, 401)

    def test_nontrigger_events_are_acknowledged_without_provider_calls(self):
        with patch.object(webhook, "deliver") as delivery:
            for event in ("taskCommentPosted", "taskDeleted", "listUpdated", None, {}):
                response = self.send({**self.payload, "event": event})
                self.assertEqual(response.status_code, 200)
                self.assertEqual(response.json(), {"status": "ignored"})
            delivery.assert_not_called()

    def test_only_current_urgent_queue_task_in_configured_list_dispatches(self):
        invalid_tasks = [
            {}, {**self.task, "id": "other"}, {**self.task, "list": {"id": "other"}},
            {**self.task, "priority": {"id": "2"}}, {**self.task, "priority": None},
            {**self.task, "priority": {"id": True}}, {**self.task, "archived": True},
            {**self.task, "status": {"status": "in progress"}},
            {**self.task, "status": {"status": "review"}},
        ]
        with patch.object(webhook, "dispatch_workflow") as dispatch:
            for task in invalid_tasks:
                with self.subTest(task=task), patch.object(webhook, "fetch_task", return_value=task):
                    response = self.send()
                    self.assertEqual(response.status_code, 200)
                    self.assertEqual(response.json(), {"status": "ignored"})
            dispatch.assert_not_called()

    def test_supported_events_dispatch_with_stable_key_and_timestamp(self):
        with patch.object(webhook, "fetch_task", return_value=self.task), patch.object(webhook, "dispatch_workflow") as dispatch:
            for event in sorted(webhook.ACCEPTED_EVENTS):
                history = self.payload["history_items"]
                if event == "taskStatusUpdated":
                    history = [{**history[0], "field": "status", "after": {"status": "to do"}}]
                response = self.send({**self.payload, "event": event, "history_items": history})
                self.assertEqual(response.status_code, 202)
            keys = {call.args[2] for call in dispatch.call_args_list}
            self.assertEqual(len(keys), 1)
            call = dispatch.call_args
            self.assertEqual(call.args[1], "task123")
            self.assertRegex(call.args[2], r"^[a-f0-9]{64}$")
            self.assertEqual(call.args[3], self.payload["history_items"][0]["date"])

    def test_nonqueue_transition_is_ignored_even_if_current_task_later_became_urgent(self):
        with patch.object(webhook, "fetch_task") as fetch, patch.object(webhook, "dispatch_workflow") as dispatch:
            for event, field, after in (("taskPriorityUpdated", "priority", {"id": "2"}),
                                        ("taskPriorityUpdated", "priority", None),
                                        ("taskStatusUpdated", "status", {"status": "in progress"})):
                payload = {**self.payload, "event": event, "history_items": [
                    {**self.payload["history_items"][0], "field": field, "after": after}]}
                self.assertEqual(self.send(payload).json(), {"status": "ignored"})
            fetch.assert_not_called()
            dispatch.assert_not_called()

    def test_task_created_with_documented_status_and_creation_history(self):
        history = self.payload["history_items"][0]
        payload = {**self.payload, "event": "taskCreated", "history_items": [
            {**history, "field": "status", "after": {"status": "to do"}},
            {**history, "id": "creation-event", "field": "task_creation", "after": None},
        ]}
        with patch.object(webhook, "fetch_task", return_value=self.task), patch.object(webhook, "dispatch_workflow"):
            self.assertEqual(self.send(payload).status_code, 202)

    def test_provider_errors_are_retryable_and_never_echo_secrets_or_bodies(self):
        for failing_function in ("fetch_task", "dispatch_workflow"):
            with patch.object(webhook, "fetch_task", return_value=self.task), patch.object(webhook, failing_function, side_effect=webhook.DeliveryUnavailable("test-github-token task title")):
                response = self.send()
                self.assertEqual(response.status_code, 503)
                self.assertEqual(response.json(), {"status": "delivery_unavailable"})
                self.assertEqual(response.headers["retry-after"], "60")
                self.assertNotIn("token", response.text)

    def test_event_key_is_order_independent_and_changes_with_history_id(self):
        first = {**self.payload, "history_items": [{"id": "b"}, {"id": "a"}]}
        second = {**first, "history_items": list(reversed(first["history_items"]))}
        self.assertEqual(webhook.event_identity(first, b"one"), webhook.event_identity(second, b"two"))
        self.assertNotEqual(webhook.event_identity(first, b"one")[0],
                            webhook.event_identity(self.payload, b"one")[0])
        no_history = {**self.payload, "history_items": []}
        self.assertEqual(webhook.event_identity(no_history, b"same"), webhook.event_identity(no_history, b"same"))
        self.assertNotEqual(webhook.event_identity(no_history, b"same"), webhook.event_identity(no_history, b"other"))

    def test_future_or_invalid_event_dates_are_not_forwarded(self):
        for value in ("nonsense", "-1", True, "9" * 10000, str(int(time.time() * 1000) + 600000)):
            payload = {**self.payload, "history_items": [{"id": "a", "date": value}]}
            self.assertEqual(webhook.event_identity(payload, b"x")[1], "")

    def test_fetch_task_uses_fixed_host_and_disables_redirects(self):
        config = webhook.WebhookConfig.from_env()
        with patch.object(webhook.requests, "get", return_value=Mock(status_code=200, content=b"{}", json=lambda: self.task)) as get:
            self.assertEqual(webhook.fetch_task(config, "task123"), self.task)
            self.assertEqual(get.call_args.args[0], "https://api.clickup.com/api/v2/task/task123")
            self.assertFalse(get.call_args.kwargs["allow_redirects"])
        for status in (301, 401, 429, 500):
            with patch.object(webhook.requests, "get", return_value=Mock(status_code=status)):
                with self.assertRaises(webhook.DeliveryUnavailable):
                    webhook.fetch_task(config, "task123")
        with patch.object(webhook.requests, "get", return_value=Mock(status_code=404)):
            self.assertEqual(webhook.fetch_task(config, "task123"), {})

    def test_dispatch_contract_contains_no_task_text_or_clickup_credentials(self):
        config = webhook.WebhookConfig.from_env()
        with patch.object(webhook.requests, "post", return_value=Mock(status_code=204)) as post:
            webhook.dispatch_workflow(config, "task123", "a" * 64, "1700000000000")
            self.assertEqual(post.call_args.args[0], "https://api.github.com/repos/yasamkaradag34/retail-ai-tool/actions/workflows/clickup-agent.yml/dispatches")
            self.assertEqual(post.call_args.kwargs["json"], {"ref": "main", "inputs": {
                "mode": "run", "task_id": "task123", "wait_for_urgent": "true",
                "event_key": "a" * 64, "urgency_changed_at": "1700000000000"}})
            self.assertNotIn("test-clickup-token", repr(post.call_args))
            self.assertFalse(post.call_args.kwargs["allow_redirects"])
        with patch.object(webhook.requests, "post", side_effect=requests.Timeout("private error")):
            with self.assertRaises(webhook.DeliveryUnavailable):
                webhook.dispatch_workflow(config, "task123", "a" * 64, "")

    def test_repository_setting_cannot_escape_github_host(self):
        for repository in ("https://attacker.invalid", "owner/repo?x=1", "owner/repo/more", "owner/repo#x"):
            with patch.dict(os.environ, {"GITHUB_AGENT_REPOSITORY": repository}):
                self.assertEqual(self.send().status_code, 503)

    def test_production_app_mounts_signed_webhook_without_google_session(self):
        import main
        raw = json.dumps(self.payload).encode()
        signature = hmac.new(b"test-signing-secret", raw, hashlib.sha256).hexdigest()
        with TestClient(main.app, base_url="https://testserver") as client:
            with patch.object(webhook, "fetch_task", return_value=self.task), patch.object(webhook, "dispatch_workflow") as dispatch:
                response = client.post("/integrations/clickup/webhook", content=raw,
                                       headers={"X-Signature": signature})
                self.assertEqual(response.status_code, 202)
                dispatch.assert_called_once()


if __name__ == "__main__":
    unittest.main()
