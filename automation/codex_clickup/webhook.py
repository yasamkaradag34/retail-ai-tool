"""Signed ClickUp events dispatch work to GitHub; no agent runs in the web app.

Delivery is at-least-once. The event key is stable across ClickUp retries, while
the workflow must recheck eligibility after its delay and serialize task claims.
"""

from __future__ import annotations

from dataclasses import dataclass
import hashlib
import hmac
import json
import os
import re
import time
from typing import Any

from fastapi import APIRouter, Request
from fastapi.responses import JSONResponse
import requests
from starlette.concurrency import run_in_threadpool


router = APIRouter()
MAX_BODY_BYTES = 1024 * 1024
ACCEPTED_EVENTS = {"taskPriorityUpdated", "taskStatusUpdated", "taskCreated"}
ID_PATTERN = re.compile(r"^[A-Za-z0-9_-]{1,128}$")
REPOSITORY_PATTERN = re.compile(r"^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$")


@dataclass(frozen=True)
class WebhookConfig:
    enabled: bool
    secret: str
    clickup_token: str
    list_id: str
    github_token: str
    repository: str
    webhook_id: str
    queue_status: str

    @classmethod
    def from_env(cls) -> "WebhookConfig":
        return cls(
            enabled=os.getenv("CLICKUP_WEBHOOK_ENABLED", "false").lower() == "true",
            secret=os.getenv("CLICKUP_WEBHOOK_SECRET", "").strip(),
            clickup_token=os.getenv("CLICKUP_API_TOKEN", "").strip(),
            list_id=os.getenv("CLICKUP_LIST_ID", "").strip(),
            github_token=os.getenv("GITHUB_AGENT_DISPATCH_TOKEN", "").strip(),
            repository=os.getenv("GITHUB_AGENT_REPOSITORY", "yasamkaradag34/retail-ai-tool").strip(),
            webhook_id=os.getenv("CLICKUP_WEBHOOK_ID", "").strip(),
            queue_status=os.getenv("CLICKUP_QUEUE_STATUS", "to do").strip().casefold(),
        )

    def ready(self) -> bool:
        return bool(self.secret and self.clickup_token and self.github_token
                    and re.fullmatch(r"[0-9]+", self.list_id)
                    and REPOSITORY_PATTERN.fullmatch(self.repository)
                    and self.queue_status)


class DeliveryUnavailable(Exception):
    """A sanitized upstream error; provider bodies never enter logs/responses."""


def valid_signature(raw: bytes, signature: str, secret: str) -> bool:
    if not re.fullmatch(r"[a-fA-F0-9]{64}", signature):
        return False
    expected = hmac.new(secret.encode(), raw, hashlib.sha256).hexdigest()
    return hmac.compare_digest(expected, signature.lower())


def event_identity(payload: dict[str, Any], raw: bytes) -> tuple[str, str]:
    """Hash official webhook/history IDs; retries without history hash raw bytes."""
    webhook_id = payload["webhook_id"]
    history = payload.get("history_items", [])
    ids = sorted({str(item["id"]) for item in history
                  if isinstance(item, dict) and isinstance(item.get("id"), (str, int))
                  and not isinstance(item.get("id"), bool)}) if isinstance(history, list) else []
    identity = [f"{webhook_id}:{item_id}" for item_id in ids]
    if not identity:
        identity = [f"{webhook_id}:{hashlib.sha256(raw).hexdigest()}"]
    dates = []
    for item in history if isinstance(history, list) else []:
        if not isinstance(item, dict):
            continue
        value = str(item.get("date", ""))
        if len(value) <= 16 and value.isdigit() and 0 < int(value) <= int(time.time() * 1000) + 300000:
            dates.append(int(value))
    changed_at = str(max(dates)) if dates else ""
    digest = hashlib.sha256(json.dumps(identity, separators=(",", ":")).encode()).hexdigest()
    return digest, changed_at


def fetch_task(config: WebhookConfig, task_id: str) -> dict[str, Any]:
    try:
        response = requests.get(
            f"https://api.clickup.com/api/v2/task/{task_id}",
            headers={"Authorization": config.clickup_token, "Accept": "application/json"},
            timeout=(3, 7), allow_redirects=False,
        )
        if response.status_code == 404:
            return {}
        if response.status_code != 200 or len(response.content) > 2 * MAX_BODY_BYTES:
            raise DeliveryUnavailable()
        task = response.json()
        if not isinstance(task, dict):
            raise DeliveryUnavailable()
        return task
    except (requests.RequestException, ValueError, TypeError) as error:
        raise DeliveryUnavailable() from error


def eligible_task(task: dict[str, Any], task_id: str, config: WebhookConfig) -> bool:
    status = task.get("status")
    priority = task.get("priority")
    task_list = task.get("list")
    return bool(
        str(task.get("id", "")) == task_id
        and isinstance(task_list, dict) and str(task_list.get("id", "")) == config.list_id
        and isinstance(status, dict) and str(status.get("status", "")).strip().casefold() == config.queue_status
        and isinstance(priority, dict) and not isinstance(priority.get("id"), bool)
        and str(priority.get("id", "")) == "1"
        and not task.get("archived", False)
    )


def relevant_transition(payload: dict[str, Any], config: WebhookConfig) -> bool:
    """A delayed non-queue event must not start work for a newer queue state."""
    history = payload.get("history_items")
    # ClickUp documents history as optional. The current task remains the
    # authoritative eligibility check when no history accompanies an event.
    if not isinstance(history, list) or not history or payload["event"] == "taskCreated":
        return True
    field = "priority" if payload["event"] == "taskPriorityUpdated" else "status"
    for item in history:
        if not isinstance(item, dict) or item.get("field") != field:
            continue
        after = item.get("after")
        if not isinstance(after, dict):
            continue
        if field == "priority" and not isinstance(after.get("id"), bool) and str(after.get("id", "")) == "1":
            return True
        if field == "status" and str(after.get("status", "")).strip().casefold() == config.queue_status:
            return True
    return False


def dispatch_workflow(config: WebhookConfig, task_id: str, event_key: str, changed_at: str) -> None:
    try:
        response = requests.post(
            f"https://api.github.com/repos/{config.repository}/actions/workflows/clickup-agent.yml/dispatches",
            headers={"Authorization": f"Bearer {config.github_token}",
                     "Accept": "application/vnd.github+json", "X-GitHub-Api-Version": "2022-11-28"},
            json={"ref": "main", "inputs": {
                "mode": "run", "task_id": task_id, "wait_for_urgent": "true",
                "event_key": event_key, "urgency_changed_at": changed_at,
            }},
            timeout=(3, 7), allow_redirects=False,
        )
        if response.status_code != 204:
            raise DeliveryUnavailable()
    except requests.RequestException as error:
        raise DeliveryUnavailable() from error


def deliver(config: WebhookConfig, payload: dict[str, Any], raw: bytes) -> bool:
    task_id = payload["task_id"]
    if not relevant_transition(payload, config):
        return False
    if not eligible_task(fetch_task(config, task_id), task_id, config):
        return False
    event_key, changed_at = event_identity(payload, raw)
    dispatch_workflow(config, task_id, event_key, changed_at)
    return True


@router.post("/integrations/clickup/webhook", include_in_schema=False)
async def clickup_webhook(request: Request):
    config = WebhookConfig.from_env()
    if not config.enabled or not config.ready():
        return JSONResponse({"status": "unavailable"}, status_code=503,
                            headers={"Retry-After": "60"})
    raw = bytearray()
    async for chunk in request.stream():
        raw.extend(chunk)
        if len(raw) > MAX_BODY_BYTES:
            return JSONResponse({"status": "payload_too_large"}, status_code=413)
    raw = bytes(raw)
    if not valid_signature(raw, request.headers.get("x-signature", ""), config.secret):
        return JSONResponse({"status": "invalid_signature"}, status_code=401)
    try:
        payload = json.loads(raw)
    except (ValueError, UnicodeDecodeError):
        return JSONResponse({"status": "invalid_payload"}, status_code=400)
    if not isinstance(payload, dict):
        return JSONResponse({"status": "invalid_payload"}, status_code=400)
    event = payload.get("event")
    if not isinstance(event, str) or event not in ACCEPTED_EVENTS:
        return JSONResponse({"status": "ignored"})
    task_id, webhook_id = payload.get("task_id"), payload.get("webhook_id")
    if (not isinstance(task_id, str) or not ID_PATTERN.fullmatch(task_id)
            or not isinstance(webhook_id, str) or not ID_PATTERN.fullmatch(webhook_id)):
        return JSONResponse({"status": "invalid_payload"}, status_code=400)
    if config.webhook_id and webhook_id != config.webhook_id:
        return JSONResponse({"status": "invalid_webhook"}, status_code=401)
    try:
        accepted = await run_in_threadpool(deliver, config, payload, raw)
    except DeliveryUnavailable:
        return JSONResponse({"status": "delivery_unavailable"}, status_code=503,
                            headers={"Retry-After": "60"})
    return JSONResponse({"status": "accepted" if accepted else "ignored"},
                        status_code=202 if accepted else 200)
