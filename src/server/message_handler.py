"""Decides what the server does when a client sends a message. Reporting is in reporting.py."""

from datetime import datetime, timezone

from server.reporting import create_report, mark_report_sent
from shared.settings import REPORT_REQUEST_MESSAGE


def handle_client_message(message: str) -> dict:
    """Returns the data to send to the client as JSON."""
    response = {"request": message, "responded_at": datetime.now(timezone.utc)}
    if message != REPORT_REQUEST_MESSAGE:
        return {**response, "status": "error", "message": "Unknown request"}
    return {**response, "status": "ok", "report": create_report()}


def handle_response_sent(message: str, response: dict):
    """Runs after a response has been sent, so a report only counts once it has gone out."""
    if message == REPORT_REQUEST_MESSAGE:
        mark_report_sent(response["report"])
