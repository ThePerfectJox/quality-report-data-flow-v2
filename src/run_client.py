"""Runs the TCP client: sends a message to the server over TLS, prints its JSON response,
and appends the report to the Excel database.

Usage: python src/run_client.py [server-ip] [message]
"""

import sys
import json
import logging

from append_data_to_excel_database import (
    AppendDatabaseError,
    ArchiveCreationError,
    append_response_to_database,
)
from client.tcp_client import ClientBindIpError, send_request
from shared.address import format_address
from shared.json_serialization import to_json
from shared.settings import (
    CLIENT_DEFAULT_MESSAGE,
    CLIENT_DEFAULT_SERVER_HOST,
    SERVER_PORT,
)
from tls.contexts import create_client_tls_context
from tls.settings import TLS_CERTIFICATE_FILE


def get_server_host() -> str:
    """Returns the IP given on the command line, or the one from the settings if there is none."""
    return sys.argv[1] if len(sys.argv) > 1 else CLIENT_DEFAULT_SERVER_HOST


def get_message() -> str:
    """Returns the message given on the command line, or the default if there is none."""
    return sys.argv[2] if len(sys.argv) > 2 else CLIENT_DEFAULT_MESSAGE


def main():
    # Print any character safely, even if the output is redirected to a limited encoding.
    sys.stdout.reconfigure(errors="backslashreplace")
    # Shows the progress messages of the Excel database append.
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

    try:
        tls_context = create_client_tls_context()
    except OSError as error:
        sys.exit(f"Could not load {TLS_CERTIFICATE_FILE}: {error}\nCopy server.crt from the server device.")

    server_host = get_server_host()
    message = get_message()
    try:
        response = send_request(server_host, message, tls_context)
    except ClientBindIpError as error:
        sys.exit(str(error))
    except ValueError as error:
        sys.exit(f"Could not complete the request: {error}")
    except OSError as error:
        sys.exit(f"The request to {format_address(server_host, SERVER_PORT)} failed: {error}")
    print("Connection closed")

    try:
        append_response_to_database(response)
    except ArchiveCreationError as error:
        sys.exit(str(error))
    except AppendDatabaseError as error:
        sys.exit(f"Nothing was written to the Excel database: {error}")

    print("Success...")
    print("Report received, and appended")

if __name__ == "__main__":
    main()
