"""Settings for the TCP server and client. TLS settings are in tls/settings.py.

IP addresses differ per device, so they can be set in each device's .env file.
"""
import socket

from managers.environment_manager import (
    ENV_CLIENT_BIND_IP,
    ENV_CLIENT_DEFAULT_SERVER_HOST,
    ENV_SERVER_LISTEN_HOST,
    EnvironmentManager,
)

_environment = EnvironmentManager()

# Used by both sides
SERVER_PORT = 5000
MAX_MESSAGE_BYTES = 1024 * 1024
REPORT_REQUEST_MESSAGE = "REQUEST QUALITY IPC REPORT DATA"

# Server
# The default, 0.0.0.0, listens on all network interfaces.
SERVER_LISTEN_HOST = _environment.get(ENV_SERVER_LISTEN_HOST, "0.0.0.0")
SERVER_SOCKET_TIMEOUT_SECONDS = (
    1  # How often waiting threads check whether the server is stopping.
)
SERVER_KEEP_ALIVE_INTERVAL_SECONDS = 1
SERVER_RECEIVE_BUFFER_BYTES = 4096
SERVER_SEND_TIMEOUT_SECONDS = 5  # A client that won't read a response is dropped.

# Client
CLIENT_DEFAULT_SERVER_HOST = _environment.get(
    ENV_CLIENT_DEFAULT_SERVER_HOST, "169.254.142.155"
)
# The local IP to connect from. Empty lets Windows choose, which can pick the wrong network.
CLIENT_BIND_IP = socket.gethostbyname(socket.gethostname())  # _environment.get(ENV_CLIENT_BIND_IP, "")
CLIENT_DEFAULT_MESSAGE = REPORT_REQUEST_MESSAGE
CLIENT_CONNECTION_TIMEOUT_SECONDS = 5
CLIENT_RESPONSE_TIMEOUT_SECONDS = 30  # The server may need time to collect the data.
CLIENT_RECEIVE_BUFFER_BYTES = 4096

# MySQL
MYSQL_DEFAULT_PORT = 3306
MYSQL_CONNECTION_TIMEOUT_SECONDS = 10
