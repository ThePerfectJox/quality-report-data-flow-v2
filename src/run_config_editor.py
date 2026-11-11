"""Opens the config editor in the browser, to edit config/client-config.json.

Usage: python src/run_config_editor.py [port]
Keep this window open while editing. Press Ctrl+C to close the editor.
"""

import sys
import webbrowser

from client.config_editor import DEFAULT_EDITOR_PORT, EDITOR_HOST, create_config_editor_server


def main():
    sys.stdout.reconfigure(errors="backslashreplace")
    try:
        port = int(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_EDITOR_PORT
    except ValueError:
        sys.exit(f"The port must be a whole number, such as {DEFAULT_EDITOR_PORT}.")

    try:
        server = create_config_editor_server(port)
    except OSError as error:
        sys.exit(
            f"Could not start the config editor on port {port}: {error}\n"
            "It may already be open. Otherwise, choose another port: "
            "python src/run_config_editor.py 8766"
        )

    url = f"http://{EDITOR_HOST}:{port}/"
    print(f"Config editor: {url}")
    print("Keep this window open while editing. Press Ctrl+C to close the editor.")
    webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
        print("Config editor closed")


if __name__ == "__main__":
    main()
