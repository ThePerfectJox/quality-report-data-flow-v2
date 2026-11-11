"""Appends a saved report JSON file to the Excel master data, without contacting the server.

run_client.py already appends each response it receives. Use this to append a response
file again, for example after fixing a problem in the workbook.

Usage: python src/run_append_to_database.py [report.json]
Without a path, REPORT_JSON_PATH from .env is used.
"""

import logging
import sys

from client.append_to_database import (
    AppendDatabaseError,
    PostAppendError,
    append_data_to_excel_database,
)


def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    input_json_path = sys.argv[1] if len(sys.argv) > 1 else None
    try:
        append_data_to_excel_database(input_json_path)
    except PostAppendError as error:
        sys.exit(str(error))
    except AppendDatabaseError as error:
        sys.exit(f"Nothing was written to the Excel master data: {error}")


if __name__ == "__main__":
    main()
