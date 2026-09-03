import datetime
import os
import sys
import traceback

ARGV_PATH = os.path.dirname(os.path.abspath(sys.argv[0]))


def error_log(msg):
    with open(os.path.join(ARGV_PATH, "error.log"), "w", encoding="utf-8") as f:
        timestamp = (
            datetime.datetime.now(datetime.timezone.utc).astimezone().isoformat()
        )
        f.write(f"Error Occurred at {timestamp}:\n")
        f.write(f"{msg}\n")


try:
    from main import main

    main()
except Exception:  # noqa: BLE001 -- record every unhandled application exception.
    error_log(traceback.format_exc())
    sys.exit(1)
