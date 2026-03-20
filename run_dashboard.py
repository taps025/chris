from __future__ import annotations

import os
from pathlib import Path
import sys

from streamlit.web import cli as stcli


ROOT = Path(__file__).resolve().parent


if __name__ == "__main__":
    os.environ["CLIENT_DASHBOARD_STREAMLIT"] = "1"
    sys.argv = [
        "streamlit",
        "run",
        str(ROOT / "app.py"),
    ]
    raise SystemExit(stcli.main())
