from __future__ import annotations

import threading
import time
import webbrowser

from app.server import run


def open_browser() -> None:
    time.sleep(1)
    webbrowser.open("http://127.0.0.1:8000")


if __name__ == "__main__":
    threading.Thread(target=open_browser, daemon=True).start()
    run(host="127.0.0.1", port=8000)

