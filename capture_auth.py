#!/usr/bin/env python3
"""One-time (per session expiry) login capture for harlo.gambitco.io.

Opens a visible browser on the load-confirmation page. Log in there, then
come back to this terminal and press Enter. Cookies and localStorage are
saved to .auth/harlo.json (gitignored) and reused by check_loads.py.
"""

import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

BASE_URL = "https://harlo.gambitco.io"
LOGIN_PAGE = f"{BASE_URL}/apps/load-confirmation"
AUTH_FILE = Path(__file__).parent / ".auth" / "harlo.json"


def main() -> int:
    AUTH_FILE.parent.mkdir(exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        page.goto(LOGIN_PAGE)
        print(f"A browser window is open at {LOGIN_PAGE}")
        print("Log in there (if prompted), wait until the load-confirmation")
        print("page is fully loaded, then press Enter here to save the session.")
        input("Press Enter when logged in... ")
        context.storage_state(path=str(AUTH_FILE))
        browser.close()
    print(f"Session saved to {AUTH_FILE}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
