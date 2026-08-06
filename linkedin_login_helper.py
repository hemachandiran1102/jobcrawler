"""
linkedin_login_helper.py
========================
Run this ONCE to log into LinkedIn and save your session.
After completing, all crawlers will reuse the saved cookies.

Usage:
    python linkedin_login_helper.py

The browser will open — log in manually (email + password + any 2FA),
then press ENTER in this terminal once you see the LinkedIn feed.
"""

import os
import sys
import time
from pathlib import Path

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("ERROR: Run:  pip install playwright && playwright install chromium")
    sys.exit(1)

SESSION_DIR = Path(__file__).parent.resolve() / ".playwright-session"
print(f"Session will be saved to: {SESSION_DIR}")
SESSION_DIR.mkdir(parents=True, exist_ok=True)

with sync_playwright() as p:
    print("\nOpening browser...")
    browser = p.chromium.launch_persistent_context(
        user_data_dir=str(SESSION_DIR),
        headless=False,
        viewport={"width": 1440, "height": 900},
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/125.0.0.0 Safari/537.36"
        ),
    )

    page = browser.new_page()
    page.goto("https://www.linkedin.com/login", wait_until="domcontentloaded")

    print("\n" + "="*60)
    print("ACTION REQUIRED:")
    print("  1. Log into LinkedIn in the browser window that opened")
    print("  2. Complete any 2FA / verification if asked")
    print("  3. Wait until you see your LinkedIn FEED page")
    print("  4. Come back here and press ENTER")
    print("="*60)
    input("\nPress ENTER once you are on the LinkedIn feed page: ")

    # Verify login
    current_url = page.url
    print(f"\nCurrent URL: {current_url}")

    cookies = browser.cookies(["https://www.linkedin.com"])
    cookie_names = [c["name"] for c in cookies]
    has_li_at = "li_at" in cookie_names

    print(f"Cookies found: {cookie_names}")
    print(f"li_at present: {has_li_at}")

    if has_li_at:
        print("\n[OK] Session saved successfully!")
        print(f"     Location: {SESSION_DIR}")
        print("\nYou can now run the full crawler:")
        print('  python "C:\\Users\\hemac\\Desktop\\JobCrawler\\full_crawl_to_word.py"')
    else:
        print("\n[!] WARNING: li_at cookie not found.")
        print("    Make sure you are fully logged into LinkedIn (on the feed page).")
        print("    Try running this script again.")

    browser.close()
