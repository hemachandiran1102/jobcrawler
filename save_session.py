"""
save_session.py
===============
Paste your li_at cookie from Chrome/Edge DevTools to save your LinkedIn session.
Run once, then the full crawler will work without any browser login needed.

How to get li_at:
  1. Open Chrome/Edge → go to linkedin.com
  2. Press F12 → Application → Cookies → https://www.linkedin.com
  3. Find 'li_at' → copy the Value
  4. Run: python save_session.py
  5. Paste the value when prompted
"""

import json
from pathlib import Path

SESSION_FILE = Path(__file__).parent / ".session.json"

print("=" * 60)
print("LinkedIn Session Setup")
print("=" * 60)
print()
print("Steps to get your li_at cookie:")
print("  1. Open Chrome/Edge")
print("  2. Go to: https://www.linkedin.com")
print("  3. Press F12 (DevTools)")
print("  4. Click 'Application' tab (top menu in DevTools)")
print("  5. Left panel: Cookies -> https://www.linkedin.com")
print("  6. Find row named 'li_at'")
print("  7. Click on it and copy the VALUE column")
print()

li_at = input("Paste your li_at cookie value here and press ENTER:\n> ").strip()

if not li_at:
    print("ERROR: No value entered. Exiting.")
    exit(1)

if len(li_at) < 50:
    print("WARNING: That value seems too short. Make sure you copied the full Value.")

session = {
    "li_at": li_at,
    "note": "Saved by save_session.py - paste li_at from Chrome/Edge DevTools"
}

SESSION_FILE.write_text(json.dumps(session, indent=2))
print()
print(f"[OK] Session saved to: {SESSION_FILE}")
print()
print("Now run the full crawler:")
print('  python full_crawl_to_word.py')
