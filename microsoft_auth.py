#!/usr/bin/env python3
"""Microsoft OAuth2 Helper for Outlook / Hotmail IMAP access."""
import os
import json
import time
import urllib.request
import urllib.parse
from pathlib import Path

MICROSOFT_CLIENT_ID = "9e5f94bc-e8a4-4e73-b8be-63364c29d753"  # Standard Public Client ID
DEVICE_CODE_URL = "https://login.microsoftonline.com/consumers/oauth2/v2.0/devicecode"
TOKEN_URL = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"
SCOPE = "https://outlook.office.com/IMAP.AccessAsUser.All offline_access"

TOKEN_FILE = Path(__file__).parent / ".outlook_tokens.json"


def get_stored_tokens() -> dict:
    if TOKEN_FILE.exists():
        try:
            with open(TOKEN_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_tokens(tokens: dict):
    with open(TOKEN_FILE, "w", encoding="utf-8") as f:
        json.dump(tokens, f, indent=2)


def refresh_access_token(account_email: str) -> str:
    tokens = get_stored_tokens()
    acc_data = tokens.get(account_email.lower(), {})
    refresh_token = acc_data.get("refresh_token")

    if not refresh_token:
        return ""

    data = urllib.parse.urlencode({
        "client_id": MICROSOFT_CLIENT_ID,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "scope": SCOPE
    }).encode("utf-8")

    try:
        req = urllib.request.Request(TOKEN_URL, data=data, headers={"Content-Type": "application/x-www-form-urlencoded"})
        with urllib.request.urlopen(req) as resp:
            token_resp = json.loads(resp.read().decode("utf-8"))
            access_token = token_resp.get("access_token")
            if token_resp.get("refresh_token"):
                acc_data["refresh_token"] = token_resp["refresh_token"]
            acc_data["access_token"] = access_token
            acc_data["expires_at"] = int(time.time()) + int(token_resp.get("expires_in", 3600))
            tokens[account_email.lower()] = acc_data
            save_tokens(tokens)
            return access_token
    except Exception:
        return ""


def get_microsoft_access_token(account_email: str, interactive: bool = True) -> str:
    """Get valid access token for Outlook/Hotmail account (refreshing or initiating device code flow)."""
    tokens = get_stored_tokens()
    acc_data = tokens.get(account_email.lower(), {})

    # Check if existing access token is still valid
    if acc_data.get("access_token") and acc_data.get("expires_at", 0) > time.time() + 60:
        return acc_data["access_token"]

    # Try refreshing
    if acc_data.get("refresh_token"):
        refreshed = refresh_access_token(account_email)
        if refreshed:
            return refreshed

    if not interactive:
        return ""

    # Initiate Device Code Flow
    data = urllib.parse.urlencode({
        "client_id": MICROSOFT_CLIENT_ID,
        "scope": SCOPE
    }).encode("utf-8")

    req = urllib.request.Request(DEVICE_CODE_URL, data=data, headers={"Content-Type": "application/x-www-form-urlencoded"})
    with urllib.request.urlopen(req) as resp:
        dev_resp = json.loads(resp.read().decode("utf-8"))

    user_code = dev_resp["user_code"]
    verification_uri = dev_resp.get("verification_uri", "https://www.microsoft.com/link")
    device_code = dev_resp["device_code"]
    interval = int(dev_resp.get("interval", 5))
    expires_in = int(dev_resp.get("expires_in", 900))

    print("\n" + "=" * 65)
    print(f"  [AUTH] MICROSOFT MODERN AUTHENTICATION -- {account_email}")
    print(f"  1. Open: {verification_uri}")
    print(f"  2. Enter Code: {user_code}")
    print(f"  3. Sign in to your Hotmail/Outlook account ({account_email})")
    print("=" * 65 + "\n", flush=True)

    # Poll for token
    start_time = time.time()
    while time.time() - start_time < expires_in:
        time.sleep(interval)
        poll_data = urllib.parse.urlencode({
            "client_id": MICROSOFT_CLIENT_ID,
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "device_code": device_code
        }).encode("utf-8")

        poll_req = urllib.request.Request(TOKEN_URL, data=poll_data, headers={"Content-Type": "application/x-www-form-urlencoded"})
        try:
            with urllib.request.urlopen(poll_req) as poll_resp:
                token_resp = json.loads(poll_resp.read().decode("utf-8"))
                access_token = token_resp["access_token"]
                refresh_token = token_resp.get("refresh_token")

                acc_data = {
                    "access_token": access_token,
                    "refresh_token": refresh_token,
                    "expires_at": int(time.time()) + int(token_resp.get("expires_in", 3600)),
                    "email": account_email
                }
                tokens[account_email.lower()] = acc_data
                save_tokens(tokens)
                print(f"[OK] Authentication Successful for {account_email}!\n")
                return access_token
        except urllib.error.HTTPError as err:
            err_body = err.read().decode("utf-8")
            if "authorization_pending" in err_body:
                continue
            elif "authorization_declined" in err_body:
                print("[!] Authorization was declined by user.")
                return ""
            elif "bad_verification_code" in err_body or "expired_token" in err_body:
                print("[!] Code expired.")
                return ""
        except Exception:
            continue

    return ""
