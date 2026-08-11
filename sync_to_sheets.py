"""
sync_to_sheets.py — Standalone Google Sheets Uploader & Sync Tool
================================================================
Reads full_crawl_jobs.csv and uploads all job records to Google Sheets
in safe, optimized batches with automated retry and backoff.

Usage:
  python sync_to_sheets.py
  python sync_to_sheets.py --batch-size 100
  python sync_to_sheets.py --start-batch 3
"""

import os
import sys
import time
import json
import argparse
import requests
import pandas as pd
from pathlib import Path
from datetime import datetime

WORK_DIR = Path(__file__).parent.resolve()
MASTER_CSV_PATH = WORK_DIR / "full_crawl_jobs.csv"
CONFIG_PATH = WORK_DIR / "google_sheets_config.json"


def log(msg, level="INFO"):
    ts = datetime.now().strftime("%H:%M:%S")
    icon = {"INFO": "[i]", "OK": "[OK]", "WARN": "[!]", "ERROR": "[ERR]"}.get(level, "")
    safe = str(msg).encode("ascii", errors="replace").decode("ascii")
    print(f"[{ts}] {icon}  {safe}", flush=True)


def get_webhook_url() -> str:
    webhook_url = os.environ.get("GOOGLE_SHEETS_WEBHOOK_URL", "").strip()
    if not webhook_url and CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                webhook_url = cfg.get("webhook_url", "").strip()
        except Exception as e:
            log(f"Error reading {CONFIG_PATH}: {e}", "WARN")
    return webhook_url


def main():
    parser = argparse.ArgumentParser(description="Sync Master Job CSV to Google Sheets")
    parser.add_argument("--csv", type=str, default=str(MASTER_CSV_PATH),
                        help="Path to CSV file (default: full_crawl_jobs.csv)")
    parser.add_argument("--batch-size", type=int, default=100,
                        help="Number of jobs per sync batch (default: 100)")
    parser.add_argument("--timeout", type=int, default=60,
                        help="Request timeout in seconds per batch (default: 60)")
    parser.add_argument("--retries", type=int, default=3,
                        help="Max retry attempts per batch (default: 3)")
    parser.add_argument("--start-batch", type=int, default=1,
                        help="Start from a specific batch number (default: 1)")
    args = parser.parse_args()

    csv_path = Path(args.csv)
    if not csv_path.exists():
        log(f"CSV file not found: {csv_path}", "ERROR")
        sys.exit(1)

    webhook_url = get_webhook_url()
    if not webhook_url:
        log("Google Sheets Webhook URL not found. Set it in google_sheets_config.json or GOOGLE_SHEETS_WEBHOOK_URL.", "ERROR")
        sys.exit(1)

    log(f"Loading CSV: {csv_path}", "INFO")
    df = pd.read_csv(csv_path, on_bad_lines="skip")
    df.dropna(subset=["Job URL"], inplace=True)
    df.fillna("", inplace=True)

    clean_jobs = df.to_dict(orient="records")
    total_jobs = len(clean_jobs)
    chunk_size = max(10, args.batch_size)
    total_batches = (total_jobs + chunk_size - 1) // chunk_size

    log("=" * 65)
    log(f"🚀 Google Sheets Uploader — {total_jobs} total jobs")
    log(f"   Webhook     : {webhook_url[:45]}...")
    log(f"   Batch Size  : {chunk_size} jobs/batch ({total_batches} total batches)")
    log(f"   Timeout     : {args.timeout}s per batch")
    log(f"   Retries     : {args.retries} attempts on failure")
    if args.start_batch > 1:
        log(f"   Resuming at : Batch {args.start_batch}")
    log("=" * 65)

    success_batches = 0
    failed_batches = 0
    start_idx = (args.start_batch - 1) * chunk_size

    for i in range(start_idx, total_jobs, chunk_size):
        batch_num = i // chunk_size + 1
        chunk = clean_jobs[i:i + chunk_size]
        payload = {
            "action": "sync_jobs",
            "jobs": chunk
        }

        batch_synced = False
        for attempt in range(1, args.retries + 1):
            try:
                log(f"Syncing batch {batch_num}/{total_batches} ({len(chunk)} jobs){f' [attempt {attempt}/{args.retries}]' if attempt > 1 else ''} …", "INFO")
                res = requests.post(webhook_url, json=payload, timeout=args.timeout)
                if res.status_code == 200:
                    log(f"Batch {batch_num}/{total_batches} synced -> {res.text[:100]}", "OK")
                    batch_synced = True
                    success_batches += 1
                    break
                else:
                    log(f"Batch {batch_num} failed HTTP {res.status_code}: {res.text[:100]}", "WARN")
            except requests.exceptions.RequestException as req_err:
                log(f"Batch {batch_num} attempt {attempt} error: {req_err}", "WARN")

            if attempt < args.retries:
                backoff = attempt * 2
                time.sleep(backoff)

        if not batch_synced:
            failed_batches += 1
            log(f"Batch {batch_num} failed after {args.retries} attempts — continuing next batch.", "WARN")
        else:
            time.sleep(0.5)

    log("\n" + "=" * 65)
    if failed_batches == 0:
        log(f"DONE! All {total_jobs} jobs ({success_batches} batches) synced to Google Sheets.", "OK")
    else:
        log(f"COMPLETED with {failed_batches} failed batches ({success_batches}/{total_batches} succeeded).", "WARN")
    log("=" * 65)


if __name__ == "__main__":
    main()
