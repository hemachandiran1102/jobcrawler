#!/usr/bin/env python3
"""
══════════════════════════════════════════════════════════════════════
 Job Compass — Master Unified Crawler & Automation Engine
 (master_crawler.py / run_all_crawlers.py)
══════════════════════════════════════════════════════════════════════
 Combines all crawling, triage, and synchronization operations into
 a single unified entry point:
   1. LinkedIn Crawl (full_crawl_to_word.py)
   2. Indeed & Glassdoor Crawl (crawl_indeed_glassdoor.py)
   3. Visa & Relocation Sponsorship Enrichment (enrich_sponsorship.py)
   4. Multi-Lingual Email Inbound Pipeline & Triage (email_triage.py)
   5. Google Sheets & Master Excel/CSV Sync (sync_to_sheets.py)
"""

import os
import sys
import time
import argparse
import subprocess
from datetime import datetime
from pathlib import Path

# Fix Windows console encoding for UTF-8 and emojis
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

WORK_DIR = Path(__file__).parent.resolve()
sys.path.insert(0, str(WORK_DIR))

# ── ANSI Color Formatting ──
GREEN = "\033[92m"
BLUE = "\033[94m"
CYAN = "\033[96m"
YELLOW = "\033[93m"
RED = "\033[91m"
MAGENTA = "\033[95m"
BOLD = "\033[1m"
DIM = "\033[2m"
RESET = "\033[0m"

def print_banner():
    print(f"\n{CYAN}{BOLD}{'═' * 70}{RESET}")
    print(f"{CYAN}{BOLD}   🧭 JOB COMPASS — UNIFIED MASTER CRAWLER & TRIAGE ENGINE{RESET}")
    print(f"{DIM}   Automated LinkedIn, Indeed, Glassdoor & Multi-Lingual Email Pipeline{RESET}")
    print(f"{CYAN}{BOLD}{'═' * 70}{RESET}\n")

def log(stage: str, msg: str, level: str = "INFO"):
    t = datetime.now().strftime("%H:%M:%S")
    color = GREEN if level == "OK" else (YELLOW if level == "WARN" else (RED if level == "ERR" else BLUE))
    print(f"[{t}] {color}[{stage}]{RESET} {msg}")

def run_step(step_name: str, cmd_args: list, critical: bool = False) -> bool:
    """Run a sub-process step and stream output cleanly."""
    log("STEP", f"Starting {step_name}...", "INFO")
    t0 = time.time()
    try:
        res = subprocess.run([sys.executable] + cmd_args, cwd=str(WORK_DIR), check=False)
        dur = time.time() - t0
        if res.returncode == 0:
            log("DONE", f"{step_name} completed in {dur:.1f}s", "OK")
            return True
        else:
            log("FAIL", f"{step_name} exited with code {res.returncode} ({dur:.1f}s)", "WARN" if not critical else "ERR")
            return False
    except Exception as e:
        log("ERR", f"Error executing {step_name}: {e}", "ERR")
        return False

def get_file_stats():
    """Gather count statistics from master CSV and JSON files."""
    stats = {
        "linkedin_jobs": 0,
        "indeed_gd_jobs": 0,
        "total_jobs": 0,
        "visa_sponsored": 0,
        "relocation_jobs": 0,
        "email_next_steps": 0,
        "interview_invites": 0,
        "assessments": 0,
        "recruiter_outreach": 0
    }
    
    # 1. full_crawl_jobs.csv (LinkedIn & Combined)
    lk_path = WORK_DIR / "full_crawl_jobs.csv"
    if lk_path.exists():
        try:
            import pandas as pd
            df = pd.read_csv(lk_path, on_bad_lines="skip")
            stats["total_jobs"] = len(df)
            if "Source" in df.columns:
                stats["linkedin_jobs"] = len(df[df["Source"].str.contains("LinkedIn|linkedin", case=False, na=True)])
            else:
                stats["linkedin_jobs"] = len(df)
            if "Visa Sponsorship Mentioned" in df.columns:
                stats["visa_sponsored"] = len(df[df["Visa Sponsorship Mentioned"] == "Visa Sponsored"])
                stats["relocation_jobs"] = len(df[df["Visa Sponsorship Mentioned"] == "Relocation Provided"])
        except Exception:
            pass

    # 2. indeed_glassdoor_jobs.csv
    ig_path = WORK_DIR / "indeed_glassdoor_jobs.csv"
    if ig_path.exists():
        try:
            import pandas as pd
            df_ig = pd.read_csv(ig_path, on_bad_lines="skip")
            stats["indeed_gd_jobs"] = len(df_ig)
        except Exception:
            pass

    # 3. interview_pipeline.json
    pipe_path = WORK_DIR / "interview_pipeline.json"
    if pipe_path.exists():
        try:
            import json
            with open(pipe_path, "r", encoding="utf-8") as f:
                pipe = json.load(f)
                stats["email_next_steps"] = len([d for d in pipe if d.get("Is Next Step") in [True, "True", "true"]])
                stats["interview_invites"] = len([d for d in pipe if d.get("Category") == "Interview Invitation"])
                stats["assessments"] = len([d for d in pipe if d.get("Category") == "Technical Assessment"])
                stats["recruiter_outreach"] = len([d for d in pipe if d.get("Category") == "Recruiter Outreach"])
        except Exception:
            pass

    return stats

def print_summary(stats: dict, duration: float):
    print(f"\n{GREEN}{BOLD}{'═' * 70}{RESET}")
    print(f"{GREEN}{BOLD}   🎉 MASTER CRAWL & TRIAGE PIPELINE COMPLETE ({duration:.1f}s){RESET}")
    print(f"{GREEN}{BOLD}{'═' * 70}{RESET}")
    print(f"   💼 Total Cumulative Opportunities : {BOLD}{stats['total_jobs']:,}{RESET}")
    print(f"      ├─ LinkedIn Opportunities       : {stats['linkedin_jobs']:,}")
    print(f"      └─ Indeed & Glassdoor           : {stats['indeed_gd_jobs']:,}")
    print(f"   🛂 Visa / Relocation Sponsorship   :")
    print(f"      ├─ 🛂 Explicit Visa Sponsored   : {stats['visa_sponsored']:,}")
    print(f"      └─ ✈️ Relocation Package        : {stats['relocation_jobs']:,}")
    print(f"   📬 Inbound Email Next Steps Ready  : {BOLD}{stats['email_next_steps']}{RESET}")
    print(f"      ├─ 🟢 Interview Invitations     : {stats['interview_invites']}")
    print(f"      ├─ 🔵 Technical Assessments     : {stats['assessments']}")
    print(f"      └─ 🟠 Recruiter Direct Outreach : {stats['recruiter_outreach']}")
    print(f"{'─' * 70}")
    print(f"   📊 Master Files & Dashboards:")
    print(f"      ├─ LinkedIn Board         : index.html")
    print(f"      ├─ Indeed & Glassdoor     : indeed_glassdoor.html")
    print(f"      ├─ Inbound Email Pipeline : inbound_pipeline.html")
    print(f"      ├─ Master Excel (Jobs)    : full_crawl_jobs.xlsx")
    print(f"      └─ Master Excel (Triage)  : interview_pipeline.xlsx")
    print(f"{GREEN}{BOLD}{'═' * 70}{RESET}\n")

def main():
    parser = argparse.ArgumentParser(description="Master Unified Crawler & Email Triage Engine")
    parser.add_argument("--all", action="store_true", help="Run complete pipeline (default if no specific mode selected)")
    parser.add_argument("--jobs-only", action="store_true", help="Run only job crawlers (LinkedIn + Indeed/Glassdoor + Sponsorship enrichment)")
    parser.add_argument("--linkedin-only", action="store_true", help="Run only LinkedIn crawl")
    parser.add_argument("--indeed-glassdoor-only", action="store_true", help="Run only Indeed & Glassdoor crawl")
    parser.add_argument("--email-only", action="store_true", help="Run only Email Triage pipeline")
    parser.add_argument("--sync-only", action="store_true", help="Run only Google Sheets sync")
    parser.add_argument("--headless", action="store_true", default=True, help="Run browser crawlers in headless mode (default: True)")
    parser.add_argument("--no-headless", dest="headless", action="store_false", help="Run browser crawlers with visible UI")
    parser.add_argument("--days", type=int, default=30, help="Days back to scan for emails / jobs (default: 30)")
    parser.add_argument("--max", type=int, default=200, help="Max items per account / search keyword (default: 200)")

    args = parser.parse_args()
    print_banner()
    t_start = time.time()

    # Determine execution scope
    is_specific = args.jobs_only or args.linkedin_only or args.indeed_glassdoor_only or args.email_only or args.sync_only
    if not is_specific:
        # Default: run all
        run_linkedin = True
        run_indeed_gd = True
        run_sponsorship = True
        run_email = True
        run_sheets = True
    else:
        run_linkedin = bool(args.linkedin_only or args.jobs_only)
        run_indeed_gd = bool(args.indeed_glassdoor_only or args.jobs_only)
        run_sponsorship = bool(run_linkedin or run_indeed_gd)
        run_email = bool(args.email_only)
        run_sheets = bool(args.sync_only or args.jobs_only)

    # ── STAGE 1: LinkedIn Crawler ──
    if run_linkedin:
        print(f"\n{MAGENTA}{BOLD}▶ STAGE 1: LINKEDIN CRAWLER{RESET}")
        lk_args = ["full_crawl_to_word.py", "--days", str(args.days)]
        if not args.headless:
            lk_args.append("--no-headless")
        run_step("LinkedIn Crawl & Export", lk_args)

    # ── STAGE 2: Indeed & Glassdoor Crawler ──
    if run_indeed_gd:
        print(f"\n{MAGENTA}{BOLD}▶ STAGE 2: INDEED & GLASSDOOR CRAWLER{RESET}")
        ig_args = ["crawl_indeed_glassdoor.py", "--days", str(args.days)]
        if not args.headless:
            ig_args.append("--no-headless")
        run_step("Indeed & Glassdoor Crawl", ig_args)

    # ── STAGE 3: Visa Sponsorship Enrichment ──
    if run_linkedin or run_indeed_gd:
        print(f"\n{MAGENTA}{BOLD}▶ STAGE 3: VISA & RELOCATION SPONSORSHIP ENRICHMENT{RESET}")
        run_step("Sponsorship Engine", ["enrich_sponsorship.py"])

    # ── STAGE 4: Email Inbound Pipeline & Multi-Lingual Triage ──
    if run_email:
        print(f"\n{MAGENTA}{BOLD}▶ STAGE 4: MULTI-LINGUAL EMAIL TRIAGE & INBOUND PIPELINE{RESET}")
        email_args = ["email_triage.py", "--days", str(args.days), "--max", str(args.max)]
        run_step("Email Inbound Triage", email_args)

    # ── STAGE 5: Google Sheets & Excel Synchronization ──
    if run_sheets:
        print(f"\n{MAGENTA}{BOLD}▶ STAGE 5: GOOGLE SHEETS SYNCHRONIZATION{RESET}")
        sync_args = ["sync_to_sheets.py"]
        run_step("Google Sheets Sync", sync_args)

    # Final summary
    total_dur = time.time() - t_start
    stats = get_file_stats()
    print_summary(stats, total_dur)


if __name__ == "__main__":
    main()
