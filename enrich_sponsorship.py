#!/usr/bin/env python3
"""
enrich_sponsorship.py
======================
Scans all jobs across LinkedIn (full_crawl_jobs.csv) and Indeed/Glassdoor (indeed_glassdoor_jobs.csv)
and accurately tags each with Visa / Relocation sponsorship details:
  - "Visa Sponsored"
  - "Relocation Provided"
  - "No Sponsorship"
  - "Visa Mentioned"
  - "Not Specified"
"""

import os
import re
import pandas as pd
from pathlib import Path

WORK_DIR = Path(__file__).parent.resolve()
LINKEDIN_CSV = WORK_DIR / "full_crawl_jobs.csv"
INDEED_GD_CSV = WORK_DIR / "indeed_glassdoor_jobs.csv"

# Comprehensive Regex Patterns
VISA_SPONSORED_PATTERNS = [
    r'visa\s+sponsorship\s+(?:is\s+)?(?:available|provided|offered|supported|possible|eligible)',
    r'(?:offers?|provides?|includes?|with)\s+visa\s+sponsorship',
    r'visa\s+(?:support|assistance|allowance|processing)\s+(?:provided|available|offered|included)',
    r'will\s+(?:sponsor|provide|support)\s+(?:a\s+)?visa',
    r'relocation\s+(?:and|&)\s+visa\s+sponsorship(?:\s+provided)?',
    r'visa\s+(?:and|&)\s+relocation(?:\s+package|\s+support|\s+provided)?',
    r'(?:eu\s+)?blue\s*card(?:\s+eligible|\s+sponsorship|\s+support)?',
    r'kennismigrant',
    r'passeport\s+talent',
    r'(?:tier\s*2|skilled\s+worker)\s+visa(?:\s+sponsorship)?',
    r'critical\s+skills(?:\s+employment)?\s+permit',
    r'red-white-red\s*card',
    r'tech\s+visa',
    r'iqama\s+(?:transferable|transfer)',
    r'work\s+permit\s+(?:provided|sponsored|supported)',
]

RELOCATION_PATTERNS = [
    r'relocation\s+(?:package|assistance|support|allowance|bonus|budget|subsidy|provided|covered|offered)',
    r'(?:offers?|provides?|includes?|with)\s+relocation',
    r'full\s+relocation(?:\s+package|\s+support)?',
    r'international\s+relocation',
    r'help\s+with\s+relocation',
    r'relocation\s+to\s+[A-Za-z]+',
]

NO_SPONSORSHIP_PATTERNS = [
    r'no\s+visa\s+sponsorship',
    r'(?:cannot|unable\s+to|will\s+not|do\s+not)\s+sponsor',
    r'not\s+eligible\s+for\s+(?:visa\s+)?sponsorship',
    r'no\s+sponsorship\s+available',
    r'must\s+(?:already\s+)?have\s+(?:the\s+)?(?:right|authorization|permit)\s+to\s+work',
    r'must\s+be\s+(?:legally\s+)?authorized\s+to\s+work',
    r'valid\s+work\s+(?:permit|authorization|visa)\s+required',
    r'only\s+applicants\s+with\s+(?:existing\s+)?(?:valid\s+)?work',
    r'must\s+hold\s+(?:a\s+)?valid\s+work',
    r'must\s+have\s+(?:an?\s+)?existing\s+right\s+to\s+work',
    r'citizens?\s+or\s+permanent\s+residents?\s+only',
    r'eu\s+citizenship\s+required',
    r'no\s+work\s+permit\s+sponsorship',
]

GENERAL_VISA_PATTERNS = [
    r'\bvisa\b',
    r'\bsponsorship\b',
    r'\bwork\s+permit\b',
    r'\brelocation\b',
    r'\bimmigration\b',
    r'\bglobal\s+talent\b',
]

def determine_sponsorship(text: str) -> str:
    """Return specific sponsorship classification for a job description/title."""
    if not text or not isinstance(text, str):
        return "Not Specified"

    t = text.lower()

    # 1. Explicit No Sponsorship check first
    for pat in NO_SPONSORSHIP_PATTERNS:
        if re.search(pat, t, flags=re.IGNORECASE):
            return "No Sponsorship"

    # 2. Positive Visa Sponsorship
    for pat in VISA_SPONSORED_PATTERNS:
        if re.search(pat, t, flags=re.IGNORECASE):
            return "Visa Sponsored"

    # 3. Positive Relocation Provided
    for pat in RELOCATION_PATTERNS:
        if re.search(pat, t, flags=re.IGNORECASE):
            return "Relocation Provided"

    # 4. General mention
    for pat in GENERAL_VISA_PATTERNS:
        if re.search(pat, t, flags=re.IGNORECASE):
            return "Visa Mentioned"

    return "Not Specified"


def enrich_file(csv_path: Path, label: str):
    if not csv_path.exists():
        print(f"[{label}] File not found: {csv_path}")
        return

    print(f"[{label}] Loading {csv_path.name}...")
    df = pd.read_csv(csv_path, on_bad_lines="skip")
    print(f"[{label}] Loaded {len(df):,} jobs.")

    # Determine sponsorship per row
    sponsorship_col = []
    for _, row in df.iterrows():
        title = str(row.get("Job Title", ""))
        skills = str(row.get("Required Skills", ""))
        keyword = str(row.get("Search Keyword", ""))
        loc = str(row.get("Location", ""))
        notes = str(row.get("Notes", ""))
        existing = str(row.get("Visa Sponsorship Mentioned", ""))

        combined = f"{title} {skills} {keyword} {loc} {notes}"
        status = determine_sponsorship(combined)

        if status == "Not Specified" and existing.strip().lower() in ["yes", "true", "1"]:
            status = "Visa Sponsored"

        sponsorship_col.append(status)

    df["Visa Sponsorship Mentioned"] = sponsorship_col

    # Print distribution
    counts = df["Visa Sponsorship Mentioned"].value_counts()
    print(f"\n[{label}] Sponsorship Distribution:")
    for k, v in counts.items():
        print(f"   {k}: {v:,} ({v/len(df)*100:.1f}%)")

    df.to_csv(csv_path, index=False)
    print(f"[{label}] Successfully updated {csv_path.name}.\n")


if __name__ == "__main__":
    enrich_file(LINKEDIN_CSV, "LinkedIn")
    enrich_file(INDEED_GD_CSV, "Indeed / Glassdoor")
