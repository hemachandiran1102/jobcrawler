"""
crawl_indeed_glassdoor.py — Indeed & Glassdoor Job Crawler -> Excel, CSV & Word
================================================================================
Crawls DevOps, Cloud, SRE, and Platform Engineering jobs from Indeed and Glassdoor
across 22 target countries & tech hubs.

Features:
- Multi-tier scraping: Indeed Country Feeds + Playwright DOM scraper + Glassdoor scraper
- Complete data schema: Job Title, Company, Location, Country, Match Score, Visa Sponsorship,
  Required Skills, Tailored Resume, Source, Job URL, Posted Date, Crawl Date, Applied Status.
- Multi-sheet Excel workbook (All Jobs master sheet + Date Tabs) with professional formatting.
- Single Master CSV with automatic deduplication based on canonical URLs.
- Word document report generation.

Usage:
  python crawl_indeed_glassdoor.py --days 2
  python crawl_indeed_glassdoor.py --countries london toronto dubai amsterdam --days 2
  python crawl_indeed_glassdoor.py --sources indeed glassdoor --days 2
  python crawl_indeed_glassdoor.py --no-headless
"""

import os
import sys
import re
import time
import json
import random
import hashlib
import argparse
import urllib.parse
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime, timedelta
import requests
import pandas as pd

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
    HAS_PLAYWRIGHT = True
except ImportError:
    HAS_PLAYWRIGHT = False

try:
    import docx
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_EXCEL = True
except ImportError:
    HAS_EXCEL = False

# ══════════════════════════════════════════════════════════════════════
# PATHS & CONFIGURATION
# ══════════════════════════════════════════════════════════════════════
_SCRIPT_DIR = Path(__file__).parent.resolve()
WORK_DIR    = Path(os.environ.get("WORK_DIR", str(_SCRIPT_DIR)))
SESSION_DIR = WORK_DIR / ".indeed-glassdoor-session"

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
RUN_DIR   = WORK_DIR / f"crawl_indeed_glassdoor_{timestamp}"
CSV_PATH  = RUN_DIR / f"indeed_glassdoor_jobs_{timestamp}.csv"
DOCX_PATH = RUN_DIR / f"Indeed_Glassdoor_Jobs_{timestamp}.docx"
EXCEL_PATH = RUN_DIR / f"indeed_glassdoor_jobs_{timestamp}.xlsx"

# Single Master Spreadsheets
MASTER_CSV_PATH   = WORK_DIR / "indeed_glassdoor_jobs.csv"
MASTER_EXCEL_PATH = WORK_DIR / "indeed_glassdoor_jobs.xlsx"
CONFIG_PATH       = WORK_DIR / "google_sheets_config.json"


# ══════════════════════════════════════════════════════════════════════
# TARGET COUNTRIES (All 22 Markets & Resumes)
# ══════════════════════════════════════════════════════════════════════
TARGET_COUNTRIES = [
    # ── Tier 1: Primary English-Speaking / High Feasibility / Major Tech Hubs ──
    {"name": "United Kingdom",       "flag": "🇬🇧", "tier": 1, "indeed_domain": "uk.indeed.com",
     "resume": "Country_Resumes/Hemachandiran_Giri_CV_UNITED_KINGDOM.docx"},
    {"name": "Canada",               "flag": "🇨🇦", "tier": 1, "indeed_domain": "ca.indeed.com",
     "resume": "Country_Resumes/Hemachandiran_Giri_CV_CANADA.docx"},
    {"name": "United Arab Emirates", "flag": "🇦🇪", "tier": 1, "indeed_domain": "ae.indeed.com",
     "resume": "Country_Resumes/Hemachandiran_Giri_CV_UAE.docx"},
    {"name": "Saudi Arabia",         "flag": "🇸🇦", "tier": 1, "indeed_domain": "sa.indeed.com",
     "resume": "Country_Resumes/Hemachandiran_Giri_CV_SAUDI_ARABIA.docx"},
    {"name": "Qatar",                "flag": "🇶🇦", "tier": 1, "indeed_domain": "qa.indeed.com",
     "resume": "Country_Resumes/Hemachandiran_Giri_CV_QATAR.docx"},
    {"name": "Netherlands",          "flag": "🇳🇱", "tier": 1, "indeed_domain": "nl.indeed.com",
     "resume": "Country_Resumes/Hemachandiran_Giri_CV_NETHERLANDS.docx"},
    {"name": "Ireland",              "flag": "🇮🇪", "tier": 1, "indeed_domain": "ie.indeed.com",
     "resume": "Country_Resumes/Hemachandiran_Giri_CV_IRELAND.docx"},
    {"name": "Sweden",               "flag": "🇸🇪", "tier": 1, "indeed_domain": "se.indeed.com",
     "resume": "Country_Resumes/Hemachandiran_Giri_CV_SWEDEN.docx"},
    {"name": "Denmark",              "flag": "🇩🇰", "tier": 1, "indeed_domain": "dk.indeed.com",
     "resume": "Country_Resumes/Hemachandiran_Giri_CV_DENMARK.docx"},
    {"name": "Finland",              "flag": "🇫🇮", "tier": 1, "indeed_domain": "fi.indeed.com",
     "resume": "Country_Resumes/Hemachandiran_Giri_CV_FINLAND.docx"},
    {"name": "Australia",            "flag": "🇦🇺", "tier": 1, "indeed_domain": "au.indeed.com",
     "resume": "Country_Resumes/Hemachandiran_Giri_CV_AUSTRALIA.docx"},
    {"name": "Singapore",            "flag": "🇸🇬", "tier": 1, "indeed_domain": "sg.indeed.com",
     "resume": "Country_Resumes/Hemachandiran_Giri_CV_SINGAPORE.docx"},
    {"name": "New Zealand",          "flag": "🇳🇿", "tier": 1, "indeed_domain": "nz.indeed.com",
     "resume": "Country_Resumes/Hemachandiran_Giri_CV_NEW_ZEALAND.docx"},

    # ── Tier 2: Strong European & Arabian Tech Markets ──
    {"name": "Kuwait",               "flag": "🇰🇼", "tier": 2, "indeed_domain": "kw.indeed.com",
     "resume": "Country_Resumes/Hemachandiran_Giri_CV_KUWAIT.docx"},
    {"name": "Bahrain",              "flag": "🇧🇭", "tier": 2, "indeed_domain": "bh.indeed.com",
     "resume": "Country_Resumes/Hemachandiran_Giri_CV_BAHRAIN.docx"},
    {"name": "Oman",                 "flag": "🇴🇲", "tier": 2, "indeed_domain": "om.indeed.com",
     "resume": "Country_Resumes/Hemachandiran_Giri_CV_OMAN.docx"},
    {"name": "France",               "flag": "🇫🇷", "tier": 2, "indeed_domain": "fr.indeed.com",
     "resume": "Country_Resumes/Hemachandiran_Giri_CV_FRANCE.docx"},
    {"name": "Portugal",             "flag": "🇵🇹", "tier": 2, "indeed_domain": "pt.indeed.com",
     "resume": "Country_Resumes/Hemachandiran_Giri_CV_PORTUGAL.docx"},
    {"name": "Poland",               "flag": "🇵🇱", "tier": 2, "indeed_domain": "pl.indeed.com",
     "resume": "Country_Resumes/Hemachandiran_Giri_CV_POLAND.docx"},
    {"name": "Belgium",              "flag": "🇧🇪", "tier": 2, "indeed_domain": "be.indeed.com",
     "resume": "Country_Resumes/Hemachandiran_Giri_CV_BELGIUM.docx"},
    {"name": "Austria",              "flag": "🇦🇹", "tier": 2, "indeed_domain": "at.indeed.com",
     "resume": "Country_Resumes/Hemachandiran_Giri_CV_AUSTRIA.docx"},
    {"name": "Malaysia",             "flag": "🇲🇾", "tier": 2, "indeed_domain": "malaysia.indeed.com",
     "resume": "Country_Resumes/Hemachandiran_Giri_CV_MALAYSIA.docx"},
]

SEARCH_TERMS = [
    "DevOps Engineer",
    "Cloud Engineer",
    "Site Reliability Engineer",
    "Platform Engineer",
]

VISA_KEYWORDS = [
    "visa", "sponsorship", "relocation", "work permit", "blue card",
    "kennismigrant", "critical skills", "passeport talent", "tech visa",
    "red-white-red", "tier 2", "skilled worker", "immigration",
    "global skills", "express entry", "green visa", "golden visa", "iqama",
]

SKILL_TAGS = [
    "AWS", "Kubernetes", "Docker", "Terraform", "CI/CD", "Python", "Linux",
    "GCP", "Azure", "Ansible", "Golang", "Kafka", "Jenkins", "ArgoCD",
    "Helm", "GitOps", "Prometheus", "Grafana", "EKS", "ECS", "Vault",
    "Pulumi", "Bash", "GitHub Actions", "GitLab CI", "Datadog",
]


# ══════════════════════════════════════════════════════════════════════
# UTILITIES & NORMALIZATION
# ══════════════════════════════════════════════════════════════════════
def log(msg, level="INFO"):
    ts   = datetime.now().strftime("%H:%M:%S")
    icon = {"INFO": "[i]", "OK": "[OK]", "WARN": "[!]", "ERROR": "[ERR]"}.get(level, "")
    safe = str(msg).encode("ascii", errors="replace").decode("ascii")
    print(f"[{ts}] {icon}  {safe}", flush=True)


def normalize_url(url: str) -> str:
    """Extract canonical job URL for Indeed, Glassdoor, or other job boards."""
    u = str(url or "").strip()
    if not u or u.lower() in {"n/a", "none", "nan", "", "-"}:
        return ""

    # Indeed: https://www.indeed.com/viewjob?jk=1234567890abcdef
    m_indeed = re.search(r"[?&]jk=([a-zA-Z0-9]+)", u)
    if m_indeed:
        return f"https://www.indeed.com/viewjob?jk={m_indeed.group(1)}"

    # Glassdoor: https://www.glassdoor.com/job-listing/?jl=1234567890
    m_gd = re.search(r"(?:jl=|jobListingId=|job-listing/.*?jl=)(\d+)", u)
    if m_gd:
        return f"https://www.glassdoor.com/job-listing/?jl={m_gd.group(1)}"

    # LinkedIn: https://www.linkedin.com/jobs/view/1234567890
    m_li = re.search(r"/jobs/view/(?:[^\s/?#]*-)?(\d{6,14})", u)
    if m_li:
        return f"https://www.linkedin.com/jobs/view/{m_li.group(1)}"

    clean = u.split("#")[0]
    base = clean.split("?")[0].rstrip("/")
    if base.startswith("http://"):
        base = "https://" + base[7:]
    return base


def extract_skills(text: str) -> str:
    t = (text or "").lower()
    found = [s for s in SKILL_TAGS if s.lower() in t]
    return ", ".join(found[:10]) if found else "AWS, Kubernetes, CI/CD, Terraform, Linux"


def extract_visa(text: str) -> str:
    t = (text or "").lower()
    return "Yes" if any(k in t for k in VISA_KEYWORDS) else "No"


def calculate_match_score(title: str, desc: str) -> str:
    t = f"{title} {desc}".lower()
    score = 80
    if any(k in t for k in ["devops", "cloud", "reliability", "platform", "infrastructure", "kubernetes", "aws"]):
        score += 10
    if any(k in t for k in ["terraform", "ci/cd", "docker", "python", "linux", "gitops"]):
        score += 5
    if any(k in t for k in VISA_KEYWORDS):
        score += 5
    return f"{min(100, score)}%"


def extract_city(location_str: str) -> str:
    loc = str(location_str or "").strip()
    if not loc or loc.lower() in {"n/a", "unknown", "remote", "global"}:
        return "Unknown"
    parts = [p.strip() for p in loc.split(",") if p.strip()]
    city_candidate = parts[0]
    city_candidate = re.sub(r"\b(Greater|Area|Metropolitan|Region|City of)\b", "", city_candidate, flags=re.IGNORECASE).strip()
    return city_candidate or "Unknown"


def filter_target_countries(targets: list, countries_filter) -> list:
    if not countries_filter:
        return targets

    alias_map = {
        "uk": ["United Kingdom"], "london": ["United Kingdom"], "manchester": ["United Kingdom"], "edinburgh": ["United Kingdom"],
        "canada": ["Canada"], "toronto": ["Canada"], "vancouver": ["Canada"], "montreal": ["Canada"], "ottawa": ["Canada"],
        "uae": ["United Arab Emirates"], "dubai": ["United Arab Emirates"], "abu dhabi": ["United Arab Emirates"],
        "saudi": ["Saudi Arabia"], "saudi arabia": ["Saudi Arabia"], "riyadh": ["Saudi Arabia"], "jeddah": ["Saudi Arabia"],
        "qatar": ["Qatar"], "doha": ["Qatar"],
        "netherlands": ["Netherlands"], "amsterdam": ["Netherlands"], "rotterdam": ["Netherlands"], "utrecht": ["Netherlands"],
        "ireland": ["Ireland"], "dublin": ["Ireland"], "cork": ["Ireland"],
        "france": ["France"], "paris": ["France"], "lyon": ["France"],
        "australia": ["Australia"], "sydney": ["Australia"], "melbourne": ["Australia"], "brisbane": ["Australia"],
        "singapore": ["Singapore"], "malaysia": ["Malaysia"], "kuala lumpur": ["Malaysia"],
        "new zealand": ["New Zealand"], "auckland": ["New Zealand"], "wellington": ["New Zealand"],
        "poland": ["Poland"], "warsaw": ["Poland"], "krakow": ["Poland"],
        "portugal": ["Portugal"], "lisbon": ["Portugal"], "porto": ["Portugal"],
        "germany": ["Germany"], "berlin": ["Germany"], "munich": ["Germany"],
        "sweden": ["Sweden"], "stockholm": ["Sweden"], "denmark": ["Denmark"], "copenhagen": ["Denmark"],
        "finland": ["Finland"], "helsinki": ["Finland"], "belgium": ["Belgium"], "brussels": ["Belgium"],
        "austria": ["Austria"], "vienna": ["Austria"], "kuwait": ["Kuwait"], "bahrain": ["Bahrain"], "oman": ["Oman"],
        "arabian": ["United Arab Emirates", "Saudi Arabia", "Qatar", "Kuwait", "Bahrain", "Oman"],
        "gcc": ["United Arab Emirates", "Saudi Arabia", "Qatar", "Kuwait", "Bahrain", "Oman"]
    }

    target_names = set()
    for f in countries_filter:
        fl = f.strip().lower()
        if fl in alias_map:
            for mapped in alias_map[fl]:
                target_names.add(mapped.lower())
        else:
            for c in targets:
                c_lower = c["name"].lower()
                if fl in c_lower or c_lower in fl:
                    target_names.add(c_lower)

    filtered = [c for c in targets if c["name"].lower() in target_names]
    return filtered or targets


def deduplicate_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Deduplicate dataframe based on normalized job URL, preserving metadata."""
    if df is None or df.empty:
        return pd.DataFrame()

    records = df.to_dict(orient="records")
    unique_map = {}
    dupes_removed = 0

    for rec in records:
        raw_url = str(rec.get("Job URL", "")).strip()
        norm_url = normalize_url(raw_url)

        if norm_url:
            key = f"URL:{norm_url}"
            rec["Job URL"] = norm_url
        else:
            comp = str(rec.get("Company", "")).strip().lower()
            title = str(rec.get("Job Title", "")).strip().lower()
            loc = str(rec.get("Location", rec.get("Country", ""))).strip().lower()
            key = f"TITLE:{comp}|{title}|{loc}"

        if key in unique_map:
            dupes_removed += 1
            existing = unique_map[key]
            # Preserve Applied Status
            if str(rec.get("Applied Status", "")).strip().lower() == "yes" or \
               str(rec.get("Applied", "")).strip().lower() == "yes":
                existing["Applied Status"] = "Yes"
                existing["Applied"] = "Yes"
            if rec.get("Notes") and not existing.get("Notes"):
                existing["Notes"] = rec["Notes"]
            existing_date = str(existing.get("Crawl Date", existing.get("Date", "")))
            rec_date = str(rec.get("Crawl Date", rec.get("Date", "")))
            if rec_date and (not existing_date or rec_date < existing_date):
                existing["Crawl Date"] = rec_date
                existing["Date"] = rec_date
            for k, v in rec.items():
                if v and (not existing.get(k) or str(existing.get(k)).strip() in {"", "N/A", "Unknown"}):
                    existing[k] = v
        else:
            unique_map[key] = rec

    if dupes_removed > 0:
        log(f"Removed {dupes_removed} duplicate job records.", "INFO")

    return pd.DataFrame(list(unique_map.values()))


# ══════════════════════════════════════════════════════════════════════
# INDEED SCRAPER (Country RSS Feeds + Web Fallback)
# ══════════════════════════════════════════════════════════════════════
def fetch_indeed_rss(country: dict, keyword: str, days: int = 2) -> list:
    """Fetch Indeed jobs via fast XML RSS endpoints."""
    domain = country.get("indeed_domain", "www.indeed.com")
    kw_enc = urllib.parse.quote_plus(keyword)
    loc_enc = urllib.parse.quote_plus(country["name"])
    
    url = f"https://{domain}/rss?q={kw_enc}&l={loc_enc}&fromage={days}&sort=date"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
        "Accept": "application/rss+xml, text/xml, application/xml;q=0.9, */*;q=0.8"
    }

    jobs = []
    try:
        resp = requests.get(url, headers=headers, timeout=15)
        if resp.status_code == 200 and resp.text:
            root = ET.fromstring(resp.content)
            items = root.findall(".//item")
            for item in items:
                title_elem = item.find("title")
                link_elem = item.find("link")
                desc_elem = item.find("description")
                pub_elem = item.find("pubDate")
                source_elem = item.find("source")

                title_raw = title_elem.text if title_elem is not None and title_elem.text else ""
                link_raw = link_elem.text if link_elem is not None and link_elem.text else ""
                desc_raw = desc_elem.text if desc_elem is not None and desc_elem.text else ""
                pub_raw = pub_elem.text if pub_elem is not None and pub_elem.text else ""
                comp_raw = source_elem.text if source_elem is not None and source_elem.text else ""

                if not comp_raw and " - " in title_raw:
                    parts = title_raw.rsplit(" - ", 1)
                    title_clean = parts[0].strip()
                    comp_raw = parts[1].strip()
                else:
                    title_clean = title_raw.strip()

                if not title_clean or not link_raw:
                    continue

                posted_date = datetime.now().strftime("%Y-%m-%d")
                if pub_raw:
                    try:
                        dt = pd.to_datetime(pub_raw)
                        posted_date = dt.strftime("%Y-%m-%d")
                    except Exception:
                        pass

                clean_job_url = normalize_url(link_raw)
                jobs.append({
                    "Job Title": title_clean,
                    "Company": comp_raw or "Company not stated",
                    "Location": country["name"],
                    "Country": country["name"],
                    "Flag": country["flag"],
                    "Tier": country["tier"],
                    "Search Keyword": keyword,
                    "Posted Date": posted_date,
                    "Crawl Date": datetime.now().strftime("%Y-%m-%d"),
                    "Source": "Indeed",
                    "Easy Apply": "Yes" if "easy apply" in desc_raw.lower() else "No",
                    "Remote / Workplace": "Remote" if "remote" in f"{title_clean} {desc_raw}".lower() else "On-site / Hybrid",
                    "Match Score": calculate_match_score(title_clean, desc_raw),
                    "Visa Sponsorship Mentioned": extract_visa(desc_raw),
                    "Required Skills": extract_skills(f"{title_clean} {desc_raw}"),
                    "Resume File Path": country["resume"],
                    "Applied Status": "No",
                    "Notes": "",
                    "Job URL": clean_job_url
                })
    except Exception as e:
        log(f"    Indeed RSS warning ({country['name']} - '{keyword}'): {e}", "WARN")

    return jobs


def crawl_indeed_playwright(page, country: dict, keyword: str, max_jobs: int = 50, days: int = 2) -> list:
    """Fallback: Scrape Indeed jobs via Playwright browser."""
    domain = country.get("indeed_domain", "www.indeed.com")
    kw_enc = urllib.parse.quote_plus(keyword)
    loc_enc = urllib.parse.quote_plus(country["name"])
    
    url = f"https://{domain}/jobs?q={kw_enc}&l={loc_enc}&fromage={days}&sort=date"
    jobs = []

    try:
        page.goto(url, wait_until="domcontentloaded", timeout=15_000)
        time.sleep(random.uniform(1.2, 2.0))

        cards = page.query_selector_all("div.job_seen_beacon, div.cardOutline, td.resultContent")
        for card in cards[:max_jobs]:
            try:
                title_el = card.query_selector("h2.jobTitle, a[data-jk]")
                comp_el = card.query_selector("span[data-testid='company-name'], span.companyName")
                loc_el = card.query_selector("div[data-testid='text-location'], div.companyLocation")
                link_el = card.query_selector("a[data-jk], a.jcs-JobTitle")

                title = title_el.inner_text().strip() if title_el else ""
                company = comp_el.inner_text().strip() if comp_el else "Company not stated"
                location = loc_el.inner_text().strip() if loc_el else country["name"]
                
                href = ""
                if link_el:
                    href = link_el.get_attribute("href") or ""
                    if href.startswith("/"):
                        href = f"https://{domain}{href}"

                jk = card.get_attribute("data-jk") or (link_el.get_attribute("data-jk") if link_el else "")
                if jk:
                    href = f"https://www.indeed.com/viewjob?jk={jk}"

                if not title or not href:
                    continue

                jobs.append({
                    "Job Title": title,
                    "Company": company,
                    "Location": location,
                    "Country": country["name"],
                    "Flag": country["flag"],
                    "Tier": country["tier"],
                    "Search Keyword": keyword,
                    "Posted Date": datetime.now().strftime("%Y-%m-%d"),
                    "Crawl Date": datetime.now().strftime("%Y-%m-%d"),
                    "Source": "Indeed",
                    "Easy Apply": "No",
                    "Remote / Workplace": "Remote" if "remote" in f"{title} {location}".lower() else "On-site / Hybrid",
                    "Match Score": calculate_match_score(title, ""),
                    "Visa Sponsorship Mentioned": "No",
                    "Required Skills": extract_skills(title),
                    "Resume File Path": country["resume"],
                    "Applied Status": "No",
                    "Notes": "",
                    "Job URL": normalize_url(href)
                })
            except Exception:
                continue
    except Exception as e:
        log(f"    Indeed Playwright error ({country['name']} - '{keyword}'): {e}", "WARN")

    return jobs


# ══════════════════════════════════════════════════════════════════════
# GLASSDOOR SCRAPER (Playwright DOM & Search API)
# ══════════════════════════════════════════════════════════════════════
def crawl_glassdoor_playwright(page, country: dict, keyword: str, max_jobs: int = 50, days: int = 2) -> list:
    """Scrape Glassdoor jobs using Playwright persistent browser."""
    kw_enc = urllib.parse.quote_plus(keyword)
    loc_enc = urllib.parse.quote_plus(country["name"])
    
    url = f"https://www.glassdoor.com/Job/jobs.htm?sc.keyword={kw_enc}&locT=N&locKeyword={loc_enc}&fromAge={days}"
    jobs = []

    try:
        page.goto(url, wait_until="domcontentloaded", timeout=30_000)
        time.sleep(random.uniform(2.5, 4.0))

        cards = page.query_selector_all("li[data-test='jobListing'], div[data-test='job-card-wrapper'], div.jobCard")
        for card in cards[:max_jobs]:
            try:
                title_el = card.query_selector("a[data-test='job-title'], a.job-title")
                comp_el = card.query_selector("span[class*='EmployerName'], div[class*='EmployerProfile']")
                loc_el = card.query_selector("div[data-test='emp-location'], span[data-test='emp-location']")
                
                title = title_el.inner_text().strip() if title_el else ""
                company = comp_el.inner_text().strip() if comp_el else "Company not stated"
                location = loc_el.inner_text().strip() if loc_el else country["name"]

                href = ""
                if title_el:
                    href = title_el.get_attribute("href") or ""
                    if href.startswith("/"):
                        href = f"https://www.glassdoor.com{href}"

                jl = card.get_attribute("data-job-id") or card.get_attribute("data-id") or ""
                if jl:
                    href = f"https://www.glassdoor.com/job-listing/?jl={jl}"

                if not title or not href:
                    continue

                jobs.append({
                    "Job Title": title,
                    "Company": company,
                    "Location": location,
                    "Country": country["name"],
                    "Flag": country["flag"],
                    "Tier": country["tier"],
                    "Search Keyword": keyword,
                    "Posted Date": datetime.now().strftime("%Y-%m-%d"),
                    "Crawl Date": datetime.now().strftime("%Y-%m-%d"),
                    "Source": "Glassdoor",
                    "Easy Apply": "No",
                    "Remote / Workplace": "Remote" if "remote" in f"{title} {location}".lower() else "On-site / Hybrid",
                    "Match Score": calculate_match_score(title, ""),
                    "Visa Sponsorship Mentioned": "No",
                    "Required Skills": extract_skills(title),
                    "Resume File Path": country["resume"],
                    "Applied Status": "No",
                    "Notes": "",
                    "Job URL": normalize_url(href)
                })
            except Exception:
                continue
    except Exception as e:
        log(f"    Glassdoor scraper error ({country['name']} - '{keyword}'): {e}", "WARN")

    return jobs


# ══════════════════════════════════════════════════════════════════════
# WORD & EXCEL GENERATORS
# ══════════════════════════════════════════════════════════════════════
def save_multi_sheet_excel(df: pd.DataFrame):
    """Save clean multi-sheet Excel with 'All Jobs' master sheet + Date Tabs."""
    if not HAS_EXCEL or df is None or df.empty:
        return

    clean_df = deduplicate_dataframe(df)
    log("\nGenerating Indeed & Glassdoor multi-sheet Excel (Master + Date Tabs) …")
    wb = openpyxl.Workbook()
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    cols = ["#","Crawl Date","Source","Country","Flag","Tier","City","City Openings",
            "Location","Company","Job Title","Search Keyword","Posted Date",
            "Easy Apply","Remote / Workplace","Match Score","Visa Sponsorship",
            "Skills","Resume","Applied","Notes","Job URL"]

    hf   = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    hfnt = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    thin = Border(left=Side(style="thin"), right=Side(style="thin"),
                  top=Side(style="thin"),  bottom=Side(style="thin"))
    vf   = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

    def populate_sheet(ws, sub_df):
        for ci, col in enumerate(cols, 1):
            c = ws.cell(row=1, column=ci, value=col)
            c.fill = hf; c.font = hfnt
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin

        for ri, (_, job) in enumerate(sub_df.iterrows(), 2):
            vv = str(job.get("Visa Sponsorship Mentioned", "No"))
            vals = [
                ri-1, str(job.get("Crawl Date", job.get("Date", ""))), str(job.get("Source", "Indeed/Glassdoor")),
                str(job.get("Country","")), str(job.get("Flag","")), str(job.get("Tier","")),
                str(job.get("City","")), str(job.get("City Openings","")),
                str(job.get("Location","")), str(job.get("Company","")),
                str(job.get("Job Title","")), str(job.get("Search Keyword","")),
                str(job.get("Posted Date","")), str(job.get("Easy Apply","No")),
                str(job.get("Remote / Workplace","On-site / Hybrid")),
                str(job.get("Match Score","85%")), vv,
                str(job.get("Required Skills","")), str(job.get("Resume File Path","")),
                str(job.get("Applied Status","No")), str(job.get("Notes","")),
                str(job.get("Job URL",""))
            ]
            for ci, val in enumerate(vals, 1):
                c = ws.cell(row=ri, column=ci, value=val)
                c.font = Font(name="Calibri", size=9); c.border = thin
                if ci == 17 and vv.lower() == "yes":
                    c.fill = vf

        for i, w in enumerate([4,12,12,14,5,5,16,14,20,22,32,20,16,12,18,12,15,30,45,8,20,55], 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"
        if len(sub_df) > 0:
            ws.auto_filter.ref = ws.dimensions

    # 1. Master Sheet "All Jobs"
    ws_all = wb.create_sheet(title="All Jobs")
    populate_sheet(ws_all, clean_df)

    # 2. Date-based sheets
    date_col = "Crawl Date" if "Crawl Date" in clean_df.columns else ("Date" if "Date" in clean_df.columns else None)
    if date_col:
        unique_dates = sorted(clean_df[date_col].dropna().astype(str).str.split("T").str[0].unique(), reverse=True)
        for date_str in unique_dates:
            if not date_str or len(date_str) < 5 or date_str == "nan":
                continue
            sub = clean_df[clean_df[date_col].astype(str).str.startswith(date_str)]
            if not sub.empty:
                ws_date = wb.create_sheet(title=date_str[:31])
                populate_sheet(ws_date, sub)

    try:
        MASTER_EXCEL_PATH.parent.mkdir(parents=True, exist_ok=True)
        EXCEL_PATH.parent.mkdir(parents=True, exist_ok=True)
        wb.save(MASTER_EXCEL_PATH)
        wb.save(EXCEL_PATH)
        log(f"Master Excel saved → {MASTER_EXCEL_PATH} (All Jobs + Date Tabs)", "OK")
        log(f"Run Excel archived → {EXCEL_PATH}", "OK")
    except PermissionError:
        alt = WORK_DIR / f"indeed_glassdoor_jobs_{int(time.time())}.xlsx"
        wb.save(alt)
        log(f"Excel locked — saved to {alt}", "WARN")


def save_master_csv(all_jobs: list) -> pd.DataFrame:
    """Save single master CSV with deduplication and metadata preservation."""
    today_str = datetime.now().strftime("%Y-%m-%d")
    incoming_df = pd.DataFrame(all_jobs)

    if MASTER_CSV_PATH.exists():
        try:
            log(f"Loading existing Indeed/Glassdoor master sheet → {MASTER_CSV_PATH}", "INFO")
            master_df = pd.read_csv(MASTER_CSV_PATH, on_bad_lines="skip")
            combined_df = pd.concat([master_df, incoming_df], ignore_index=True)
        except Exception as e:
            log(f"Could not read existing master CSV ({e}) — creating new", "WARN")
            combined_df = incoming_df
    else:
        combined_df = incoming_df

    df = deduplicate_dataframe(combined_df)
    df["City"] = df["Location"].apply(extract_city)
    city_counts = df["City"].value_counts().to_dict()
    df["City Openings"] = df["City"].map(city_counts).fillna(0).astype(int)

    RUN_DIR.mkdir(parents=True, exist_ok=True)
    df.to_csv(MASTER_CSV_PATH, index=False)
    df.to_csv(CSV_PATH, index=False)
    log(f"Master CSV updated → {MASTER_CSV_PATH} ({len(df)} unique jobs maintaining single master sheet)", "OK")
    log(f"Run CSV archived  → {CSV_PATH}", "OK")
    return df


def generate_word_report(jobs_df: pd.DataFrame):
    """Generate professional Word Document report for Indeed & Glassdoor crawl."""
    if not HAS_DOCX or jobs_df.empty:
        return
    log("\nGenerating Indeed & Glassdoor Word document report …", "INFO")
    doc = Document()

    # Title
    title_p = doc.add_paragraph()
    r = title_p.add_run(f"Indeed & Glassdoor Job Intelligence Report")
    r.bold = True
    r.font.size = Pt(20)
    r.font.color.rgb = RGBColor(31, 78, 121)

    sub_p = doc.add_paragraph()
    sub_p.add_run(f"Generated on {datetime.now().strftime('%Y-%m-%d %H:%M')} · {len(jobs_df)} Total Roles across {jobs_df['Country'].nunique()} Countries")
    sub_p.runs[0].font.size = Pt(10)
    sub_p.runs[0].font.italic = True

    # Country summary table
    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdrs = ["Country", "Flag", "Roles Found", "Sources"]
    for i, h in enumerate(hdrs):
        cell = table.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True

    for country in TARGET_COUNTRIES:
        c_name = country["name"]
        sub = jobs_df[jobs_df["Country"].str.lower() == c_name.lower()]
        if not sub.empty:
            row = table.add_row().cells
            row[0].text = c_name
            row[1].text = country["flag"]
            row[2].text = str(len(sub))
            row[3].text = ", ".join(sub["Source"].unique())

    try:
        DOCX_PATH.parent.mkdir(parents=True, exist_ok=True)
        doc.save(DOCX_PATH)
        log(f"Word document saved → {DOCX_PATH}", "OK")
    except Exception as e:
        log(f"Could not save Word document: {e}", "WARN")


# ══════════════════════════════════════════════════════════════════════
# MAIN CRAWL PIPELINE
# ══════════════════════════════════════════════════════════════════════
def run_crawler(headless: bool, countries_filter, sources: list, max_per_kw: int, days: int) -> list:
    targets = filter_target_countries(TARGET_COUNTRIES, countries_filter)
    all_jobs = []

    log("=" * 65)
    log(f"🚀 Starting Indeed & Glassdoor Job Crawler")
    log(f"   Target Markets : {len(targets)} Countries")
    log(f"   Sources        : {', '.join(sources).upper()}")
    log(f"   Time Window    : Past {days} Days")
    log(f"   Headless Mode  : {headless}")
    log("=" * 65)

    # 1. Fast Indeed Country Feeds
    if "indeed" in sources:
        log("\n── Phase 1: Fast Indeed Extraction ──", "INFO")
        for country in targets:
            log(f"Scanning Indeed for {country['flag']} {country['name']} …")
            for kw in SEARCH_TERMS:
                rss_jobs = fetch_indeed_rss(country, kw, days=days)
                if rss_jobs:
                    log(f"   {country['flag']} {country['name']} | '{kw}' → {len(rss_jobs)} jobs via Indeed RSS", "OK")
                    all_jobs.extend(rss_jobs)
                time.sleep(random.uniform(0.5, 1.2))

    # 2. Playwright Browser Fallback / Glassdoor Scraper
    needs_browser = ("glassdoor" in sources) or (len(all_jobs) == 0 and HAS_PLAYWRIGHT)
    if needs_browser and HAS_PLAYWRIGHT:
        log("\n── Phase 2: Browser Scraper (Indeed & Glassdoor) ──", "INFO")
        SESSION_DIR.mkdir(parents=True, exist_ok=True)
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch_persistent_context(
                    user_data_dir=str(SESSION_DIR),
                    headless=headless,
                    args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-blink-features=AutomationControlled"],
                    user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
                    viewport={"width": 1440, "height": 900},
                    locale="en-US",
                )
                page = browser.new_page()

                for country in targets:
                    for kw in SEARCH_TERMS:
                        if "glassdoor" in sources:
                            log(f"   [Glassdoor] {country['flag']} {country['name']} | '{kw}' …")
                            gd_jobs = crawl_glassdoor_playwright(page, country, kw, max_jobs=max_per_kw, days=days)
                            if gd_jobs:
                                log(f"      + {len(gd_jobs)} jobs from Glassdoor", "OK")
                                all_jobs.extend(gd_jobs)
                        
                        if "indeed" in sources:
                            ind_jobs = crawl_indeed_playwright(page, country, kw, max_jobs=max_per_kw, days=days)
                            if ind_jobs:
                                log(f"      + {len(ind_jobs)} jobs from Indeed Web", "OK")
                                all_jobs.extend(ind_jobs)

                browser.close()
        except Exception as e:
            log(f"Browser scraping exception: {e}", "WARN")

    return all_jobs


def main():
    parser = argparse.ArgumentParser(description="Indeed & Glassdoor Job Crawler -> Excel/CSV/Word")
    parser.add_argument("--no-headless", action="store_true",
                        help="Show browser window (for CAPTCHA solving or visual verification)")
    parser.add_argument("--countries", nargs="*", default=None,
                        help="Country/City names to scrape (e.g. london toronto dubai amsterdam)")
    parser.add_argument("--sources", nargs="*", default=["indeed", "glassdoor"],
                        help="Sources to scrape: indeed, glassdoor (default: both)")
    parser.add_argument("--max-per-keyword", type=int, default=50,
                        help="Max jobs per keyword per source (default: 50)")
    parser.add_argument("--days", type=int, default=2,
                        help="Number of past days to crawl (default: 2)")
    args = parser.parse_args()

    headless = not args.no_headless
    sources = [s.lower() for s in args.sources]
    countries_filter = [c.lower() for c in args.countries] if args.countries else None

    all_jobs = run_crawler(headless, countries_filter, sources, args.max_per_keyword, args.days)

    if not all_jobs:
        if MASTER_CSV_PATH.exists():
            log("No new live jobs extracted — loading existing master database …", "WARN")
            jobs_df = pd.read_csv(MASTER_CSV_PATH, on_bad_lines="skip")
        else:
            log("No job records found to write.", "ERROR")
            return
    else:
        jobs_df = save_master_csv(all_jobs)

    if not jobs_df.empty:
        save_multi_sheet_excel(jobs_df)
        generate_word_report(jobs_df)

    log("\n" + "=" * 65)
    log("✅  INDEED & GLASSDOOR CRAWL COMPLETE", "OK")
    log(f"   Total Unique Jobs : {len(jobs_df)}")
    log(f"   📊 Master Excel   : {MASTER_EXCEL_PATH}")
    log(f"   📋 Master CSV     : {MASTER_CSV_PATH}")
    log(f"   📄 Word Report    : {DOCX_PATH}")
    log("=" * 65)


if __name__ == "__main__":
    main()
