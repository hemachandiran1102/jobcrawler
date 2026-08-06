# Job Compass — LinkedIn Job Crawler & Opportunity Board

> **Job Compass** is a high-performance LinkedIn job crawler, single master spreadsheet engine, and interactive dashboard system designed to automate job discovery across **14 countries**, track application statuses in real-time, and synchronize bi-directionally with **Google Sheets**.

---

## 🌟 Key Features

- ⚡ **Voyager API Crawler**: Direct JSON endpoint extraction — immune to DOM layout changes and 10x faster than traditional browser automation.
- 📊 **Single Master Spreadsheet**: Maintains a unified, deduplicated dataset ([full_crawl_jobs.csv](file:///c:/Users/hemac/Desktop/JobCrawler/full_crawl_jobs.csv) and [full_crawl_jobs.xlsx](file:///c:/Users/hemac/Desktop/JobCrawler/full_crawl_jobs.xlsx)) at your project root. Existing notes and `Applied Status` are automatically preserved across crawls.
- 📅 **Flexible Time-Window Filters**: Scrape past 24 hours (`24h` / `1d`), past week (`1w`), past month (`1m`), or all available postings.
- ☁️ **Real-Time Google Sheets Sync**: Ultra-fast batch streaming to Google Sheets with bi-directional status updates.
- 🖥️ **Interactive Web Dashboard ([index.html](file:///c:/Users/hemac/Desktop/JobCrawler/index.html))**: Dark mode dashboard featuring search, multi-field filtering, sorting, bookmarking, and 1-click `Mark Applied` toggles.
- 🔄 **Cross-Device Sync**: Mark a job as "Applied" on one computer and instantly view the status update in Google Sheets and on other devices.

---

## 📁 Project Structure

```
JobCrawler/
├── full_crawl_to_word.py      # Main Python crawler script
├── index.html                 # Interactive Job Compass web dashboard
├── app.js                     # Dashboard logic, IndexedDB persistence, Google Sheets sync
├── styles.css                 # Premium responsive CSS styling system
├── google_apps_script.js      # Apps Script backend for Google Sheets integration
├── google_sheets_config.json  # Configuration file holding the Google Webhook URL
├── full_crawl_jobs.csv        # Master CSV dataset (all cumulative jobs)
├── full_crawl_jobs.xlsx       # Master Excel spreadsheet (all cumulative jobs)
├── linkedin_login_helper.py   # Login session helper script
└── save_session.py            # Playwright session state extractor
```

---

## 🚀 Quick Start

### 1. Prerequisites & Installation

Ensure you have Python 3.10+ installed, then install the required dependencies:

```bash
pip install pandas requests openpyxl python-docx playwright
playwright install chromium
```

---

### 2. Running Job Crawls

Run the crawler with optional arguments to filter by time window, target countries, or job caps:

```bash
# Past 24 hours (Default)
python full_crawl_to_word.py --time-window 24h

# Past 1 week (7 Days)
python full_crawl_to_word.py --time-window 1w

# Past 1 month (30 Days) for specific countries
python full_crawl_to_word.py --time-window 1m --countries netherlands ireland

# Cap jobs per keyword
python full_crawl_to_word.py --max-per-keyword 50
```

#### CLI Options:
- `--time-window`: `24h` / `1d` (past 24 hrs), `1w` / `7d` (past week), `1m` / `30d` (past month), `all`.
- `--countries`: Space-separated country list (e.g. `netherlands ireland sweden`).
- `--max-per-keyword`: Maximum jobs per keyword per country.
- `--no-headless`: Show browser window (for initial LinkedIn login or CAPTCHA resolution).

---

### 3. Setting Up Google Sheets Integration

1. Open your Google Spreadsheet (or create a new one).
2. Go to **Extensions** → **Apps Script**.
3. Copy all contents of [`google_apps_script.js`](file:///c:/Users/hemac/Desktop/JobCrawler/google_apps_script.js) into the editor.
4. Set your Spreadsheet ID on **Line 15** (optional if bound to the sheet):
   ```javascript
   const SPREADSHEET_ID = '1kUPHPL8hPRKG2d_5D5j6Qb6r72MJxpohh6xL1TEYYWI';
   ```
5. Click **Deploy** → **New deployment**:
   - Type: `Web app`
   - Execute as: `Me`
   - Who has access: `Anyone`
6. Copy the generated **Web app URL** and paste it into [`google_sheets_config.json`](file:///c:/Users/hemac/Desktop/JobCrawler/google_sheets_config.json):
   ```json
   {
     "webhook_url": "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec",
     "enabled": true
   }
   ```

---

### 4. Using the Dashboard (`index.html`)

1. Open [`index.html`](file:///c:/Users/hemac/Desktop/JobCrawler/index.html) in any modern browser or visit your GitHub Pages site.
2. **🔐 Master Password Login**: Enter your Master Password (`JobCompass2026!`) to unlock the board. Unauthenticated visitors cannot view your job queue or metrics.
3. **Search & Filter**: Filter opportunities by Country, Match Score, Workplace, Date Added, or Applied Status.
4. **Mark Applied**: Click **Mark Applied** on any job row. The status updates locally in IndexedDB and syncs instantly to your Google Spreadsheet.
5. **Fetch / Pull**: Click **📥 Fetch from Google Sheet** in the sidebar to load the latest records from Google Sheets on any device.

---

## 📊 Target Countries & Keywords

**Coverage (14 Countries)**:
🇳🇱 Netherlands · 🇮🇪 Ireland · 🇸🇪 Sweden · 🇩🇰 Denmark · 🇫🇮 Finland · 🇫🇷 France · 🇵🇹 Portugal · 🇵🇱 Poland · 🇧🇪 Belgium · 🇦🇹 Austria · 🇦🇺 Australia · 🇸🇬 Singapore · 🇲🇾 Malaysia · 🇳🇿 New Zealand

**Search Keywords**:
`DevOps Engineer`, `Cloud Engineer`, `Site Reliability Engineer`, `Platform Engineer`

---

## 👤 Author
**Hemachandiran Giri** · AWS Certified DevOps Professional
