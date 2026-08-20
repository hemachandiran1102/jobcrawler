/* ══════════════════════════════════════════════════════════════════════
   Job Compass — app.js
   Persistent Excel/CSV import with deduplication via IndexedDB
   ══════════════════════════════════════════════════════════════════════ */

const SOURCE_FILE = 'full_crawl_jobs.csv';
const PAGE_SIZE = 12;
const DB_NAME = 'job-compass-db';
const DB_VERSION = 2;
const STORE_NAME = 'jobs';

const state = { jobs: [], filtered: [], page: 1, sort: { key: 'Match Score', asc: false }, filters: { country: '', source: '', score: '', workplace: '', posted: '', recent: '', status: '' }, query: '', compact: false };
const $ = (id) => document.getElementById(id);
const esc = (value = '') => String(value).replace(/[&<>"']/g, (s) => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#039;' }[s]));
const keyFor = (job) => {
  const norm = normalizeJobUrl(job['Job URL']);
  return norm || `${job.Company}|${job['Job Title']}|${job.Location}`;
};
const shortlist = () => new Set(JSON.parse(localStorage.getItem('job-compass-shortlist') || '[]'));
const saveShortlist = (keys) => localStorage.setItem('job-compass-shortlist', JSON.stringify([...keys]));
const appliedSet = () => new Set(JSON.parse(localStorage.getItem('job-compass-applied') || '[]'));
const saveAppliedSet = (keys) => localStorage.setItem('job-compass-applied', JSON.stringify([...keys]));

function isJobApplied(job) {
  if (!job) return false;
  if (appliedSet().has(keyFor(job))) return true;
  const val = String(job['Applied Status'] || job.AppliedStatus || job.Applied || '').trim().toLowerCase();
  return ['yes', 'applied', 'true', '1', 'done', 'submitted'].includes(val);
}


/* ══════════════════════════════════════════════════════════════════════
   IndexedDB Persistence Layer
   ══════════════════════════════════════════════════════════════════════ */

function openDB() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (db.objectStoreNames.contains(STORE_NAME)) {
        db.deleteObjectStore(STORE_NAME);
      }
      db.createObjectStore(STORE_NAME, { keyPath: '_dedupKey' });
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

/** Read all stored jobs from IndexedDB */
async function loadStoredJobs() {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readonly');
    const store = tx.objectStore(STORE_NAME);
    const request = store.getAll();
    request.onsuccess = () => resolve(request.result || []);
    request.onerror = () => reject(request.error);
  });
}

/** Write an array of jobs into IndexedDB (upsert — won't duplicate by _dedupKey) */
async function storeJobs(jobs) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    for (const job of jobs) {
      store.put(job);
    }
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

/** Clear all jobs from IndexedDB */
async function clearStoredJobs() {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    const request = store.clear();
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}


/* ══════════════════════════════════════════════════════════════════════
   Deduplication Layer
   ══════════════════════════════════════════════════════════════════════ */

/** Normalize LinkedIn, Indeed, Glassdoor URLs to canonical form */
function normalizeJobUrl(url) {
  const u = String(url || '').trim();
  if (!u || ['n/a', 'none', 'nan', '', '-'].includes(u.toLowerCase())) return '';
  const mLi = u.match(/\/jobs\/view\/(?:[^\s/?#]*-)?(\d{6,14})/);
  if (mLi) return `https://www.linkedin.com/jobs/view/${mLi[1]}`;
  const mLiParam = u.match(/[?&]currentJobId=(\d{6,14})/);
  if (mLiParam) return `https://www.linkedin.com/jobs/view/${mLiParam[1]}`;
  const mIndeed = u.match(/[?&]jk=([a-zA-Z0-9]+)/);
  if (mIndeed) return `https://www.indeed.com/viewjob?jk=${mIndeed[1]}`;
  const mGd = u.match(/(?:jl=|jobListingId=|job-listing\/.*?jl=)(\d+)/);
  if (mGd) return `https://www.glassdoor.com/job-listing/?jl=${mGd[1]}`;
  return u.split('#')[0].split('?')[0].replace(/\/+$/, '');
}

/** Alias for backward compatibility */
const normalizeLinkedInUrl = normalizeJobUrl;

/** Generate a unique dedup key for a job */
function dedupKey(job) {
  const normUrl = normalizeJobUrl(job['Job URL']);
  if (normUrl) return `url:${normUrl}`;
  const company = String(job.Company || '').trim().toLowerCase().replace(/&amp;/g, '&');
  const title = String(job['Job Title'] || '').trim().toLowerCase().replace(/&amp;/g, '&');
  const location = String(job.Location || job.Country || '').trim().toLowerCase();
  return `combo:${company}|${title}|${location}`;
}

/** Deduplicate an array of jobs based on normalized LinkedIn URL and merge metadata */
function deduplicateJobsArray(jobsList) {
  if (!Array.isArray(jobsList)) return [];
  const uniqueMap = new Map();

  for (const job of jobsList) {
    if (!job || !job['Job Title'] || job['Job Title'] === 'Job Title') continue;
    const normUrl = normalizeLinkedInUrl(job['Job URL']);
    const key = dedupKey(job);
    const cleanJob = { ...job };
    if (normUrl) cleanJob['Job URL'] = normUrl;
    cleanJob._dedupKey = key;

    if (uniqueMap.has(key)) {
      const existing = uniqueMap.get(key);
      // 1. Preserve Applied Status = Yes
      const isApplied = String(cleanJob['Applied Status'] || cleanJob.Applied || '').trim().toLowerCase() === 'yes';
      const existApplied = String(existing['Applied Status'] || existing.Applied || '').trim().toLowerCase() === 'yes';
      if (isApplied || existApplied) {
        existing['Applied Status'] = 'Yes';
        existing.Applied = 'Yes';
      }
      // 2. Preserve Notes
      if (cleanJob.Notes && !existing.Notes) {
        existing.Notes = cleanJob.Notes;
      }
      // 3. Preserve earliest Crawl Date
      const existDate = String(existing['Crawl Date'] || existing.Date || '');
      const cleanDate = String(cleanJob['Crawl Date'] || cleanJob.Date || '');
      if (cleanDate && (!existDate || cleanDate < existDate)) {
        existing['Crawl Date'] = cleanDate;
        existing.Date = cleanDate;
      }
      // 4. Backfill any missing fields
      for (const [k, v] of Object.entries(cleanJob)) {
        if (v && (!existing[k] || ['N/A', 'Unknown', '-', ''].includes(String(existing[k]).trim()))) {
          existing[k] = v;
        }
      }
    } else {
      uniqueMap.set(key, cleanJob);
    }
  }

  return Array.from(uniqueMap.values());
}

/** Tag each job with a _dedupKey field (used as IndexedDB keyPath) */
function tagJobs(jobs) {
  return deduplicateJobsArray(jobs);
}


/* ══════════════════════════════════════════════════════════════════════
   CSV & Excel Parsing
   ══════════════════════════════════════════════════════════════════════ */

function parseCSV(text) {
  const rows = []; let row = [], value = '', quoted = false;
  for (let i = 0; i < text.length; i++) { const c = text[i], next = text[i + 1];
    if (c === '"' && quoted && next === '"') { value += '"'; i++; }
    else if (c === '"') quoted = !quoted;
    else if (c === ',' && !quoted) { row.push(value); value = ''; }
    else if ((c === '\n' || c === '\r') && !quoted) { if (c === '\r' && next === '\n') i++; row.push(value); if (row.some(Boolean)) rows.push(row); row = []; value = ''; }
    else value += c;
  }
  if (value || row.length) { row.push(value); rows.push(row); }
  const headers = (rows.shift() || []).map((h) => h.replace(/^\uFEFF/, '').trim());
  return rows.map((r) => Object.fromEntries(headers.map((h, i) => [h, (r[i] || '').trim()]))).filter((r) => r['Job Title'] || r.Company);
}

/** Parse an Excel ArrayBuffer using SheetJS into the same job object format */
function parseExcel(arrayBuffer) {
  if (typeof XLSX === 'undefined') {
    showToast('Excel support is loading — please try again in a moment.');
    return [];
  }
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });
  // Normalise all values to trimmed strings
  return rows.map((row) => {
    const job = {};
    for (const [key, val] of Object.entries(row)) {
      job[key.trim()] = String(val ?? '').trim();
    }
    return job;
  }).filter((r) => r['Job Title'] || r.Company);
}


/* ══════════════════════════════════════════════════════════════════════
   Core Application Logic (filters, sorting, rendering)
   ══════════════════════════════════════════════════════════════════════ */

function scoreOf(job) { return Number(String(job['Match Score'] || '').replace(/[^0-9]/g, '')) || 0; }

function parseDate(rawValue) {
  const raw = String(rawValue || '').trim();
  if (!raw || /^(unknown|undefined|none|null|—)$/i.test(raw)) return null;

  // Direct Date parsing (handles ISO 8601 strings like 2026-08-05T18:30:00.000Z)
  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) return parsed;

  // Regex fallback for YYYY-MM-DD or YYYY/MM/DD
  const match = raw.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (match) {
    return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  }

  // Excel Serial Number fallback (e.g., 45500)
  if (/^\d{5}(\.\d+)?$/.test(raw)) {
    const excelEpoch = new Date(1899, 11, 30);
    return new Date(excelEpoch.getTime() + Number(raw) * 86400000);
  }

  return null;
}

function crawlDateOf(job) {
  return parseDate(job['Crawl Date']) || parseDate(job.Date) || parseDate(job.CrawlDate);
}

function postedDateOf(job) {
  return parseDate(job['Posted Date']) || parseDate(job.PostedDate) || crawlDateOf(job);
}

function formatDisplayDate(val) {
  const d = parseDate(val);
  if (!d) {
    const s = String(val || '').trim();
    if (!s || /^(unknown|undefined|none|null)$/i.test(s)) return '—';
    return s.split('T')[0] || s;
  }
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function latestReferenceDate(jobs) {
  let maxDate = null;
  for (const j of jobs) {
    const d = crawlDateOf(j) || postedDateOf(j);
    if (d && (!maxDate || d > maxDate)) maxDate = d;
  }
  return maxDate || new Date();
}
function countBy(list, field) { return [...list.reduce((m, x) => m.set(x[field] || 'Unknown', (m.get(x[field] || 'Unknown') || 0) + 1), new Map())].sort((a,b) => b[1] - a[1]); }
function titleCase(value) { return value || 'Not specified'; }
function roleText(job) { return [job['Job Title'], job.Company, job.Location, job['Required Skills'], job['Search Keyword']].join(' ').toLowerCase(); }

function renderDateTabs() {
  const container = $('date-tabs-bar');
  if (!container) return;

  const dateCounts = new Map();
  let totalCount = 0;

  state.jobs.forEach((job) => {
    totalCount++;
    const dt = formatDisplayDate(crawlDateOf(job) || postedDateOf(job));
    if (dt && dt !== '—') {
      dateCounts.set(dt, (dateCounts.get(dt) || 0) + 1);
    }
  });

  const sortedDates = [...dateCounts.entries()].sort((a, b) => b[0].localeCompare(a[0]));
  
  let html = `<button class="date-tab ${state.selectedDate === 'all' || !state.selectedDate ? 'active' : ''}" data-date="all">All Dates (${totalCount.toLocaleString()})</button>`;
  
  sortedDates.forEach(([dateStr, count]) => {
    const isActive = state.selectedDate === dateStr;
    html += `<button class="date-tab ${isActive ? 'active' : ''}" data-date="${esc(dateStr)}">📅 ${esc(dateStr)} (${count.toLocaleString()})</button>`;
  });

  container.innerHTML = html;

  container.querySelectorAll('[data-date]').forEach((btn) => {
    btn.onclick = () => {
      state.selectedDate = btn.dataset.date;
      state.page = 1;
      renderDateTabs();
      renderAll();
    };
  });
}

function loadJobs(jobs, sourceLabel = 'your export') {
  const cleanJobs = deduplicateJobsArray(jobs);
  state.jobs = cleanJobs; state.page = 1;
  state.filters = { country: '', source: '', score: '', workplace: '', posted: '', recent: '', visa: '', status: '' }; state.query = ''; $('search').value = '';
  populateFilters(); renderDateTabs(); renderAll(); updateStoredCount();
  showToast(`Loaded ${cleanJobs.length.toLocaleString()} opportunities from ${sourceLabel}.`);
}

function populateFilters() {
  const update = (id, values, label) => {
    const select = $(id);
    if (!select) return;
    const current = select.value;
    select.innerHTML = `<option value="">${label}</option>` + values.map((x) => `<option value="${esc(x)}">${esc(x)}</option>`).join('');
    select.value = current;
  };
  update('country-filter', countBy(state.jobs, 'Country').map(([x]) => x), 'All countries');
  update('workplace-filter', countBy(state.jobs, 'Remote / Workplace').map(([x]) => x), 'Any workplace');
  const sources = Array.from(new Set(state.jobs.map(j => (j.Source || 'LinkedIn').trim()).filter(Boolean)));
  if (sources.length > 1) {
    update('source-filter', sources, 'All platforms');
  }
}

function applyFilters() {
  const s = shortlist(), f = state.filters, q = state.query.toLowerCase();
  const referenceDate = latestReferenceDate(state.jobs);

  state.filtered = state.jobs.filter((job) => {
    const applied = (job['Applied Status'] || '').toLowerCase();
    const crawlDt = crawlDateOf(job);
    const postedDt = postedDateOf(job);
    const src = (job.Source || 'LinkedIn').trim();
    
    const crawlAgeInDays = crawlDt ? Math.max(0, (referenceDate - crawlDt) / 86400000) : null;
    const postedAgeInDays = postedDt ? Math.max(0, (referenceDate - postedDt) / 86400000) : null;

    const isCrawlRecent = !f.recent || (crawlAgeInDays !== null && crawlAgeInDays <= Number(f.recent));
    const isPostedRecent = !f.posted || (postedAgeInDays !== null && postedAgeInDays <= Number(f.posted));
    
    const displayCrawlDate = formatDisplayDate(crawlDt || postedDt);
    const matchSelectedDate = !state.selectedDate || state.selectedDate === 'all' || displayCrawlDate === state.selectedDate;

    const isApplied = isJobApplied(job);
    const matchVisa = !f.visa || String(job['Visa Sponsorship Mentioned'] || job['Visa Sponsorship'] || '').toLowerCase().includes(f.visa.toLowerCase());

    return (!q || roleText(job).includes(q)) &&
           (!f.country || job.Country === f.country) &&
           (!f.source || src.toLowerCase().includes(f.source.toLowerCase())) &&
           (!f.workplace || job['Remote / Workplace'] === f.workplace) &&
           (!f.score || scoreOf(job) >= Number(f.score)) &&
           isCrawlRecent && isPostedRecent && matchSelectedDate && matchVisa &&
           (!f.status || (f.status === 'shortlisted' && s.has(keyFor(job))) || (f.status === 'applied' && isApplied) || (f.status === 'not-applied' && !isApplied));
  });

  const { key, asc } = state.sort;
  state.filtered.sort((a,b) => {
    let av = '', bv = '';
    if (key === 'Match Score') {
      av = scoreOf(a); bv = scoreOf(b);
    } else if (key === 'Posted Date') {
      av = postedDateOf(a) ? postedDateOf(a).getTime() : 0;
      bv = postedDateOf(b) ? postedDateOf(b).getTime() : 0;
    } else if (key === 'Crawl Date') {
      av = crawlDateOf(a) ? crawlDateOf(a).getTime() : 0;
      bv = crawlDateOf(b) ? crawlDateOf(b).getTime() : 0;
    } else {
      av = String(a[key] || '').toLowerCase();
      bv = String(b[key] || '').toLowerCase();
    }
    return av < bv ? (asc ? -1 : 1) : av > bv ? (asc ? 1 : -1) : 0;
  });
}

const COUNTRY_FLAGS = {
  'United Kingdom': '🇬🇧',
  'Canada': '🇨🇦',
  'United Arab Emirates': '🇦🇪',
  'Saudi Arabia': '🇸🇦',
  'Qatar': '🇶🇦',
  'Kuwait': '🇰🇼',
  'Bahrain': '🇧🇭',
  'Oman': '🇴🇲',
  'Netherlands': '🇳🇱',
  'Ireland': '🇮🇪',
  'Sweden': '🇸🇪',
  'Denmark': '🇩🇰',
  'Finland': '🇫🇮',
  'France': '🇫🇷',
  'Portugal': '🇵🇹',
  'Poland': '🇵🇱',
  'Belgium': '🇧🇪',
  'Austria': '🇦🇹',
  'Australia': '🇦🇺',
  'Singapore': '🇸🇬',
  'Malaysia': '🇲🇾',
  'New Zealand': '🇳🇿',
  'Germany': '🇩🇪',
  'United States': '🇺🇸',
  'Switzerland': '🇨🇭',
  'Norway': '🇳🇴',
  'Spain': '🇪🇸',
  'Italy': '🇮🇹',
  'India': '🇮🇳',
  'Remote': '🌐'
};

let countryChartSort = 'volume';
let countryChartQuery = '';

function renderMetrics() {
  const all = state.jobs, saved = shortlist(), high = all.filter((j) => scoreOf(j) >= 85), countries = countBy(all, 'Country');
  const appliedJobs = all.filter(isJobApplied);
  const totalRoles = all.length || 1;
  const appliedPct = ((appliedJobs.length / totalRoles) * 100).toFixed(1);

  $('metric-total').textContent = all.length.toLocaleString();
  $('metric-total-note').textContent = `${countBy(all, 'Company').length.toLocaleString()} companies represented`;
  $('metric-high-match').textContent = high.length.toLocaleString();
  
  if ($('metric-applied')) {
    $('metric-applied').textContent = appliedJobs.length.toLocaleString();
  }
  if ($('metric-applied-note')) {
    $('metric-applied-note').textContent = appliedJobs.length ? `${appliedPct}% submitted · Click to filter` : 'Track sent applications';
  }

  $('metric-shortlisted').textContent = saved.size.toLocaleString();
  $('metric-shortlist-note').textContent = saved.size ? 'Your roles to revisit' : 'Start saving standout roles';
  $('metric-countries').textContent = countries.length;
  $('metric-country-note').textContent = countries.length ? `Across ${countries.length} target markets` : 'Primary focus areas';
  
  $('nav-shortlist').textContent = saved.size;
  if ($('nav-applied')) {
    $('nav-applied').textContent = appliedJobs.length.toLocaleString();
  }

  const igCount = all.filter((j) => {
    const s = String(j.Source || '').toLowerCase();
    return s.includes('indeed') || s.includes('glassdoor');
  }).length;
  if ($('nav-ig-total')) {
    $('nav-ig-total').textContent = igCount.toLocaleString();
  }

  // Update Opportunity Queue status tab counts
  if ($('tab-count-all')) $('tab-count-all').textContent = all.length.toLocaleString();
  if ($('tab-count-high')) $('tab-count-high').textContent = high.length.toLocaleString();
  if ($('tab-count-shortlisted')) $('tab-count-shortlisted').textContent = saved.size.toLocaleString();
  if ($('tab-count-applied')) $('tab-count-applied').textContent = appliedJobs.length.toLocaleString();

  updateActiveStatusTab();

  $('dataset-summary').textContent = `${all.length.toLocaleString()} opportunities across ${countries.length} countries — filter the noise, keep what matters, and make every application count.`;

  if ($('country-data-label')) {
    $('country-data-label').textContent = `${countries.length} TARGET MARKETS`;
  }
  if ($('country-chart-subtitle')) {
    $('country-chart-subtitle').textContent = `Showing all ${countries.length} markets · ${all.length.toLocaleString()} total roles`;
  }

  // Showcase all target countries data
  renderCountryBars(countries, totalRoles);

  const best = all.slice().sort((a,b) => scoreOf(b) - scoreOf(a))[0];
  if (best) {
    $('focus-score').innerHTML = `${scoreOf(best)}<small>%</small>`;
    $('focus-title').textContent = best['Job Title'];
    $('focus-company').textContent = `${best.Company || 'Company not stated'} · ${best.Location || best.Country || 'Location not stated'}`;
    $('view-best-fit').onclick = () => focusJob(best);
  }
}

function updateActiveStatusTab() {
  const tabs = document.querySelectorAll('[data-status-tab]');
  tabs.forEach((tab) => {
    const tabType = tab.dataset.statusTab;
    let isActive = false;
    if (tabType === 'all' && !state.filters.status && !state.filters.score) {
      isActive = true;
    } else if (tabType === 'high' && state.filters.score === '85' && !state.filters.status) {
      isActive = true;
    } else if (tabType === 'shortlisted' && state.filters.status === 'shortlisted') {
      isActive = true;
    } else if (tabType === 'applied' && state.filters.status === 'applied') {
      isActive = true;
    }
    tab.classList.toggle('active', isActive);
  });

  const appliedCard = $('metric-applied-card');
  if (appliedCard) {
    appliedCard.classList.toggle('active', state.filters.status === 'applied');
  }
  const shortlistedCard = $('metric-shortlisted-card');
  if (shortlistedCard) {
    shortlistedCard.classList.toggle('active', state.filters.status === 'shortlisted');
  }
}

/* ══════════════════════════════════════════════════════════════════════
   Inbound Pipeline Badge Helper
   ══════════════════════════════════════════════════════════════════════ */

async function loadInboundPipelineBadge() {
  if (location.protocol === 'file:') return;
  try {
    const resp = await fetch('interview_pipeline.json');
    if (resp.ok) {
      const data = await resp.json();
      if (Array.isArray(data)) {
        const nextSteps = data.filter((d) => d['Is Next Step'] === true || d['Is Next Step'] === 'True' || d['Is Next Step'] === 'true' || ['Interview Invitation', 'Technical Assessment', 'Availability / Inquiry'].includes(d.Category));
        if ($('nav-nextsteps-count')) $('nav-nextsteps-count').textContent = nextSteps.length.toString();
      }
    }
  } catch {}
}

function renderCountryBars(countries, totalRoles) {
  const container = $('country-bars');
  if (!container) return;

  if (!countries.length) {
    container.innerHTML = '<p class="empty-hint">No country data loaded yet.</p>';
    return;
  }

  let list = countries.slice();
  if (countryChartQuery) {
    const q = countryChartQuery.toLowerCase();
    list = list.filter(([c]) => c.toLowerCase().includes(q));
  }

  if (countryChartSort === 'alpha') {
    list.sort((a, b) => a[0].localeCompare(b[0]));
  } else {
    list.sort((a, b) => b[1] - a[1]);
  }

  const max = countries[0]?.[1] || 1;
  const activeCountry = (state.filters.country || '').trim().toLowerCase();

  const clearBtn = $('country-filter-clear-btn');
  if (clearBtn) {
    clearBtn.style.display = activeCountry ? 'inline-block' : 'none';
    clearBtn.onclick = (e) => {
      e.stopPropagation();
      state.filters.country = '';
      if ($('country-filter')) $('country-filter').value = '';
      state.page = 1;
      renderAll();
    };
  }

const COUNTRY_CODES = {
  'United Kingdom': 'GB',
  'Canada': 'CA',
  'United Arab Emirates': 'AE',
  'Saudi Arabia': 'SA',
  'Qatar': 'QA',
  'Kuwait': 'KW',
  'Bahrain': 'BH',
  'Oman': 'OM',
  'Netherlands': 'NL',
  'Ireland': 'IE',
  'Sweden': 'SE',
  'Denmark': 'DK',
  'Finland': 'FI',
  'France': 'FR',
  'Portugal': 'PT',
  'Poland': 'PL',
  'Belgium': 'BE',
  'Austria': 'AT',
  'Australia': 'AU',
  'Singapore': 'SG',
  'Malaysia': 'MY',
  'New Zealand': 'NZ',
  'Germany': 'DE',
  'United States': 'US',
  'Switzerland': 'CH',
  'Norway': 'NO',
  'Spain': 'ES',
  'Italy': 'IT',
  'India': 'IN',
  'Remote': 'GL'
};

  container.innerHTML = list.map(([country, n]) => {
    const code = COUNTRY_CODES[country] || (country.length >= 2 ? country.slice(0, 2).toUpperCase() : 'GL');
    const percent = ((n / totalRoles) * 100).toFixed(1);
    const barWidth = Math.max(2, (n / max * 100)).toFixed(1);
    const isActive = activeCountry && (activeCountry === country.toLowerCase());
    return `
      <div class="country-row ${isActive ? 'is-active' : ''}" data-country="${esc(country)}" title="Click to filter by ${esc(country)} (${n.toLocaleString()} roles · ${percent}%)">
        <div class="country-name-cell">
          <span class="country-flag-code">${code}</span>
          <span class="country-name-text">${esc(country)}</span>
        </div>
        <div class="bar">
          <i style="width:${barWidth}%"></i>
        </div>
        <div class="country-count-badge">
          <b>${n.toLocaleString()}</b>
          <span class="country-pct">${percent}%</span>
        </div>
      </div>
    `;
  }).join('');

  container.querySelectorAll('.country-row').forEach((row) => {
    row.onclick = () => {
      const selected = row.dataset.country;
      if (state.filters.country && state.filters.country.toLowerCase() === selected.toLowerCase()) {
        state.filters.country = '';
      } else {
        state.filters.country = selected;
      }
      if ($('country-filter')) $('country-filter').value = state.filters.country;
      state.page = 1;
      renderAll();
      const rolesSection = $('roles');
      if (rolesSection && state.filters.country) {
        rolesSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    };
  });
}

function initCountryChartToolbar() {
  const searchInput = $('country-search-input');
  if (searchInput) {
    searchInput.oninput = (e) => {
      countryChartQuery = e.target.value.trim();
      const countries = countBy(state.jobs, 'Country');
      renderCountryBars(countries, state.jobs.length || 1);
    };
  }

  const btnVolume = $('sort-country-count');
  const btnAlpha = $('sort-country-alpha');
  if (btnVolume && btnAlpha) {
    btnVolume.onclick = () => {
      countryChartSort = 'volume';
      btnVolume.classList.add('active');
      btnAlpha.classList.remove('active');
      const countries = countBy(state.jobs, 'Country');
      renderCountryBars(countries, state.jobs.length || 1);
    };
    btnAlpha.onclick = () => {
      countryChartSort = 'alpha';
      btnAlpha.classList.add('active');
      btnVolume.classList.remove('active');
      const countries = countBy(state.jobs, 'Country');
      renderCountryBars(countries, state.jobs.length || 1);
    };
  }
}

function renderFilters() {
  const labels = { country: 'Country', source: 'Source', score: 'Match ≥', workplace: 'Workplace', posted: 'Job Posted', recent: 'Crawl Date', visa: 'Visa / Reloc', status: 'Status' };
  const active = Object.entries(state.filters).filter(([,v]) => v);
  if ($('filter-count')) $('filter-count').textContent = active.length;
  if ($('filter-button')) $('filter-button').classList.toggle('has-filters', active.length > 0);
  if ($('active-filters')) $('active-filters').innerHTML = active.map(([key,value]) => `<span class="filter-chip">${labels[key] || key}: ${esc(key === 'score' ? value + '%' : (key === 'recent' || key === 'posted') ? `last ${value} days` : value)} <button data-clear="${key}" aria-label="Clear ${labels[key] || key}">×</button></span>`).join('');
  document.querySelectorAll('[data-clear]').forEach((b) => b.onclick = () => { state.filters[b.dataset.clear] = ''; if ($(`${b.dataset.clear}-filter`)) $(`${b.dataset.clear}-filter`).value = ''; state.page = 1; renderAll(); });
}

/* ══════════════════════════════════════════════════════════════════════
   Google Sheets Synchronization
   ══════════════════════════════════════════════════════════════════════ */

const DEFAULT_WEBHOOK_URL = 'https://script.google.com/macros/s/AKfycbxTW2hZQRdGUgkUwH-Ast0K_2080xOSPbKw-gzOjFCmDLIJTiXwAJAwLiCDMoMDxAzL/exec';

function getSheetsWebhookUrl() {
  return localStorage.getItem('job-compass-sheets-webhook') || DEFAULT_WEBHOOK_URL;
}

function saveSheetsWebhookUrl(url) {
  localStorage.setItem('job-compass-sheets-webhook', (url || '').trim());
}

['import-button','hero-import'].forEach((id) => $(id).onclick = () => $('file-input').click());
$('file-input').onchange = (e) => { importFile(e.target.files[0]); e.target.value = ''; };
$('search').oninput = (e) => { state.query = e.target.value; state.page = 1; renderAll(); };
['country','source','score','workplace','posted','recent','visa','status'].forEach((id) => {
  if ($(`${id}-filter`)) {
    $(`${id}-filter`).onchange = (e) => { state.filters[id] = e.target.value; state.page = 1; renderAll(); };
  }
});
$('filter-button').onclick = () => { const panel=$('filters-panel'); panel.hidden=!panel.hidden; }; $('clear-filters').onclick=clearFilters;
document.querySelectorAll('th[data-sort]').forEach((th) => th.onclick = () => { const key=th.dataset.sort; state.sort={key,asc: state.sort.key===key ? !state.sort.asc : key !== 'Match Score'}; renderTable(); });
$('previous-page').onclick=()=>{ if(state.page>1){state.page--;renderTable();} }; $('next-page').onclick=()=>{state.page++;renderTable();};
$('table-view-button').onclick=()=>{state.compact=false;renderAll();}; $('compact-view-button').onclick=()=>{state.compact=true;renderAll();};
$('theme-button').onclick=()=>document.body.classList.toggle('dark'); $('mobile-menu').onclick=()=>document.querySelector('.sidebar').classList.toggle('open');
$('clear-data').onclick = handleClearData;

function fetchJobsViaJSONP(webhookUrl) {
  return new Promise(async (resolve, reject) => {
    // Attempt 1: Standard Fetch (CORS with redirect follow)
    try {
      const sep = webhookUrl.includes('?') ? '&' : '?';
      const fetchResp = await fetch(`${webhookUrl}${sep}action=get_jobs`, { method: 'GET', redirect: 'follow' });
      if (fetchResp.ok) {
        const text = await fetchResp.text();
        try {
          const parsed = JSON.parse(text);
          if (parsed && (parsed.jobs || parsed.success)) {
            return resolve(parsed);
          }
        } catch {
          // If response was wrapped in callback
          const jsonMatch = text.match(/^[a-zA-Z0-9_$]+\(([\s\S]*)\)\s*;?$/);
          if (jsonMatch) {
            return resolve(JSON.parse(jsonMatch[1]));
          }
        }
      }
    } catch {
      // Fall through to JSONP script injection
    }

    // Attempt 2: Script Tag Injection JSONP with 60s timeout
    const callbackName = 'gsCallback_' + Math.random().toString(36).substring(2);
    const script = document.createElement('script');
    const timer = setTimeout(() => {
      cleanup();
      reject(new Error('Google Sheets request timed out after 60 seconds.'));
    }, 60000);

    const cleanup = () => {
      clearTimeout(timer);
      delete window[callbackName];
      if (script.parentNode) script.parentNode.removeChild(script);
    };

    window[callbackName] = (data) => {
      cleanup();
      resolve(data);
    };

    script.onerror = () => {
      cleanup();
      reject(new Error('Network error connecting to Google Apps Script.'));
    };

    const separator = webhookUrl.includes('?') ? '&' : '?';
    script.src = `${webhookUrl}${separator}action=get_jobs&callback=${callbackName}`;
    document.head.appendChild(script);
  });
}

async function syncSingleJobStatusToGoogleSheet(jobUrl, appliedStatus) {
  const webhookUrl = getSheetsWebhookUrl();
  if (!webhookUrl || !jobUrl) return;
  
  try {
    const payload = {
      action: 'update_applied',
      jobUrl: jobUrl,
      appliedStatus: appliedStatus
    };
    await fetch(webhookUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload)
    });
    console.log('[Google Sheets] Synced Applied status:', appliedStatus, 'for', jobUrl);
  } catch (err) {
    console.warn('[Google Sheets] Status sync failed:', err);
  }
}

async function syncAllJobsToGoogleSheet() {
  const webhookUrl = getSheetsWebhookUrl();
  if (!webhookUrl) {
    showToast('Please enter your Google Apps Script Webhook URL first.');
    return;
  }
  if (!state.jobs.length) {
    showToast('No job opportunities loaded to sync.');
    return;
  }
  
  showLoading(`Syncing ${state.jobs.length.toLocaleString()} jobs to Google Sheets…`);
  await yieldUI();
  
  try {
    const payload = {
      action: 'sync_jobs',
      jobs: state.jobs
    };
    const res = await fetch(webhookUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload)
    });
    const data = await res.json();
    console.log('[Google Sheets] Full sync result:', data);
    showToast(`Google Sheets synced! Total rows: ${(data.total || state.jobs.length).toLocaleString()}`);
  } catch (err) {
    console.error('[Google Sheets] Sync failed:', err);
    showToast('Sync failed — check Webhook URL or network connection.');
  } finally {
    hideLoading();
  }
}


function renderTable() {
  applyFilters(); const total = state.filtered.length, pages = Math.max(1, Math.ceil(total/PAGE_SIZE)); if (state.page > pages) state.page = pages;
  const slice = state.filtered.slice((state.page-1)*PAGE_SIZE, state.page*PAGE_SIZE), saved = shortlist();
  $('visible-count').textContent = total.toLocaleString();
  $('jobs-body').innerHTML = slice.length ? slice.map((job) => {
    const score = scoreOf(job), work = job['Remote / Workplace'] || 'Not stated', remote = /remote/i.test(work);
    const postedDate = formatDisplayDate(job['Posted Date'] || job.PostedDate);
    const crawlDate = formatDisplayDate(job['Crawl Date'] || job.Date);
    const isApplied = String(job['Applied Status'] || '').toLowerCase() === 'yes';
    const sourceRaw = String(job.Source || '').trim();
    let sourceBadge = '';
    if (sourceRaw.toLowerCase().includes('glassdoor')) {
      sourceBadge = '<span class="source-tag source-glassdoor">Glassdoor</span>';
    } else if (sourceRaw.toLowerCase().includes('indeed')) {
      sourceBadge = '<span class="source-tag source-indeed">Indeed</span>';
    } else if (sourceRaw.toLowerCase().includes('linkedin') || !sourceRaw) {
      sourceBadge = '<span class="source-tag source-linkedin">LinkedIn</span>';
    } else {
      sourceBadge = `<span class="source-tag">${esc(sourceRaw)}</span>`;
    }

    const visaVal = String(job['Visa Sponsorship Mentioned'] || job['Visa Sponsorship'] || '').trim();
    let visaBadge = '<span class="work-pill" style="color:var(--text-muted); opacity:0.5;">—</span>';
    if (visaVal === 'Visa Sponsored') {
      visaBadge = '<span class="work-pill" style="background:rgba(16,185,129,0.15); color:#10b981; border:1px solid rgba(16,185,129,0.3); font-weight:700;">🛂 Visa Sponsored</span>';
    } else if (visaVal === 'Relocation Provided') {
      visaBadge = '<span class="work-pill" style="background:rgba(59,130,246,0.15); color:#60a5fa; border:1px solid rgba(59,130,246,0.3); font-weight:700;">✈️ Relocation</span>';
    } else if (visaVal === 'Visa Mentioned') {
      visaBadge = '<span class="work-pill" style="background:rgba(245,158,11,0.15); color:#fbbf24; border:1px solid rgba(245,158,11,0.3); font-weight:700;">🌐 Visa Mentioned</span>';
    } else if (visaVal === 'No Sponsorship') {
      visaBadge = '<span class="work-pill" style="background:rgba(239,68,68,0.12); color:#f87171; border:1px solid rgba(239,68,68,0.25);">🚫 No Visa</span>';
    }

    return `<tr>
    <td class="bookmark-cell"><button class="bookmark ${saved.has(keyFor(job)) ? 'active' : ''}" data-save="${esc(keyFor(job))}" title="${saved.has(keyFor(job)) ? 'Remove from shortlist' : 'Add to shortlist'}">${saved.has(keyFor(job)) ? '★' : '☆'}</button></td>
    <td class="col-role"><span class="role-title">${esc(job['Job Title'] || 'Untitled role')}</span><span class="role-sub">${esc(job.Company || 'Not stated')} ${sourceBadge}</span></td>
    <td class="col-country">${esc(job.Country || 'Global')}</td>
    <td class="col-posted"><span class="track">${esc(postedDate)}</span></td>
    <td class="col-crawl"><span class="track">${esc(crawlDate)}</span></td>
    <td class="col-workplace"><span class="work-pill ${remote ? 'remote' : ''}">${esc(work)}</span></td>
    <td class="col-visa">${visaBadge}</td>
    <td class="col-match"><span class="match-pill ${score ? `score-${score}` : 'score-none'}">${score ? score + '%' : '—'}</span></td>
    <td class="col-applied">
      <button class="applied-badge ${isApplied ? 'applied' : ''}" data-applied="${esc(keyFor(job))}">
        ${isApplied ? '✓ Applied' : 'Mark Applied'}
      </button>
    </td>
    <td class="col-link open-cell">${job['Job URL'] ? `<a class="open-link" href="${esc(job['Job URL'])}" target="_blank" rel="noopener" title="Open job posting">↗</a>` : '—'}</td></tr>`;
  }).join('') : `<tr><td colspan="10" class="empty">No opportunities match these filters.<br/><button class="clear-filters" id="empty-clear">Clear filters</button></td></tr>`;

  document.querySelectorAll('[data-save]').forEach((button) => button.onclick = () => { const keys = shortlist(), key = button.dataset.save; const adding = !keys.has(key); adding ? keys.add(key) : keys.delete(key); saveShortlist(keys); renderAll(); showToast(adding ? 'Added to your shortlist.' : 'Removed from your shortlist.'); });

  document.querySelectorAll('[data-applied]').forEach((button) => button.onclick = async () => {
    const key = button.dataset.applied;
    const targetJob = state.jobs.find((j) => keyFor(j) === key);
    if (!targetJob) return;
    
    const currentlyApplied = isJobApplied(targetJob);
    const newStatus = currentlyApplied ? 'No' : 'Yes';
    targetJob['Applied Status'] = newStatus;

    const appKeys = appliedSet();
    if (newStatus === 'Yes') {
      appKeys.add(key);
    } else {
      appKeys.delete(key);
    }
    saveAppliedSet(appKeys);
    
    try {
      await storeJobs([targetJob]);
    } catch { /* IndexedDB error fallback */ }
    
    renderAll();
    showToast(newStatus === 'Yes' ? `Marked "${targetJob['Job Title']}" as Applied.` : `Marked "${targetJob['Job Title']}" as Not Applied.`);
    
    // Sync status change to Google Sheets Webhook
    if (targetJob['Job URL']) {
      syncSingleJobStatusToGoogleSheet(targetJob['Job URL'], newStatus);
    }
  });

  $('empty-clear')?.addEventListener('click', clearFilters); $('page-info').textContent = total ? `Showing ${(state.page-1)*PAGE_SIZE+1}–${Math.min(state.page*PAGE_SIZE,total)} of ${total.toLocaleString()} roles` : 'No roles to show';
  $('previous-page').disabled = state.page === 1; $('next-page').disabled = state.page === pages; $('page-buttons').innerHTML = pagination(pages); document.querySelectorAll('[data-page]').forEach((b) => b.onclick = () => { state.page = Number(b.dataset.page); renderTable(); });
}

function pagination(pages) { const set = new Set([1,pages,state.page-1,state.page,state.page+1].filter((n) => n >= 1 && n <= pages)); let previous = 0; return [...set].sort((a,b)=>a-b).map((n) => `${n-previous > 1 ? '<span>…</span>' : ''}<button class="${n===state.page?'active':''}" data-page="${n}">${n}</button>`).join(''); }
function renderAll() { renderMetrics(); renderFilters(); renderTable(); $('table-view-button').classList.toggle('active', !state.compact); $('compact-view-button').classList.toggle('active', state.compact); document.querySelector('.table-wrap').classList.toggle('compact', state.compact); }
function clearFilters() { state.filters = { country:'', source:'', score:'', workplace:'', posted:'', recent:'', visa:'', status:'' }; state.query = ''; state.page = 1; $('search').value=''; ['country','source','score','workplace','posted','recent','visa','status'].forEach((id)=>{ if($(`${id}-filter`)) $(`${id}-filter`).value=''; }); renderAll(); }
function focusJob(job) { clearFilters(); state.query = job['Job Title'] || ''; $('search').value = state.query; applyFilters(); const exact = state.filtered.findIndex((x)=>keyFor(x)===keyFor(job)); state.page = Math.floor(Math.max(exact,0)/PAGE_SIZE)+1; renderTable(); $('roles').scrollIntoView({ behavior:'smooth', block:'start' }); }
function showToast(message) { const toast = $('toast'); toast.textContent = message; toast.classList.add('show'); clearTimeout(showToast.timer); showToast.timer = setTimeout(()=>toast.classList.remove('show'), 3500); }


/** Show/hide the loading overlay with a status message */
function showLoading(text) {
  const overlay = $('loading-overlay');
  $('loading-text').textContent = text || 'Importing…';
  overlay.hidden = false;
}
function hideLoading() { $('loading-overlay').hidden = true; }

/** Yield to the browser so it can paint the loading UI */
function yieldUI() { return new Promise((r) => setTimeout(r, 0)); }

async function importFile(file) {
  if (!file) return;
  const isExcel = /\.xlsx$/i.test(file.name);
  const isCSV = /\.csv$/i.test(file.name);
  if (!isExcel && !isCSV) return showToast('Please choose a CSV or Excel (.xlsx) file.');

  showLoading(`Reading ${file.name}…`);
  await yieldUI();

  try {
    console.log('[Import] Starting import of', file.name, '(' + (file.size / 1024).toFixed(0) + ' KB)');

    let incoming = [];
    if (isExcel) {
      console.log('[Import] Reading file as ArrayBuffer…');
      const buffer = await file.arrayBuffer();
      console.log('[Import] Buffer ready, parsing Excel…');
      showLoading('Parsing Excel data…');
      await yieldUI();
      incoming = parseExcel(buffer);
      console.log('[Import] Excel parsed:', incoming.length, 'rows');
    } else {
      console.log('[Import] Reading CSV…');
      const text = await file.text();
      incoming = parseCSV(text);
      console.log('[Import] CSV parsed:', incoming.length, 'rows');
    }

    if (!incoming.length) {
      showToast('That file did not contain readable job rows.');
      return;
    }

    showLoading(`Processing ${incoming.length.toLocaleString()} jobs…`);
    await yieldUI();

    const tagged = tagJobs(incoming);
    console.log('[Import] Tagged', tagged.length, 'jobs with dedup keys');

    showLoading('Checking for duplicates…');
    await yieldUI();

    let existingJobs = [];
    try {
      existingJobs = await loadStoredJobs();
    } catch (dbErr) {
      console.warn('[Import] Could not load stored jobs:', dbErr);
    }
    const existingKeys = new Set(existingJobs.map((j) => j._dedupKey));
    console.log('[Import] Existing stored jobs:', existingJobs.length);

    const newJobs = [];
    let dupeCount = 0;
    for (const job of tagged) {
      if (existingKeys.has(job._dedupKey)) {
        dupeCount++;
      } else {
        newJobs.push(job);
        existingKeys.add(job._dedupKey);
      }
    }
    console.log('[Import] New:', newJobs.length, '| Duplicates:', dupeCount);

    if (newJobs.length > 0) {
      showLoading(`Saving ${newJobs.length.toLocaleString()} new jobs…`);
      await yieldUI();
      try {
        await storeJobs(newJobs);
        console.log('[Import] Saved to IndexedDB');
      } catch (dbErr) {
        console.warn('[Import] IndexedDB write failed:', dbErr);
      }
    }

    const allJobs = [...existingJobs, ...newJobs];
    showLoading('Loading dashboard…');
    await yieldUI();
    loadJobs(allJobs, file.name);
    console.log('[Import] Done. Total:', allJobs.length);

    showToast(`Imported ${newJobs.length.toLocaleString()} new job${newJobs.length !== 1 ? 's' : ''} (${dupeCount.toLocaleString()} duplicate${dupeCount !== 1 ? 's' : ''} skipped). Total: ${allJobs.length.toLocaleString()}`);
  } catch (err) {
    console.error('[Import] Fatal error:', err);
    showToast('Import failed: ' + (err.message || 'unknown error'));
  } finally {
    hideLoading();
  }
}


/* ══════════════════════════════════════════════════════════════════════
   Sidebar Stored Count & Clear Data
   ══════════════════════════════════════════════════════════════════════ */

async function updateStoredCount() {
  try {
    const jobs = await loadStoredJobs();
    const count = jobs.length;
    const el = $('stored-count');
    const clearBtn = $('clear-data');
    if (count > 0) {
      el.textContent = `${count.toLocaleString()} jobs stored`;
      clearBtn.style.display = 'block';
    } else {
      el.textContent = '';
      clearBtn.style.display = 'none';
    }
  } catch { /* IndexedDB may not be available */ }
}

async function handleClearData() {
  if (!confirm('This will permanently remove all stored job data. Continue?')) return;
  await clearStoredJobs();
  state.jobs = []; state.filtered = []; state.page = 1;
  renderAll(); updateStoredCount();
  $('dataset-summary').textContent = 'All stored data cleared. Import a CSV or Excel file to start fresh.';
  $('jobs-body').innerHTML = '<tr><td colspan="8" class="empty">Your dashboard is ready. Import a CSV or Excel file to begin.</td></tr>';
  showToast('All stored job data has been cleared.');
}


function loadInboundPipelineBadge() {
  fetch('interview_pipeline.json')
    .then((r) => r.json())
    .then((data) => {
      const nextSteps = (data || []).filter((d) => d['Is Next Step'] === true || d['Is Next Step'] === 'True' || d['Is Next Step'] === 'true');
      if ($('nav-nextsteps-count')) $('nav-nextsteps-count').textContent = nextSteps.length.toString();
    })
    .catch(() => {});
}

/* ══════════════════════════════════════════════════════════════════════
   AUTHENTICATION & SECURITY LAYER (Web Crypto API SHA-256)
   ══════════════════════════════════════════════════════════════════════ */

const MASTER_PASSWORD_HASH = '5a8dc1ec9f6708f0e7071d8fbf7bb455c0edd294046c5d0a7d9dbf72f2a16f4b';
const ALT_PASSWORD_HASH = 'e2a23afe0cdeccbaf4fab5e2387a134d32ba1064d96d11454af926917e2a5383';
const MASTER_PASSWORD_SALT = 'jobcompass_salt_2026';
const AUTH_SESSION_KEY = 'job-compass-auth-active';

/** Hash password with salt using Web Crypto API SHA-256 */
async function hashPassword(password, salt) {
  const encoder = new TextEncoder();
  const data = encoder.encode(password + '::' + salt);
  const hashBuffer = await crypto.subtle.digest('SHA-256', data);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map((b) => b.toString(16).padStart(2, '0')).join('');
}

/** Check if user is currently authenticated */
function isAuthActive() {
  return sessionStorage.getItem(AUTH_SESSION_KEY) === 'true';
}

function showAuthModal() {
  if (!$('auth-overlay')) {
    unlockDashboard();
    return;
  }
  if ($('app-shell')) {
    $('app-shell').hidden = false;
    $('app-shell').style.filter = 'blur(4px)';
    $('app-shell').style.pointerEvents = 'none';
    $('app-shell').style.userSelect = 'none';
  }
  if ($('auth-overlay')) $('auth-overlay').hidden = false;
  if ($('auth-error')) $('auth-error').hidden = true;
  if ($('auth-password')) $('auth-password').value = '';
  if ($('auth-confirm-wrapper')) $('auth-confirm-wrapper').hidden = true;
  if ($('password-strength-bar')) $('password-strength-bar').hidden = true;
  if ($('auth-title')) $('auth-title').textContent = '🔐 Protected Access';
  if ($('auth-subtitle')) $('auth-subtitle').textContent = 'Enter your master password to unlock your job board.';
  if ($('auth-submit-btn')) $('auth-submit-btn').textContent = 'Unlock Dashboard →';
}

/** Unlock dashboard */
function unlockDashboard() {
  sessionStorage.setItem(AUTH_SESSION_KEY, 'true');
  if ($('auth-overlay')) $('auth-overlay').hidden = true;
  if ($('app-shell')) {
    $('app-shell').hidden = false;
    $('app-shell').style.filter = 'none';
    $('app-shell').style.pointerEvents = 'auto';
    $('app-shell').style.userSelect = 'auto';
  }
  initDashboardData();
}

/** Lock dashboard session */
function lockDashboard() {
  sessionStorage.removeItem(AUTH_SESSION_KEY);
  showAuthModal();
  showToast('Dashboard session locked.');
}

/** Initialize Auth Listeners */
function initAuth() {
  const toggleBtn = $('toggle-auth-password');
  if (toggleBtn) {
    toggleBtn.onclick = () => {
      const input = $('auth-password');
      if (input.type === 'password') {
        input.type = 'text';
        toggleBtn.textContent = '🙈';
      } else {
        input.type = 'password';
        toggleBtn.textContent = '👁️';
      }
    };
  }

  const authForm = $('auth-form');
  if (authForm) {
    authForm.onsubmit = async (e) => {
      e.preventDefault();
      const password = ($('auth-password').value || '').trim();
      if (!password) return;

      const computedHash = await hashPassword(password, MASTER_PASSWORD_SALT);
      const customHash = localStorage.getItem('job-compass-custom-hash');

      if (computedHash === MASTER_PASSWORD_HASH || computedHash === ALT_PASSWORD_HASH || (customHash && computedHash === customHash)) {
        unlockDashboard();
        showToast('Dashboard unlocked!');
      } else {
        $('auth-error').textContent = 'Incorrect master password. Access denied.';
        $('auth-error').hidden = false;
      }
    };
  }

  const lockBtn = $('lock-dashboard-btn');
  if (lockBtn) lockBtn.onclick = lockDashboard;
  const lockHeaderBtn = $('lock-header-btn');
  if (lockHeaderBtn) lockHeaderBtn.onclick = lockDashboard;
}

async function fetchJobsFromGoogleSheet(silent = false) {
  const webhookUrl = getSheetsWebhookUrl();
  if (!webhookUrl) return false;
  if (!silent) showLoading('Fetching opportunities from Google Sheets…');
  try {
    const data = await fetchJobsViaJSONP(webhookUrl);
    if (data && data.jobs && Array.isArray(data.jobs) && data.jobs.length > 0) {
      const deduped = deduplicateJobsArray(data.jobs);
      await storeJobs(deduped);
      loadJobs(deduped, 'Google Sheets');
      if (!silent) showToast(`Fetched ${deduped.length.toLocaleString()} jobs from Google Sheets.`);
      return true;
    }
  } catch (err) {
    if (!silent) showToast('Could not fetch from Google Sheets: ' + (err.message || 'network error'));
  } finally {
    if (!silent) hideLoading();
  }
  return false;
}

// Google Sheets Modal Event Listeners
$('sheets-config-button').onclick = () => {
  $('sheets-url-input').value = getSheetsWebhookUrl();
  $('sheets-modal').hidden = false;
};
$('sheets-fetch-button').onclick = () => fetchJobsFromGoogleSheet(false);
$('close-sheets-modal').onclick = () => { $('sheets-modal').hidden = true; };
$('save-sheets-config').onclick = () => {
  const url = $('sheets-url-input').value.trim();
  saveSheetsWebhookUrl(url);
  $('sheets-modal').hidden = true;
  showToast(url ? 'Google Sheets Webhook URL saved!' : 'Google Sheets Webhook URL cleared.');
};
$('sync-sheets-now').onclick = () => {
  const url = $('sheets-url-input').value.trim();
  if (url) saveSheetsWebhookUrl(url);
  syncAllJobsToGoogleSheet();
};

// ── Startup: Authentication Gate check & Data Loading ──
async function initDashboardData() {
  loadInboundPipelineBadge();

  // Load default Webhook URL from config file if not set in localStorage
  try {
    const cfgResp = await fetch('google_sheets_config.json');
    if (cfgResp.ok) {
      const cfg = await cfgResp.json();
      const defaultWebhook = (cfg.webhook_url || '').trim();
      if (defaultWebhook && !getSheetsWebhookUrl()) {
        saveSheetsWebhookUrl(defaultWebhook);
      }
    }
  } catch {}

  // 1. Try fetching freshest local master CSV files (LinkedIn + Indeed + Glassdoor)
  try {
      const combinedJobs = [];
      const cacheBust = `?_t=${Date.now()}`;

      // 1A. Load LinkedIn jobs (full_crawl_jobs.csv)
      try {
        const resp1 = await fetch(SOURCE_FILE + cacheBust);
        if (resp1.ok) {
          const text1 = await resp1.text();
          const j1 = parseCSV(text1).map((j) => ({ ...j, Source: j.Source || 'LinkedIn' }));
          combinedJobs.push(...j1);
        }
      } catch {}

      // 1B. Load Indeed & Glassdoor jobs (indeed_glassdoor_jobs.csv)
      try {
        const resp2 = await fetch('indeed_glassdoor_jobs.csv' + cacheBust);
        if (resp2.ok) {
          const text2 = await resp2.text();
          const j2 = parseCSV(text2);
          combinedJobs.push(...j2);
        }
      } catch {}

      if (combinedJobs.length > 0) {
        const deduped = deduplicateJobsArray(combinedJobs);
        await storeJobs(deduped);
        loadJobs(deduped, 'All Platforms (LinkedIn, Indeed, Glassdoor)');
        return;
      }
    } catch (err) {
      console.warn('[Data Init] Local CSV fetch error:', err);
    }

  // 2. Try stored IndexedDB
  try {
    const storedJobs = await loadStoredJobs();
    if (storedJobs.length > 0) {
      const deduped = deduplicateJobsArray(storedJobs);
      if (deduped.length !== storedJobs.length) {
        await clearStoredJobs();
        await storeJobs(deduped);
      }
      loadJobs(deduped, 'stored data');
      return;
    }
  } catch { /* IndexedDB unavailable */ }

  // 3. Try pulling live data from Google Sheets if available
  try {
    const pulledFromSheets = await fetchJobsFromGoogleSheet(true);
    if (pulledFromSheets) return;
  } catch {}

  // Nothing to load
  $('dataset-summary').textContent = 'Import your CSV or Excel job export to start ranking and shortlisting roles.';
  $('jobs-body').innerHTML = '<tr><td colspan="9" class="empty">Your dashboard is ready. Import a CSV or Excel file to begin.</td></tr>';
  updateStoredCount();
}

function initStatusTabs() {
  document.querySelectorAll('[data-status-tab]').forEach((tab) => {
    tab.onclick = () => {
      const type = tab.dataset.statusTab;
      if (type === 'all') {
        state.filters.status = '';
        state.filters.score = '';
        if ($('score-filter')) $('score-filter').value = '';
        if ($('status-filter')) $('status-filter').value = '';
      } else if (type === 'high') {
        state.filters.status = '';
        state.filters.score = '85';
        if ($('score-filter')) $('score-filter').value = '85';
        if ($('status-filter')) $('status-filter').value = '';
      } else if (type === 'shortlisted') {
        state.filters.status = 'shortlisted';
        if ($('status-filter')) $('status-filter').value = 'shortlisted';
      } else if (type === 'applied') {
        state.filters.status = 'applied';
        if ($('status-filter')) $('status-filter').value = 'applied';
      }
      state.page = 1;
      renderAll();
    };
  });

  const appliedCard = $('metric-applied-card');
  if (appliedCard) {
    appliedCard.onclick = () => {
      if (state.filters.status === 'applied') {
        state.filters.status = '';
      } else {
        state.filters.status = 'applied';
      }
      if ($('status-filter')) $('status-filter').value = state.filters.status;
      state.page = 1;
      renderAll();
      const rolesSec = $('roles');
      if (rolesSec && state.filters.status) rolesSec.scrollIntoView({ behavior: 'smooth', block: 'start' });
    };
  }

  const shortlistedCard = $('metric-shortlisted-card');
  if (shortlistedCard) {
    shortlistedCard.onclick = () => {
      if (state.filters.status === 'shortlisted') {
        state.filters.status = '';
      } else {
        state.filters.status = 'shortlisted';
      }
      if ($('status-filter')) $('status-filter').value = state.filters.status;
      state.page = 1;
      renderAll();
      const rolesSec = $('roles');
      if (rolesSec && state.filters.status) rolesSec.scrollIntoView({ behavior: 'smooth', block: 'start' });
    };
  }

  const navAppliedLink = $('nav-applied-link');
  if (navAppliedLink) {
    navAppliedLink.onclick = (e) => {
      e.preventDefault();
      state.filters.status = 'applied';
      if ($('status-filter')) $('status-filter').value = 'applied';
      state.page = 1;
      renderAll();
      const rolesSec = $('roles');
      if (rolesSec) rolesSec.scrollIntoView({ behavior: 'smooth', block: 'start' });
    };
  }

  const navShortlistLink = $('nav-shortlist-link');
  if (navShortlistLink) {
    navShortlistLink.onclick = (e) => {
      e.preventDefault();
      state.filters.status = 'shortlisted';
      if ($('status-filter')) $('status-filter').value = 'shortlisted';
      state.page = 1;
      renderAll();
      const rolesSec = $('roles');
      if (rolesSec) rolesSec.scrollIntoView({ behavior: 'smooth', block: 'start' });
    };
  }

  const navRolesLink = $('nav-roles-link');
  if (navRolesLink) {
    navRolesLink.onclick = (e) => {
      e.preventDefault();
      state.filters.status = '';
      state.filters.score = '';
      if ($('status-filter')) $('status-filter').value = '';
      if ($('score-filter')) $('score-filter').value = '';
      state.page = 1;
      renderAll();
      const rolesSec = $('roles');
      if (rolesSec) rolesSec.scrollIntoView({ behavior: 'smooth', block: 'start' });
    };
  }
}

(function init() {
  initAuth();
  initCountryChartToolbar();
  initStatusTabs();
  loadInboundPipelineBadge();
  if ($('auth-overlay') && !isAuthActive()) {
    showAuthModal();
  } else {
    unlockDashboard();
  }
})();
