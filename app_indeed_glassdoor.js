/* ══════════════════════════════════════════════════════════════════════
   Job Compass — Indeed & Glassdoor Intelligence (app_indeed_glassdoor.js)
   ══════════════════════════════════════════════════════════════════════ */

const SOURCE_FILE = 'indeed_glassdoor_jobs.csv';
const PAGE_SIZE = 12;
const DB_NAME = 'job-compass-ig-db';
const DB_VERSION = 1;
const STORE_NAME = 'jobs';

const state = {
  jobs: [],
  filtered: [],
  page: 1,
  sort: { key: 'Match Score', asc: false },
  filters: { country: '', source: '', score: '', workplace: '', posted: '', recent: '', status: '' },
  query: '',
  selectedDate: 'all',
  compact: false
};

const $ = (id) => document.getElementById(id);
const esc = (value = '') => String(value).replace(/[&<>"']/g, (s) => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#039;' }[s]));

function normalizeJobUrl(url) {
  const u = String(url || '').trim();
  if (!u || ['n/a', 'none', 'nan', '', '-'].includes(u.toLowerCase())) return '';
  const mIndeed = u.match(/[?&]jk=([a-zA-Z0-9]+)/);
  if (mIndeed) return `https://www.indeed.com/viewjob?jk=${mIndeed[1]}`;
  const mGd = u.match(/(?:jl=|jobListingId=|job-listing\/.*?jl=)(\d+)/);
  if (mGd) return `https://www.glassdoor.com/job-listing/?jl=${mGd[1]}`;
  return u.split('#')[0].split('?')[0].replace(/\/+$/, '');
}

const keyFor = (job) => {
  const norm = normalizeJobUrl(job['Job URL']);
  return norm || `${job.Company}|${job['Job Title']}|${job.Location}`;
};

const shortlist = () => new Set(JSON.parse(localStorage.getItem('job-compass-ig-shortlist') || '[]'));
const saveShortlist = (keys) => localStorage.setItem('job-compass-ig-shortlist', JSON.stringify([...keys]));
const appliedSet = () => new Set(JSON.parse(localStorage.getItem('job-compass-applied') || '[]'));
const saveAppliedSet = (keys) => localStorage.setItem('job-compass-applied', JSON.stringify([...keys]));

function isJobApplied(job) {
  if (!job) return false;
  if (appliedSet().has(keyFor(job))) return true;
  const val = String(job['Applied Status'] || job.AppliedStatus || job.Applied || '').trim().toLowerCase();
  return ['yes', 'applied', 'true', '1', 'done', 'submitted'].includes(val);
}


/* ══════════════════════════════════════════════════════════════════════
   IndexedDB Layer
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

function dedupKey(job) {
  const normUrl = normalizeJobUrl(job['Job URL']);
  if (normUrl) return `url:${normUrl}`;
  const company = String(job.Company || '').trim().toLowerCase().replace(/&amp;/g, '&');
  const title = String(job['Job Title'] || '').trim().toLowerCase().replace(/&amp;/g, '&');
  const location = String(job.Location || job.Country || '').trim().toLowerCase();
  return `combo:${company}|${title}|${location}`;
}

function deduplicateJobsArray(jobsList) {
  if (!Array.isArray(jobsList)) return [];
  const uniqueMap = new Map();

  for (const job of jobsList) {
    if (!job || !job['Job Title'] || job['Job Title'] === 'Job Title') continue;
    const normUrl = normalizeJobUrl(job['Job URL']);
    const key = dedupKey(job);
    const cleanJob = { ...job };
    if (normUrl) cleanJob['Job URL'] = normUrl;
    cleanJob._dedupKey = key;

    if (uniqueMap.has(key)) {
      const existing = uniqueMap.get(key);
      const isApplied = String(cleanJob['Applied Status'] || cleanJob.Applied || '').trim().toLowerCase() === 'yes';
      const existApplied = String(existing['Applied Status'] || existing.Applied || '').trim().toLowerCase() === 'yes';
      if (isApplied || existApplied) {
        existing['Applied Status'] = 'Yes';
        existing.Applied = 'Yes';
      }
      if (cleanJob.Notes && !existing.Notes) {
        existing.Notes = cleanJob.Notes;
      }
      const existDate = String(existing['Crawl Date'] || existing.Date || '');
      const cleanDate = String(cleanJob['Crawl Date'] || cleanJob.Date || '');
      if (cleanDate && (!existDate || cleanDate < existDate)) {
        existing['Crawl Date'] = cleanDate;
        existing.Date = cleanDate;
      }
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


/* ══════════════════════════════════════════════════════════════════════
   CSV Parsing
   ══════════════════════════════════════════════════════════════════════ */

function parseCSV(text) {
  const rows = []; let row = [], value = '', quoted = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i], next = text[i + 1];
    if (c === '"' && quoted && next === '"') { value += '"'; i++; }
    else if (c === '"') quoted = !quoted;
    else if (c === ',' && !quoted) { row.push(value); value = ''; }
    else if ((c === '\n' || c === '\r') && !quoted) {
      if (c === '\r' && next === '\n') i++;
      row.push(value);
      if (row.some(Boolean)) rows.push(row);
      row = []; value = '';
    } else value += c;
  }
  if (value || row.length) { row.push(value); rows.push(row); }
  const headers = (rows.shift() || []).map((h) => h.replace(/^\uFEFF/, '').trim());
  return rows.map((r) => Object.fromEntries(headers.map((h, i) => [h, (r[i] || '').trim()]))).filter((r) => r['Job Title'] || r.Company);
}


/* ══════════════════════════════════════════════════════════════════════
   Helper & Date Functions
   ══════════════════════════════════════════════════════════════════════ */

function scoreOf(job) { return Number(String(job['Match Score'] || '').replace(/[^0-9]/g, '')) || 0; }

function parseDate(rawValue) {
  const raw = String(rawValue || '').trim();
  if (!raw || /^(unknown|undefined|none|null|—)$/i.test(raw)) return null;
  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) return parsed;
  const match = raw.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (match) return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  return null;
}

function crawlDateOf(job) { return parseDate(job['Crawl Date']) || parseDate(job.Date); }
function postedDateOf(job) { return parseDate(job['Posted Date']) || parseDate(job.PostedDate) || crawlDateOf(job); }

function formatDisplayDate(val) {
  const d = parseDate(val);
  if (!d) {
    const s = String(val || '').trim();
    return (!s || /^(unknown|undefined|none|null)$/i.test(s)) ? '—' : (s.split('T')[0] || s);
  }
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function countBy(list, field) {
  return [...list.reduce((m, x) => m.set(x[field] || 'Unknown', (m.get(x[field] || 'Unknown') || 0) + 1), new Map())].sort((a,b) => b[1] - a[1]);
}

function roleText(job) {
  return [job['Job Title'], job.Company, job.Location, job['Required Skills'], job['Search Keyword'], job.Source].join(' ').toLowerCase();
}


/* ══════════════════════════════════════════════════════════════════════
   Rendering Logic
   ══════════════════════════════════════════════════════════════════════ */

function loadJobs(jobs, sourceLabel = 'Indeed & Glassdoor Export') {
  const cleanJobs = deduplicateJobsArray(jobs);
  state.jobs = cleanJobs; state.page = 1;
  state.filters = { country: '', source: '', score: '', workplace: '', posted: '', recent: '', status: '' };
  state.query = '';
  if ($('search')) $('search').value = '';
  populateFilters();
  renderDateTabs();
  renderAll();
  updateStoredCount();
  showToast(`Loaded ${cleanJobs.length.toLocaleString()} opportunities from ${sourceLabel}.`);
}

function updateStoredCount() {
  const c = state.jobs.length;
  if ($('stored-count')) $('stored-count').textContent = `${c.toLocaleString()} opportunities stored`;
  if ($('nav-all')) $('nav-all').textContent = c.toLocaleString();
  if ($('nav-ig-total')) $('nav-ig-total').textContent = c.toLocaleString();
  const appliedCount = state.jobs.filter(isJobApplied).length;
  if ($('nav-applied')) $('nav-applied').textContent = appliedCount.toLocaleString();
  const savedCount = shortlist().size;
  if ($('nav-shortlist')) $('nav-shortlist').textContent = savedCount.toLocaleString();

  // Load Inbound count badge if available
  fetch('interview_pipeline.json').then((r) => r.json()).then((data) => {
    const nextSteps = data.filter((d) => d['Is Next Step'] === true || d['Is Next Step'] === 'True' || d['Is Next Step'] === 'true');
    if ($('nav-inbound-total')) $('nav-inbound-total').textContent = nextSteps.length;
  }).catch(() => {});
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
}

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
  container.querySelectorAll('.date-tab').forEach((btn) => {
    btn.onclick = () => {
      state.selectedDate = btn.dataset.date;
      state.page = 1;
      container.querySelectorAll('.date-tab').forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      renderAll();
    };
  });
}

function applyFilters() {
  const s = shortlist(), f = state.filters, q = state.query.toLowerCase();
  const refDate = new Date();

  state.filtered = state.jobs.filter((job) => {
    const applied = (job['Applied Status'] || '').toLowerCase();
    const source = (job.Source || '').toLowerCase();
    const crawlDt = crawlDateOf(job);
    const postedDt = postedDateOf(job);

    const crawlAgeInDays = crawlDt ? Math.max(0, (refDate - crawlDt) / 86400000) : null;
    const postedAgeInDays = postedDt ? Math.max(0, (refDate - postedDt) / 86400000) : null;

    const isCrawlRecent = !f.recent || (crawlAgeInDays !== null && crawlAgeInDays <= Number(f.recent));
    const isPostedRecent = !f.posted || (postedAgeInDays !== null && postedAgeInDays <= Number(f.posted));
    const displayCrawlDate = formatDisplayDate(crawlDt || postedDt);
    const matchSelectedDate = !state.selectedDate || state.selectedDate === 'all' || displayCrawlDate === state.selectedDate;

    const isApplied = isJobApplied(job);

    return (!q || roleText(job).includes(q)) &&
           (!f.country || job.Country === f.country) &&
           (!f.source || source.includes(f.source.toLowerCase())) &&
           (!f.workplace || job['Remote / Workplace'] === f.workplace) &&
           (!f.score || scoreOf(job) >= Number(f.score)) &&
           isCrawlRecent && isPostedRecent && matchSelectedDate &&
           (!f.status || (f.status === 'shortlisted' && s.has(keyFor(job))) || (f.status === 'applied' && isApplied) || (f.status === 'not-applied' && !isApplied) || (f.status === 'indeed' && source.includes('indeed')) || (f.status === 'glassdoor' && source.includes('glassdoor')));
  });

  const { key, asc } = state.sort;
  state.filtered.sort((a, b) => {
    let av = '', bv = '';
    if (key === 'Match Score') {
      av = scoreOf(a); bv = scoreOf(b);
    } else {
      av = String(a[key] || '').toLowerCase();
      bv = String(b[key] || '').toLowerCase();
    }
    return av < bv ? (asc ? -1 : 1) : av > bv ? (asc ? 1 : -1) : 0;
  });
}

function renderMetrics() {
  const total = state.jobs.length;
  const companies = new Set(state.jobs.map((j) => j.Company)).size;
  const highMatch = state.jobs.filter((j) => scoreOf(j) >= 85).length;
  const appliedCount = state.jobs.filter((j) => String(j['Applied Status'] || '').toLowerCase() === 'yes').length;
  const savedCount = shortlist().size;
  const countries = countBy(state.jobs, 'Country');
  const indeedCount = state.jobs.filter((j) => (j.Source || '').toLowerCase().includes('indeed')).length;
  const gdCount = state.jobs.filter((j) => (j.Source || '').toLowerCase().includes('glassdoor')).length;

  $('metric-total').textContent = total.toLocaleString();
  $('metric-total-note').textContent = `${companies.toLocaleString()} companies · ${indeedCount} Indeed / ${gdCount} Glassdoor`;
  $('metric-high-match').textContent = highMatch.toLocaleString();
  $('metric-applied').textContent = appliedCount.toLocaleString();
  $('metric-applied-note').textContent = `${total ? ((appliedCount / total) * 100).toFixed(1) : 0}% submitted · Click to filter`;
  $('metric-shortlisted').textContent = savedCount.toLocaleString();
  $('metric-countries').textContent = countries.length.toString();

  if ($('tab-count-indeed')) $('tab-count-indeed').textContent = indeedCount.toLocaleString();
  if ($('tab-count-glassdoor')) $('tab-count-glassdoor').textContent = gdCount.toLocaleString();
  if ($('tab-count-applied')) $('tab-count-applied').textContent = appliedCount.toLocaleString();
  if ($('tab-count-shortlisted')) $('tab-count-shortlisted').textContent = savedCount.toLocaleString();
  if ($('nav-ig-total')) $('nav-ig-total').textContent = total.toLocaleString();
  if ($('nav-all')) $('nav-all').textContent = total.toLocaleString();
  if ($('nav-shortlist')) $('nav-shortlist').textContent = savedCount.toLocaleString();
  if ($('nav-applied')) $('nav-applied').textContent = appliedCount.toLocaleString();

  renderCountryBars(countries, total || 1);

  const bestFit = [...state.jobs].sort((a,b) => scoreOf(b) - scoreOf(a))[0];
  if (bestFit) {
    $('focus-score').innerHTML = `${scoreOf(bestFit)}<small>%</small>`;
    $('focus-title').textContent = bestFit['Job Title'] || 'Featured Opportunity';
    $('focus-company').textContent = `${bestFit.Company || 'Company'} · ${bestFit.Location || bestFit.Country} (${bestFit.Source || 'Indeed/Glassdoor'})`;
    $('view-best-fit').onclick = () => focusJob(bestFit);
  }
}

const COUNTRY_FLAGS = {
  'United Kingdom': '🇬🇧', 'Canada': '🇨🇦', 'United Arab Emirates': '🇦🇪',
  'Saudi Arabia': '🇸🇦', 'Qatar': '🇶🇦', 'Netherlands': '🇳🇱', 'Ireland': '🇮🇪',
  'Sweden': '🇸🇪', 'Denmark': '🇩🇰', 'Finland': '🇫🇮', 'Australia': '🇦🇺',
  'Singapore': '🇸🇬', 'New Zealand': '🇳🇿', 'Kuwait': '🇰🇼', 'Bahrain': '🇧🇭',
  'Oman': '🇴🇲', 'France': '🇫🇷', 'Portugal': '🇵🇹', 'Poland': '🇵🇱',
  'Belgium': '🇧🇪', 'Austria': '🇦🇹', 'Malaysia': '🇲🇾'
};

const COUNTRY_ISO = {
  'United Kingdom': 'GB', 'Canada': 'CA', 'United Arab Emirates': 'AE',
  'Saudi Arabia': 'SA', 'Qatar': 'QA', 'Netherlands': 'NL', 'Ireland': 'IE',
  'Sweden': 'SE', 'Denmark': 'DK', 'Finland': 'FI', 'Australia': 'AU',
  'Singapore': 'SG', 'New Zealand': 'NZ', 'Kuwait': 'KW', 'Bahrain': 'BH',
  'Oman': 'OM', 'France': 'FR', 'Portugal': 'PT', 'Poland': 'PL',
  'Belgium': 'BE', 'Austria': 'AT', 'Malaysia': 'MY'
};

let countryChartSort = 'volume';
let countryChartQuery = '';

function renderCountryBars(countries, totalJobs) {
  const container = $('country-bars');
  if (!container) return;

  const filtered = countries.filter(([name]) => !countryChartQuery || name.toLowerCase().includes(countryChartQuery.toLowerCase()));
  if (countryChartSort === 'alpha') {
    filtered.sort((a, b) => a[0].localeCompare(b[0]));
  } else {
    filtered.sort((a, b) => b[1] - a[1]);
  }

  const maxCount = Math.max(...countries.map(([, c]) => c), 1);
  container.innerHTML = filtered.map(([name, count]) => {
    const pct = totalJobs > 0 ? ((count / totalJobs) * 100).toFixed(1) : 0;
    const barWidth = Math.max(4, Math.round((count / maxCount) * 100));
    const isSelected = state.filters.country && state.filters.country.toLowerCase() === name.toLowerCase();
    const isoCode = COUNTRY_ISO[name] || name.slice(0, 2).toUpperCase();

    return `
      <div class="country-row ${isSelected ? 'selected' : ''}" data-country="${esc(name)}">
        <div class="country-name-badge">
          <span class="country-iso-tag">${isoCode}</span>
          <span class="country-full-name">${esc(name)}</span>
        </div>
        <div class="country-bar-track">
          <div class="country-bar-fill" style="width: ${barWidth}%;"></div>
        </div>
        <span class="country-metric-group">
          <span class="country-row-count">${count.toLocaleString()}</span>
          <span class="country-row-pct">${pct}%</span>
        </span>
      </div>
    `;
  }).join('');

  container.querySelectorAll('.country-row').forEach((row) => {
    row.onclick = () => {
      const selected = row.dataset.country;
      state.filters.country = (state.filters.country.toLowerCase() === selected.toLowerCase()) ? '' : selected;
      if ($('country-filter')) $('country-filter').value = state.filters.country;
      state.page = 1;
      renderAll();
    };
  });
}

function renderTable() {
  applyFilters();
  const total = state.filtered.length;
  const pages = Math.max(1, Math.ceil(total / PAGE_SIZE));
  if (state.page > pages) state.page = pages;
  const slice = state.filtered.slice((state.page - 1) * PAGE_SIZE, state.page * PAGE_SIZE);
  const saved = shortlist();

  $('visible-count').textContent = total.toLocaleString();
  $('jobs-body').innerHTML = slice.length ? slice.map((job) => {
    const score = scoreOf(job);
    const work = job['Remote / Workplace'] || 'Not stated';
    const remote = /remote/i.test(work);
    const postedDate = formatDisplayDate(job['Posted Date'] || job.PostedDate);
    const crawlDate = formatDisplayDate(job['Crawl Date'] || job.Date);
    const isApplied = String(job['Applied Status'] || '').toLowerCase() === 'yes';
    const source = (job.Source || 'Indeed').trim();
    const sourceClass = source.toLowerCase().includes('glassdoor') ? 'source-glassdoor' : 'source-indeed';

    return `<tr>
    <td class="bookmark-cell"><button class="bookmark ${saved.has(keyFor(job)) ? 'active' : ''}" data-save="${esc(keyFor(job))}">${saved.has(keyFor(job)) ? '★' : '☆'}</button></td>
    <td class="col-role"><span class="role-title">${esc(job['Job Title'] || 'Untitled role')}</span><span class="role-sub">${esc(job.Company || 'Not stated')}</span></td>
    <td class="col-source"><span class="source-badge ${sourceClass}">${esc(source)}</span></td>
    <td class="col-country">${esc(job.Country || 'Global')}</td>
    <td class="col-posted"><span class="track">${esc(postedDate)}</span></td>
    <td class="col-crawl"><span class="track">${esc(crawlDate)}</span></td>
    <td class="col-workplace"><span class="work-pill ${remote ? 'remote' : ''}">${esc(work)}</span></td>
    <td class="col-match"><span class="match-pill ${score ? `score-${score}` : 'score-none'}">${score ? score + '%' : '—'}</span></td>
    <td class="col-applied">
      <button class="applied-badge ${isApplied ? 'applied' : ''}" data-applied="${esc(keyFor(job))}">
        ${isApplied ? '✓ Applied' : 'Mark Applied'}
      </button>
    </td>
    <td class="col-link open-cell">${job['Job URL'] ? `<a class="open-link" href="${esc(job['Job URL'])}" target="_blank" rel="noopener" title="Open job posting">↗</a>` : '—'}</td></tr>`;
  }).join('') : `<tr><td colspan="10" class="empty">No Indeed or Glassdoor opportunities match these filters.</td></tr>`;

  document.querySelectorAll('[data-save]').forEach((b) => b.onclick = () => {
    const keys = shortlist(), key = b.dataset.save;
    keys.has(key) ? keys.delete(key) : keys.add(key);
    saveShortlist(keys);
    renderAll();
  });

  document.querySelectorAll('[data-applied]').forEach((b) => b.onclick = async () => {
    const key = b.dataset.applied;
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

    try { await storeJobs([targetJob]); } catch {}
    renderAll();
    showToast(newStatus === 'Yes' ? `Marked "${targetJob['Job Title']}" as Applied.` : `Marked "${targetJob['Job Title']}" as Not Applied.`);
  });

  $('page-info').textContent = total ? `Showing ${(state.page-1)*PAGE_SIZE+1}–${Math.min(state.page*PAGE_SIZE,total)} of ${total.toLocaleString()} roles` : 'No roles to show';
  $('previous-page').disabled = state.page === 1;
  $('next-page').disabled = state.page === pages;
  $('page-buttons').innerHTML = pagination(pages);
  document.querySelectorAll('[data-page]').forEach((b) => b.onclick = () => { state.page = Number(b.dataset.page); renderTable(); });
}

function pagination(pages) {
  const set = new Set([1, pages, state.page - 1, state.page, state.page + 1].filter((n) => n >= 1 && n <= pages));
  let previous = 0;
  return [...set].sort((a, b) => a - b).map((n) => `${n - previous > 1 ? '<span>…</span>' : ''}<button class="${n === state.page ? 'active' : ''}" data-page="${n}">${n}</button>`).join('');
}

function renderFilters() {
  const labels = { country: 'Country', source: 'Source', score: 'Match ≥', workplace: 'Workplace', posted: 'Job Posted', recent: 'Crawl Date', status: 'Status' };
  const active = Object.entries(state.filters).filter(([, v]) => v);
  if ($('filter-count')) $('filter-count').textContent = active.length;
  if ($('filter-button')) $('filter-button').classList.toggle('has-filters', active.length > 0);
  if ($('active-filters')) {
    $('active-filters').innerHTML = active.map(([k, v]) => `<span class="filter-chip">${labels[k]}: ${esc(v)} <button data-clear="${k}">×</button></span>`).join('');
    document.querySelectorAll('[data-clear]').forEach((b) => b.onclick = () => {
      state.filters[b.dataset.clear] = '';
      if ($(`${b.dataset.clear}-filter`)) $(`${b.dataset.clear}-filter`).value = '';
      state.page = 1;
      renderAll();
    });
  }
}

function renderAll() {
  renderMetrics();
  renderFilters();
  renderTable();
  updateStoredCount();
}

function focusJob(job) {
  state.query = job['Job Title'] || '';
  if ($('search')) $('search').value = state.query;
  applyFilters();
  const exact = state.filtered.findIndex((x) => keyFor(x) === keyFor(job));
  state.page = Math.floor(Math.max(exact, 0) / PAGE_SIZE) + 1;
  renderTable();
  $('roles').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function showToast(msg) {
  const toast = $('toast');
  toast.textContent = msg;
  toast.classList.add('show');
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => toast.classList.remove('show'), 3500);
}


/* ══════════════════════════════════════════════════════════════════════
   AUTHENTICATION GATE (SHA-256)
   ══════════════════════════════════════════════════════════════════════ */

const MASTER_PASSWORD_HASH = '5a8dc1ec9f6708f0e7071d8fbf7bb455c0edd294046c5d0a7d9dbf72f2a16f4b';
const ALT_PASSWORD_HASH = 'e2a23afe0cdeccbaf4fab5e2387a134d32ba1064d96d11454af926917e2a5383';
const MASTER_PASSWORD_SALT = 'jobcompass_salt_2026';
const AUTH_SESSION_KEY = 'job-compass-auth-active';

async function hashPassword(password, salt) {
  const encoder = new TextEncoder();
  const data = encoder.encode(password + '::' + salt);
  const hashBuffer = await crypto.subtle.digest('SHA-256', data);
  return Array.from(new Uint8Array(hashBuffer)).map((b) => b.toString(16).padStart(2, '0')).join('');
}

function isAuthActive() { return sessionStorage.getItem(AUTH_SESSION_KEY) === 'true'; }

function showAuthModal() {
  if (!$('auth-overlay')) {
    unlockDashboard();
    return;
  }
  if ($('app-shell')) {
    $('app-shell').style.filter = 'blur(4px)';
    $('app-shell').style.pointerEvents = 'none';
  }
  if ($('auth-overlay')) $('auth-overlay').hidden = false;
}

function unlockDashboard() {
  sessionStorage.setItem(AUTH_SESSION_KEY, 'true');
  if ($('auth-overlay')) $('auth-overlay').hidden = true;
  if ($('app-shell')) {
    $('app-shell').style.filter = 'none';
    $('app-shell').style.pointerEvents = 'auto';
  }
  initData();
}

function lockDashboard() {
  sessionStorage.removeItem(AUTH_SESSION_KEY);
  showAuthModal();
}

function initAuth() {
  const authForm = $('auth-form');
  if (authForm) {
    authForm.onsubmit = async (e) => {
      e.preventDefault();
      const p = ($('auth-password').value || '').trim();
      const h = await hashPassword(p, MASTER_PASSWORD_SALT);
      if (h === MASTER_PASSWORD_HASH || h === ALT_PASSWORD_HASH) {
        unlockDashboard();
      } else {
        $('auth-error').textContent = 'Incorrect master password.';
        $('auth-error').hidden = false;
      }
    };
  }
  if ($('lock-header-btn')) $('lock-header-btn').onclick = lockDashboard;
}


/* ══════════════════════════════════════════════════════════════════════
   DATA INITIALIZATION
   ══════════════════════════════════════════════════════════════════════ */

async function initData() {
  // 1. Try fetching fresh CSV first
  if (location.protocol !== 'file:') {
    try {
      const resp = await fetch(SOURCE_FILE);
      if (resp.ok) {
        const text = await resp.text();
        const jobs = parseCSV(text);
        if (jobs.length) {
          const deduped = deduplicateJobsArray(jobs);
          await storeJobs(deduped);
          loadJobs(deduped, SOURCE_FILE);
          return;
        }
      }
    } catch (err) {
      console.warn('[IG Board] Could not fetch CSV:', err);
    }
  }

  // 2. Fallback to stored IndexedDB
  try {
    const stored = await loadStoredJobs();
    if (stored.length > 0) {
      loadJobs(stored, 'stored data');
      return;
    }
  } catch {}

  loadJobs([], 'clean board');
}

function initUIListeners() {
  if ($('filter-button')) {
    $('filter-button').onclick = () => {
      const panel = $('filters-panel');
      panel.hidden = !panel.hidden;
    };
  }

  if ($('search')) {
    $('search').oninput = (e) => {
      state.query = e.target.value;
      state.page = 1;
      renderTable();
    };
  }

  ['country', 'source', 'score', 'workplace', 'posted', 'recent', 'status'].forEach((k) => {
    const sel = $(`${k}-filter`);
    if (sel) {
      sel.onchange = (e) => {
        state.filters[k] = e.target.value;
        state.page = 1;
        renderAll();
      };
    }
  });

  if ($('clear-filters')) {
    $('clear-filters').onclick = () => {
      state.filters = { country: '', source: '', score: '', workplace: '', posted: '', recent: '', status: '' };
      ['country', 'source', 'score', 'workplace', 'posted', 'recent', 'status'].forEach((k) => {
        if ($(`${k}-filter`)) $(`${k}-filter`).value = '';
      });
      state.page = 1;
      renderAll();
    };
  }

  document.querySelectorAll('[data-status-tab]').forEach((tab) => {
    tab.onclick = () => {
      document.querySelectorAll('[data-status-tab]').forEach((t) => t.classList.remove('active'));
      tab.classList.add('active');
      const t = tab.dataset.statusTab;
      if (t === 'all') state.filters.status = '';
      else if (t === 'high') { state.filters.status = ''; state.filters.score = '85'; }
      else state.filters.status = t;
      state.page = 1;
      renderAll();
    };
  });

  const fileInput = $('file-input');
  if ($('import-button')) $('import-button').onclick = () => fileInput.click();
  if ($('sidebar-import-button')) $('sidebar-import-button').onclick = () => fileInput.click();

  if (fileInput) {
    fileInput.onchange = async (e) => {
      const file = e.target.files[0];
      if (!file) return;
      const text = await file.text();
      const jobs = parseCSV(text);
      if (jobs.length) {
        const deduped = deduplicateJobsArray(jobs);
        await storeJobs(deduped);
        loadJobs(deduped, file.name);
      }
    };
  }

  if ($('clear-data-button')) {
    $('clear-data-button').onclick = async () => {
      if (confirm('Clear all stored Indeed & Glassdoor opportunities from this browser?')) {
        await clearStoredJobs();
        loadJobs([], 'cleared data');
      }
    };
  }
}

(function init() {
  initAuth();
  initUIListeners();
  if ($('auth-overlay') && !isAuthActive()) {
    showAuthModal();
  } else {
    unlockDashboard();
  }
})();
