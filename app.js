/* ══════════════════════════════════════════════════════════════════════
   Job Compass — app.js
   Persistent Excel/CSV import with deduplication via IndexedDB
   ══════════════════════════════════════════════════════════════════════ */

const SOURCE_FILE = 'full_crawl_jobs_20260804_103131.csv';
const PAGE_SIZE = 12;
const DB_NAME = 'job-compass-db';
const DB_VERSION = 2;
const STORE_NAME = 'jobs';

const state = { jobs: [], filtered: [], page: 1, sort: { key: 'Match Score', asc: false }, filters: { country: '', score: '', workplace: '', recent: '', status: '' }, query: '', compact: false };
const $ = (id) => document.getElementById(id);
const esc = (value = '') => String(value).replace(/[&<>"']/g, (s) => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'\&#039;' }[s]));
const keyFor = (job) => job['Job URL'] || `${job.Company}|${job['Job Title']}|${job.Location}`;
const shortlist = () => new Set(JSON.parse(localStorage.getItem('job-compass-shortlist') || '[]'));
const saveShortlist = (keys) => localStorage.setItem('job-compass-shortlist', JSON.stringify([...keys]));


/* ══════════════════════════════════════════════════════════════════════
   IndexedDB Persistence Layer
   ══════════════════════════════════════════════════════════════════════ */

function openDB() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = (e) => {
      const db = e.target.result;
      // Drop old store if it exists (schema change)
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
      store.put(job);  // put = upsert on keyPath _dedupKey
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
   Deduplication
   ══════════════════════════════════════════════════════════════════════ */

/** Generate a dedup key for a job: prefer Job URL, else hash Company+Title+Location */
function dedupKey(job) {
  const url = (job['Job URL'] || '').trim();
  if (url) return `url:${url}`;
  const company = (job.Company || '').trim().toLowerCase();
  const title = (job['Job Title'] || '').trim().toLowerCase();
  const location = (job.Location || '').trim().toLowerCase();
  return `combo:${company}|${title}|${location}`;
}

/** Tag each job with a _dedupKey field (used as IndexedDB keyPath) */
function tagJobs(jobs) {
  return jobs.map((job) => ({ ...job, _dedupKey: dedupKey(job) }));
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
  if (!raw || /^unknown$/i.test(raw)) return null;
  const parsed = new Date(raw);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}
function dateOf(job) { return parseDate(job['Posted Date']); }
function latestCrawlDate(jobs) { return jobs.reduce((latest, job) => { const date = parseDate(job.Date); return date && (!latest || date > latest) ? date : latest; }, null) || new Date(); }
function countBy(list, field) { return [...list.reduce((m, x) => m.set(x[field] || 'Unknown', (m.get(x[field] || 'Unknown') || 0) + 1), new Map())].sort((a,b) => b[1] - a[1]); }
function titleCase(value) { return value || 'Not specified'; }
function roleText(job) { return [job['Job Title'], job.Company, job.Location, job['Required Skills'], job['Search Keyword']].join(' ').toLowerCase(); }

function loadJobs(jobs, sourceLabel = 'your export') {
  state.jobs = jobs; state.page = 1;
  state.filters = { country: '', score: '', workplace: '', recent: '', status: '' }; state.query = ''; $('search').value = '';
  populateFilters(); renderAll(); updateStoredCount();
  showToast(`Loaded ${jobs.length.toLocaleString()} opportunities from ${sourceLabel}.`);
}

function populateFilters() {
  const update = (id, values, label) => { const select = $(id); const current = select.value; select.innerHTML = `<option value="">${label}</option>` + values.map((x) => `<option value="${esc(x)}">${esc(x)}</option>`).join(''); select.value = current; };
  update('country-filter', countBy(state.jobs, 'Country').map(([x]) => x), 'All countries');
  update('workplace-filter', countBy(state.jobs, 'Remote / Workplace').map(([x]) => x), 'Any workplace');
}

function applyFilters() {
  const s = shortlist(), f = state.filters, q = state.query.toLowerCase();
  const referenceDate = latestCrawlDate(state.jobs);
  state.filtered = state.jobs.filter((job) => {
    const applied = (job['Applied Status'] || '').toLowerCase();
    const posted = dateOf(job);
    const ageInDays = posted ? (referenceDate - posted) / 86400000 : Infinity;
    const isRecent = !f.recent || (ageInDays >= 0 && ageInDays <= Number(f.recent));
    return (!q || roleText(job).includes(q)) && (!f.country || job.Country === f.country) && (!f.workplace || job['Remote / Workplace'] === f.workplace) && (!f.score || scoreOf(job) >= Number(f.score)) && isRecent && (!f.status || (f.status === 'shortlisted' && s.has(keyFor(job))) || (f.status === 'applied' && applied === 'yes') || (f.status === 'not-applied' && applied !== 'yes'));
  });
  const { key, asc } = state.sort;
  state.filtered.sort((a,b) => { const av = key === 'Match Score' ? scoreOf(a) : (a[key] || '').toLowerCase(); const bv = key === 'Match Score' ? scoreOf(b) : (b[key] || '').toLowerCase(); return av < bv ? (asc ? -1 : 1) : av > bv ? (asc ? 1 : -1) : 0; });
}

function renderMetrics() {
  const all = state.jobs, saved = shortlist(), high = all.filter((j) => scoreOf(j) >= 85), countries = countBy(all, 'Country');
  $('metric-total').textContent = all.length.toLocaleString(); $('metric-total-note').textContent = `${countBy(all, 'Company').length.toLocaleString()} companies represented`;
  $('metric-high-match').textContent = high.length.toLocaleString(); $('metric-shortlisted').textContent = saved.size.toLocaleString(); $('metric-shortlist-note').textContent = saved.size ? 'Your roles to revisit' : 'Start saving standout roles';
  $('metric-countries').textContent = countries.length; $('metric-country-note').textContent = countries.slice(0,2).map(([x]) => x).join(' & ') + (countries.length > 2 ? ' + more' : '');
  $('nav-shortlist').textContent = saved.size;
  $('dataset-summary').textContent = `${all.length.toLocaleString()} opportunities across ${countries.length} countries — filter the noise, keep what matters, and make every application count.`;
  const top = countries.slice(0, 5), max = top[0]?.[1] || 1;
  $('country-bars').innerHTML = top.map(([country, n]) => `<div class="country-row"><span>${esc(country)}</span><div class="bar"><i style="width:${(n/max*100).toFixed(1)}%"></i></div><b>${n}</b></div>`).join('');
  const best = all.slice().sort((a,b) => scoreOf(b) - scoreOf(a))[0];
  if (best) { $('focus-score').innerHTML = `${scoreOf(best)}<small>%</small>`; $('focus-title').textContent = best['Job Title']; $('focus-company').textContent = `${best.Company || 'Company not stated'} · ${best.Location || best.Country || 'Location not stated'}`; $('view-best-fit').onclick = () => focusJob(best); }
}

function renderFilters() {
  const labels = { country: 'Country', score: 'Match ≥', workplace: 'Workplace', recent: 'Added', status: 'Status' };
  const active = Object.entries(state.filters).filter(([,v]) => v);
  $('filter-count').textContent = active.length; $('filter-button').classList.toggle('has-filters', active.length > 0);
  $('active-filters').innerHTML = active.map(([key,value]) => `<span class="filter-chip">${labels[key]}: ${esc(key === 'score' ? value + '%' : key === 'recent' ? `last ${value} days` : value)} <button data-clear="${key}" aria-label="Clear ${labels[key]}">×</button></span>`).join('');
  document.querySelectorAll('[data-clear]').forEach((b) => b.onclick = () => { state.filters[b.dataset.clear] = ''; $(`${b.dataset.clear}-filter`).value = ''; state.page = 1; renderAll(); });
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

function fetchJobsViaJSONP(webhookUrl) {
  return new Promise((resolve, reject) => {
    const callbackName = 'gsCallback_' + Math.random().toString(36).substring(2);
    const script = document.createElement('script');
    const timer = setTimeout(() => {
      cleanup();
      reject(new Error('JSONP timeout'));
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

    script.onerror = (err) => {
      cleanup();
      reject(err);
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
    const crawlDate = esc(job['Crawl Date'] || job.Date || '—');
    const isApplied = String(job['Applied Status'] || '').toLowerCase() === 'yes';
    return `<tr>
    <td class="bookmark-cell"><button class="bookmark ${saved.has(keyFor(job)) ? 'active' : ''}" data-save="${esc(keyFor(job))}" title="${saved.has(keyFor(job)) ? 'Remove from shortlist' : 'Add to shortlist'}">${saved.has(keyFor(job)) ? '★' : '☆'}</button></td>
    <td><span class="track">${crawlDate}</span></td>
    <td><span class="role-title">${esc(job['Job Title'] || 'Untitled role')}</span><span class="role-sub">${esc(job.Country || '')}</span></td>
    <td><span class="company">${esc(job.Company || 'Not stated')}</span></td>
    <td class="location">${esc(job.Location || job.City || 'Not stated')}</td>
    <td><span class="match-pill ${score ? `score-${score}` : 'score-none'}">${score ? score + '%' : '—'}</span></td>
    <td><span class="work-pill ${remote ? 'remote' : ''}">${esc(work)}</span></td>
    <td>
      <button class="applied-badge ${isApplied ? 'applied' : ''}" data-applied="${esc(keyFor(job))}">
        ${isApplied ? '✓ Applied' : 'Mark Applied'}
      </button>
    </td>
    <td class="open-cell">${job['Job URL'] ? `<a class="open-link" href="${esc(job['Job URL'])}" target="_blank" rel="noopener" title="Open job posting">↗</a>` : '—'}</td></tr>`;
  }).join('') : `<tr><td colspan="9" class="empty">No opportunities match these filters.<br/><button class="clear-filters" id="empty-clear">Clear filters</button></td></tr>`;

  document.querySelectorAll('[data-save]').forEach((button) => button.onclick = () => { const keys = shortlist(), key = button.dataset.save; const adding = !keys.has(key); adding ? keys.add(key) : keys.delete(key); saveShortlist(keys); renderAll(); showToast(adding ? 'Added to your shortlist.' : 'Removed from your shortlist.'); });

  document.querySelectorAll('[data-applied]').forEach((button) => button.onclick = async () => {
    const key = button.dataset.applied;
    const targetJob = state.jobs.find((j) => keyFor(j) === key);
    if (!targetJob) return;
    
    const currentlyApplied = String(targetJob['Applied Status'] || '').toLowerCase() === 'yes';
    const newStatus = currentlyApplied ? 'No' : 'Yes';
    targetJob['Applied Status'] = newStatus;
    
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
function clearFilters() { state.filters = { country:'',score:'',workplace:'',recent:'',status:'' }; state.query = ''; state.page = 1; $('search').value=''; ['country','score','workplace','recent','status'].forEach((id)=>$(`${id}-filter`).value=''); renderAll(); }
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


/* ══════════════════════════════════════════════════════════════════════
   Startup & Event Wiring
   ══════════════════════════════════════════════════════════════════════ */

['import-button','hero-import'].forEach((id) => $(id).onclick = () => $('file-input').click());
$('file-input').onchange = (e) => { importFile(e.target.files[0]); e.target.value = ''; };
$('search').oninput = (e) => { state.query = e.target.value; state.page = 1; renderAll(); };
['country','score','workplace','recent','status'].forEach((id) => $(`${id}-filter`).onchange = (e) => { state.filters[id] = e.target.value; state.page = 1; renderAll(); });
$('filter-button').onclick = () => { const panel=$('filters-panel'); panel.hidden=!panel.hidden; }; $('clear-filters').onclick=clearFilters;
document.querySelectorAll('th[data-sort]').forEach((th) => th.onclick = () => { const key=th.dataset.sort; state.sort={key,asc: state.sort.key===key ? !state.sort.asc : key !== 'Match Score'}; renderTable(); });
$('previous-page').onclick=()=>{ if(state.page>1){state.page--;renderTable();} }; $('next-page').onclick=()=>{state.page++;renderTable();};
$('table-view-button').onclick=()=>{state.compact=false;renderAll();}; $('compact-view-button').onclick=()=>{state.compact=true;renderAll();};
$('theme-button').onclick=()=>document.body.classList.toggle('dark'); $('mobile-menu').onclick=()=>document.querySelector('.sidebar').classList.toggle('open');
$('clear-data').onclick = handleClearData;

function fetchJobsViaJSONP(webhookUrl) {
  return new Promise((resolve, reject) => {
    const callbackName = 'gsCallback_' + Math.random().toString(36).substring(2);
    const script = document.createElement('script');
    const timer = setTimeout(() => {
      cleanup();
      reject(new Error('JSONP timeout'));
    }, 15000);

    const cleanup = () => {
      clearTimeout(timer);
      delete window[callbackName];
      if (script.parentNode) script.parentNode.removeChild(script);
    };

    window[callbackName] = (data) => {
      cleanup();
      resolve(data);
    };

    script.onerror = (err) => {
      cleanup();
      reject(err);
    };

    const separator = webhookUrl.includes('?') ? '&' : '?';
    script.src = `${webhookUrl}${separator}action=get_jobs&callback=${callbackName}`;
    document.head.appendChild(script);
  });
}

async function fetchJobsFromGoogleSheet(silent = false) {
  let webhookUrl = getSheetsWebhookUrl();
  if (!webhookUrl) {
    if (!silent) showToast('Please enter your Google Apps Script Webhook URL first.');
    return false;
  }

  if (!silent) showLoading('Fetching latest job data from Google Sheets…');
  try {
    let data = null;

    // 1. Primary: JSONP script injection (bypasses CORS & 302 redirects on file:// / localhost)
    try {
      data = await fetchJobsViaJSONP(webhookUrl);
    } catch {
      // 2. Fallback: GET fetch
      try {
        const fetchUrl = webhookUrl + (webhookUrl.includes('?') ? '&' : '?') + 'action=get_jobs';
        const res = await fetch(fetchUrl, { method: 'GET', redirect: 'follow' });
        const text = await res.text();
        data = JSON.parse(text);
      } catch {
        // 3. Fallback: POST fetch
        const postRes = await fetch(webhookUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify({ action: 'get_jobs' })
        });
        data = await postRes.json();
      }
    }

    if (data && data.success && Array.isArray(data.jobs) && data.jobs.length > 0) {
      const normalizedJobs = data.jobs.map((j) => ({
        ...j,
        'Job URL': String(j['Job URL'] || j.JobURL || j.url || '').trim(),
        'Company': String(j.Company || j.company || 'N/A').trim(),
        'Job Title': String(j['Job Title'] || j.JobTitle || j.title || 'Role').trim(),
        'Country': String(j.Country || j.country || 'Global').trim(),
        'Location': String(j.Location || j.location || j.Country || '').trim(),
        'Crawl Date': String(j['Crawl Date'] || j.CrawlDate || j.Date || '').trim(),
        'Applied Status': String(j['Applied Status'] || j.AppliedStatus || j.Applied || 'No').trim(),
        'Match Score': String(j['Match Score'] !== undefined ? j['Match Score'] : (j.MatchScore !== undefined ? j.MatchScore : '85%')).trim(),
        'Remote / Workplace': String(j['Remote / Workplace'] || j.Workplace || 'On-site / Hybrid').trim(),
        'Easy Apply': String(j['Easy Apply'] || j.EasyApply || 'No').trim(),
        'Visa Sponsorship Mentioned': String(j['Visa Sponsorship'] || j['Visa Sponsorship Mentioned'] || 'No').trim(),
        'Required Skills': String(j['Required Skills'] || j.Skills || '').trim(),
        'Notes': String(j.Notes || '').trim()
      })).filter((j) => j['Job Title'] && j['Job Title'] !== 'Job Title');

      if (normalizedJobs.length > 0) {
        const tagged = tagJobs(normalizedJobs);
        await storeJobs(tagged);
        loadJobs(tagged, 'Google Sheets');
        if (!silent) showToast(`Loaded ${tagged.length.toLocaleString()} opportunities live from Google Sheets!`);
        return true;
      }
    }
    
    if (!silent) showToast('Google Sheets returned 0 jobs or empty response.');
  } catch (err) {
    console.warn('[Google Sheets] Fetch failed:', err);
    if (!silent) showToast('Could not fetch from Google Sheets — check network connection.');
  } finally {
    if (!silent) hideLoading();
  }
  return false;
}

/* ══════════════════════════════════════════════════════════════════════
   AUTHENTICATION & SECURITY LAYER (Web Crypto API SHA-256)
   ══════════════════════════════════════════════════════════════════════ */

const MASTER_PASSWORD_HASH = 'e2a23afe0cdeccbaf4fab5e2387a134d32ba1064d96d11454af926917e2a5383';
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

/** Show authentication modal */
function showAuthModal() {
  if ($('app-shell')) $('app-shell').hidden = true;
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
  if ($('app-shell')) $('app-shell').hidden = false;
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

      if (computedHash === MASTER_PASSWORD_HASH || (customHash && computedHash === customHash)) {
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

  // 1. Try pulling live data from Google Sheets
  const pulledFromSheets = await fetchJobsFromGoogleSheet(true);
  if (pulledFromSheets) return;

  // 2. Fallback to stored IndexedDB
  try {
    const storedJobs = await loadStoredJobs();
    if (storedJobs.length > 0) {
      loadJobs(storedJobs, 'stored data');
      return;
    }
  } catch { /* IndexedDB unavailable */ }

  // 3. Fallback: fetch the default CSV
  if (location.protocol !== 'file:') {
    try {
      const resp = await fetch(SOURCE_FILE);
      if (!resp.ok) throw new Error('Not found');
      const text = await resp.text();
      const jobs = parseCSV(text);
      if (jobs.length) {
        const tagged = tagJobs(jobs);
        await storeJobs(tagged);
        loadJobs(tagged, SOURCE_FILE);
        return;
      }
    } catch { /* CSV not available */ }
  }

  // Nothing to load
  $('dataset-summary').textContent = 'Import your CSV or Excel job export to start ranking and shortlisting roles.';
  $('jobs-body').innerHTML = '<tr><td colspan="9" class="empty">Your dashboard is ready. Import a CSV or Excel file to begin.</td></tr>';
  updateStoredCount();
}

(function init() {
  initAuth();
  if (!isAuthActive()) {
    showAuthModal();
  } else {
    unlockDashboard();
  }
})();
