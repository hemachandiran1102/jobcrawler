/* ══════════════════════════════════════════════════════════════════════
   Job Compass — app_inbound_pipeline.js
   Dedicated Inbound Interview & Next Steps Pipeline Dashboard
   ══════════════════════════════════════════════════════════════════════ */

const state = {
  inbound: [],
  selectedAccount: 'all',
  selectedTab: 'next-steps',
  searchQuery: '',
};

const $ = (id) => document.getElementById(id);
const esc = (value = '') => String(value).replace(/[&<>"']/g, (s) => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#039;' }[s]));

function showToast(message) {
  const toast = $('toast');
  if (!toast) return;
  toast.textContent = message;
  toast.classList.add('show');
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => toast.classList.remove('show'), 3500);
}

// ── Authentication Gate ──
const MASTER_PASSWORD_HASH = '5a8dc1ec9f6708f0e7071d8fbf7bb455c0edd294046c5d0a7d9dbf72f2a16f4b';
const ALT_PASSWORD_HASH = 'e2a23afe0cdeccbaf4fab5e2387a134d32ba1064d96d11454af926917e2a5383';
const MASTER_PASSWORD_SALT = 'jobcompass_salt_2026';
const AUTH_SESSION_KEY = 'job-compass-auth-active';

async function hashPassword(password, salt) {
  const encoder = new TextEncoder();
  const data = encoder.encode(password + '::' + salt);
  const hashBuffer = await crypto.subtle.digest('SHA-256', data);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map((b) => b.toString(16).padStart(2, '0')).join('');
}

function isAuthActive() {
  return sessionStorage.getItem(AUTH_SESSION_KEY) === 'true';
}

function showAuthModal() {
  const overlay = $('auth-overlay');
  if (overlay) overlay.hidden = false;
  if ($('app-shell')) $('app-shell').style.filter = 'blur(4px)';
}

function unlockDashboard() {
  sessionStorage.setItem(AUTH_SESSION_KEY, 'true');
  const overlay = $('auth-overlay');
  if (overlay) overlay.hidden = true;
  if ($('app-shell')) $('app-shell').style.filter = 'none';
  loadPipelineData();
}

function lockDashboard() {
  sessionStorage.removeItem(AUTH_SESSION_KEY);
  showAuthModal();
  showToast('Dashboard locked.');
}

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
        showToast('Inbound Pipeline unlocked!');
      } else {
        $('auth-error').textContent = 'Incorrect master password. Access denied.';
        $('auth-error').hidden = false;
      }
    };
  }

  const lockBtn = $('lock-header-btn');
  if (lockBtn) lockBtn.onclick = lockDashboard;
}

// ── Data Loading & Parsing ──
async function loadPipelineData() {
  try {
    let items = [];
    try {
      const resp = await fetch('interview_pipeline.json');
      if (resp.ok) {
        items = await resp.json();
      }
    } catch {}

    if (!items || !items.length) {
      try {
        const resp2 = await fetch('interview_pipeline.csv');
        if (resp2.ok) {
          const text = await resp2.text();
          items = parseCSV(text);
        }
      } catch {}
    }

    state.inbound = items || [];
    renderAll();
  } catch (err) {
    console.error('Error loading pipeline data:', err);
  }
}

function parseCSV(text) {
  const lines = text.split('\n').map((l) => l.trim()).filter(Boolean);
  if (lines.length < 2) return [];
  const headers = lines[0].split(',').map((h) => h.replace(/^["']|["']$/g, '').trim());
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    // Regex for CSV with quoted strings
    const match = lines[i].match(/(".*?"|[^",\s]+)(?=\s*,|\s*$)/g);
    if (!match) continue;
    const obj = {};
    headers.forEach((h, idx) => {
      obj[h] = (match[idx] || '').replace(/^["']|["']$/g, '').trim();
    });
    if (obj['Company'] || obj['Subject']) rows.push(obj);
  }
  return rows;
}

// ── Rendering Engine ──
function renderMetrics() {
  const items = state.inbound || [];
  const nextSteps = items.filter((d) => d['Is Next Step'] === true || d['Is Next Step'] === 'True' || d['Is Next Step'] === 'true');
  const invites = items.filter((d) => d.Category === 'Interview Invitation');
  const assessments = items.filter((d) => d.Category === 'Technical Assessment');
  const inquiries = items.filter((d) => (d.Category || '').includes('Availability') || (d.Category || '').includes('Inquiry'));
  const rejections = items.filter((d) => d.Category === 'Rejection');

  if ($('metric-action-count')) $('metric-action-count').textContent = nextSteps.length;
  if ($('metric-invites-count')) $('metric-invites-count').textContent = invites.length;
  if ($('metric-assessments-count')) $('metric-assessments-count').textContent = assessments.length;
  if ($('metric-inquiries-count')) $('metric-inquiries-count').textContent = inquiries.length;
  if ($('metric-total-emails')) $('metric-total-emails').textContent = items.length;

  if ($('header-total-count')) $('header-total-count').textContent = `${nextSteps.length} Actionable Next Steps`;
  if ($('nav-inbound-badge')) $('nav-inbound-badge').textContent = nextSteps.length;

  // Tab count badges
  if ($('tab-action-count')) $('tab-action-count').textContent = nextSteps.length;
  if ($('tab-invites-count')) $('tab-invites-count').textContent = invites.length;
  if ($('tab-assessments-count')) $('tab-assessments-count').textContent = assessments.length;
  if ($('tab-inquiries-count')) $('tab-inquiries-count').textContent = inquiries.length;
  if ($('tab-all-count')) $('tab-all-count').textContent = items.length;
  if ($('tab-rejections-count')) $('tab-rejections-count').textContent = rejections.length;

  // Account count badges
  if ($('acc-count-all')) $('acc-count-all').textContent = items.length;
  if ($('acc-count-hotmail1')) $('acc-count-hotmail1').textContent = items.filter((d) => (d.Account || '').includes('Hotmail Primary')).length;
  if ($('acc-count-gmail')) $('acc-count-gmail').textContent = items.filter((d) => (d.Account || '').includes('Gmail Primary')).length;
  if ($('acc-count-hotmail2')) $('acc-count-hotmail2').textContent = items.filter((d) => (d.Account || '').includes('Hotmail Secondary')).length;
}

function getFilteredItems() {
  let list = (state.inbound || []).slice();

  // 1. Filter by Account
  if (state.selectedAccount && state.selectedAccount !== 'all') {
    list = list.filter((item) => (item.Account || '').toLowerCase().includes(state.selectedAccount.toLowerCase()));
  }

  // 2. Filter by Category Tab
  if (state.selectedTab === 'next-steps') {
    list = list.filter((d) => d['Is Next Step'] === true || d['Is Next Step'] === 'True' || d['Is Next Step'] === 'true');
  } else if (state.selectedTab === 'invites') {
    list = list.filter((d) => d.Category === 'Interview Invitation');
  } else if (state.selectedTab === 'assessments') {
    list = list.filter((d) => d.Category === 'Technical Assessment');
  } else if (state.selectedTab === 'inquiries') {
    list = list.filter((d) => (d.Category || '').includes('Availability') || (d.Category || '').includes('Inquiry'));
  } else if (state.selectedTab === 'rejections') {
    list = list.filter((d) => d.Category === 'Rejection');
  }

  // 3. Filter by Search Query
  if (state.searchQuery) {
    const q = state.searchQuery.toLowerCase();
    list = list.filter((item) => {
      const text = [
        item.Company, item['Job Title'], item.Sender, item.Subject,
        item['Email Snippet'], item['Full Email Body'], item.Language
      ].join(' ').toLowerCase();
      return text.includes(q);
    });
  }

  return list;
}

function renderCards() {
  const container = $('inbound-cards-container');
  if (!container) return;

  const items = getFilteredItems();
  if ($('visible-inbound-count')) $('visible-inbound-count').textContent = items.length;

  if (!items.length) {
    container.innerHTML = `
      <div class="inbound-empty-card" style="grid-column: 1 / -1; padding: 3rem; background: rgba(255,255,255,0.03); border: 1px dashed rgba(255,255,255,0.15); border-radius: 12px; text-align: center; color: #94a3b8;">
        <p style="margin: 0 0 10px 0; font-size: 16px; font-weight:600; color:#e2e8f0;">No emails match the selected filters.</p>
        <button class="filter-button" onclick="resetFilters()" style="padding: 8px 16px;">Reset Filter Tabs</button>
      </div>
    `;
    return;
  }

  container.innerHTML = items.map((item, idx) => {
    const cat = item.Category || 'General';
    let cardClass = 'inbound-card';
    let badgeClass = 'badge-general';
    let icon = '⚪';

    if (cat === 'Interview Invitation') {
      cardClass += ' is-interview';
      badgeClass = 'badge-interview';
      icon = '🟢';
    } else if (cat === 'Technical Assessment') {
      cardClass += ' is-assessment';
      badgeClass = 'badge-assessment';
      icon = '🔵';
    } else if (cat.includes('Availability')) {
      cardClass += ' is-inquiry';
      badgeClass = 'badge-inquiry';
      icon = '🟡';
    } else if (cat === 'Rejection') {
      cardClass += ' is-rejection';
      badgeClass = 'badge-rejection';
      icon = '🔴';
    }

    const actionUrl = item['Action URL'] || '';
    const hasBookingLink = actionUrl && /^https?:\/\//i.test(actionUrl);
    const sender = item.Sender || 'Recruiter';
    const cleanSenderEmail = (sender.match(/<([^>]+)>/) || [null, sender])[1].trim();
    const replySubject = encodeURIComponent(`Re: ${item.Subject || 'Interview Next Steps'}`);
    const replyBody = encodeURIComponent(`Hi ${item.Company} Team,\n\nThank you for reaching out regarding the ${item['Job Title']} role!\n\nI would be delighted to proceed with the next steps. Please let me know your available times or feel free to send over any additional details.\n\nBest regards,\nHemachandiran Giri`);
    const mailtoLink = `mailto:${cleanSenderEmail}?subject=${replySubject}&body=${replyBody}`;

    const accountName = item.Account || 'Inbox';
    const lang = item.Language || 'English';

    return `
      <article class="${cardClass}">
        <div class="inbound-head">
          <div>
            <div style="display:flex; align-items:center; gap:6px; margin-bottom:4px;">
              <span class="source-tag source-indeed" style="margin:0; font-size:10px; padding:1px 6px;">📥 ${esc(accountName)}</span>
              ${lang !== 'English' ? `<span class="source-tag source-glassdoor" style="margin:0; font-size:10px; padding:1px 6px;">🌐 ${esc(lang)}</span>` : ''}
            </div>
            <h3 class="inbound-company">${esc(item.Company || 'Company')}</h3>
            <p class="inbound-role">${esc(item['Job Title'] || 'Role')}</p>
          </div>
          <span class="inbound-badge ${badgeClass}">${icon} ${esc(cat)}</span>
        </div>

        <div class="inbound-action-box">
          <span>⚡</span>
          <div><b>Action:</b> ${esc(item['Action Required'] || 'Review recruiter response')}</div>
        </div>

        <p class="inbound-snippet">${esc(item['Email Snippet'] || item.Subject || 'No snippet available.')}</p>

        <div class="inbound-actions">
          <button class="inbound-btn inbound-btn-view" onclick="openEmailModal(${idx})">📖 View Full Email</button>
          ${hasBookingLink ? `<a href="${esc(actionUrl)}" target="_blank" rel="noopener noreferrer" class="inbound-btn inbound-btn-primary">📅 Book Call / Test →</a>` : ''}
          <a href="${mailtoLink}" class="inbound-btn inbound-btn-reply">✉️ Reply Email</a>
          <button class="inbound-btn inbound-btn-done ${item['Replied Status'] === 'Yes' ? 'is-done' : ''}" onclick="toggleInboundDone(${idx})">
            ${item['Replied Status'] === 'Yes' ? '✓ Replied' : '○ Mark Replied'}
          </button>
        </div>

        <div class="inbound-meta">
          <span>From: ${esc(sender.split('<')[0].trim() || sender)}</span>
          <span>📅 ${esc(item['Date Received'] || '')}</span>
        </div>
      </article>
    `;
  }).join('');
}

function renderAll() {
  renderMetrics();
  renderCards();
}

window.resetFilters = () => {
  state.selectedAccount = 'all';
  state.selectedTab = 'next-steps';
  state.searchQuery = '';
  if ($('search-inbound')) $('search-inbound').value = '';
  document.querySelectorAll('#account-filter-bar button').forEach((b) => b.classList.toggle('active', b.dataset.account === 'all'));
  document.querySelectorAll('#inbound-category-tabs button').forEach((b) => b.classList.toggle('active', b.dataset.tab === 'next-steps'));
  renderAll();
};

window.toggleInboundDone = (idx) => {
  const filtered = getFilteredItems();
  const item = filtered[idx];
  if (!item) return;

  const cur = item['Replied Status'];
  item['Replied Status'] = cur === 'Yes' ? 'No' : 'Yes';
  renderAll();
  showToast(`Updated status for ${item.Company}.`);
};

// ── Full Email Modal ──
window.openEmailModal = (idx) => {
  const filtered = getFilteredItems();
  const item = filtered[idx];
  if (!item) return;

  const modal = $('inbound-email-modal');
  if (!modal) return;

  const cat = item.Category || 'General';
  let icon = '📬';
  if (cat === 'Interview Invitation') icon = '🟢';
  else if (cat === 'Technical Assessment') icon = '🔵';
  else if (cat.includes('Availability')) icon = '🟡';
  else if (cat === 'Rejection') icon = '🔴';

  $('modal-email-icon').textContent = icon;
  $('modal-email-title').textContent = item.Subject || `${item.Company} - Next Steps`;

  const sender = item.Sender || 'Recruiter';
  const cleanSenderEmail = (sender.match(/<([^>]+)>/) || [null, sender])[1].trim();
  const accountName = item.Account || 'Email Inbox';
  const inboxEmail = item['Inbox Email'] || '';
  const dateReceived = item['Date Received'] || '';
  const lang = item.Language || 'English';

  $('modal-email-meta').innerHTML = `
    <div class="email-modal-meta-item">
      <span class="email-modal-meta-label">From</span>
      <span class="email-modal-meta-val">${esc(sender)}</span>
    </div>
    <div class="email-modal-meta-item">
      <span class="email-modal-meta-label">To / Inbox</span>
      <span class="email-modal-meta-val">${esc(accountName)} ${inboxEmail ? `(${esc(inboxEmail)})` : ''}</span>
    </div>
    <div class="email-modal-meta-item">
      <span class="email-modal-meta-label">Company & Role</span>
      <span class="email-modal-meta-val"><b>${esc(item.Company || 'Company')}</b> — ${esc(item['Job Title'] || 'Role')}</span>
    </div>
    <div class="email-modal-meta-item">
      <span class="email-modal-meta-label">Date & Language</span>
      <span class="email-modal-meta-val">📅 ${esc(dateReceived)} · 🌐 ${esc(lang)}</span>
    </div>
  `;

  $('modal-email-action-box').innerHTML = `
    <span>⚡</span>
    <div><b>Action Required:</b> ${esc(item['Action Required'] || 'Review recruiter email')}</div>
  `;

  const fullBody = item['Full Email Body'] || item['Email Snippet'] || item.Subject || 'No email content text available.';
  $('modal-email-full-text').textContent = fullBody;

  const linksSec = $('modal-email-links-section');
  const linksList = $('modal-email-links-list');
  const allLinksStr = item['All Links'] || item['Action URL'] || '';
  const links = allLinksStr.split('|').map((s) => s.trim()).filter((s) => /^https?:\/\//i.test(s));

  if (linksSec && linksList) {
    if (links.length > 0) {
      linksSec.style.display = 'block';
      linksList.innerHTML = links.map((l) => `<a href="${esc(l)}" target="_blank" rel="noopener noreferrer" class="email-link-chip">🔗 ${esc(l)}</a>`).join('');
    } else {
      linksSec.style.display = 'none';
    }
  }

  const actionUrl = item['Action URL'] || '';
  const hasBookingLink = actionUrl && /^https?:\/\//i.test(actionUrl);
  const replySubject = encodeURIComponent(`Re: ${item.Subject || 'Interview Next Steps'}`);
  const replyBody = encodeURIComponent(`Hi ${item.Company} Team,\n\nThank you for reaching out regarding the ${item['Job Title']} role!\n\nI would be delighted to proceed with the next steps. Please let me know your available times or feel free to send over any additional details.\n\nBest regards,\nHemachandiran Giri`);
  const mailtoLink = `mailto:${cleanSenderEmail}?subject=${replySubject}&body=${replyBody}`;

  $('modal-email-action-btns').innerHTML = `
    ${hasBookingLink ? `<a href="${esc(actionUrl)}" target="_blank" rel="noopener noreferrer" class="inbound-btn inbound-btn-primary" style="padding: 8px 16px;">📅 Book Call / Test →</a>` : ''}
    <a href="${mailtoLink}" class="inbound-btn inbound-btn-reply" style="padding: 8px 16px;">✉️ Reply via Email</a>
  `;

  modal.hidden = false;
};

window.closeEmailModal = () => {
  const modal = $('inbound-email-modal');
  if (modal) modal.hidden = true;
};

// ── Listeners ──
function initListeners() {
  // Category tabs
  const catBar = $('inbound-category-tabs');
  if (catBar) {
    catBar.querySelectorAll('[data-tab]').forEach((btn) => {
      btn.onclick = () => {
        state.selectedTab = btn.dataset.tab;
        catBar.querySelectorAll('[data-tab]').forEach((b) => b.classList.remove('active'));
        btn.classList.add('active');
        renderCards();
      };
    });
  }

  // Account filter bar
  const accBar = $('account-filter-bar');
  if (accBar) {
    accBar.querySelectorAll('[data-account]').forEach((btn) => {
      btn.onclick = () => {
        state.selectedAccount = btn.dataset.account;
        accBar.querySelectorAll('[data-account]').forEach((b) => b.classList.remove('active'));
        btn.classList.add('active');
        renderCards();
      };
    });
  }

  // Metric card clicks
  const actionCard = $('metric-action-card');
  if (actionCard) {
    actionCard.onclick = () => {
      state.selectedTab = 'next-steps';
      document.querySelectorAll('#inbound-category-tabs button').forEach((b) => b.classList.toggle('active', b.dataset.tab === 'next-steps'));
      renderCards();
    };
  }
  const invitesCard = $('metric-invites-card');
  if (invitesCard) {
    invitesCard.onclick = () => {
      state.selectedTab = 'invites';
      document.querySelectorAll('#inbound-category-tabs button').forEach((b) => b.classList.toggle('active', b.dataset.tab === 'invites'));
      renderCards();
    };
  }
  const assessCard = $('metric-assessments-card');
  if (assessCard) {
    assessCard.onclick = () => {
      state.selectedTab = 'assessments';
      document.querySelectorAll('#inbound-category-tabs button').forEach((b) => b.classList.toggle('active', b.dataset.tab === 'assessments'));
      renderCards();
    };
  }
  const inqCard = $('metric-inquiries-card');
  if (inqCard) {
    inqCard.onclick = () => {
      state.selectedTab = 'inquiries';
      document.querySelectorAll('#inbound-category-tabs button').forEach((b) => b.classList.toggle('active', b.dataset.tab === 'inquiries'));
      renderCards();
    };
  }

  // Search input
  const searchInput = $('search-inbound');
  if (searchInput) {
    searchInput.oninput = (e) => {
      state.searchQuery = e.target.value.trim();
      renderCards();
    };
  }

  // Refresh pipeline button
  const refreshBtn = $('refresh-pipeline-btn');
  if (refreshBtn) {
    refreshBtn.onclick = () => {
      loadPipelineData();
      showToast('Refreshed Inbound Pipeline.');
    };
  }

  // Modal close handlers
  const closeX = $('close-email-modal');
  if (closeX) closeX.onclick = closeEmailModal;
  const closeBtn = $('close-email-modal-btn');
  if (closeBtn) closeBtn.onclick = closeEmailModal;
  const emailModal = $('inbound-email-modal');
  if (emailModal) {
    emailModal.onclick = (e) => {
      if (e.target === emailModal) closeEmailModal();
    };
  }
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeEmailModal();
  });
}

// ── Init ──
(function init() {
  initAuth();
  initListeners();
  if ($('auth-overlay') && !isAuthActive()) {
    showAuthModal();
  } else {
    unlockDashboard();
  }
})();
