/**
 * Google Apps Script for Job Compass / JobCrawler
 * ======================================================================
 * Paste this script into your Google Spreadsheet (Extensions -> Apps Script),
 * then click "Deploy" -> "New deployment" -> Select type: "Web app"
 * - Execute as: "Me"
 * - Who has access: "Anyone"
 *
 * Copy the resulting Web Application URL and enter it in your Job Compass Dashboard
 * or save it in `google_sheets_config.json`.
 * ======================================================================
 */

const SHEET_NAME = 'All Jobs';
const SPREADSHEET_ID = '1kUPHPL8hPRKG2d_5D5j6Qb6r72MJxpohh6xL1TEYYWI';

function setupSheet() {
  let ss = null;
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim()) {
    try { ss = SpreadsheetApp.openById(SPREADSHEET_ID.trim()); } catch (e) {}
  }
  if (!ss) {
    try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  }
  if (!ss) {
    const files = DriveApp.getFilesByType(MimeType.GOOGLE_SHEETS);
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      ss = SpreadsheetApp.create("Job Compass Opportunities");
    }
  }
  
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  
  const headers = [
    "#", "Crawl Date", "Country", "Flag", "Tier", "City", "City Openings",
    "Location", "Company", "Job Title", "Search Keyword", "Posted Date",
    "Easy Apply", "Remote / Workplace", "Match Score", "Visa Sponsorship",
    "Skills", "Resume", "Applied Status", "Notes", "Job URL"
  ];
  
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1F4E79');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return { ss: ss, sheet: sheet, headers: headers };
}

function syncDateTabs(ss, headers, values) {
  try {
    const dateGroups = new Map();
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const rawDate = String(row[1] || '').trim();
      const dateStr = rawDate.split('T')[0] || 'Unknown Date';
      if (!dateGroups.has(dateStr)) dateGroups.set(dateStr, []);
      dateGroups.get(dateStr).push(row);
    }
    
    dateGroups.forEach((rows, dateStr) => {
      if (!dateStr || dateStr.length < 5) return;
      let dateSheet = ss.getSheetByName(dateStr);
      if (!dateSheet) {
        dateSheet = ss.insertSheet(dateStr);
      } else {
        dateSheet.clearContents();
      }
      
      const sheetData = [headers, ...rows];
      dateSheet.getRange(1, 1, sheetData.length, headers.length).setValues(sheetData);
      
      const headerRange = dateSheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#1F4E79');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setFontWeight('bold');
      dateSheet.setFrozenRows(1);
    });
  } catch (e) {
    Logger.log("Error syncing date tabs: " + e.toString());
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const setup = setupSheet();
    const sheet = setup.sheet;
    const ss = setup.ss;
    const headers = setup.headers;
    
    // Action 1: Update single job applied status (from index.html UI)
    if (data.action === 'update_applied') {
      const targetUrl = (data.jobUrl || '').trim();
      const newStatus = data.appliedStatus || 'Yes';
      
      if (!targetUrl) {
        return responseJSON({ success: false, error: 'Missing jobUrl' });
      }
      
      const values = sheet.getDataRange().getValues();
      const urlColIdx = 20; // 21st column (0-indexed 20) is Job URL
      const appliedColIdx = 18; // 19th column (0-indexed 18) is Applied Status
      
      let updated = false;
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][urlColIdx]).trim() === targetUrl) {
          sheet.getRange(i + 1, appliedColIdx + 1).setValue(newStatus);
          updated = true;
          break;
        }
      }
      
      return responseJSON({ success: true, updated: updated, status: newStatus });
    }
    
    // Action 2: Batch import/sync (from crawler or dashboard full sync)
    if (data.action === 'sync_jobs' || Array.isArray(data.jobs)) {
      const jobs = data.jobs || [];
      if (!jobs.length) {
        return responseJSON({ success: true, count: 0 });
      }
      
      const existingValues = sheet.getDataRange().getValues();
      const urlMap = new Map();
      const comboMap = new Map();
      
      for (let i = 1; i < existingValues.length; i++) {
        const row = existingValues[i];
        const url = String(row[20] || '').trim();
        const company = String(row[8] || '').trim().toLowerCase();
        const title = String(row[9] || '').trim().toLowerCase();
        if (url) urlMap.set(url, i);
        if (company && title) comboMap.set(`${company}|${title}`, i);
      }
      
      let added = 0;
      let updated = 0;
      const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
      const rowsToAppend = [];

      jobs.forEach((job) => {
        const url = (job['Job URL'] || '').trim();
        const company = (job.Company || '').trim();
        const title = (job['Job Title'] || '').trim();
        const comboKey = `${company.toLowerCase()}|${title.toLowerCase()}`;
        
        const existingRowIdx = (url && urlMap.get(url)) ?? comboMap.get(comboKey);
        
        const rawCrawlDate = job['Crawl Date'] || job.Date || today;
        const crawlDate = String(rawCrawlDate).split('T')[0] || today;
        let appliedStatus = job['Applied Status'] || job.Applied || 'No';
        
        if (existingRowIdx !== undefined && existingRowIdx < existingValues.length) {
          const currentApplied = String(existingValues[existingRowIdx][18] || '').trim();
          if (currentApplied.toLowerCase() === 'yes' && appliedStatus.toLowerCase() === 'no') {
            appliedStatus = 'Yes';
          }
        }

        const rowData = [
          existingRowIdx !== undefined ? existingRowIdx : existingValues.length + rowsToAppend.length,
          crawlDate,
          job.Country || '',
          job.Flag || '',
          job.Tier || '',
          job.City || '',
          job['City Openings'] || '',
          job.Location || '',
          company,
          title,
          job['Search Keyword'] || '',
          job['Posted Date'] || '',
          job['Easy Apply'] || 'No',
          job['Remote / Workplace'] || 'On-site / Hybrid',
          job['Match Score'] || '85%',
          job['Visa Sponsorship Mentioned'] || job['Visa Sponsorship'] || 'No',
          job['Required Skills'] || job.Skills || '',
          job['Resume File Path'] || job.Resume || '',
          appliedStatus,
          job.Notes || '',
          url
        ];
        
        if (existingRowIdx !== undefined && existingRowIdx < existingValues.length) {
          existingValues[existingRowIdx] = rowData;
          updated++;
        } else {
          rowsToAppend.push(rowData);
          added++;
        }
      });
      
      // Bulk update existing rows in memory
      if (updated > 0 && existingValues.length > 1) {
        sheet.getRange(1, 1, existingValues.length, existingValues[0].length).setValues(existingValues);
      }
      
      // Bulk append new rows
      if (rowsToAppend.length > 0) {
        const startRow = sheet.getLastRow() + 1;
        sheet.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
      }
      
      // Maintain Date Tabs in spreadsheet automatically
      const finalValues = sheet.getDataRange().getValues();
      syncDateTabs(ss, headers, finalValues);
      
      return responseJSON({ success: true, added: added, updated: updated, total: sheet.getLastRow() - 1 });
    }
    
    // Action 3: Fetch all jobs from Google Sheet
    if (data.action === 'get_jobs') {
      const existingValues = sheet.getDataRange().getValues();
      const headers = existingValues[0] || [];
      const jobs = [];
      for (let i = 1; i < existingValues.length; i++) {
        const row = existingValues[i];
        const job = {};
        headers.forEach((h, idx) => { job[h] = row[idx]; });
        jobs.push(job);
      }
      return responseJSON({ success: true, jobs: jobs });
    }

    return responseJSON({ success: false, error: 'Unknown action' });
  } catch (err) {
    return responseJSON({ success: false, error: err.toString() });
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'get_jobs') {
    try {
      const setup = setupSheet();
      const sheet = setup.sheet;
      const existingValues = sheet.getDataRange().getValues();
      const headers = existingValues[0] || [];
      const jobs = [];
      for (let i = 1; i < existingValues.length; i++) {
        const row = existingValues[i];
        const job = {};
        headers.forEach((h, idx) => { job[h] = row[idx]; });
        jobs.push(job);
      }
      
      const callback = e.parameter.callback;
      if (callback) {
        return ContentService.createTextOutput(callback + '(' + JSON.stringify({ success: true, jobs: jobs }) + ')')
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }

      return responseJSON({ success: true, jobs: jobs });
    } catch (err) {
      const callback = e && e.parameter && e.parameter.callback;
      if (callback) {
        return ContentService.createTextOutput(callback + '(' + JSON.stringify({ success: false, error: err.toString() }) + ')')
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return responseJSON({ success: false, error: err.toString() });
    }
  }

  return responseJSON({
    name: "Job Compass Google Apps Script Webhook API",
    status: "active",
    spreadsheet_url: "https://docs.google.com/spreadsheets/d/" + SPREADSHEET_ID + "/edit"
  });
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
