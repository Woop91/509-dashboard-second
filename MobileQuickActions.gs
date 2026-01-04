/**
 * ============================================================================
 * MOBILE INTERFACE & QUICK ACTIONS
 * ============================================================================
 * Mobile-optimized views and context-aware quick actions
 * Includes automatic device detection for responsive experience
 */

// ==================== MOBILE CONFIGURATION ====================

var MOBILE_CONFIG = {
  MAX_COLUMNS_MOBILE: 8,
  CARD_LAYOUT_ENABLED: true,
  TOUCH_TARGET_SIZE: '44px',
  MOBILE_BREAKPOINT: 768,  // Width in pixels below which is considered mobile
  TABLET_BREAKPOINT: 1024  // Width in pixels below which is considered tablet
};

// ==================== DEVICE DETECTION ====================

/**
 * Shows a smart dashboard that automatically detects the device type
 * and displays the appropriate interface (mobile or desktop)
 */
function showSmartDashboard() {
  var html = HtmlService.createHtmlOutput(getSmartDashboardHtml())
    .setWidth(800)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 509 Dashboard');
}

/**
 * Returns the HTML for the smart dashboard with device detection
 */
function getSmartDashboardHtml() {
  var stats = getMobileDashboardStats();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<base target="_top">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<style>' +
    // CSS Reset and base styles
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh}' +

    // Responsive container
    '.container{padding:15px;max-width:1200px;margin:0 auto}' +

    // Header - responsive
    '.header{background:linear-gradient(135deg,#1a73e8,#1557b0);color:white;padding:20px;text-align:center}' +
    '.header h1{font-size:clamp(18px,5vw,28px);margin-bottom:5px}' +
    '.header .subtitle{font-size:clamp(12px,3vw,14px);opacity:0.9}' +
    '.device-badge{display:inline-block;padding:4px 12px;background:rgba(255,255,255,0.2);border-radius:20px;font-size:11px;margin-top:8px}' +

    // Stats grid - responsive
    '.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:20px}' +
    '.stat-card{background:white;padding:20px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);text-align:center;transition:transform 0.2s}' +
    '.stat-card:hover{transform:translateY(-2px)}' +
    '.stat-value{font-size:clamp(24px,6vw,36px);font-weight:bold;color:#1a73e8}' +
    '.stat-label{font-size:clamp(11px,2.5vw,13px);color:#666;text-transform:uppercase;margin-top:5px}' +

    // Section titles
    '.section-title{font-size:clamp(14px,3.5vw,18px);font-weight:600;color:#333;margin:20px 0 12px;padding-left:5px}' +

    // Action buttons - responsive grid
    '.actions{display:grid;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));gap:10px}' +
    '.action-btn{background:white;border:none;padding:16px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);' +
    'width:100%;text-align:left;display:flex;align-items:center;gap:15px;font-size:15px;cursor:pointer;' +
    'min-height:' + MOBILE_CONFIG.TOUCH_TARGET_SIZE + ';transition:all 0.2s}' +
    '.action-btn:hover{background:#e8f0fe;transform:translateX(4px)}' +
    '.action-btn:active{transform:scale(0.98)}' +
    '.action-icon{font-size:24px;width:44px;height:44px;display:flex;align-items:center;justify-content:center;background:#e8f0fe;border-radius:10px;flex-shrink:0}' +
    '.action-label{font-weight:500}' +
    '.action-desc{font-size:12px;color:#666;margin-top:2px}' +

    // FAB (Floating Action Button)
    '.fab{position:fixed;bottom:20px;right:20px;width:56px;height:56px;background:#1a73e8;color:white;' +
    'border:none;border-radius:50%;font-size:24px;box-shadow:0 4px 12px rgba(0,0,0,0.3);cursor:pointer;z-index:1000}' +
    '.fab:hover{background:#1557b0}' +

    // Desktop-only elements
    '.desktop-only{display:none}' +

    // Mobile-specific adjustments
    '@media (max-width:' + MOBILE_CONFIG.MOBILE_BREAKPOINT + 'px){' +
    '  .stats{grid-template-columns:repeat(2,1fr)}' +
    '  .actions{grid-template-columns:1fr}' +
    '  .container{padding:10px}' +
    '  .header{padding:15px}' +
    '}' +

    // Tablet adjustments
    '@media (min-width:' + MOBILE_CONFIG.MOBILE_BREAKPOINT + 'px) and (max-width:' + MOBILE_CONFIG.TABLET_BREAKPOINT + 'px){' +
    '  .stats{grid-template-columns:repeat(2,1fr)}' +
    '  .actions{grid-template-columns:repeat(2,1fr)}' +
    '}' +

    // Desktop view
    '@media (min-width:' + MOBILE_CONFIG.TABLET_BREAKPOINT + 'px){' +
    '  .stats{grid-template-columns:repeat(4,1fr)}' +
    '  .actions{grid-template-columns:repeat(2,1fr)}' +
    '  .desktop-only{display:block}' +
    '}' +

    '</style>' +
    '</head><body>' +

    // Header with dynamic device badge
    '<div class="header">' +
    '<h1>📱 509 Dashboard</h1>' +
    '<div class="subtitle">Union Grievance Management</div>' +
    '<div class="device-badge" id="deviceBadge">Detecting device...</div>' +
    '</div>' +

    '<div class="container">' +

    // Stats section
    '<div class="stats">' +
    '<div class="stat-card"><div class="stat-value">' + stats.totalGrievances + '</div><div class="stat-label">Total</div></div>' +
    '<div class="stat-card"><div class="stat-value">' + stats.activeGrievances + '</div><div class="stat-label">Active</div></div>' +
    '<div class="stat-card"><div class="stat-value">' + stats.pendingGrievances + '</div><div class="stat-label">Pending</div></div>' +
    '<div class="stat-card"><div class="stat-value">' + stats.overdueGrievances + '</div><div class="stat-label">Overdue</div></div>' +
    '</div>' +

    // Quick Actions
    '<div class="section-title">⚡ Quick Actions</div>' +
    '<div class="actions">' +

    '<button class="action-btn" onclick="google.script.run.showMobileGrievanceList()">' +
    '<div class="action-icon">📋</div>' +
    '<div><div class="action-label">View Grievances</div><div class="action-desc">Browse and filter all grievances</div></div>' +
    '</button>' +

    '<button class="action-btn" onclick="google.script.run.showMobileUnifiedSearch()">' +
    '<div class="action-icon">🔍</div>' +
    '<div><div class="action-label">Search</div><div class="action-desc">Find grievances or members</div></div>' +
    '</button>' +

    '<button class="action-btn" onclick="google.script.run.showMyAssignedGrievances()">' +
    '<div class="action-icon">👤</div>' +
    '<div><div class="action-label">My Cases</div><div class="action-desc">View your assigned grievances</div></div>' +
    '</button>' +

    '<button class="action-btn" onclick="google.script.run.showQuickActionsMenu()">' +
    '<div class="action-icon">⚡</div>' +
    '<div><div class="action-label">Row Actions</div><div class="action-desc">Quick actions for selected row</div></div>' +
    '</button>' +

    '</div>' +

    // Desktop-only additional info
    '<div class="desktop-only">' +
    '<div class="section-title">ℹ️ Dashboard Info</div>' +
    '<p style="color:#666;font-size:14px;padding:15px;background:white;border-radius:8px;">' +
    'This responsive dashboard automatically adjusts to your screen size. ' +
    'On mobile devices, you\'ll see a touch-optimized interface with larger buttons. ' +
    'Use the menu items above to manage grievances and member information.' +
    '</p>' +
    '</div>' +

    '</div>' +

    // FAB for refresh
    '<button class="fab" onclick="location.reload()" title="Refresh">🔄</button>' +

    // Device detection script
    '<script>' +
    'function detectDevice(){' +
    '  var w=window.innerWidth;' +
    '  var badge=document.getElementById("deviceBadge");' +
    '  var isTouchDevice="ontouchstart" in window||navigator.maxTouchPoints>0;' +
    '  var userAgent=navigator.userAgent.toLowerCase();' +
    '  var isMobileUA=/android|webos|iphone|ipad|ipod|blackberry|iemobile|opera mini/i.test(userAgent);' +
    '  ' +
    '  if(w<' + MOBILE_CONFIG.MOBILE_BREAKPOINT + '||isMobileUA){' +
    '    badge.textContent="📱 Mobile View";' +
    '    badge.style.background="rgba(76,175,80,0.3)";' +
    '  }else if(w<' + MOBILE_CONFIG.TABLET_BREAKPOINT + '){' +
    '    badge.textContent="📱 Tablet View";' +
    '    badge.style.background="rgba(255,152,0,0.3)";' +
    '  }else{' +
    '    badge.textContent="🖥️ Desktop View";' +
    '    badge.style.background="rgba(33,150,243,0.3)";' +
    '  }' +
    '  ' +
    '  if(isTouchDevice){' +
    '    document.body.classList.add("touch-device");' +
    '  }' +
    '}' +
    'detectDevice();' +
    'window.addEventListener("resize",detectDevice);' +
    '</script>' +

    '</body></html>';
}

/**
 * Check if the current context appears to be mobile
 * Note: This is a server-side heuristic based on available info
 * Real detection happens client-side in the HTML
 */
function isMobileContext() {
  // Server-side we can't reliably detect mobile
  // This function exists for potential future use with session properties
  return false;
}

// ==================== MOBILE DASHBOARD ====================

function showMobileDashboard() {
  var stats = getMobileDashboardStats();
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no"><style>*{box-sizing:border-box}body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial;padding:0;margin:0;background:#f5f5f5}.header{background:linear-gradient(135deg,#1a73e8,#1557b0);color:white;padding:20px}.header h1{margin:0;font-size:24px}.container{padding:15px}.stats{display:grid;grid-template-columns:repeat(2,1fr);gap:12px;margin-bottom:20px}.stat-card{background:white;padding:20px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);text-align:center}.stat-value{font-size:32px;font-weight:bold;color:#1a73e8}.stat-label{font-size:13px;color:#666;text-transform:uppercase}.section-title{font-size:16px;font-weight:600;color:#333;margin:20px 0 12px;padding-left:5px}.action-btn{background:white;border:none;padding:16px;margin-bottom:10px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);width:100%;text-align:left;display:flex;align-items:center;gap:15px;font-size:15px;cursor:pointer;min-height:' + MOBILE_CONFIG.TOUCH_TARGET_SIZE + '}.action-btn:active{transform:scale(0.98)}.action-icon{font-size:24px;width:40px;height:40px;display:flex;align-items:center;justify-content:center;background:#e8f0fe;border-radius:10px}.action-label{font-weight:500}.action-desc{font-size:12px;color:#666}.fab{position:fixed;bottom:20px;right:20px;width:56px;height:56px;background:#1a73e8;color:white;border:none;border-radius:50%;font-size:24px;box-shadow:0 4px 12px rgba(0,0,0,0.3);cursor:pointer}</style></head><body><div class="header"><h1>📱 509 Dashboard</h1><div style="font-size:14px;opacity:0.9">Mobile View</div></div><div class="container"><div class="stats"><div class="stat-card"><div class="stat-value">' + stats.totalGrievances + '</div><div class="stat-label">Total</div></div><div class="stat-card"><div class="stat-value">' + stats.activeGrievances + '</div><div class="stat-label">Active</div></div><div class="stat-card"><div class="stat-value">' + stats.pendingGrievances + '</div><div class="stat-label">Pending</div></div><div class="stat-card"><div class="stat-value">' + stats.overdueGrievances + '</div><div class="stat-label">Overdue</div></div></div><div class="section-title">⚡ Quick Actions</div><button class="action-btn" onclick="google.script.run.showMobileGrievanceList()"><div class="action-icon">📋</div><div><div class="action-label">View Grievances</div><div class="action-desc">Browse all grievances</div></div></button><button class="action-btn" onclick="google.script.run.showMobileUnifiedSearch()"><div class="action-icon">🔍</div><div><div class="action-label">Search</div><div class="action-desc">Find grievances or members</div></div></button><button class="action-btn" onclick="google.script.run.showMyAssignedGrievances()"><div class="action-icon">👤</div><div><div class="action-label">My Cases</div><div class="action-desc">View assigned grievances</div></div></button></div><button class="fab" onclick="location.reload()">🔄</button></body></html>'
  ).setWidth(400).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '📱 Mobile Dashboard');
}

function getMobileDashboardStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet || sheet.getLastRow() <= 1) return { totalGrievances: 0, activeGrievances: 0, pendingGrievances: 0, overdueGrievances: 0 };
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, GRIEVANCE_COLS.DAYS_TO_DEADLINE).getValues();
  var stats = { totalGrievances: data.length, activeGrievances: 0, pendingGrievances: 0, overdueGrievances: 0 };
  var today = new Date(); today.setHours(0, 0, 0, 0);
  data.forEach(function(row) {
    var status = row[GRIEVANCE_COLS.STATUS - 1];
    var daysTo = row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
    if (status && status !== 'Resolved' && status !== 'Withdrawn') stats.activeGrievances++;
    if (status === 'Pending Info') stats.pendingGrievances++;
    if (daysTo !== null && daysTo !== '' && daysTo < 0 && status === 'Open') stats.overdueGrievances++;
  });
  return stats;
}

function getRecentGrievancesForMobile(limit) {
  limit = limit || 5;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, GRIEVANCE_COLS.RESOLUTION).getValues();
  return data.map(function(row, idx) {
    var filed = row[GRIEVANCE_COLS.DATE_FILED - 1];
    var deadline = row[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];
    return {
      id: row[GRIEVANCE_COLS.GRIEVANCE_ID - 1],
      memberName: row[GRIEVANCE_COLS.FIRST_NAME - 1] + ' ' + row[GRIEVANCE_COLS.LAST_NAME - 1],
      issueType: row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1],
      status: row[GRIEVANCE_COLS.STATUS - 1],
      filedDate: filed instanceof Date ? Utilities.formatDate(filed, Session.getScriptTimeZone(), 'MM/dd/yyyy') : filed,
      deadline: deadline instanceof Date ? Utilities.formatDate(deadline, Session.getScriptTimeZone(), 'MM/dd/yyyy') : null,
      filedDateObj: filed
    };
  }).sort(function(a, b) {
    var da = a.filedDateObj instanceof Date ? a.filedDateObj : new Date(0);
    var db = b.filedDateObj instanceof Date ? b.filedDateObj : new Date(0);
    return db - da;
  }).slice(0, limit);
}

function showMobileGrievanceList() {
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<style>' +
    '*{box-sizing:border-box}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial;margin:0;padding:0;background:#f5f5f5}' +
    '.header{background:#1a73e8;color:white;padding:15px;position:sticky;top:0;z-index:100}' +
    '.header h2{margin:0;font-size:clamp(18px,4vw,24px)}' +
    '.search{width:100%;padding:clamp(10px,2.5vw,14px);border:none;border-radius:8px;font-size:clamp(14px,3vw,16px);margin-top:10px}' +
    '.filters{display:flex;overflow-x:auto;padding:10px;background:white;gap:8px;-webkit-overflow-scrolling:touch}' +
    '.filter{padding:clamp(6px,1.5vw,10px) clamp(12px,3vw,18px);border-radius:20px;background:#f0f0f0;white-space:nowrap;cursor:pointer;font-size:clamp(12px,2.5vw,14px);border:none;min-height:' + MOBILE_CONFIG.TOUCH_TARGET_SIZE + ';display:flex;align-items:center}' +
    '.filter.active{background:#1a73e8;color:white}' +
    '.list{padding:10px;display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:12px}' +
    '.card{background:white;padding:15px;border-radius:12px;box-shadow:0 2px 4px rgba(0,0,0,0.08)}' +
    '.card-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;flex-wrap:wrap;gap:8px}' +
    '.card-id{font-weight:bold;color:#1a73e8;font-size:clamp(14px,3vw,16px)}' +
    '.card-status{padding:4px 10px;border-radius:12px;font-size:clamp(10px,2vw,12px);font-weight:bold;background:#e8f0fe}' +
    '.card-row{font-size:clamp(12px,2.5vw,14px);margin:5px 0;color:#666}' +
    '@media (min-width:768px){.list{grid-template-columns:repeat(2,1fr)}}' +
    '@media (min-width:1024px){.list{grid-template-columns:repeat(3,1fr)}}' +
    '</style></head><body>' +
    '<div class="header"><h2>📋 Grievances</h2><input type="text" class="search" placeholder="Search..." oninput="filter(this.value)"></div>' +
    '<div class="filters"><button class="filter active" onclick="filterStatus(\'all\',this)">All</button><button class="filter" onclick="filterStatus(\'Open\',this)">Open</button><button class="filter" onclick="filterStatus(\'Pending Info\',this)">Pending</button><button class="filter" onclick="filterStatus(\'Resolved\',this)">Resolved</button></div>' +
    '<div class="list" id="list"><div style="text-align:center;padding:40px;color:#666;grid-column:1/-1">Loading...</div></div>' +
    '<script>var all=[];google.script.run.withSuccessHandler(function(data){all=data;render(data)}).getRecentGrievancesForMobile(100);function render(data){var c=document.getElementById("list");if(!data||data.length===0){c.innerHTML="<div style=\'text-align:center;padding:40px;color:#999;grid-column:1/-1\'>No grievances</div>";return}c.innerHTML=data.map(function(g){return"<div class=\'card\'><div class=\'card-header\'><div class=\'card-id\'>#"+g.id+"</div><div class=\'card-status\'>"+(g.status||"Filed")+"</div></div><div class=\'card-row\'><strong>Member:</strong> "+g.memberName+"</div><div class=\'card-row\'><strong>Issue:</strong> "+(g.issueType||"N/A")+"</div><div class=\'card-row\'><strong>Filed:</strong> "+g.filedDate+"</div></div>"}).join("")}function filterStatus(s,btn){document.querySelectorAll(".filter").forEach(function(f){f.classList.remove("active")});btn.classList.add("active");render(s==="all"?all:all.filter(function(g){return g.status===s}))}function filter(q){render(all.filter(function(g){q=q.toLowerCase();return g.id.toLowerCase().indexOf(q)>=0||g.memberName.toLowerCase().indexOf(q)>=0||(g.issueType||"").toLowerCase().indexOf(q)>=0}))}</script></body></html>'
  ).setWidth(800).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Grievance List');
}

function showMobileUnifiedSearch() {
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<style>' +
    '*{box-sizing:border-box}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial;margin:0;padding:0;background:#f5f5f5}' +
    '.header{background:linear-gradient(135deg,#1a73e8,#1557b0);color:white;padding:15px}' +
    '.header h2{margin:0 0 12px 0;font-size:clamp(18px,4vw,22px)}' +
    '.search-container{position:relative}' +
    '.search-input{width:100%;padding:clamp(12px,3vw,16px) clamp(12px,3vw,16px) clamp(12px,3vw,16px) 45px;border:none;border-radius:10px;font-size:clamp(14px,3vw,16px);background:white}' +
    '.search-icon{position:absolute;left:14px;top:50%;transform:translateY(-50%);font-size:18px}' +
    '.tabs{display:flex;background:white;border-bottom:1px solid #e0e0e0}' +
    '.tab{flex:1;padding:clamp(10px,2.5vw,14px);text-align:center;font-size:clamp(12px,2.5vw,14px);font-weight:500;color:#666;border:none;background:none;cursor:pointer;border-bottom:3px solid transparent;min-height:' + MOBILE_CONFIG.TOUCH_TARGET_SIZE + '}' +
    '.tab.active{color:#1a73e8;border-bottom-color:#1a73e8}' +
    '.results{padding:10px;display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:10px}' +
    '.result-card{background:white;padding:15px;border-radius:12px;box-shadow:0 2px 4px rgba(0,0,0,0.08)}' +
    '.result-title{font-weight:bold;color:#1a73e8;margin-bottom:5px;font-size:clamp(14px,3vw,16px)}' +
    '.result-detail{font-size:clamp(11px,2.5vw,13px);color:#666;margin:3px 0}' +
    '.empty-state{text-align:center;padding:60px;color:#999;grid-column:1/-1}' +
    '@media (min-width:768px){.results{grid-template-columns:repeat(2,1fr)}}' +
    '@media (min-width:1024px){.results{grid-template-columns:repeat(3,1fr)}}' +
    '</style></head><body>' +
    '<div class="header"><h2>🔍 Search</h2><div class="search-container"><span class="search-icon">🔍</span><input type="text" class="search-input" id="q" placeholder="Search members or grievances..." oninput="search(this.value)"></div></div>' +
    '<div class="tabs"><button class="tab active" onclick="setTab(\'all\',this)">All</button><button class="tab" onclick="setTab(\'members\',this)">Members</button><button class="tab" onclick="setTab(\'grievances\',this)">Grievances</button></div>' +
    '<div class="results" id="results"><div class="empty-state">Type to search...</div></div>' +
    '<script>var tab="all";function setTab(t,btn){tab=t;document.querySelectorAll(".tab").forEach(function(tb){tb.classList.remove("active")});btn.classList.add("active");search(document.getElementById("q").value)}function search(q){if(!q||q.length<2){document.getElementById("results").innerHTML="<div class=\'empty-state\'>Type to search...</div>";return}google.script.run.withSuccessHandler(function(data){render(data)}).getMobileSearchData(q,tab)}function render(data){var c=document.getElementById("results");if(!data||data.length===0){c.innerHTML="<div class=\'empty-state\'>No results</div>";return}c.innerHTML=data.map(function(r){return"<div class=\'result-card\'><div class=\'result-title\'>"+(r.type==="member"?"👤 ":"📋 ")+r.title+"</div><div class=\'result-detail\'>"+r.subtitle+"</div>"+(r.detail?"<div class=\'result-detail\'>"+r.detail+"</div>":"")+"</div>"}).join("")}</script></body></html>'
  ).setWidth(800).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '🔍 Search');
}

function getMobileSearchData(query, tab) {
  var results = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  query = query.toLowerCase();
  if (tab === 'all' || tab === 'members') {
    var mSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (mSheet && mSheet.getLastRow() > 1) {
      var mData = mSheet.getRange(2, 1, mSheet.getLastRow() - 1, MEMBER_COLS.EMAIL).getValues();
      mData.forEach(function(row) {
        var id = row[MEMBER_COLS.MEMBER_ID - 1] || '';
        var name = (row[MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' + (row[MEMBER_COLS.LAST_NAME - 1] || '');
        var email = row[MEMBER_COLS.EMAIL - 1] || '';
        if (id.toLowerCase().indexOf(query) >= 0 || name.toLowerCase().indexOf(query) >= 0 || email.toLowerCase().indexOf(query) >= 0) {
          results.push({ type: 'member', title: name, subtitle: id, detail: email });
        }
      });
    }
  }
  if (tab === 'all' || tab === 'grievances') {
    var gSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (gSheet && gSheet.getLastRow() > 1) {
      var gData = gSheet.getRange(2, 1, gSheet.getLastRow() - 1, GRIEVANCE_COLS.STATUS).getValues();
      gData.forEach(function(row) {
        var id = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '';
        var name = (row[GRIEVANCE_COLS.FIRST_NAME - 1] || '') + ' ' + (row[GRIEVANCE_COLS.LAST_NAME - 1] || '');
        var status = row[GRIEVANCE_COLS.STATUS - 1] || '';
        if (id.toLowerCase().indexOf(query) >= 0 || name.toLowerCase().indexOf(query) >= 0) {
          results.push({ type: 'grievance', title: id, subtitle: name, detail: status });
        }
      });
    }
  }
  return results.slice(0, 20);
}

function showMyAssignedGrievances() {
  var email = Session.getActiveUser().getEmail();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet || sheet.getLastRow() <= 1) { SpreadsheetApp.getUi().alert('No grievances found'); return; }
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, GRIEVANCE_COLS.STEWARD).getValues();
  var mine = data.filter(function(row) { var steward = row[GRIEVANCE_COLS.STEWARD - 1]; return steward && steward.indexOf(email) >= 0; });
  if (mine.length === 0) { SpreadsheetApp.getUi().alert('No grievances assigned to you'); return; }
  var msg = 'You have ' + mine.length + ' assigned grievance(s):\n\n';
  mine.slice(0, 10).forEach(function(row) { msg += '#' + row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] + ' - ' + row[GRIEVANCE_COLS.FIRST_NAME - 1] + ' ' + row[GRIEVANCE_COLS.LAST_NAME - 1] + ' (' + row[GRIEVANCE_COLS.STATUS - 1] + ')\n'; });
  if (mine.length > 10) msg += '\n... and ' + (mine.length - 10) + ' more';
  SpreadsheetApp.getUi().alert('My Cases', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ==================== QUICK ACTIONS ====================

/**
 * Show context-aware Quick Actions menu
 *
 * HOW IT WORKS:
 * Quick Actions provides contextual shortcuts based on your current selection.
 *
 * AVAILABLE ON:
 * - Member Directory: Start new grievance, send email, view history, copy ID
 * - Grievance Log: Sync to calendar, setup folder, update status, copy ID
 *
 * HOW TO USE:
 * 1. Navigate to Member Directory or Grievance Log
 * 2. Click on any data row (not the header)
 * 3. Run Quick Actions from the menu
 * 4. A popup will show relevant actions for that row
 */
function showQuickActionsMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getActiveSheet();
  var name = sheet.getName();
  var selection = sheet.getActiveRange();

  if (!selection) {
    ui.alert('⚡ Quick Actions - How to Use',
      'Quick Actions provides contextual shortcuts for the selected row.\n\n' +
      'TO USE:\n' +
      '1. Go to Member Directory or Grievance Log\n' +
      '2. Click on a data row (not the header)\n' +
      '3. Run this menu item again\n\n' +
      'MEMBER DIRECTORY ACTIONS:\n' +
      '• Start new grievance for member\n' +
      '• Send email to member\n' +
      '• View grievance history\n' +
      '• Copy Member ID\n\n' +
      'GRIEVANCE LOG ACTIONS:\n' +
      '• Sync deadlines to calendar\n' +
      '• Setup Drive folder\n' +
      '• Quick status update\n' +
      '• Copy Grievance ID',
      ui.ButtonSet.OK);
    return;
  }

  var row = selection.getRow();
  if (row < 2) {
    ui.alert('Quick Actions',
      'Please select a data row, not the header row.\n\n' +
      'Click on row 2 or below to use Quick Actions.',
      ui.ButtonSet.OK);
    return;
  }

  if (name === SHEETS.MEMBER_DIR) {
    showMemberQuickActions(row);
  } else if (name === SHEETS.GRIEVANCE_LOG) {
    showGrievanceQuickActions(row);
  } else {
    ui.alert('⚡ Quick Actions',
      'Quick Actions is available for:\n\n' +
      '• Member Directory - actions for members\n' +
      '• Grievance Log - actions for grievances\n\n' +
      'Current sheet: ' + name + '\n\n' +
      'Navigate to one of the supported sheets and select a row.',
      ui.ButtonSet.OK);
  }
}

function showMemberQuickActions(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  var memberId = data[MEMBER_COLS.MEMBER_ID - 1];
  var name = data[MEMBER_COLS.FIRST_NAME - 1] + ' ' + data[MEMBER_COLS.LAST_NAME - 1];
  var email = data[MEMBER_COLS.EMAIL - 1];
  var hasOpen = data[MEMBER_COLS.HAS_OPEN_GRIEVANCE - 1];
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:Arial;padding:20px;background:#f5f5f5}.container{background:white;padding:25px;border-radius:8px}h2{color:#1a73e8;display:flex;align-items:center;gap:10px}.info{background:#e8f4fd;padding:15px;border-radius:8px;margin-bottom:20px}.name{font-size:18px;font-weight:bold}.id{color:#666;font-size:14px}.status{margin-top:10px}.badge{display:inline-block;padding:3px 10px;border-radius:12px;font-size:12px;font-weight:bold}.open{background:#ffebee;color:#c62828}.none{background:#e8f5e9;color:#2e7d32}.actions{display:flex;flex-direction:column;gap:10px}.action-btn{display:flex;align-items:center;gap:12px;padding:15px;border:none;border-radius:8px;cursor:pointer;font-size:14px;text-align:left;background:#f8f9fa}.action-btn:hover{background:#e8f4fd}.icon{font-size:24px}.title{font-weight:bold}.desc{font-size:12px;color:#666;margin-top:2px}.close{width:100%;margin-top:15px;padding:12px;background:#6c757d;color:white;border:none;border-radius:8px;cursor:pointer}</style></head><body><div class="container"><h2>⚡ Quick Actions</h2><div class="info"><div class="name">' + name + '</div><div class="id">' + memberId + ' | ' + (email || 'No email') + '</div><div class="status">' + (hasOpen === 'Yes' ? '<span class="badge open">🔴 Has Open Grievance</span>' : '<span class="badge none">🟢 No Open Grievances</span>') + '</div></div><div class="actions"><button class="action-btn" onclick="google.script.run.openGrievanceFormForMember(' + row + ');google.script.host.close()"><span class="icon">📋</span><span><div class="title">Start New Grievance</div><div class="desc">Create a grievance for this member</div></span></button>' + (email ? '<button class="action-btn" onclick="google.script.run.composeEmailForMember(\'' + memberId + '\');google.script.host.close()"><span class="icon">📧</span><span><div class="title">Send Email</div><div class="desc">Compose email to ' + email + '</div></span></button>' : '') + '<button class="action-btn" onclick="google.script.run.showMemberGrievanceHistory(\'' + memberId + '\');google.script.host.close()"><span class="icon">📁</span><span><div class="title">View Grievance History</div><div class="desc">See all grievances for this member</div></span></button><button class="action-btn" onclick="navigator.clipboard.writeText(\'' + memberId + '\');alert(\'Copied!\')"><span class="icon">📋</span><span><div class="title">Copy Member ID</div><div class="desc">' + memberId + '</div></span></button></div><button class="close" onclick="google.script.host.close()">Close</button></div></body></html>'
  ).setWidth(400).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Member Quick Actions');
}

function showGrievanceQuickActions(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  var grievanceId = data[GRIEVANCE_COLS.GRIEVANCE_ID - 1];
  var memberId = data[GRIEVANCE_COLS.MEMBER_ID - 1];
  var name = data[GRIEVANCE_COLS.FIRST_NAME - 1] + ' ' + data[GRIEVANCE_COLS.LAST_NAME - 1];
  var status = data[GRIEVANCE_COLS.STATUS - 1];
  var step = data[GRIEVANCE_COLS.CURRENT_STEP - 1];
  var daysTo = data[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
  var isOpen = status === 'Open';
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:Arial;padding:20px;background:#f5f5f5}.container{background:white;padding:25px;border-radius:8px}h2{color:#DC2626}.info{background:#fff5f5;padding:15px;border-radius:8px;margin-bottom:20px;border-left:4px solid #DC2626}.gid{font-size:18px;font-weight:bold}.gmem{color:#666;font-size:14px}.gstatus{margin-top:10px;display:flex;gap:10px;flex-wrap:wrap}.badge{display:inline-block;padding:3px 10px;border-radius:12px;font-size:12px;font-weight:bold}.actions{display:flex;flex-direction:column;gap:10px}.action-btn{display:flex;align-items:center;gap:12px;padding:15px;border:none;border-radius:8px;cursor:pointer;font-size:14px;text-align:left;background:#f8f9fa}.action-btn:hover{background:#fff5f5}.icon{font-size:24px}.title{font-weight:bold}.desc{font-size:12px;color:#666;margin-top:2px}.divider{height:1px;background:#e0e0e0;margin:10px 0}.status-section{margin-top:15px;padding:15px;background:#f8f9fa;border-radius:8px}.status-section h4{margin:0 0 10px}select{width:100%;padding:10px;border:2px solid #ddd;border-radius:4px;font-size:14px}.close{width:100%;margin-top:15px;padding:12px;background:#6c757d;color:white;border:none;border-radius:8px;cursor:pointer}</style></head><body><div class="container"><h2>⚡ Grievance Actions</h2><div class="info"><div class="gid">' + grievanceId + '</div><div class="gmem">' + name + ' (' + memberId + ')</div><div class="gstatus"><span class="badge">' + status + '</span><span class="badge">' + step + '</span>' + (daysTo !== null && daysTo !== '' ? '<span class="badge" style="background:' + (daysTo < 0 ? '#ffebee;color:#c62828' : '#e3f2fd;color:#1565c0') + '">' + (daysTo < 0 ? '⚠️ Overdue' : '📅 ' + daysTo + ' days') + '</span>' : '') + '</div></div><div class="actions"><button class="action-btn" onclick="google.script.run.syncSingleGrievanceToCalendar(\'' + grievanceId + '\');google.script.host.close()"><span class="icon">📅</span><span><div class="title">Sync to Calendar</div><div class="desc">Add deadlines to Google Calendar</div></span></button><button class="action-btn" onclick="google.script.run.setupDriveFolderForGrievance();google.script.host.close()"><span class="icon">📁</span><span><div class="title">Setup Drive Folder</div><div class="desc">Create document folder</div></span></button><button class="action-btn" onclick="navigator.clipboard.writeText(\'' + grievanceId + '\');alert(\'Copied!\')"><span class="icon">📋</span><span><div class="title">Copy Grievance ID</div><div class="desc">' + grievanceId + '</div></span></button></div>' + (isOpen ? '<div class="status-section"><h4>Quick Status Update</h4><select id="statusSelect"><option value="">-- Select --</option><option value="Open">Open</option><option value="Pending Info">Pending Info</option><option value="Settled">Settled</option><option value="Withdrawn">Withdrawn</option><option value="Closed">Closed</option></select><button class="action-btn" style="margin-top:10px" onclick="var s=document.getElementById(\'statusSelect\').value;if(!s){alert(\'Select status\');return}google.script.run.withSuccessHandler(function(){alert(\'Updated!\');google.script.host.close()}).quickUpdateGrievanceStatus(' + row + ',s)"><span class="icon">✓</span><span><div class="title">Update Status</div></span></button></div>' : '') + '<button class="close" onclick="google.script.host.close()">Close</button></div></body></html>'
  ).setWidth(400).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Grievance Quick Actions');
}

function quickUpdateGrievanceStatus(row, newStatus) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) throw new Error('Grievance Log not found');
  sheet.getRange(row, GRIEVANCE_COLS.STATUS).setValue(newStatus);
  if (['Closed', 'Settled', 'Withdrawn'].indexOf(newStatus) >= 0) {
    var closeCol = GRIEVANCE_COLS.DATE_CLOSED;
    if (!sheet.getRange(row, closeCol).getValue()) sheet.getRange(row, closeCol).setValue(new Date());
  }
  ss.toast('Grievance status updated to: ' + newStatus, 'Status Updated', 3);
}

function composeEmailForMember(memberId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet) return;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, MEMBER_COLS.EMAIL).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][MEMBER_COLS.MEMBER_ID - 1] === memberId) {
      var email = data[i][MEMBER_COLS.EMAIL - 1];
      var name = data[i][MEMBER_COLS.FIRST_NAME - 1] + ' ' + data[i][MEMBER_COLS.LAST_NAME - 1];
      if (!email) { SpreadsheetApp.getUi().alert('No email on file.'); return; }
      var html = HtmlService.createHtmlOutput(
        '<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:Arial;padding:20px;background:#f5f5f5}.container{background:white;padding:25px;border-radius:8px}h2{color:#1a73e8}.info{background:#e8f4fd;padding:15px;border-radius:8px;margin-bottom:20px}.form-group{margin:15px 0}label{display:block;font-weight:bold;margin-bottom:5px}input,textarea{width:100%;padding:10px;border:2px solid #ddd;border-radius:4px;font-size:14px;box-sizing:border-box}textarea{min-height:200px}input:focus,textarea:focus{outline:none;border-color:#1a73e8}.buttons{display:flex;gap:10px;margin-top:20px}button{padding:12px 24px;font-size:14px;border:none;border-radius:4px;cursor:pointer;flex:1}.primary{background:#1a73e8;color:white}.secondary{background:#6c757d;color:white}</style></head><body><div class="container"><h2>📧 Email to Member</h2><div class="info"><strong>' + name + '</strong> (' + memberId + ')<br>' + email + '</div><div class="form-group"><label>Subject:</label><input type="text" id="subject" placeholder="Email subject"></div><div class="form-group"><label>Message:</label><textarea id="message" placeholder="Type your message..."></textarea></div><div class="buttons"><button class="primary" onclick="send()">📤 Send</button><button class="secondary" onclick="google.script.host.close()">Cancel</button></div></div><script>function send(){var s=document.getElementById("subject").value.trim();var m=document.getElementById("message").value.trim();if(!s||!m){alert("Fill in subject and message");return}google.script.run.withSuccessHandler(function(){alert("Email sent!");google.script.host.close()}).withFailureHandler(function(e){alert("Error: "+e.message)}).sendQuickEmail("' + email + '",s,m,"' + memberId + '")}</script></body></html>'
      ).setWidth(600).setHeight(500);
      SpreadsheetApp.getUi().showModalDialog(html, '📧 Compose Email');
      return;
    }
  }
}

function sendQuickEmail(to, subject, body, memberId) {
  try {
    MailApp.sendEmail({ to: to, subject: subject, body: body, name: 'SEIU Local 509 Dashboard' });
    return { success: true };
  } catch (e) { throw new Error('Failed to send: ' + e.message); }
}

function showMemberGrievanceHistory(memberId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet || sheet.getLastRow() <= 1) { SpreadsheetApp.getUi().alert('No grievances found.'); return; }
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, GRIEVANCE_COLS.DATE_CLOSED).getValues();
  var mine = [];
  data.forEach(function(row) {
    if (row[GRIEVANCE_COLS.MEMBER_ID - 1] === memberId) {
      mine.push({ id: row[GRIEVANCE_COLS.GRIEVANCE_ID - 1], status: row[GRIEVANCE_COLS.STATUS - 1], step: row[GRIEVANCE_COLS.CURRENT_STEP - 1], issue: row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1], filed: row[GRIEVANCE_COLS.DATE_FILED - 1], closed: row[GRIEVANCE_COLS.DATE_CLOSED - 1] });
    }
  });
  if (mine.length === 0) { SpreadsheetApp.getUi().alert('No grievances for this member.'); return; }
  var list = mine.map(function(g) {
    return '<div style="background:#f8f9fa;padding:12px;margin:8px 0;border-radius:4px;border-left:4px solid ' + (g.status === 'Open' ? '#f44336' : '#4caf50') + '"><strong>' + g.id + '</strong><br><span style="color:#666">Status: ' + g.status + ' | Step: ' + g.step + '</span><br><span style="color:#888;font-size:12px">' + g.issue + ' | Filed: ' + (g.filed ? new Date(g.filed).toLocaleDateString() : 'N/A') + '</span></div>';
  }).join('');
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:Arial;padding:20px}h2{color:#1a73e8}.summary{background:#e8f4fd;padding:15px;border-radius:8px;margin-bottom:20px}</style></head><body><h2>📁 Grievance History</h2><div class="summary"><strong>Member ID:</strong> ' + memberId + '<br><strong>Total:</strong> ' + mine.length + '<br><strong>Open:</strong> ' + mine.filter(function(g) { return g.status === 'Open'; }).length + '<br><strong>Closed:</strong> ' + mine.filter(function(g) { return g.status !== 'Open'; }).length + '</div>' + list + '</body></html>'
  ).setWidth(500).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Grievance History - ' + memberId);
}

function openGrievanceFormForMember(row) {
  SpreadsheetApp.getUi().alert('ℹ️ New Grievance', 'To start a new grievance for this member, navigate to the Grievance Log sheet and add a new row with their Member ID.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function syncSingleGrievanceToCalendar(grievanceId) {
  SpreadsheetApp.getActiveSpreadsheet().toast('📅 Syncing ' + grievanceId + '...', 'Calendar', 3);
  if (typeof syncDeadlinesToCalendar === 'function') syncDeadlinesToCalendar();
}

// ╔═══════════════════════════════════════════════════════════════════════════╗
// ║                                                                           ║
// ║   ██████╗ ██████╗  ██████╗ ████████╗███████╗ ██████╗████████╗███████╗██████╗  ║
// ║   ██╔══██╗██╔══██╗██╔═══██╗╚══██╔══╝██╔════╝██╔════╝╚══██╔══╝██╔════╝██╔══██╗ ║
// ║   ██████╔╝██████╔╝██║   ██║   ██║   █████╗  ██║        ██║   █████╗  ██║  ██║ ║
// ║   ██╔═══╝ ██╔══██╗██║   ██║   ██║   ██╔══╝  ██║        ██║   ██╔══╝  ██║  ██║ ║
// ║   ██║     ██║  ██║╚██████╔╝   ██║   ███████╗╚██████╗   ██║   ███████╗██████╔╝ ║
// ║   ╚═╝     ╚═╝  ╚═╝ ╚═════╝    ╚═╝   ╚══════╝ ╚═════╝   ╚═╝   ╚══════╝╚═════╝  ║
// ║                                                                           ║
// ║         ⚠️  DO NOT MODIFY THIS SECTION - PROTECTED CODE  ⚠️              ║
// ║                                                                           ║
// ╠═══════════════════════════════════════════════════════════════════════════╣
// ║  INTERACTIVE DASHBOARD TAB - Modal Popup with Tabbed Interface           ║
// ╠═══════════════════════════════════════════════════════════════════════════╣
// ║                                                                           ║
// ║  This code block is PROTECTED and should NOT be modified or removed.     ║
// ║                                                                           ║
// ║  Protected Functions:                                                     ║
// ║  • showInteractiveDashboardTab() - Opens the modal dialog                 ║
// ║  • getInteractiveDashboardHtml() - Returns the HTML/CSS/JS for the UI     ║
// ║  • getInteractiveOverviewData()  - Fetches overview statistics            ║
// ║  • getInteractiveMemberData()    - Fetches member list data               ║
// ║  • getInteractiveGrievanceData() - Fetches grievance list data            ║
// ║  • getInteractiveAnalyticsData() - Fetches analytics/charts data          ║
// ║                                                                           ║
// ║  Features:                                                                ║
// ║  • 4 Tabs: Overview, Members, Grievances, Analytics                       ║
// ║  • Live search and status filtering                                       ║
// ║  • Mobile-responsive design with touch targets                            ║
// ║  • Bar charts for status distribution and categories                      ║
// ║                                                                           ║
// ║  Menu Location: 👤 Dashboard > 🎯 Custom View                  ║
// ║                                                                           ║
// ║  Added: December 29, 2025 (commit c75c1cc)                                ║
// ║  Status: USER APPROVED - DO NOT CHANGE                                    ║
// ║                                                                           ║
// ╚═══════════════════════════════════════════════════════════════════════════╝

/**
 * Shows the Custom View with tabbed interface
 * Features: Overview, Members, Grievances, and Analytics tabs
 *
 * ⚠️ PROTECTED FUNCTION - DO NOT MODIFY ⚠️
 */
function showInteractiveDashboardTab() {
  var html = HtmlService.createHtmlOutput(getInteractiveDashboardHtml())
    .setWidth(900)
    .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, '🎯 Custom View');
}

/**
 * Returns the HTML for the interactive dashboard with tabs
 */
function getInteractiveDashboardHtml() {
  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<base target="_top">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<style>' +
    // CSS Reset and base styles
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh}' +

    // Header
    '.header{background:linear-gradient(135deg,#7C3AED,#5B21B6);color:white;padding:20px;text-align:center}' +
    '.header h1{font-size:clamp(18px,4vw,24px);margin-bottom:5px}' +
    '.header .subtitle{font-size:clamp(11px,2.5vw,13px);opacity:0.9}' +

    // Tab navigation
    '.tabs{display:flex;background:white;border-bottom:2px solid #e0e0e0;position:sticky;top:0;z-index:100}' +
    '.tab{flex:1;padding:clamp(12px,3vw,16px);text-align:center;font-size:clamp(12px,2.5vw,14px);font-weight:600;color:#666;' +
    'border:none;background:none;cursor:pointer;border-bottom:3px solid transparent;transition:all 0.2s;min-height:' + MOBILE_CONFIG.TOUCH_TARGET_SIZE + '}' +
    '.tab:hover{background:#f8f9fa;color:#7C3AED}' +
    '.tab.active{color:#7C3AED;border-bottom-color:#7C3AED;background:#f8f4ff}' +
    '.tab-icon{display:block;font-size:18px;margin-bottom:4px}' +

    // Tab content
    '.tab-content{display:none;padding:15px;animation:fadeIn 0.3s}' +
    '.tab-content.active{display:block}' +
    '@keyframes fadeIn{from{opacity:0}to{opacity:1}}' +

    // Stats grid
    '.stats-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:12px;margin-bottom:20px}' +
    '.stat-card{background:white;padding:20px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);text-align:center;transition:transform 0.2s}' +
    '.stat-card:hover{transform:translateY(-2px)}' +
    '.stat-value{font-size:clamp(24px,5vw,32px);font-weight:bold;color:#7C3AED}' +
    '.stat-label{font-size:clamp(10px,2vw,12px);color:#666;text-transform:uppercase;margin-top:5px}' +
    '.stat-card.green .stat-value{color:#059669}' +
    '.stat-card.red .stat-value{color:#DC2626}' +
    '.stat-card.orange .stat-value{color:#F97316}' +

    // Data table
    '.data-table{width:100%;border-collapse:collapse;background:white;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08)}' +
    '.data-table th{background:#7C3AED;color:white;padding:12px;text-align:left;font-size:13px}' +
    '.data-table td{padding:12px;border-bottom:1px solid #eee;font-size:13px}' +
    '.data-table tr:hover{background:#f8f4ff}' +
    '.data-table tr:last-child td{border-bottom:none}' +

    // Status badges
    '.badge{display:inline-block;padding:4px 10px;border-radius:12px;font-size:11px;font-weight:bold}' +
    '.badge-open{background:#fee2e2;color:#dc2626}' +
    '.badge-pending{background:#fef3c7;color:#d97706}' +
    '.badge-closed{background:#d1fae5;color:#059669}' +

    // Action buttons
    '.action-btn{display:inline-flex;align-items:center;gap:8px;padding:10px 16px;border:none;border-radius:8px;' +
    'cursor:pointer;font-size:13px;font-weight:500;transition:all 0.2s;min-height:' + MOBILE_CONFIG.TOUCH_TARGET_SIZE + '}' +
    '.action-btn-primary{background:#7C3AED;color:white}' +
    '.action-btn-primary:hover{background:#5B21B6}' +
    '.action-btn-secondary{background:#f3f4f6;color:#374151}' +
    '.action-btn-secondary:hover{background:#e5e7eb}' +

    // List items
    '.list-container{display:flex;flex-direction:column;gap:10px}' +
    '.list-item{background:white;padding:15px;border-radius:8px;box-shadow:0 2px 4px rgba(0,0,0,0.06);display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px}' +
    '.list-item:hover{box-shadow:0 4px 8px rgba(0,0,0,0.1)}' +
    '.list-item-main{flex:1;min-width:200px}' +
    '.list-item-title{font-weight:600;color:#1f2937;margin-bottom:3px}' +
    '.list-item-subtitle{font-size:12px;color:#666}' +

    // Search input
    '.search-container{position:relative;margin-bottom:15px}' +
    '.search-input{width:100%;padding:12px 12px 12px 40px;border:2px solid #e5e7eb;border-radius:8px;font-size:14px;transition:border-color 0.2s}' +
    '.search-input:focus{outline:none;border-color:#7C3AED}' +
    '.search-icon{position:absolute;left:12px;top:50%;transform:translateY(-50%);font-size:16px;color:#9ca3af}' +

    // Charts section
    '.chart-container{background:white;padding:20px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.08);margin-bottom:15px}' +
    '.chart-title{font-weight:600;color:#1f2937;margin-bottom:15px;font-size:14px}' +
    '.bar-chart{display:flex;flex-direction:column;gap:10px}' +
    '.bar-row{display:flex;align-items:center;gap:10px}' +
    '.bar-label{width:100px;font-size:12px;color:#666;text-align:right}' +
    '.bar-container{flex:1;background:#e5e7eb;border-radius:4px;height:24px;overflow:hidden}' +
    '.bar-fill{height:100%;border-radius:4px;transition:width 0.5s}' +
    '.bar-value{width:50px;font-size:12px;font-weight:600;color:#374151}' +

    // Empty state
    '.empty-state{text-align:center;padding:40px;color:#9ca3af}' +
    '.empty-state-icon{font-size:48px;margin-bottom:10px}' +

    // Loading
    '.loading{text-align:center;padding:40px;color:#666}' +
    '.spinner{display:inline-block;width:24px;height:24px;border:3px solid #e5e7eb;border-top-color:#7C3AED;border-radius:50%;animation:spin 1s linear infinite}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +

    // Sankey Diagram
    '.sankey-container{position:relative;padding:20px 0}' +
    '.sankey-nodes{display:flex;justify-content:space-between;position:relative;z-index:2}' +
    '.sankey-column{display:flex;flex-direction:column;gap:8px;align-items:center}' +
    '.sankey-node{padding:10px 15px;border-radius:8px;color:white;font-weight:600;font-size:12px;text-align:center;min-width:80px;box-shadow:0 2px 4px rgba(0,0,0,0.2)}' +
    '.sankey-node.source{background:linear-gradient(135deg,#7C3AED,#9333EA)}' +
    '.sankey-node.status-open{background:linear-gradient(135deg,#dc2626,#ef4444)}' +
    '.sankey-node.status-pending{background:linear-gradient(135deg,#f97316,#fb923c)}' +
    '.sankey-node.status-closed{background:linear-gradient(135deg,#059669,#10b981)}' +
    '.sankey-node.resolution{background:linear-gradient(135deg,#1d4ed8,#3b82f6)}' +
    '.sankey-flows{position:absolute;top:0;left:0;right:0;bottom:0;z-index:1}' +
    '.sankey-flow{position:absolute;height:4px;border-radius:2px;opacity:0.6;transition:opacity 0.2s}' +
    '.sankey-flow:hover{opacity:1}' +
    '.sankey-flow.flow-filed{background:linear-gradient(90deg,#7C3AED,#dc2626)}' +
    '.sankey-flow.flow-open{background:linear-gradient(90deg,#dc2626,#1d4ed8)}' +
    '.sankey-flow.flow-pending{background:linear-gradient(90deg,#f97316,#059669)}' +
    '.sankey-flow.flow-closed{background:linear-gradient(90deg,#059669,#1d4ed8)}' +
    '.sankey-label{font-size:11px;color:#666;margin-top:5px}' +
    '.sankey-legend{display:flex;flex-wrap:wrap;gap:15px;justify-content:center;margin-top:15px;padding-top:15px;border-top:1px solid #e5e7eb}' +
    '.sankey-legend-item{display:flex;align-items:center;gap:5px;font-size:11px;color:#666}' +
    '.sankey-legend-color{width:12px;height:12px;border-radius:3px}' +

    // Responsive
    '@media (max-width:600px){' +
    '  .stats-grid{grid-template-columns:repeat(2,1fr)}' +
    '  .list-item{flex-direction:column;align-items:flex-start}' +
    '  .tab-icon{font-size:16px}' +
    '}' +

    '</style>' +
    '</head><body>' +

    // Header
    '<div class="header">' +
    '<h1>🎯 Custom View</h1>' +
    '<div class="subtitle">Real-time union data at your fingertips</div>' +
    '</div>' +

    // Tab Navigation
    '<div class="tabs">' +
    '<button class="tab active" onclick="switchTab(\'overview\',this)" id="tab-overview"><span class="tab-icon">📊</span>Overview</button>' +
    '<button class="tab" onclick="switchTab(\'members\',this)" id="tab-members"><span class="tab-icon">👥</span>Members</button>' +
    '<button class="tab" onclick="switchTab(\'grievances\',this)" id="tab-grievances"><span class="tab-icon">📋</span>Grievances</button>' +
    '<button class="tab" onclick="switchTab(\'analytics\',this)" id="tab-analytics"><span class="tab-icon">📈</span>Analytics</button>' +
    '</div>' +

    // Overview Tab
    '<div class="tab-content active" id="content-overview">' +
    '<div class="stats-grid" id="overview-stats"><div class="loading"><div class="spinner"></div><p>Loading stats...</p></div></div>' +
    '<div id="overview-actions" style="margin-top:15px;display:flex;flex-wrap:wrap;gap:10px"></div>' +
    '</div>' +

    // Members Tab
    '<div class="tab-content" id="content-members">' +
    '<div class="search-container"><span class="search-icon">🔍</span><input type="text" class="search-input" id="member-search" placeholder="Search members..." oninput="filterMembers(this.value)"></div>' +
    '<div class="list-container" id="members-list"><div class="loading"><div class="spinner"></div><p>Loading members...</p></div></div>' +
    '</div>' +

    // Grievances Tab
    '<div class="tab-content" id="content-grievances">' +
    '<div class="search-container"><span class="search-icon">🔍</span><input type="text" class="search-input" id="grievance-search" placeholder="Search grievances..." oninput="filterGrievances(this.value)"></div>' +
    '<div style="display:flex;gap:8px;margin-bottom:15px;flex-wrap:wrap">' +
    '<button class="action-btn action-btn-primary" onclick="filterGrievanceStatus(\'all\')">All</button>' +
    '<button class="action-btn action-btn-secondary" onclick="filterGrievanceStatus(\'Open\')">Open</button>' +
    '<button class="action-btn action-btn-secondary" onclick="filterGrievanceStatus(\'Pending Info\')">Pending</button>' +
    '<button class="action-btn action-btn-secondary" onclick="filterGrievanceStatus(\'Closed\')">Closed</button>' +
    '</div>' +
    '<div class="list-container" id="grievances-list"><div class="loading"><div class="spinner"></div><p>Loading grievances...</p></div></div>' +
    '</div>' +

    // Analytics Tab
    '<div class="tab-content" id="content-analytics">' +
    '<div id="analytics-charts"><div class="loading"><div class="spinner"></div><p>Loading analytics...</p></div></div>' +
    '</div>' +

    // JavaScript
    '<script>' +
    'var allMembers=[];var allGrievances=[];var currentGrievanceFilter="all";' +

    // Tab switching
    'function switchTab(tabName,btn){' +
    '  document.querySelectorAll(".tab").forEach(function(t){t.classList.remove("active")});' +
    '  document.querySelectorAll(".tab-content").forEach(function(c){c.classList.remove("active")});' +
    '  btn.classList.add("active");' +
    '  document.getElementById("content-"+tabName).classList.add("active");' +
    '  if(tabName==="members"&&allMembers.length===0)loadMembers();' +
    '  if(tabName==="grievances"&&allGrievances.length===0)loadGrievances();' +
    '  if(tabName==="analytics")loadAnalytics();' +
    '}' +

    // Load overview data
    'function loadOverview(){' +
    '  google.script.run.withSuccessHandler(function(data){renderOverview(data)}).getInteractiveOverviewData();' +
    '}' +

    // Render overview
    'function renderOverview(data){' +
    '  var html="";' +
    '  html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+data.totalMembers+"</div><div class=\\"stat-label\\">Total Members</div></div>";' +
    '  html+="<div class=\\"stat-card green\\"><div class=\\"stat-value\\">"+data.activeStewards+"</div><div class=\\"stat-label\\">Stewards</div></div>";' +
    '  html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+data.totalGrievances+"</div><div class=\\"stat-label\\">Total Grievances</div></div>";' +
    '  html+="<div class=\\"stat-card red\\"><div class=\\"stat-value\\">"+data.openGrievances+"</div><div class=\\"stat-label\\">Open Cases</div></div>";' +
    '  html+="<div class=\\"stat-card orange\\"><div class=\\"stat-value\\">"+data.pendingInfo+"</div><div class=\\"stat-label\\">Pending Info</div></div>";' +
    '  html+="<div class=\\"stat-card green\\"><div class=\\"stat-value\\">"+data.winRate+"</div><div class=\\"stat-label\\">Win Rate</div></div>";' +
    '  document.getElementById("overview-stats").innerHTML=html;' +
    '  var actions="";' +
    '  actions+="<button class=\\"action-btn action-btn-primary\\" onclick=\\"google.script.run.showMobileUnifiedSearch()\\">🔍 Quick Search</button>";' +
    '  actions+="<button class=\\"action-btn action-btn-secondary\\" onclick=\\"google.script.run.showMobileGrievanceList()\\">📋 View All Grievances</button>";' +
    '  actions+="<button class=\\"action-btn action-btn-secondary\\" onclick=\\"google.script.run.showMyAssignedGrievances()\\">👤 My Cases</button>";' +
    '  actions+="<button class=\\"action-btn action-btn-secondary\\" onclick=\\"location.reload()\\">🔄 Refresh</button>";' +
    '  document.getElementById("overview-actions").innerHTML=actions;' +
    '}' +

    // Load members
    'function loadMembers(){' +
    '  google.script.run.withSuccessHandler(function(data){allMembers=data;renderMembers(data)}).getInteractiveMemberData();' +
    '}' +

    // Render members
    'function renderMembers(data){' +
    '  var c=document.getElementById("members-list");' +
    '  if(!data||data.length===0){c.innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-state-icon\\">👥</div><p>No members found</p></div>";return}' +
    '  c.innerHTML=data.slice(0,50).map(function(m){' +
    '    return "<div class=\\"list-item\\"><div class=\\"list-item-main\\"><div class=\\"list-item-title\\">"+m.name+"</div><div class=\\"list-item-subtitle\\">"+m.id+" • "+m.title+" • "+m.location+"</div></div><div><span class=\\"badge "+(m.isSteward?"badge-closed":"badge-open")+"\\">"+(m.isSteward?"Steward":"Member")+"</span></div></div>"' +
    '  }).join("");' +
    '  if(data.length>50)c.innerHTML+="<div class=\\"empty-state\\"><p>Showing 50 of "+data.length+" members</p></div>";' +
    '}' +

    // Filter members
    'function filterMembers(query){' +
    '  if(!query||query.length<2){renderMembers(allMembers);return}' +
    '  query=query.toLowerCase();' +
    '  var filtered=allMembers.filter(function(m){return m.name.toLowerCase().indexOf(query)>=0||m.id.toLowerCase().indexOf(query)>=0||m.title.toLowerCase().indexOf(query)>=0});' +
    '  renderMembers(filtered);' +
    '}' +

    // Load grievances
    'function loadGrievances(){' +
    '  google.script.run.withSuccessHandler(function(data){allGrievances=data;renderGrievances(data)}).getInteractiveGrievanceData();' +
    '}' +

    // Render grievances
    'function renderGrievances(data){' +
    '  var c=document.getElementById("grievances-list");' +
    '  if(!data||data.length===0){c.innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-state-icon\\">📋</div><p>No grievances found</p></div>";return}' +
    '  c.innerHTML=data.slice(0,50).map(function(g){' +
    '    var badgeClass=g.status==="Open"?"badge-open":(g.status==="Pending Info"?"badge-pending":"badge-closed");' +
    '    return "<div class=\\"list-item\\"><div class=\\"list-item-main\\"><div class=\\"list-item-title\\">"+g.id+" - "+g.memberName+"</div><div class=\\"list-item-subtitle\\">"+g.issueType+" • Filed: "+g.filedDate+"</div></div><div><span class=\\"badge "+badgeClass+"\\">"+g.status+"</span></div></div>"' +
    '  }).join("");' +
    '  if(data.length>50)c.innerHTML+="<div class=\\"empty-state\\"><p>Showing 50 of "+data.length+" grievances</p></div>";' +
    '}' +

    // Filter grievances
    'function filterGrievances(query){' +
    '  if(!query||query.length<2){applyGrievanceFilters();return}' +
    '  query=query.toLowerCase();' +
    '  var filtered=allGrievances.filter(function(g){return g.id.toLowerCase().indexOf(query)>=0||g.memberName.toLowerCase().indexOf(query)>=0||(g.issueType||"").toLowerCase().indexOf(query)>=0});' +
    '  if(currentGrievanceFilter!=="all")filtered=filtered.filter(function(g){return g.status===currentGrievanceFilter});' +
    '  renderGrievances(filtered);' +
    '}' +

    // Filter by status
    'function filterGrievanceStatus(status){' +
    '  currentGrievanceFilter=status;' +
    '  applyGrievanceFilters();' +
    '}' +

    // Apply filters
    'function applyGrievanceFilters(){' +
    '  var query=document.getElementById("grievance-search").value.toLowerCase();' +
    '  var filtered=allGrievances;' +
    '  if(currentGrievanceFilter!=="all")filtered=filtered.filter(function(g){return g.status===currentGrievanceFilter});' +
    '  if(query&&query.length>=2)filtered=filtered.filter(function(g){return g.id.toLowerCase().indexOf(query)>=0||g.memberName.toLowerCase().indexOf(query)>=0});' +
    '  renderGrievances(filtered);' +
    '}' +

    // Load analytics
    'function loadAnalytics(){' +
    '  google.script.run.withSuccessHandler(function(data){renderAnalytics(data)}).getInteractiveAnalyticsData();' +
    '}' +

    // Render analytics
    'function renderAnalytics(data){' +
    '  var c=document.getElementById("analytics-charts");' +
    '  var html="";' +
    // Member Directory Stats section
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">👥 Member Directory Statistics</div>";' +
    '  html+="<div class=\\"stats-grid\\" style=\\"margin-bottom:15px\\">";' +
    '  html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+data.memberStats.total+"</div><div class=\\"stat-label\\">Total Members</div></div>";' +
    '  html+="<div class=\\"stat-card green\\"><div class=\\"stat-value\\">"+data.memberStats.stewards+"</div><div class=\\"stat-label\\">Stewards</div></div>";' +
    '  html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+data.memberStats.withOpenGrievance+"</div><div class=\\"stat-label\\">With Open Case</div></div>";' +
    '  html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+data.memberStats.stewardRatio+"</div><div class=\\"stat-label\\">Member:Steward</div></div>";' +
    '  html+="</div></div>";' +
    // Members by Location chart
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📍 Members by Location</div><div class=\\"bar-chart\\">";' +
    '  var maxLoc=Math.max.apply(null,data.memberStats.byLocation.map(function(l){return l.count}))||1;' +
    '  data.memberStats.byLocation.forEach(function(loc){' +
    '    html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\" style=\\"width:120px\\">"+loc.name+"</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+(loc.count/maxLoc*100)+"%25;background:#059669\\"></div></div><div class=\\"bar-value\\">"+loc.count+"</div></div>";' +
    '  });' +
    '  if(data.memberStats.byLocation.length===0)html+="<div class=\\"empty-state\\">No location data</div>";' +
    '  html+="</div></div>";' +
    // Members by Unit chart
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">🏢 Members by Unit</div><div class=\\"bar-chart\\">";' +
    '  var maxUnit=Math.max.apply(null,data.memberStats.byUnit.map(function(u){return u.count}))||1;' +
    '  data.memberStats.byUnit.forEach(function(unit){' +
    '    html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\" style=\\"width:120px\\">"+unit.name+"</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+(unit.count/maxUnit*100)+"%25;background:#1a73e8\\"></div></div><div class=\\"bar-value\\">"+unit.count+"</div></div>";' +
    '  });' +
    '  if(data.memberStats.byUnit.length===0)html+="<div class=\\"empty-state\\">No unit data</div>";' +
    '  html+="</div></div>";' +
    // Status distribution chart
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📊 Grievance Status Distribution</div><div class=\\"bar-chart\\">";' +
    '  var total=data.statusCounts.open+data.statusCounts.pending+data.statusCounts.closed;' +
    '  if(total>0){' +
    '    html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\">Open</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+(data.statusCounts.open/total*100)+"%25;background:#dc2626\\"></div></div><div class=\\"bar-value\\">"+data.statusCounts.open+"</div></div>";' +
    '    html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\">Pending</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+(data.statusCounts.pending/total*100)+"%25;background:#f97316\\"></div></div><div class=\\"bar-value\\">"+data.statusCounts.pending+"</div></div>";' +
    '    html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\">Closed</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+(data.statusCounts.closed/total*100)+"%25;background:#059669\\"></div></div><div class=\\"bar-value\\">"+data.statusCounts.closed+"</div></div>";' +
    '  }else{html+="<div class=\\"empty-state\\">No grievances</div>"}' +
    '  html+="</div></div>";' +
    // Issue category chart
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">📋 Top Issue Categories</div><div class=\\"bar-chart\\">";' +
    '  var maxCat=Math.max.apply(null,data.topCategories.map(function(c){return c.count}))||1;' +
    '  data.topCategories.forEach(function(cat){' +
    '    html+="<div class=\\"bar-row\\"><div class=\\"bar-label\\" style=\\"width:120px\\">"+cat.name+"</div><div class=\\"bar-container\\"><div class=\\"bar-fill\\" style=\\"width:"+(cat.count/maxCat*100)+"%25;background:#7C3AED\\"></div></div><div class=\\"bar-value\\">"+cat.count+"</div></div>";' +
    '  });' +
    '  if(data.topCategories.length===0)html+="<div class=\\"empty-state\\">No data</div>";' +
    '  html+="</div></div>";' +
    // Resolution summary
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">🏆 Resolution Summary</div>";' +
    '  html+="<div class=\\"stats-grid\\" style=\\"margin:0\\">";' +
    '  html+="<div class=\\"stat-card green\\"><div class=\\"stat-value\\">"+data.resolutions.won+"</div><div class=\\"stat-label\\">Won</div></div>";' +
    '  html+="<div class=\\"stat-card orange\\"><div class=\\"stat-value\\">"+data.resolutions.settled+"</div><div class=\\"stat-label\\">Settled</div></div>";' +
    '  html+="<div class=\\"stat-card\\"><div class=\\"stat-value\\">"+data.resolutions.withdrawn+"</div><div class=\\"stat-label\\">Withdrawn</div></div>";' +
    '  html+="<div class=\\"stat-card red\\"><div class=\\"stat-value\\">"+data.resolutions.denied+"</div><div class=\\"stat-label\\">Denied</div></div>";' +
    '  html+="</div></div>";' +
    // Sankey Diagram - Grievance Flow
    '  var totalGrievances=data.statusCounts.open+data.statusCounts.pending+data.statusCounts.closed;' +
    '  if(totalGrievances>0){' +
    '  html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">🔀 Grievance Flow (Sankey Diagram)</div>";' +
    '  html+="<div class=\\"sankey-container\\">";' +
    '  html+="<div class=\\"sankey-nodes\\">";' +
    // Source column (Filed)
    '  html+="<div class=\\"sankey-column\\">";' +
    '  html+="<div class=\\"sankey-node source\\">Filed<br/>"+totalGrievances+"</div>";' +
    '  html+="<div class=\\"sankey-label\\">Total Filed</div>";' +
    '  html+="</div>";' +
    // Status column
    '  html+="<div class=\\"sankey-column\\">";' +
    '  if(data.statusCounts.open>0)html+="<div class=\\"sankey-node status-open\\">Open<br/>"+data.statusCounts.open+"</div>";' +
    '  if(data.statusCounts.pending>0)html+="<div class=\\"sankey-node status-pending\\">Pending<br/>"+data.statusCounts.pending+"</div>";' +
    '  if(data.statusCounts.closed>0)html+="<div class=\\"sankey-node status-closed\\">Closed<br/>"+data.statusCounts.closed+"</div>";' +
    '  html+="<div class=\\"sankey-label\\">Current Status</div>";' +
    '  html+="</div>";' +
    // Resolution column
    '  html+="<div class=\\"sankey-column\\">";' +
    '  var totalResolved=data.resolutions.won+data.resolutions.settled+data.resolutions.withdrawn+data.resolutions.denied;' +
    '  if(data.resolutions.won>0)html+="<div class=\\"sankey-node resolution\\" style=\\"background:linear-gradient(135deg,#059669,#10b981)\\">Won<br/>"+data.resolutions.won+"</div>";' +
    '  if(data.resolutions.settled>0)html+="<div class=\\"sankey-node resolution\\" style=\\"background:linear-gradient(135deg,#f97316,#fb923c)\\">Settled<br/>"+data.resolutions.settled+"</div>";' +
    '  if(data.resolutions.withdrawn>0)html+="<div class=\\"sankey-node resolution\\" style=\\"background:linear-gradient(135deg,#6b7280,#9ca3af)\\">Withdrawn<br/>"+data.resolutions.withdrawn+"</div>";' +
    '  if(data.resolutions.denied>0)html+="<div class=\\"sankey-node resolution\\" style=\\"background:linear-gradient(135deg,#dc2626,#ef4444)\\">Denied<br/>"+data.resolutions.denied+"</div>";' +
    '  if(totalResolved===0)html+="<div class=\\"sankey-node resolution\\" style=\\"background:#ccc\\">Pending<br/>Resolution</div>";' +
    '  html+="<div class=\\"sankey-label\\">Outcome</div>";' +
    '  html+="</div>";' +
    '  html+="</div>";' +  // End sankey-nodes
    // Legend
    '  html+="<div class=\\"sankey-legend\\">";' +
    '  html+="<div class=\\"sankey-legend-item\\"><div class=\\"sankey-legend-color\\" style=\\"background:#7C3AED\\"></div>Filed</div>";' +
    '  html+="<div class=\\"sankey-legend-item\\"><div class=\\"sankey-legend-color\\" style=\\"background:#dc2626\\"></div>Open</div>";' +
    '  html+="<div class=\\"sankey-legend-item\\"><div class=\\"sankey-legend-color\\" style=\\"background:#f97316\\"></div>Pending</div>";' +
    '  html+="<div class=\\"sankey-legend-item\\"><div class=\\"sankey-legend-color\\" style=\\"background:#059669\\"></div>Closed/Won</div>";' +
    '  html+="</div>";' +
    '  html+="</div></div>";' +  // End sankey-container and chart-container
    '  }' +
    '  c.innerHTML=html;' +
    '}' +

    // Initialize
    'loadOverview();' +
    '</script>' +

    '</body></html>';
}

/**
 * Get overview data for interactive dashboard
 */
function getInteractiveOverviewData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = {
    totalMembers: 0,
    activeStewards: 0,
    totalGrievances: 0,
    openGrievances: 0,
    pendingInfo: 0,
    winRate: '0%'
  };

  // Get member stats
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (memberSheet && memberSheet.getLastRow() > 1) {
    var memberData = memberSheet.getRange(2, 1, memberSheet.getLastRow() - 1, MEMBER_COLS.IS_STEWARD).getValues();
    data.totalMembers = memberData.length;
    memberData.forEach(function(row) {
      if (row[MEMBER_COLS.IS_STEWARD - 1] === 'Yes') data.activeStewards++;
    });
  }

  // Get grievance stats
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (grievanceSheet && grievanceSheet.getLastRow() > 1) {
    var grievanceData = grievanceSheet.getRange(2, 1, grievanceSheet.getLastRow() - 1, GRIEVANCE_COLS.RESOLUTION).getValues();
    data.totalGrievances = grievanceData.length;
    var wonCount = 0;
    var closedCount = 0;
    grievanceData.forEach(function(row) {
      var status = row[GRIEVANCE_COLS.STATUS - 1];
      var resolution = row[GRIEVANCE_COLS.RESOLUTION - 1] || '';
      if (status === 'Open') data.openGrievances++;
      if (status === 'Pending Info') data.pendingInfo++;
      if (status !== 'Open' && status !== 'Pending Info') closedCount++;
      if (resolution.toLowerCase().indexOf('won') >= 0) wonCount++;
    });
    if (closedCount > 0) {
      data.winRate = Math.round(wonCount / closedCount * 100) + '%';
    }
  }

  return data;
}

/**
 * Get member data for interactive dashboard
 */
function getInteractiveMemberData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, MEMBER_COLS.IS_STEWARD).getValues();
  return data.map(function(row) {
    return {
      id: row[MEMBER_COLS.MEMBER_ID - 1] || '',
      name: (row[MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' + (row[MEMBER_COLS.LAST_NAME - 1] || ''),
      title: row[MEMBER_COLS.JOB_TITLE - 1] || 'N/A',
      location: row[MEMBER_COLS.WORK_LOCATION - 1] || 'N/A',
      isSteward: row[MEMBER_COLS.IS_STEWARD - 1] === 'Yes'
    };
  }).filter(function(m) { return m.id; });
}

/**
 * Get grievance data for interactive dashboard
 */
function getInteractiveGrievanceData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, GRIEVANCE_COLS.RESOLUTION).getValues();
  return data.map(function(row) {
    var filed = row[GRIEVANCE_COLS.DATE_FILED - 1];
    return {
      id: row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '',
      memberName: (row[GRIEVANCE_COLS.FIRST_NAME - 1] || '') + ' ' + (row[GRIEVANCE_COLS.LAST_NAME - 1] || ''),
      status: row[GRIEVANCE_COLS.STATUS - 1] || 'Filed',
      issueType: row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || 'N/A',
      filedDate: filed instanceof Date ? Utilities.formatDate(filed, Session.getScriptTimeZone(), 'MM/dd/yyyy') : (filed || 'N/A')
    };
  }).filter(function(g) { return g.id; });
}

/**
 * Get analytics data for interactive dashboard
 */
function getInteractiveAnalyticsData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var data = {
    memberStats: {
      total: 0,
      stewards: 0,
      withOpenGrievance: 0,
      stewardRatio: '0:0',
      byLocation: [],
      byUnit: []
    },
    statusCounts: { open: 0, pending: 0, closed: 0 },
    topCategories: [],
    resolutions: { won: 0, settled: 0, withdrawn: 0, denied: 0 }
  };

  // Get Member Directory statistics
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (memberSheet && memberSheet.getLastRow() > 1) {
    var memberData = memberSheet.getRange(2, 1, memberSheet.getLastRow() - 1, MEMBER_COLS.HAS_OPEN_GRIEVANCE).getValues();
    var locationMap = {};
    var unitMap = {};

    memberData.forEach(function(row) {
      if (!row[MEMBER_COLS.MEMBER_ID - 1]) return; // Skip empty rows

      data.memberStats.total++;

      // Count stewards
      if (row[MEMBER_COLS.IS_STEWARD - 1] === 'Yes') data.memberStats.stewards++;

      // Count members with open grievances
      if (row[MEMBER_COLS.HAS_OPEN_GRIEVANCE - 1] === 'Yes') data.memberStats.withOpenGrievance++;

      // Count by location
      var location = row[MEMBER_COLS.WORK_LOCATION - 1] || 'Unknown';
      if (!locationMap[location]) locationMap[location] = 0;
      locationMap[location]++;

      // Count by unit
      var unit = row[MEMBER_COLS.UNIT - 1] || 'Unknown';
      if (!unitMap[unit]) unitMap[unit] = 0;
      unitMap[unit]++;
    });

    // Calculate steward ratio
    if (data.memberStats.stewards > 0) {
      var ratio = Math.round(data.memberStats.total / data.memberStats.stewards);
      data.memberStats.stewardRatio = ratio + ':1';
    } else {
      data.memberStats.stewardRatio = 'N/A';
    }

    // Get top 5 locations
    data.memberStats.byLocation = Object.keys(locationMap).map(function(key) {
      return { name: key, count: locationMap[key] };
    }).sort(function(a, b) { return b.count - a.count; }).slice(0, 5);

    // Get top 5 units
    data.memberStats.byUnit = Object.keys(unitMap).map(function(key) {
      return { name: key, count: unitMap[key] };
    }).sort(function(a, b) { return b.count - a.count; }).slice(0, 5);
  }

  // Get Grievance Log statistics
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (grievanceSheet && grievanceSheet.getLastRow() > 1) {
    var rows = grievanceSheet.getRange(2, 1, grievanceSheet.getLastRow() - 1, GRIEVANCE_COLS.RESOLUTION).getValues();
    var categoryMap = {};

    rows.forEach(function(row) {
      var status = row[GRIEVANCE_COLS.STATUS - 1] || '';
      var category = row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || 'Other';
      var resolution = (row[GRIEVANCE_COLS.RESOLUTION - 1] || '').toLowerCase();

      // Status counts
      if (status === 'Open') data.statusCounts.open++;
      else if (status === 'Pending Info') data.statusCounts.pending++;
      else data.statusCounts.closed++;

      // Category counts
      if (!categoryMap[category]) categoryMap[category] = 0;
      categoryMap[category]++;

      // Resolution counts
      if (resolution.indexOf('won') >= 0) data.resolutions.won++;
      else if (resolution.indexOf('settled') >= 0) data.resolutions.settled++;
      else if (resolution.indexOf('withdrawn') >= 0) data.resolutions.withdrawn++;
      else if (resolution.indexOf('denied') >= 0 || resolution.indexOf('lost') >= 0) data.resolutions.denied++;
    });

    // Get top 5 categories
    data.topCategories = Object.keys(categoryMap).map(function(key) {
      return { name: key, count: categoryMap[key] };
    }).sort(function(a, b) { return b.count - a.count; }).slice(0, 5);
  }

  return data;
}

// ╔═══════════════════════════════════════════════════════════════════════════╗
// ║                                                                           ║
// ║         ⚠️  END OF PROTECTED SECTION - INTERACTIVE DASHBOARD  ⚠️         ║
// ║                                                                           ║
// ╚═══════════════════════════════════════════════════════════════════════════╝
