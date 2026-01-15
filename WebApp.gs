/**
 * ============================================================================
 * WEB APP DEPLOYMENT FOR MOBILE ACCESS
 * ============================================================================
 * This file enables the dashboard to be deployed as a standalone web app
 * that can be accessed directly via URL on mobile devices.
 *
 * DEPLOYMENT INSTRUCTIONS:
 * 1. Go to Extensions → Apps Script
 * 2. Click "Deploy" → "New deployment"
 * 3. Select "Web app" as the deployment type
 * 4. Set "Execute as" to your account
 * 5. Set "Who has access" to your organization or anyone
 * 6. Click "Deploy" and copy the URL
 * 7. Bookmark this URL on your mobile device for easy access
 */

/**
 * Web app entry point - serves the mobile dashboard
 * @param {Object} e - Event object with query parameters
 * @returns {HtmlOutput} The HTML page to display
 */
function doGet(e) {
  var page = e && e.parameter && e.parameter.page ? e.parameter.page : 'dashboard';

  var html;
  switch (page) {
    case 'search':
      html = getWebAppSearchHtml();
      break;
    case 'grievances':
      html = getWebAppGrievanceListHtml();
      break;
    case 'members':
      html = getWebAppMemberListHtml();
      break;
    case 'links':
      html = getWebAppLinksHtml();
      break;
    case 'dashboard':
    default:
      html = getWebAppDashboardHtml();
      break;
  }

  return HtmlService.createHtmlOutput(html)
    .setTitle('509 Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

/**
 * Returns the main dashboard HTML for web app - FUTURISTIC REDESIGN
 * Features: Dark theme, glassmorphism, neon accents, all interactive dashboard features
 */
function getWebAppDashboardHtml() {
  var stats = getWebAppDashboardStats();
  var baseUrl = ScriptApp.getService().getUrl();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<meta name="apple-mobile-web-app-capable" content="yes">' +
    '<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">' +
    '<meta name="theme-color" content="#0a0a0f">' +
    '<title>509 Dashboard</title>' +
    '<style>' +
    ':root{--bg-dark:#0a0a0f;--bg-card:rgba(20,20,30,0.8);--bg-glass:rgba(255,255,255,0.03);--neon-purple:#a855f7;--neon-cyan:#22d3ee;--neon-pink:#ec4899;--neon-green:#10b981;--neon-orange:#f97316;--neon-red:#ef4444;--text-primary:#f1f5f9;--text-secondary:#94a3b8;--text-muted:#64748b;--border-glow:rgba(168,85,247,0.3)}' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:"SF Pro Display",-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;background:var(--bg-dark);color:var(--text-primary);min-height:100vh;overflow-x:hidden}' +
    'body::before{content:"";position:fixed;top:0;left:0;right:0;bottom:0;background:radial-gradient(ellipse at top,rgba(168,85,247,0.15) 0%,transparent 50%),radial-gradient(ellipse at bottom right,rgba(34,211,238,0.1) 0%,transparent 50%);pointer-events:none;z-index:-1}' +

    // Header - Sleek gradient with glow
    '.header{background:linear-gradient(180deg,rgba(168,85,247,0.2) 0%,transparent 100%);padding:24px 20px 20px;text-align:center;position:relative;border-bottom:1px solid var(--border-glow)}' +
    '.header::after{content:"";position:absolute;bottom:0;left:50%;transform:translateX(-50%);width:60%;height:1px;background:linear-gradient(90deg,transparent,var(--neon-purple),transparent)}' +
    '.header h1{font-size:28px;font-weight:700;background:linear-gradient(135deg,#fff 0%,var(--neon-purple) 50%,var(--neon-cyan) 100%);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;letter-spacing:-0.5px}' +
    '.header .subtitle{font-size:12px;color:var(--text-secondary);margin-top:4px;text-transform:uppercase;letter-spacing:2px}' +
    '.refresh-btn{position:absolute;right:16px;top:24px;background:var(--bg-glass);backdrop-filter:blur(10px);border:1px solid var(--border-glow);color:var(--neon-purple);width:44px;height:44px;border-radius:12px;font-size:18px;cursor:pointer;transition:all 0.3s}' +
    '.refresh-btn:active{transform:scale(0.95);box-shadow:0 0 20px var(--neon-purple)}' +

    // Tab Navigation - Futuristic pills
    '.tabs{display:flex;gap:6px;padding:16px;overflow-x:auto;scrollbar-width:none;-ms-overflow-style:none;background:rgba(0,0,0,0.3)}' +
    '.tabs::-webkit-scrollbar{display:none}' +
    '.tab{flex-shrink:0;padding:10px 16px;border-radius:20px;font-size:12px;font-weight:600;color:var(--text-secondary);background:var(--bg-glass);border:1px solid transparent;cursor:pointer;transition:all 0.3s;white-space:nowrap}' +
    '.tab.active{background:linear-gradient(135deg,var(--neon-purple),var(--neon-cyan));color:white;border-color:transparent;box-shadow:0 0 20px rgba(168,85,247,0.4)}' +
    '.tab:not(.active):hover{border-color:var(--border-glow);color:var(--text-primary)}' +

    // Main Container
    '.main{padding:0 16px 100px}' +

    // Tab Content
    '.tab-content{display:none;animation:fadeSlide 0.4s ease}' +
    '.tab-content.active{display:block}' +
    '@keyframes fadeSlide{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}' +

    // Stats Grid - Glass cards with glow
    '.stats-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:20px}' +
    '.stat-card{background:var(--bg-card);backdrop-filter:blur(20px);border:1px solid var(--border-glow);border-radius:16px;padding:16px 12px;text-align:center;cursor:pointer;transition:all 0.3s;text-decoration:none;display:block;position:relative;overflow:hidden}' +
    '.stat-card::before{content:"";position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,transparent,var(--neon-purple),transparent);opacity:0;transition:opacity 0.3s}' +
    '.stat-card:hover::before,.stat-card:active::before{opacity:1}' +
    '.stat-card:active{transform:scale(0.97);box-shadow:0 0 30px rgba(168,85,247,0.2)}' +
    '.stat-value{font-size:28px;font-weight:700;color:var(--neon-purple);text-shadow:0 0 20px rgba(168,85,247,0.5)}' +
    '.stat-value.cyan{color:var(--neon-cyan);text-shadow:0 0 20px rgba(34,211,238,0.5)}' +
    '.stat-value.green{color:var(--neon-green);text-shadow:0 0 20px rgba(16,185,129,0.5)}' +
    '.stat-value.orange{color:var(--neon-orange);text-shadow:0 0 20px rgba(249,115,22,0.5)}' +
    '.stat-value.red{color:var(--neon-red);text-shadow:0 0 20px rgba(239,68,68,0.5)}' +
    '.stat-label{font-size:10px;color:var(--text-muted);text-transform:uppercase;letter-spacing:1px;margin-top:6px}' +

    // Section titles
    '.section-title{font-size:14px;font-weight:600;color:var(--text-secondary);margin:24px 0 12px;text-transform:uppercase;letter-spacing:1px;display:flex;align-items:center;gap:8px}' +
    '.section-title::before{content:"";width:3px;height:14px;background:linear-gradient(180deg,var(--neon-purple),var(--neon-cyan));border-radius:2px}' +

    // Overdue Alert - Danger glow
    '.alert-card{background:linear-gradient(135deg,rgba(239,68,68,0.1),rgba(239,68,68,0.05));border:1px solid rgba(239,68,68,0.3);border-radius:16px;padding:16px;margin-bottom:20px;position:relative;overflow:hidden}' +
    '.alert-card::before{content:"";position:absolute;top:0;left:0;right:0;height:2px;background:var(--neon-red);box-shadow:0 0 10px var(--neon-red)}' +
    '.alert-title{font-size:13px;font-weight:600;color:var(--neon-red);margin-bottom:12px;display:flex;align-items:center;gap:8px}' +
    '.alert-item{background:rgba(0,0,0,0.3);border-radius:10px;padding:12px;margin-bottom:8px;border-left:3px solid var(--neon-red)}' +
    '.alert-item:last-child{margin-bottom:0}' +
    '.alert-item-id{font-weight:600;color:var(--neon-purple);font-size:12px}' +
    '.alert-item-name{color:var(--text-primary);font-size:13px;margin-top:2px}' +
    '.alert-item-detail{color:var(--text-muted);font-size:11px;margin-top:2px}' +
    '.alert-btn{width:100%;padding:12px;margin-top:12px;background:var(--neon-red);border:none;border-radius:10px;color:white;font-weight:600;font-size:13px;cursor:pointer;transition:all 0.3s}' +
    '.alert-btn:active{transform:scale(0.98);box-shadow:0 0 20px var(--neon-red)}' +

    // Action Cards - Glass with icons
    '.action-grid{display:flex;flex-direction:column;gap:10px}' +
    '.action-card{background:var(--bg-card);backdrop-filter:blur(20px);border:1px solid var(--border-glow);border-radius:16px;padding:16px;display:flex;align-items:center;gap:14px;cursor:pointer;transition:all 0.3s;text-decoration:none}' +
    '.action-card:active{transform:scale(0.98);border-color:var(--neon-purple);box-shadow:0 0 20px rgba(168,85,247,0.2)}' +
    '.action-icon{width:48px;height:48px;border-radius:14px;background:linear-gradient(135deg,rgba(168,85,247,0.2),rgba(34,211,238,0.1));display:flex;align-items:center;justify-content:center;font-size:22px;flex-shrink:0}' +
    '.action-content{flex:1}' +
    '.action-label{font-weight:600;color:var(--text-primary);font-size:15px}' +
    '.action-desc{font-size:12px;color:var(--text-muted);margin-top:2px}' +
    '.action-arrow{color:var(--text-muted);font-size:18px}' +

    // List Items - Expandable glass cards
    '.list-container{display:flex;flex-direction:column;gap:10px}' +
    '.list-item{background:var(--bg-card);backdrop-filter:blur(20px);border:1px solid var(--border-glow);border-radius:14px;padding:14px;cursor:pointer;transition:all 0.3s}' +
    '.list-item:active{border-color:var(--neon-purple)}' +
    '.list-item.expanded{border-color:var(--neon-purple);box-shadow:0 0 20px rgba(168,85,247,0.1)}' +
    '.list-header{display:flex;justify-content:space-between;align-items:flex-start;gap:10px}' +
    '.list-main{flex:1}' +
    '.list-title{font-weight:600;color:var(--text-primary);font-size:14px}' +
    '.list-subtitle{font-size:12px;color:var(--text-muted);margin-top:2px}' +
    '.list-details{display:none;margin-top:14px;padding-top:14px;border-top:1px solid var(--border-glow)}' +
    '.list-item.expanded .list-details{display:block}' +
    '.detail-row{display:flex;justify-content:space-between;padding:6px 0;font-size:12px}' +
    '.detail-label{color:var(--text-muted)}' +
    '.detail-value{color:var(--text-primary);font-weight:500}' +

    // Status Badges
    '.badge{display:inline-flex;align-items:center;padding:4px 10px;border-radius:20px;font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:0.5px}' +
    '.badge-open{background:rgba(239,68,68,0.15);color:var(--neon-red);border:1px solid rgba(239,68,68,0.3)}' +
    '.badge-pending{background:rgba(249,115,22,0.15);color:var(--neon-orange);border:1px solid rgba(249,115,22,0.3)}' +
    '.badge-closed{background:rgba(16,185,129,0.15);color:var(--neon-green);border:1px solid rgba(16,185,129,0.3)}' +
    '.badge-overdue{background:rgba(239,68,68,0.25);color:#fca5a5;border:1px solid var(--neon-red);animation:pulse 2s infinite}' +
    '.badge-steward{background:rgba(168,85,247,0.15);color:var(--neon-purple);border:1px solid rgba(168,85,247,0.3)}' +
    '@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.7}}' +

    // Search Input
    '.search-box{position:relative;margin-bottom:16px}' +
    '.search-input{width:100%;padding:14px 14px 14px 44px;background:var(--bg-card);border:1px solid var(--border-glow);border-radius:14px;color:var(--text-primary);font-size:14px;transition:all 0.3s}' +
    '.search-input::placeholder{color:var(--text-muted)}' +
    '.search-input:focus{outline:none;border-color:var(--neon-purple);box-shadow:0 0 20px rgba(168,85,247,0.2)}' +
    '.search-icon{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:var(--text-muted);font-size:16px}' +

    // Filter Pills
    '.filter-pills{display:flex;gap:8px;margin-bottom:16px;overflow-x:auto;padding-bottom:4px;scrollbar-width:none}' +
    '.filter-pills::-webkit-scrollbar{display:none}' +
    '.filter-pill{flex-shrink:0;padding:8px 14px;border-radius:20px;font-size:11px;font-weight:600;color:var(--text-secondary);background:var(--bg-glass);border:1px solid var(--border-glow);cursor:pointer;transition:all 0.3s}' +
    '.filter-pill.active{background:var(--neon-purple);color:white;border-color:var(--neon-purple);box-shadow:0 0 15px rgba(168,85,247,0.4)}' +

    // Charts - Minimal bar charts
    '.chart-container{background:var(--bg-card);border:1px solid var(--border-glow);border-radius:16px;padding:16px;margin-bottom:16px}' +
    '.chart-title{font-size:13px;font-weight:600;color:var(--text-primary);margin-bottom:16px}' +
    '.chart-bar{display:flex;align-items:center;gap:10px;margin-bottom:10px}' +
    '.chart-bar:last-child{margin-bottom:0}' +
    '.chart-label{font-size:11px;color:var(--text-muted);width:80px;flex-shrink:0}' +
    '.chart-track{flex:1;height:8px;background:var(--bg-glass);border-radius:4px;overflow:hidden}' +
    '.chart-fill{height:100%;border-radius:4px;background:linear-gradient(90deg,var(--neon-purple),var(--neon-cyan));transition:width 0.5s ease}' +
    '.chart-value{font-size:11px;color:var(--text-primary);font-weight:600;width:30px;text-align:right}' +

    // Analytics Cards
    '.analytics-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:12px;margin-bottom:16px}' +
    '.analytics-card{background:var(--bg-card);border:1px solid var(--border-glow);border-radius:14px;padding:14px;text-align:center}' +
    '.analytics-value{font-size:24px;font-weight:700;color:var(--neon-cyan);text-shadow:0 0 15px rgba(34,211,238,0.5)}' +
    '.analytics-label{font-size:10px;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.5px;margin-top:4px}' +

    // Bottom Navigation - Glass effect
    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:rgba(10,10,15,0.9);backdrop-filter:blur(20px);border-top:1px solid var(--border-glow);display:flex;justify-content:space-around;padding:10px 0 max(10px,env(safe-area-inset-bottom));z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:6px 12px;text-decoration:none;color:var(--text-muted);font-size:9px;font-weight:500;transition:all 0.3s;position:relative}' +
    '.nav-item.active{color:var(--neon-purple)}' +
    '.nav-item.active::before{content:"";position:absolute;top:-10px;left:50%;transform:translateX(-50%);width:20px;height:3px;background:var(--neon-purple);border-radius:2px;box-shadow:0 0 10px var(--neon-purple)}' +
    '.nav-icon{font-size:20px;margin-bottom:4px}' +

    // Loading States
    '.loading{text-align:center;padding:40px;color:var(--text-muted)}' +
    '.spinner{width:24px;height:24px;border:2px solid var(--border-glow);border-top-color:var(--neon-purple);border-radius:50%;animation:spin 0.8s linear infinite;margin:0 auto 12px}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +

    // Success State
    '.success-msg{text-align:center;padding:20px;color:var(--neon-green);font-size:13px}' +

    // Empty State
    '.empty-state{text-align:center;padding:40px 20px;color:var(--text-muted)}' +
    '.empty-icon{font-size:48px;margin-bottom:12px;opacity:0.5}' +
    '.empty-text{font-size:13px}' +

    '</style></head><body>' +

    // Header
    '<div class="header">' +
    '<button class="refresh-btn" onclick="location.reload()">↻</button>' +
    '<h1>509 DASHBOARD</h1>' +
    '<div class="subtitle">Grievance Command Center</div>' +
    '</div>' +

    // Tab Navigation
    '<div class="tabs">' +
    '<button class="tab active" onclick="switchTab(\'overview\')">📊 Overview</button>' +
    '<button class="tab" onclick="switchTab(\'mycases\')">👤 My Cases</button>' +
    '<button class="tab" onclick="switchTab(\'grievances\')">📋 Cases</button>' +
    '<button class="tab" onclick="switchTab(\'members\')">👥 Members</button>' +
    '<button class="tab" onclick="switchTab(\'analytics\')">📈 Analytics</button>' +
    '<button class="tab" onclick="switchTab(\'links\')">🔗 Links</button>' +
    '</div>' +

    '<div class="main">' +

    // ===== OVERVIEW TAB =====
    '<div id="tab-overview" class="tab-content active">' +
    '<div class="stats-grid">' +
    '<a class="stat-card" href="' + baseUrl + '?page=members"><div class="stat-value cyan">' + stats.totalMembers + '</div><div class="stat-label">Members</div></a>' +
    '<a class="stat-card" href="' + baseUrl + '?page=grievances"><div class="stat-value">' + stats.totalGrievances + '</div><div class="stat-label">Total Cases</div></a>' +
    '<a class="stat-card" href="' + baseUrl + '?page=grievances&filter=open"><div class="stat-value green">' + stats.activeGrievances + '</div><div class="stat-label">Active</div></a>' +
    '<a class="stat-card" href="' + baseUrl + '?page=grievances&filter=pending"><div class="stat-value orange">' + stats.pendingGrievances + '</div><div class="stat-label">Pending</div></a>' +
    '<a class="stat-card" href="' + baseUrl + '?page=grievances&filter=overdue"><div class="stat-value red">' + stats.overdueGrievances + '</div><div class="stat-label">Overdue</div></a>' +
    '<div class="stat-card"><div class="stat-value green">' + stats.winRate + '</div><div class="stat-label">Win Rate</div></div>' +
    '</div>' +
    '<div id="overdue-preview"><div class="loading"><div class="spinner"></div><div>Loading alerts...</div></div></div>' +
    '<div class="section-title">Quick Actions</div>' +
    '<div class="action-grid">' +
    '<a class="action-card" href="' + baseUrl + '?page=search"><div class="action-icon">🔍</div><div class="action-content"><div class="action-label">Search</div><div class="action-desc">Find members or grievances</div></div><div class="action-arrow">›</div></a>' +
    '<a class="action-card" href="' + baseUrl + '?page=grievances"><div class="action-icon">📋</div><div class="action-content"><div class="action-label">All Cases</div><div class="action-desc">Browse grievance log</div></div><div class="action-arrow">›</div></a>' +
    '<a class="action-card" href="' + baseUrl + '?page=members"><div class="action-icon">👥</div><div class="action-content"><div class="action-label">Members</div><div class="action-desc">View directory</div></div><div class="action-arrow">›</div></a>' +
    '</div>' +
    '</div>' +

    // ===== MY CASES TAB =====
    '<div id="tab-mycases" class="tab-content">' +
    '<div id="mycases-stats" class="stats-grid"></div>' +
    '<div class="filter-pills" id="mycases-filters">' +
    '<button class="filter-pill active" onclick="filterMyCases(\'all\')">All</button>' +
    '<button class="filter-pill" onclick="filterMyCases(\'open\')">Open</button>' +
    '<button class="filter-pill" onclick="filterMyCases(\'pending\')">Pending</button>' +
    '<button class="filter-pill" onclick="filterMyCases(\'overdue\')">Overdue</button>' +
    '</div>' +
    '<div id="mycases-list" class="list-container"><div class="loading"><div class="spinner"></div><div>Loading your cases...</div></div></div>' +
    '</div>' +

    // ===== GRIEVANCES TAB =====
    '<div id="tab-grievances" class="tab-content">' +
    '<div class="search-box"><span class="search-icon">🔍</span><input type="text" class="search-input" placeholder="Search cases..." id="grievance-search" oninput="filterGrievances()"></div>' +
    '<div class="filter-pills" id="grievance-filters">' +
    '<button class="filter-pill active" onclick="setGrievanceFilter(\'all\')">All</button>' +
    '<button class="filter-pill" onclick="setGrievanceFilter(\'open\')">Open</button>' +
    '<button class="filter-pill" onclick="setGrievanceFilter(\'pending\')">Pending</button>' +
    '<button class="filter-pill" onclick="setGrievanceFilter(\'overdue\')">Overdue</button>' +
    '<button class="filter-pill" onclick="setGrievanceFilter(\'closed\')">Closed</button>' +
    '</div>' +
    '<div id="grievance-list" class="list-container"><div class="loading"><div class="spinner"></div><div>Loading cases...</div></div></div>' +
    '</div>' +

    // ===== MEMBERS TAB =====
    '<div id="tab-members" class="tab-content">' +
    '<div class="search-box"><span class="search-icon">🔍</span><input type="text" class="search-input" placeholder="Search members..." id="member-search" oninput="filterMembers()"></div>' +
    '<div id="member-list" class="list-container"><div class="loading"><div class="spinner"></div><div>Loading members...</div></div></div>' +
    '</div>' +

    // ===== ANALYTICS TAB =====
    '<div id="tab-analytics" class="tab-content">' +
    '<div id="analytics-content"><div class="loading"><div class="spinner"></div><div>Loading analytics...</div></div></div>' +
    '</div>' +

    // ===== LINKS TAB =====
    '<div id="tab-links" class="tab-content">' +
    '<div class="section-title">Forms & Resources</div>' +
    '<div id="links-content" class="action-grid"><div class="loading"><div class="spinner"></div><div>Loading links...</div></div></div>' +
    '</div>' +

    '</div>' +

    // Bottom Navigation
    '<nav class="bottom-nav">' +
    '<a class="nav-item active" href="' + baseUrl + '"><span class="nav-icon">📊</span>Home</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=search"><span class="nav-icon">🔍</span>Search</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=grievances"><span class="nav-icon">📋</span>Cases</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=members"><span class="nav-icon">👥</span>Members</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=links"><span class="nav-icon">🔗</span>Links</a>' +
    '</nav>' +

    // JavaScript
    '<script>' +
    'var baseUrl="' + baseUrl + '";' +
    'var allGrievances=[],allMembers=[],myCases=[];' +
    'var currentGrievanceFilter="all",currentMyCasesFilter="all";' +

    // Tab switching
    'function switchTab(tab){' +
    '  document.querySelectorAll(".tab").forEach(function(t){t.classList.remove("active")});' +
    '  document.querySelectorAll(".tab-content").forEach(function(c){c.classList.remove("active")});' +
    '  event.target.classList.add("active");' +
    '  document.getElementById("tab-"+tab).classList.add("active");' +
    '  if(tab==="mycases"&&myCases.length===0)loadMyCases();' +
    '  if(tab==="grievances"&&allGrievances.length===0)loadGrievances();' +
    '  if(tab==="members"&&allMembers.length===0)loadMembers();' +
    '  if(tab==="analytics")loadAnalytics();' +
    '  if(tab==="links")loadLinks();' +
    '}' +

    // Load overdue preview
    'function loadOverdue(){' +
    '  google.script.run.withSuccessHandler(function(data){' +
    '    if(!data||!Array.isArray(data)){document.getElementById("overdue-preview").innerHTML="<div class=\\"success-msg\\">✅ All cases on track!</div>";return}' +
    '    var overdue=data.filter(function(g){return g&&g.isOverdue});' +
    '    if(overdue.length===0){document.getElementById("overdue-preview").innerHTML="<div class=\\"success-msg\\">✅ No overdue cases</div>";return}' +
    '    var html="<div class=\\"alert-card\\"><div class=\\"alert-title\\">⚠️ OVERDUE ALERTS ("+overdue.length+")</div>";' +
    '    overdue.slice(0,3).forEach(function(g){' +
    '      html+="<div class=\\"alert-item\\"><div class=\\"alert-item-id\\">"+g.id+"</div><div class=\\"alert-item-name\\">"+g.name+"</div><div class=\\"alert-item-detail\\">"+g.category+" • "+g.step+"</div></div>";' +
    '    });' +
    '    if(overdue.length>3)html+="<button class=\\"alert-btn\\" onclick=\\"location.href=baseUrl+\'?page=grievances&filter=overdue\'\\">View All "+overdue.length+" Overdue</button>";' +
    '    html+="</div>";' +
    '    document.getElementById("overdue-preview").innerHTML=html;' +
    '  }).withFailureHandler(function(){' +
    '    document.getElementById("overdue-preview").innerHTML="";' +
    '  }).getWebAppGrievanceList();' +
    '}' +

    // Load My Cases
    'function loadMyCases(){' +
    '  google.script.run.withSuccessHandler(function(data){' +
    '    myCases=data||[];' +
    '    var open=myCases.filter(function(c){return c.status==="Open"}).length;' +
    '    var pending=myCases.filter(function(c){return c.status==="Pending Info"}).length;' +
    '    var overdue=myCases.filter(function(c){return c.isOverdue}).length;' +
    '    document.getElementById("mycases-stats").innerHTML=' +
    '      "<div class=\\"stat-card\\"><div class=\\"stat-value cyan\\">"+myCases.length+"</div><div class=\\"stat-label\\">Assigned</div></div>"' +
    '      +"<div class=\\"stat-card\\"><div class=\\"stat-value green\\">"+open+"</div><div class=\\"stat-label\\">Open</div></div>"' +
    '      +"<div class=\\"stat-card\\"><div class=\\"stat-value orange\\">"+pending+"</div><div class=\\"stat-label\\">Pending</div></div>"' +
    '      +"<div class=\\"stat-card\\"><div class=\\"stat-value red\\">"+overdue+"</div><div class=\\"stat-label\\">Overdue</div></div>";' +
    '    renderMyCases();' +
    '  }).withFailureHandler(function(){' +
    '    document.getElementById("mycases-list").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">📋</div><div class=\\"empty-text\\">Could not load cases</div></div>";' +
    '  }).getMyStewardCases();' +
    '}' +

    'function filterMyCases(filter){' +
    '  currentMyCasesFilter=filter;' +
    '  document.querySelectorAll("#mycases-filters .filter-pill").forEach(function(p){p.classList.remove("active")});' +
    '  event.target.classList.add("active");' +
    '  renderMyCases();' +
    '}' +

    'function renderMyCases(){' +
    '  var filtered=myCases;' +
    '  if(currentMyCasesFilter==="open")filtered=myCases.filter(function(c){return c.status==="Open"});' +
    '  if(currentMyCasesFilter==="pending")filtered=myCases.filter(function(c){return c.status==="Pending Info"});' +
    '  if(currentMyCasesFilter==="overdue")filtered=myCases.filter(function(c){return c.isOverdue});' +
    '  if(filtered.length===0){document.getElementById("mycases-list").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">📋</div><div class=\\"empty-text\\">No cases found</div></div>";return}' +
    '  var html="";' +
    '  filtered.forEach(function(c,i){' +
    '    var badge=c.isOverdue?"badge-overdue":c.status==="Open"?"badge-open":c.status==="Pending Info"?"badge-pending":"badge-closed";' +
    '    html+="<div class=\\"list-item\\" onclick=\\"this.classList.toggle(\'expanded\')\\">"' +
    '      +"<div class=\\"list-header\\"><div class=\\"list-main\\"><div class=\\"list-title\\">"+c.id+" - "+c.memberName+"</div><div class=\\"list-subtitle\\">"+c.category+"</div></div>"' +
    '      +"<span class=\\"badge "+badge+"\\">"+c.status+"</span></div>"' +
    '      +"<div class=\\"list-details\\"><div class=\\"detail-row\\"><span class=\\"detail-label\\">Filed</span><span class=\\"detail-value\\">"+(c.filedDate||"-")+"</span></div>"' +
    '      +"<div class=\\"detail-row\\"><span class=\\"detail-label\\">Step</span><span class=\\"detail-value\\">"+(c.step||"-")+"</span></div>"' +
    '      +"<div class=\\"detail-row\\"><span class=\\"detail-label\\">Location</span><span class=\\"detail-value\\">"+(c.location||"-")+"</span></div></div></div>";' +
    '  });' +
    '  document.getElementById("mycases-list").innerHTML=html;' +
    '}' +

    // Load Grievances
    'function loadGrievances(){' +
    '  google.script.run.withSuccessHandler(function(data){' +
    '    allGrievances=data||[];' +
    '    renderGrievances();' +
    '  }).withFailureHandler(function(){' +
    '    document.getElementById("grievance-list").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">📋</div><div class=\\"empty-text\\">Could not load cases</div></div>";' +
    '  }).getWebAppGrievanceList();' +
    '}' +

    'function setGrievanceFilter(filter){' +
    '  currentGrievanceFilter=filter;' +
    '  document.querySelectorAll("#grievance-filters .filter-pill").forEach(function(p){p.classList.remove("active")});' +
    '  event.target.classList.add("active");' +
    '  renderGrievances();' +
    '}' +

    'function filterGrievances(){renderGrievances()}' +

    'function renderGrievances(){' +
    '  var search=(document.getElementById("grievance-search").value||"").toLowerCase();' +
    '  var filtered=allGrievances.filter(function(g){' +
    '    if(search&&g.id.toLowerCase().indexOf(search)<0&&g.name.toLowerCase().indexOf(search)<0)return false;' +
    '    if(currentGrievanceFilter==="open")return g.status==="Open";' +
    '    if(currentGrievanceFilter==="pending")return g.status==="Pending Info";' +
    '    if(currentGrievanceFilter==="overdue")return g.isOverdue;' +
    '    if(currentGrievanceFilter==="closed")return g.status==="Closed"||g.status==="Resolved";' +
    '    return true;' +
    '  });' +
    '  if(filtered.length===0){document.getElementById("grievance-list").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">📋</div><div class=\\"empty-text\\">No cases found</div></div>";return}' +
    '  var html="";' +
    '  filtered.slice(0,50).forEach(function(g){' +
    '    var badge=g.isOverdue?"badge-overdue":g.status==="Open"?"badge-open":g.status==="Pending Info"?"badge-pending":"badge-closed";' +
    '    html+="<div class=\\"list-item\\" onclick=\\"this.classList.toggle(\'expanded\')\\">"' +
    '      +"<div class=\\"list-header\\"><div class=\\"list-main\\"><div class=\\"list-title\\">"+g.id+" - "+g.name+"</div><div class=\\"list-subtitle\\">"+g.category+"</div></div>"' +
    '      +"<span class=\\"badge "+badge+"\\">"+g.status+"</span></div>"' +
    '      +"<div class=\\"list-details\\"><div class=\\"detail-row\\"><span class=\\"detail-label\\">Filed</span><span class=\\"detail-value\\">"+(g.filed||"-")+"</span></div>"' +
    '      +"<div class=\\"detail-row\\"><span class=\\"detail-label\\">Step</span><span class=\\"detail-value\\">"+(g.step||"-")+"</span></div>"' +
    '      +"<div class=\\"detail-row\\"><span class=\\"detail-label\\">Steward</span><span class=\\"detail-value\\">"+(g.steward||"-")+"</span></div>"' +
    '      +"<div class=\\"detail-row\\"><span class=\\"detail-label\\">Location</span><span class=\\"detail-value\\">"+(g.location||"-")+"</span></div></div></div>";' +
    '  });' +
    '  document.getElementById("grievance-list").innerHTML=html;' +
    '}' +

    // Load Members
    'function loadMembers(){' +
    '  google.script.run.withSuccessHandler(function(data){' +
    '    allMembers=data||[];' +
    '    renderMembers();' +
    '  }).withFailureHandler(function(){' +
    '    document.getElementById("member-list").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">👥</div><div class=\\"empty-text\\">Could not load members</div></div>";' +
    '  }).getWebAppMemberList();' +
    '}' +

    'function filterMembers(){renderMembers()}' +

    'function renderMembers(){' +
    '  var search=(document.getElementById("member-search").value||"").toLowerCase();' +
    '  var filtered=allMembers.filter(function(m){' +
    '    if(!search)return true;' +
    '    return m.name.toLowerCase().indexOf(search)>=0||m.id.toLowerCase().indexOf(search)>=0||(m.title||"").toLowerCase().indexOf(search)>=0;' +
    '  });' +
    '  if(filtered.length===0){document.getElementById("member-list").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">👥</div><div class=\\"empty-text\\">No members found</div></div>";return}' +
    '  var html="";' +
    '  filtered.slice(0,50).forEach(function(m){' +
    '    html+="<div class=\\"list-item\\" onclick=\\"this.classList.toggle(\'expanded\')\\">"' +
    '      +"<div class=\\"list-header\\"><div class=\\"list-main\\"><div class=\\"list-title\\">"+m.name+"</div><div class=\\"list-subtitle\\">"+m.id+" • "+(m.title||"")+"</div></div>"' +
    '      +(m.isSteward?"<span class=\\"badge badge-steward\\">Steward</span>":"")+"</div>"' +
    '      +"<div class=\\"list-details\\"><div class=\\"detail-row\\"><span class=\\"detail-label\\">Location</span><span class=\\"detail-value\\">"+(m.location||"-")+"</span></div>"' +
    '      +"<div class=\\"detail-row\\"><span class=\\"detail-label\\">Unit</span><span class=\\"detail-value\\">"+(m.unit||"-")+"</span></div>"' +
    '      +"<div class=\\"detail-row\\"><span class=\\"detail-label\\">Email</span><span class=\\"detail-value\\">"+(m.email||"-")+"</span></div>"' +
    '      +"<div class=\\"detail-row\\"><span class=\\"detail-label\\">Phone</span><span class=\\"detail-value\\">"+(m.phone||"-")+"</span></div></div></div>";' +
    '  });' +
    '  document.getElementById("member-list").innerHTML=html;' +
    '}' +

    // Load Analytics
    'function loadAnalytics(){' +
    '  google.script.run.withSuccessHandler(function(data){' +
    '    if(!data){document.getElementById("analytics-content").innerHTML="<div class=\\"empty-state\\">No data available</div>";return}' +
    '    var html="<div class=\\"analytics-grid\\">"' +
    '      +"<div class=\\"analytics-card\\"><div class=\\"analytics-value\\">"+data.totalMembers+"</div><div class=\\"analytics-label\\">Total Members</div></div>"' +
    '      +"<div class=\\"analytics-card\\"><div class=\\"analytics-value\\">"+data.stewardCount+"</div><div class=\\"analytics-label\\">Stewards</div></div>"' +
    '      +"<div class=\\"analytics-card\\"><div class=\\"analytics-value\\">"+(data.memberToStewardRatio||"N/A")+"</div><div class=\\"analytics-label\\">Member:Steward</div></div>"' +
    '      +"<div class=\\"analytics-card\\"><div class=\\"analytics-value\\">"+data.membersWithCases+"</div><div class=\\"analytics-label\\">With Cases</div></div></div>";' +
    '    html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">Case Status Distribution</div>";' +
    '    var statusData=data.statusCounts||{};' +
    '    var maxStatus=Math.max(statusData.Open||0,statusData.Pending||0,statusData.Closed||0,1);' +
    '    html+="<div class=\\"chart-bar\\"><span class=\\"chart-label\\">Open</span><div class=\\"chart-track\\"><div class=\\"chart-fill\\" style=\\"width:"+(((statusData.Open||0)/maxStatus)*100)+"%\\"></div></div><span class=\\"chart-value\\">"+(statusData.Open||0)+"</span></div>";' +
    '    html+="<div class=\\"chart-bar\\"><span class=\\"chart-label\\">Pending</span><div class=\\"chart-track\\"><div class=\\"chart-fill\\" style=\\"width:"+(((statusData.Pending||0)/maxStatus)*100)+"%\\"></div></div><span class=\\"chart-value\\">"+(statusData.Pending||0)+"</span></div>";' +
    '    html+="<div class=\\"chart-bar\\"><span class=\\"chart-label\\">Closed</span><div class=\\"chart-track\\"><div class=\\"chart-fill\\" style=\\"width:"+(((statusData.Closed||0)/maxStatus)*100)+"%\\"></div></div><span class=\\"chart-value\\">"+(statusData.Closed||0)+"</span></div>";' +
    '    html+="</div>";' +
    '    if(data.topCategories&&data.topCategories.length>0){' +
    '      html+="<div class=\\"chart-container\\"><div class=\\"chart-title\\">Top Issue Categories</div>";' +
    '      var maxCat=data.topCategories[0].count||1;' +
    '      data.topCategories.slice(0,5).forEach(function(c){' +
    '        html+="<div class=\\"chart-bar\\"><span class=\\"chart-label\\">"+c.name.substring(0,12)+"</span><div class=\\"chart-track\\"><div class=\\"chart-fill\\" style=\\"width:"+((c.count/maxCat)*100)+"%\\"></div></div><span class=\\"chart-value\\">"+c.count+"</span></div>";' +
    '      });' +
    '      html+="</div>";' +
    '    }' +
    '    document.getElementById("analytics-content").innerHTML=html;' +
    '  }).withFailureHandler(function(){' +
    '    document.getElementById("analytics-content").innerHTML="<div class=\\"empty-state\\">Could not load analytics</div>";' +
    '  }).getInteractiveAnalyticsData();' +
    '}' +

    // Load Links
    'function loadLinks(){' +
    '  google.script.run.withSuccessHandler(function(data){' +
    '    if(!data){document.getElementById("links-content").innerHTML="<div class=\\"empty-state\\">No links configured</div>";return}' +
    '    var html="";' +
    '    if(data.grievanceFormUrl)html+="<a class=\\"action-card\\" href=\\""+data.grievanceFormUrl+"\\" target=\\"_blank\\"><div class=\\"action-icon\\">📝</div><div class=\\"action-content\\"><div class=\\"action-label\\">Grievance Form</div><div class=\\"action-desc\\">Submit new grievance</div></div><div class=\\"action-arrow\\">›</div></a>";' +
    '    if(data.contactFormUrl)html+="<a class=\\"action-card\\" href=\\""+data.contactFormUrl+"\\" target=\\"_blank\\"><div class=\\"action-icon\\">📧</div><div class=\\"action-content\\"><div class=\\"action-label\\">Contact Form</div><div class=\\"action-desc\\">Update member info</div></div><div class=\\"action-arrow\\">›</div></a>";' +
    '    if(data.surveyUrl)html+="<a class=\\"action-card\\" href=\\""+data.surveyUrl+"\\" target=\\"_blank\\"><div class=\\"action-icon\\">📊</div><div class=\\"action-content\\"><div class=\\"action-label\\">Satisfaction Survey</div><div class=\\"action-desc\\">Member feedback</div></div><div class=\\"action-arrow\\">›</div></a>";' +
    '    if(data.spreadsheetUrl)html+="<a class=\\"action-card\\" href=\\""+data.spreadsheetUrl+"\\" target=\\"_blank\\"><div class=\\"action-icon\\">📁</div><div class=\\"action-content\\"><div class=\\"action-label\\">Full Spreadsheet</div><div class=\\"action-desc\\">View all data</div></div><div class=\\"action-arrow\\">›</div></a>";' +
    '    if(data.orgWebsite)html+="<a class=\\"action-card\\" href=\\""+data.orgWebsite+"\\" target=\\"_blank\\"><div class=\\"action-icon\\">🌐</div><div class=\\"action-content\\"><div class=\\"action-label\\">Organization Website</div><div class=\\"action-desc\\">"+data.orgWebsite+"</div></div><div class=\\"action-arrow\\">›</div></a>";' +
    '    if(!html)html="<div class=\\"empty-state\\">No links configured</div>";' +
    '    document.getElementById("links-content").innerHTML=html;' +
    '  }).withFailureHandler(function(){' +
    '    document.getElementById("links-content").innerHTML="<div class=\\"empty-state\\">Could not load links</div>";' +
    '  }).getInteractiveResourceLinks();' +
    '}' +

    // Initialize
    'loadOverdue();' +
    '</script>' +

    '</body></html>';
}

/**
 * Returns the search page HTML for web app
 */
function getWebAppSearchHtml() {
  var baseUrl = ScriptApp.getService().getUrl();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<title>Search - 509 Dashboard</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh;padding-bottom:80px}' +

    // Header with search
    '.header{background:linear-gradient(135deg,#7C3AED,#5B21B6);color:white;padding:15px;position:sticky;top:0;z-index:100}' +
    '.header h2{font-size:clamp(18px,4vw,22px);margin-bottom:12px;text-align:center}' +
    '.search-container{position:relative}' +
    '.search-input{width:100%;padding:14px 14px 14px 45px;border:none;border-radius:12px;font-size:16px;background:white;-webkit-appearance:none}' +
    '.search-input:focus{outline:none;box-shadow:0 0 0 3px rgba(124,58,237,0.3)}' +
    '.search-icon{position:absolute;left:14px;top:50%;transform:translateY(-50%);font-size:20px;color:#666}' +
    '.clear-btn{position:absolute;right:14px;top:50%;transform:translateY(-50%);font-size:20px;color:#999;background:none;border:none;cursor:pointer;display:none}' +

    // Tabs
    '.tabs{display:flex;background:white;border-bottom:1px solid #e0e0e0;position:sticky;top:76px;z-index:99}' +
    '.tab{flex:1;padding:14px;text-align:center;font-size:14px;font-weight:500;color:#666;border:none;background:none;cursor:pointer;border-bottom:3px solid transparent;min-height:48px}' +
    '.tab.active{color:#7C3AED;border-bottom-color:#7C3AED}' +

    // Results
    '.results{padding:15px}' +
    '.result-card{background:white;padding:16px;border-radius:14px;box-shadow:0 2px 6px rgba(0,0,0,0.06);margin-bottom:12px}' +
    '.result-type{font-size:12px;color:#7C3AED;font-weight:600;text-transform:uppercase;margin-bottom:6px}' +
    '.result-title{font-size:17px;font-weight:600;color:#333;margin-bottom:4px}' +
    '.result-detail{font-size:14px;color:#666;margin-top:4px}' +
    '.result-badge{display:inline-block;padding:4px 10px;border-radius:20px;font-size:12px;font-weight:500;margin-top:8px}' +
    '.badge-open{background:#FEE2E2;color:#DC2626}' +
    '.badge-pending{background:#FEF3C7;color:#D97706}' +
    '.badge-resolved{background:#D1FAE5;color:#059669}' +

    // Empty state
    '.empty-state{text-align:center;padding:60px 20px;color:#999}' +
    '.empty-icon{font-size:48px;margin-bottom:15px}' +
    '.empty-text{font-size:16px}' +

    // Loading
    '.loading{text-align:center;padding:40px;color:#666}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:24px;height:24px;border:3px solid #e0e0e0;border-top-color:#7C3AED;border-radius:50%;animation:spin 0.8s linear infinite}' +

    // Bottom nav - 5 items
    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;display:flex;justify-content:space-around;padding:8px 0 max(8px,env(safe-area-inset-bottom));box-shadow:0 -2px 10px rgba(0,0,0,0.1);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:6px 10px;text-decoration:none;color:#666;font-size:10px;min-width:60px}' +
    '.nav-item.active{color:#7C3AED}' +
    '.nav-icon{font-size:22px;margin-bottom:3px}' +

    '</style></head><body>' +

    '<div class="header">' +
    '<h2>🔍 Search</h2>' +
    '<div class="search-container">' +
    '<span class="search-icon">🔍</span>' +
    '<input type="text" class="search-input" id="searchInput" placeholder="Search members or grievances..." oninput="handleSearch(this.value)" autocomplete="off" autocapitalize="off">' +
    '<button class="clear-btn" id="clearBtn" onclick="clearSearch()">✕</button>' +
    '</div></div>' +

    '<div class="tabs">' +
    '<button class="tab active" data-tab="all" onclick="setTab(\'all\',this)">All</button>' +
    '<button class="tab" data-tab="members" onclick="setTab(\'members\',this)">Members</button>' +
    '<button class="tab" data-tab="grievances" onclick="setTab(\'grievances\',this)">Grievances</button>' +
    '</div>' +

    '<div class="results" id="results">' +
    '<div class="empty-state"><div class="empty-icon">🔍</div><div class="empty-text">Type to search members or grievances</div></div>' +
    '</div>' +

    // Bottom Navigation - 5 items
    '<nav class="bottom-nav">' +
    '<a class="nav-item" href="' + baseUrl + '">' +
    '<span class="nav-icon">📊</span>Home</a>' +
    '<a class="nav-item active" href="' + baseUrl + '?page=search">' +
    '<span class="nav-icon">🔍</span>Search</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=grievances">' +
    '<span class="nav-icon">📋</span>Cases</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=members">' +
    '<span class="nav-icon">👥</span>Members</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=links">' +
    '<span class="nav-icon">🔗</span>Links</a>' +
    '</nav>' +

    '<script>' +
    'var currentTab="all";' +
    'var searchTimeout=null;' +
    'var lastQuery="";' +

    'function setTab(tab,btn){' +
    '  currentTab=tab;' +
    '  document.querySelectorAll(".tab").forEach(function(t){t.classList.remove("active")});' +
    '  btn.classList.add("active");' +
    '  if(lastQuery.length>=2)performSearch(lastQuery);' +
    '}' +

    'function handleSearch(q){' +
    '  lastQuery=q;' +
    '  document.getElementById("clearBtn").style.display=q?"block":"none";' +
    '  if(searchTimeout)clearTimeout(searchTimeout);' +
    '  if(!q||q.length<2){' +
    '    showEmpty("Type at least 2 characters to search");' +
    '    return;' +
    '  }' +
    '  showLoading();' +
    '  searchTimeout=setTimeout(function(){performSearch(q)},300);' +
    '}' +

    'function performSearch(q){' +
    '  google.script.run.withSuccessHandler(renderResults).withFailureHandler(showError).getWebAppSearchResults(q,currentTab);' +
    '}' +

    'function clearSearch(){' +
    '  document.getElementById("searchInput").value="";' +
    '  document.getElementById("clearBtn").style.display="none";' +
    '  lastQuery="";' +
    '  showEmpty("Type to search members or grievances");' +
    '}' +

    'function showEmpty(msg){' +
    '  document.getElementById("results").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">🔍</div><div class=\\"empty-text\\">"+msg+"</div></div>";' +
    '}' +

    'function showLoading(){' +
    '  document.getElementById("results").innerHTML="<div class=\\"loading\\"><div class=\\"spinner\\"></div><div style=\\"margin-top:15px\\">Searching...</div></div>";' +
    '}' +

    'function showError(err){' +
    '  document.getElementById("results").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">⚠️</div><div class=\\"empty-text\\">Error: "+(err.message||"Unknown error")+"</div></div>";' +
    '}' +

    'function getBadgeClass(status){' +
    '  if(!status)return"";' +
    '  var s=status.toLowerCase();' +
    '  if(s.indexOf("open")>=0)return"badge-open";' +
    '  if(s.indexOf("pending")>=0)return"badge-pending";' +
    '  if(s.indexOf("resolved")>=0||s.indexOf("closed")>=0||s.indexOf("withdrawn")>=0)return"badge-resolved";' +
    '  return"";' +
    '}' +

    'function renderResults(data){' +
    '  var c=document.getElementById("results");' +
    '  if(!data||data.length===0){' +
    '    showEmpty("No results found");' +
    '    return;' +
    '  }' +
    '  c.innerHTML=data.map(function(r){' +
    '    var badge=r.status?"<span class=\\"result-badge "+getBadgeClass(r.status)+"\\">"+r.status+"</span>":"";' +
    '    return"<div class=\\"result-card\\">"+"<div class=\\"result-type\\">"+(r.type==="member"?"👤 Member":"📋 Grievance")+"</div>"+"<div class=\\"result-title\\">"+r.title+"</div>"+"<div class=\\"result-detail\\">"+r.subtitle+"</div>"+(r.detail?"<div class=\\"result-detail\\">"+r.detail+"</div>":"")+badge+"</div>";' +
    '  }).join("");' +
    '}' +

    // Auto-focus search on load
    'document.getElementById("searchInput").focus();' +
    '</script>' +

    '</body></html>';
}

/**
 * Returns the grievance list HTML for web app (enhanced with Overdue filter, expandable details)
 */
function getWebAppGrievanceListHtml() {
  var baseUrl = ScriptApp.getService().getUrl();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<title>Grievances - 509 Dashboard</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh;padding-bottom:80px}' +

    // Header
    '.header{background:linear-gradient(135deg,#7C3AED,#5B21B6);color:white;padding:15px 15px 12px;position:sticky;top:0;z-index:100}' +
    '.header h2{font-size:clamp(18px,4vw,22px);text-align:center;margin-bottom:12px}' +

    // Filter pills with Overdue
    '.filters{display:flex;gap:8px;overflow-x:auto;padding:2px 0;-webkit-overflow-scrolling:touch}' +
    '.filter-pill{flex-shrink:0;padding:8px 16px;border-radius:20px;font-size:13px;font-weight:500;border:none;cursor:pointer;background:rgba(255,255,255,0.2);color:white}' +
    '.filter-pill.active{background:white;color:#7C3AED}' +
    '.filter-pill.danger{background:#DC2626;color:white}' +
    '.filter-pill.danger.active{background:#FEE2E2;color:#DC2626}' +

    // List with expandable cards
    '.grievance-list{padding:15px}' +
    '.grievance-card{background:white;padding:16px;border-radius:14px;box-shadow:0 2px 6px rgba(0,0,0,0.06);margin-bottom:12px;cursor:pointer;transition:all 0.2s}' +
    '.grievance-card:active{transform:scale(0.99)}' +
    '.grievance-card.overdue{border-left:4px solid #DC2626}' +
    '.grievance-header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:8px}' +
    '.grievance-id{font-size:15px;font-weight:700;color:#7C3AED}' +
    '.grievance-status{padding:4px 10px;border-radius:20px;font-size:11px;font-weight:600}' +
    '.status-open{background:#FEE2E2;color:#DC2626}' +
    '.status-pending{background:#FEF3C7;color:#D97706}' +
    '.status-resolved{background:#D1FAE5;color:#059669}' +
    '.status-overdue{background:#DC2626;color:white;animation:pulse 2s infinite}' +
    '@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.7}}' +
    '.grievance-name{font-size:16px;font-weight:500;color:#333;margin-bottom:4px}' +
    '.grievance-detail{font-size:13px;color:#666;margin-top:4px}' +
    '.grievance-step{display:inline-block;padding:3px 8px;background:#E0E7FF;color:#4F46E5;border-radius:6px;font-size:11px;font-weight:500;margin-top:8px}' +

    // Expandable details
    '.grievance-details{display:none;margin-top:12px;padding-top:12px;border-top:1px solid #eee;font-size:13px}' +
    '.grievance-card.expanded .grievance-details{display:block}' +
    '.detail-row{display:flex;gap:8px;margin-bottom:6px}' +
    '.detail-label{color:#666;min-width:90px}' +
    '.detail-value{color:#333;font-weight:500}' +
    '.detail-value.danger{color:#DC2626}' +

    // Empty state
    '.empty-state{text-align:center;padding:60px 20px;color:#999}' +
    '.empty-icon{font-size:48px;margin-bottom:15px}' +

    // Loading
    '.loading{text-align:center;padding:40px;color:#666}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:24px;height:24px;border:3px solid #e0e0e0;border-top-color:#7C3AED;border-radius:50%;animation:spin 0.8s linear infinite}' +

    // Count badge
    '.count-badge{background:rgba(255,255,255,0.2);padding:4px 12px;border-radius:20px;font-size:12px;display:inline-block;margin-top:8px}' +

    // Load more button
    '.load-more-btn{background:#7C3AED;color:white;border:none;padding:14px;border-radius:12px;font-size:14px;font-weight:500;width:100%;margin:15px 0;cursor:pointer}' +
    '.load-more-btn:active{background:#5B21B6}' +

    // Bottom nav - 5 items
    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;display:flex;justify-content:space-around;padding:8px 0 max(8px,env(safe-area-inset-bottom));box-shadow:0 -2px 10px rgba(0,0,0,0.1);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:6px 10px;text-decoration:none;color:#666;font-size:10px;min-width:60px}' +
    '.nav-item.active{color:#7C3AED}' +
    '.nav-icon{font-size:22px;margin-bottom:3px}' +

    '</style></head><body>' +

    '<div class="header">' +
    '<h2>📋 Grievances</h2>' +
    '<div class="filters">' +
    '<button class="filter-pill active" data-filter="all" onclick="setFilter(\'all\',this)">All</button>' +
    '<button class="filter-pill" data-filter="open" onclick="setFilter(\'open\',this)">Open</button>' +
    '<button class="filter-pill" data-filter="pending" onclick="setFilter(\'pending\',this)">Pending</button>' +
    '<button class="filter-pill danger" data-filter="overdue" onclick="setFilter(\'overdue\',this)">⚠️ Overdue</button>' +
    '<button class="filter-pill" data-filter="resolved" onclick="setFilter(\'resolved\',this)">Resolved</button>' +
    '</div>' +
    '<div class="count-badge" id="countBadge">Loading...</div>' +
    '</div>' +

    '<div class="grievance-list" id="grievanceList">' +
    '<div class="loading"><div class="spinner"></div><div style="margin-top:15px">Loading grievances...</div></div>' +
    '</div>' +

    // Bottom Navigation - 5 items
    '<nav class="bottom-nav">' +
    '<a class="nav-item" href="' + baseUrl + '">' +
    '<span class="nav-icon">📊</span>Home</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=search">' +
    '<span class="nav-icon">🔍</span>Search</a>' +
    '<a class="nav-item active" href="' + baseUrl + '?page=grievances">' +
    '<span class="nav-icon">📋</span>Cases</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=members">' +
    '<span class="nav-icon">👥</span>Members</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=links">' +
    '<span class="nav-icon">🔗</span>Links</a>' +
    '</nav>' +

    '<script>' +
    'var allData=[];' +
    'var currentFilter="all";' +
    'var PAGE_SIZE=20;' +
    'var displayLimit=PAGE_SIZE;' +

    // Offline detection
    'function isOnline(){return navigator.onLine!==false}' +
    'function showOfflineWarning(){' +
    '  document.getElementById("grievanceList").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">📶</div><div>You appear to be offline</div><div style=\\"font-size:12px;color:#666;margin-top:8px\\">Check your connection and try again</div><button class=\\"load-more-btn\\" style=\\"margin-top:15px;max-width:200px\\" onclick=\\"loadData()\\">Retry</button></div>";' +
    '}' +

    // Memory cleanup on page unload
    'window.addEventListener("pagehide",function(){allData=[];});' +

    // Check URL for filter parameter
    'var urlParams=new URLSearchParams(window.location.search);' +
    'var initialFilter=urlParams.get("filter");' +

    'function setFilter(filter,btn){' +
    '  currentFilter=filter;' +
    '  displayLimit=PAGE_SIZE;' +
    '  document.querySelectorAll(".filter-pill").forEach(function(p){p.classList.remove("active")});' +
    '  if(btn)btn.classList.add("active");' +
    '  renderList();' +
    '}' +

    'function loadMore(){' +
    '  displayLimit+=PAGE_SIZE;' +
    '  renderList();' +
    '}' +

    'function getStatusClass(g){' +
    '  if(g.isOverdue)return"status-overdue";' +
    '  if(!g.status)return"";' +
    '  var s=g.status.toLowerCase();' +
    '  if(s.indexOf("open")>=0)return"status-open";' +
    '  if(s.indexOf("pending")>=0)return"status-pending";' +
    '  return"status-resolved";' +
    '}' +

    'function getStatusText(g){' +
    '  if(g.isOverdue)return"⚠️ OVERDUE";' +
    '  return g.status||"";' +
    '}' +

    'function matchesFilter(g){' +
    '  if(currentFilter==="all")return true;' +
    '  if(currentFilter==="overdue")return g.isOverdue;' +
    '  if(!g.status)return false;' +
    '  var s=g.status.toLowerCase();' +
    '  if(currentFilter==="open")return s.indexOf("open")>=0;' +
    '  if(currentFilter==="pending")return s.indexOf("pending")>=0;' +
    '  if(currentFilter==="resolved")return s.indexOf("resolved")>=0||s.indexOf("withdrawn")>=0||s.indexOf("closed")>=0;' +
    '  return true;' +
    '}' +

    'function toggleCard(el){el.classList.toggle("expanded")}' +

    'function renderList(){' +
    '  var filtered=allData.filter(function(g){return matchesFilter(g)});' +
    '  var showing=Math.min(displayLimit,filtered.length);' +
    '  document.getElementById("countBadge").textContent="Showing "+showing+" of "+filtered.length+(filtered.length<allData.length?" (filtered)":"");' +
    '  var c=document.getElementById("grievanceList");' +
    '  if(filtered.length===0){' +
    '    c.innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">📋</div><div>No grievances found</div></div>";' +
    '    return;' +
    '  }' +
    '  var html=filtered.slice(0,displayLimit).map(function(g){' +
    '    var cardClass="grievance-card"+(g.isOverdue?" overdue":"");' +
    '    var daysInfo=g.isOverdue?"<span class=\\"detail-value danger\\">⚠️ PAST DUE</span>":(typeof g.daysToDeadline==="number"?"<span class=\\"detail-value\\">"+g.daysToDeadline+" days</span>":"<span class=\\"detail-value\\">N/A</span>");' +
    '    return"<div class=\\""+cardClass+"\\" onclick=\\"toggleCard(this)\\">"+"<div class=\\"grievance-header\\">"+"<span class=\\"grievance-id\\">"+g.id+"</span>"+"<span class=\\"grievance-status "+getStatusClass(g)+"\\">"+getStatusText(g)+"</span>"+"</div>"+"<div class=\\"grievance-name\\">"+g.name+"</div>"+(g.category?"<div class=\\"grievance-detail\\">"+g.category+"</div>":"")+(g.step?"<span class=\\"grievance-step\\">"+g.step+"</span>":"")+"<div class=\\"grievance-details\\">"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">📅 Filed:</span><span class=\\"detail-value\\">"+g.filedDate+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">🔔 Incident:</span><span class=\\"detail-value\\">"+g.incidentDate+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">⏰ Next Due:</span>"+daysInfo+"</div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">⏱️ Days Open:</span><span class=\\"detail-value\\">"+g.daysOpen+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">📍 Location:</span><span class=\\"detail-value\\">"+g.location+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">📜 Articles:</span><span class=\\"detail-value\\">"+g.articles+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">🛡️ Steward:</span><span class=\\"detail-value\\">"+g.steward+"</span></div>"+(g.resolution?"<div class=\\"detail-row\\"><span class=\\"detail-label\\">✅ Resolution:</span><span class=\\"detail-value\\">"+g.resolution+"</span></div>":"")+"</div>"+"</div>";' +
    '  }).join("");' +
    '  if(filtered.length>displayLimit){' +
    '    html+="<button class=\\"load-more-btn\\" onclick=\\"loadMore()\\">Load More ("+(filtered.length-displayLimit)+" remaining)</button>";' +
    '  }' +
    '  c.innerHTML=html;' +
    '}' +

    'var CACHE_KEY="grievance_data";' +
    'var CACHE_DURATION=300000;' +  // 5 minutes

    'function loadData(){' +
    '  console.log("Loading grievance data...");' +
    '  try{' +
    '    var cached=sessionStorage.getItem(CACHE_KEY);' +
    '    if(cached){' +
    '      var parsed=JSON.parse(cached);' +
    '      if(Date.now()-parsed.time<CACHE_DURATION){' +
    '        console.log("Using cached data");' +
    '        allData=parsed.data||[];' +
    '        applyInitialFilter();' +
    '        renderList();' +
    '        return;' +
    '      }' +
    '    }' +
    '  }catch(e){console.log("Cache read error:",e)}' +
    '  if(!isOnline()){showOfflineWarning();return}' +
    '  google.script.run.withSuccessHandler(function(data){' +
    '    console.log("Data received:",data?data.length:0,"items");' +
    '    allData=data||[];' +
    '    try{sessionStorage.setItem(CACHE_KEY,JSON.stringify({data:allData,time:Date.now()}))}catch(e){}' +
    '    applyInitialFilter();' +
    '    renderList();' +
    '  }).withFailureHandler(function(err){' +
    '    console.error("Failed to load data:",err);' +
    '    document.getElementById("grievanceList").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">⚠️</div><div>Error loading data</div><div style=\\"font-size:11px;color:#999;margin-top:8px\\">"+String(err||"Unknown error")+"</div></div>";' +
    '  }).getWebAppGrievanceList();' +
    '}' +

    'function applyInitialFilter(){' +
    '  if(initialFilter){' +
    '    currentFilter=initialFilter;' +
    '    var btn=document.querySelector("[data-filter=\\""+initialFilter+"\\"]");' +
    '    if(btn){document.querySelectorAll(".filter-pill").forEach(function(p){p.classList.remove("active")});btn.classList.add("active")}' +
    '  }' +
    '}' +

    'loadData();' +
    '</script>' +

    '</body></html>';
}

/**
 * Returns the member list HTML for web app
 */
function getWebAppMemberListHtml() {
  var baseUrl = ScriptApp.getService().getUrl();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<title>Members - 509 Dashboard</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh;padding-bottom:80px}' +

    // Header with search
    '.header{background:linear-gradient(135deg,#7C3AED,#5B21B6);color:white;padding:15px;position:sticky;top:0;z-index:100}' +
    '.header h2{font-size:clamp(18px,4vw,22px);text-align:center;margin-bottom:12px}' +
    '.search-container{position:relative}' +
    '.search-input{width:100%;padding:12px 12px 12px 40px;border:none;border-radius:12px;font-size:15px;background:white}' +
    '.search-input:focus{outline:none;box-shadow:0 0 0 3px rgba(124,58,237,0.3)}' +
    '.search-icon{position:absolute;left:12px;top:50%;transform:translateY(-50%);font-size:18px;color:#666}' +

    // Filter pills
    '.filters{display:flex;gap:8px;overflow-x:auto;padding:8px 0 2px;-webkit-overflow-scrolling:touch}' +
    '.filter-pill{flex-shrink:0;padding:8px 14px;border-radius:20px;font-size:12px;font-weight:500;border:none;cursor:pointer;background:rgba(255,255,255,0.2);color:white}' +
    '.filter-pill.active{background:white;color:#7C3AED}' +

    // Count badge
    '.count-badge{background:rgba(255,255,255,0.2);padding:4px 12px;border-radius:20px;font-size:12px;display:inline-block;margin-top:8px}' +

    // Member list
    '.member-list{padding:15px}' +
    '.member-card{background:white;padding:16px;border-radius:14px;box-shadow:0 2px 6px rgba(0,0,0,0.06);margin-bottom:12px;cursor:pointer;transition:all 0.2s}' +
    '.member-card:active{transform:scale(0.99)}' +
    '.member-card.has-grievance{border-left:4px solid #F97316}' +
    '.member-header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:4px}' +
    '.member-name{font-size:16px;font-weight:600;color:#333}' +
    '.member-id{font-size:12px;color:#7C3AED;font-weight:500}' +
    '.member-title{font-size:14px;color:#666;margin-bottom:4px}' +
    '.member-location{font-size:13px;color:#888}' +
    '.member-badges{display:flex;gap:6px;margin-top:8px;flex-wrap:wrap}' +
    '.badge{padding:3px 8px;border-radius:12px;font-size:11px;font-weight:500}' +
    '.badge-steward{background:#DDD6FE;color:#7C3AED}' +
    '.badge-grievance{background:#FEF3C7;color:#D97706}' +

    // Expandable details
    '.member-details{display:none;margin-top:12px;padding-top:12px;border-top:1px solid #eee;font-size:13px}' +
    '.member-card.expanded .member-details{display:block}' +
    '.detail-row{display:flex;gap:8px;margin-bottom:6px}' +
    '.detail-label{color:#666;min-width:80px}' +
    '.detail-value{color:#333;font-weight:500}' +

    // Empty state
    '.empty-state{text-align:center;padding:60px 20px;color:#999}' +
    '.empty-icon{font-size:48px;margin-bottom:15px}' +

    // Loading
    '.loading{text-align:center;padding:40px;color:#666}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:24px;height:24px;border:3px solid #e0e0e0;border-top-color:#7C3AED;border-radius:50%;animation:spin 0.8s linear infinite}' +

    // Load more button
    '.load-more-btn{background:#7C3AED;color:white;border:none;padding:14px;border-radius:12px;font-size:14px;font-weight:500;width:100%;margin:15px 0;cursor:pointer}' +
    '.load-more-btn:active{background:#5B21B6}' +

    // Bottom nav - 5 items
    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;display:flex;justify-content:space-around;padding:8px 0 max(8px,env(safe-area-inset-bottom));box-shadow:0 -2px 10px rgba(0,0,0,0.1);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:6px 10px;text-decoration:none;color:#666;font-size:10px;min-width:60px}' +
    '.nav-item.active{color:#7C3AED}' +
    '.nav-icon{font-size:22px;margin-bottom:3px}' +

    '</style></head><body>' +

    '<div class="header">' +
    '<h2>👥 Members</h2>' +
    '<div class="search-container">' +
    '<span class="search-icon">🔍</span>' +
    '<input type="text" class="search-input" id="searchInput" placeholder="Search by name, ID, title..." oninput="handleSearch()">' +
    '</div>' +
    '<div class="filters">' +
    '<button class="filter-pill active" data-filter="all" onclick="setFilter(\'all\',this)">All</button>' +
    '<button class="filter-pill" data-filter="steward" onclick="setFilter(\'steward\',this)">Stewards</button>' +
    '<button class="filter-pill" data-filter="grievance" onclick="setFilter(\'grievance\',this)">With Grievance</button>' +
    '</div>' +
    '<div class="count-badge" id="countBadge">Loading...</div>' +
    '</div>' +

    '<div class="member-list" id="memberList">' +
    '<div class="loading"><div class="spinner"></div><div style="margin-top:15px">Loading members...</div></div>' +
    '</div>' +

    // Bottom Navigation - 5 items
    '<nav class="bottom-nav">' +
    '<a class="nav-item" href="' + baseUrl + '">' +
    '<span class="nav-icon">📊</span>Home</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=search">' +
    '<span class="nav-icon">🔍</span>Search</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=grievances">' +
    '<span class="nav-icon">📋</span>Cases</a>' +
    '<a class="nav-item active" href="' + baseUrl + '?page=members">' +
    '<span class="nav-icon">👥</span>Members</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=links">' +
    '<span class="nav-icon">🔗</span>Links</a>' +
    '</nav>' +

    '<script>' +
    'var allData=[];' +
    'var currentFilter="all";' +
    'var PAGE_SIZE=25;' +
    'var displayLimit=PAGE_SIZE;' +
    'var searchTimeout=null;' +

    // Offline detection
    'function isOnline(){return navigator.onLine!==false}' +
    'function showOfflineWarning(){' +
    '  document.getElementById("memberList").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">📶</div><div>You appear to be offline</div><div style=\\"font-size:12px;color:#666;margin-top:8px\\">Check your connection and try again</div><button class=\\"load-more-btn\\" style=\\"margin-top:15px;max-width:200px\\" onclick=\\"loadData()\\">Retry</button></div>";' +
    '}' +

    // Memory cleanup on page unload
    'window.addEventListener("pagehide",function(){allData=[];});' +

    'function setFilter(filter,btn){' +
    '  currentFilter=filter;' +
    '  displayLimit=PAGE_SIZE;' +
    '  document.querySelectorAll(".filter-pill").forEach(function(p){p.classList.remove("active")});' +
    '  btn.classList.add("active");' +
    '  filterMembers();' +
    '}' +

    'function loadMore(){' +
    '  displayLimit+=PAGE_SIZE;' +
    '  filterMembers();' +
    '}' +

    'function toggleCard(el){el.classList.toggle("expanded")}' +

    'function handleSearch(){' +
    '  clearTimeout(searchTimeout);' +
    '  searchTimeout=setTimeout(function(){' +
    '    displayLimit=PAGE_SIZE;' +
    '    filterMembers();' +
    '  },300);' +
    '}' +

    'function filterMembers(){' +
    '  var query=(document.getElementById("searchInput").value||"").toLowerCase();' +
    '  var filtered=allData.filter(function(m){' +
    '    var matchesQuery=!query||query.length<2||m.name.toLowerCase().indexOf(query)>=0||m.id.toLowerCase().indexOf(query)>=0||(m.title||"").toLowerCase().indexOf(query)>=0||(m.location||"").toLowerCase().indexOf(query)>=0;' +
    '    var matchesFilter=currentFilter==="all"||(currentFilter==="steward"&&m.isSteward)||(currentFilter==="grievance"&&m.hasOpenGrievance);' +
    '    return matchesQuery&&matchesFilter;' +
    '  });' +
    '  var showing=Math.min(displayLimit,filtered.length);' +
    '  document.getElementById("countBadge").textContent="Showing "+showing+" of "+filtered.length+(filtered.length<allData.length?" (filtered)":"");' +
    '  renderList(filtered);' +
    '}' +

    'function renderList(data){' +
    '  var c=document.getElementById("memberList");' +
    '  if(!data||data.length===0){' +
    '    c.innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">👥</div><div>No members found</div></div>";' +
    '    return;' +
    '  }' +
    '  var html=data.slice(0,displayLimit).map(function(m){' +
    '    var cardClass="member-card"+(m.hasOpenGrievance?" has-grievance":"");' +
    '    var badges="";' +
    '    if(m.isSteward)badges+="<span class=\\"badge badge-steward\\">🛡️ Steward</span>";' +
    '    if(m.hasOpenGrievance)badges+="<span class=\\"badge badge-grievance\\">⚠️ Open Grievance</span>";' +
    '    return"<div class=\\""+cardClass+"\\" onclick=\\"toggleCard(this)\\">"+"<div class=\\"member-header\\"><span class=\\"member-name\\">"+m.name+"</span><span class=\\"member-id\\">"+m.id+"</span></div>"+"<div class=\\"member-title\\">"+m.title+"</div>"+"<div class=\\"member-location\\">📍 "+m.location+"</div>"+(badges?"<div class=\\"member-badges\\">"+badges+"</div>":"")+"<div class=\\"member-details\\">"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">📧 Email:</span><span class=\\"detail-value\\">"+(m.email||"N/A")+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">📞 Phone:</span><span class=\\"detail-value\\">"+(m.phone||"N/A")+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">🏢 Unit:</span><span class=\\"detail-value\\">"+m.unit+"</span></div>"+"<div class=\\"detail-row\\"><span class=\\"detail-label\\">👔 Supervisor:</span><span class=\\"detail-value\\">"+m.supervisor+"</span></div>"+"</div>"+"</div>";' +
    '  }).join("");' +
    '  if(data.length>displayLimit){' +
    '    html+="<button class=\\"load-more-btn\\" onclick=\\"loadMore()\\">Load More ("+(data.length-displayLimit)+" remaining)</button>";' +
    '  }' +
    '  c.innerHTML=html;' +
    '}' +

    'var CACHE_KEY="member_data";' +
    'var CACHE_DURATION=300000;' +  // 5 minutes

    'function loadData(){' +
    '  try{' +
    '    var cached=sessionStorage.getItem(CACHE_KEY);' +
    '    if(cached){' +
    '      var parsed=JSON.parse(cached);' +
    '      if(Date.now()-parsed.time<CACHE_DURATION){' +
    '        console.log("Using cached member data");' +
    '        allData=parsed.data||[];' +
    '        filterMembers();' +
    '        return;' +
    '      }' +
    '    }' +
    '  }catch(e){console.log("Cache read error:",e)}' +
    '  if(!isOnline()){showOfflineWarning();return}' +
    '  google.script.run.withSuccessHandler(function(data){' +
    '    allData=data||[];' +
    '    try{sessionStorage.setItem(CACHE_KEY,JSON.stringify({data:allData,time:Date.now()}))}catch(e){}' +
    '    filterMembers();' +
    '  }).withFailureHandler(function(err){' +
    '    document.getElementById("memberList").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">⚠️</div><div>Error loading data</div></div>";' +
    '  }).getWebAppMemberList();' +
    '}' +

    'loadData();' +
    '</script>' +

    '</body></html>';
}

/**
 * Returns the links/resources page HTML for web app
 */
function getWebAppLinksHtml() {
  var baseUrl = ScriptApp.getService().getUrl();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<title>Links - 509 Dashboard</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh;padding-bottom:80px}' +

    // Header
    '.header{background:linear-gradient(135deg,#7C3AED,#5B21B6);color:white;padding:20px;text-align:center;position:sticky;top:0;z-index:100}' +
    '.header h2{font-size:clamp(18px,4vw,22px)}' +
    '.header .subtitle{font-size:13px;opacity:0.9;margin-top:5px}' +

    // Container
    '.container{padding:15px;max-width:600px;margin:0 auto}' +

    // Section titles
    '.section-title{font-size:clamp(14px,3.5vw,18px);font-weight:600;color:#333;margin:20px 0 12px;padding-left:5px}' +

    // Link cards
    '.link-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:12px}' +
    '.link-card{background:white;padding:20px 16px;border-radius:16px;box-shadow:0 2px 8px rgba(0,0,0,0.08);text-decoration:none;display:flex;flex-direction:column;align-items:center;text-align:center;gap:10px;transition:all 0.2s}' +
    '.link-card:active{transform:scale(0.96);background:#f8f4ff}' +
    '.link-icon{font-size:32px}' +
    '.link-label{font-weight:600;color:#333;font-size:14px}' +
    '.link-desc{font-size:12px;color:#666}' +

    // Full-width link
    '.link-card.full{grid-column:span 2;flex-direction:row;padding:16px;justify-content:flex-start;text-align:left}' +
    '.link-card.full .link-icon{font-size:28px}' +
    '.link-card.full .link-content{flex:1}' +

    // GitHub special styling
    '.link-card.github{background:linear-gradient(135deg,#24292e,#1a1e22);color:white}' +
    '.link-card.github .link-label{color:white}' +
    '.link-card.github .link-desc{color:rgba(255,255,255,0.7)}' +

    // Loading
    '.loading{text-align:center;padding:40px;color:#666}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:24px;height:24px;border:3px solid #e0e0e0;border-top-color:#7C3AED;border-radius:50%;animation:spin 0.8s linear infinite}' +

    // Bottom nav - 5 items
    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;display:flex;justify-content:space-around;padding:8px 0 max(8px,env(safe-area-inset-bottom));box-shadow:0 -2px 10px rgba(0,0,0,0.1);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:6px 10px;text-decoration:none;color:#666;font-size:10px;min-width:60px}' +
    '.nav-item.active{color:#7C3AED}' +
    '.nav-icon{font-size:22px;margin-bottom:3px}' +

    '</style></head><body>' +

    '<div class="header">' +
    '<h2>🔗 Links & Resources</h2>' +
    '<div class="subtitle">Quick access to forms and tools</div>' +
    '</div>' +

    '<div class="container" id="linksContent">' +
    '<div class="loading"><div class="spinner"></div><div style="margin-top:15px">Loading links...</div></div>' +
    '</div>' +

    // Bottom Navigation - 5 items
    '<nav class="bottom-nav">' +
    '<a class="nav-item" href="' + baseUrl + '">' +
    '<span class="nav-icon">📊</span>Home</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=search">' +
    '<span class="nav-icon">🔍</span>Search</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=grievances">' +
    '<span class="nav-icon">📋</span>Cases</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=members">' +
    '<span class="nav-icon">👥</span>Members</a>' +
    '<a class="nav-item active" href="' + baseUrl + '?page=links">' +
    '<span class="nav-icon">🔗</span>Links</a>' +
    '</nav>' +

    '<script>' +
    'function loadLinks(){' +
    '  google.script.run.withSuccessHandler(function(links){' +
    '    renderLinks(links);' +
    '  }).withFailureHandler(function(err){' +
    '    document.getElementById("linksContent").innerHTML="<div class=\\"loading\\">⚠️ Error loading links</div>";' +
    '  }).getWebAppResourceLinks();' +
    '}' +

    'function renderLinks(links){' +
    '  var html="";' +

    '  // Forms section' +
    '  html+="<div class=\\"section-title\\">📝 Forms</div>";' +
    '  html+="<div class=\\"link-grid\\">";' +
    '  if(links.grievanceForm){html+="<a class=\\"link-card\\" href=\\""+links.grievanceForm+"\\" target=\\"_blank\\"><span class=\\"link-icon\\">📋</span><span class=\\"link-label\\">Grievance Form</span><span class=\\"link-desc\\">File a grievance</span></a>";}' +
    '  if(links.contactForm){html+="<a class=\\"link-card\\" href=\\""+links.contactForm+"\\" target=\\"_blank\\"><span class=\\"link-icon\\">✉️</span><span class=\\"link-label\\">Contact Form</span><span class=\\"link-desc\\">Send a message</span></a>";}' +
    '  if(links.satisfactionForm){html+="<a class=\\"link-card\\" href=\\""+links.satisfactionForm+"\\" target=\\"_blank\\"><span class=\\"link-icon\\">📊</span><span class=\\"link-label\\">Satisfaction Survey</span><span class=\\"link-desc\\">Give feedback</span></a>";}' +
    '  if(!links.grievanceForm&&!links.contactForm&&!links.satisfactionForm){html+="<div class=\\"link-card full\\"><span class=\\"link-icon\\">ℹ️</span><div class=\\"link-content\\"><span class=\\"link-label\\">No Forms Configured</span><span class=\\"link-desc\\">Add form URLs to Config sheet</span></div></div>";}' +
    '  html+="</div>";' +

    '  // Resources section' +
    '  html+="<div class=\\"section-title\\">🔧 Resources</div>";' +
    '  html+="<div class=\\"link-grid\\">";' +
    '  html+="<a class=\\"link-card\\" href=\\""+links.spreadsheetUrl+"\\" target=\\"_blank\\"><span class=\\"link-icon\\">📊</span><span class=\\"link-label\\">Spreadsheet</span><span class=\\"link-desc\\">Open full dashboard</span></a>";' +
    '  html+="<a class=\\"link-card github\\" href=\\""+links.githubRepo+"\\" target=\\"_blank\\"><span class=\\"link-icon\\">📦</span><span class=\\"link-label\\">GitHub Repo</span><span class=\\"link-desc\\">Source code</span></a>";' +
    '  html+="</div>";' +

    '  document.getElementById("linksContent").innerHTML=html;' +
    '}' +

    'loadLinks();' +
    '</script>' +

    '</body></html>';
}

/**
 * API function to get search results for web app
 * @param {string} query - Search query
 * @param {string} tab - Tab filter (all, members, grievances)
 * @returns {Array} Search results
 */
function getWebAppSearchResults(query, tab) {
  return getMobileSearchData(query, tab);
}

/**
 * API function to get grievance list for web app (full fields like Interactive Dashboard)
 * @returns {Array} Grievance data with all fields
 */
function getWebAppGrievanceList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      Logger.log('getWebAppGrievanceList: No active spreadsheet');
      return [];
    }

    var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (!sheet) {
      Logger.log('getWebAppGrievanceList: Grievance Log sheet not found');
      return [];
    }

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('getWebAppGrievanceList: No data rows in sheet');
      return [];
    }

    var data = sheet.getRange(2, 1, lastRow - 1, GRIEVANCE_COLS.QUICK_ACTIONS).getValues();
    var tz = Session.getScriptTimeZone();

    var result = data.map(function(row) {
      var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '';
      // Skip blank rows - must have a valid grievance ID starting with G
      if (!grievanceId || (typeof grievanceId === 'string' && !grievanceId.toString().match(/^G/i))) return null;

      var filed = row[GRIEVANCE_COLS.DATE_FILED - 1];
      var incident = row[GRIEVANCE_COLS.INCIDENT_DATE - 1];
      var nextDue = row[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];
      var daysToDeadline = row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];

      return {
        id: grievanceId,
        memberId: row[GRIEVANCE_COLS.MEMBER_ID - 1] || '',
        name: ((row[GRIEVANCE_COLS.FIRST_NAME - 1] || '') + ' ' + (row[GRIEVANCE_COLS.LAST_NAME - 1] || '')).trim(),
        status: row[GRIEVANCE_COLS.STATUS - 1] || 'Filed',
        step: row[GRIEVANCE_COLS.CURRENT_STEP - 1] || 'Step I',
        category: row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || 'N/A',
        articles: row[GRIEVANCE_COLS.ARTICLES - 1] || 'N/A',
        filedDate: filed instanceof Date ? Utilities.formatDate(filed, tz, 'MM/dd/yyyy') : (filed || 'N/A'),
        incidentDate: incident instanceof Date ? Utilities.formatDate(incident, tz, 'MM/dd/yyyy') : (incident || 'N/A'),
        nextActionDue: nextDue instanceof Date ? Utilities.formatDate(nextDue, tz, 'MM/dd/yyyy') : (nextDue || 'N/A'),
        daysToDeadline: daysToDeadline,
        isOverdue: daysToDeadline === 'Overdue' || (typeof daysToDeadline === 'number' && daysToDeadline < 0),
        daysOpen: row[GRIEVANCE_COLS.DAYS_OPEN - 1] || 0,
        location: row[GRIEVANCE_COLS.LOCATION - 1] || 'N/A',
        unit: row[GRIEVANCE_COLS.UNIT - 1] || 'N/A',
        steward: row[GRIEVANCE_COLS.STEWARD - 1] || 'N/A',
        resolution: row[GRIEVANCE_COLS.RESOLUTION - 1] || ''
      };
    }).filter(function(g) { return g !== null; }).slice(0, 100);

    Logger.log('getWebAppGrievanceList: Returning ' + result.length + ' grievances');
    return result;
  } catch (e) {
    Logger.log('getWebAppGrievanceList error: ' + e.toString());
    throw new Error('Failed to load grievances: ' + e.message);
  }
}

/**
 * API function to get member list for web app
 * @returns {Array} Member data
 */
function getWebAppMemberList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      Logger.log('getWebAppMemberList: No active spreadsheet');
      return [];
    }

    var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (!sheet) {
      Logger.log('getWebAppMemberList: Member Directory sheet not found');
      return [];
    }

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('getWebAppMemberList: No data rows in sheet');
      return [];
    }

    var data = sheet.getRange(2, 1, lastRow - 1, MEMBER_COLS.QUICK_ACTIONS).getValues();

    var result = data.map(function(row) {
      var memberId = row[MEMBER_COLS.MEMBER_ID - 1] || '';
      // Skip blank rows - must have a valid member ID starting with M
      if (!memberId || (typeof memberId === 'string' && !memberId.toString().match(/^M/i))) return null;

      return {
        id: memberId,
        firstName: row[MEMBER_COLS.FIRST_NAME - 1] || '',
        lastName: row[MEMBER_COLS.LAST_NAME - 1] || '',
        name: ((row[MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' + (row[MEMBER_COLS.LAST_NAME - 1] || '')).trim(),
        title: row[MEMBER_COLS.JOB_TITLE - 1] || 'N/A',
        location: row[MEMBER_COLS.WORK_LOCATION - 1] || 'N/A',
        unit: row[MEMBER_COLS.UNIT - 1] || 'N/A',
        email: row[MEMBER_COLS.EMAIL - 1] || '',
        phone: row[MEMBER_COLS.PHONE - 1] || '',
        isSteward: row[MEMBER_COLS.IS_STEWARD - 1] === 'Yes',
        supervisor: row[MEMBER_COLS.SUPERVISOR - 1] || 'N/A',
        hasOpenGrievance: row[MEMBER_COLS.HAS_OPEN_GRIEVANCE - 1] === 'Yes'
      };
    }).filter(function(m) { return m !== null; }).slice(0, 100);

    Logger.log('getWebAppMemberList: Returning ' + result.length + ' members');
    return result;
  } catch (e) {
    Logger.log('getWebAppMemberList error: ' + e.toString());
    throw new Error('Failed to load members: ' + e.message);
  }
}

/**
 * API function to get resource links for web app
 * @returns {Object} Resource links
 */
function getWebAppResourceLinks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  var links = {
    grievanceForm: '',
    contactForm: '',
    satisfactionForm: '',
    spreadsheetUrl: ss.getUrl(),
    orgWebsite: '',
    githubRepo: 'https://github.com/Woop91/509-dashboard-second'
  };

  // Try to get form URLs from Config sheet
  if (configSheet) {
    try {
      var configData = configSheet.getDataRange().getValues();
      for (var i = 0; i < configData.length; i++) {
        var row = configData[i];
        for (var j = 0; j < row.length; j++) {
          var val = String(row[j] || '').toLowerCase();
          if (val.indexOf('grievance') >= 0 && val.indexOf('form') >= 0 && row[j + 1]) {
            links.grievanceForm = String(row[j + 1]);
          } else if (val.indexOf('contact') >= 0 && val.indexOf('form') >= 0 && row[j + 1]) {
            links.contactForm = String(row[j + 1]);
          } else if (val.indexOf('satisfaction') >= 0 && val.indexOf('form') >= 0 && row[j + 1]) {
            links.satisfactionForm = String(row[j + 1]);
          }
        }
      }
    } catch (e) {
      // Ignore errors reading config
    }
  }

  return links;
}

/**
 * API function to get dashboard stats with win rate for web app
 * @returns {Object} Dashboard statistics
 */
function getWebAppDashboardStats() {
  var stats = getMobileDashboardStats();

  // Calculate win rate
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (sheet && sheet.getLastRow() > 1) {
    var resolutions = sheet.getRange(2, GRIEVANCE_COLS.RESOLUTION, sheet.getLastRow() - 1, 1).getValues();
    var won = 0, total = 0;
    resolutions.forEach(function(row) {
      var res = (row[0] || '').toString().toLowerCase();
      if (res) {
        total++;
        if (res.indexOf('won') >= 0 || res.indexOf('favorable') >= 0) {
          won++;
        }
      }
    });
    stats.winRate = total > 0 ? Math.round((won / total) * 100) + '%' : 'N/A';
  } else {
    stats.winRate = 'N/A';
  }

  // Also get total members
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (memberSheet && memberSheet.getLastRow() > 1) {
    var memberIds = memberSheet.getRange(2, MEMBER_COLS.MEMBER_ID, memberSheet.getLastRow() - 1, 1).getValues();
    var validMembers = memberIds.filter(function(row) {
      var id = row[0] || '';
      return id && id.toString().match(/^M/i);
    }).length;
    stats.totalMembers = validMembers;
  } else {
    stats.totalMembers = 0;
  }

  return stats;
}

/**
 * Menu function to get the web app URL
 */
function showWebAppUrl() {
  var url = ScriptApp.getService().getUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert(
      '📱 Web App Not Deployed',
      'To access the dashboard on mobile:\n\n' +
      '1. Go to Extensions → Apps Script\n' +
      '2. Click "Deploy" → "New deployment"\n' +
      '3. Select "Web app"\n' +
      '4. Set "Who has access" appropriately\n' +
      '5. Click "Deploy"\n' +
      '6. Copy the URL and open it on your phone',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    '📱 Mobile Dashboard URL',
    'Open this URL on your mobile device:\n\n' + url + '\n\n' +
    'Tip: Add it to your home screen for quick access!',
    ui.ButtonSet.OK
  );
}
