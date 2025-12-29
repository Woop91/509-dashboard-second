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
 * Returns the main dashboard HTML for web app
 */
function getWebAppDashboardHtml() {
  var stats = getMobileDashboardStats();
  var baseUrl = ScriptApp.getService().getUrl();

  return '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">' +
    '<meta name="apple-mobile-web-app-capable" content="yes">' +
    '<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">' +
    '<link rel="apple-touch-icon" href="data:image/svg+xml,<svg xmlns=\'http://www.w3.org/2000/svg\' viewBox=\'0 0 100 100\'><text y=\'.9em\' font-size=\'90\'>📊</text></svg>">' +
    '<title>509 Dashboard</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f5f5f5;min-height:100vh;padding-bottom:80px}' +

    // Header
    '.header{background:linear-gradient(135deg,#7C3AED,#5B21B6);color:white;padding:20px;text-align:center;position:sticky;top:0;z-index:100}' +
    '.header h1{font-size:clamp(20px,5vw,28px);margin-bottom:5px}' +
    '.header .subtitle{font-size:clamp(12px,3vw,14px);opacity:0.9}' +

    // Container
    '.container{padding:15px;max-width:600px;margin:0 auto}' +

    // Stats grid
    '.stats{display:grid;grid-template-columns:repeat(2,1fr);gap:12px;margin-bottom:20px}' +
    '.stat-card{background:white;padding:20px 15px;border-radius:16px;box-shadow:0 2px 8px rgba(0,0,0,0.08);text-align:center}' +
    '.stat-value{font-size:clamp(28px,8vw,40px);font-weight:bold;color:#7C3AED}' +
    '.stat-value.warning{color:#F97316}' +
    '.stat-value.danger{color:#DC2626}' +
    '.stat-value.success{color:#059669}' +
    '.stat-label{font-size:clamp(11px,2.5vw,13px);color:#666;text-transform:uppercase;margin-top:5px;letter-spacing:0.5px}' +

    // Section titles
    '.section-title{font-size:clamp(14px,3.5vw,18px);font-weight:600;color:#333;margin:20px 0 12px;padding-left:5px}' +

    // Action buttons
    '.actions{display:flex;flex-direction:column;gap:10px}' +
    '.action-btn{background:white;border:none;padding:18px 16px;border-radius:16px;box-shadow:0 2px 8px rgba(0,0,0,0.08);' +
    'width:100%;text-align:left;display:flex;align-items:center;gap:15px;font-size:16px;cursor:pointer;' +
    'text-decoration:none;color:inherit;min-height:64px;transition:all 0.2s}' +
    '.action-btn:active{transform:scale(0.98);background:#f0f0f0}' +
    '.action-icon{font-size:28px;width:50px;height:50px;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,#EDE9FE,#DDD6FE);border-radius:14px;flex-shrink:0}' +
    '.action-label{font-weight:600;color:#333}' +
    '.action-desc{font-size:13px;color:#666;margin-top:3px}' +

    // Bottom nav
    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;display:flex;justify-content:space-around;padding:10px 0 max(10px,env(safe-area-inset-bottom));box-shadow:0 -2px 10px rgba(0,0,0,0.1);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:8px 16px;text-decoration:none;color:#666;font-size:11px;min-width:70px}' +
    '.nav-item.active{color:#7C3AED}' +
    '.nav-icon{font-size:24px;margin-bottom:4px}' +

    // Refresh indicator
    '.refresh-btn{position:absolute;right:15px;top:50%;transform:translateY(-50%);background:rgba(255,255,255,0.2);border:none;color:white;width:40px;height:40px;border-radius:50%;font-size:20px;cursor:pointer}' +

    '</style></head><body>' +

    // Header
    '<div class="header">' +
    '<button class="refresh-btn" onclick="location.reload()">🔄</button>' +
    '<h1>📊 509 Dashboard</h1>' +
    '<div class="subtitle">Union Grievance Management</div>' +
    '</div>' +

    '<div class="container">' +

    // Stats section
    '<div class="stats">' +
    '<div class="stat-card"><div class="stat-value">' + stats.totalGrievances + '</div><div class="stat-label">Total</div></div>' +
    '<div class="stat-card"><div class="stat-value success">' + stats.activeGrievances + '</div><div class="stat-label">Active</div></div>' +
    '<div class="stat-card"><div class="stat-value warning">' + stats.pendingGrievances + '</div><div class="stat-label">Pending</div></div>' +
    '<div class="stat-card"><div class="stat-value danger">' + stats.overdueGrievances + '</div><div class="stat-label">Overdue</div></div>' +
    '</div>' +

    // Quick Actions
    '<div class="section-title">⚡ Quick Actions</div>' +
    '<div class="actions">' +

    '<a class="action-btn" href="' + baseUrl + '?page=search">' +
    '<div class="action-icon">🔍</div>' +
    '<div><div class="action-label">Search</div><div class="action-desc">Find members or grievances</div></div>' +
    '</a>' +

    '<a class="action-btn" href="' + baseUrl + '?page=grievances">' +
    '<div class="action-icon">📋</div>' +
    '<div><div class="action-label">All Grievances</div><div class="action-desc">Browse and filter grievances</div></div>' +
    '</a>' +

    '</div>' +
    '</div>' +

    // Bottom Navigation
    '<nav class="bottom-nav">' +
    '<a class="nav-item active" href="' + baseUrl + '">' +
    '<span class="nav-icon">📊</span>Dashboard</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=search">' +
    '<span class="nav-icon">🔍</span>Search</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=grievances">' +
    '<span class="nav-icon">📋</span>Grievances</a>' +
    '</nav>' +

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

    // Bottom nav
    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;display:flex;justify-content:space-around;padding:10px 0 max(10px,env(safe-area-inset-bottom));box-shadow:0 -2px 10px rgba(0,0,0,0.1);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:8px 16px;text-decoration:none;color:#666;font-size:11px;min-width:70px}' +
    '.nav-item.active{color:#7C3AED}' +
    '.nav-icon{font-size:24px;margin-bottom:4px}' +

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

    // Bottom Navigation
    '<nav class="bottom-nav">' +
    '<a class="nav-item" href="' + baseUrl + '">' +
    '<span class="nav-icon">📊</span>Dashboard</a>' +
    '<a class="nav-item active" href="' + baseUrl + '?page=search">' +
    '<span class="nav-icon">🔍</span>Search</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=grievances">' +
    '<span class="nav-icon">📋</span>Grievances</a>' +
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
 * Returns the grievance list HTML for web app
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

    // Filter pills
    '.filters{display:flex;gap:8px;overflow-x:auto;padding:2px 0;-webkit-overflow-scrolling:touch}' +
    '.filter-pill{flex-shrink:0;padding:8px 16px;border-radius:20px;font-size:13px;font-weight:500;border:none;cursor:pointer;background:rgba(255,255,255,0.2);color:white}' +
    '.filter-pill.active{background:white;color:#7C3AED}' +

    // List
    '.grievance-list{padding:15px}' +
    '.grievance-card{background:white;padding:16px;border-radius:14px;box-shadow:0 2px 6px rgba(0,0,0,0.06);margin-bottom:12px}' +
    '.grievance-header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:8px}' +
    '.grievance-id{font-size:15px;font-weight:700;color:#7C3AED}' +
    '.grievance-status{padding:4px 10px;border-radius:20px;font-size:11px;font-weight:600}' +
    '.status-open{background:#FEE2E2;color:#DC2626}' +
    '.status-pending{background:#FEF3C7;color:#D97706}' +
    '.status-resolved{background:#D1FAE5;color:#059669}' +
    '.grievance-name{font-size:16px;font-weight:500;color:#333;margin-bottom:4px}' +
    '.grievance-detail{font-size:13px;color:#666;margin-top:4px}' +
    '.grievance-step{display:inline-block;padding:3px 8px;background:#E0E7FF;color:#4F46E5;border-radius:6px;font-size:11px;font-weight:500;margin-top:8px}' +

    // Empty state
    '.empty-state{text-align:center;padding:60px 20px;color:#999}' +
    '.empty-icon{font-size:48px;margin-bottom:15px}' +

    // Loading
    '.loading{text-align:center;padding:40px;color:#666}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.spinner{display:inline-block;width:24px;height:24px;border:3px solid #e0e0e0;border-top-color:#7C3AED;border-radius:50%;animation:spin 0.8s linear infinite}' +

    // Count badge
    '.count-badge{background:rgba(255,255,255,0.2);padding:4px 12px;border-radius:20px;font-size:12px;display:inline-block;margin-top:8px}' +

    // Bottom nav
    '.bottom-nav{position:fixed;bottom:0;left:0;right:0;background:white;display:flex;justify-content:space-around;padding:10px 0 max(10px,env(safe-area-inset-bottom));box-shadow:0 -2px 10px rgba(0,0,0,0.1);z-index:100}' +
    '.nav-item{display:flex;flex-direction:column;align-items:center;padding:8px 16px;text-decoration:none;color:#666;font-size:11px;min-width:70px}' +
    '.nav-item.active{color:#7C3AED}' +
    '.nav-icon{font-size:24px;margin-bottom:4px}' +

    '</style></head><body>' +

    '<div class="header">' +
    '<h2>📋 Grievances</h2>' +
    '<div class="filters">' +
    '<button class="filter-pill active" data-filter="all" onclick="setFilter(\'all\',this)">All</button>' +
    '<button class="filter-pill" data-filter="open" onclick="setFilter(\'open\',this)">Open</button>' +
    '<button class="filter-pill" data-filter="pending" onclick="setFilter(\'pending\',this)">Pending</button>' +
    '<button class="filter-pill" data-filter="resolved" onclick="setFilter(\'resolved\',this)">Resolved</button>' +
    '</div>' +
    '<div class="count-badge" id="countBadge">Loading...</div>' +
    '</div>' +

    '<div class="grievance-list" id="grievanceList">' +
    '<div class="loading"><div class="spinner"></div><div style="margin-top:15px">Loading grievances...</div></div>' +
    '</div>' +

    // Bottom Navigation
    '<nav class="bottom-nav">' +
    '<a class="nav-item" href="' + baseUrl + '">' +
    '<span class="nav-icon">📊</span>Dashboard</a>' +
    '<a class="nav-item" href="' + baseUrl + '?page=search">' +
    '<span class="nav-icon">🔍</span>Search</a>' +
    '<a class="nav-item active" href="' + baseUrl + '?page=grievances">' +
    '<span class="nav-icon">📋</span>Grievances</a>' +
    '</nav>' +

    '<script>' +
    'var allData=[];' +
    'var currentFilter="all";' +

    'function setFilter(filter,btn){' +
    '  currentFilter=filter;' +
    '  document.querySelectorAll(".filter-pill").forEach(function(p){p.classList.remove("active")});' +
    '  btn.classList.add("active");' +
    '  renderList();' +
    '}' +

    'function getStatusClass(status){' +
    '  if(!status)return"";' +
    '  var s=status.toLowerCase();' +
    '  if(s.indexOf("open")>=0)return"status-open";' +
    '  if(s.indexOf("pending")>=0)return"status-pending";' +
    '  return"status-resolved";' +
    '}' +

    'function matchesFilter(status){' +
    '  if(currentFilter==="all")return true;' +
    '  if(!status)return false;' +
    '  var s=status.toLowerCase();' +
    '  if(currentFilter==="open")return s.indexOf("open")>=0;' +
    '  if(currentFilter==="pending")return s.indexOf("pending")>=0;' +
    '  if(currentFilter==="resolved")return s.indexOf("resolved")>=0||s.indexOf("withdrawn")>=0||s.indexOf("closed")>=0;' +
    '  return true;' +
    '}' +

    'function renderList(){' +
    '  var filtered=allData.filter(function(g){return matchesFilter(g.status)});' +
    '  document.getElementById("countBadge").textContent="Showing "+filtered.length+" of "+allData.length;' +
    '  var c=document.getElementById("grievanceList");' +
    '  if(filtered.length===0){' +
    '    c.innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">📋</div><div>No grievances found</div></div>";' +
    '    return;' +
    '  }' +
    '  c.innerHTML=filtered.map(function(g){' +
    '    return"<div class=\\"grievance-card\\">"+"<div class=\\"grievance-header\\">"+"<span class=\\"grievance-id\\">"+g.id+"</span>"+"<span class=\\"grievance-status "+getStatusClass(g.status)+"\\">"+g.status+"</span>"+"</div>"+"<div class=\\"grievance-name\\">"+g.name+"</div>"+(g.category?"<div class=\\"grievance-detail\\">"+g.category+"</div>":"")+(g.step?"<span class=\\"grievance-step\\">"+g.step+"</span>":"")+"</div>";' +
    '  }).join("");' +
    '}' +

    'function loadData(){' +
    '  google.script.run.withSuccessHandler(function(data){' +
    '    allData=data||[];' +
    '    renderList();' +
    '  }).withFailureHandler(function(err){' +
    '    document.getElementById("grievanceList").innerHTML="<div class=\\"empty-state\\"><div class=\\"empty-icon\\">⚠️</div><div>Error loading data</div></div>";' +
    '  }).getWebAppGrievanceList();' +
    '}' +

    'loadData();' +
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
 * API function to get grievance list for web app
 * @returns {Array} Grievance data
 */
function getWebAppGrievanceList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, GRIEVANCE_COLS.ISSUE_CATEGORY).getValues();

  return data.map(function(row) {
    return {
      id: row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '',
      name: (row[GRIEVANCE_COLS.FIRST_NAME - 1] || '') + ' ' + (row[GRIEVANCE_COLS.LAST_NAME - 1] || ''),
      status: row[GRIEVANCE_COLS.STATUS - 1] || '',
      step: row[GRIEVANCE_COLS.CURRENT_STEP - 1] || '',
      category: row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || ''
    };
  }).filter(function(g) { return g.id; }).slice(0, 100);
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
