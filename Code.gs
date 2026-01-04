/**
 * 509 Dashboard - Main Entry Point
 *
 * Core setup functions, menu system, and sheet creation.
 *
 * ⚠️ WARNING: DO NOT DEPLOY THIS FILE DIRECTLY
 * This is a source file used to generate ConsolidatedDashboard.gs.
 * Deploy ONLY ConsolidatedDashboard.gs to avoid function conflicts.
 *
 * @version 1.0.0
 * @license Free for use by non-profit collective bargaining groups and unions
 */

// ============================================================================
// MENU SYSTEM
// ============================================================================

/**
 * Creates the menu system when the spreadsheet opens
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  // Main Dashboard Menu
  ui.createMenu('👤 Dashboard')
    .addItem('📊 Smart Dashboard (Auto-Detect)', 'showSmartDashboard')
    .addItem('🎯 Custom View', 'showInteractiveDashboardTab')
    .addSeparator()
    .addItem('📋 View Active Grievances', 'viewActiveGrievances')
    .addItem('📱 Mobile Dashboard', 'showMobileDashboard')
    .addItem('📱 Get Mobile App URL', 'showWebAppUrl')
    .addItem('⚡ Quick Actions', 'showQuickActionsMenu')
    .addSeparator()
    .addSubMenu(ui.createMenu('📋 Grievance Tools')
      .addItem('➕ Start New Grievance', 'startNewGrievance')
      .addItem('🔄 Refresh Grievance Formulas', 'recalcAllGrievancesBatched')
      .addItem('🔄 Refresh Member Directory Data', 'refreshMemberDirectoryFormulas')
      .addSeparator()
      .addItem('🔗 Setup Live Grievance Links', 'setupLiveGrievanceFormulas')
      .addItem('👤 Setup Member ID Dropdown', 'setupGrievanceMemberDropdown')
      .addItem('🔧 Fix Overdue Text Data', 'fixOverdueTextToNumbers'))
    .addToUi();

  // Member Search Menu (standalone for quick access)
  ui.createMenu('🔍 Search')
    .addItem('🔍 Search Members', 'searchMembers')
    .addToUi();

  // View Menu - Timeline and display controls
  ui.createMenu('👁️ View')
    .addItem('📅 Simplify Timeline (Hide Steps)', 'simplifyTimelineView')
    .addItem('📅 Show Full Timeline', 'showFullTimelineView')
    .addSeparator()
    .addItem('🎨 Apply Step Highlighting', 'applyStepHighlighting')
    .addItem('🔲 Setup Column Groups', 'setupTimelineColumnGroups')
    .addSeparator()
    .addItem('❄️ Freeze Key Columns', 'freezeKeyColumns')
    .addItem('🔓 Unfreeze All Columns', 'unfreezeAllColumns')
    .addToUi();

  // Sheet Manager Menu
  ui.createMenu('📊 Sheet Manager')
    .addItem('📊 Rebuild Dashboard', 'rebuildDashboard')
    .addItem('📈 Refresh Interactive Charts', 'refreshInteractiveCharts')
    .addItem('🔄 Refresh All Formulas', 'refreshAllFormulas')
    .addSeparator()
    .addSubMenu(ui.createMenu('📁 Google Drive')
      .addItem('📁 Setup Folder for Grievance', 'setupDriveFolderForGrievance')
      .addItem('📁 View Grievance Files', 'showGrievanceFiles')
      .addItem('📁 Batch Create Folders', 'batchCreateGrievanceFolders'))
    .addSubMenu(ui.createMenu('📅 Calendar')
      .addItem('📅 Sync Deadlines to Calendar', 'syncDeadlinesToCalendar')
      .addItem('📅 View Upcoming Deadlines', 'showUpcomingDeadlinesFromCalendar')
      .addItem('🗑️ Clear Calendar Events', 'clearAllCalendarEvents'))
    .addSubMenu(ui.createMenu('📬 Notifications')
      .addItem('⚙️ Notification Settings', 'showNotificationSettings')
      .addItem('⚙️ Alert Settings', 'configureAlertSettings')
      .addSeparator()
      .addItem('📧 Send Steward Alerts Now', 'sendStewardAlertsNow')
      .addItem('🧪 Test Notifications', 'testDeadlineNotifications'))
    .addToUi();

  // Tools Menu (NEW)
  ui.createMenu('🔧 Tools')
    .addSubMenu(ui.createMenu('♿ ADHD & Accessibility')
      .addItem('♿ ADHD Control Panel', 'showADHDControlPanel')
      .addItem('🎯 Focus Mode', 'activateFocusMode')
      .addItem('🔲 Toggle Zebra Stripes', 'toggleZebraStripes')
      .addItem('📝 Quick Capture', 'showQuickCaptureNotepad')
      .addItem('🍅 Pomodoro Timer', 'startPomodoroTimer'))
    .addSubMenu(ui.createMenu('🎨 Theming')
      .addItem('🎨 Theme Manager', 'showThemeManager')
      .addItem('🌙 Toggle Dark Mode', 'quickToggleDarkMode')
      .addItem('🔄 Reset Theme', 'resetToDefaultTheme'))
    .addSeparator()
    .addSubMenu(ui.createMenu('☑️ Multi-Select')
      .addItem('📝 Open Editor', 'showMultiSelectDialog')
      .addSeparator()
      .addItem('⚡ Enable Auto-Open', 'installMultiSelectTrigger')
      .addItem('🚫 Disable Auto-Open', 'removeMultiSelectTrigger'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🗄️ Cache & Performance')
      .addItem('🗄️ Cache Status', 'showCacheStatusDashboard')
      .addItem('🔥 Warm Up Caches', 'warmUpCaches')
      .addItem('🗑️ Clear All Caches', 'invalidateAllCaches'))
    .addSeparator()
    .addSubMenu(ui.createMenu('✅ Validation')
      .addItem('🔍 Run Bulk Validation', 'runBulkValidation')
      .addItem('⚙️ Validation Settings', 'showValidationSettings')
      .addItem('🧹 Clear Indicators', 'clearValidationIndicators')
      .addItem('⚡ Install Validation Trigger', 'installValidationTrigger'))
    .addToUi();

  // Setup Menu
  ui.createMenu('🏗️ Setup')
    .addItem('🔧 REPAIR DASHBOARD', 'REPAIR_DASHBOARD')
    .addSeparator()
    .addItem('⚙️ Setup Data Validations', 'setupDataValidations')
    .addItem('🎨 Setup ADHD Defaults', 'setupADHDDefaults')
    .addItem('↩️ Undo ADHD Defaults', 'undoADHDDefaults')
    .addToUi();

  // Demo Menu - only show if demo mode is not disabled
  if (!isDemoModeDisabled()) {
    ui.createMenu('🎭 Demo')
      .addItem('🚀 Seed All Sample Data', 'SEED_SAMPLE_DATA')
      .addSeparator()
      .addSubMenu(ui.createMenu('🗑️ Nuke Data')
        .addItem('☢️ NUKE SEEDED DATA', 'NUKE_SEEDED_DATA')
        .addItem('🧹 Clear Config Dropdowns Only', 'NUKE_CONFIG_DROPDOWNS')
        .addSeparator()
        .addItem('🔄 Restore Config & Dropdowns', 'restoreConfigAndDropdowns'))
      .addToUi();
  }

  // Testing Menu (NEW)
  ui.createMenu('🧪 Testing')
    .addItem('🧪 Run All Tests', 'runAllTests')
    .addItem('⚡ Run Quick Tests', 'runQuickTests')
    .addSeparator()
    .addItem('📊 View Test Results', 'viewTestResults')
    .addToUi();

  // Administrator Menu
  ui.createMenu('⚙️ Administrator')
    .addItem('🔍 DIAGNOSE SETUP', 'DIAGNOSE_SETUP')
    .addItem('🔍 Verify Hidden Sheets', 'verifyHiddenSheets')
    .addSeparator()
    .addSubMenu(ui.createMenu('🔧 Setup & Triggers')
      .addItem('🔧 Setup All Hidden Sheets', 'setupAllHiddenSheets')
      .addItem('🔧 Repair All Hidden Sheets', 'repairAllHiddenSheets')
      .addItem('⚡ Install Auto-Sync Trigger', 'installAutoSyncTrigger')
      .addItem('🚫 Remove Auto-Sync Trigger', 'removeAutoSyncTrigger'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔄 Manual Sync')
      .addItem('🔄 Sync All Data Now', 'syncAllData')
      .addItem('🔄 Sync Grievance → Members', 'syncGrievanceToMemberDirectory')
      .addItem('🔄 Sync Members → Grievances', 'syncMemberToGrievanceLog'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📋 Audit Log')
      .addItem('📋 View Audit Log', 'viewAuditLog')
      .addItem('🔧 Setup Audit Log', 'setupAuditLogSheet')
      .addSeparator()
      .addItem('⚡ Enable Audit Tracking', 'installAuditTrigger')
      .addItem('🚫 Disable Audit Tracking', 'removeAuditTrigger')
      .addSeparator()
      .addItem('🗑️ Clear Old Entries (30+ days)', 'clearOldAuditEntries'))
    .addSubMenu(ui.createMenu('🩺 Data Quality')
      .addItem('🔍 Check Data Quality', 'fixDataQualityIssues')
      .addItem('📋 View Missing Member IDs', 'showGrievancesWithMissingMemberIds'))
    .addToUi();
}

// ============================================================================
// MAIN SETUP FUNCTION
// ============================================================================

/**
 * Main setup function - creates the complete 509 Dashboard
 * Creates the core sheets with proper structure and formatting
 */
function CREATE_509_DASHBOARD() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // Confirm with user
  var response = ui.alert(
    '🏗️ Create 509 Dashboard',
    'This will create the 509 Dashboard with:\n\n' +
    '• Config (dropdown sources)\n' +
    '• Member Directory\n' +
    '• Grievance Log\n' +
    '• 💼 Dashboard (Executive metrics)\n' +
    '• 🎯 Interactive (Customizable view)\n\n' +
    'Plus 6 hidden calculation sheets for self-healing formulas.\n\n' +
    'Existing sheets with matching names will be recreated.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Setup cancelled.');
    return;
  }

  ss.toast('Starting dashboard creation...', '🏗️ Setup', 5);

  try {
    // Create core data sheets
    createConfigSheet(ss);
    ss.toast('Created Config sheet', '🏗️ Progress', 2);

    createMemberDirectory(ss);
    ss.toast('Created Member Directory', '🏗️ Progress', 2);

    createGrievanceLog(ss);
    ss.toast('Created Grievance Log', '🏗️ Progress', 2);

    // Setup hidden calculation sheets (needed before dashboards for formula references)
    ss.toast('Setting up hidden sheets...', '🏗️ Progress', 3);
    setupHiddenSheets(ss);

    // Create dashboard sheets (after hidden sheets so formulas can reference them)
    createDashboard(ss);
    ss.toast('Created Dashboard', '🏗️ Progress', 2);

    createInteractiveDashboard(ss);
    ss.toast('Created Custom View', '🏗️ Progress', 2);

    // Setup data validations
    ss.toast('Setting up validations...', '🏗️ Progress', 3);
    setupDataValidations();

    // Move Config to first position
    var configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (configSheet) {
      ss.setActiveSheet(configSheet);
      ss.moveActiveSheet(1);
    }

    ss.toast('Dashboard creation complete!', '✅ Success', 5);
    ui.alert('✅ Success', '509 Dashboard has been created successfully!\n\n' +
      '5 sheets created:\n' +
      '• Config, Member Directory, Grievance Log (data)\n' +
      '• 💼 Dashboard, 🎯 Interactive (views)\n\n' +
      'Plus 6 hidden calculation sheets with self-healing formulas.\n\n' +
      '⚡ Auto-sync trigger installed - dates and deadlines will\n' +
      'update automatically when you edit the sheets.\n\n' +
      'Use the Demo menu to seed sample data.', ui.ButtonSet.OK);

  } catch (error) {
    Logger.log('Error in CREATE_509_DASHBOARD: ' + error.message);
    ui.alert('❌ Error', 'An error occurred: ' + error.message, ui.ButtonSet.OK);
  }
}

// ============================================================================
// SHEET CREATION FUNCTIONS
// ============================================================================

/**
 * Create or recreate the Config sheet with dropdown values
 * Comprehensive configuration with section groupings and organization settings
 */
function createConfigSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.CONFIG);
  sheet.clear();

  // Row 1: Section Headers (grouped categories)
  var sectionHeaders = [
    '── EMPLOYMENT INFO ──', '', '', '', '',           // A-E (5 cols)
    '── SUPERVISION ──', '',                            // F-G (2 cols)
    '── STEWARD INFO ──', '',                           // H-I (2 cols)
    '── GRIEVANCE SETTINGS ──', '', '', '',             // J-M (4 cols)
    '── LINKS & COORDINATORS ──', '', '', '',           // N-Q (4 cols)
    '── NOTIFICATIONS ──', '', '',                      // R-T (3 cols)
    '── ORGANIZATION ──', '', '', '',                   // U-X (4 cols)
    '── INTEGRATION ──', '',                            // Y-Z (2 cols)
    '── DEADLINES ──', '', '', '',                      // AA-AD (4 cols)
    '── MULTI-SELECT OPTIONS ──', '',                   // AE-AF (2 cols)
    '── CONTRACT & LEGAL ──', '', '', '',               // AG-AJ (4 cols)
    '── ORG IDENTITY ──', '', '',                       // AK-AM (3 cols)
    '── EXTENDED CONTACT ──', '', '', ''                // AN-AQ (4 cols)
  ];

  // Row 2: Column Headers
  var columnHeaders = [
    'Job Titles', 'Office Locations', 'Units', 'Office Days', 'Yes/No (Dropdowns)',
    'Supervisors', 'Managers',
    'Stewards', 'Steward Committees',
    'Grievance Status', 'Grievance Step', 'Issue Category', 'Articles Violated',
    'Communication Methods', 'Grievance Coordinators', 'Grievance Form URL', 'Contact Form URL',
    'Admin Emails', 'Alert Days Before Deadline', 'Notification Recipients',
    'Organization Name', 'Local Number', 'Main Office Address', 'Main Phone',
    'Google Drive Folder ID', 'Google Calendar ID',
    'Filing Deadline Days', 'Step I Response Days', 'Step II Appeal Days', 'Step II Response Days',
    'Best Times to Contact', 'Home Towns',
    'Contract Article (Grievance)', 'Contract Article (Discipline)', 'Contract Article (Workload)', 'Contract Name',
    'Union Parent', 'State/Region', 'Organization Website',
    'Office Addresses', 'Main Fax', 'Main Contact Name', 'Main Contact Email'
  ];

  // Apply section headers (Row 1)
  sheet.getRange(1, 1, 1, sectionHeaders.length).setValues([sectionHeaders])
    .setBackground(COLORS.LIGHT_GRAY)
    .setFontColor(COLORS.TEXT_DARK)
    .setFontWeight('bold')
    .setFontStyle('italic')
    .setHorizontalAlignment('center');

  // Apply column headers (Row 2)
  sheet.getRange(2, 1, 1, columnHeaders.length).setValues([columnHeaders])
    .setBackground(COLORS.PRIMARY_PURPLE)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Add default dropdown values (Row 3+)
  // Office Days (D)
  sheet.getRange(3, CONFIG_COLS.OFFICE_DAYS, DEFAULT_CONFIG.OFFICE_DAYS.length, 1)
    .setValues(DEFAULT_CONFIG.OFFICE_DAYS.map(function(v) { return [v]; }));

  // Yes/No (E)
  sheet.getRange(3, CONFIG_COLS.YES_NO, DEFAULT_CONFIG.YES_NO.length, 1)
    .setValues(DEFAULT_CONFIG.YES_NO.map(function(v) { return [v]; }));

  // Steward Committees (I)
  var committees = ['Grievance Committee', 'Bargaining Committee', 'Health & Safety Committee',
                    'Political Action Committee', 'Membership Committee', 'Executive Board'];
  sheet.getRange(3, CONFIG_COLS.STEWARD_COMMITTEES, committees.length, 1)
    .setValues(committees.map(function(v) { return [v]; }));

  // Grievance Status (J)
  sheet.getRange(3, CONFIG_COLS.GRIEVANCE_STATUS, DEFAULT_CONFIG.GRIEVANCE_STATUS.length, 1)
    .setValues(DEFAULT_CONFIG.GRIEVANCE_STATUS.map(function(v) { return [v]; }));

  // Grievance Step (K)
  sheet.getRange(3, CONFIG_COLS.GRIEVANCE_STEP, DEFAULT_CONFIG.GRIEVANCE_STEP.length, 1)
    .setValues(DEFAULT_CONFIG.GRIEVANCE_STEP.map(function(v) { return [v]; }));

  // Issue Category (L)
  sheet.getRange(3, CONFIG_COLS.ISSUE_CATEGORY, DEFAULT_CONFIG.ISSUE_CATEGORY.length, 1)
    .setValues(DEFAULT_CONFIG.ISSUE_CATEGORY.map(function(v) { return [v]; }));

  // Articles (M)
  sheet.getRange(3, CONFIG_COLS.ARTICLES, DEFAULT_CONFIG.ARTICLES.length, 1)
    .setValues(DEFAULT_CONFIG.ARTICLES.map(function(v) { return [v]; }));

  // Communication Methods (N)
  sheet.getRange(3, CONFIG_COLS.COMM_METHODS, DEFAULT_CONFIG.COMM_METHODS.length, 1)
    .setValues(DEFAULT_CONFIG.COMM_METHODS.map(function(v) { return [v]; }));

  // Alert Days (S) - default notification intervals
  sheet.getRange(3, CONFIG_COLS.ALERT_DAYS, 1, 1).setValue('3, 7, 14');

  // Organization defaults (SEIU Local 509)
  sheet.getRange(3, CONFIG_COLS.ORG_NAME, 1, 1).setValue('SEIU Local 509');
  sheet.getRange(3, CONFIG_COLS.LOCAL_NUMBER, 1, 1).setValue('509');
  sheet.getRange(3, CONFIG_COLS.MAIN_ADDRESS, 1, 1).setValue('293 Boston Post Road West, 4th Floor, Marlborough, MA 01752');
  sheet.getRange(3, CONFIG_COLS.MAIN_PHONE, 1, 1).setValue('774-843-7509');

  // Deadline defaults (in days)
  sheet.getRange(3, CONFIG_COLS.FILING_DEADLINE_DAYS, 1, 1).setValue(21);
  sheet.getRange(3, CONFIG_COLS.STEP1_RESPONSE_DAYS, 1, 1).setValue(30);
  sheet.getRange(3, CONFIG_COLS.STEP2_APPEAL_DAYS, 1, 1).setValue(10);
  sheet.getRange(3, CONFIG_COLS.STEP2_RESPONSE_DAYS, 1, 1).setValue(30);

  // Best Times to Contact (AE)
  var bestTimes = ['Morning (8am-12pm)', 'Afternoon (12pm-5pm)', 'Evening (5pm-8pm)', 'Weekends', 'Flexible'];
  sheet.getRange(3, CONFIG_COLS.BEST_TIMES, bestTimes.length, 1)
    .setValues(bestTimes.map(function(v) { return [v]; }));

  // Contract articles
  sheet.getRange(3, CONFIG_COLS.CONTRACT_GRIEVANCE, 1, 1).setValue('Article 23A');
  sheet.getRange(3, CONFIG_COLS.CONTRACT_DISCIPLINE, 1, 1).setValue('Article 12');
  sheet.getRange(3, CONFIG_COLS.CONTRACT_WORKLOAD, 1, 1).setValue('Article 15');
  sheet.getRange(3, CONFIG_COLS.CONTRACT_NAME, 1, 1).setValue('2023-2026 CBA');

  // Org identity
  sheet.getRange(3, CONFIG_COLS.UNION_PARENT, 1, 1).setValue('SEIU');
  sheet.getRange(3, CONFIG_COLS.STATE_REGION, 1, 1).setValue('Massachusetts');
  sheet.getRange(3, CONFIG_COLS.ORG_WEBSITE, 1, 1).setValue('https://www.seiu509.org/');

  // Extended contact
  sheet.getRange(3, CONFIG_COLS.MAIN_FAX, 1, 1).setValue('508-485-8529');
  sheet.getRange(3, CONFIG_COLS.MAIN_CONTACT_NAME, 1, 1).setValue('Marc');
  sheet.getRange(3, CONFIG_COLS.MAIN_CONTACT_EMAIL, 1, 1).setValue('marc@seiu509.org');

  // Freeze header rows (1 and 2)
  sheet.setFrozenRows(2);

  // Auto-resize all columns
  sheet.autoResizeColumns(1, columnHeaders.length);

  // Set minimum column widths for readability
  for (var i = 1; i <= columnHeaders.length; i++) {
    if (sheet.getColumnWidth(i) < 100) {
      sheet.setColumnWidth(i, 100);
    }
  }
}

/**
 * Create or recreate the Member Directory sheet
 */
function createMemberDirectory(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.MEMBER_DIR);
  sheet.clear();

  var headers = getMemberHeaders();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(COLORS.PRIMARY_PURPLE)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold');

  // Freeze header row
  sheet.setFrozenRows(1);

  // Set column widths
  sheet.setColumnWidth(MEMBER_COLS.MEMBER_ID, 100);
  sheet.setColumnWidth(MEMBER_COLS.FIRST_NAME, 120);
  sheet.setColumnWidth(MEMBER_COLS.LAST_NAME, 120);
  sheet.setColumnWidth(MEMBER_COLS.EMAIL, 200);
  sheet.setColumnWidth(MEMBER_COLS.CONTACT_NOTES, 250);

  // Add checkbox for Start Grievance column (pre-allocate for future rows)
  sheet.getRange(2, MEMBER_COLS.START_GRIEVANCE, 4999, 1).insertCheckboxes();

  // Format date columns (MM/dd/yyyy)
  var dateColumns = [
    MEMBER_COLS.LAST_VIRTUAL_MTG,
    MEMBER_COLS.LAST_INPERSON_MTG,
    MEMBER_COLS.RECENT_CONTACT_DATE
  ];
  dateColumns.forEach(function(col) {
    sheet.getRange(2, col, 998, 1).setNumberFormat('MM/dd/yyyy');
  });

  // Format numeric columns with comma separators
  sheet.getRange(2, MEMBER_COLS.OPEN_RATE, 998, 1).setNumberFormat('#,##0.0');       // S - Open Rate %
  sheet.getRange(2, MEMBER_COLS.VOLUNTEER_HOURS, 998, 1).setNumberFormat('#,##0');  // T - Volunteer Hours

  // Columns AB-AD (Has Open Grievance?, Grievance Status, Days to Deadline)
  // are populated by syncGrievanceToMemberDirectory() with STATIC values
  // No formulas in visible sheets - all calculations done by script

  // Auto-resize other columns
  sheet.autoResizeColumns(1, headers.length);

  // ═══════════════════════════════════════════════════════════════════════════
  // COLUMN GROUPS: Group and hide optional columns for cleaner view
  // ═══════════════════════════════════════════════════════════════════════════
  try {
    // Group 1: Engagement Metrics (Q-T, columns 17-20) - Hidden by default
    sheet.getRange(1, MEMBER_COLS.LAST_VIRTUAL_MTG, sheet.getMaxRows(), 4).shiftColumnGroupDepth(1);
    sheet.collapseAllColumnGroups();

    // Group 2: Member Interests (U-X, columns 21-24) - Hidden by default
    sheet.getRange(1, MEMBER_COLS.INTEREST_LOCAL, sheet.getMaxRows(), 4).shiftColumnGroupDepth(1);
    sheet.collapseAllColumnGroups();

    sheet.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.AFTER);
  } catch (e) {
    Logger.log('Member Directory column group setup skipped: ' + e.toString());
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // CONDITIONAL FORMATTING: Highlight members with open grievances
  // ═══════════════════════════════════════════════════════════════════════════
  var lastRow = Math.max(sheet.getLastRow(), 2);
  var hasOpenGrievanceRange = sheet.getRange(2, MEMBER_COLS.HAS_OPEN_GRIEVANCE, 4999, 1);

  var redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Yes')
    .setBackground('#ffebee')  // Light red background
    .setFontColor('#c62828')   // Dark red text
    .setBold(true)
    .setRanges([hasOpenGrievanceRange])
    .build();

  var rules = sheet.getConditionalFormatRules();
  rules.push(redRule);
  sheet.setConditionalFormatRules(rules);
}

/**
 * Create or recreate the Grievance Log sheet
 * NOTE: Calculated columns (First Name, Last Name, Email, Deadlines, Days Open, etc.)
 * are managed by the hidden _Grievance_Formulas sheet for self-healing capability.
 * Users can't accidentally erase formulas because they're in the hidden sheet.
 */
function createGrievanceLog(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.GRIEVANCE_LOG);
  sheet.clear();

  var headers = getGrievanceHeaders();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(COLORS.PRIMARY_PURPLE)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold');

  // Freeze header row
  sheet.setFrozenRows(1);

  // Set column widths
  sheet.setColumnWidth(GRIEVANCE_COLS.GRIEVANCE_ID, 100);
  sheet.setColumnWidth(GRIEVANCE_COLS.RESOLUTION, 250);
  sheet.setColumnWidth(GRIEVANCE_COLS.COORDINATOR_MESSAGE, 250);

  // Add checkbox for Message Alert column (pre-allocate for future rows)
  sheet.getRange(2, GRIEVANCE_COLS.MESSAGE_ALERT, 4999, 1).insertCheckboxes();

  // Format date columns
  var dateColumns = [
    GRIEVANCE_COLS.INCIDENT_DATE,
    GRIEVANCE_COLS.FILING_DEADLINE,
    GRIEVANCE_COLS.DATE_FILED,
    GRIEVANCE_COLS.STEP1_DUE,
    GRIEVANCE_COLS.STEP1_RCVD,
    GRIEVANCE_COLS.STEP2_APPEAL_DUE,
    GRIEVANCE_COLS.STEP2_APPEAL_FILED,
    GRIEVANCE_COLS.STEP2_DUE,
    GRIEVANCE_COLS.STEP2_RCVD,
    GRIEVANCE_COLS.STEP3_APPEAL_DUE,
    GRIEVANCE_COLS.STEP3_APPEAL_FILED,
    GRIEVANCE_COLS.DATE_CLOSED,
    GRIEVANCE_COLS.NEXT_ACTION_DUE
  ];

  dateColumns.forEach(function(col) {
    sheet.getRange(2, col, 998, 1).setNumberFormat('MM/dd/yyyy');
  });

  // Format Days Open (S) and Days to Deadline (U) as whole numbers with comma separators
  sheet.getRange(2, GRIEVANCE_COLS.DAYS_OPEN, 998, 1).setNumberFormat('#,##0');
  // Days to Deadline can show "Overdue" text, so use General format that handles both
  sheet.getRange(2, GRIEVANCE_COLS.DAYS_TO_DEADLINE, 998, 1).setNumberFormat('#,##0');

  // Auto-resize other columns
  sheet.autoResizeColumns(1, headers.length);

  // Setup column groups for timeline (Step I, II, III collapsible)
  try {
    sheet.getRange(1, GRIEVANCE_COLS.STEP1_DUE, sheet.getMaxRows(), 2).shiftColumnGroupDepth(1);
    sheet.getRange(1, GRIEVANCE_COLS.STEP2_APPEAL_DUE, sheet.getMaxRows(), 4).shiftColumnGroupDepth(1);
    sheet.getRange(1, GRIEVANCE_COLS.STEP3_APPEAL_DUE, sheet.getMaxRows(), 2).shiftColumnGroupDepth(1);
    sheet.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.AFTER);
  } catch (e) {
    Logger.log('Column group setup skipped: ' + e.toString());
  }
}


/**
 * Create or recreate the unified Dashboard sheet (Executive Dashboard theme)
 * Combines member metrics, grievance metrics, and key performance indicators
 * Uses Executive Dashboard's green QUICK STATS theme
 */
function createDashboard(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.DASHBOARD);
  sheet.clear();

  // Title - Executive Dashboard style
  sheet.getRange('A1').setValue('💼 Executive Dashboard')
    .setFontSize(24)
    .setFontWeight('bold')
    .setFontColor(COLORS.PRIMARY_PURPLE);
  sheet.getRange('A1:F1').merge();

  // Subtitle with last refresh
  sheet.getRange('A2').setValue('Real-time metrics from Member Directory & Grievance Log')
    .setFontSize(10)
    .setFontStyle('italic')
    .setFontColor(COLORS.TEXT_DARK);
  sheet.getRange('A2:F2').merge();

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 1: QUICK STATS (Executive Dashboard green theme)
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A4').setValue('QUICK STATS')
    .setFontWeight('bold')
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A4:F4').merge();

  // Quick stats row - pulls from hidden _Dashboard_Calc sheet
  var quickStatsLabels = [
    ['Total Members', 'Active Stewards', 'Active Grievances', 'Win Rate', 'Overdue Cases', 'Due This Week']
  ];
  sheet.getRange('A5:F5').setValues(quickStatsLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  // Quick stats formulas - reference hidden calculation sheet
  var quickStatsFormulas = [
    [
      '=IFERROR(\'' + SHEETS.DASHBOARD_CALC + '\'!B2,0)',
      '=IFERROR(\'' + SHEETS.DASHBOARD_CALC + '\'!B3,0)',
      '=IFERROR(\'' + SHEETS.DASHBOARD_CALC + '\'!B5+\'' + SHEETS.DASHBOARD_CALC + '\'!B6,0)',
      '=IFERROR(TEXT(\'' + SHEETS.DASHBOARD_CALC + '\'!B11/100,"0%"),"-")',
      '=IFERROR(\'' + SHEETS.DASHBOARD_CALC + '\'!B13,0)',
      '=IFERROR(\'' + SHEETS.DASHBOARD_CALC + '\'!B14,0)'
    ]
  ];
  sheet.getRange('A6:F6').setFormulas(quickStatsFormulas)
    .setFontSize(20)
    .setHorizontalAlignment('center')
    .setFontWeight('bold');

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 2: MEMBER METRICS
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A8').setValue('MEMBER METRICS')
    .setFontWeight('bold')
    .setBackground(COLORS.PRIMARY_BLUE)
    .setFontColor(COLORS.TEXT_DARK);
  sheet.getRange('A8:D8').merge();

  var memberMetricLabels = [['Total Members', 'Active Stewards', 'Avg Open Rate', 'YTD Vol. Hours']];
  sheet.getRange('A9:D9').setValues(memberMetricLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  var mIdCol = getColumnLetter(MEMBER_COLS.MEMBER_ID);
  var mStewardCol = getColumnLetter(MEMBER_COLS.IS_STEWARD);
  var mOpenRateCol = getColumnLetter(MEMBER_COLS.OPEN_RATE);
  var mVolHoursCol = getColumnLetter(MEMBER_COLS.VOLUNTEER_HOURS);

  var memberMetricFormulas = [
    [
      '=COUNTA(\'' + SHEETS.MEMBER_DIR + '\'!' + mIdCol + ':' + mIdCol + ')-1',
      '=COUNTIF(\'' + SHEETS.MEMBER_DIR + '\'!' + mStewardCol + ':' + mStewardCol + ',"Yes")',
      '=IFERROR(ROUND(AVERAGE(\'' + SHEETS.MEMBER_DIR + '\'!' + mOpenRateCol + ':' + mOpenRateCol + '),1)&"%","-")',
      '=SUM(\'' + SHEETS.MEMBER_DIR + '\'!' + mVolHoursCol + ':' + mVolHoursCol + ')'
    ]
  ];
  sheet.getRange('A10:D10').setFormulas(memberMetricFormulas)
    .setFontSize(18)
    .setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 3: GRIEVANCE METRICS
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A12').setValue('GRIEVANCE METRICS')
    .setFontWeight('bold')
    .setBackground(COLORS.ACCENT_ORANGE)
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A12:F12').merge();

  var grievanceLabels = [['Open', 'Pending Info', 'Settled', 'Won', 'Denied', 'Withdrawn']];
  sheet.getRange('A13:F13').setValues(grievanceLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  var gStatusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);
  var gResolutionCol = getColumnLetter(GRIEVANCE_COLS.RESOLUTION);

  var grievanceFormulas = [
    [
      '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Open")',
      '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Pending Info")',
      '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Settled")',
      '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Won")',
      '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Denied")',
      '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Withdrawn")'
    ]
  ];
  sheet.getRange('A14:F14').setFormulas(grievanceFormulas)
    .setFontSize(18)
    .setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 4: TIMELINE METRICS
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A16').setValue('TIMELINE & PERFORMANCE')
    .setFontWeight('bold')
    .setBackground(COLORS.PRIMARY_PURPLE)
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A16:D16').merge();

  var timelineLabels = [['Avg Days Open', 'Filed This Month', 'Closed This Month', 'Avg Resolution Days']];
  sheet.getRange('A17:D17').setValues(timelineLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  var gDaysOpenCol = getColumnLetter(GRIEVANCE_COLS.DAYS_OPEN);
  var gDateFiledCol = getColumnLetter(GRIEVANCE_COLS.DATE_FILED);
  var gDateClosedCol = getColumnLetter(GRIEVANCE_COLS.DATE_CLOSED);

  var timelineFormulas = [
    [
      '=IFERROR(ROUND(AVERAGE(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysOpenCol + ':' + gDaysOpenCol + '),1),0)',
      '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateFiledCol + ':' + gDateFiledCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateFiledCol + ':' + gDateFiledCol + ',"<="&TODAY())',
      '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',"<="&TODAY())',
      '=IFERROR(ROUND(AVERAGEIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',"<>",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysOpenCol + ':' + gDaysOpenCol + '),1),0)'
    ]
  ];
  sheet.getRange('A18:D18').setFormulas(timelineFormulas)
    .setFontSize(18)
    .setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 5: TYPE ANALYSIS (by Issue Category)
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A20').setValue('📊 TYPE ANALYSIS (by Issue Category)')
    .setFontWeight('bold')
    .setBackground('#6366F1')  // Indigo
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A20:F20').merge();

  var typeLabels = [['Issue Category', 'Total Cases', 'Open', 'Resolved', 'Win Rate', 'Avg Days']];
  sheet.getRange('A21:F21').setValues(typeLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  var gIssueCatCol = getColumnLetter(GRIEVANCE_COLS.ISSUE_CATEGORY);
  var gIdCol = getColumnLetter(GRIEVANCE_COLS.GRIEVANCE_ID);

  // Top 5 issue categories with metrics - using QUERY formulas
  var categoryFormulas = [];
  var defaultCategories = ['Contract Violation', 'Discipline', 'Workload', 'Safety', 'Discrimination'];
  for (var c = 0; c < 5; c++) {
    var cat = defaultCategories[c];
    categoryFormulas.push([
      '="' + cat + '"',  // Wrap text as formula for setFormulas()
      '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gIssueCatCol + ':' + gIssueCatCol + ',"' + cat + '")',
      '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gIssueCatCol + ':' + gIssueCatCol + ',"' + cat + '",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Open")',
      '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gIssueCatCol + ':' + gIssueCatCol + ',"' + cat + '",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"<>Open")-COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gIssueCatCol + ':' + gIssueCatCol + ',"' + cat + '",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Pending Info")',
      '=IFERROR(TEXT(COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gIssueCatCol + ':' + gIssueCatCol + ',"' + cat + '",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Won")/COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gIssueCatCol + ':' + gIssueCatCol + ',"' + cat + '"),"0%"),"-")',
      '=IFERROR(ROUND(AVERAGEIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysOpenCol + ':' + gDaysOpenCol + ',\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gIssueCatCol + ':' + gIssueCatCol + ',"' + cat + '"),1),"-")'
    ]);
  }
  sheet.getRange('A22:F26').setFormulas(categoryFormulas)
    .setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 6: LOCATION BREAKDOWN
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A28').setValue('🗺️ LOCATION BREAKDOWN')
    .setFontWeight('bold')
    .setBackground('#0891B2')  // Cyan
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A28:F28').merge();

  var locationLabels = [['Location', 'Members', 'Grievances', 'Open Cases', 'Win Rate', 'Avg Satisfaction']];
  sheet.getRange('A29:F29').setValues(locationLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  var mLocationCol = getColumnLetter(MEMBER_COLS.WORK_LOCATION);
  var gLocationCol = getColumnLetter(GRIEVANCE_COLS.LOCATION);  // GRIEVANCE_COLS uses LOCATION not WORK_LOCATION
  var configLocCol = getColumnLetter(CONFIG_COLS.OFFICE_LOCATIONS);

  // Location formulas - pulls top 5 locations from Config and calculates metrics
  var locationFormulas = [];
  for (var loc = 0; loc < 5; loc++) {
    var configRow = 3 + loc;  // Config data starts at row 3
    var locRef = '\'' + SHEETS.CONFIG + '\'!' + configLocCol + configRow;
    locationFormulas.push([
      '=IFERROR(' + locRef + ',"")',
      '=IF(' + locRef + '<>"",COUNTIF(\'' + SHEETS.MEMBER_DIR + '\'!' + mLocationCol + ':' + mLocationCol + ',' + locRef + '),"")',
      '=IF(' + locRef + '<>"",COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gLocationCol + ':' + gLocationCol + ',' + locRef + '),"")',
      '=IF(' + locRef + '<>"",COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gLocationCol + ':' + gLocationCol + ',' + locRef + ',\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Open"),"")',
      '=IF(AND(' + locRef + '<>"",COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gLocationCol + ':' + gLocationCol + ',' + locRef + ')>0),TEXT(COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gLocationCol + ':' + gLocationCol + ',' + locRef + ',\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Won")/COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gLocationCol + ':' + gLocationCol + ',' + locRef + '),"0%"),"-")',
      '="-"'  // Satisfaction requires separate tracking
    ]);
  }
  sheet.getRange('A30:F34').setFormulas(locationFormulas)
    .setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 7: MONTH-OVER-MONTH TRENDS (moved up from rows 40-44)
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A36').setValue('📈 MONTH-OVER-MONTH TRENDS')
    .setFontWeight('bold')
    .setBackground('#DC2626')  // Red
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A36:F36').merge();

  var trendLabels = [['Metric', 'This Month', 'Last Month', 'Change', '% Change', 'Trend']];
  sheet.getRange('A37:F37').setValues(trendLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  // Trend rows: Filed, Closed, Won
  var trendData = [
    [
      '="Grievances Filed"',
      '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateFiledCol + ':' + gDateFiledCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateFiledCol + ':' + gDateFiledCol + ',"<="&TODAY())',
      '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateFiledCol + ':' + gDateFiledCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY())-1,1),\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateFiledCol + ':' + gDateFiledCol + ',"<"&DATE(YEAR(TODAY()),MONTH(TODAY()),1))',
      '=B38-C38',
      '=IFERROR(TEXT((B38-C38)/C38,"0%"),"-")',
      '=IF(B38>C38,"📈",IF(B38<C38,"📉","➡️"))'
    ],
    [
      '="Grievances Closed"',
      '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',"<="&TODAY())',
      '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY())-1,1),\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',"<"&DATE(YEAR(TODAY()),MONTH(TODAY()),1))',
      '=B39-C39',
      '=IFERROR(TEXT((B39-C39)/C39,"0%"),"-")',
      '=IF(B39>C39,"📈",IF(B39<C39,"📉","➡️"))'
    ],
    [
      '="Cases Won"',
      '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Won")',
      '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY())-1,1),\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',"<"&DATE(YEAR(TODAY()),MONTH(TODAY()),1),\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Won")',
      '=B40-C40',
      '=IFERROR(TEXT((B40-C40)/C40,"0%"),"-")',
      '=IF(B40>C40,"📈",IF(B40<C40,"📉","➡️"))'
    ]
  ];
  sheet.getRange('A38:F40').setFormulas(trendData)
    .setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 8: STATUS LEGEND (compact - fits in A-F only)
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A42').setValue('STATUS LEGEND')
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange('A42:F42').merge();

  // Deadline status - compact version
  var legendDeadline = [
    ['🟢 >7d', '🟡 4-7d', '🟠 1-3d', '🔴 Overdue', '✅ Won', '❌ Denied']
  ];
  sheet.getRange('A43:F43').setValues(legendDeadline)
    .setHorizontalAlignment('center')
    .setFontSize(9);

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 9: STEWARD PERFORMANCE SUMMARY (moved to near bottom)
  // ═══════════════════════════════════════════════════════════════════════════
  var mStewardCol = getColumnLetter(MEMBER_COLS.IS_STEWARD);
  var gAssignedStewardCol = getColumnLetter(GRIEVANCE_COLS.STEWARD);
  var mContactDateCol = getColumnLetter(MEMBER_COLS.RECENT_CONTACT_DATE);

  sheet.getRange('A45').setValue('👨‍⚖️ STEWARD PERFORMANCE SUMMARY')
    .setFontWeight('bold')
    .setBackground('#7C3AED')  // Purple
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A45:F45').merge();

  var stewardLabels = [['Total Stewards', 'Active (w/ Cases)', 'Avg Cases/Steward', 'Total Vol Hours', 'Contacts This Month', '']];
  sheet.getRange('A46:F46').setValues(stewardLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  var stewardFormulas = [
    [
      '=COUNTIF(\'' + SHEETS.MEMBER_DIR + '\'!' + mStewardCol + ':' + mStewardCol + ',"Yes")',
      '=IFERROR(ROWS(UNIQUE(FILTER(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gAssignedStewardCol + ':' + gAssignedStewardCol + ',\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gAssignedStewardCol + ':' + gAssignedStewardCol + '<>"",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + '="Open"))),0)',
      '=IFERROR(ROUND((COUNTA(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gIdCol + ':' + gIdCol + ')-1)/COUNTIF(\'' + SHEETS.MEMBER_DIR + '\'!' + mStewardCol + ':' + mStewardCol + ',"Yes"),1),"-")',
      '=SUM(\'' + SHEETS.MEMBER_DIR + '\'!' + mVolHoursCol + ':' + mVolHoursCol + ')',
      '=COUNTIFS(\'' + SHEETS.MEMBER_DIR + '\'!' + mContactDateCol + ':' + mContactDateCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),\'' + SHEETS.MEMBER_DIR + '\'!' + mContactDateCol + ':' + mContactDateCol + ',"<="&TODAY())',
      '=""'
    ]
  ];
  sheet.getRange('A47:F47').setFormulas(stewardFormulas)
    .setFontSize(16)
    .setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 10: TOP 30 BUSIEST STEWARDS (by active grievance count)
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A49').setValue('📊 TOP 30 BUSIEST STEWARDS (Active Cases)')
    .setFontWeight('bold')
    .setBackground('#B91C1C')  // Dark red
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A49:F49').merge();

  var busiestLabels = [['Rank', 'Steward Name', 'Active Cases', 'Open', 'Pending Info', 'Total Ever']];
  sheet.getRange('A50:F50').setValues(busiestLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  // Use QUERY to get top 30 stewards sorted by active case count (Open + Pending Info)
  // This properly sorts by workload, not alphabetically
  var queryFormula = '=IFERROR(QUERY({' +
    '\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gAssignedStewardCol + ':' + gAssignedStewardCol + ',' +
    '\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + '},' +
    '"SELECT Col1, COUNT(Col1) WHERE Col1 <> \'\' AND Col1 <> \'Assigned Steward\' AND (Col2 = \'Open\' OR Col2 = \'Pending Info\') ' +
    'GROUP BY Col1 ORDER BY COUNT(Col1) DESC LIMIT 30 LABEL COUNT(Col1) \'\'",' +
    '0),{"",""})';

  // Place QUERY result starting at B51 - this returns steward names and their active counts
  sheet.getRange('B51').setFormula(queryFormula);

  // Generate rank numbers and additional metrics for rows 51-80
  for (var rank = 1; rank <= 30; rank++) {
    var row = 50 + rank;
    // Rank number
    sheet.getRange('A' + row).setFormula('=IF(B' + row + '<>"",' + rank + ',"")');
    // Open cases (just Open status)
    sheet.getRange('D' + row).setFormula('=IF(B' + row + '<>"",COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gAssignedStewardCol + ':' + gAssignedStewardCol + ',B' + row + ',\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Open"),"")');
    // Pending Info cases
    sheet.getRange('E' + row).setFormula('=IF(B' + row + '<>"",COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gAssignedStewardCol + ':' + gAssignedStewardCol + ',B' + row + ',\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Pending Info"),"")');
    // Total Ever (all cases for this steward)
    sheet.getRange('F' + row).setFormula('=IF(B' + row + '<>"",COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gAssignedStewardCol + ':' + gAssignedStewardCol + ',B' + row + '),"")');
  }

  // Set horizontal alignment for all data rows
  sheet.getRange('A51:F80').setHorizontalAlignment('center');

  // Alternate row coloring for busiest stewards list
  for (var r = 51; r <= 80; r++) {
    if (r % 2 === 1) {
      sheet.getRange('A' + r + ':F' + r).setBackground('#F9FAFB');
    }
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 11: TOP 10 PERFORMERS BY SCORE
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A82').setValue('🏆 TOP 10 PERFORMERS BY SCORE')
    .setFontWeight('bold')
    .setBackground('#059669')  // Green
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A82:F82').merge();

  var topPerfLabels = [['Rank', 'Steward Name', 'Score', 'Win Rate %', 'Avg Days', 'Overdue']];
  sheet.getRange('A83:F83').setValues(topPerfLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  // Query hidden sheet for top 10 by Performance Score (descending)
  var topPerfQuery = '=IFERROR(QUERY(\'' + SHEETS.STEWARD_PERFORMANCE_CALC + '\'!A:J,' +
    '"SELECT A, J, F, G, H WHERE A <> \'\' AND A <> \'Steward\' ORDER BY J DESC LIMIT 10",' +
    '0),{"","","","",""})';
  sheet.getRange('B84').setFormula(topPerfQuery);

  // Add rank numbers for top performers
  for (var rank = 1; rank <= 10; rank++) {
    var row = 83 + rank;
    sheet.getRange('A' + row).setFormula('=IF(B' + row + '<>"",' + rank + ',"")');
  }

  // Alternate row coloring for top performers
  for (var r = 84; r <= 93; r++) {
    if (r % 2 === 0) {
      sheet.getRange('A' + r + ':F' + r).setBackground('#ECFDF5');  // Light green
    }
  }
  sheet.getRange('A84:F93').setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 12: STEWARDS NEEDING SUPPORT (Bottom 10 by Score)
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A95').setValue('⚠️ STEWARDS NEEDING SUPPORT (Lowest Scores)')
    .setFontWeight('bold')
    .setBackground('#DC2626')  // Red
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A95:F95').merge();

  var lowPerfLabels = [['Rank', 'Steward Name', 'Score', 'Win Rate %', 'Avg Days', 'Overdue']];
  sheet.getRange('A96:F96').setValues(lowPerfLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  // Query hidden sheet for bottom 10 by Performance Score (ascending)
  var lowPerfQuery = '=IFERROR(QUERY(\'' + SHEETS.STEWARD_PERFORMANCE_CALC + '\'!A:J,' +
    '"SELECT A, J, F, G, H WHERE A <> \'\' AND A <> \'Steward\' ORDER BY J ASC LIMIT 10",' +
    '0),{"","","","",""})';
  sheet.getRange('B97').setFormula(lowPerfQuery);

  // Add rank numbers for bottom performers (1 = lowest score)
  for (var rank = 1; rank <= 10; rank++) {
    var row = 96 + rank;
    sheet.getRange('A' + row).setFormula('=IF(B' + row + '<>"",' + rank + ',"")');
  }

  // Alternate row coloring for bottom performers
  for (var r = 97; r <= 106; r++) {
    if (r % 2 === 1) {
      sheet.getRange('A' + r + ':F' + r).setBackground('#FEF2F2');  // Light red
    }
  }
  sheet.getRange('A97:F106').setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════════════════════
  // FORMATTING AND CLEANUP
  // ═══════════════════════════════════════════════════════════════════════════

  // Auto-resize and format
  sheet.autoResizeColumns(1, 6);
  sheet.setFrozenRows(3);

  // Set minimum column widths
  for (var i = 1; i <= 6; i++) {
    if (sheet.getColumnWidth(i) < 120) {
      sheet.setColumnWidth(i, 120);
    }
  }

  // Hide column G onwards (if any exist)
  try {
    var maxCols = sheet.getMaxColumns();
    if (maxCols > 6) {
      sheet.deleteColumns(7, maxCols - 6);
    }
  } catch (e) {
    Logger.log('Could not delete extra columns: ' + e.toString());
  }

  // NUMBER FORMATTING: Use comma separators for all numeric values (1,000)
  var numberFormat = '#,##0';
  // Quick Stats row (row 6)
  sheet.getRange('A6:C6').setNumberFormat(numberFormat);
  sheet.getRange('E6:F6').setNumberFormat(numberFormat);
  // Member Metrics row (row 10)
  sheet.getRange('A10:B10').setNumberFormat(numberFormat);
  sheet.getRange('D10').setNumberFormat(numberFormat);
  // Grievance Metrics row (row 14)
  sheet.getRange('A14:F14').setNumberFormat(numberFormat);
  // Timeline row (row 18)
  sheet.getRange('A18:D18').setNumberFormat(numberFormat);
  // Type Analysis rows (rows 22-26)
  sheet.getRange('B22:D26').setNumberFormat(numberFormat);
  // Location Breakdown rows (rows 30-34)
  sheet.getRange('B30:D34').setNumberFormat(numberFormat);
  // Trends rows (rows 38-40)
  sheet.getRange('B38:D40').setNumberFormat(numberFormat);
  // Steward Performance row (row 47)
  sheet.getRange('A47:E47').setNumberFormat(numberFormat);
  // Top 30 Busiest Stewards (rows 51-80)
  sheet.getRange('C51:F80').setNumberFormat(numberFormat);
  // Top 10 Performers (rows 84-93) - Score and Win Rate have decimals
  var decimalFormat = '#,##0.0';
  sheet.getRange('C84:D93').setNumberFormat(decimalFormat);  // Score, Win Rate
  sheet.getRange('E84:F93').setNumberFormat(numberFormat);   // Avg Days, Overdue
  // Stewards Needing Support (rows 97-106)
  sheet.getRange('C97:D106').setNumberFormat(decimalFormat);  // Score, Win Rate
  sheet.getRange('E97:F106').setNumberFormat(numberFormat);   // Avg Days, Overdue
}

/**
 * Create or recreate the Custom View sheet
 * Allows users to select metrics and visualization preferences
 */
function createInteractiveDashboard(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.INTERACTIVE);
  sheet.clear();

  // Title
  sheet.getRange('A1').setValue('🎯 Custom View')
    .setFontSize(20)
    .setFontWeight('bold')
    .setFontColor(COLORS.PRIMARY_PURPLE);
  sheet.getRange('A1:F1').merge();

  // Instructions
  sheet.getRange('A3').setValue('Select metrics and chart types using the dropdowns below. Metrics auto-update from live data.')
    .setFontStyle('italic');
  sheet.getRange('A3:F3').merge();

  // ═══════════════════════════════════════════════════════════════════════════
  // METRIC SELECTION ROW
  // ═══════════════════════════════════════════════════════════════════════════
  var controlLabels = [['Metric 1', 'Metric 2', 'Metric 3', 'Time Range', 'Show Trend', 'Theme']];
  sheet.getRange('A5:F5').setValues(controlLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  // Default selections
  var defaultSelections = [['Total Members', 'Open Grievances', 'Win Rate', 'All Time', 'Yes', 'Default']];
  sheet.getRange('A6:F6').setValues(defaultSelections);

  // ═══════════════════════════════════════════════════════════════════════════
  // SELECTED METRICS DISPLAY
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A8').setValue('SELECTED METRICS')
    .setFontWeight('bold')
    .setBackground(COLORS.UNION_GREEN)
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A8:F8').merge();

  // Headers for metrics
  sheet.getRange('A9:C9').setValues([['Metric', 'Current Value', 'Description']])
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  // Dynamic metric formulas based on selection
  var mIdCol = getColumnLetter(MEMBER_COLS.MEMBER_ID);
  var mStewardCol = getColumnLetter(MEMBER_COLS.IS_STEWARD);
  var gIdCol = getColumnLetter(GRIEVANCE_COLS.GRIEVANCE_ID);
  var gStatusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);
  var gResolutionCol = getColumnLetter(GRIEVANCE_COLS.RESOLUTION);

  // Metric lookup table (row 10-17)
  var metricData = [
    ['Total Members', '=COUNTA(\'' + SHEETS.MEMBER_DIR + '\'!' + mIdCol + ':' + mIdCol + ')-1', 'Total union members in directory'],
    ['Active Stewards', '=COUNTIF(\'' + SHEETS.MEMBER_DIR + '\'!' + mStewardCol + ':' + mStewardCol + ',"Yes")', 'Members marked as stewards'],
    ['Total Grievances', '=COUNTA(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gIdCol + ':' + gIdCol + ')-1', 'All grievances filed'],
    ['Open Grievances', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Open")', 'Currently open cases'],
    ['Pending Info', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Pending Info")', 'Cases awaiting information'],
    ['Settled', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Settled")', 'Cases settled'],
    ['Won', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Won")', 'Cases won'],
    ['Win Rate', '=IFERROR(ROUND(COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Won")/(COUNTA(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gIdCol + ':' + gIdCol + ')-1)*100,1)&"%","0%")', 'Win percentage of all cases']
  ];

  for (var i = 0; i < metricData.length; i++) {
    sheet.getRange(10 + i, 1).setValue(metricData[i][0]);
    sheet.getRange(10 + i, 2).setFormula(metricData[i][1]);
    sheet.getRange(10 + i, 3).setValue(metricData[i][2]);
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // DROPDOWN VALIDATIONS
  // ═══════════════════════════════════════════════════════════════════════════
  // Metric dropdown options
  var metricOptions = ['Total Members', 'Active Stewards', 'Total Grievances', 'Open Grievances', 'Pending Info', 'Settled', 'Won', 'Win Rate'];
  var metricRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(metricOptions, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('A6').setDataValidation(metricRule);
  sheet.getRange('B6').setDataValidation(metricRule);
  sheet.getRange('C6').setDataValidation(metricRule);

  // Time range options
  var timeOptions = ['All Time', 'This Month', 'This Quarter', 'This Year', 'Last 30 Days', 'Last 90 Days'];
  var timeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(timeOptions, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('D6').setDataValidation(timeRule);

  // Yes/No options
  var yesNoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('E6').setDataValidation(yesNoRule);

  // Theme options
  var themeOptions = ['Default', 'Dark', 'High Contrast', 'Print Friendly'];
  var themeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(themeOptions, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('F6').setDataValidation(themeRule);

  // Format
  sheet.autoResizeColumns(1, 6);
  sheet.setColumnWidth(3, 250);
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Get existing sheet or create new one
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {string} name - Sheet name
 * @returns {Sheet} Sheet object
 */
function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  return ss.insertSheet(name);
}

/**
 * Setup hidden calculation sheets for cross-sheet data sync
 * Calls the full implementation in HiddenSheets.gs
 */
function setupHiddenSheets(ss) {
  // Call the full hidden sheet setup from HiddenSheets.gs
  setupAllHiddenSheets();

  // Install the auto-sync trigger
  installAutoSyncTrigger();
}

// ============================================================================
// DATA VALIDATION
// ============================================================================

/**
 * Setup all data validations for Member Directory and Grievance Log
 */
function setupDataValidations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!configSheet || !memberSheet || !grievanceSheet) {
    SpreadsheetApp.getUi().alert('Error: Required sheets not found. Please run CREATE_509_DASHBOARD first.');
    return;
  }

  // Member Directory Validations
  // Single-select dropdowns (strict validation)
  setDropdownValidation(memberSheet, MEMBER_COLS.JOB_TITLE, configSheet, CONFIG_COLS.JOB_TITLES);
  setDropdownValidation(memberSheet, MEMBER_COLS.WORK_LOCATION, configSheet, CONFIG_COLS.OFFICE_LOCATIONS);
  setDropdownValidation(memberSheet, MEMBER_COLS.UNIT, configSheet, CONFIG_COLS.UNITS);
  setDropdownValidation(memberSheet, MEMBER_COLS.IS_STEWARD, configSheet, CONFIG_COLS.YES_NO);
  setDropdownValidation(memberSheet, MEMBER_COLS.SUPERVISOR, configSheet, CONFIG_COLS.SUPERVISORS);
  setDropdownValidation(memberSheet, MEMBER_COLS.MANAGER, configSheet, CONFIG_COLS.MANAGERS);
  setDropdownValidation(memberSheet, MEMBER_COLS.INTEREST_LOCAL, configSheet, CONFIG_COLS.YES_NO);
  setDropdownValidation(memberSheet, MEMBER_COLS.INTEREST_CHAPTER, configSheet, CONFIG_COLS.YES_NO);
  setDropdownValidation(memberSheet, MEMBER_COLS.INTEREST_ALLIED, configSheet, CONFIG_COLS.YES_NO);
  setDropdownValidation(memberSheet, MEMBER_COLS.HOME_TOWN, configSheet, CONFIG_COLS.HOME_TOWNS);
  setDropdownValidation(memberSheet, MEMBER_COLS.CONTACT_STEWARD, configSheet, CONFIG_COLS.STEWARDS);

  // Multi-select dropdowns (allow comma-separated values)
  setMultiSelectValidation(memberSheet, MEMBER_COLS.OFFICE_DAYS, configSheet, CONFIG_COLS.OFFICE_DAYS);
  setMultiSelectValidation(memberSheet, MEMBER_COLS.PREFERRED_COMM, configSheet, CONFIG_COLS.COMM_METHODS);
  setMultiSelectValidation(memberSheet, MEMBER_COLS.BEST_TIME, configSheet, CONFIG_COLS.BEST_TIMES);
  setMultiSelectValidation(memberSheet, MEMBER_COLS.COMMITTEES, configSheet, CONFIG_COLS.STEWARD_COMMITTEES);
  setMultiSelectValidation(memberSheet, MEMBER_COLS.ASSIGNED_STEWARD, configSheet, CONFIG_COLS.STEWARDS);

  // Grievance Log Validations
  // Member ID dropdown - links to valid Member IDs from Member Directory
  setMemberIdValidation(grievanceSheet, memberSheet);

  setDropdownValidation(grievanceSheet, GRIEVANCE_COLS.STATUS, configSheet, CONFIG_COLS.GRIEVANCE_STATUS);
  setDropdownValidation(grievanceSheet, GRIEVANCE_COLS.CURRENT_STEP, configSheet, CONFIG_COLS.GRIEVANCE_STEP);
  setDropdownValidation(grievanceSheet, GRIEVANCE_COLS.ISSUE_CATEGORY, configSheet, CONFIG_COLS.ISSUE_CATEGORY);
  setDropdownValidation(grievanceSheet, GRIEVANCE_COLS.ARTICLES, configSheet, CONFIG_COLS.ARTICLES);

  // Note: Unit, Location, Steward columns now use formulas for auto-lookup
  // No need for manual dropdown validation on those columns

  SpreadsheetApp.getActiveSpreadsheet().toast('Data validations applied successfully!', '✅ Success', 3);
}

/**
 * Set Member ID validation dropdown from Member Directory
 * @param {Sheet} grievanceSheet - Grievance Log sheet
 * @param {Sheet} memberSheet - Member Directory sheet
 */
function setMemberIdValidation(grievanceSheet, memberSheet) {
  // Get the Member ID column from Member Directory
  var memberIdCol = getColumnLetter(MEMBER_COLS.MEMBER_ID);
  var sourceRange = memberSheet.getRange(memberIdCol + '2:' + memberIdCol + '1000');

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(false)
    .build();

  var targetRange = grievanceSheet.getRange(2, GRIEVANCE_COLS.MEMBER_ID, 998, 1);
  targetRange.setDataValidation(rule);
}

/**
 * Set dropdown validation from Config sheet
 * @param {Sheet} targetSheet - Sheet to apply validation
 * @param {number} targetCol - Column number in target sheet
 * @param {Sheet} configSheet - Config sheet with source values
 * @param {number} sourceCol - Column number in Config sheet
 */
function setDropdownValidation(targetSheet, targetCol, configSheet, sourceCol) {
  // Config data starts at row 3 (row 1 = section headers, row 2 = column headers)
  var sourceRange = configSheet.getRange(3, sourceCol, 100, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(false)
    .build();

  var targetRange = targetSheet.getRange(2, targetCol, 998, 1);
  targetRange.setDataValidation(rule);
}

/**
 * Set multi-select validation (allows comma-separated values)
 * Shows dropdown for convenience but accepts any text
 * @param {Sheet} targetSheet - Sheet to apply validation
 * @param {number} targetCol - Column number in target sheet
 * @param {Sheet} configSheet - Config sheet with source values
 * @param {number} sourceCol - Column number in Config sheet
 */
function setMultiSelectValidation(targetSheet, targetCol, configSheet, sourceCol) {
  // Config data starts at row 3
  var sourceRange = configSheet.getRange(3, sourceCol, 100, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(true)  // Allow comma-separated values
    .build();

  var targetRange = targetSheet.getRange(2, targetCol, 998, 1);
  targetRange.setDataValidation(rule);
}

// ============================================================================
// MULTI-SELECT FUNCTIONALITY
// ============================================================================

// Store the target cell for multi-select dialog
var multiSelectTarget_ = null;

/**
 * Show multi-select dialog for the current cell
 * Called from menu or double-click on multi-select column
 */
function showMultiSelectDialog() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cell = sheet.getActiveCell();
  var sheetName = sheet.getName();

  // Only works on Member Directory
  if (sheetName !== SHEETS.MEMBER_DIR) {
    SpreadsheetApp.getUi().alert('Multi-select is only available in Member Directory.');
    return;
  }

  var col = cell.getColumn();
  var row = cell.getRow();

  // Must be in data row (not header)
  if (row < 2) {
    SpreadsheetApp.getUi().alert('Please select a data cell (row 2 or below).');
    return;
  }

  var config = getMultiSelectConfig(col);
  if (!config) {
    SpreadsheetApp.getUi().alert('This column does not support multi-select.\n\nMulti-select columns: Office Days, Preferred Communication, Best Time, Committees, Assigned Steward(s)');
    return;
  }

  // Store target cell info in PropertiesService for the callback
  var props = PropertiesService.getDocumentProperties();
  props.setProperty('multiSelectRow', row.toString());
  props.setProperty('multiSelectCol', col.toString());

  // Get options from Config sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  var options = getConfigValues(configSheet, config.configCol);

  // Get current value and parse into array
  var currentValue = cell.getValue() || '';
  var currentValues = currentValue ? currentValue.split(/,\s*/) : [];

  // Dialog data
  var dialogData = {
    label: config.label,
    options: options,
    currentValues: currentValues
  };

  // Create inline HTML for multi-select dialog
  var htmlContent = '<!DOCTYPE html><html><head><base target="_top"><style>' +
    '*{box-sizing:border-box;font-family:"Google Sans",Arial,sans-serif}' +
    'body{margin:0;padding:16px;background:#fff}' +
    'h3{margin:0 0 12px 0;color:#7C3AED;font-size:16px}' +
    '.options-container{max-height:250px;overflow-y:auto;border:1px solid #e0e0e0;border-radius:8px;padding:8px;margin-bottom:16px}' +
    '.option-item{display:flex;align-items:center;padding:8px 12px;margin:2px 0;border-radius:6px;cursor:pointer;transition:background 0.15s}' +
    '.option-item:hover{background:#f3f4f6}.option-item.selected{background:#ede9fe}' +
    '.option-item input[type="checkbox"]{margin-right:10px;width:18px;height:18px;accent-color:#7C3AED}' +
    '.option-item label{flex:1;cursor:pointer;font-size:14px;color:#1f2937}' +
    '.button-row{display:flex;gap:8px;justify-content:flex-end}' +
    'button{padding:10px 20px;border:none;border-radius:6px;font-size:14px;font-weight:500;cursor:pointer;transition:all 0.15s}' +
    '.btn-primary{background:#7C3AED;color:white}.btn-primary:hover{background:#6d28d9}' +
    '.btn-secondary{background:#f3f4f6;color:#374151}.btn-secondary:hover{background:#e5e7eb}' +
    '.btn-clear{background:#fef2f2;color:#dc2626}.btn-clear:hover{background:#fee2e2}' +
    '.selected-count{font-size:12px;color:#6b7280;margin-bottom:8px}' +
    '.quick-actions{display:flex;gap:8px;margin-bottom:12px}.quick-actions button{padding:6px 12px;font-size:12px}' +
    '</style></head><body>' +
    '<h3 id="dialogTitle">Select Options</h3>' +
    '<div class="quick-actions"><button class="btn-secondary" onclick="selectAll()">Select All</button>' +
    '<button class="btn-clear" onclick="clearAll()">Clear All</button></div>' +
    '<div class="selected-count" id="selectedCount">0 selected</div>' +
    '<div class="options-container" id="optionsContainer"></div>' +
    '<div class="button-row"><button class="btn-secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn-primary" onclick="saveSelection()">Save</button></div>' +
    '<script>' +
    'var options=[];var currentSelection=[];' +
    'function initDialog(data){document.getElementById("dialogTitle").textContent="Select "+data.label;options=data.options;currentSelection=data.currentValues||[];renderOptions();updateCount()}' +
    'function renderOptions(){var c=document.getElementById("optionsContainer");c.innerHTML="";options.forEach(function(o,i){var s=currentSelection.indexOf(o)!==-1;var d=document.createElement("div");d.className="option-item"+(s?" selected":"");d.onclick=function(){toggleOption(i)};var cb=document.createElement("input");cb.type="checkbox";cb.id="opt_"+i;cb.checked=s;var l=document.createElement("label");l.htmlFor="opt_"+i;l.textContent=o;d.appendChild(cb);d.appendChild(l);c.appendChild(d)})}' +
    'function toggleOption(i){var o=options[i];var p=currentSelection.indexOf(o);if(p===-1){currentSelection.push(o)}else{currentSelection.splice(p,1)}renderOptions();updateCount()}' +
    'function selectAll(){currentSelection=options.slice();renderOptions();updateCount()}' +
    'function clearAll(){currentSelection=[];renderOptions();updateCount()}' +
    'function updateCount(){document.getElementById("selectedCount").textContent=currentSelection.length+" selected"}' +
    'function saveSelection(){var v=currentSelection.join(", ");google.script.run.withSuccessHandler(function(){google.script.host.close()}).withFailureHandler(function(e){alert("Error: "+e.message)}).applyMultiSelectValue(v)}' +
    '</script>' +
    '<script>initDialog(' + JSON.stringify(dialogData) + ');</script>' +
    '</body></html>';

  var html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(350)
    .setHeight(420);

  SpreadsheetApp.getUi().showModalDialog(html, 'Multi-Select: ' + config.label);
}

/**
 * Get values from a Config sheet column
 * Note: Row 1 = section headers, Row 2 = column headers, Row 3+ = data
 * @param {Sheet} configSheet - The Config sheet
 * @param {number} col - Column number
 * @returns {Array} Array of non-empty values
 */
function getConfigValues(configSheet, col) {
  var lastRow = configSheet.getLastRow();
  if (lastRow < 3) return [];

  var data = configSheet.getRange(3, col, lastRow - 2, 1).getValues();
  var values = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() !== '') {
      values.push(data[i][0].toString());
    }
  }
  return values;
}

/**
 * Apply the multi-select value to the stored cell
 * Called from the dialog
 * @param {string} value - Comma-separated selected values
 */
function applyMultiSelectValue(value) {
  var props = PropertiesService.getDocumentProperties();
  var row = parseInt(props.getProperty('multiSelectRow'), 10);
  var col = parseInt(props.getProperty('multiSelectCol'), 10);

  if (!row || !col) {
    throw new Error('Target cell not found. Please try again.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  sheet.getRange(row, col).setValue(value);

  // Clear stored properties
  props.deleteProperty('multiSelectRow');
  props.deleteProperty('multiSelectCol');
}

/**
 * Handle edit events to trigger multi-select dialog
 * This is installed as an onEdit trigger
 */
function onEditMultiSelect(e) {
  // Only process single cell edits
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();

  // Only Member Directory
  if (sheetName !== SHEETS.MEMBER_DIR) return;

  var col = e.range.getColumn();
  var row = e.range.getRow();

  // Skip header row
  if (row < 2) return;

  // Check if this is a multi-select column
  var config = getMultiSelectConfig(col);
  if (!config) return;

  // If user typed something, show the dialog to help them select properly
  // Only trigger if the new value isn't already comma-separated (user might be pasting)
  var newValue = e.value || '';
  var oldValue = e.oldValue || '';

  // If user cleared the cell or pasted valid data, don't interrupt
  if (newValue === '' || newValue.indexOf(',') !== -1) return;

  // Show helpful toast
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Tip: Use Dashboard menu > "Multi-Select Editor" for easier selection of multiple values.',
    config.label,
    5
  );
}

/**
 * Handle selection change to auto-open multi-select dialog
 * This is installed as an onSelectionChange trigger
 */
function onSelectionChangeMultiSelect(e) {
  // Only process if we have a valid range
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();

  // Only Member Directory
  if (sheetName !== SHEETS.MEMBER_DIR) return;

  var col = e.range.getColumn();
  var row = e.range.getRow();

  // Skip header row and multi-cell selections
  if (row < 2) return;
  if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;

  // Check if this is a multi-select column
  var config = getMultiSelectConfig(col);
  if (!config) return;

  // Check if we already showed dialog for this cell (avoid repeated opens)
  var props = PropertiesService.getDocumentProperties();
  var lastCell = props.getProperty('lastMultiSelectCell');
  var currentCell = row + ',' + col;

  if (lastCell === currentCell) return;

  // Store current cell
  props.setProperty('lastMultiSelectCell', currentCell);

  // Auto-open the multi-select dialog
  showMultiSelectDialog();
}

/**
 * Install the multi-select auto-open trigger
 * Run this once to enable auto-open on cell selection
 */
function installMultiSelectTrigger() {
  var ui = SpreadsheetApp.getUi();

  // Note: onSelectionChange triggers cannot be created programmatically
  // User must set this up manually in Apps Script editor
  ui.alert('☑️ Multi-Select Auto-Open Setup',
    'To enable auto-open for multi-select cells:\n\n' +
    '1. Go to Extensions → Apps Script\n' +
    '2. Click the clock icon (Triggers) in the left sidebar\n' +
    '3. Click "+ Add Trigger"\n' +
    '4. Choose function: onSelectionChangeMultiSelect\n' +
    '5. Select event type: "On change" or "On edit"\n' +
    '6. Click Save\n\n' +
    'Alternatively, use the manual method:\n' +
    '• Select a multi-select cell\n' +
    '• Go to Tools → Multi-Select → Open Editor',
    ui.ButtonSet.OK);
}

/**
 * Remove the multi-select auto-open trigger
 */
function removeMultiSelectTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getUserTriggers(ss);
  var removed = false;

  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onSelectionChangeMultiSelect') {
      ScriptApp.deleteTrigger(trigger);
      removed = true;
    }
  });

  if (removed) {
    SpreadsheetApp.getUi().alert('Multi-Select auto-open has been disabled.');
  } else {
    SpreadsheetApp.getUi().alert('No multi-select trigger was found.');
  }
}

// ============================================================================
// DIAGNOSE FUNCTION
// ============================================================================

/**
 * System health check - validates sheets and column counts
 */
function DIAGNOSE_SETUP() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var report = [];

  report.push('🔍 509 DASHBOARD DIAGNOSTIC REPORT');
  report.push('================================');
  report.push('');

  // Check all required sheets (core + dashboard sheets)
  var requiredSheets = [
    SHEETS.CONFIG,
    SHEETS.MEMBER_DIR,
    SHEETS.GRIEVANCE_LOG,
    SHEETS.DASHBOARD,
    SHEETS.INTERACTIVE
  ];

  report.push('📋 SHEET CHECK:');
  var allSheetsPresent = true;
  requiredSheets.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      report.push('  ✅ ' + sheetName);
    } else {
      report.push('  ❌ ' + sheetName + ' - MISSING');
      allSheetsPresent = false;
    }
  });

  report.push('');

  // Check column counts
  report.push('📊 COLUMN CHECK:');

  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (memberSheet) {
    var memberCols = memberSheet.getLastColumn();
    var expectedMemberCols = 31;
    if (memberCols >= expectedMemberCols) {
      report.push('  ✅ Member Directory: ' + memberCols + ' columns (expected ' + expectedMemberCols + ')');
    } else {
      report.push('  ⚠️ Member Directory: ' + memberCols + ' columns (expected ' + expectedMemberCols + ')');
    }
  }

  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (grievanceSheet) {
    var grievanceCols = grievanceSheet.getLastColumn();
    var expectedGrievanceCols = 34;
    if (grievanceCols >= expectedGrievanceCols) {
      report.push('  ✅ Grievance Log: ' + grievanceCols + ' columns (expected ' + expectedGrievanceCols + ')');
    } else {
      report.push('  ⚠️ Grievance Log: ' + grievanceCols + ' columns (expected ' + expectedGrievanceCols + ')');
    }
  }

  report.push('');

  // Check hidden sheets (6 hidden calculation sheets)
  report.push('🔒 HIDDEN SHEETS:');
  var hiddenSheets = [
    SHEETS.GRIEVANCE_CALC,
    SHEETS.GRIEVANCE_FORMULAS,
    SHEETS.MEMBER_LOOKUP,
    SHEETS.STEWARD_CONTACT_CALC,
    SHEETS.DASHBOARD_CALC,
    SHEETS.STEWARD_PERFORMANCE_CALC
  ];

  hiddenSheets.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      report.push('  ✅ ' + sheetName + (sheet.isSheetHidden() ? ' (hidden)' : ' (visible)'));
    } else {
      report.push('  ⚠️ ' + sheetName + ' - not created');
    }
  });

  report.push('');
  report.push('================================');
  report.push(allSheetsPresent ? '✅ Core sheets OK' : '❌ Some sheets missing');

  // Display report
  ui.alert('Diagnostic Report', report.join('\n'), ui.ButtonSet.OK);

  // Also log it
  Logger.log(report.join('\n'));
}

// ============================================================================
// REPAIR FUNCTION
// ============================================================================

/**
 * Repair dashboard - recreates hidden sheets, triggers, and syncs data
 */
function REPAIR_DASHBOARD() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert(
    '🔧 Repair Dashboard',
    'This will:\n\n' +
    '• Recreate all 6 hidden calculation sheets with formulas\n' +
    '• Install auto-sync trigger\n' +
    '• Sync all cross-sheet data\n' +
    '• Reapply data validations\n\n' +
    'Your data will NOT be affected.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  ss.toast('Repairing dashboard...', '🔧 Repair', 3);

  try {
    // Use the full repair function from HiddenSheets.gs
    repairAllHiddenSheets();

    // Also reapply data validations
    setupDataValidations();

    // Create Menu Checklist sheet
    createMenuChecklistSheet_();

    // Final message handled by repairAllHiddenSheets
  } catch (error) {
    Logger.log('Error in REPAIR_DASHBOARD: ' + error.message);
    ui.alert('❌ Error', 'Repair failed: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Creates the Menu Checklist sheet with all menu items
 * Called automatically during dashboard repair/creation
 * @private
 */
function createMenuChecklistSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = SHEETS.MENU_CHECKLIST || 'Menu Checklist';

  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  // Menu items organized by optimal testing order
  var menuItems = [
    // ═══ PHASE 1: Foundation & Setup (Test these first!) ═══
    ['1️⃣ Foundation', '🏗️ Setup', '🔧 REPAIR DASHBOARD', 'REPAIR_DASHBOARD'],
    ['1️⃣ Foundation', '⚙️ Administrator', '🔍 DIAGNOSE SETUP', 'DIAGNOSE_SETUP'],
    ['1️⃣ Foundation', '⚙️ Administrator', '🔍 Verify Hidden Sheets', 'verifyHiddenSheets'],
    ['1️⃣ Foundation', '⚙️ Admin > Setup', '🔧 Setup All Hidden Sheets', 'setupAllHiddenSheets'],
    ['1️⃣ Foundation', '⚙️ Admin > Setup', '🔧 Repair All Hidden Sheets', 'repairAllHiddenSheets'],
    ['1️⃣ Foundation', '🏗️ Setup', '⚙️ Setup Data Validations', 'setupDataValidations'],
    ['1️⃣ Foundation', '🏗️ Setup', '🎨 Setup ADHD Defaults', 'setupADHDDefaults'],

    // ═══ PHASE 2: Triggers & Data Sync ═══
    ['2️⃣ Sync', '⚙️ Admin > Setup', '⚡ Install Auto-Sync Trigger', 'installAutoSyncTrigger'],
    ['2️⃣ Sync', '⚙️ Admin > Sync', '🔄 Sync All Data Now', 'syncAllData'],
    ['2️⃣ Sync', '⚙️ Admin > Sync', '🔄 Sync Grievance → Members', 'syncGrievanceToMemberDirectory'],
    ['2️⃣ Sync', '⚙️ Admin > Sync', '🔄 Sync Members → Grievances', 'syncMemberToGrievanceLog'],
    ['2️⃣ Sync', '⚙️ Admin > Setup', '🚫 Remove Auto-Sync Trigger', 'removeAutoSyncTrigger'],

    // ═══ PHASE 3: Core Dashboards ═══
    ['3️⃣ Dashboards', '👤 Dashboard', '📊 Smart Dashboard (Auto-Detect)', 'showSmartDashboard'],
    ['3️⃣ Dashboards', '👤 Dashboard', '🎯 Custom View', 'showInteractiveDashboardTab'],
    ['3️⃣ Dashboards', '👤 Dashboard', '📋 View Active Grievances', 'viewActiveGrievances'],
    ['3️⃣ Dashboards', '👤 Dashboard', '📱 Mobile Dashboard', 'showMobileDashboard'],
    ['3️⃣ Dashboards', '👤 Dashboard', '📱 Get Mobile App URL', 'showWebAppUrl'],
    ['3️⃣ Dashboards', '👤 Dashboard', '⚡ Quick Actions', 'showQuickActionsMenu'],
    ['3️⃣ Dashboards', '📊 Sheet Manager', '📊 Rebuild Dashboard', 'rebuildDashboard'],
    ['3️⃣ Dashboards', '📊 Sheet Manager', '📈 Refresh Interactive Charts', 'refreshInteractiveCharts'],
    ['3️⃣ Dashboards', '📊 Sheet Manager', '🔄 Refresh All Formulas', 'refreshAllFormulas'],

    // ═══ PHASE 4: Search ═══
    ['4️⃣ Search', '🔍 Search', '🔍 Search Members', 'searchMembers'],

    // ═══ PHASE 5: Grievance Management ═══
    ['5️⃣ Grievances', '👤 Grievance Tools', '➕ Start New Grievance', 'startNewGrievance'],
    ['5️⃣ Grievances', '👤 Grievance Tools', '🔄 Refresh Grievance Formulas', 'recalcAllGrievancesBatched'],
    ['5️⃣ Grievances', '👤 Grievance Tools', '🔄 Refresh Member Directory Data', 'refreshMemberDirectoryFormulas'],
    ['5️⃣ Grievances', '👤 Grievance Tools', '🔗 Setup Live Grievance Links', 'setupLiveGrievanceFormulas'],
    ['5️⃣ Grievances', '👤 Grievance Tools', '👤 Setup Member ID Dropdown', 'setupGrievanceMemberDropdown'],
    ['5️⃣ Grievances', '👤 Grievance Tools', '🔧 Fix Overdue Text Data', 'fixOverdueTextToNumbers'],

    // ═══ PHASE 6: Google Drive ═══
    ['6️⃣ Drive', '📊 Google Drive', '📁 Setup Folder for Grievance', 'setupDriveFolderForGrievance'],
    ['6️⃣ Drive', '📊 Google Drive', '📁 View Grievance Files', 'showGrievanceFiles'],
    ['6️⃣ Drive', '📊 Google Drive', '📁 Batch Create Folders', 'batchCreateGrievanceFolders'],

    // ═══ PHASE 7: Calendar ═══
    ['7️⃣ Calendar', '📊 Calendar', '📅 Sync Deadlines to Calendar', 'syncDeadlinesToCalendar'],
    ['7️⃣ Calendar', '📊 Calendar', '📅 View Upcoming Deadlines', 'showUpcomingDeadlinesFromCalendar'],
    ['7️⃣ Calendar', '📊 Calendar', '🗑️ Clear Calendar Events', 'clearAllCalendarEvents'],

    // ═══ PHASE 8: Notifications ═══
    ['8️⃣ Notify', '📊 Notifications', '⚙️ Notification Settings', 'showNotificationSettings'],
    ['8️⃣ Notify', '📊 Notifications', '🧪 Test Notifications', 'testDeadlineNotifications'],

    // ═══ PHASE 9: Accessibility & Theming ═══
    ['9️⃣ Access', '🔧 ADHD', '♿ ADHD Control Panel', 'showADHDControlPanel'],
    ['9️⃣ Access', '🔧 ADHD', '🎯 Focus Mode', 'activateFocusMode'],
    ['9️⃣ Access', '🔧 ADHD', '🔲 Toggle Zebra Stripes', 'toggleZebraStripes'],
    ['9️⃣ Access', '🔧 ADHD', '📝 Quick Capture', 'showQuickCaptureNotepad'],
    ['9️⃣ Access', '🔧 ADHD', '🍅 Pomodoro Timer', 'startPomodoroTimer'],
    ['9️⃣ Access', '🔧 Theming', '🎨 Theme Manager', 'showThemeManager'],
    ['9️⃣ Access', '🔧 Theming', '🌙 Toggle Dark Mode', 'quickToggleDarkMode'],
    ['9️⃣ Access', '🔧 Theming', '🔄 Reset Theme', 'resetToDefaultTheme'],

    // ═══ PHASE 10: Productivity Tools ═══
    ['🔟 Tools', '🔧 Multi-Select', '📝 Open Editor', 'showMultiSelectDialog'],
    ['🔟 Tools', '🔧 Multi-Select', '⚡ Enable Auto-Open', 'installMultiSelectTrigger'],
    ['🔟 Tools', '🔧 Multi-Select', '🚫 Disable Auto-Open', 'removeMultiSelectTrigger'],

    // ═══ PHASE 11: Performance & Cache ═══
    ['1️⃣1️⃣ Perf', '🔧 Cache', '🗄️ Cache Status', 'showCacheStatusDashboard'],
    ['1️⃣1️⃣ Perf', '🔧 Cache', '🔥 Warm Up Caches', 'warmUpCaches'],
    ['1️⃣1️⃣ Perf', '🔧 Cache', '🗑️ Clear All Caches', 'invalidateAllCaches'],

    // ═══ PHASE 12: Validation ═══
    ['1️⃣2️⃣ Valid', '🔧 Validation', '🔍 Run Bulk Validation', 'runBulkValidation'],
    ['1️⃣2️⃣ Valid', '🔧 Validation', '⚙️ Validation Settings', 'showValidationSettings'],
    ['1️⃣2️⃣ Valid', '🔧 Validation', '🧹 Clear Indicators', 'clearValidationIndicators'],
    ['1️⃣2️⃣ Valid', '🔧 Validation', '⚡ Install Validation Trigger', 'installValidationTrigger'],

    // ═══ PHASE 13: Testing (Run last to verify everything) ═══
    ['1️⃣3️⃣ Test', '🧪 Testing', '🧪 Run All Tests', 'runAllTests'],
    ['1️⃣3️⃣ Test', '🧪 Testing', '⚡ Run Quick Tests', 'runQuickTests'],
    ['1️⃣3️⃣ Test', '🧪 Testing', '📊 View Test Results', 'viewTestResults']
  ];

  // Build rows with header
  var rows = [['✓', 'Phase', 'Menu', 'Item', 'Function', 'Notes']];
  for (var i = 0; i < menuItems.length; i++) {
    rows.push([false, menuItems[i][0], menuItems[i][1], menuItems[i][2], menuItems[i][3], '']);
  }

  // Write all data
  sheet.getRange(1, 1, rows.length, 6).setValues(rows);

  // Format header
  sheet.getRange(1, 1, 1, 6)
    .setFontWeight('bold')
    .setBackground(COLORS.PRIMARY_PURPLE || '#7C3AED')
    .setFontColor(COLORS.WHITE || '#FFFFFF')
    .setHorizontalAlignment('center');

  // Add checkboxes
  if (rows.length > 1) {
    sheet.getRange(2, 1, rows.length - 1, 1).insertCheckboxes();
  }

  // Set column widths
  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 250);
  sheet.setColumnWidth(5, 250);
  sheet.setColumnWidth(6, 200);

  // Freeze header
  sheet.setFrozenRows(1);

  // Alternating colors
  for (var r = 2; r <= rows.length; r++) {
    if (r % 2 === 0) {
      sheet.getRange(r, 1, 1, 6).setBackground('#F9FAFB');
    }
  }

  // Conditional formatting for checked items
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$A2=TRUE')
    .setBackground('#E8F5E9')
    .setRanges([sheet.getRange(2, 1, rows.length - 1, 6)])
    .build();
  sheet.setConditionalFormatRules([rule]);

  return sheet;
}

// ============================================================================
// MENU HANDLER FUNCTIONS
// ============================================================================

/**
 * Show desktop search modal - comprehensive search for members and grievances
 * Enhanced version of mobile search with more fields and filtering options
 */
function searchMembers() {
  showDesktopSearch();
}

/**
 * Show the desktop unified search dialog
 * Optimized for larger screens with advanced filtering
 */
function showDesktopSearch() {
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top">' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:#f8fafc;color:#1e293b}' +

    /* Header */
    '.header{background:linear-gradient(135deg,#7C3AED 0%,#5B21B6 100%);color:white;padding:24px 30px}' +
    '.header h1{font-size:24px;font-weight:600;margin-bottom:8px;display:flex;align-items:center;gap:12px}' +
    '.header p{opacity:0.9;font-size:14px}' +

    /* Search Container */
    '.search-section{padding:20px 30px;background:white;border-bottom:1px solid #e2e8f0}' +
    '.search-row{display:flex;gap:12px;align-items:center}' +
    '.search-input-container{flex:1;position:relative}' +
    '.search-input{width:100%;padding:14px 16px 14px 48px;border:2px solid #e2e8f0;border-radius:10px;font-size:16px;transition:all 0.2s}' +
    '.search-input:focus{outline:none;border-color:#7C3AED;box-shadow:0 0 0 3px rgba(124,58,237,0.1)}' +
    '.search-icon{position:absolute;left:16px;top:50%;transform:translateY(-50%);font-size:20px;color:#94a3b8}' +

    /* Tabs */
    '.tabs{display:flex;gap:4px;background:#f1f5f9;padding:4px;border-radius:8px;width:fit-content}' +
    '.tab{padding:10px 20px;border:none;background:transparent;border-radius:6px;cursor:pointer;font-size:14px;font-weight:500;color:#64748b;transition:all 0.2s}' +
    '.tab:hover{background:#e2e8f0}' +
    '.tab.active{background:white;color:#7C3AED;box-shadow:0 1px 3px rgba(0,0,0,0.1)}' +

    /* Filters */
    '.filters-section{padding:16px 30px;background:#f8fafc;border-bottom:1px solid #e2e8f0}' +
    '.filters-row{display:flex;gap:12px;flex-wrap:wrap;align-items:center}' +
    '.filter-group{display:flex;align-items:center;gap:8px}' +
    '.filter-label{font-size:13px;color:#64748b;font-weight:500}' +
    '.filter-select{padding:8px 12px;border:1px solid #e2e8f0;border-radius:6px;font-size:13px;background:white;cursor:pointer;min-width:140px}' +
    '.filter-select:focus{outline:none;border-color:#7C3AED}' +
    '.clear-filters{padding:8px 16px;border:1px solid #e2e8f0;border-radius:6px;background:white;cursor:pointer;font-size:13px;color:#64748b}' +
    '.clear-filters:hover{background:#f1f5f9}' +

    /* Results */
    '.results-section{padding:20px 30px;flex:1;overflow-y:auto;max-height:400px}' +
    '.results-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:16px}' +
    '.results-count{font-size:14px;color:#64748b}' +
    '.results-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:16px}' +

    /* Result Card */
    '.result-card{background:white;border-radius:12px;padding:20px;border:1px solid #e2e8f0;cursor:pointer;transition:all 0.2s}' +
    '.result-card:hover{border-color:#7C3AED;box-shadow:0 4px 12px rgba(124,58,237,0.15);transform:translateY(-1px)}' +
    '.card-header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:12px}' +
    '.card-type{display:inline-flex;align-items:center;gap:6px;padding:4px 10px;border-radius:20px;font-size:12px;font-weight:500}' +
    '.type-member{background:#dbeafe;color:#1d4ed8}' +
    '.type-grievance{background:#fee2e2;color:#dc2626}' +
    '.card-id{font-size:12px;color:#94a3b8;font-family:monospace}' +
    '.card-title{font-size:16px;font-weight:600;color:#1e293b;margin-bottom:8px}' +
    '.card-details{display:flex;flex-direction:column;gap:6px}' +
    '.detail-row{display:flex;align-items:center;gap:8px;font-size:13px;color:#64748b}' +
    '.detail-icon{width:16px;text-align:center}' +
    '.detail-label{color:#94a3b8;min-width:60px}' +
    '.status-badge{display:inline-block;padding:3px 10px;border-radius:12px;font-size:11px;font-weight:600}' +
    '.status-open{background:#fef3c7;color:#d97706}' +
    '.status-pending{background:#fef3c7;color:#b45309}' +
    '.status-closed{background:#d1fae5;color:#059669}' +

    /* Empty State */
    '.empty-state{text-align:center;padding:60px 20px;color:#94a3b8}' +
    '.empty-icon{font-size:48px;margin-bottom:16px}' +
    '.empty-title{font-size:18px;font-weight:600;color:#64748b;margin-bottom:8px}' +
    '.empty-text{font-size:14px}' +

    /* Loading */
    '.loading{text-align:center;padding:40px;color:#64748b}' +
    '.spinner{display:inline-block;width:24px;height:24px;border:3px solid #e2e8f0;border-top-color:#7C3AED;border-radius:50%;animation:spin 1s linear infinite}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +

    /* Footer */
    '.footer{padding:16px 30px;background:white;border-top:1px solid #e2e8f0;display:flex;justify-content:space-between;align-items:center}' +
    '.footer-info{font-size:13px;color:#94a3b8}' +
    '.close-btn{padding:10px 24px;background:#64748b;color:white;border:none;border-radius:8px;cursor:pointer;font-size:14px;font-weight:500}' +
    '.close-btn:hover{background:#475569}' +
    '</style></head><body>' +

    '<div class="header">' +
    '  <h1>🔍 Search Dashboard</h1>' +
    '  <p>Search across members and grievances with advanced filters</p>' +
    '</div>' +

    '<div class="search-section">' +
    '  <div class="search-row">' +
    '    <div class="search-input-container">' +
    '      <span class="search-icon">🔍</span>' +
    '      <input type="text" class="search-input" id="searchQuery" placeholder="Search by name, ID, email, job title, location, or issue type..." autofocus>' +
    '    </div>' +
    '    <div class="tabs">' +
    '      <button class="tab active" data-tab="all">All</button>' +
    '      <button class="tab" data-tab="members">Members</button>' +
    '      <button class="tab" data-tab="grievances">Grievances</button>' +
    '    </div>' +
    '  </div>' +
    '</div>' +

    '<div class="filters-section">' +
    '  <div class="filters-row">' +
    '    <div class="filter-group" id="statusFilter" style="display:none">' +
    '      <span class="filter-label">Status:</span>' +
    '      <select class="filter-select" id="filterStatus">' +
    '        <option value="">All Statuses</option>' +
    '        <option value="Open">Open</option>' +
    '        <option value="Pending Info">Pending Info</option>' +
    '        <option value="Settled">Settled</option>' +
    '        <option value="Closed">Closed</option>' +
    '        <option value="Won">Won</option>' +
    '        <option value="Denied">Denied</option>' +
    '      </select>' +
    '    </div>' +
    '    <div class="filter-group" id="locationFilter">' +
    '      <span class="filter-label">Location:</span>' +
    '      <select class="filter-select" id="filterLocation">' +
    '        <option value="">All Locations</option>' +
    '      </select>' +
    '    </div>' +
    '    <div class="filter-group" id="stewardFilter">' +
    '      <span class="filter-label">Is Steward:</span>' +
    '      <select class="filter-select" id="filterSteward">' +
    '        <option value="">Any</option>' +
    '        <option value="Yes">Yes</option>' +
    '        <option value="No">No</option>' +
    '      </select>' +
    '    </div>' +
    '    <button class="clear-filters" onclick="clearFilters()">Clear Filters</button>' +
    '  </div>' +
    '</div>' +

    '<div class="results-section" id="resultsSection">' +
    '  <div class="empty-state">' +
    '    <div class="empty-icon">🔍</div>' +
    '    <div class="empty-title">Start searching</div>' +
    '    <div class="empty-text">Type at least 2 characters to search across all records</div>' +
    '  </div>' +
    '</div>' +

    '<div class="footer">' +
    '  <div class="footer-info">Press Enter to search • Click a result to navigate</div>' +
    '  <button class="close-btn" onclick="google.script.host.close()">Close</button>' +
    '</div>' +

    '<script>' +
    'var currentTab = "all";' +
    'var currentResults = [];' +
    'var locations = [];' +

    /* Initialize */
    'document.addEventListener("DOMContentLoaded", function() {' +
    '  loadLocations();' +
    '  setupEventListeners();' +
    '});' +

    /* Load locations for filter */
    'function loadLocations() {' +
    '  google.script.run.withSuccessHandler(function(locs) {' +
    '    locations = locs || [];' +
    '    var select = document.getElementById("filterLocation");' +
    '    locations.forEach(function(loc) {' +
    '      if (loc) {' +
    '        var opt = document.createElement("option");' +
    '        opt.value = loc;' +
    '        opt.textContent = loc;' +
    '        select.appendChild(opt);' +
    '      }' +
    '    });' +
    '  }).getDesktopSearchLocations();' +
    '}' +

    /* Event Listeners */
    'function setupEventListeners() {' +
    '  var input = document.getElementById("searchQuery");' +
    '  var debounceTimer;' +
    '  input.addEventListener("input", function() {' +
    '    clearTimeout(debounceTimer);' +
    '    debounceTimer = setTimeout(function() { doSearch(); }, 300);' +
    '  });' +
    '  input.addEventListener("keydown", function(e) {' +
    '    if (e.key === "Enter") { doSearch(); }' +
    '  });' +
    '  document.querySelectorAll(".tab").forEach(function(tab) {' +
    '    tab.addEventListener("click", function() {' +
    '      document.querySelectorAll(".tab").forEach(function(t) { t.classList.remove("active"); });' +
    '      this.classList.add("active");' +
    '      currentTab = this.dataset.tab;' +
    '      updateFilterVisibility();' +
    '      doSearch();' +
    '    });' +
    '  });' +
    '  document.querySelectorAll(".filter-select").forEach(function(sel) {' +
    '    sel.addEventListener("change", function() { doSearch(); });' +
    '  });' +
    '}' +

    /* Update filter visibility based on tab */
    'function updateFilterVisibility() {' +
    '  var statusFilter = document.getElementById("statusFilter");' +
    '  var stewardFilter = document.getElementById("stewardFilter");' +
    '  if (currentTab === "grievances") {' +
    '    statusFilter.style.display = "flex";' +
    '    stewardFilter.style.display = "none";' +
    '  } else if (currentTab === "members") {' +
    '    statusFilter.style.display = "none";' +
    '    stewardFilter.style.display = "flex";' +
    '  } else {' +
    '    statusFilter.style.display = "flex";' +
    '    stewardFilter.style.display = "flex";' +
    '  }' +
    '}' +

    /* Clear all filters */
    'function clearFilters() {' +
    '  document.getElementById("filterStatus").value = "";' +
    '  document.getElementById("filterLocation").value = "";' +
    '  document.getElementById("filterSteward").value = "";' +
    '  doSearch();' +
    '}' +

    /* Perform search */
    'function doSearch() {' +
    '  var query = document.getElementById("searchQuery").value.trim();' +
    '  var filters = {' +
    '    status: document.getElementById("filterStatus").value,' +
    '    location: document.getElementById("filterLocation").value,' +
    '    isSteward: document.getElementById("filterSteward").value' +
    '  };' +
    '  if (query.length < 2 && !filters.status && !filters.location && !filters.isSteward) {' +
    '    showEmptyState();' +
    '    return;' +
    '  }' +
    '  showLoading();' +
    '  google.script.run' +
    '    .withSuccessHandler(function(data) { renderResults(data); })' +
    '    .withFailureHandler(function(err) { showError(err.message); })' +
    '    .getDesktopSearchData(query, currentTab, filters);' +
    '}' +

    /* Show loading state */
    'function showLoading() {' +
    '  document.getElementById("resultsSection").innerHTML = ' +
    '    \'<div class="loading"><div class="spinner"></div><p style="margin-top:12px">Searching...</p></div>\';' +
    '}' +

    /* Show empty state */
    'function showEmptyState() {' +
    '  document.getElementById("resultsSection").innerHTML = ' +
    '    \'<div class="empty-state"><div class="empty-icon">🔍</div>\' +' +
    '    \'<div class="empty-title">Start searching</div>\' +' +
    '    \'<div class="empty-text">Type at least 2 characters to search</div></div>\';' +
    '}' +

    /* Show error */
    'function showError(msg) {' +
    '  document.getElementById("resultsSection").innerHTML = ' +
    '    \'<div class="empty-state"><div class="empty-icon">⚠️</div>\' +' +
    '    \'<div class="empty-title">Error</div>\' +' +
    '    \'<div class="empty-text">\' + msg + \'</div></div>\';' +
    '}' +

    /* Render results */
    'function renderResults(data) {' +
    '  currentResults = data || [];' +
    '  var container = document.getElementById("resultsSection");' +
    '  if (currentResults.length === 0) {' +
    '    container.innerHTML = \'<div class="empty-state"><div class="empty-icon">📭</div>\' +' +
    '      \'<div class="empty-title">No results found</div>\' +' +
    '      \'<div class="empty-text">Try adjusting your search or filters</div></div>\';' +
    '    return;' +
    '  }' +
    '  var html = \'<div class="results-header"><div class="results-count">\' + currentResults.length + \' result\' + (currentResults.length !== 1 ? "s" : "") + \' found</div></div>\';' +
    '  html += \'<div class="results-grid">\';' +
    '  currentResults.forEach(function(r, i) {' +
    '    var statusClass = "";' +
    '    if (r.status) {' +
    '      statusClass = r.status === "Open" ? "status-open" : (r.status.indexOf("Pending") >= 0 ? "status-pending" : "status-closed");' +
    '    }' +
    '    html += \'<div class="result-card" onclick="navigateToResult(\' + i + \')">\';' +
    '    html += \'<div class="card-header">\';' +
    '    html += \'<span class="card-type \' + (r.type === "member" ? "type-member" : "type-grievance") + \'">\' + (r.type === "member" ? "👤 Member" : "📋 Grievance") + \'</span>\';' +
    '    html += \'<span class="card-id">\' + r.id + \'</span>\';' +
    '    html += \'</div>\';' +
    '    html += \'<div class="card-title">\' + r.title + \'</div>\';' +
    '    html += \'<div class="card-details">\';' +
    '    if (r.email) html += \'<div class="detail-row"><span class="detail-icon">📧</span>\' + r.email + \'</div>\';' +
    '    if (r.location) html += \'<div class="detail-row"><span class="detail-icon">📍</span>\' + r.location + \'</div>\';' +
    '    if (r.jobTitle) html += \'<div class="detail-row"><span class="detail-icon">💼</span>\' + r.jobTitle + \'</div>\';' +
    '    if (r.status) html += \'<div class="detail-row"><span class="detail-icon">📌</span><span class="status-badge \' + statusClass + \'">\' + r.status + \'</span></div>\';' +
    '    if (r.issueType) html += \'<div class="detail-row"><span class="detail-icon">⚖️</span>\' + r.issueType + \'</div>\';' +
    '    if (r.steward) html += \'<div class="detail-row"><span class="detail-icon">🛡️</span>Steward: \' + r.steward + \'</div>\';' +
    '    if (r.filedDate) html += \'<div class="detail-row"><span class="detail-icon">📅</span>Filed: \' + r.filedDate + \'</div>\';' +
    '    if (r.isSteward === "Yes") html += \'<div class="detail-row"><span class="detail-icon">⭐</span>Union Steward</div>\';' +
    '    html += \'</div></div>\';' +
    '  });' +
    '  html += \'</div>\';' +
    '  container.innerHTML = html;' +
    '}' +

    /* Navigate to result */
    'function navigateToResult(index) {' +
    '  var r = currentResults[index];' +
    '  if (!r) return;' +
    '  google.script.run.withSuccessHandler(function() {' +
    '    google.script.host.close();' +
    '  }).navigateToSearchResult(r.type, r.id, r.row);' +
    '}' +
    '</script></body></html>'
  ).setWidth(900).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '🔍 Search Dashboard');
}

/**
 * Get locations for desktop search filter dropdown
 * @returns {Array} Array of unique locations
 */
function getDesktopSearchLocations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var locations = [];

  // Get locations from Member Directory
  var mSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  if (mSheet && mSheet.getLastRow() > 1) {
    var mData = mSheet.getRange(2, MEMBER_COLS.WORK_LOCATION, mSheet.getLastRow() - 1, 1).getValues();
    mData.forEach(function(row) {
      var loc = row[0];
      if (loc && locations.indexOf(loc) === -1) {
        locations.push(loc);
      }
    });
  }

  return locations.sort();
}

/**
 * Get search data for desktop search
 * Searches more fields than mobile: job title, location, issue type, etc.
 * @param {string} query - Search query
 * @param {string} tab - Tab filter: 'all', 'members', 'grievances'
 * @param {Object} filters - Additional filters: status, location, isSteward
 * @returns {Array} Array of search results
 */
function getDesktopSearchData(query, tab, filters) {
  var results = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var q = (query || '').toLowerCase();
  filters = filters || {};

  // Search Members
  if (tab === 'all' || tab === 'members') {
    var mSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (mSheet && mSheet.getLastRow() > 1) {
      var lastCol = Math.max(MEMBER_COLS.IS_STEWARD, MEMBER_COLS.WORK_LOCATION, MEMBER_COLS.JOB_TITLE, MEMBER_COLS.EMAIL);
      var mData = mSheet.getRange(2, 1, mSheet.getLastRow() - 1, lastCol).getValues();

      mData.forEach(function(row, index) {
        var memberId = row[MEMBER_COLS.MEMBER_ID - 1] || '';
        var firstName = row[MEMBER_COLS.FIRST_NAME - 1] || '';
        var lastName = row[MEMBER_COLS.LAST_NAME - 1] || '';
        var fullName = firstName + ' ' + lastName;
        var email = row[MEMBER_COLS.EMAIL - 1] || '';
        var jobTitle = row[MEMBER_COLS.JOB_TITLE - 1] || '';
        var location = row[MEMBER_COLS.WORK_LOCATION - 1] || '';
        var isSteward = row[MEMBER_COLS.IS_STEWARD - 1] || '';

        // Apply filters
        if (filters.location && location !== filters.location) return;
        if (filters.isSteward && isSteward !== filters.isSteward) return;

        // Search across fields
        var searchable = (memberId + ' ' + fullName + ' ' + email + ' ' + jobTitle + ' ' + location).toLowerCase();
        if (q.length >= 2 && searchable.indexOf(q) === -1) return;

        // Skip if no query and no filters
        if (q.length < 2 && !filters.location && !filters.isSteward) return;

        results.push({
          type: 'member',
          id: memberId,
          title: fullName.trim() || 'Unnamed Member',
          email: email,
          jobTitle: jobTitle,
          location: location,
          isSteward: isSteward,
          row: index + 2 // 1-indexed + header row
        });
      });
    }
  }

  // Search Grievances
  if (tab === 'all' || tab === 'grievances') {
    var gSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (gSheet && gSheet.getLastRow() > 1) {
      var lastGCol = Math.max(GRIEVANCE_COLS.STATUS, GRIEVANCE_COLS.ISSUE_CATEGORY, GRIEVANCE_COLS.LOCATION, GRIEVANCE_COLS.STEWARD, GRIEVANCE_COLS.DATE_FILED);
      var gData = gSheet.getRange(2, 1, gSheet.getLastRow() - 1, lastGCol).getValues();

      gData.forEach(function(row, index) {
        var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || '';
        var firstName = row[GRIEVANCE_COLS.FIRST_NAME - 1] || '';
        var lastName = row[GRIEVANCE_COLS.LAST_NAME - 1] || '';
        var fullName = firstName + ' ' + lastName;
        var status = row[GRIEVANCE_COLS.STATUS - 1] || '';
        var issueType = row[GRIEVANCE_COLS.ISSUE_CATEGORY - 1] || '';
        var location = row[GRIEVANCE_COLS.LOCATION - 1] || '';
        var steward = row[GRIEVANCE_COLS.STEWARD - 1] || '';
        var dateFiled = row[GRIEVANCE_COLS.DATE_FILED - 1] || '';

        // Apply filters
        if (filters.status && status !== filters.status) return;
        if (filters.location && location !== filters.location) return;

        // Search across fields
        var searchable = (grievanceId + ' ' + fullName + ' ' + status + ' ' + issueType + ' ' + location + ' ' + steward).toLowerCase();
        if (q.length >= 2 && searchable.indexOf(q) === -1) return;

        // Skip if no query and no filters
        if (q.length < 2 && !filters.status && !filters.location) return;

        // Format date
        var filedDateStr = '';
        if (dateFiled) {
          try {
            filedDateStr = Utilities.formatDate(new Date(dateFiled), Session.getScriptTimeZone(), 'MM/dd/yyyy');
          } catch(e) {
            filedDateStr = dateFiled.toString();
          }
        }

        results.push({
          type: 'grievance',
          id: grievanceId,
          title: fullName.trim() || 'Unknown Member',
          status: status,
          issueType: issueType,
          location: location,
          steward: steward,
          filedDate: filedDateStr,
          row: index + 2
        });
      });
    }
  }

  // Limit results
  return results.slice(0, 50);
}

/**
 * Navigate to a search result in the spreadsheet
 * @param {string} type - 'member' or 'grievance'
 * @param {string} id - The record ID
 * @param {number} row - The row number
 */
function navigateToSearchResult(type, id, row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = type === 'member' ? SHEETS.MEMBER_DIR : SHEETS.GRIEVANCE_LOG;
  var sheet = ss.getSheetByName(sheetName);

  if (sheet && row) {
    ss.setActiveSheet(sheet);
    sheet.setActiveRange(sheet.getRange(row, 1));
    SpreadsheetApp.flush();
  }
}

function viewActiveGrievances() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (sheet) {
    ss.setActiveSheet(sheet);
  }
}

function startNewGrievance() {
  SpreadsheetApp.getUi().alert('Start New Grievance feature - Coming soon!\n\nFor now, add grievances directly to the Grievance Log sheet.');
}

/**
 * Recalculate all grievance deadlines and sync to Member Directory
 * Uses hidden sheet formulas for self-healing calculations
 */
function recalcAllGrievancesBatched() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Recalculating grievances...', '🔄 Refresh', 3);

  // Step 1: Sync calculated formulas from hidden sheet to Grievance Log
  // This updates: Filing Deadline, Step Due dates, Days Open, Next Action Due, Days to Deadline
  // Also updates: First Name, Last Name, Email, Unit, Location, Steward from Member Directory
  syncGrievanceFormulasToLog();

  // Step 2: Repair checkboxes (in case they were overwritten)
  repairGrievanceCheckboxes();

  // Step 3: Sync grievance data to member directory (updates AB-AD columns)
  syncGrievanceToMemberDirectory();

  ss.toast('Grievance data refreshed!', '✅ Success', 3);
}

/**
 * Refresh Member Directory calculated columns (AB-AD: Has Open Grievance, Status, Next Deadline)
 */
function refreshMemberDirectoryFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Refreshing Member Directory...', '🔄 Refresh', 3);

  // Step 1: Refresh grievance formulas first (to get latest Next Action Due dates)
  syncGrievanceFormulasToLog();

  // Step 2: Sync grievance data to member directory (updates AB-AD columns)
  syncGrievanceToMemberDirectory();

  // Step 3: Repair member checkboxes
  repairMemberCheckboxes();

  ss.toast('Member Directory refreshed!', '✅ Success', 3);
}

function rebuildDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Rebuilding dashboard sheets...', '🔄 Rebuild', 3);

  // Recreate dashboard sheets with latest layout
  createDashboard(ss);
  createInteractiveDashboard(ss);

  // Refresh hidden sheet formulas and sync data
  refreshAllHiddenFormulas();

  // Reapply data validations
  setupDataValidations();

  ss.toast('Dashboard rebuilt with all 9 sections!', '✅ Success', 3);
}

/**
 * Refresh all formulas and sync all data
 */
function refreshAllFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Refreshing all formulas and syncing data...', '🔄 Refresh', 3);

  // Use the full refresh from HiddenSheets.gs
  refreshAllHiddenFormulas();
}

// ============================================================================
// VIEW CONTROLS - Timeline Simplification
// ============================================================================

/**
 * Simplify the Grievance Log timeline view
 * Hides Step II and Step III columns, keeping only essential dates
 * Shows: Incident Date, Date Filed, Date Closed, Days Open, Next Action Due, Days to Deadline
 */
function simplifyTimelineView() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Grievance Log sheet not found.');
    return;
  }

  ss.toast('Simplifying timeline view...', '👁️ View', 2);

  // Hide Step I detail columns (J-K): Step I Due, Step I Rcvd
  sheet.hideColumns(GRIEVANCE_COLS.STEP1_DUE, 2);

  // Hide Step II columns (L-O): Appeal Due, Appeal Filed, Due, Rcvd
  sheet.hideColumns(GRIEVANCE_COLS.STEP2_APPEAL_DUE, 4);

  // Hide Step III columns (P-Q): Appeal Due, Appeal Filed
  sheet.hideColumns(GRIEVANCE_COLS.STEP3_APPEAL_DUE, 2);

  // Hide Filing Deadline (H) - auto-calculated, less important once filed
  sheet.hideColumns(GRIEVANCE_COLS.FILING_DEADLINE, 1);

  ss.toast('Timeline simplified! Showing only key dates: Incident, Filed, Closed, Next Due', '✅ Done', 3);
}

/**
 * Show the full timeline view
 * Unhides all date columns
 */
function showFullTimelineView() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Grievance Log sheet not found.');
    return;
  }

  ss.toast('Showing full timeline...', '👁️ View', 2);

  // Show all timeline columns (H through Q)
  sheet.showColumns(GRIEVANCE_COLS.FILING_DEADLINE, 10); // H through Q

  ss.toast('Full timeline view restored!', '✅ Done', 3);
}

/**
 * Setup column groups for the timeline
 * Creates expandable/collapsible groups for Step II and Step III
 */
function setupTimelineColumnGroups() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Grievance Log sheet not found.');
    return;
  }

  ss.toast('Setting up column groups...', '👁️ View', 2);

  // Group Step I columns (J-K)
  var step1Range = sheet.getRange(1, GRIEVANCE_COLS.STEP1_DUE, 1, 2);
  sheet.getColumnGroup(GRIEVANCE_COLS.STEP1_DUE, 1);

  // Group Step II columns (L-O)
  sheet.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.AFTER);
  var step2Group = sheet.getRange(1, GRIEVANCE_COLS.STEP2_APPEAL_DUE, 1, 4);
  sheet.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.AFTER);

  // Group Step III columns (P-Q)
  var step3Group = sheet.getRange(1, GRIEVANCE_COLS.STEP3_APPEAL_DUE, 1, 2);

  // Create the groups
  try {
    sheet.getRange(1, GRIEVANCE_COLS.STEP1_DUE, sheet.getMaxRows(), 2).shiftColumnGroupDepth(1);
    sheet.getRange(1, GRIEVANCE_COLS.STEP2_APPEAL_DUE, sheet.getMaxRows(), 4).shiftColumnGroupDepth(1);
    sheet.getRange(1, GRIEVANCE_COLS.STEP3_APPEAL_DUE, sheet.getMaxRows(), 2).shiftColumnGroupDepth(1);

    // Collapse Step II and III by default (Step I usually visible)
    sheet.collapseAllColumnGroups();

    ss.toast('Column groups created! Click +/- to expand/collapse step details', '✅ Done', 5);
  } catch (e) {
    Logger.log('Column group error: ' + e.toString());
    ss.toast('Column groups may already exist or require manual setup', '⚠️ Note', 3);
  }
}

/**
 * Apply conditional formatting to highlight the current step's dates
 * Grays out dates for steps not yet reached
 */
function applyStepHighlighting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Grievance Log sheet not found.');
    return;
  }

  ss.toast('Applying step highlighting...', '🎨 Format', 3);

  var lastRow = Math.max(sheet.getLastRow(), 2);
  var rules = sheet.getConditionalFormatRules();

  // Colors
  var grayText = SpreadsheetApp.newColor().setRgbColor('#9e9e9e').build();
  var greenBg = SpreadsheetApp.newColor().setRgbColor('#e8f5e9').build();
  var currentStepCol = GRIEVANCE_COLS.CURRENT_STEP; // Column F

  // Rule 1: Gray out Step I columns (J-K) if current step is Informal
  var step1Range = sheet.getRange(2, GRIEVANCE_COLS.STEP1_DUE, lastRow - 1, 2);
  var rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2="Informal"')
    .setFontColor('#9e9e9e')
    .setRanges([step1Range])
    .build();

  // Rule 2: Gray out Step II columns (L-O) if current step is Informal or Step I
  var step2Range = sheet.getRange(2, GRIEVANCE_COLS.STEP2_APPEAL_DUE, lastRow - 1, 4);
  var rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($F2="Informal",$F2="Step I")')
    .setFontColor('#9e9e9e')
    .setRanges([step2Range])
    .build();

  // Rule 3: Gray out Step III columns (P-Q) if not at Step III or beyond
  var step3Range = sheet.getRange(2, GRIEVANCE_COLS.STEP3_APPEAL_DUE, lastRow - 1, 2);
  var rule3 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($F2="Informal",$F2="Step I",$F2="Step II")')
    .setFontColor('#9e9e9e')
    .setRanges([step3Range])
    .build();

  // -------------------------------------------------------------------------
  // DEADLINE STATUS RULES (Days to Deadline column U)
  // Order matters: more specific rules first, then broader ones
  // -------------------------------------------------------------------------

  var daysDeadlineRange = sheet.getRange(2, GRIEVANCE_COLS.DAYS_TO_DEADLINE, lastRow - 1, 1);
  var nextDueRange = sheet.getRange(2, GRIEVANCE_COLS.NEXT_ACTION_DUE, lastRow - 1, 1);

  // Rule 4: 🔴 Red - Overdue (Days to Deadline shows "Overdue" or negative/0)
  var rule4 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($U2="Overdue",AND(ISNUMBER($U2),$U2<=0))')
    .setBackground('#ffebee')
    .setFontColor('#c62828')
    .setBold(true)
    .setRanges([daysDeadlineRange])
    .build();

  // Rule 5: 🟠 Orange - Due in 1-3 days
  var rule5 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($U2),$U2>=1,$U2<=3)')
    .setBackground('#fff3e0')
    .setFontColor('#e65100')
    .setBold(true)
    .setRanges([daysDeadlineRange])
    .build();

  // Rule 6: 🟡 Yellow - Due in 4-7 days
  var rule6 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($U2),$U2>=4,$U2<=7)')
    .setBackground('#fffde7')
    .setFontColor('#f57f17')
    .setBold(false)
    .setRanges([daysDeadlineRange])
    .build();

  // Rule 7: 🟢 Green - On Track (more than 7 days remaining)
  var rule7 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(ISNUMBER($U2),$U2>7)')
    .setBackground('#e8f5e9')
    .setFontColor('#2e7d32')
    .setBold(false)
    .setRanges([daysDeadlineRange])
    .build();

  // Rule 8: Red highlight for Next Action Due if overdue
  var rule8 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($T2<>"",$T2<TODAY())')
    .setBackground('#ffebee')
    .setFontColor('#c62828')
    .setBold(true)
    .setRanges([nextDueRange])
    .build();

  // Rule 9: Orange for Next Action Due within 3 days
  var rule9 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($T2<>"",($T2-TODAY())>=0,($T2-TODAY())<=3)')
    .setBackground('#fff3e0')
    .setFontColor('#e65100')
    .setBold(true)
    .setRanges([nextDueRange])
    .build();

  // -------------------------------------------------------------------------
  // OUTCOME STATUS RULES (Status column E)
  // -------------------------------------------------------------------------

  var statusRange = sheet.getRange(2, GRIEVANCE_COLS.STATUS, lastRow - 1, 1);

  // Rule 10: ✅ Green - Won
  var rule10 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Won')
    .setBackground('#e8f5e9')
    .setFontColor('#2e7d32')
    .setBold(true)
    .setRanges([statusRange])
    .build();

  // Rule 11: ❌ Red - Denied
  var rule11 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Denied')
    .setBackground('#ffebee')
    .setFontColor('#c62828')
    .setBold(true)
    .setRanges([statusRange])
    .build();

  // Rule 12: 🤝 Blue - Settled
  var rule12 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Settled')
    .setBackground('#e3f2fd')
    .setFontColor('#1565c0')
    .setBold(true)
    .setRanges([statusRange])
    .build();

  // Add new rules (keep existing rules)
  rules.push(rule1, rule2, rule3, rule4, rule5, rule6, rule7, rule8, rule9, rule10, rule11, rule12);
  sheet.setConditionalFormatRules(rules);

  ss.toast('Formatting applied! Deadline colors (🟢🟡🟠🔴) and outcome status (Won/Denied/Settled)', '✅ Done', 5);
}

/**
 * Freeze key columns for easier scrolling
 * Freezes A-F (Identity & Status) so they're always visible
 */
function freezeKeyColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Grievance Log sheet not found.');
    return;
  }

  // Freeze first 6 columns (A-F: ID, Member ID, Name, Status, Step)
  sheet.setFrozenColumns(6);
  // Freeze header row
  sheet.setFrozenRows(1);

  ss.toast('Frozen columns A-F and header row. Scroll right to see timeline.', '❄️ Frozen', 3);
}

/**
 * Unfreeze all columns
 */
function unfreezeAllColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Grievance Log sheet not found.');
    return;
  }

  sheet.setFrozenColumns(0);
  // Keep header row frozen
  sheet.setFrozenRows(1);

  ss.toast('Columns unfrozen. Header row still frozen.', '🔓 Unfrozen', 3);
}

// ============================================================================
// TESTING FUNCTIONS
// ============================================================================

function viewTestResults() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.TEST_RESULTS);
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('No test results yet. Run tests first using 🧪 Testing menu.');
  }
}

// ============================================================================
// NAVIGATION FUNCTIONS (Menu Items)
// ============================================================================

/**
 * Refresh Custom View charts and data
 */
function refreshInteractiveCharts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Refreshing Custom View...', '📈 Refresh', 2);

  var sheet = ss.getSheetByName(SHEETS.INTERACTIVE);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Custom View not found. Run REPAIR DASHBOARD to create it.');
    return;
  }

  // Force recalculation by flushing
  SpreadsheetApp.flush();

  // Navigate to it
  ss.setActiveSheet(sheet);
  ss.toast('Custom View refreshed!', '✅ Done', 2);
}

// ============================================================================
// GOOGLE DRIVE INTEGRATION
// ============================================================================

/**
 * Create a Google Drive folder for the selected grievance
 */
function setupDriveFolderForGrievance() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  // Must be on Grievance Log
  if (sheet.getName() !== SHEETS.GRIEVANCE_LOG) {
    ui.alert('📁 Setup Folder', 'Please select a row in the Grievance Log first.', ui.ButtonSet.OK);
    return;
  }

  var row = sheet.getActiveRange().getRow();
  if (row < 2) {
    ui.alert('📁 Setup Folder', 'Please select a grievance row (not the header).', ui.ButtonSet.OK);
    return;
  }

  var grievanceId = sheet.getRange(row, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();
  var memberId = sheet.getRange(row, GRIEVANCE_COLS.MEMBER_ID).getValue();

  if (!grievanceId) {
    ui.alert('📁 Setup Folder', 'This row has no Grievance ID.', ui.ButtonSet.OK);
    return;
  }

  ss.toast('Creating Drive folder for ' + grievanceId + '...', '📁 Drive', 3);

  try {
    // Get or create root folder
    var rootFolder = getOrCreateDashboardFolder_();

    // Create grievance folder
    var folderName = grievanceId + ' - ' + (memberId || 'Unknown');
    var folders = rootFolder.getFoldersByName(folderName);
    var folder;

    if (folders.hasNext()) {
      folder = folders.next();
      ss.toast('Folder already exists!', '📁 Drive', 2);
    } else {
      folder = rootFolder.createFolder(folderName);
      ss.toast('Folder created!', '✅ Success', 2);
    }

    // Save folder ID and URL to grievance row
    sheet.getRange(row, GRIEVANCE_COLS.DRIVE_FOLDER_ID).setValue(folder.getId());
    sheet.getRange(row, GRIEVANCE_COLS.DRIVE_FOLDER_URL).setValue(folder.getUrl());

    // Open folder in new tab
    var html = HtmlService.createHtmlOutput(
      '<script>window.open("' + folder.getUrl() + '", "_blank");google.script.host.close();</script>'
    ).setWidth(1).setHeight(1);
    ui.showModalDialog(html, 'Opening folder...');

  } catch (e) {
    ui.alert('❌ Error', 'Failed to create folder: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Get or create the root 509 Dashboard folder in Drive
 */
function getOrCreateDashboardFolder_() {
  var folderName = '509 Dashboard - Grievance Files';
  var folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  return DriveApp.createFolder(folderName);
}

/**
 * Show files in the selected grievance's folder
 */
function showGrievanceFiles() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  if (sheet.getName() !== SHEETS.GRIEVANCE_LOG) {
    ui.alert('📁 View Files', 'Please select a row in the Grievance Log first.', ui.ButtonSet.OK);
    return;
  }

  var row = sheet.getActiveRange().getRow();
  if (row < 2) {
    ui.alert('📁 View Files', 'Please select a grievance row (not the header).', ui.ButtonSet.OK);
    return;
  }

  var folderId = sheet.getRange(row, GRIEVANCE_COLS.DRIVE_FOLDER_ID).getValue();
  var folderUrl = sheet.getRange(row, GRIEVANCE_COLS.DRIVE_FOLDER_URL).getValue();
  var grievanceId = sheet.getRange(row, GRIEVANCE_COLS.GRIEVANCE_ID).getValue();

  if (!folderId) {
    var response = ui.alert('📁 No Folder',
      'No folder exists for ' + grievanceId + '.\n\nWould you like to create one?',
      ui.ButtonSet.YES_NO);
    if (response === ui.Button.YES) {
      setupDriveFolderForGrievance();
    }
    return;
  }

  try {
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();
    var fileList = [];

    while (files.hasNext()) {
      var file = files.next();
      fileList.push('• ' + file.getName());
    }

    if (fileList.length === 0) {
      var response = ui.alert('📁 ' + grievanceId + ' Files',
        'Folder is empty.\n\nWould you like to open the folder to add files?',
        ui.ButtonSet.YES_NO);
      if (response === ui.Button.YES) {
        var html = HtmlService.createHtmlOutput(
          '<script>window.open("' + folderUrl + '", "_blank");google.script.host.close();</script>'
        ).setWidth(1).setHeight(1);
        ui.showModalDialog(html, 'Opening folder...');
      }
    } else {
      var response = ui.alert('📁 ' + grievanceId + ' Files (' + fileList.length + ')',
        fileList.join('\n') + '\n\nOpen folder in Drive?',
        ui.ButtonSet.YES_NO);
      if (response === ui.Button.YES) {
        var html = HtmlService.createHtmlOutput(
          '<script>window.open("' + folderUrl + '", "_blank");google.script.host.close();</script>'
        ).setWidth(1).setHeight(1);
        ui.showModalDialog(html, 'Opening folder...');
      }
    }
  } catch (e) {
    ui.alert('❌ Error', 'Could not access folder: ' + e.message + '\n\nThe folder may have been deleted.', ui.ButtonSet.OK);
  }
}

/**
 * Batch create folders for all grievances without folders
 */
function batchCreateGrievanceFolders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    ui.alert('❌ Error', 'Grievance Log not found.', ui.ButtonSet.OK);
    return;
  }

  var response = ui.alert('📁 Batch Create Folders',
    'This will create Google Drive folders for all grievances that don\'t have one.\n\nContinue?',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) return;

  var data = sheet.getDataRange().getValues();
  var rootFolder = getOrCreateDashboardFolder_();
  var created = 0;
  var skipped = 0;

  ss.toast('Creating folders...', '📁 Batch', -1);

  for (var i = 1; i < data.length; i++) {
    var grievanceId = data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    var memberId = data[i][GRIEVANCE_COLS.MEMBER_ID - 1];
    var existingFolderId = data[i][GRIEVANCE_COLS.DRIVE_FOLDER_ID - 1];

    if (!grievanceId) continue;

    if (existingFolderId) {
      skipped++;
      continue;
    }

    var folderName = grievanceId + ' - ' + (memberId || 'Unknown');
    var folder = rootFolder.createFolder(folderName);

    sheet.getRange(i + 1, GRIEVANCE_COLS.DRIVE_FOLDER_ID).setValue(folder.getId());
    sheet.getRange(i + 1, GRIEVANCE_COLS.DRIVE_FOLDER_URL).setValue(folder.getUrl());
    created++;

    if (created % 5 === 0) {
      ss.toast('Created ' + created + ' folders...', '📁 Batch', 2);
    }
  }

  ss.toast('Done! Created ' + created + ', skipped ' + skipped, '✅ Complete', 5);
  ui.alert('📁 Batch Complete',
    'Created: ' + created + ' new folders\n' +
    'Skipped: ' + skipped + ' (already had folders)\n\n' +
    'Root folder: ' + rootFolder.getUrl(),
    ui.ButtonSet.OK);
}

// ============================================================================
// GOOGLE CALENDAR INTEGRATION
// ============================================================================

/**
 * Sync grievance deadlines to Google Calendar with rate limit handling
 */
function syncDeadlinesToCalendar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    ui.alert('❌ Error', 'Grievance Log not found.', ui.ButtonSet.OK);
    return;
  }

  var data = sheet.getDataRange().getValues();
  var eventsToCreate = [];
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  // First pass: count events to create
  for (var i = 1; i < data.length; i++) {
    var grievanceId = data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    var status = data[i][GRIEVANCE_COLS.STATUS - 1];
    var nextActionDue = data[i][GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];
    var currentStep = data[i][GRIEVANCE_COLS.CURRENT_STEP - 1];

    if (!grievanceId || !nextActionDue) continue;

    var closedStatuses = ['Closed', 'Settled', 'Won', 'Denied', 'Withdrawn'];
    if (closedStatuses.indexOf(status) !== -1) continue;

    var dueDate = new Date(nextActionDue);
    if (isNaN(dueDate.getTime()) || dueDate < today) continue;

    eventsToCreate.push({
      grievanceId: grievanceId,
      status: status,
      currentStep: currentStep,
      dueDate: dueDate
    });
  }

  // Warn if too many events
  if (eventsToCreate.length > 50) {
    var response = ui.alert('⚠️ Large Sync Warning',
      'You are about to create up to ' + eventsToCreate.length + ' calendar events.\n\n' +
      'Creating too many events in a short time may trigger Google\'s rate limits.\n\n' +
      'Recommended: Sync in batches of 50 or less.\n\n' +
      'Continue with first 50 events only?',
      ui.ButtonSet.YES_NO);

    if (response !== ui.Button.YES) return;
    eventsToCreate = eventsToCreate.slice(0, 50);
  } else {
    var response = ui.alert('📅 Sync to Calendar',
      'This will create calendar events for ' + eventsToCreate.length + ' upcoming grievance deadlines.\n\n' +
      'Events will be created in your primary Google Calendar.\n\nContinue?',
      ui.ButtonSet.YES_NO);

    if (response !== ui.Button.YES) return;
  }

  ss.toast('Syncing deadlines to calendar...', '📅 Calendar', -1);

  var calendar;
  try {
    calendar = CalendarApp.getDefaultCalendar();
  } catch (error) {
    ui.alert('❌ Calendar Error',
      'Could not access your calendar.\n\n' +
      'Error: ' + error.message + '\n\n' +
      'Make sure you have authorized calendar access.',
      ui.ButtonSet.OK);
    return;
  }

  var created = 0;
  var skipped = 0;
  var rateLimited = false;

  for (var j = 0; j < eventsToCreate.length; j++) {
    var event = eventsToCreate[j];

    try {
      // Check if event already exists
      var eventTitle = '⚠️ ' + event.grievanceId + ' - ' + event.currentStep + ' Due';
      var existingEvents = calendar.getEventsForDay(event.dueDate, {search: event.grievanceId});

      if (existingEvents.length > 0) {
        skipped++;
        continue;
      }

      // Create all-day event
      calendar.createAllDayEvent(eventTitle, event.dueDate, {
        description: 'Grievance: ' + event.grievanceId + '\nStatus: ' + event.status + '\nStep: ' + event.currentStep + '\n\nCreated by 509 Dashboard'
      });
      created++;

      // Throttle to avoid rate limits (1 per 100ms)
      if (j < eventsToCreate.length - 1) {
        Utilities.sleep(100);
      }

    } catch (error) {
      if (error.message.indexOf('too many') !== -1 ||
          error.message.indexOf('rate') !== -1 ||
          error.message.indexOf('quota') !== -1) {
        rateLimited = true;
        Logger.log('Rate limit hit at event ' + (j + 1));
        break;
      }
      Logger.log('Error creating event for ' + event.grievanceId + ': ' + error.message);
      skipped++;
    }
  }

  ss.toast('Done! Created ' + created + ' events', '✅ Complete', 5);

  if (rateLimited) {
    ui.alert('⚠️ Rate Limit Reached',
      'Google Calendar rate limit was reached.\n\n' +
      'Created: ' + created + ' events before limit\n' +
      'Remaining: ' + (eventsToCreate.length - j) + ' events not synced\n\n' +
      'Please wait 1-2 minutes and try again to sync remaining events.',
      ui.ButtonSet.OK);
  } else {
    ui.alert('📅 Sync Complete',
      'Created: ' + created + ' calendar events\n' +
      'Skipped: ' + skipped + ' (already exists or error)\n' +
      'Total processed: ' + eventsToCreate.length,
      ui.ButtonSet.OK);
  }
}

/**
 * Show upcoming deadlines from calendar with member names
 */
function showUpcomingDeadlinesFromCalendar() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    var calendar = CalendarApp.getDefaultCalendar();
    var today = new Date();
    var nextWeek = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000);

    var events = calendar.getEvents(today, nextWeek, {search: 'Grievance'});

    if (events.length === 0) {
      ui.alert('📅 Upcoming Deadlines',
        'No grievance deadlines in the next 7 days!\n\n' +
        'Use "Sync Deadlines to Calendar" to add deadline events.',
        ui.ButtonSet.OK);
      return;
    }

    // Build a lookup of grievance IDs to member names
    var memberLookup = buildGrievanceMemberLookup();

    var eventList = events.map(function(e) {
      var date = Utilities.formatDate(e.getStartTime(), Session.getScriptTimeZone(), 'MM/dd');
      var title = e.getTitle();

      // Extract grievance ID from title (format: "Grievance GR-XXXX: Step X Due")
      var match = title.match(/Grievance\s+(GR-\d+)/i);
      var memberInfo = '';

      if (match && match[1] && memberLookup[match[1]]) {
        memberInfo = ' (' + memberLookup[match[1]] + ')';
      }

      return '• ' + date + ': ' + title + memberInfo;
    });

    ui.alert('📅 Upcoming Deadlines (Next 7 Days)',
      'Events with member names:\n\n' + eventList.join('\n'),
      ui.ButtonSet.OK);

  } catch (error) {
    if (error.message.indexOf('too many') !== -1 || error.message.indexOf('rate') !== -1) {
      ui.alert('⚠️ Calendar Rate Limit',
        'Google Calendar is temporarily limiting requests.\n\n' +
        'Please wait a few minutes and try again.\n\n' +
        'Tip: Avoid running calendar operations repeatedly in quick succession.',
        ui.ButtonSet.OK);
    } else {
      ui.alert('❌ Calendar Error', error.message, ui.ButtonSet.OK);
    }
  }
}

/**
 * Build a lookup map of grievance IDs to member names
 * @return {Object} Map of grievanceId -> "First Last"
 */
function buildGrievanceMemberLookup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var lookup = {};

  if (!sheet || sheet.getLastRow() <= 1) return lookup;

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, GRIEVANCE_COLS.LAST_NAME).getValues();

  data.forEach(function(row) {
    var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    var firstName = row[GRIEVANCE_COLS.FIRST_NAME - 1] || '';
    var lastName = row[GRIEVANCE_COLS.LAST_NAME - 1] || '';

    if (grievanceId) {
      lookup[grievanceId] = (firstName + ' ' + lastName).trim() || 'Unknown';
    }
  });

  return lookup;
}

/**
 * Clear all 509 Dashboard calendar events
 */
function clearAllCalendarEvents() {
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert('🗑️ Clear Calendar Events',
    'This will delete ALL calendar events containing "Grievance" from the past 6 months to 1 year ahead.\n\n' +
    'This cannot be undone. Continue?',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Removing calendar events...', '🗑️ Calendar', -1);

  var calendar = CalendarApp.getDefaultCalendar();
  var sixMonthsAgo = new Date(Date.now() - 180 * 24 * 60 * 60 * 1000);
  var oneYearAhead = new Date(Date.now() + 365 * 24 * 60 * 60 * 1000);

  var events = calendar.getEvents(sixMonthsAgo, oneYearAhead, {search: 'Grievance'});
  var deleted = 0;

  for (var i = 0; i < events.length; i++) {
    events[i].deleteEvent();
    deleted++;
  }

  ss.toast('Deleted ' + deleted + ' events', '✅ Complete', 3);
  ui.alert('🗑️ Events Cleared', 'Deleted ' + deleted + ' grievance calendar events.', ui.ButtonSet.OK);
}

// ============================================================================
// EMAIL NOTIFICATIONS
// ============================================================================

/**
 * Show notification settings dialog
 */
function showNotificationSettings() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var enabled = props.getProperty('notifications_enabled') === 'true';
  var email = props.getProperty('notification_email') || Session.getEffectiveUser().getEmail();

  var response = ui.alert('📬 Notification Settings',
    'Daily deadline notifications: ' + (enabled ? 'ENABLED ✅' : 'DISABLED ❌') + '\n' +
    'Email: ' + email + '\n\n' +
    'Notifications are sent daily at 8 AM for grievances due within 3 days.\n\n' +
    'Would you like to ' + (enabled ? 'DISABLE' : 'ENABLE') + ' notifications?',
    ui.ButtonSet.YES_NO);

  if (response === ui.Button.YES) {
    if (enabled) {
      // Disable
      props.setProperty('notifications_enabled', 'false');
      removeDailyTrigger_();
      ui.alert('📬 Notifications Disabled', 'Daily deadline notifications have been turned off.', ui.ButtonSet.OK);
    } else {
      // Enable
      props.setProperty('notifications_enabled', 'true');
      props.setProperty('notification_email', email);
      installDailyTrigger_();
      ui.alert('📬 Notifications Enabled',
        'Daily notifications enabled!\n\n' +
        'You will receive an email at 8 AM when grievances are due within 3 days.\n\n' +
        'Email: ' + email, ui.ButtonSet.OK);
    }
  }
}

/**
 * Install daily trigger for notifications
 */
function installDailyTrigger_() {
  // Remove existing triggers
  removeDailyTrigger_();

  // Create new daily trigger at 8 AM
  ScriptApp.newTrigger('checkDeadlinesAndNotify_')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();
}

/**
 * Remove daily notification trigger
 */
function removeDailyTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkDeadlinesAndNotify_') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * Check deadlines and send notification email (called by trigger)
 */
function checkDeadlinesAndNotify_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('notifications_enabled') !== 'true') return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  if (!sheet) return;

  var email = props.getProperty('notification_email');
  if (!email) return;

  var data = sheet.getDataRange().getValues();
  var today = new Date();
  var threeDaysAhead = new Date(today.getTime() + 3 * 24 * 60 * 60 * 1000);
  var urgent = [];

  for (var i = 1; i < data.length; i++) {
    var grievanceId = data[i][GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    var status = data[i][GRIEVANCE_COLS.STATUS - 1];
    var daysToDeadline = data[i][GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
    var currentStep = data[i][GRIEVANCE_COLS.CURRENT_STEP - 1];

    var closedStatuses = ['Closed', 'Settled', 'Won', 'Denied', 'Withdrawn'];
    if (closedStatuses.indexOf(status) !== -1) continue;

    if (daysToDeadline !== '' && daysToDeadline <= 3) {
      urgent.push({
        id: grievanceId,
        step: currentStep,
        days: daysToDeadline
      });
    }
  }

  if (urgent.length === 0) return;

  var subject = '⚠️ 509 Dashboard: ' + urgent.length + ' Grievance Deadline(s) Approaching';
  var body = 'The following grievances have deadlines within 3 days:\n\n';

  for (var j = 0; j < urgent.length; j++) {
    var g = urgent[j];
    body += '• ' + g.id + ' (' + g.step + ') - ' +
      (g.days <= 0 ? 'OVERDUE!' : g.days + ' day(s) remaining') + '\n';
  }

  body += '\n\nView your dashboard: ' + ss.getUrl();

  MailApp.sendEmail(email, subject, body);
}

/**
 * Test the notification system
 */
function testDeadlineNotifications() {
  var ui = SpreadsheetApp.getUi();
  var email = Session.getEffectiveUser().getEmail();

  var response = ui.alert('🧪 Test Notifications',
    'This will send a test notification email to:\n' + email + '\n\nSend test email?',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) return;

  try {
    MailApp.sendEmail(email,
      '🧪 509 Dashboard Test Notification',
      'This is a test notification from your 509 Dashboard.\n\n' +
      'If you received this email, notifications are working correctly!\n\n' +
      'Dashboard: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
    );
    ui.alert('✅ Test Sent', 'Test email sent to ' + email + '\n\nCheck your inbox!', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ Error', 'Failed to send test email: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * Send daily digest to all stewards with their assigned grievance deadlines
 * Each steward gets their own personalized email
 */
function sendStewardDeadlineAlerts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet || !memberSheet) {
    Logger.log('Required sheets not found for steward alerts');
    return;
  }

  var props = PropertiesService.getScriptProperties();
  var alertDays = parseInt(props.getProperty('alert_days') || '7', 10);

  var grievanceData = sheet.getDataRange().getValues();
  var memberData = memberSheet.getDataRange().getValues();
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  // Build member lookup for steward emails
  var memberLookup = {};
  for (var m = 1; m < memberData.length; m++) {
    var memberId = memberData[m][MEMBER_COLS.MEMBER_ID - 1];
    if (memberId) {
      memberLookup[memberId] = {
        name: (memberData[m][MEMBER_COLS.FIRST_NAME - 1] || '') + ' ' + (memberData[m][MEMBER_COLS.LAST_NAME - 1] || ''),
        steward: memberData[m][MEMBER_COLS.ASSIGNED_STEWARD - 1] || ''
      };
    }
  }

  // Group grievances by steward
  var stewardGrievances = {};
  var closedStatuses = ['Closed', 'Settled', 'Won', 'Denied', 'Withdrawn'];

  for (var i = 1; i < grievanceData.length; i++) {
    var row = grievanceData[i];
    var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    var memberId = row[GRIEVANCE_COLS.MEMBER_ID - 1];
    var status = row[GRIEVANCE_COLS.STATUS - 1];
    var currentStep = row[GRIEVANCE_COLS.CURRENT_STEP - 1];
    var nextDue = row[GRIEVANCE_COLS.NEXT_ACTION_DUE - 1];
    var daysToDeadline = row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
    var steward = row[GRIEVANCE_COLS.ASSIGNED_STEWARD - 1] || '';

    // Skip closed grievances
    if (closedStatuses.indexOf(status) !== -1) continue;
    if (!grievanceId) continue;

    // Check if deadline is within alert window
    var daysRemaining = null;
    if (daysToDeadline === 'Overdue') {
      daysRemaining = -1;
    } else if (typeof daysToDeadline === 'number') {
      daysRemaining = daysToDeadline;
    } else {
      continue; // No deadline
    }

    if (daysRemaining > alertDays) continue;

    // Get member info
    var memberInfo = memberLookup[memberId] || { name: 'Unknown', steward: '' };
    var assignedSteward = steward || memberInfo.steward || 'Unassigned';

    if (!stewardGrievances[assignedSteward]) {
      stewardGrievances[assignedSteward] = [];
    }

    stewardGrievances[assignedSteward].push({
      id: grievanceId,
      memberName: memberInfo.name,
      step: currentStep,
      status: status,
      daysRemaining: daysRemaining,
      nextDue: nextDue
    });
  }

  // Get steward emails from Config sheet
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  var stewardEmails = {};
  if (configSheet) {
    var configData = configSheet.getDataRange().getValues();
    // Look for Steward Emails column (assume it's after Stewards column)
    for (var c = 1; c < configData.length; c++) {
      var stewardName = configData[c][CONFIG_COLS.STEWARDS - 1];
      var stewardEmail = configData[c][CONFIG_COLS.STEWARDS]; // Next column
      if (stewardName && stewardEmail && stewardEmail.indexOf('@') !== -1) {
        stewardEmails[stewardName] = stewardEmail;
      }
    }
  }

  // Send emails to each steward
  var emailsSent = 0;
  var adminEmail = Session.getEffectiveUser().getEmail();

  for (var stewardName in stewardGrievances) {
    var grievances = stewardGrievances[stewardName];
    if (grievances.length === 0) continue;

    // Sort by days remaining (most urgent first)
    grievances.sort(function(a, b) { return a.daysRemaining - b.daysRemaining; });

    var email = stewardEmails[stewardName] || adminEmail;

    // Build email body
    var overdue = grievances.filter(function(g) { return g.daysRemaining < 0; });
    var urgent = grievances.filter(function(g) { return g.daysRemaining >= 0 && g.daysRemaining <= 3; });
    var upcoming = grievances.filter(function(g) { return g.daysRemaining > 3; });

    var body = '📋 509 GRIEVANCE DEADLINE ALERT\n';
    body += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n';
    body += 'Steward: ' + stewardName + '\n';
    body += 'Date: ' + Utilities.formatDate(today, Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy') + '\n\n';

    if (overdue.length > 0) {
      body += '🔴 OVERDUE (' + overdue.length + ')\n';
      body += '─────────────────────\n';
      for (var o = 0; o < overdue.length; o++) {
        body += '  ⚠️ ' + overdue[o].id + ' - ' + overdue[o].memberName + '\n';
        body += '     Step: ' + overdue[o].step + ' | Status: ' + overdue[o].status + '\n';
        body += '     OVERDUE by ' + Math.abs(overdue[o].daysRemaining) + ' day(s)\n\n';
      }
    }

    if (urgent.length > 0) {
      body += '🟠 URGENT - Due within 3 days (' + urgent.length + ')\n';
      body += '─────────────────────\n';
      for (var u = 0; u < urgent.length; u++) {
        body += '  ⏰ ' + urgent[u].id + ' - ' + urgent[u].memberName + '\n';
        body += '     Step: ' + urgent[u].step + ' | Status: ' + urgent[u].status + '\n';
        body += '     Due in ' + urgent[u].daysRemaining + ' day(s)\n\n';
      }
    }

    if (upcoming.length > 0) {
      body += '🟡 UPCOMING - Due within ' + alertDays + ' days (' + upcoming.length + ')\n';
      body += '─────────────────────\n';
      for (var up = 0; up < upcoming.length; up++) {
        body += '  📅 ' + upcoming[up].id + ' - ' + upcoming[up].memberName + '\n';
        body += '     Step: ' + upcoming[up].step + ' | Due in ' + upcoming[up].daysRemaining + ' day(s)\n\n';
      }
    }

    body += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
    body += '📊 Dashboard: ' + ss.getUrl() + '\n';
    body += 'Total grievances requiring attention: ' + grievances.length + '\n';

    var subject = (overdue.length > 0 ? '🔴 OVERDUE: ' : '⏰ ') +
      grievances.length + ' Grievance Deadline(s) - ' + stewardName;

    try {
      MailApp.sendEmail({
        to: email,
        subject: subject,
        body: body,
        name: 'SEIU Local 509 Dashboard'
      });
      emailsSent++;
      Logger.log('Sent alert to ' + stewardName + ' (' + email + '): ' + grievances.length + ' grievances');
    } catch (e) {
      Logger.log('Failed to send to ' + email + ': ' + e.message);
    }
  }

  Logger.log('Steward deadline alerts complete. Sent ' + emailsSent + ' emails.');
  return emailsSent;
}

/**
 * Manual trigger to send steward alerts now
 */
function sendStewardAlertsNow() {
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert('📬 Send Steward Alerts',
    'This will send deadline alert emails to all stewards with upcoming deadlines.\n\n' +
    'Each steward will receive their own personalized digest.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) return;

  var emailsSent = sendStewardDeadlineAlerts();

  ui.alert('✅ Alerts Sent',
    'Sent ' + emailsSent + ' steward alert email(s).\n\n' +
    'Check the Logs for details.',
    ui.ButtonSet.OK);
}

/**
 * Configure alert settings
 */
function configureAlertSettings() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();

  var currentDays = props.getProperty('alert_days') || '7';
  var stewardAlerts = props.getProperty('steward_alerts_enabled') === 'true';

  var response = ui.prompt('⚙️ Alert Settings',
    'Current settings:\n' +
    '• Alert window: ' + currentDays + ' days before deadline\n' +
    '• Per-steward alerts: ' + (stewardAlerts ? 'ENABLED' : 'DISABLED') + '\n\n' +
    'Enter new alert window (days before deadline):\n' +
    '(Enter 3, 7, 14, or 30)',
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) return;

  var newDays = parseInt(response.getResponseText(), 10);
  if (isNaN(newDays) || newDays < 1 || newDays > 30) {
    ui.alert('Invalid input. Please enter a number between 1 and 30.');
    return;
  }

  props.setProperty('alert_days', newDays.toString());

  // Ask about per-steward alerts
  var stewardResponse = ui.alert('Per-Steward Alerts',
    'Enable per-steward email alerts?\n\n' +
    'When enabled, each steward receives their own personalized deadline digest.\n\n' +
    'Enable per-steward alerts?',
    ui.ButtonSet.YES_NO);

  props.setProperty('steward_alerts_enabled', stewardResponse === ui.Button.YES ? 'true' : 'false');

  ui.alert('✅ Settings Saved',
    'Alert window: ' + newDays + ' days\n' +
    'Per-steward alerts: ' + (stewardResponse === ui.Button.YES ? 'ENABLED' : 'DISABLED'),
    ui.ButtonSet.OK);
}

// ============================================================================
// AUDIT LOGGING - Multi-Steward Accountability
// ============================================================================

/**
 * Setup the hidden audit log sheet
 * Tracks all changes to Member Directory and Grievance Log
 */
function setupAuditLogSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.AUDIT_LOG);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.AUDIT_LOG);
  }

  sheet.clear();

  // Headers
  var headers = [
    'Timestamp',
    'User Email',
    'Sheet',
    'Row',
    'Column',
    'Field Name',
    'Old Value',
    'New Value',
    'Record ID',
    'Action Type'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format header row
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground(COLORS.PRIMARY_PURPLE);
  headerRange.setFontColor(COLORS.WHITE);
  headerRange.setFontWeight('bold');
  sheet.setFrozenRows(1);

  // Set column widths
  sheet.setColumnWidth(1, 160); // Timestamp
  sheet.setColumnWidth(2, 200); // User Email
  sheet.setColumnWidth(3, 120); // Sheet
  sheet.setColumnWidth(4, 50);  // Row
  sheet.setColumnWidth(5, 50);  // Column
  sheet.setColumnWidth(6, 150); // Field Name
  sheet.setColumnWidth(7, 200); // Old Value
  sheet.setColumnWidth(8, 200); // New Value
  sheet.setColumnWidth(9, 100); // Record ID
  sheet.setColumnWidth(10, 100); // Action Type

  // Hide the sheet
  sheet.hideSheet();

  SpreadsheetApp.getActiveSpreadsheet().toast('Audit log sheet created and hidden.', '✅ Setup Complete', 3);
}

/**
 * Log an audit event
 * @param {string} sheetName - Name of the sheet where change occurred
 * @param {number} row - Row number
 * @param {number} col - Column number
 * @param {string} fieldName - Name of the field/column
 * @param {string} oldValue - Previous value
 * @param {string} newValue - New value
 * @param {string} recordId - ID of the record (Member ID or Grievance ID)
 * @param {string} actionType - Type of action (Edit, Delete, Create)
 */
function logAuditEvent(sheetName, row, col, fieldName, oldValue, newValue, recordId, actionType) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var auditSheet = ss.getSheetByName(SHEETS.AUDIT_LOG);

    if (!auditSheet) {
      setupAuditLogSheet();
      auditSheet = ss.getSheetByName(SHEETS.AUDIT_LOG);
    }

    var userEmail = Session.getEffectiveUser().getEmail();
    var timestamp = new Date();

    var logEntry = [
      timestamp,
      userEmail,
      sheetName,
      row,
      col,
      fieldName,
      String(oldValue || ''),
      String(newValue || ''),
      recordId || '',
      actionType || 'Edit'
    ];

    auditSheet.appendRow(logEntry);
  } catch (e) {
    Logger.log('Audit log error: ' + e.message);
  }
}

/**
 * onEdit trigger for audit logging
 * Tracks changes to Member Directory and Grievance Log
 */
function onEditAudit(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();

  // Only track changes to Member Directory and Grievance Log
  if (sheetName !== SHEETS.MEMBER_DIR && sheetName !== SHEETS.GRIEVANCE_LOG) {
    return;
  }

  var row = e.range.getRow();
  var col = e.range.getColumn();

  // Skip header row
  if (row < 2) return;

  var oldValue = e.oldValue || '';
  var newValue = e.value || '';

  // Skip if no actual change
  if (oldValue === newValue) return;

  // Get field name from header
  var fieldName = sheet.getRange(1, col).getValue() || ('Column ' + col);

  // Get record ID (column A for both sheets)
  var recordId = sheet.getRange(row, 1).getValue() || '';

  // Determine action type
  var actionType = 'Edit';
  if (!oldValue && newValue) {
    actionType = 'Create';
  } else if (oldValue && !newValue) {
    actionType = 'Delete';
  }

  logAuditEvent(sheetName, row, col, fieldName, oldValue, newValue, recordId, actionType);
}

/**
 * Install the audit trigger
 */
function installAuditTrigger() {
  // Remove existing audit triggers
  removeAuditTrigger();

  // Create new onEdit trigger
  ScriptApp.newTrigger('onEditAudit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  // Ensure audit sheet exists
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEETS.AUDIT_LOG)) {
    setupAuditLogSheet();
  }

  SpreadsheetApp.getUi().alert('✅ Audit Tracking Enabled',
    'All changes to Member Directory and Grievance Log will now be logged.\n\n' +
    'View the audit log via:\n⚙️ Administrator > 📋 Audit Log > 📋 View Audit Log',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Remove the audit trigger
 */
function removeAuditTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onEditAudit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Audit tracking disabled.', '🚫 Disabled', 3);
}

/**
 * View the audit log sheet
 */
function viewAuditLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.AUDIT_LOG);

  if (!sheet) {
    var response = SpreadsheetApp.getUi().alert('📋 Audit Log Not Found',
      'The audit log sheet does not exist yet.\n\nWould you like to create it now?',
      SpreadsheetApp.getUi().ButtonSet.YES_NO);

    if (response === SpreadsheetApp.getUi().Button.YES) {
      setupAuditLogSheet();
      sheet = ss.getSheetByName(SHEETS.AUDIT_LOG);
    } else {
      return;
    }
  }

  // Show the hidden sheet temporarily
  sheet.showSheet();
  ss.setActiveSheet(sheet);

  // Sort by timestamp descending (newest first)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).sort({column: 1, ascending: false});
  }

  SpreadsheetApp.getUi().alert('📋 Audit Log',
    'Viewing audit log.\n\n' +
    'Total entries: ' + Math.max(0, sheet.getLastRow() - 1) + '\n\n' +
    'The sheet will be hidden again when you navigate away.\n' +
    'To keep it visible, right-click the tab and select "Unhide".',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Clear audit entries older than 30 days
 */
function clearOldAuditEntries() {
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert('🗑️ Clear Old Audit Entries',
    'This will delete all audit entries older than 30 days.\n\n' +
    'This action cannot be undone.\n\nContinue?',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.AUDIT_LOG);

  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert('No audit entries to clear.');
    return;
  }

  var cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - 30);

  var data = sheet.getDataRange().getValues();
  var rowsToDelete = [];

  // Find rows older than 30 days (skip header)
  for (var i = data.length - 1; i >= 1; i--) {
    var timestamp = data[i][0];
    if (timestamp instanceof Date && timestamp < cutoffDate) {
      rowsToDelete.push(i + 1); // +1 for 1-indexed rows
    }
  }

  // Delete rows from bottom to top to maintain correct indices
  for (var j = 0; j < rowsToDelete.length; j++) {
    sheet.deleteRow(rowsToDelete[j]);
  }

  ui.alert('✅ Cleanup Complete',
    'Deleted ' + rowsToDelete.length + ' entries older than 30 days.\n\n' +
    'Remaining entries: ' + Math.max(0, sheet.getLastRow() - 1),
    ui.ButtonSet.OK);
}

/**
 * Get audit summary for a specific record
 * @param {string} recordId - Member ID or Grievance ID
 * @returns {Array} Array of audit entries for this record
 */
function getAuditHistory(recordId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.AUDIT_LOG);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  var data = sheet.getDataRange().getValues();
  var history = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][8] === recordId) { // Column I is Record ID
      history.push({
        timestamp: data[i][0],
        user: data[i][1],
        field: data[i][5],
        oldValue: data[i][6],
        newValue: data[i][7],
        action: data[i][9]
      });
    }
  }

  return history;
}

// ============================================================================
// GRIEVANCE TOOLS - ADDITIONAL FUNCTIONS
// ============================================================================

/**
 * Setup Live Grievance Links - Syncs grievance data to Member Directory
 * Uses static values (no formulas) to avoid #REF! errors
 */
function setupLiveGrievanceFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error: Member Directory not found.');
    return;
  }

  ss.toast('Syncing grievance data to Member Directory...', '🔄 Sync', 3);

  // Use sync function to populate with static values (no formulas)
  syncGrievanceToMemberDirectory();

  ss.toast('Grievance data synced! Columns AB-AD updated with static values.', '✅ Success', 3);
}

/**
 * Setup Member ID dropdown in Grievance Log
 * Creates data validation that references Member Directory Member IDs
 */
function setupGrievanceMemberDropdown() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!memberSheet || !grievanceSheet) {
    SpreadsheetApp.getUi().alert('Error: Required sheets not found.');
    return;
  }

  ss.toast('Setting up Member ID dropdown...', '🔄 Setup', 3);

  // Get column letter for Member ID in Member Directory
  var mMemberIdCol = getColumnLetter(MEMBER_COLS.MEMBER_ID);

  // Create data validation rule that references Member Directory Member IDs
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(memberSheet.getRange(mMemberIdCol + '2:' + mMemberIdCol), true)
    .setAllowInvalid(true)  // Allow manual entry too
    .build();

  // Apply to Member ID column in Grievance Log (column B, rows 2-1000)
  grievanceSheet.getRange(2, GRIEVANCE_COLS.MEMBER_ID, 998, 1).setDataValidation(rule);

  ss.toast('Member ID dropdown set up!', '✅ Success', 3);
}

/**
 * Fix existing "Overdue" text in Days to Deadline column
 * Converts text back to negative numbers for proper counting
 */
function fixOverdueTextToNumbers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Grievance Log not found.');
    return;
  }

  ss.toast('Fixing overdue data...', '🔧 Fix', 3);

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var daysCol = GRIEVANCE_COLS.DAYS_TO_DEADLINE;
  var nextActionCol = GRIEVANCE_COLS.NEXT_ACTION_DUE;

  var daysData = sheet.getRange(2, daysCol, lastRow - 1, 1).getValues();
  var nextActionData = sheet.getRange(2, nextActionCol, lastRow - 1, 1).getValues();

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var updates = [];
  var fixCount = 0;

  for (var i = 0; i < daysData.length; i++) {
    var currentValue = daysData[i][0];
    var nextAction = nextActionData[i][0];

    if (currentValue === 'Overdue' && nextAction instanceof Date) {
      var days = Math.floor((nextAction - today) / (1000 * 60 * 60 * 24));
      updates.push([days]);
      fixCount++;
    } else {
      updates.push([currentValue]);
    }
  }

  if (fixCount > 0) {
    sheet.getRange(2, daysCol, updates.length, 1).setValues(updates);
    ss.toast('Fixed ' + fixCount + ' overdue entries!', '✅ Success', 3);
  } else {
    ss.toast('No "Overdue" text found to fix.', '✅ All Good', 3);
  }
}
