/**
 * 509 Dashboard - Main Entry Point
 *
 * Core setup functions, menu system, and sheet creation.
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
    .addItem('🎯 Interactive Dashboard', 'showInteractiveDashboardTab')
    .addSeparator()
    .addItem('🔍 Search Members', 'searchMembers')
    .addItem('📋 View Active Grievances', 'viewActiveGrievances')
    .addItem('📱 Mobile Dashboard', 'showMobileDashboard')
    .addItem('📱 Get Mobile App URL', 'showWebAppUrl')
    .addItem('⚡ Quick Actions', 'showQuickActionsMenu')
    .addSeparator()
    .addSubMenu(ui.createMenu('📋 Grievance Tools')
      .addItem('➕ Start New Grievance', 'startNewGrievance')
      .addItem('🔄 Refresh Grievance Formulas', 'recalcAllGrievancesBatched')
      .addItem('🔄 Refresh Member Directory Data', 'refreshMemberDirectoryFormulas'))
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
    .addSubMenu(ui.createMenu('↩️ Undo/Redo')
      .addItem('↩️ Undo Last Action', 'undoLastAction')
      .addItem('↪️ Redo Action', 'redoLastAction')
      .addItem('📋 View History', 'showUndoRedoPanel')
      .addItem('🗑️ Clear History', 'clearUndoHistory'))
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
    .addItem('🏗️ CREATE 509 DASHBOARD', 'CREATE_509_DASHBOARD')
    .addItem('🔧 REPAIR DASHBOARD', 'REPAIR_DASHBOARD')
    .addSeparator()
    .addItem('⚙️ Setup Data Validations', 'setupDataValidations')
    .addItem('🎨 Setup ADHD Defaults', 'setupADHDDefaults')
    .addToUi();

  // Demo Menu - only show if demo mode is not disabled
  if (!isDemoModeDisabled()) {
    ui.createMenu('🎭 Demo')
      .addItem('🚀 Seed All Sample Data', 'SEED_SAMPLE_DATA')
      .addSeparator()
      .addSubMenu(ui.createMenu('🌱 Seed Data')
        .addItem('⚙️ Seed Config Dropdowns Only', 'seedConfigData')
        .addSeparator()
        .addItem('👥 Seed Members (Custom Count)', 'SEED_MEMBERS_DIALOG')
        .addItem('📋 Seed Grievances (Custom Count)', 'SEED_GRIEVANCES_DIALOG')
        .addSeparator()
        .addItem('👥 Seed 50 Members', 'seed50Members')
        .addItem('📋 Seed 25 Grievances', 'seed25Grievances'))
      .addSeparator()
      .addSubMenu(ui.createMenu('🗑️ Nuke Data')
        .addItem('☢️ NUKE SEEDED DATA', 'NUKE_SEEDED_DATA')
        .addItem('🧹 Clear Config Dropdowns Only', 'NUKE_CONFIG_DROPDOWNS'))
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
    'Plus 5 hidden calculation sheets for self-healing formulas.\n\n' +
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
    ss.toast('Created Interactive Dashboard', '🏗️ Progress', 2);

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
      'Plus 5 hidden calculation sheets with self-healing formulas.\n\n' +
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

  // Add checkbox for Start Grievance column
  var checkboxRange = sheet.getRange(2, MEMBER_COLS.START_GRIEVANCE, 998, 1);
  checkboxRange.insertCheckboxes();

  // Auto-resize other columns
  sheet.autoResizeColumns(1, headers.length);
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

  // Add checkbox for Message Alert column (AC)
  var checkboxRange = sheet.getRange(2, GRIEVANCE_COLS.MESSAGE_ALERT, 998, 1);
  checkboxRange.insertCheckboxes();

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
    sheet.getRange(2, col, 998, 1).setNumberFormat('dd-mm-yyyy');
  });

  // Format Days Open (S) and Days to Deadline (U) as whole numbers
  sheet.getRange(2, GRIEVANCE_COLS.DAYS_OPEN, 998, 1).setNumberFormat('0');
  sheet.getRange(2, GRIEVANCE_COLS.DAYS_TO_DEADLINE, 998, 1).setNumberFormat('0');

  // Auto-resize other columns
  sheet.autoResizeColumns(1, headers.length);
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
      '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gResolutionCol + ':' + gResolutionCol + ',"*Won*")',
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
      '=IFERROR(TEXT(COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gIssueCatCol + ':' + gIssueCatCol + ',"' + cat + '",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gResolutionCol + ':' + gResolutionCol + ',"*Won*")/COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gIssueCatCol + ':' + gIssueCatCol + ',"' + cat + '"),"0%"),"-")',
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
      '=IF(AND(' + locRef + '<>"",COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gLocationCol + ':' + gLocationCol + ',' + locRef + ')>0),TEXT(COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gLocationCol + ':' + gLocationCol + ',' + locRef + ',\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gResolutionCol + ':' + gResolutionCol + ',"*Won*")/COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gLocationCol + ':' + gLocationCol + ',' + locRef + '),"0%"),"-")',
      '="-"'  // Satisfaction requires separate tracking
    ]);
  }
  sheet.getRange('A30:F34').setFormulas(locationFormulas)
    .setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 7: STEWARD PERFORMANCE
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A36').setValue('👨‍⚖️ STEWARD PERFORMANCE')
    .setFontWeight('bold')
    .setBackground('#7C3AED')  // Purple
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A36:F36').merge();

  var stewardLabels = [['Total Stewards', 'Active (w/ Cases)', 'Avg Cases/Steward', 'Best Win Rate', 'Total Vol Hours', 'Contacts This Month']];
  sheet.getRange('A37:F37').setValues(stewardLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  var mStewardCol = getColumnLetter(MEMBER_COLS.IS_STEWARD);
  var gAssignedStewardCol = getColumnLetter(GRIEVANCE_COLS.STEWARD);  // GRIEVANCE_COLS uses STEWARD not ASSIGNED_STEWARD
  var mContactDateCol = getColumnLetter(MEMBER_COLS.RECENT_CONTACT_DATE);

  var stewardFormulas = [
    [
      '=COUNTIF(\'' + SHEETS.MEMBER_DIR + '\'!' + mStewardCol + ':' + mStewardCol + ',"Yes")',
      '=IFERROR(ROWS(UNIQUE(FILTER(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gAssignedStewardCol + ':' + gAssignedStewardCol + ',\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gAssignedStewardCol + ':' + gAssignedStewardCol + '<>"",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + '="Open"))),0)',
      '=IFERROR(ROUND((COUNTA(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gIdCol + ':' + gIdCol + ')-1)/COUNTIF(\'' + SHEETS.MEMBER_DIR + '\'!' + mStewardCol + ':' + mStewardCol + ',"Yes"),1),"-")',
      '="-"',  // Best win rate requires complex calculation
      '=SUM(\'' + SHEETS.MEMBER_DIR + '\'!' + mVolHoursCol + ':' + mVolHoursCol + ')',
      '=COUNTIFS(\'' + SHEETS.MEMBER_DIR + '\'!' + mContactDateCol + ':' + mContactDateCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),\'' + SHEETS.MEMBER_DIR + '\'!' + mContactDateCol + ':' + mContactDateCol + ',"<="&TODAY())'
    ]
  ];
  sheet.getRange('A38:F38').setFormulas(stewardFormulas)
    .setFontSize(16)
    .setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 8: MONTH-OVER-MONTH TRENDS
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A40').setValue('📈 MONTH-OVER-MONTH TRENDS')
    .setFontWeight('bold')
    .setBackground('#DC2626')  // Red
    .setFontColor(COLORS.WHITE);
  sheet.getRange('A40:F40').merge();

  var trendLabels = [['Metric', 'This Month', 'Last Month', 'Change', '% Change', 'Trend']];
  sheet.getRange('A41:F41').setValues(trendLabels)
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY)
    .setHorizontalAlignment('center');

  // Trend rows: Filed, Closed, Won
  var trendData = [
    [
      '="Grievances Filed"',  // Wrap text as formula for setFormulas()
      '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateFiledCol + ':' + gDateFiledCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateFiledCol + ':' + gDateFiledCol + ',"<="&TODAY())',
      '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateFiledCol + ':' + gDateFiledCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY())-1,1),\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateFiledCol + ':' + gDateFiledCol + ',"<"&DATE(YEAR(TODAY()),MONTH(TODAY()),1))',
      '=B42-C42',
      '=IFERROR(TEXT((B42-C42)/C42,"0%"),"-")',
      '=IF(B42>C42,"📈",IF(B42<C42,"📉","➡️"))'
    ],
    [
      '="Grievances Closed"',  // Wrap text as formula for setFormulas()
      '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',"<="&TODAY())',
      '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY())-1,1),\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',"<"&DATE(YEAR(TODAY()),MONTH(TODAY()),1))',
      '=B43-C43',
      '=IFERROR(TEXT((B43-C43)/C43,"0%"),"-")',
      '=IF(B43>C43,"📈",IF(B43<C43,"📉","➡️"))'
    ],
    [
      '="Cases Won"',  // Wrap text as formula for setFormulas()
      '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gResolutionCol + ':' + gResolutionCol + ',"*Won*")',
      '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY())-1,1),\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',"<"&DATE(YEAR(TODAY()),MONTH(TODAY()),1),\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gResolutionCol + ':' + gResolutionCol + ',"*Won*")',
      '=B44-C44',
      '=IFERROR(TEXT((B44-C44)/C44,"0%"),"-")',
      '=IF(B44>C44,"📈",IF(B44<C44,"📉","➡️"))'
    ]
  ];
  sheet.getRange('A42:F44').setFormulas(trendData)
    .setHorizontalAlignment('center');

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 9: STATUS LEGEND
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange('A46').setValue('STATUS LEGEND')
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange('A46:F46').merge();

  var legend = [
    ['🟢 On Track', '🟡 Due in 7 days', '🟠 Due in 3 days', '🔴 Overdue', '✅ Won', '❌ Denied']
  ];
  sheet.getRange('A47:F47').setValues(legend)
    .setHorizontalAlignment('center')
    .setFontSize(10);

  // Auto-resize and format
  sheet.autoResizeColumns(1, 6);
  sheet.setFrozenRows(3);

  // Set minimum column widths
  for (var i = 1; i <= 6; i++) {
    if (sheet.getColumnWidth(i) < 120) {
      sheet.setColumnWidth(i, 120);
    }
  }
}

/**
 * Create or recreate the Interactive Dashboard sheet
 * Allows users to select metrics and visualization preferences
 */
function createInteractiveDashboard(ss) {
  var sheet = getOrCreateSheet(ss, SHEETS.INTERACTIVE);
  sheet.clear();

  // Title
  sheet.getRange('A1').setValue('🎯 Interactive Dashboard')
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
    ['Won', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gResolutionCol + ':' + gResolutionCol + ',"*Won*")', 'Cases won (full or partial)'],
    ['Win Rate', '=IFERROR(ROUND(COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gResolutionCol + ':' + gResolutionCol + ',"*Won*")/(COUNTA(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gIdCol + ':' + gIdCol + ')-1)*100,1)&"%","0%")', 'Win percentage of all cases']
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

  // Create and show dialog
  var html = HtmlService.createHtmlOutputFromFile('MultiSelectDialog')
    .setWidth(350)
    .setHeight(420);

  // Pass data to dialog
  var dialogData = {
    label: config.label,
    options: options,
    currentValues: currentValues
  };

  html.append('<script>initDialog(' + JSON.stringify(dialogData) + ');</script>');

  SpreadsheetApp.getUi().showModalDialog(html, 'Multi-Select: ' + config.label);
}

/**
 * Get values from a Config sheet column
 * @param {Sheet} configSheet - The Config sheet
 * @param {number} col - Column number
 * @returns {Array} Array of non-empty values
 */
function getConfigValues(configSheet, col) {
  var data = configSheet.getRange(3, col, 100, 1).getValues();
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Remove existing triggers to avoid duplicates
  var triggers = ScriptApp.getUserTriggers(ss);
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onSelectionChangeMultiSelect') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new trigger
  ScriptApp.newTrigger('onSelectionChangeMultiSelect')
    .forSpreadsheet(ss)
    .onSelectionChange()
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ Multi-Select Auto-Open Enabled!\n\n' +
    'Now when you click on a multi-select cell (Office Days, Preferred Comm, etc.), ' +
    'the selection dialog will automatically appear.'
  );
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

  // Check hidden sheets (5 hidden calculation sheets)
  report.push('🔒 HIDDEN SHEETS:');
  var hiddenSheets = [
    SHEETS.GRIEVANCE_CALC,
    SHEETS.GRIEVANCE_FORMULAS,
    SHEETS.MEMBER_LOOKUP,
    SHEETS.STEWARD_CONTACT_CALC,
    SHEETS.DASHBOARD_CALC
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
    '• Recreate all 5 hidden calculation sheets with formulas\n' +
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

    // Final message handled by repairAllHiddenSheets
  } catch (error) {
    Logger.log('Error in REPAIR_DASHBOARD: ' + error.message);
    ui.alert('❌ Error', 'Repair failed: ' + error.message, ui.ButtonSet.OK);
  }
}

// ============================================================================
// MENU HANDLER FUNCTIONS
// ============================================================================

function searchMembers() {
  SpreadsheetApp.getUi().alert('Search Members feature - Coming soon!');
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

/**
 * Run all tests (stub - TestingValidation.gs not included)
 */
function runAllTests() {
  SpreadsheetApp.getUi().alert('🧪 Run All Tests',
    'Test framework not yet implemented.\n\n' +
    'To add tests, create TestingValidation.gs with test functions.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Run quick tests (stub - TestingValidation.gs not included)
 */
function runQuickTests() {
  SpreadsheetApp.getUi().alert('⚡ Run Quick Tests',
    'Test framework not yet implemented.\n\n' +
    'To add tests, create TestingValidation.gs with test functions.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function viewTestResults() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.TEST_RESULTS);
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('No test results yet. Run tests first using 🧪 Testing menu.');
  }
}
