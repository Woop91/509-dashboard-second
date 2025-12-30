/**
 * 509 Dashboard - Menu Checklist Sheet
 *
 * Creates a checklist sheet with all menu items for tracking feature testing/usage.
 *
 * @version 1.0.0
 */

/**
 * Menu items data structure organized by menu category
 * Excludes Demo menu items
 */
var MENU_ITEMS = {
  'Dashboard': {
    icon: '👤',
    items: [
      { name: 'Smart Dashboard (Auto-Detect)', func: 'showSmartDashboard', icon: '📊' },
      { name: 'Interactive Dashboard', func: 'showInteractiveDashboardTab', icon: '🎯' },
      { name: 'View Active Grievances', func: 'viewActiveGrievances', icon: '📋' },
      { name: 'Mobile Dashboard', func: 'showMobileDashboard', icon: '📱' },
      { name: 'Get Mobile App URL', func: 'showWebAppUrl', icon: '📱' },
      { name: 'Quick Actions', func: 'showQuickActionsMenu', icon: '⚡' }
    ],
    submenus: {
      'Grievance Tools': {
        items: [
          { name: 'Start New Grievance', func: 'startNewGrievance', icon: '➕' },
          { name: 'Refresh Grievance Formulas', func: 'recalcAllGrievancesBatched', icon: '🔄' },
          { name: 'Refresh Member Directory Data', func: 'refreshMemberDirectoryFormulas', icon: '🔄' },
          { name: 'Setup Live Grievance Links', func: 'setupLiveGrievanceFormulas', icon: '🔗' },
          { name: 'Setup Member ID Dropdown', func: 'setupGrievanceMemberDropdown', icon: '👤' },
          { name: 'Fix Overdue Text Data', func: 'fixOverdueTextToNumbers', icon: '🔧' }
        ]
      }
    }
  },
  'Search': {
    icon: '🔍',
    items: [
      { name: 'Search Members', func: 'searchMembers', icon: '🔍' }
    ]
  },
  'Sheet Manager': {
    icon: '📊',
    items: [
      { name: 'Rebuild Dashboard', func: 'rebuildDashboard', icon: '📊' },
      { name: 'Refresh Interactive Charts', func: 'refreshInteractiveCharts', icon: '📈' },
      { name: 'Refresh All Formulas', func: 'refreshAllFormulas', icon: '🔄' }
    ],
    submenus: {
      'Google Drive': {
        items: [
          { name: 'Setup Folder for Grievance', func: 'setupDriveFolderForGrievance', icon: '📁' },
          { name: 'View Grievance Files', func: 'showGrievanceFiles', icon: '📁' },
          { name: 'Batch Create Folders', func: 'batchCreateGrievanceFolders', icon: '📁' }
        ]
      },
      'Calendar': {
        items: [
          { name: 'Sync Deadlines to Calendar', func: 'syncDeadlinesToCalendar', icon: '📅' },
          { name: 'View Upcoming Deadlines', func: 'showUpcomingDeadlinesFromCalendar', icon: '📅' },
          { name: 'Clear Calendar Events', func: 'clearAllCalendarEvents', icon: '🗑️' }
        ]
      },
      'Notifications': {
        items: [
          { name: 'Notification Settings', func: 'showNotificationSettings', icon: '⚙️' },
          { name: 'Test Notifications', func: 'testDeadlineNotifications', icon: '🧪' }
        ]
      }
    }
  },
  'Tools': {
    icon: '🔧',
    submenus: {
      'ADHD & Accessibility': {
        items: [
          { name: 'ADHD Control Panel', func: 'showADHDControlPanel', icon: '♿' },
          { name: 'Focus Mode', func: 'activateFocusMode', icon: '🎯' },
          { name: 'Toggle Zebra Stripes', func: 'toggleZebraStripes', icon: '🔲' },
          { name: 'Quick Capture', func: 'showQuickCaptureNotepad', icon: '📝' },
          { name: 'Pomodoro Timer', func: 'startPomodoroTimer', icon: '🍅' }
        ]
      },
      'Theming': {
        items: [
          { name: 'Theme Manager', func: 'showThemeManager', icon: '🎨' },
          { name: 'Toggle Dark Mode', func: 'quickToggleDarkMode', icon: '🌙' },
          { name: 'Reset Theme', func: 'resetToDefaultTheme', icon: '🔄' }
        ]
      },
      'Multi-Select': {
        items: [
          { name: 'Open Editor', func: 'showMultiSelectDialog', icon: '📝' },
          { name: 'Enable Auto-Open', func: 'installMultiSelectTrigger', icon: '⚡' },
          { name: 'Disable Auto-Open', func: 'removeMultiSelectTrigger', icon: '🚫' }
        ]
      },
      'Undo/Redo': {
        items: [
          { name: 'Undo Last Action', func: 'undoLastAction', icon: '↩️' },
          { name: 'Redo Action', func: 'redoLastAction', icon: '↪️' },
          { name: 'View History', func: 'showUndoRedoPanel', icon: '📋' },
          { name: 'Clear History', func: 'clearUndoHistory', icon: '🗑️' }
        ]
      },
      'Cache & Performance': {
        items: [
          { name: 'Cache Status', func: 'showCacheStatusDashboard', icon: '🗄️' },
          { name: 'Warm Up Caches', func: 'warmUpCaches', icon: '🔥' },
          { name: 'Clear All Caches', func: 'invalidateAllCaches', icon: '🗑️' }
        ]
      },
      'Validation': {
        items: [
          { name: 'Run Bulk Validation', func: 'runBulkValidation', icon: '🔍' },
          { name: 'Validation Settings', func: 'showValidationSettings', icon: '⚙️' },
          { name: 'Clear Indicators', func: 'clearValidationIndicators', icon: '🧹' },
          { name: 'Install Validation Trigger', func: 'installValidationTrigger', icon: '⚡' }
        ]
      }
    }
  },
  'Setup': {
    icon: '🏗️',
    items: [
      { name: 'REPAIR DASHBOARD', func: 'REPAIR_DASHBOARD', icon: '🔧' },
      { name: 'Setup Data Validations', func: 'setupDataValidations', icon: '⚙️' },
      { name: 'Setup ADHD Defaults', func: 'setupADHDDefaults', icon: '🎨' },
      { name: 'Create Menu Checklist', func: 'createMenuChecklist', icon: '📋' },
      { name: 'View Checklist Progress', func: 'showMenuChecklistProgress', icon: '📊' }
    ]
  },
  'Testing': {
    icon: '🧪',
    items: [
      { name: 'Run All Tests', func: 'runAllTests', icon: '🧪' },
      { name: 'Run Quick Tests', func: 'runQuickTests', icon: '⚡' },
      { name: 'View Test Results', func: 'viewTestResults', icon: '📊' }
    ]
  },
  'Administrator': {
    icon: '⚙️',
    items: [
      { name: 'DIAGNOSE SETUP', func: 'DIAGNOSE_SETUP', icon: '🔍' },
      { name: 'Verify Hidden Sheets', func: 'verifyHiddenSheets', icon: '🔍' }
    ],
    submenus: {
      'Setup & Triggers': {
        items: [
          { name: 'Setup All Hidden Sheets', func: 'setupAllHiddenSheets', icon: '🔧' },
          { name: 'Repair All Hidden Sheets', func: 'repairAllHiddenSheets', icon: '🔧' },
          { name: 'Install Auto-Sync Trigger', func: 'installAutoSyncTrigger', icon: '⚡' },
          { name: 'Remove Auto-Sync Trigger', func: 'removeAutoSyncTrigger', icon: '🚫' }
        ]
      },
      'Manual Sync': {
        items: [
          { name: 'Sync All Data Now', func: 'syncAllData', icon: '🔄' },
          { name: 'Sync Grievance → Members', func: 'syncGrievanceToMemberDirectory', icon: '🔄' },
          { name: 'Sync Members → Grievances', func: 'syncMemberToGrievanceLog', icon: '🔄' }
        ]
      }
    }
  }
};

/**
 * Creates or rebuilds the Menu Checklist sheet
 * Displays all menu items with checkboxes for tracking
 */
function createMenuChecklist() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = SHEETS.MENU_CHECKLIST || 'Menu Checklist';

  // Get or create sheet
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  // Build data rows
  var rows = [];
  var rowIndex = 0;

  // Header row
  rows.push(['✓', 'Menu', 'Submenu', 'Item', 'Function', 'Notes']);
  rowIndex++;

  // Process each menu
  for (var menuName in MENU_ITEMS) {
    var menu = MENU_ITEMS[menuName];
    var menuLabel = menu.icon + ' ' + menuName;

    // Add top-level menu items
    if (menu.items) {
      for (var i = 0; i < menu.items.length; i++) {
        var item = menu.items[i];
        rows.push([
          false,  // Checkbox
          menuLabel,
          '',  // No submenu
          item.icon + ' ' + item.name,
          item.func,
          ''  // Notes
        ]);
        rowIndex++;
      }
    }

    // Add submenu items
    if (menu.submenus) {
      for (var submenuName in menu.submenus) {
        var submenu = menu.submenus[submenuName];
        for (var j = 0; j < submenu.items.length; j++) {
          var subItem = submenu.items[j];
          rows.push([
            false,  // Checkbox
            menuLabel,
            submenuName,
            subItem.icon + ' ' + subItem.name,
            subItem.func,
            ''  // Notes
          ]);
          rowIndex++;
        }
      }
    }
  }

  // Write all data
  sheet.getRange(1, 1, rows.length, 6).setValues(rows);

  // Format header row
  var headerRange = sheet.getRange(1, 1, 1, 6);
  headerRange.setFontWeight('bold')
    .setBackground(COLORS.PRIMARY_PURPLE || '#7C3AED')
    .setFontColor(COLORS.WHITE || '#FFFFFF')
    .setHorizontalAlignment('center');

  // Add checkboxes to column A (except header)
  if (rows.length > 1) {
    var checkboxRange = sheet.getRange(2, 1, rows.length - 1, 1);
    checkboxRange.insertCheckboxes();
  }

  // Set column widths
  sheet.setColumnWidth(1, 40);   // Checkbox
  sheet.setColumnWidth(2, 150);  // Menu
  sheet.setColumnWidth(3, 150);  // Submenu
  sheet.setColumnWidth(4, 250);  // Item
  sheet.setColumnWidth(5, 250);  // Function
  sheet.setColumnWidth(6, 200);  // Notes

  // Freeze header row
  sheet.setFrozenRows(1);

  // Add alternating row colors for readability
  if (rows.length > 1) {
    for (var r = 2; r <= rows.length; r++) {
      if (r % 2 === 0) {
        sheet.getRange(r, 1, 1, 6).setBackground('#F9FAFB');
      }
    }
  }

  // Add conditional formatting for checked items (strikethrough)
  var dataRange = sheet.getRange(2, 1, rows.length - 1, 6);
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$A2=TRUE')
    .setBackground('#E8F5E9')
    .setRanges([dataRange])
    .build();

  var rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);

  // Activate the sheet
  sheet.activate();

  SpreadsheetApp.getUi().alert(
    'Menu Checklist Created',
    'Created checklist with ' + (rows.length - 1) + ' menu items.\n\n' +
    'Use the checkboxes to track which features you\'ve tested or used.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  return sheet;
}

/**
 * Gets count of checked/total items from the menu checklist
 * @returns {Object} Object with checked and total counts
 */
function getMenuChecklistProgress() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = SHEETS.MENU_CHECKLIST || 'Menu Checklist';
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return { checked: 0, total: 0, percentage: 0 };
  }

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { checked: 0, total: 0, percentage: 0 };
  }

  var checkboxes = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var total = checkboxes.length;
  var checked = 0;

  for (var i = 0; i < checkboxes.length; i++) {
    if (checkboxes[i][0] === true) {
      checked++;
    }
  }

  return {
    checked: checked,
    total: total,
    percentage: total > 0 ? Math.round((checked / total) * 100) : 0
  };
}

/**
 * Shows menu checklist progress in a dialog
 */
function showMenuChecklistProgress() {
  var progress = getMenuChecklistProgress();

  if (progress.total === 0) {
    SpreadsheetApp.getUi().alert(
      'No Checklist Found',
      'The Menu Checklist sheet has not been created yet.\n\n' +
      'Go to Setup > Create Menu Checklist to create it.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  SpreadsheetApp.getUi().alert(
    'Menu Checklist Progress',
    'Completed: ' + progress.checked + ' / ' + progress.total + ' items\n' +
    'Progress: ' + progress.percentage + '%',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
