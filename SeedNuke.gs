/**
 * 509 Dashboard - Seed and Nuke Functions
 *
 * Functions for seeding sample data and clearing data.
 * Seeded data is tracked separately from manually entered data.
 * NUKE only removes seeded data, preserving manual entries.
 *
 * ⚠️ WARNING: DO NOT DEPLOY THIS FILE DIRECTLY
 * This is a source file used to generate ConsolidatedDashboard.gs.
 * Deploy ONLY ConsolidatedDashboard.gs to avoid function conflicts.
 *
 * @version 1.0.0
 * @license Free for use by non-profit collective bargaining groups and unions
 */

// ============================================================================
// DEMO MODE TRACKING
// ============================================================================

/**
 * Check if demo mode has been disabled (after nuke)
 * @returns {boolean} True if demo mode is disabled
 */
function isDemoModeDisabled() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty('DEMO_MODE_DISABLED') === 'true';
}

/**
 * Disable demo mode permanently (called after nuke)
 */
function disableDemoMode() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('DEMO_MODE_DISABLED', 'true');
  // Clear tracked IDs since they're no longer needed
  props.deleteProperty('SEEDED_MEMBER_IDS');
  props.deleteProperty('SEEDED_GRIEVANCE_IDS');
}

/**
 * Track a seeded member ID
 * @param {string} memberId - The member ID to track
 */
function trackSeededMemberId(memberId) {
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty('SEEDED_MEMBER_IDS') || '';
  var ids = existing ? existing.split(',') : [];
  if (ids.indexOf(memberId) === -1) {
    ids.push(memberId);
    props.setProperty('SEEDED_MEMBER_IDS', ids.join(','));
  }
}

/**
 * Track a seeded grievance ID
 * @param {string} grievanceId - The grievance ID to track
 */
function trackSeededGrievanceId(grievanceId) {
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty('SEEDED_GRIEVANCE_IDS') || '';
  var ids = existing ? existing.split(',') : [];
  if (ids.indexOf(grievanceId) === -1) {
    ids.push(grievanceId);
    props.setProperty('SEEDED_GRIEVANCE_IDS', ids.join(','));
  }
}

/**
 * Get all tracked seeded member IDs
 * @returns {Object} Object with member IDs as keys for quick lookup
 */
function getSeededMemberIds() {
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty('SEEDED_MEMBER_IDS') || '';
  var ids = existing ? existing.split(',') : [];
  var lookup = {};
  ids.forEach(function(id) { if (id) lookup[id] = true; });
  return lookup;
}

/**
 * Get all tracked seeded grievance IDs
 * @returns {Object} Object with grievance IDs as keys for quick lookup
 */
function getSeededGrievanceIds() {
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty('SEEDED_GRIEVANCE_IDS') || '';
  var ids = existing ? existing.split(',') : [];
  var lookup = {};
  ids.forEach(function(id) { if (id) lookup[id] = true; });
  return lookup;
}

/**
 * Batch track multiple seeded member IDs (more efficient than individual calls)
 * @param {Array<string>} memberIds - Array of member IDs to track
 */
function trackSeededMemberIdsBatch(memberIds) {
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty('SEEDED_MEMBER_IDS') || '';
  var ids = existing ? existing.split(',') : [];
  memberIds.forEach(function(id) {
    if (id && ids.indexOf(id) === -1) ids.push(id);
  });
  props.setProperty('SEEDED_MEMBER_IDS', ids.join(','));
}

/**
 * Batch track multiple seeded grievance IDs (more efficient than individual calls)
 * @param {Array<string>} grievanceIds - Array of grievance IDs to track
 */
function trackSeededGrievanceIdsBatch(grievanceIds) {
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty('SEEDED_GRIEVANCE_IDS') || '';
  var ids = existing ? existing.split(',') : [];
  grievanceIds.forEach(function(id) {
    if (id && ids.indexOf(id) === -1) ids.push(id);
  });
  props.setProperty('SEEDED_GRIEVANCE_IDS', ids.join(','));
}

// ============================================================================
// SEED FUNCTIONS
// ============================================================================

/**
 * Seed all sample data: Config + 1,000 members + 300 grievances
 * Grievances are randomly distributed - some members may have multiple
 * Auto-installs the sync trigger for live updates between sheets
 */
function SEED_SAMPLE_DATA() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var response = ui.alert(
    '🚀 Seed Sample Data',
    'This will seed:\n' +
    '• Config dropdowns (Job Titles, Locations, etc.)\n' +
    '• 1,000 sample members\n' +
    '• 300 sample grievances (30%)\n' +
    '• Auto-sync trigger for live updates\n\n' +
    'Note: Some members may have multiple grievances.\n' +
    'Member Directory will auto-update when Grievance Log changes.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  ss.toast('Seeding config data...', '🌱 Seeding', 3);
  seedConfigData();

  ss.toast('Seeding 1,000 members + 300 grievances (this may take a moment)...', '🌱 Seeding', 10);
  SEED_MEMBERS(1000, 30);

  ss.toast('Installing auto-sync trigger...', '🔧 Setup', 3);
  installAutoSyncTriggerQuick();

  ss.toast('Sample data seeded successfully!', '✅ Success', 5);
  ui.alert('✅ Success', 'Sample data has been seeded!\n\n' +
    '• Config dropdowns populated\n' +
    '• 1,000 members added\n' +
    '• 300 grievances added (30%)\n' +
    '• Auto-sync trigger installed\n\n' +
    'Member Directory columns (Has Open Grievance?, Grievance Status, Days to Deadline) ' +
    'will now auto-update when you edit the Grievance Log.', ui.ButtonSet.OK);
}

/**
 * Seed Config sheet with dropdown values
 * Note: Data starts at row 3 (row 1 = section headers, row 2 = column headers)
 */
function seedConfigData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error: Config sheet not found. Please run CREATE_509_DASHBOARD first.');
    return;
  }

  // Data row start (after section headers row 1 and column headers row 2)
  var dataStartRow = 3;

  // Helper: only seed column if it's empty (preserves user data)
  function seedIfEmpty(column, values) {
    var existing = getConfigValues(sheet, column);
    if (existing.length === 0) {
      sheet.getRange(dataStartRow, column, values.length, 1)
        .setValues(values.map(function(v) { return [v]; }));
      return true;
    }
    return false;
  }

  var seededAny = false;

  // Job Titles (Column A)
  if (seedIfEmpty(CONFIG_COLS.JOB_TITLES, [
    'Social Worker', 'Case Manager', 'Program Coordinator', 'Administrative Assistant',
    'Supervisor', 'Director', 'Clinician', 'Counselor', 'Specialist', 'Analyst',
    'Manager', 'Senior Social Worker', 'Lead Case Manager', 'Program Manager',
    'Executive Assistant', 'HR Coordinator', 'Finance Associate', 'IT Support',
    'Communications Specialist', 'Outreach Worker'
  ])) seededAny = true;

  // Office Locations (Column B)
  if (seedIfEmpty(CONFIG_COLS.OFFICE_LOCATIONS, [
    'Boston Main Office', 'Worcester Regional', 'Springfield Center', 'Cambridge Branch',
    'Lowell Office', 'Brockton Center', 'Quincy Regional', 'New Bedford Office',
    'Fall River Branch', 'Lawrence Center', 'Framingham Office', 'Somerville Branch',
    'Lynn Regional', 'Haverhill Center', 'Malden Office', 'Medford Branch',
    'Waltham Regional', 'Newton Center', 'Brookline Office', 'Salem Branch'
  ])) seededAny = true;

  // Units (Column C)
  if (seedIfEmpty(CONFIG_COLS.UNITS, [
    'Child Welfare', 'Adult Services', 'Mental Health', 'Disability Services',
    'Elder Affairs', 'Housing Assistance', 'Employment Services', 'Youth Services',
    'Family Support', 'Administration'
  ])) seededAny = true;

  // Supervisors (Column F)
  if (seedIfEmpty(CONFIG_COLS.SUPERVISORS, [
    'Maria Rodriguez', 'James Wilson', 'Sarah Chen', 'Michael Brown',
    'Jennifer Davis', 'Robert Taylor', 'Lisa Anderson', 'David Martinez',
    'Emily Johnson', 'Christopher Lee', 'Amanda White', 'Daniel Garcia'
  ])) seededAny = true;

  // Managers (Column G)
  if (seedIfEmpty(CONFIG_COLS.MANAGERS, [
    'Patricia Thompson', 'William Jackson', 'Elizabeth Moore', 'Richard Harris',
    'Susan Clark', 'Joseph Lewis', 'Margaret Robinson', 'Charles Walker'
  ])) seededAny = true;

  // Stewards (Column H)
  if (seedIfEmpty(CONFIG_COLS.STEWARDS, [
    'John Smith', 'Mary Johnson', 'Robert Williams', 'Patricia Jones',
    'Michael Davis', 'Linda Miller', 'William Brown', 'Barbara Wilson',
    'David Moore', 'Susan Taylor', 'James Anderson', 'Karen Thomas'
  ])) seededAny = true;

  // Steward Committees (Column I)
  if (seedIfEmpty(CONFIG_COLS.STEWARD_COMMITTEES, [
    'Grievance Committee', 'Bargaining Committee', 'Health & Safety Committee',
    'Political Action Committee', 'Membership Committee', 'Executive Board'
  ])) seededAny = true;

  // Home Towns (Column AF)
  if (seedIfEmpty(CONFIG_COLS.HOME_TOWNS, [
    'Boston', 'Worcester', 'Springfield', 'Cambridge', 'Lowell', 'Brockton',
    'Quincy', 'New Bedford', 'Fall River', 'Lawrence', 'Framingham', 'Somerville',
    'Lynn', 'Haverhill', 'Malden', 'Medford', 'Waltham', 'Newton', 'Brookline'
  ])) seededAny = true;

  if (seededAny) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Config data seeded!', '✅ Success', 3);
  }
}

/**
 * Restore Config dropdowns AND re-apply dropdown validations to Member Directory and Grievance Log
 * This is the full restore function for use after nuking or when dropdowns are missing
 */
function restoreConfigAndDropdowns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  ss.toast('Restoring Config values...', '🔄 Restoring', 2);

  // First, seed the Config values
  seedConfigData();

  ss.toast('Applying dropdown validations...', '🔄 Restoring', 2);

  // Then, re-apply dropdown validations to Member Directory and Grievance Log
  setupDataValidations();

  ss.toast('Config and dropdowns restored!', '✅ Success', 3);
}

/**
 * Seed sample members with optional grievances
 * @param {number} count - Number of members to seed (max 2000)
 * @param {number} grievancePercent - Percentage of members to give grievances (0-100, default 30)
 */
function SEED_MEMBERS(count, grievancePercent) {
  count = Math.min(count || 50, 2000);
  grievancePercent = (grievancePercent !== undefined) ? grievancePercent : 30; // Default 30% get grievances

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!sheet || !configSheet) {
    SpreadsheetApp.getUi().alert('Error: Required sheets not found.');
    return;
  }

  // Always ensure Config has data for all required columns
  ss.toast('Ensuring Config data exists...', '🌱 Seeding', 2);
  seedConfigData();  // seedConfigData now only populates EMPTY columns

  // Now get all config values (will have data from seedConfigData or user's existing data)
  var jobTitles = getConfigValues(configSheet, CONFIG_COLS.JOB_TITLES);
  var locations = getConfigValues(configSheet, CONFIG_COLS.OFFICE_LOCATIONS);
  var units = getConfigValues(configSheet, CONFIG_COLS.UNITS);
  var supervisors = getConfigValues(configSheet, CONFIG_COLS.SUPERVISORS);
  var managers = getConfigValues(configSheet, CONFIG_COLS.MANAGERS);
  var stewards = getConfigValues(configSheet, CONFIG_COLS.STEWARDS);
  var homeTowns = getConfigValues(configSheet, CONFIG_COLS.HOME_TOWNS);
  var committees = getConfigValues(configSheet, CONFIG_COLS.STEWARD_COMMITTEES);

  // Expanded name pools for better variety (100+ names each = 10,000+ unique combinations)
  var firstNames = [
    'James', 'Mary', 'John', 'Patricia', 'Robert', 'Jennifer', 'Michael', 'Linda', 'William', 'Elizabeth',
    'David', 'Barbara', 'Richard', 'Susan', 'Joseph', 'Jessica', 'Thomas', 'Sarah', 'Charles', 'Karen',
    'Christopher', 'Nancy', 'Daniel', 'Lisa', 'Matthew', 'Betty', 'Anthony', 'Margaret', 'Mark', 'Sandra',
    'Donald', 'Ashley', 'Steven', 'Kimberly', 'Paul', 'Emily', 'Andrew', 'Donna', 'Joshua', 'Michelle',
    'Kenneth', 'Dorothy', 'Kevin', 'Carol', 'Brian', 'Amanda', 'George', 'Melissa', 'Timothy', 'Deborah',
    'Ronald', 'Stephanie', 'Edward', 'Rebecca', 'Jason', 'Sharon', 'Jeffrey', 'Laura', 'Ryan', 'Cynthia',
    'Jacob', 'Kathleen', 'Gary', 'Amy', 'Nicholas', 'Angela', 'Eric', 'Shirley', 'Jonathan', 'Anna',
    'Stephen', 'Brenda', 'Larry', 'Pamela', 'Justin', 'Emma', 'Scott', 'Nicole', 'Brandon', 'Helen',
    'Benjamin', 'Samantha', 'Samuel', 'Katherine', 'Raymond', 'Christine', 'Gregory', 'Debra', 'Frank', 'Rachel',
    'Alexander', 'Carolyn', 'Patrick', 'Janet', 'Jack', 'Catherine', 'Dennis', 'Maria', 'Jerry', 'Heather',
    'Tyler', 'Diane', 'Aaron', 'Ruth', 'Jose', 'Julie', 'Adam', 'Olivia', 'Nathan', 'Joyce',
    'Henry', 'Virginia', 'Douglas', 'Victoria', 'Zachary', 'Kelly', 'Peter', 'Lauren', 'Kyle', 'Christina'
  ];
  var lastNames = [
    'Smith', 'Johnson', 'Williams', 'Brown', 'Jones', 'Garcia', 'Miller', 'Davis', 'Rodriguez', 'Martinez',
    'Hernandez', 'Lopez', 'Gonzalez', 'Wilson', 'Anderson', 'Thomas', 'Taylor', 'Moore', 'Jackson', 'Martin',
    'Lee', 'Perez', 'Thompson', 'White', 'Harris', 'Sanchez', 'Clark', 'Ramirez', 'Lewis', 'Robinson',
    'Walker', 'Young', 'Allen', 'King', 'Wright', 'Scott', 'Torres', 'Nguyen', 'Hill', 'Flores',
    'Green', 'Adams', 'Nelson', 'Baker', 'Hall', 'Rivera', 'Campbell', 'Mitchell', 'Carter', 'Roberts',
    'Gomez', 'Phillips', 'Evans', 'Turner', 'Diaz', 'Parker', 'Cruz', 'Edwards', 'Collins', 'Reyes',
    'Stewart', 'Morris', 'Morales', 'Murphy', 'Cook', 'Rogers', 'Gutierrez', 'Ortiz', 'Morgan', 'Cooper',
    'Peterson', 'Bailey', 'Reed', 'Kelly', 'Howard', 'Ramos', 'Kim', 'Cox', 'Ward', 'Richardson',
    'Watson', 'Brooks', 'Chavez', 'Wood', 'James', 'Bennett', 'Gray', 'Mendoza', 'Ruiz', 'Hughes',
    'Price', 'Alvarez', 'Castillo', 'Sanders', 'Patel', 'Myers', 'Long', 'Ross', 'Foster', 'Jimenez',
    'Powell', 'Jenkins', 'Perry', 'Russell', 'Sullivan', 'Bell', 'Coleman', 'Butler', 'Henderson', 'Barnes',
    'Gonzales', 'Fisher', 'Vasquez', 'Simmons', 'Stokes', 'Burns', 'Fox', 'Alexander', 'Rice', 'Stone'
  ];
  var officeDays = DEFAULT_CONFIG.OFFICE_DAYS;
  var commMethods = DEFAULT_CONFIG.COMM_METHODS;

  var startRow = Math.max(sheet.getLastRow() + 1, 2);

  // Build set of existing member IDs to prevent duplicates
  var existingMemberIds = {};
  if (startRow > 2) {
    var existingData = sheet.getRange(2, MEMBER_COLS.MEMBER_ID, startRow - 2, 1).getValues();
    for (var e = 0; e < existingData.length; e++) {
      if (existingData[e][0]) {
        existingMemberIds[existingData[e][0]] = true;
      }
    }
  }

  var rows = [];
  var seededIds = []; // Track IDs for this seeding session
  var batchSize = 50;
  var today = new Date();

  for (var i = 0; i < count; i++) {
    var firstName = randomChoice(firstNames);
    var lastName = randomChoice(lastNames);
    var memberId = generateNameBasedId('M', firstName, lastName, existingMemberIds);
    existingMemberIds[memberId] = true; // Track new ID to prevent duplicates in same batch
    seededIds.push(memberId); // Track for persistence
    var email = firstName.toLowerCase() + '.' + lastName.toLowerCase() + '.' + memberId.toLowerCase() + '@example.org';
    var phone = '617-555-' + String(Math.floor(Math.random() * 9000) + 1000);
    var isSteward = Math.random() < 0.1 ? 'Yes' : 'No';
    var assignedSteward = randomChoice(stewards);

    // Generate recent contact data (50% chance of having recent contact)
    var hasRecentContact = Math.random() < 0.5;
    var recentContactDate = hasRecentContact ? randomDate(new Date(today.getTime() - 60 * 24 * 60 * 60 * 1000), today) : '';
    var contactSteward = hasRecentContact ? assignedSteward : '';

    // Sample contact notes for members who have been contacted
    var sampleContactNotes = [
      'Discussed workload concerns',
      'Follow up on scheduling issue',
      'Interested in becoming steward',
      'Addressed safety complaint',
      'Positive feedback received',
      'Needs info on benefits',
      'Question about contract language',
      'Planning to attend next meeting',
      'Grievance update provided',
      'Initial outreach - new member',
      'Discussed upcoming negotiations',
      'Shared resources on workplace rights',
      'Scheduling meeting for next week'
    ];
    var contactNotes = hasRecentContact ? randomChoice(sampleContactNotes) : '';

    var row = generateSingleMemberRow(
      memberId, firstName, lastName,
      randomChoice(jobTitles),
      randomChoice(locations),
      randomChoice(units),
      randomChoice(officeDays),
      email, phone,
      randomChoice(commMethods),
      'Morning',
      randomChoice(supervisors),
      randomChoice(managers),
      isSteward,
      isSteward === 'Yes' ? randomChoice(committees) : '',
      assignedSteward,
      randomChoice(homeTowns),
      recentContactDate,
      contactSteward,
      contactNotes
    );

    rows.push(row);

    // Write in batches
    if (rows.length >= batchSize || i === count - 1) {
      sheet.getRange(startRow, 1, rows.length, 31).setValues(rows);
      startRow += rows.length;
      rows = [];
      Utilities.sleep(100);
    }
  }

  // Re-apply checkboxes to Start Grievance column (AE) - setValues overwrites them
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, MEMBER_COLS.START_GRIEVANCE, lastRow - 1, 1).insertCheckboxes();
  }

  // Sync grievance data to Member Directory (populates AB-AD: Has Open Grievance, Status, Next Deadline)
  syncGrievanceToMemberDirectory();

  // Track seeded IDs for later cleanup (nuke only removes seeded data)
  trackSeededMemberIdsBatch(seededIds);

  // Seed grievances if percentage > 0
  var grievanceCount = Math.floor(count * grievancePercent / 100);
  if (grievanceCount > 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(count + ' members seeded! Now seeding ' + grievanceCount + ' grievances...', '✅ Members Done', 2);
    SEED_GRIEVANCES(grievanceCount);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast(count + ' members seeded!', '✅ Success', 3);
  }
}

/**
 * Seed members only (no automatic grievance seeding)
 * Used by SEED_SAMPLE_DATA to separate member and grievance seeding
 * @param {number} count - Number of members to seed (max 2000)
 */
function SEED_MEMBERS_ONLY(count) {
  count = Math.min(count || 50, 2000);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!sheet || !configSheet) {
    SpreadsheetApp.getUi().alert('Error: Required sheets not found.');
    return;
  }

  // Always ensure Config has data for all required columns
  seedConfigData();

  // Get all config values
  var jobTitles = getConfigValues(configSheet, CONFIG_COLS.JOB_TITLES);
  var locations = getConfigValues(configSheet, CONFIG_COLS.OFFICE_LOCATIONS);
  var units = getConfigValues(configSheet, CONFIG_COLS.UNITS);
  var supervisors = getConfigValues(configSheet, CONFIG_COLS.SUPERVISORS);
  var managers = getConfigValues(configSheet, CONFIG_COLS.MANAGERS);
  var stewards = getConfigValues(configSheet, CONFIG_COLS.STEWARDS);
  var homeTowns = getConfigValues(configSheet, CONFIG_COLS.HOME_TOWNS);
  var committees = getConfigValues(configSheet, CONFIG_COLS.STEWARD_COMMITTEES);

  // Expanded name pools for better variety
  var firstNames = [
    'James', 'Mary', 'John', 'Patricia', 'Robert', 'Jennifer', 'Michael', 'Linda', 'William', 'Elizabeth',
    'David', 'Barbara', 'Richard', 'Susan', 'Joseph', 'Jessica', 'Thomas', 'Sarah', 'Charles', 'Karen',
    'Christopher', 'Nancy', 'Daniel', 'Lisa', 'Matthew', 'Betty', 'Anthony', 'Margaret', 'Mark', 'Sandra',
    'Donald', 'Ashley', 'Steven', 'Kimberly', 'Paul', 'Emily', 'Andrew', 'Donna', 'Joshua', 'Michelle',
    'Kenneth', 'Dorothy', 'Kevin', 'Carol', 'Brian', 'Amanda', 'George', 'Melissa', 'Timothy', 'Deborah',
    'Ronald', 'Stephanie', 'Edward', 'Rebecca', 'Jason', 'Sharon', 'Jeffrey', 'Laura', 'Ryan', 'Cynthia',
    'Jacob', 'Kathleen', 'Gary', 'Amy', 'Nicholas', 'Angela', 'Eric', 'Shirley', 'Jonathan', 'Anna',
    'Stephen', 'Brenda', 'Larry', 'Pamela', 'Justin', 'Emma', 'Scott', 'Nicole', 'Brandon', 'Helen',
    'Benjamin', 'Samantha', 'Samuel', 'Katherine', 'Raymond', 'Christine', 'Gregory', 'Debra', 'Frank', 'Rachel',
    'Alexander', 'Carolyn', 'Patrick', 'Janet', 'Jack', 'Catherine', 'Dennis', 'Maria', 'Jerry', 'Heather',
    'Tyler', 'Diane', 'Aaron', 'Ruth', 'Jose', 'Julie', 'Adam', 'Olivia', 'Nathan', 'Joyce',
    'Henry', 'Virginia', 'Douglas', 'Victoria', 'Zachary', 'Kelly', 'Peter', 'Lauren', 'Kyle', 'Christina'
  ];
  var lastNames = [
    'Smith', 'Johnson', 'Williams', 'Brown', 'Jones', 'Garcia', 'Miller', 'Davis', 'Rodriguez', 'Martinez',
    'Hernandez', 'Lopez', 'Gonzalez', 'Wilson', 'Anderson', 'Thomas', 'Taylor', 'Moore', 'Jackson', 'Martin',
    'Lee', 'Perez', 'Thompson', 'White', 'Harris', 'Sanchez', 'Clark', 'Ramirez', 'Lewis', 'Robinson',
    'Walker', 'Young', 'Allen', 'King', 'Wright', 'Scott', 'Torres', 'Nguyen', 'Hill', 'Flores',
    'Green', 'Adams', 'Nelson', 'Baker', 'Hall', 'Rivera', 'Campbell', 'Mitchell', 'Carter', 'Roberts',
    'Gomez', 'Phillips', 'Evans', 'Turner', 'Diaz', 'Parker', 'Cruz', 'Edwards', 'Collins', 'Reyes',
    'Stewart', 'Morris', 'Morales', 'Murphy', 'Cook', 'Rogers', 'Gutierrez', 'Ortiz', 'Morgan', 'Cooper',
    'Peterson', 'Bailey', 'Reed', 'Kelly', 'Howard', 'Ramos', 'Kim', 'Cox', 'Ward', 'Richardson',
    'Watson', 'Brooks', 'Chavez', 'Wood', 'James', 'Bennett', 'Gray', 'Mendoza', 'Ruiz', 'Hughes',
    'Price', 'Alvarez', 'Castillo', 'Sanders', 'Patel', 'Myers', 'Long', 'Ross', 'Foster', 'Jimenez',
    'Powell', 'Jenkins', 'Perry', 'Russell', 'Sullivan', 'Bell', 'Coleman', 'Butler', 'Henderson', 'Barnes',
    'Gonzales', 'Fisher', 'Vasquez', 'Simmons', 'Stokes', 'Burns', 'Fox', 'Alexander', 'Rice', 'Stone'
  ];
  var officeDays = DEFAULT_CONFIG.OFFICE_DAYS;
  var commMethods = DEFAULT_CONFIG.COMM_METHODS;

  var startRow = Math.max(sheet.getLastRow() + 1, 2);

  // Build set of existing member IDs to prevent duplicates
  var existingMemberIds = {};
  if (startRow > 2) {
    var existingData = sheet.getRange(2, MEMBER_COLS.MEMBER_ID, startRow - 2, 1).getValues();
    for (var e = 0; e < existingData.length; e++) {
      if (existingData[e][0]) {
        existingMemberIds[existingData[e][0]] = true;
      }
    }
  }

  var rows = [];
  var seededIds = [];
  var batchSize = 100; // Larger batches for 1000 members
  var today = new Date();

  for (var i = 0; i < count; i++) {
    var firstName = randomChoice(firstNames);
    var lastName = randomChoice(lastNames);
    var memberId = generateNameBasedId('M', firstName, lastName, existingMemberIds);
    existingMemberIds[memberId] = true;
    seededIds.push(memberId);
    var email = firstName.toLowerCase() + '.' + lastName.toLowerCase() + '.' + memberId.toLowerCase() + '@example.org';
    var phone = '617-555-' + String(Math.floor(Math.random() * 9000) + 1000);
    var isSteward = Math.random() < 0.1 ? 'Yes' : 'No';
    var assignedSteward = randomChoice(stewards);

    var hasRecentContact = Math.random() < 0.5;
    var recentContactDate = hasRecentContact ? randomDate(new Date(today.getTime() - 60 * 24 * 60 * 60 * 1000), today) : '';
    var contactSteward = hasRecentContact ? assignedSteward : '';

    var sampleContactNotes = [
      'Discussed workload concerns', 'Follow up on scheduling issue', 'Interested in becoming steward',
      'Addressed safety complaint', 'Positive feedback received', 'Needs info on benefits',
      'Question about contract language', 'Planning to attend next meeting', 'Grievance update provided',
      'Initial outreach - new member', 'Discussed upcoming negotiations', 'Shared resources on workplace rights'
    ];
    var contactNotes = hasRecentContact ? randomChoice(sampleContactNotes) : '';

    var row = generateSingleMemberRow(
      memberId, firstName, lastName,
      randomChoice(jobTitles), randomChoice(locations), randomChoice(units), randomChoice(officeDays),
      email, phone, randomChoice(commMethods), 'Morning',
      randomChoice(supervisors), randomChoice(managers),
      isSteward, isSteward === 'Yes' ? randomChoice(committees) : '', assignedSteward,
      randomChoice(homeTowns), recentContactDate, contactSteward, contactNotes
    );

    rows.push(row);

    // Write in batches
    if (rows.length >= batchSize || i === count - 1) {
      sheet.getRange(startRow, 1, rows.length, 31).setValues(rows);
      startRow += rows.length;
      rows = [];
      Utilities.sleep(50);
    }
  }

  // Re-apply checkboxes
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, MEMBER_COLS.START_GRIEVANCE, lastRow - 1, 1).insertCheckboxes();
  }

  // Track seeded IDs
  trackSeededMemberIdsBatch(seededIds);

  SpreadsheetApp.getActiveSpreadsheet().toast(count + ' members seeded!', '✅ Success', 3);
}

/**
 * Generate a single member row with all 31 columns
 * @param {string} memberId - Member ID
 * @param {string} firstName - First name
 * @param {string} lastName - Last name
 * @param {string} jobTitle - Job title
 * @param {string} location - Work location
 * @param {string} unit - Unit
 * @param {string} officeDays - Office days
 * @param {string} email - Email
 * @param {string} phone - Phone
 * @param {string} prefComm - Preferred communication
 * @param {string} bestTime - Best time to contact
 * @param {string} supervisor - Supervisor
 * @param {string} manager - Manager
 * @param {string} isSteward - Is steward (Yes/No)
 * @param {string} committees - Committees
 * @param {string} assignedSteward - Assigned steward
 * @param {string} homeTown - Home town
 * @param {Date|string} recentContactDate - Recent contact date
 * @param {string} contactSteward - Steward who made contact
 */
function generateSingleMemberRow(memberId, firstName, lastName, jobTitle, location, unit, officeDays, email, phone, prefComm, bestTime, supervisor, manager, isSteward, committees, assignedSteward, homeTown, recentContactDate, contactSteward, contactNotes) {
  var today = new Date();
  var lastMonth = new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000);

  return [
    memberId,                                    // 1: Member ID (A)
    firstName,                                   // 2: First Name (B)
    lastName,                                    // 3: Last Name (C)
    jobTitle,                                    // 4: Job Title (D)
    location,                                    // 5: Work Location (E)
    unit,                                        // 6: Unit (F)
    officeDays,                                  // 7: Office Days (G)
    email,                                       // 8: Email (H)
    phone,                                       // 9: Phone (I)
    prefComm,                                    // 10: Preferred Communication (J)
    bestTime,                                    // 11: Best Time (K)
    supervisor,                                  // 12: Supervisor (L)
    manager,                                     // 13: Manager (M)
    isSteward,                                   // 14: Is Steward (N)
    committees,                                  // 15: Committees (O)
    assignedSteward,                             // 16: Assigned Steward (P)
    randomDate(lastMonth, today),                // 17: Last Virtual Mtg (Q)
    randomDate(lastMonth, today),                // 18: Last In-Person Mtg (R)
    Math.floor(Math.random() * 100),             // 19: Open Rate % (S)
    Math.floor(Math.random() * 20),              // 20: Volunteer Hours (T)
    Math.random() < 0.3 ? 'Yes' : 'No',          // 21: Interest Local (U)
    Math.random() < 0.2 ? 'Yes' : 'No',          // 22: Interest Chapter (V)
    Math.random() < 0.1 ? 'Yes' : 'No',          // 23: Interest Allied (W)
    homeTown || '',                              // 24: Home Town (X)
    recentContactDate || '',                     // 25: Recent Contact Date (Y)
    contactSteward || '',                        // 26: Contact Steward (Z)
    contactNotes || '',                          // 27: Contact Notes (AA) - seeded if contacted
    '',                                          // 28: Has Open Grievance (AB) - auto-calculated from Grievance Log
    '',                                          // 29: Grievance Status (AC) - auto-calculated from Grievance Log
    '',                                          // 30: Next Deadline (AD) - auto-calculated from Grievance Log
    false                                        // 31: Start Grievance (AE) - checkbox (re-applied after setValues)
  ];
}

/**
 * Seed N grievances
 * @param {number} count - Number of grievances to seed (max 300)
 */
function SEED_GRIEVANCES(count) {
  count = Math.min(count || 25, 300);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!grievanceSheet || !memberSheet || !configSheet) {
    SpreadsheetApp.getUi().alert('Error: Required sheets not found.');
    return;
  }

  // Get members
  var memberData = memberSheet.getDataRange().getValues();
  if (memberData.length < 2) {
    SpreadsheetApp.getUi().alert('Error: No members found. Please seed members first.');
    return;
  }

  var statuses = DEFAULT_CONFIG.GRIEVANCE_STATUS;
  var steps = DEFAULT_CONFIG.GRIEVANCE_STEP;
  var categories = DEFAULT_CONFIG.ISSUE_CATEGORY;
  var articles = DEFAULT_CONFIG.ARTICLES;

  // Get config values for stewards
  var stewards = getConfigValues(configSheet, CONFIG_COLS.STEWARDS);
  if (stewards.length === 0) stewards = ['Mary Steward'];

  var startRow = Math.max(grievanceSheet.getLastRow() + 1, 2);

  // Build set of existing grievance IDs to prevent duplicates
  var existingGrievanceIds = {};
  if (startRow > 2) {
    var existingData = grievanceSheet.getRange(2, GRIEVANCE_COLS.GRIEVANCE_ID, startRow - 2, 1).getValues();
    for (var e = 0; e < existingData.length; e++) {
      if (existingData[e][0]) {
        existingGrievanceIds[existingData[e][0]] = true;
      }
    }
  }

  var rows = [];
  var seededIds = []; // Track IDs for this seeding session
  var batchSize = 25;
  var today = new Date();

  // Create shuffled list of member indices (excluding header row)
  // This ensures each member is used before repeating
  var memberIndices = [];
  for (var m = 1; m < memberData.length; m++) {
    if (memberData[m][MEMBER_COLS.MEMBER_ID - 1]) { // Only include rows with valid member ID
      memberIndices.push(m);
    }
  }

  // Check if we found any valid members
  if (memberIndices.length === 0) {
    SpreadsheetApp.getUi().alert('Error: No members with valid Member IDs found. Please seed members first.');
    return;
  }

  var shuffledMembers = shuffleArray(memberIndices);
  var memberIndex = 0;

  // Distribute incident dates across the 90-day range to avoid clustering
  var dateRangeStart = new Date(today.getTime() - 90 * 24 * 60 * 60 * 1000);
  var dateSpread = 90 / count; // Days between each grievance's base date

  for (var i = 0; i < count; i++) {
    // Use shuffled members - cycle through if more grievances than members
    if (memberIndex >= shuffledMembers.length) {
      shuffledMembers = shuffleArray(memberIndices); // Reshuffle for next cycle
      memberIndex = 0;
    }
    var memberRow = memberData[shuffledMembers[memberIndex]];
    memberIndex++;

    var memberId = memberRow[MEMBER_COLS.MEMBER_ID - 1];

    // Skip if no member ID (shouldn't happen due to filtering above, but just in case)
    if (!memberId) continue;

    // Get member data directly from member row
    var firstName = memberRow[MEMBER_COLS.FIRST_NAME - 1] || '';
    var lastName = memberRow[MEMBER_COLS.LAST_NAME - 1] || '';
    var memberEmail = memberRow[MEMBER_COLS.EMAIL - 1] || '';
    var memberUnit = memberRow[MEMBER_COLS.UNIT - 1] || '';
    var memberLocation = memberRow[MEMBER_COLS.WORK_LOCATION - 1] || '';
    var memberSteward = memberRow[MEMBER_COLS.ASSIGNED_STEWARD - 1] || randomChoice(stewards);

    // Generate grievance ID using member's name with G prefix
    var grievanceId = generateNameBasedId('G', firstName, lastName, existingGrievanceIds);
    existingGrievanceIds[grievanceId] = true; // Track to prevent duplicates in same batch
    seededIds.push(grievanceId); // Track for persistence

    // Distribute incident dates across the 90-day range with some randomness
    // Each grievance gets a "slot" in the timeline, with +/- 2 days variation
    var baseDate = new Date(dateRangeStart.getTime() + (i * dateSpread * 24 * 60 * 60 * 1000));
    var variation = (Math.random() - 0.5) * 4 * 24 * 60 * 60 * 1000; // +/- 2 days
    var incidentDate = new Date(Math.min(baseDate.getTime() + variation, today.getTime()));

    var status = randomChoice(statuses);
    var step = randomChoice(steps);

    // Generate grievance row with all data populated directly
    var row = generateSingleGrievanceRow(
      grievanceId,
      memberId,
      firstName,
      lastName,
      status,
      step,
      incidentDate,
      randomChoice(articles),
      randomChoice(categories),
      memberEmail,
      memberUnit,
      memberLocation,
      memberSteward
    );

    rows.push(row);

    // Write in batches
    if (rows.length >= batchSize || i === count - 1) {
      grievanceSheet.getRange(startRow, 1, rows.length, 34).setValues(rows);
      startRow += rows.length;
      rows = [];
      Utilities.sleep(100);
    }
  }

  // Re-apply checkboxes to Message Alert column (AC) - setValues overwrites them
  var lastRow = grievanceSheet.getLastRow();
  if (lastRow >= 2) {
    grievanceSheet.getRange(2, GRIEVANCE_COLS.MESSAGE_ALERT, lastRow - 1, 1).insertCheckboxes();
  }

  // Sync data from hidden formulas sheet (self-healing - keeps data updated on edits)
  syncGrievanceFormulasToLog();

  // Sync grievance data to Member Directory (populates Has Open Grievance?, Status, Days to Deadline)
  syncGrievanceToMemberDirectory();

  // Track seeded IDs for later cleanup (nuke only removes seeded data)
  trackSeededGrievanceIdsBatch(seededIds);

  SpreadsheetApp.getActiveSpreadsheet().toast(count + ' grievances seeded!', '✅ Success', 3);
}

/**
 * Generate a grievance row for seeding - leaves formula columns empty
 * Formula columns (C, D, H, J, L, N, P, S, T, U, X, Y, Z, AA) will be auto-populated
 */
function generateSingleGrievanceRowForSeed(grievanceId, memberId, status, step, incidentDate, articles, category) {
  var today = new Date();

  // Calculate dates for manual entry columns only
  // Date Filed - always set if status is not brand new
  var dateFiled = addDays(incidentDate, Math.floor(Math.random() * 14) + 1);

  // Initialize timeline variables for manual entry columns
  var step1Rcvd = '';
  var step2AppealFiled = '';
  var step2Rcvd = '';
  var step3AppealFiled = '';
  var dateClosed = '';

  // Determine if case is closed
  var isClosed = (status === 'Settled' || status === 'Withdrawn' || status === 'Denied' || status === 'Won' || status === 'Closed');

  // Populate timeline based on current step
  var stepIndex = ['Informal', 'Step I', 'Step II', 'Step III', 'Mediation', 'Arbitration'].indexOf(step);

  // Step I Due calculated by formula, but Step I Rcvd is manual
  var step1Due = dateFiled ? addDays(dateFiled, 30) : '';

  // If Step I or beyond, Step I has been completed
  if (stepIndex >= 1 || isClosed) {
    step1Rcvd = addDays(step1Due, Math.floor(Math.random() * 10) - 5);
    if (step1Rcvd < dateFiled) step1Rcvd = addDays(dateFiled, 15);
  }

  // If Step II or beyond
  if (stepIndex >= 2 || (isClosed && stepIndex >= 1)) {
    step2AppealFiled = addDays(step1Rcvd, Math.floor(Math.random() * 8) + 1);
    var step2Due = addDays(step2AppealFiled, 30);
    step2Rcvd = addDays(step2Due, Math.floor(Math.random() * 10) - 5);
    if (step2Rcvd < step2AppealFiled) step2Rcvd = addDays(step2AppealFiled, 15);
  }

  // If Step III or beyond
  if (stepIndex >= 3 || (isClosed && stepIndex >= 2)) {
    step3AppealFiled = addDays(step2Rcvd, Math.floor(Math.random() * 20) + 1);
  }

  // Set date closed for resolved cases
  if (isClosed) {
    var lastDate = step3AppealFiled || step2Rcvd || step1Rcvd || dateFiled;
    dateClosed = addDays(lastDate, Math.floor(Math.random() * 30) + 5);
  }

  var resolutions = ['Won - Full remedy', 'Won - Partial remedy', 'Settled - Compromise', 'Denied', 'Withdrawn', 'Pending'];
  var resolution = dateClosed ? randomChoice(resolutions) : '';

  // Return row with empty strings for formula columns
  // Formula columns: C, D (names), H, J, L, N, P (deadlines), S, T, U (metrics), X, Y, Z, AA (member info)
  return [
    grievanceId,              // 1: Grievance ID (A)
    memberId,                 // 2: Member ID (B)
    '',                       // 3: First Name (C) - FORMULA
    '',                       // 4: Last Name (D) - FORMULA
    status,                   // 5: Status (E)
    step,                     // 6: Current Step (F)
    incidentDate,             // 7: Incident Date (G)
    '',                       // 8: Filing Deadline (H) - FORMULA
    dateFiled,                // 9: Date Filed (I)
    '',                       // 10: Step I Due (J) - FORMULA
    step1Rcvd,                // 11: Step I Rcvd (K)
    '',                       // 12: Step II Appeal Due (L) - FORMULA
    step2AppealFiled,         // 13: Step II Appeal Filed (M)
    '',                       // 14: Step II Due (N) - FORMULA
    step2Rcvd,                // 15: Step II Rcvd (O)
    '',                       // 16: Step III Appeal Due (P) - FORMULA
    step3AppealFiled,         // 17: Step III Appeal Filed (Q)
    dateClosed,               // 18: Date Closed (R)
    '',                       // 19: Days Open (S) - FORMULA
    '',                       // 20: Next Action Due (T) - FORMULA
    '',                       // 21: Days to Deadline (U) - FORMULA
    articles,                 // 22: Articles Violated (V)
    category,                 // 23: Issue Category (W)
    '',                       // 24: Member Email (X) - FORMULA
    '',                       // 25: Unit (Y) - FORMULA
    '',                       // 26: Location (Z) - FORMULA
    '',                       // 27: Steward (AA) - FORMULA
    resolution,               // 28: Resolution (AB)
    false,                    // 29: Message Alert (AC) - checkbox
    '',                       // 30: Coordinator Message (AD)
    '',                       // 31: Acknowledged By (AE)
    '',                       // 32: Acknowledged Date (AF)
    '',                       // 33: Drive Folder ID (AG)
    ''                        // 34: Drive Folder URL (AH)
  ];
}

/**
 * Generate a single grievance row with all 34 columns
 * Properly calculates all timeline fields based on grievance step progression
 */
function generateSingleGrievanceRow(grievanceId, memberId, firstName, lastName, status, step, incidentDate, articles, category, email, unit, location, steward) {
  var today = new Date();
  var filingDeadline = addDays(incidentDate, 21);

  // Date Filed - always set if status is not brand new
  var dateFiled = addDays(incidentDate, Math.floor(Math.random() * 14) + 1);

  // Step I Due (30 days after filing)
  var step1Due = dateFiled ? addDays(dateFiled, 30) : '';

  // Initialize timeline variables
  var step1Rcvd = '';
  var step2AppealDue = '';
  var step2AppealFiled = '';
  var step2Due = '';
  var step2Rcvd = '';
  var step3AppealDue = '';
  var step3AppealFiled = '';
  var dateClosed = '';

  // Determine if case is closed
  var isClosed = (status === 'Settled' || status === 'Withdrawn' || status === 'Denied' || status === 'Won' || status === 'Closed');

  // Populate timeline based on current step
  var stepIndex = ['Informal', 'Step I', 'Step II', 'Step III', 'Mediation', 'Arbitration'].indexOf(step);

  // If Step I or beyond, Step I has been completed
  if (stepIndex >= 1 || isClosed) {
    // Step I Decision Received (sometime after Step I Due or earlier)
    step1Rcvd = addDays(step1Due, Math.floor(Math.random() * 10) - 5);
    if (step1Rcvd < dateFiled) step1Rcvd = addDays(dateFiled, 15);

    // Step II Appeal Due (10 days after Step I Decision Received)
    step2AppealDue = addDays(step1Rcvd, 10);
  }

  // If Step II or beyond
  if (stepIndex >= 2 || (isClosed && stepIndex >= 1)) {
    // Step II Appeal Filed (within appeal window)
    step2AppealFiled = addDays(step1Rcvd, Math.floor(Math.random() * 8) + 1);

    // Step II Decision Due (30 days after appeal filed)
    step2Due = addDays(step2AppealFiled, 30);

    // Step II Decision Received
    step2Rcvd = addDays(step2Due, Math.floor(Math.random() * 10) - 5);
    if (step2Rcvd < step2AppealFiled) step2Rcvd = addDays(step2AppealFiled, 15);

    // Step III Appeal Due (30 days after Step II Decision Received)
    step3AppealDue = addDays(step2Rcvd, 30);
  }

  // If Step III or beyond
  if (stepIndex >= 3 || (isClosed && stepIndex >= 2)) {
    // Step III Appeal Filed
    step3AppealFiled = addDays(step2Rcvd, Math.floor(Math.random() * 20) + 1);
  }

  // Set date closed for resolved cases
  if (isClosed) {
    var lastDate = step3AppealFiled || step2Rcvd || step1Rcvd || dateFiled;
    dateClosed = addDays(lastDate, Math.floor(Math.random() * 30) + 5);
  }

  // Calculate Days Open
  var daysOpen = '';
  if (dateFiled) {
    var endDate = dateClosed || today;
    daysOpen = Math.floor((endDate - dateFiled) / (1000 * 60 * 60 * 24));
  }

  // Calculate Next Action Due based on current step
  var nextActionDue = '';
  if (!isClosed) {
    switch (step) {
      case 'Informal':
        nextActionDue = filingDeadline;
        break;
      case 'Step I':
        nextActionDue = step1Due || filingDeadline;
        break;
      case 'Step II':
        nextActionDue = step2Due || step2AppealDue || step1Due;
        break;
      case 'Step III':
      case 'Mediation':
      case 'Arbitration':
        nextActionDue = step3AppealDue || step2Due;
        break;
    }
  }

  // Calculate Days to Deadline
  var daysToDeadline = '';
  if (nextActionDue && !isClosed) {
    daysToDeadline = Math.floor((nextActionDue - today) / (1000 * 60 * 60 * 24));
  }

  var resolutions = ['Won - Full remedy', 'Won - Partial remedy', 'Settled - Compromise', 'Denied', 'Withdrawn', 'Pending'];
  var resolution = dateClosed ? randomChoice(resolutions) : '';

  return [
    grievanceId,              // 1: Grievance ID (A)
    memberId,                 // 2: Member ID (B)
    firstName,                // 3: First Name (C)
    lastName,                 // 4: Last Name (D)
    status,                   // 5: Status (E)
    step,                     // 6: Current Step (F)
    incidentDate,             // 7: Incident Date (G)
    filingDeadline,           // 8: Filing Deadline (H)
    dateFiled,                // 9: Date Filed (I)
    step1Due,                 // 10: Step I Due (J)
    step1Rcvd,                // 11: Step I Rcvd (K)
    step2AppealDue,           // 12: Step II Appeal Due (L)
    step2AppealFiled,         // 13: Step II Appeal Filed (M)
    step2Due,                 // 14: Step II Due (N)
    step2Rcvd,                // 15: Step II Rcvd (O)
    step3AppealDue,           // 16: Step III Appeal Due (P)
    step3AppealFiled,         // 17: Step III Appeal Filed (Q)
    dateClosed,               // 18: Date Closed (R)
    daysOpen,                 // 19: Days Open (S)
    nextActionDue,            // 20: Next Action Due (T)
    daysToDeadline,           // 21: Days to Deadline (U)
    articles,                 // 22: Articles Violated (V)
    category,                 // 23: Issue Category (W)
    email,                    // 24: Member Email (X)
    unit,                     // 25: Unit (Y)
    location,                 // 26: Location (Z)
    steward,                  // 27: Steward (AA)
    resolution,               // 28: Resolution (AB)
    false,                    // 29: Message Alert (AC) - checkbox
    '',                       // 30: Coordinator Message (AD)
    '',                       // 31: Acknowledged By (AE)
    '',                       // 32: Acknowledged Date (AF)
    '',                       // 33: Drive Folder ID (AG)
    ''                        // 34: Drive Folder URL (AH)
  ];
}

// ============================================================================
// DIALOG FUNCTIONS
// ============================================================================

/**
 * Show dialog to seed custom number of members (30% get grievances by default)
 */
function SEED_MEMBERS_DIALOG() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    '👥 Seed Members & Grievances',
    'How many members to seed? (30% will get grievances)\nMax 2000:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    var count = parseInt(response.getResponseText(), 10);
    if (isNaN(count) || count < 1) {
      ui.alert('Please enter a valid number.');
      return;
    }
    SEED_MEMBERS(count, 30);
  }
}

/**
 * Show dialog to seed custom number of members with custom grievance percentage
 */
function SEED_MEMBERS_ADVANCED_DIALOG() {
  var ui = SpreadsheetApp.getUi();

  var countResponse = ui.prompt(
    '👥 Seed Members (Step 1/2)',
    'How many members to seed? (max 2000)',
    ui.ButtonSet.OK_CANCEL
  );

  if (countResponse.getSelectedButton() !== ui.Button.OK) return;

  var count = parseInt(countResponse.getResponseText(), 10);
  if (isNaN(count) || count < 1) {
    ui.alert('Please enter a valid number.');
    return;
  }

  var percentResponse = ui.prompt(
    '📋 Grievance Percentage (Step 2/2)',
    'What percentage of members should have grievances? (0-100)\nDefault: 30',
    ui.ButtonSet.OK_CANCEL
  );

  if (percentResponse.getSelectedButton() !== ui.Button.OK) return;

  var percent = parseInt(percentResponse.getResponseText(), 10);
  if (isNaN(percent)) percent = 30;
  percent = Math.max(0, Math.min(100, percent));

  SEED_MEMBERS(count, percent);
}

/**
 * Show dialog to seed custom number of grievances
 */
function SEED_GRIEVANCES_DIALOG() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    '📋 Seed Grievances',
    'How many grievances to seed? (max 300)',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    var count = parseInt(response.getResponseText(), 10);
    if (isNaN(count) || count < 1) {
      ui.alert('Please enter a valid number.');
      return;
    }
    SEED_GRIEVANCES(count);
  }
}

/**
 * Seed 50 members with 30% grievances (shortcut)
 */
function seed50Members() {
  SEED_MEMBERS(50, 30);
}

/**
 * Seed 100 members with 50% grievances (shortcut)
 */
function seed100MembersWithGrievances() {
  SEED_MEMBERS(100, 50);
}

/**
 * Seed 25 grievances for existing members (shortcut)
 */
function seed25Grievances() {
  SEED_GRIEVANCES(25);
}

// ============================================================================
// NUKE FUNCTIONS
// ============================================================================

/**
 * Delete all seeded data from Member Directory and Grievance Log
 * Uses pattern matching (M/G + 4 letters + 3 digits) to identify seeded IDs
 * This is more reliable than Script Properties tracking which has size limits
 */
function NUKE_SEEDED_DATA() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Check if already disabled
  if (isDemoModeDisabled()) {
    ui.alert('Demo Mode Disabled', 'Demo mode has already been disabled. The Demo menu will be removed on next refresh.', ui.ButtonSet.OK);
    return;
  }

  // Pattern for seeded IDs: M/G + 4 uppercase letters + 3 digits (e.g., MJOSM123, GJOSM456)
  var seededIdPattern = /^[MG][A-Z]{4}\d{3}$/;

  // Count seeded data by pattern
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  var memberCount = 0;
  var grievanceCount = 0;

  if (memberSheet && memberSheet.getLastRow() > 1) {
    var memberIds = memberSheet.getRange(2, 1, memberSheet.getLastRow() - 1, 1).getValues();
    memberIds.forEach(function(row) {
      if (row[0] && seededIdPattern.test(String(row[0]))) memberCount++;
    });
  }

  if (grievanceSheet && grievanceSheet.getLastRow() > 1) {
    var grievanceIds = grievanceSheet.getRange(2, 1, grievanceSheet.getLastRow() - 1, 1).getValues();
    grievanceIds.forEach(function(row) {
      if (row[0] && seededIdPattern.test(String(row[0]))) grievanceCount++;
    });
  }

  var response = ui.alert(
    '☢️ NUKE SEEDED DATA',
    '⚠️ This will permanently delete seeded/demo data:\n\n' +
    '• ' + memberCount + ' seeded members (ID pattern: M****###)\n' +
    '• ' + grievanceCount + ' seeded grievances (ID pattern: G****###)\n' +
    '• Config dropdown values\n\n' +
    '✅ Manually entered data with different ID formats will be PRESERVED.\n\n' +
    '⚠️ After nuke, the Demo menu will be permanently disabled.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  // Double confirm
  var response2 = ui.alert(
    '☢️ FINAL CONFIRMATION',
    'This will:\n' +
    '1. Delete ' + memberCount + ' seeded members\n' +
    '2. Delete ' + grievanceCount + ' seeded grievances\n' +
    '3. Permanently disable the Demo menu\n\n' +
    'Are you sure?',
    ui.ButtonSet.YES_NO
  );

  if (response2 !== ui.Button.YES) {
    return;
  }

  ss.toast('Nuking seeded data...', '☢️ NUKE', 3);

  try {
    var deletedMembers = 0;
    var deletedGrievances = 0;

    // Delete seeded grievances first (they reference members)
    if (grievanceSheet && grievanceSheet.getLastRow() > 1) {
      var grievanceData = grievanceSheet.getRange(2, 1, grievanceSheet.getLastRow() - 1, 1).getValues();
      // Delete from bottom up to preserve row indices
      for (var g = grievanceData.length - 1; g >= 0; g--) {
        var gId = String(grievanceData[g][0] || '');
        if (seededIdPattern.test(gId)) {
          grievanceSheet.deleteRow(g + 2);
          deletedGrievances++;
        }
      }
    }

    // Delete seeded members
    if (memberSheet && memberSheet.getLastRow() > 1) {
      var memberData = memberSheet.getRange(2, 1, memberSheet.getLastRow() - 1, 1).getValues();
      // Delete from bottom up to preserve row indices
      for (var m = memberData.length - 1; m >= 0; m--) {
        var mId = String(memberData[m][0] || '');
        if (seededIdPattern.test(mId)) {
          memberSheet.deleteRow(m + 2);
          deletedMembers++;
        }
      }
    }

    // Clear Config dropdowns (keep headers and default values)
    NUKE_CONFIG_DROPDOWNS();

    // Clear tracked IDs from Script Properties
    var props = PropertiesService.getScriptProperties();
    props.deleteProperty('SEEDED_MEMBER_IDS');
    props.deleteProperty('SEEDED_GRIEVANCE_IDS');

    // Disable demo mode permanently
    disableDemoMode();

    ss.toast('Seeded data nuked! Demo mode disabled.', '☢️ Complete', 5);
    ui.alert('☢️ Complete',
      'Seeded data has been deleted:\n' +
      '• ' + deletedMembers + ' members removed\n' +
      '• ' + deletedGrievances + ' grievances removed\n\n' +
      'Demo mode has been permanently disabled.\n' +
      'Refresh the page to remove the Demo menu.',
      ui.ButtonSet.OK);

  } catch (error) {
    Logger.log('Error in NUKE_SEEDED_DATA: ' + error.message);
    ui.alert('❌ Error', 'Nuke failed: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * Clear only Config dropdown values (user-populated columns)
 */
function NUKE_CONFIG_DROPDOWNS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!sheet) {
    return;
  }

  // Clear user-populated columns only (keep system defaults)
  var userColumns = [
    CONFIG_COLS.JOB_TITLES,
    CONFIG_COLS.OFFICE_LOCATIONS,
    CONFIG_COLS.UNITS,
    CONFIG_COLS.SUPERVISORS,
    CONFIG_COLS.MANAGERS,
    CONFIG_COLS.STEWARDS,
    CONFIG_COLS.STEWARD_COMMITTEES,
    CONFIG_COLS.GRIEVANCE_COORDINATORS,
    CONFIG_COLS.HOME_TOWNS
  ];

  userColumns.forEach(function(col) {
    var lastRow = sheet.getLastRow();
    if (lastRow > 2) {
      sheet.getRange(3, col, lastRow - 2, 1).clear();
    }
  });

  SpreadsheetApp.getActiveSpreadsheet().toast('Config dropdowns cleared!', '🧹 Cleared', 3);
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Get random element from array
 */
function randomChoice(arr) {
  return arr[Math.floor(Math.random() * arr.length)];
}

/**
 * Shuffle array using Fisher-Yates algorithm
 * Returns a new shuffled array (does not modify original)
 */
function shuffleArray(arr) {
  var shuffled = arr.slice(); // Create a copy
  for (var i = shuffled.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = shuffled[i];
    shuffled[i] = shuffled[j];
    shuffled[j] = temp;
  }
  return shuffled;
}

/**
 * Get random date between two dates
 */
function randomDate(start, end) {
  var startTime = start.getTime();
  var endTime = end.getTime();
  var randomTime = startTime + Math.random() * (endTime - startTime);
  return new Date(randomTime);
}

/**
 * Add days to a date
 */
function addDays(date, days) {
  if (!date) return '';
  var result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}
