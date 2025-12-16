/**
 * 509 Dashboard - Seed and Nuke Functions
 *
 * Functions for seeding sample data and clearing data.
 *
 * @version 1.0.0
 * @license Free for use by non-profit collective bargaining groups and unions
 */

// ============================================================================
// SEED FUNCTIONS
// ============================================================================

/**
 * Seed all sample data: Config + 50 members + 25 grievances
 */
function SEED_SAMPLE_DATA() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var response = ui.alert(
    '🚀 Seed Sample Data',
    'This will seed:\n' +
    '• Config dropdowns (Job Titles, Locations, etc.)\n' +
    '• 50 sample members\n' +
    '• 25 sample grievances\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  ss.toast('Seeding config data...', '🌱 Seeding', 3);
  seedConfigData();

  ss.toast('Seeding members...', '🌱 Seeding', 3);
  SEED_MEMBERS(50);

  ss.toast('Seeding grievances...', '🌱 Seeding', 3);
  SEED_GRIEVANCES(25);

  ss.toast('Sample data seeded successfully!', '✅ Success', 5);
  ui.alert('✅ Success', 'Sample data has been seeded!\n\n' +
    '• Config dropdowns populated\n' +
    '• 50 members added\n' +
    '• 25 grievances added', ui.ButtonSet.OK);
}

/**
 * Seed Config sheet with dropdown values
 */
function seedConfigData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error: Config sheet not found. Please run CREATE_509_DASHBOARD first.');
    return;
  }

  // Job Titles (Column A)
  var jobTitles = [
    'Social Worker', 'Case Manager', 'Program Coordinator', 'Administrative Assistant',
    'Supervisor', 'Director', 'Clinician', 'Counselor', 'Specialist', 'Analyst',
    'Manager', 'Senior Social Worker', 'Lead Case Manager', 'Program Manager',
    'Executive Assistant', 'HR Coordinator', 'Finance Associate', 'IT Support',
    'Communications Specialist', 'Outreach Worker'
  ];
  sheet.getRange(2, CONFIG_COLS.JOB_TITLES, jobTitles.length, 1)
    .setValues(jobTitles.map(function(v) { return [v]; }));

  // Office Locations (Column B)
  var locations = [
    'Boston Main Office', 'Worcester Regional', 'Springfield Center', 'Cambridge Branch',
    'Lowell Office', 'Brockton Center', 'Quincy Regional', 'New Bedford Office',
    'Fall River Branch', 'Lawrence Center', 'Framingham Office', 'Somerville Branch',
    'Lynn Regional', 'Haverhill Center', 'Malden Office', 'Medford Branch',
    'Waltham Regional', 'Newton Center', 'Brookline Office', 'Salem Branch'
  ];
  sheet.getRange(2, CONFIG_COLS.OFFICE_LOCATIONS, locations.length, 1)
    .setValues(locations.map(function(v) { return [v]; }));

  // Units (Column C)
  var units = [
    'Child Welfare', 'Adult Services', 'Mental Health', 'Disability Services',
    'Elder Affairs', 'Housing Assistance', 'Employment Services', 'Youth Services',
    'Family Support', 'Administration'
  ];
  sheet.getRange(2, CONFIG_COLS.UNITS, units.length, 1)
    .setValues(units.map(function(v) { return [v]; }));

  // Supervisors (Column F)
  var supervisors = [
    'Maria Rodriguez', 'James Wilson', 'Sarah Chen', 'Michael Brown',
    'Jennifer Davis', 'Robert Taylor', 'Lisa Anderson', 'David Martinez',
    'Emily Johnson', 'Christopher Lee', 'Amanda White', 'Daniel Garcia'
  ];
  sheet.getRange(2, CONFIG_COLS.SUPERVISORS, supervisors.length, 1)
    .setValues(supervisors.map(function(v) { return [v]; }));

  // Managers (Column G)
  var managers = [
    'Patricia Thompson', 'William Jackson', 'Elizabeth Moore', 'Richard Harris',
    'Susan Clark', 'Joseph Lewis', 'Margaret Robinson', 'Charles Walker'
  ];
  sheet.getRange(2, CONFIG_COLS.MANAGERS, managers.length, 1)
    .setValues(managers.map(function(v) { return [v]; }));

  // Stewards (Column H)
  var stewards = [
    'John Smith', 'Mary Johnson', 'Robert Williams', 'Patricia Jones',
    'Michael Davis', 'Linda Miller', 'William Brown', 'Barbara Wilson',
    'David Moore', 'Susan Taylor', 'James Anderson', 'Karen Thomas'
  ];
  sheet.getRange(2, CONFIG_COLS.STEWARDS, stewards.length, 1)
    .setValues(stewards.map(function(v) { return [v]; }));

  // Home Towns (Column AF - 32)
  var homeTowns = [
    'Boston', 'Worcester', 'Springfield', 'Cambridge', 'Lowell', 'Brockton',
    'Quincy', 'New Bedford', 'Fall River', 'Lawrence', 'Framingham', 'Somerville',
    'Lynn', 'Haverhill', 'Malden', 'Medford', 'Waltham', 'Newton', 'Brookline'
  ];
  sheet.getRange(2, CONFIG_COLS.HOME_TOWNS, homeTowns.length, 1)
    .setValues(homeTowns.map(function(v) { return [v]; }));

  SpreadsheetApp.getActiveSpreadsheet().toast('Config data seeded!', '✅ Success', 3);
}

/**
 * Seed N members
 * @param {number} count - Number of members to seed (max 2000)
 */
function SEED_MEMBERS(count) {
  count = Math.min(count || 50, 2000);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);

  if (!sheet || !configSheet) {
    SpreadsheetApp.getUi().alert('Error: Required sheets not found.');
    return;
  }

  // Get config values
  var jobTitles = getConfigValues(configSheet, CONFIG_COLS.JOB_TITLES);
  var locations = getConfigValues(configSheet, CONFIG_COLS.OFFICE_LOCATIONS);
  var units = getConfigValues(configSheet, CONFIG_COLS.UNITS);
  var supervisors = getConfigValues(configSheet, CONFIG_COLS.SUPERVISORS);
  var managers = getConfigValues(configSheet, CONFIG_COLS.MANAGERS);
  var stewards = getConfigValues(configSheet, CONFIG_COLS.STEWARDS);

  // If config is empty, use defaults
  if (jobTitles.length === 0) jobTitles = ['Social Worker', 'Case Manager', 'Supervisor'];
  if (locations.length === 0) locations = ['Boston Main Office', 'Worcester Regional'];
  if (units.length === 0) units = ['Child Welfare', 'Adult Services'];
  if (supervisors.length === 0) supervisors = ['Jane Supervisor'];
  if (managers.length === 0) managers = ['John Manager'];
  if (stewards.length === 0) stewards = ['Mary Steward'];

  var firstNames = ['James', 'Mary', 'John', 'Patricia', 'Robert', 'Jennifer', 'Michael', 'Linda', 'William', 'Elizabeth', 'David', 'Barbara', 'Richard', 'Susan', 'Joseph', 'Jessica', 'Thomas', 'Sarah', 'Charles', 'Karen'];
  var lastNames = ['Smith', 'Johnson', 'Williams', 'Brown', 'Jones', 'Garcia', 'Miller', 'Davis', 'Rodriguez', 'Martinez', 'Hernandez', 'Lopez', 'Gonzalez', 'Wilson', 'Anderson', 'Thomas', 'Taylor', 'Moore', 'Jackson', 'Martin'];
  var officeDays = DEFAULT_CONFIG.OFFICE_DAYS;
  var commMethods = DEFAULT_CONFIG.COMM_METHODS;

  var startRow = Math.max(sheet.getLastRow() + 1, 2);
  var existingCount = startRow - 2;

  var rows = [];
  var batchSize = 50;

  for (var i = 0; i < count; i++) {
    var memberId = 'M' + String(existingCount + i + 1).padStart(6, '0');
    var firstName = randomChoice(firstNames);
    var lastName = randomChoice(lastNames);
    var email = firstName.toLowerCase() + '.' + lastName.toLowerCase() + (existingCount + i + 1) + '@example.org';
    var phone = '617-555-' + String(1000 + i).padStart(4, '0');
    var isSteward = Math.random() < 0.1 ? 'Yes' : 'No';

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
      isSteward === 'Yes' ? 'Grievance Committee' : '',
      randomChoice(stewards)
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

  SpreadsheetApp.getActiveSpreadsheet().toast(count + ' members seeded!', '✅ Success', 3);
}

/**
 * Generate a single member row with all 31 columns
 */
function generateSingleMemberRow(memberId, firstName, lastName, jobTitle, location, unit, officeDays, email, phone, prefComm, bestTime, supervisor, manager, isSteward, committees, assignedSteward) {
  var today = new Date();
  var lastMonth = new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000);

  return [
    memberId,                                    // 1: Member ID
    firstName,                                   // 2: First Name
    lastName,                                    // 3: Last Name
    jobTitle,                                    // 4: Job Title
    location,                                    // 5: Work Location
    unit,                                        // 6: Unit
    officeDays,                                  // 7: Office Days
    email,                                       // 8: Email
    phone,                                       // 9: Phone
    prefComm,                                    // 10: Preferred Communication
    bestTime,                                    // 11: Best Time
    supervisor,                                  // 12: Supervisor
    manager,                                     // 13: Manager
    isSteward,                                   // 14: Is Steward
    committees,                                  // 15: Committees
    assignedSteward,                             // 16: Assigned Steward
    randomDate(lastMonth, today),                // 17: Last Virtual Mtg
    randomDate(lastMonth, today),                // 18: Last In-Person Mtg
    Math.floor(Math.random() * 100),             // 19: Open Rate %
    Math.floor(Math.random() * 20),              // 20: Volunteer Hours
    Math.random() < 0.3 ? 'Yes' : 'No',          // 21: Interest Local
    Math.random() < 0.2 ? 'Yes' : 'No',          // 22: Interest Chapter
    Math.random() < 0.1 ? 'Yes' : 'No',          // 23: Interest Allied
    '',                                          // 24: Home Town
    '',                                          // 25: Recent Contact Date
    '',                                          // 26: Contact Steward
    '',                                          // 27: Contact Notes
    '',                                          // 28: Has Open Grievance (calculated)
    '',                                          // 29: Grievance Status (calculated)
    '',                                          // 30: Next Deadline (calculated)
    false                                        // 31: Start Grievance (checkbox)
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

  var startRow = Math.max(grievanceSheet.getLastRow() + 1, 2);
  var existingCount = startRow - 2;

  var rows = [];
  var batchSize = 25;
  var today = new Date();

  for (var i = 0; i < count; i++) {
    // Pick a random member - ensure we get a valid member with an ID
    var memberRow = memberData[1 + Math.floor(Math.random() * (memberData.length - 1))];
    var memberId = memberRow[MEMBER_COLS.MEMBER_ID - 1];

    // Skip if no member ID
    if (!memberId) continue;

    var grievanceId = 'G-' + String(existingCount + i + 1).padStart(5, '0');
    var incidentDate = randomDate(new Date(today.getTime() - 90 * 24 * 60 * 60 * 1000), today);
    var status = randomChoice(statuses);
    var step = randomChoice(steps);

    // Generate grievance row - First Name, Last Name, Email, Unit, Location, Steward
    // will be auto-populated by formulas based on Member ID
    var row = generateSingleGrievanceRowForSeed(
      grievanceId,
      memberId,
      status,
      step,
      incidentDate,
      randomChoice(articles),
      randomChoice(categories)
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

  // Sync data from hidden formulas sheet (self-healing)
  syncGrievanceFormulasToLog();

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
 * Show dialog to seed custom number of members
 */
function SEED_MEMBERS_DIALOG() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    '👥 Seed Members',
    'How many members to seed? (max 2000)',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    var count = parseInt(response.getResponseText(), 10);
    if (isNaN(count) || count < 1) {
      ui.alert('Please enter a valid number.');
      return;
    }
    SEED_MEMBERS(count);
  }
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
 * Seed 50 members shortcut
 */
function seed50Members() {
  SEED_MEMBERS(50);
}

/**
 * Seed 25 grievances shortcut
 */
function seed25Grievances() {
  SEED_GRIEVANCES(25);
}

// ============================================================================
// NUKE FUNCTIONS
// ============================================================================

/**
 * Clear all member and grievance data
 */
function NUKE_ALL_DATA() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var response = ui.alert(
    '☢️ NUKE ALL DATA',
    '⚠️ WARNING: This will permanently delete:\n\n' +
    '• All members in Member Directory\n' +
    '• All grievances in Grievance Log\n' +
    '• All Config dropdown values\n\n' +
    'This cannot be undone!\n\n' +
    'Are you absolutely sure?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  // Double confirm
  var response2 = ui.alert(
    '☢️ FINAL CONFIRMATION',
    'Type "NUKE" to confirm deletion of all data.',
    ui.ButtonSet.YES_NO
  );

  if (response2 !== ui.Button.YES) {
    return;
  }

  ss.toast('Nuking all data...', '☢️ NUKE', 3);

  try {
    // Clear Member Directory (keep headers)
    var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);
    if (memberSheet && memberSheet.getLastRow() > 1) {
      memberSheet.getRange(2, 1, memberSheet.getLastRow() - 1, memberSheet.getLastColumn()).clear();
    }

    // Clear Grievance Log (keep headers)
    var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
    if (grievanceSheet && grievanceSheet.getLastRow() > 1) {
      grievanceSheet.getRange(2, 1, grievanceSheet.getLastRow() - 1, grievanceSheet.getLastColumn()).clear();
    }

    // Clear Config dropdowns (keep headers and default values)
    NUKE_CONFIG_DROPDOWNS();

    ss.toast('All data has been nuked!', '☢️ Complete', 5);
    ui.alert('☢️ Complete', 'All data has been deleted.', ui.ButtonSet.OK);

  } catch (error) {
    Logger.log('Error in NUKE_ALL_DATA: ' + error.message);
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
    CONFIG_COLS.GRIEVANCE_COORDINATORS,
    CONFIG_COLS.HOME_TOWNS
  ];

  userColumns.forEach(function(col) {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, col, lastRow - 1, 1).clear();
    }
  });

  SpreadsheetApp.getActiveSpreadsheet().toast('Config dropdowns cleared!', '🧹 Cleared', 3);
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Get values from a Config column (excluding header and empty cells)
 */
function getConfigValues(configSheet, column) {
  var lastRow = configSheet.getLastRow();
  if (lastRow < 2) return [];

  var values = configSheet.getRange(2, column, lastRow - 1, 1).getValues();
  return values
    .map(function(row) { return row[0]; })
    .filter(function(v) { return v !== '' && v !== null; });
}

/**
 * Get random element from array
 */
function randomChoice(arr) {
  return arr[Math.floor(Math.random() * arr.length)];
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
