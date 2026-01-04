/**
 * 509 Dashboard - Hidden Sheet Architecture
 *
 * Self-healing hidden calculation sheets with auto-sync triggers.
 * Provides automatic cross-sheet data population.
 *
 * ⚠️ WARNING: DO NOT DEPLOY THIS FILE DIRECTLY
 * This is a source file used to generate ConsolidatedDashboard.gs.
 * Deploy ONLY ConsolidatedDashboard.gs to avoid function conflicts.
 *
 * @version 1.0.0
 * @license Free for use by non-profit collective bargaining groups and unions
 */

// ============================================================================
// HIDDEN SHEET 1: _Grievance_Calc
// Source: Grievance Log → Destination: Member Directory (AB-AD)
// ============================================================================

/**
 * Setup the _Grievance_Calc hidden sheet with self-healing formulas
 * Calculates: Has Open Grievance, Grievance Status, Next Deadline per member
 */
function setupGrievanceCalcSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_CALC);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.GRIEVANCE_CALC);
  }

  sheet.clear();

  // Headers
  var headers = ['Member ID', 'Has Open Grievance', 'Grievance Status', 'Days to Deadline', 'Total Count', 'Win Rate %', 'Last Grievance Date'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  // Get column letters for dynamic formulas
  var memberIdCol = getColumnLetter(MEMBER_COLS.MEMBER_ID);
  var gMemberIdCol = getColumnLetter(GRIEVANCE_COLS.MEMBER_ID);
  var gStatusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);
  var gNextActionCol = getColumnLetter(GRIEVANCE_COLS.NEXT_ACTION_DUE);
  var gResolutionCol = getColumnLetter(GRIEVANCE_COLS.RESOLUTION);
  var gDateFiledCol = getColumnLetter(GRIEVANCE_COLS.DATE_FILED);

  // Formula for Member IDs (Column A) - pulls unique member IDs from Member Directory
  var memberIdFormula = '=IFERROR(FILTER(\'' + SHEETS.MEMBER_DIR + '\'!' + memberIdCol + ':' + memberIdCol + ',\'' + SHEETS.MEMBER_DIR + '\'!' + memberIdCol + ':' + memberIdCol + '<>"Member ID"),"")';
  sheet.getRange('A2').setFormula(memberIdFormula);

  // Formulas for calculations (using ARRAYFORMULA for efficiency)
  // Column B: Has Open Grievance
  var hasOpenFormula = '=ARRAYFORMULA(IF(A2:A="","",IF(COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Open")+COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Pending Info")>0,"Yes","No")))';
  sheet.getRange('B2').setFormula(hasOpenFormula);

  // Column C: Grievance Status (most urgent: Open > Pending Info, blank if all closed)
  var statusFormula = '=ARRAYFORMULA(IF(A2:A="","",IF(COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Open")>0,"Open",IF(COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Pending Info")>0,"Pending Info",""))))';
  sheet.getRange('C2').setFormula(statusFormula);

  // Column D: Days to Deadline (minimum/most urgent deadline for open grievances only)
  // Excludes all closed statuses: Closed, Settled, Withdrawn, Denied, Won
  var gDaysToDeadlineCol = getColumnLetter(GRIEVANCE_COLS.DAYS_TO_DEADLINE);
  var deadlineFormula = '=ARRAYFORMULA(IF(A2:A="","",IFERROR(MINIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysToDeadlineCol + ':' + gDaysToDeadlineCol + ',\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"<>Closed",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"<>Settled",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"<>Withdrawn",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"<>Denied",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"<>Won"),"")))';
  sheet.getRange('D2').setFormula(deadlineFormula);

  // Column E: Total Grievance Count
  var countFormula = '=ARRAYFORMULA(IF(A2:A="","",COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A)))';
  sheet.getRange('E2').setFormula(countFormula);

  // Column F: Win Rate %
  var winRateFormula = '=ARRAYFORMULA(IF(A2:A="","",IFERROR(COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gResolutionCol + ':' + gResolutionCol + ',"*Won*")/COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A)*100,0)))';
  sheet.getRange('F2').setFormula(winRateFormula);

  // Column G: Last Grievance Date
  var lastDateFormula = '=ARRAYFORMULA(IF(A2:A="","",IFERROR(MAXIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateFiledCol + ':' + gDateFiledCol + ',\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A),"")))';
  sheet.getRange('G2').setFormula(lastDateFormula);

  // Hide the sheet
  sheet.hideSheet();

  Logger.log('_Grievance_Calc sheet setup complete');
}

/**
 * Sync grievance data directly from Grievance Log to Member Directory
 * Calculates Has Open Grievance, Status, and Days to Deadline per member
 * Fixed in v1.6.0: Now calculates directly instead of using MINIFS (which ignores "Overdue" text)
 */
function syncGrievanceToMemberDirectory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!grievanceSheet || !memberSheet) {
    Logger.log('Required sheets not found for grievance sync');
    return;
  }

  // Get grievance data
  var grievanceData = grievanceSheet.getDataRange().getValues();
  if (grievanceData.length < 2) return;

  // Closed statuses - grievances with these statuses don't count as "open"
  var closedStatuses = ['Closed', 'Settled', 'Withdrawn', 'Denied', 'Won'];

  // Build lookup map: memberId -> {hasOpen, status, deadline}
  // Calculate directly from grievance data (handles "Overdue" text properly)
  var lookup = {};

  for (var i = 1; i < grievanceData.length; i++) {
    var row = grievanceData[i];
    var memberId = row[GRIEVANCE_COLS.MEMBER_ID - 1];
    if (!memberId) continue;

    var status = row[GRIEVANCE_COLS.STATUS - 1] || '';
    var daysToDeadline = row[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
    var isClosed = closedStatuses.indexOf(status) !== -1;

    // Initialize member entry if not exists
    if (!lookup[memberId]) {
      lookup[memberId] = {
        hasOpen: 'No',
        status: '',
        deadline: '',
        minDeadline: Infinity,  // Track minimum numeric deadline
        hasOverdue: false       // Track if any grievance is overdue
      };
    }

    // Check if this grievance is open/pending
    if (!isClosed) {
      lookup[memberId].hasOpen = 'Yes';

      // Set status priority: Open > Pending Info
      if (status === 'Open') {
        lookup[memberId].status = 'Open';
      } else if (status === 'Pending Info' && lookup[memberId].status !== 'Open') {
        lookup[memberId].status = 'Pending Info';
      }

      // Handle Days to Deadline (can be number or "Overdue" text)
      if (daysToDeadline === 'Overdue') {
        lookup[memberId].hasOverdue = true;
      } else if (typeof daysToDeadline === 'number' && daysToDeadline < lookup[memberId].minDeadline) {
        lookup[memberId].minDeadline = daysToDeadline;
      }
    }
  }

  // Finalize deadline values
  for (var mid in lookup) {
    var data = lookup[mid];
    if (data.hasOpen === 'Yes') {
      if (data.minDeadline !== Infinity) {
        // Has a numeric deadline - use the minimum
        data.deadline = data.minDeadline;
      } else if (data.hasOverdue) {
        // All open grievances are overdue
        data.deadline = 'Overdue';
      }
    }
  }

  // Get member data
  var memberData = memberSheet.getDataRange().getValues();
  if (memberData.length < 2) return;

  // Update columns AB-AD (Has Open Grievance?, Grievance Status, Days to Deadline)
  var updates = [];
  for (var j = 1; j < memberData.length; j++) {
    var memberId = memberData[j][MEMBER_COLS.MEMBER_ID - 1];
    var memberInfo = lookup[memberId] || {hasOpen: 'No', status: '', deadline: ''};
    updates.push([memberInfo.hasOpen, memberInfo.status, memberInfo.deadline]);
  }

  if (updates.length > 0) {
    memberSheet.getRange(2, MEMBER_COLS.HAS_OPEN_GRIEVANCE, updates.length, 3).setValues(updates);
  }

  Logger.log('Synced grievance data to ' + updates.length + ' members');
}

// ============================================================================
// HIDDEN SHEET 2: _Grievance_Formulas (SELF-HEALING)
// Source: Grievance Log → Destination: Grievance Log (calculated columns)
// This sheet contains all auto-calculated formulas and syncs them back
// ============================================================================

/**
 * Setup the _Grievance_Formulas hidden sheet with self-healing formulas
 * Calculates: First Name, Last Name, Email, Unit, Location, Steward (from Member Dir)
 *            Filing Deadline, Step I-III dates, Days Open, Next Action Due, Days to Deadline
 */
function setupGrievanceFormulasSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_FORMULAS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.GRIEVANCE_FORMULAS);
  }

  sheet.clear();

  // Headers matching Grievance Log columns that need formulas
  var headers = [
    'Row Index',           // A - For tracking which row in Grievance Log
    'Member ID',           // B - From Grievance Log
    'First Name',          // C - Lookup from Member Directory
    'Last Name',           // D - Lookup from Member Directory
    'Incident Date',       // E - From Grievance Log
    'Date Filed',          // F - From Grievance Log
    'Step I Rcvd',         // G - From Grievance Log
    'Step II Appeal Filed',// H - From Grievance Log
    'Step II Rcvd',        // I - From Grievance Log
    'Status',              // J - From Grievance Log
    'Current Step',        // K - From Grievance Log
    'Date Closed',         // L - From Grievance Log
    'Filing Deadline',     // M - CALCULATED
    'Step I Due',          // N - CALCULATED
    'Step II Appeal Due',  // O - CALCULATED
    'Step II Due',         // P - CALCULATED
    'Step III Appeal Due', // Q - CALCULATED
    'Days Open',           // R - CALCULATED
    'Next Action Due',     // S - CALCULATED
    'Days to Deadline',    // T - CALCULATED
    'Member Email',        // U - Lookup from Member Directory
    'Unit',                // V - Lookup from Member Directory
    'Location',            // W - Lookup from Member Directory
    'Steward'              // X - Lookup from Member Directory
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  // Get column letters for Grievance Log source data
  var gGrievanceIdCol = getColumnLetter(GRIEVANCE_COLS.GRIEVANCE_ID);     // A
  var gMemberIdCol = getColumnLetter(GRIEVANCE_COLS.MEMBER_ID);           // B
  var gIncidentDateCol = getColumnLetter(GRIEVANCE_COLS.INCIDENT_DATE);   // G
  var gDateFiledCol = getColumnLetter(GRIEVANCE_COLS.DATE_FILED);         // I
  var gStep1RcvdCol = getColumnLetter(GRIEVANCE_COLS.STEP1_RCVD);         // K
  var gStep2AppealFiledCol = getColumnLetter(GRIEVANCE_COLS.STEP2_APPEAL_FILED); // M
  var gStep2RcvdCol = getColumnLetter(GRIEVANCE_COLS.STEP2_RCVD);         // O
  var gStatusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);                // E
  var gCurrentStepCol = getColumnLetter(GRIEVANCE_COLS.CURRENT_STEP);     // F
  var gDateClosedCol = getColumnLetter(GRIEVANCE_COLS.DATE_CLOSED);       // R

  // Member Directory columns for lookups
  var mMemberIdCol = getColumnLetter(MEMBER_COLS.MEMBER_ID);
  var mStewardCol = getColumnLetter(MEMBER_COLS.ASSIGNED_STEWARD);
  var memberRange = "'" + SHEETS.MEMBER_DIR + "'!" + mMemberIdCol + ":" + mStewardCol;

  // Column A: Row Index (ROW()-1 to match Grievance Log rows)
  // Pull unique grievance IDs to create row mapping
  sheet.getRange('A2').setFormula(
    '=IFERROR(FILTER(ROW(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gGrievanceIdCol + '2:' + gGrievanceIdCol + ')-1,' +
    '\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gGrievanceIdCol + '2:' + gGrievanceIdCol + '<>""),"")'
  );

  // Column B: Member ID (from Grievance Log)
  sheet.getRange('B2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',A2:A+1)))'
  );

  // Column C: First Name (VLOOKUP from Member Directory)
  sheet.getRange('C2').setFormula(
    '=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,' + memberRange + ',' + MEMBER_COLS.FIRST_NAME + ',FALSE),"")))'
  );

  // Column D: Last Name (VLOOKUP from Member Directory)
  sheet.getRange('D2').setFormula(
    '=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,' + memberRange + ',' + MEMBER_COLS.LAST_NAME + ',FALSE),"")))'
  );

  // Column E: Incident Date (from Grievance Log)
  sheet.getRange('E2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gIncidentDateCol + ':' + gIncidentDateCol + ',A2:A+1)))'
  );

  // Column F: Date Filed (from Grievance Log)
  sheet.getRange('F2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateFiledCol + ':' + gDateFiledCol + ',A2:A+1)))'
  );

  // Column G: Step I Rcvd (from Grievance Log)
  sheet.getRange('G2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStep1RcvdCol + ':' + gStep1RcvdCol + ',A2:A+1)))'
  );

  // Column H: Step II Appeal Filed (from Grievance Log)
  sheet.getRange('H2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStep2AppealFiledCol + ':' + gStep2AppealFiledCol + ',A2:A+1)))'
  );

  // Column I: Step II Rcvd (from Grievance Log)
  sheet.getRange('I2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStep2RcvdCol + ':' + gStep2RcvdCol + ',A2:A+1)))'
  );

  // Column J: Status (from Grievance Log)
  sheet.getRange('J2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',A2:A+1)))'
  );

  // Column K: Current Step (from Grievance Log)
  sheet.getRange('K2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gCurrentStepCol + ':' + gCurrentStepCol + ',A2:A+1)))'
  );

  // Column L: Date Closed (from Grievance Log)
  sheet.getRange('L2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",INDEX(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',A2:A+1)))'
  );

  // =========== CALCULATED COLUMNS ===========

  // Column M: Filing Deadline = Incident Date + 21 days
  sheet.getRange('M2').setFormula(
    '=ARRAYFORMULA(IF(E2:E="","",E2:E+21))'
  );

  // Column N: Step I Due = Date Filed + 30 days
  sheet.getRange('N2').setFormula(
    '=ARRAYFORMULA(IF(F2:F="","",F2:F+30))'
  );

  // Column O: Step II Appeal Due = Step I Rcvd + 10 days
  sheet.getRange('O2').setFormula(
    '=ARRAYFORMULA(IF(G2:G="","",G2:G+10))'
  );

  // Column P: Step II Due = Step II Appeal Filed + 30 days
  sheet.getRange('P2').setFormula(
    '=ARRAYFORMULA(IF(H2:H="","",H2:H+30))'
  );

  // Column Q: Step III Appeal Due = Step II Rcvd + 30 days
  sheet.getRange('Q2').setFormula(
    '=ARRAYFORMULA(IF(I2:I="","",I2:I+30))'
  );

  // Column R: Days Open = IF closed: Date Closed - Date Filed, ELSE: Today - Date Filed
  sheet.getRange('R2').setFormula(
    '=ARRAYFORMULA(IF(F2:F="","",IF(L2:L<>"",L2:L-F2:F,TODAY()-F2:F)))'
  );

  // Column S: Next Action Due = Based on current step and status
  // If closed status, leave blank; otherwise return appropriate deadline
  sheet.getRange('S2').setFormula(
    '=ARRAYFORMULA(IF(J2:J="","",' +
    'IF(OR(J2:J="Settled",J2:J="Withdrawn",J2:J="Denied",J2:J="Won",J2:J="Closed"),"",' +
    'IF(K2:K="Informal",M2:M,' +
    'IF(K2:K="Step I",N2:N,' +
    'IF(K2:K="Step II",P2:P,' +
    'Q2:Q))))))'
  );

  // Column T: Days to Deadline = Next Action Due - Today
  sheet.getRange('T2').setFormula(
    '=ARRAYFORMULA(IF(S2:S="","",S2:S-TODAY()))'
  );

  // =========== MEMBER LOOKUP COLUMNS ===========

  // Column U: Member Email (VLOOKUP from Member Directory)
  sheet.getRange('U2').setFormula(
    '=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,' + memberRange + ',' + MEMBER_COLS.EMAIL + ',FALSE),"")))'
  );

  // Column V: Unit (VLOOKUP from Member Directory)
  sheet.getRange('V2').setFormula(
    '=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,' + memberRange + ',' + MEMBER_COLS.UNIT + ',FALSE),"")))'
  );

  // Column W: Location (VLOOKUP from Member Directory)
  sheet.getRange('W2').setFormula(
    '=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,' + memberRange + ',' + MEMBER_COLS.WORK_LOCATION + ',FALSE),"")))'
  );

  // Column X: Steward (VLOOKUP from Member Directory)
  sheet.getRange('X2').setFormula(
    '=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,' + memberRange + ',' + MEMBER_COLS.ASSIGNED_STEWARD + ',FALSE),"")))'
  );

  // Format date columns (MM/dd/yyyy)
  sheet.getRange('E:E').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('F:F').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('G:G').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('H:H').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('I:I').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('L:L').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('M:M').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('N:N').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('O:O').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('P:P').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('Q:Q').setNumberFormat('MM/dd/yyyy');
  sheet.getRange('S:S').setNumberFormat('MM/dd/yyyy');

  // Hide the sheet
  sheet.hideSheet();

  Logger.log('_Grievance_Formulas sheet setup complete');
}

/**
 * Sync calculated formulas from hidden sheet to Grievance Log
 * This is the self-healing function - it copies calculated values to the Grievance Log
 * Member data (Name, Email, Unit, Location, Steward) is looked up directly from Member Directory
 */
function syncGrievanceFormulasToLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!grievanceSheet || !memberSheet) {
    Logger.log('Required sheets not found for grievance formula sync');
    return;
  }

  // Get Member Directory data and create lookup by Member ID
  var memberData = memberSheet.getDataRange().getValues();
  var memberLookup = {};
  for (var i = 1; i < memberData.length; i++) {
    var memberId = memberData[i][MEMBER_COLS.MEMBER_ID - 1];
    if (memberId) {
      memberLookup[memberId] = {
        firstName: memberData[i][MEMBER_COLS.FIRST_NAME - 1] || '',
        lastName: memberData[i][MEMBER_COLS.LAST_NAME - 1] || '',
        email: memberData[i][MEMBER_COLS.EMAIL - 1] || '',
        unit: memberData[i][MEMBER_COLS.UNIT - 1] || '',
        location: memberData[i][MEMBER_COLS.WORK_LOCATION - 1] || '',
        steward: memberData[i][MEMBER_COLS.ASSIGNED_STEWARD - 1] || ''
      };
    }
  }

  // Get grievance data
  var grievanceData = grievanceSheet.getDataRange().getValues();
  if (grievanceData.length < 2) return;

  var today = new Date();
  today.setHours(0, 0, 0, 0); // Normalize to start of day

  // Closed statuses that should not have Next Action Due
  var closedStatuses = ['Settled', 'Withdrawn', 'Denied', 'Won', 'Closed'];

  // Prepare updates
  var nameUpdates = [];           // Columns C-D
  var deadlineUpdates = [];       // Columns H, J, L, N, P (Filing Deadline, Step I Due, Step II Appeal Due, Step II Due, Step III Appeal Due)
  var metricsUpdates = [];        // Columns S, T, U (Days Open, Next Action Due, Days to Deadline)
  var contactUpdates = [];        // Columns X, Y, Z, AA (Email, Unit, Location, Steward)

  // Track data quality issues
  var orphanedGrievances = [];    // Grievances with non-existent Member IDs
  var missingMemberIds = [];      // Grievances with no Member ID

  for (var j = 1; j < grievanceData.length; j++) {
    var row = grievanceData[j];
    var memberId = row[GRIEVANCE_COLS.MEMBER_ID - 1];
    var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1] || ('Row ' + (j + 1));

    // Track data quality issues
    if (!memberId) {
      missingMemberIds.push(grievanceId);
      Logger.log('WARNING: Grievance ' + grievanceId + ' has no Member ID');
    } else if (!memberLookup[memberId]) {
      orphanedGrievances.push(grievanceId + ' (Member ID: ' + memberId + ')');
      Logger.log('WARNING: Grievance ' + grievanceId + ' references non-existent Member ID: ' + memberId);
    }

    var memberInfo = memberLookup[memberId] || {};

    // Names (C-D) - from Member Directory
    nameUpdates.push([
      memberInfo.firstName || '',
      memberInfo.lastName || ''
    ]);

    // Get date values from grievance row for deadline calculations
    var incidentDate = row[GRIEVANCE_COLS.INCIDENT_DATE - 1];
    var dateFiled = row[GRIEVANCE_COLS.DATE_FILED - 1];
    var step1Rcvd = row[GRIEVANCE_COLS.STEP1_RCVD - 1];
    var step2AppealFiled = row[GRIEVANCE_COLS.STEP2_APPEAL_FILED - 1];
    var step2Rcvd = row[GRIEVANCE_COLS.STEP2_RCVD - 1];
    var dateClosed = row[GRIEVANCE_COLS.DATE_CLOSED - 1];
    var status = row[GRIEVANCE_COLS.STATUS - 1];
    var currentStep = row[GRIEVANCE_COLS.CURRENT_STEP - 1];

    // Calculate deadline dates
    var filingDeadline = '';
    var step1Due = '';
    var step2AppealDue = '';
    var step2Due = '';
    var step3AppealDue = '';

    if (incidentDate instanceof Date) {
      filingDeadline = new Date(incidentDate.getTime() + 21 * 24 * 60 * 60 * 1000);
    }
    if (dateFiled instanceof Date) {
      step1Due = new Date(dateFiled.getTime() + 30 * 24 * 60 * 60 * 1000);
    }
    if (step1Rcvd instanceof Date) {
      step2AppealDue = new Date(step1Rcvd.getTime() + 10 * 24 * 60 * 60 * 1000);
    }
    if (step2AppealFiled instanceof Date) {
      step2Due = new Date(step2AppealFiled.getTime() + 30 * 24 * 60 * 60 * 1000);
    }
    if (step2Rcvd instanceof Date) {
      step3AppealDue = new Date(step2Rcvd.getTime() + 30 * 24 * 60 * 60 * 1000);
    }

    // Deadlines (H, J, L, N, P)
    deadlineUpdates.push([
      filingDeadline,
      step1Due,
      step2AppealDue,
      step2Due,
      step3AppealDue
    ]);

    // Calculate Days Open directly
    var daysOpen = '';
    if (dateFiled instanceof Date) {
      if (dateClosed instanceof Date) {
        daysOpen = Math.floor((dateClosed - dateFiled) / (1000 * 60 * 60 * 24));
      } else {
        daysOpen = Math.floor((today - dateFiled) / (1000 * 60 * 60 * 24));
      }
    }

    // Calculate Next Action Due based on current step and status
    var nextActionDue = '';
    var isClosed = closedStatuses.indexOf(status) !== -1;

    if (!isClosed && currentStep) {
      if (currentStep === 'Informal' && filingDeadline) {
        nextActionDue = filingDeadline;
      } else if (currentStep === 'Step I' && step1Due) {
        nextActionDue = step1Due;
      } else if (currentStep === 'Step II' && step2Due) {
        nextActionDue = step2Due;
      } else if (currentStep === 'Step III' && step3AppealDue) {
        nextActionDue = step3AppealDue;
      }
    }

    // Calculate Days to Deadline directly
    var daysToDeadline = '';
    if (nextActionDue instanceof Date) {
      var days = Math.floor((nextActionDue - today) / (1000 * 60 * 60 * 24));
      daysToDeadline = days < 0 ? 'Overdue' : days;
    }

    // Metrics (S, T, U)
    metricsUpdates.push([
      daysOpen,
      nextActionDue,
      daysToDeadline
    ]);

    // Contact info (X, Y, Z, AA)
    contactUpdates.push([
      memberInfo.email || '',
      memberInfo.unit || '',
      memberInfo.location || '',
      memberInfo.steward || ''
    ]);
  }

  // Apply updates to Grievance Log
  if (nameUpdates.length > 0) {
    // C-D: First Name, Last Name
    grievanceSheet.getRange(2, GRIEVANCE_COLS.FIRST_NAME, nameUpdates.length, 2).setValues(nameUpdates);

    // H: Filing Deadline (column 8)
    grievanceSheet.getRange(2, GRIEVANCE_COLS.FILING_DEADLINE, deadlineUpdates.length, 1)
      .setValues(deadlineUpdates.map(function(r) { return [r[0]]; }));

    // J: Step I Due (column 10)
    grievanceSheet.getRange(2, GRIEVANCE_COLS.STEP1_DUE, deadlineUpdates.length, 1)
      .setValues(deadlineUpdates.map(function(r) { return [r[1]]; }));

    // L: Step II Appeal Due (column 12)
    grievanceSheet.getRange(2, GRIEVANCE_COLS.STEP2_APPEAL_DUE, deadlineUpdates.length, 1)
      .setValues(deadlineUpdates.map(function(r) { return [r[2]]; }));

    // N: Step II Due (column 14)
    grievanceSheet.getRange(2, GRIEVANCE_COLS.STEP2_DUE, deadlineUpdates.length, 1)
      .setValues(deadlineUpdates.map(function(r) { return [r[3]]; }));

    // P: Step III Appeal Due (column 16)
    grievanceSheet.getRange(2, GRIEVANCE_COLS.STEP3_APPEAL_DUE, deadlineUpdates.length, 1)
      .setValues(deadlineUpdates.map(function(r) { return [r[4]]; }));

    // Format deadline columns as dates (MM/dd/yyyy)
    grievanceSheet.getRange(2, GRIEVANCE_COLS.FILING_DEADLINE, deadlineUpdates.length, 1).setNumberFormat('MM/dd/yyyy');
    grievanceSheet.getRange(2, GRIEVANCE_COLS.STEP1_DUE, deadlineUpdates.length, 1).setNumberFormat('MM/dd/yyyy');
    grievanceSheet.getRange(2, GRIEVANCE_COLS.STEP2_APPEAL_DUE, deadlineUpdates.length, 1).setNumberFormat('MM/dd/yyyy');
    grievanceSheet.getRange(2, GRIEVANCE_COLS.STEP2_DUE, deadlineUpdates.length, 1).setNumberFormat('MM/dd/yyyy');
    grievanceSheet.getRange(2, GRIEVANCE_COLS.STEP3_APPEAL_DUE, deadlineUpdates.length, 1).setNumberFormat('MM/dd/yyyy');

    // S, T, U: Days Open, Next Action Due, Days to Deadline
    grievanceSheet.getRange(2, GRIEVANCE_COLS.DAYS_OPEN, metricsUpdates.length, 3).setValues(metricsUpdates);

    // Format Days Open (S) and Days to Deadline (U) as whole numbers
    grievanceSheet.getRange(2, GRIEVANCE_COLS.DAYS_OPEN, metricsUpdates.length, 1).setNumberFormat('0');
    grievanceSheet.getRange(2, GRIEVANCE_COLS.NEXT_ACTION_DUE, metricsUpdates.length, 1).setNumberFormat('MM/dd/yyyy');
    grievanceSheet.getRange(2, GRIEVANCE_COLS.DAYS_TO_DEADLINE, metricsUpdates.length, 1).setNumberFormat('0');

    // X, Y, Z, AA: Email, Unit, Location, Steward
    grievanceSheet.getRange(2, GRIEVANCE_COLS.MEMBER_EMAIL, contactUpdates.length, 4).setValues(contactUpdates);
  }

  Logger.log('Synced grievance formulas to ' + nameUpdates.length + ' grievances');

  // Show warnings to user if data quality issues found
  var warnings = [];
  if (missingMemberIds.length > 0) {
    warnings.push(missingMemberIds.length + ' grievance(s) have no Member ID');
    Logger.log('Missing Member IDs: ' + missingMemberIds.join(', '));
  }
  if (orphanedGrievances.length > 0) {
    warnings.push(orphanedGrievances.length + ' grievance(s) reference non-existent members');
    Logger.log('Orphaned grievances: ' + orphanedGrievances.join(', '));
  }

  if (warnings.length > 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '⚠️ Data issues found:\n' + warnings.join('\n') + '\n\nCheck Logs for details.',
      '⚠️ Sync Warning',
      10
    );
  }
}

/**
 * Auto-sort the Grievance Log by status priority
 * Active cases (Open, Pending Info, In Arbitration, Appealed) appear first,
 * resolved cases (Settled, Won, Denied, Withdrawn, Closed) appear last
 */
function sortGrievanceLogByStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return; // Need at least 2 data rows to sort

  // Get all data (excluding header row)
  var dataRange = sheet.getRange(2, 1, lastRow - 1, 34);
  var data = dataRange.getValues();

  // Sort by status priority (column E = index 4)
  data.sort(function(a, b) {
    var statusA = a[GRIEVANCE_COLS.STATUS - 1] || '';
    var statusB = b[GRIEVANCE_COLS.STATUS - 1] || '';

    // Get priority (default to 99 for unknown statuses)
    var priorityA = GRIEVANCE_STATUS_PRIORITY[statusA] || 99;
    var priorityB = GRIEVANCE_STATUS_PRIORITY[statusB] || 99;

    // Primary sort: by status priority (lower number = higher priority)
    if (priorityA !== priorityB) {
      return priorityA - priorityB;
    }

    // Secondary sort: by Days to Deadline (column U = index 20) - most urgent first
    var daysA = a[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];
    var daysB = b[GRIEVANCE_COLS.DAYS_TO_DEADLINE - 1];

    // Handle empty/non-numeric values
    if (daysA === '' || daysA === null) daysA = 9999;
    if (daysB === '' || daysB === null) daysB = 9999;

    return daysA - daysB;
  });

  // Write sorted data back
  dataRange.setValues(data);

  // Re-apply checkboxes to Message Alert column (AC) - setValues overwrites them
  if (lastRow >= 2) {
    sheet.getRange(2, GRIEVANCE_COLS.MESSAGE_ALERT, lastRow - 1, 1).insertCheckboxes();
  }

  Logger.log('Grievance Log sorted by status priority');
}

// ============================================================================
// HIDDEN SHEET 3: _Member_Lookup
// Source: Member Directory → Destination: Grievance Log (C,D,X-AA)
// ============================================================================

/**
 * Setup the _Member_Lookup hidden sheet with self-healing formulas
 * Looks up: First Name, Last Name, Email, Unit, Location, Steward from Member Directory
 */
function setupMemberLookupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MEMBER_LOOKUP);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.MEMBER_LOOKUP);
  }

  sheet.clear();

  // Headers
  var headers = ['Member ID', 'First Name', 'Last Name', 'Email', 'Unit', 'Location', 'Assigned Steward'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  // Get column letters
  var mIdCol = getColumnLetter(MEMBER_COLS.MEMBER_ID);
  var mFirstCol = getColumnLetter(MEMBER_COLS.FIRST_NAME);
  var mLastCol = getColumnLetter(MEMBER_COLS.LAST_NAME);
  var mEmailCol = getColumnLetter(MEMBER_COLS.EMAIL);
  var mUnitCol = getColumnLetter(MEMBER_COLS.UNIT);
  var mLocCol = getColumnLetter(MEMBER_COLS.WORK_LOCATION);
  var mStewardCol = getColumnLetter(MEMBER_COLS.ASSIGNED_STEWARD);

  // Formula to get unique member IDs from Grievance Log
  var gMemberIdCol = getColumnLetter(GRIEVANCE_COLS.MEMBER_ID);
  var memberIdFormula = '=IFERROR(UNIQUE(FILTER(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + ',\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + '<>"Member ID",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gMemberIdCol + ':' + gMemberIdCol + '<>"")),"")';
  sheet.getRange('A2').setFormula(memberIdFormula);

  // VLOOKUP formulas for member data
  var vlookupBase = 'VLOOKUP(A2:A,\'' + SHEETS.MEMBER_DIR + '\'!' + mIdCol + ':' + mStewardCol + ',';

  sheet.getRange('B2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(' + vlookupBase + '2,FALSE),"")))'); // First Name
  sheet.getRange('C2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(' + vlookupBase + '3,FALSE),"")))'); // Last Name
  sheet.getRange('D2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(' + vlookupBase + '8,FALSE),"")))'); // Email
  sheet.getRange('E2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(' + vlookupBase + '6,FALSE),"")))'); // Unit
  sheet.getRange('F2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(' + vlookupBase + '5,FALSE),"")))'); // Location
  sheet.getRange('G2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(' + vlookupBase + '16,FALSE),"")))'); // Steward

  // Hide the sheet
  sheet.hideSheet();

  Logger.log('_Member_Lookup sheet setup complete');
}

/**
 * Sync member data from hidden sheet to Grievance Log
 */
function syncMemberToGrievanceLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lookupSheet = ss.getSheetByName(SHEETS.MEMBER_LOOKUP);
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!lookupSheet || !grievanceSheet) {
    Logger.log('Required sheets not found for member sync');
    return;
  }

  // Get lookup data
  var lookupData = lookupSheet.getDataRange().getValues();
  if (lookupData.length < 2) return;

  // Create lookup map
  var lookup = {};
  for (var i = 1; i < lookupData.length; i++) {
    var memberId = lookupData[i][0];
    if (memberId) {
      lookup[memberId] = {
        firstName: lookupData[i][1],
        lastName: lookupData[i][2],
        email: lookupData[i][3],
        unit: lookupData[i][4],
        location: lookupData[i][5],
        steward: lookupData[i][6]
      };
    }
  }

  // Get grievance data
  var grievanceData = grievanceSheet.getDataRange().getValues();
  if (grievanceData.length < 2) return;

  // Update grievance rows
  var nameUpdates = [];
  var infoUpdates = [];

  for (var j = 1; j < grievanceData.length; j++) {
    var memberId = grievanceData[j][GRIEVANCE_COLS.MEMBER_ID - 1];
    var data = lookup[memberId] || {firstName: '', lastName: '', email: '', unit: '', location: '', steward: ''};
    nameUpdates.push([data.firstName, data.lastName]);
    infoUpdates.push([data.email, data.unit, data.location, data.steward]);
  }

  if (nameUpdates.length > 0) {
    // Update C-D (First Name, Last Name)
    grievanceSheet.getRange(2, GRIEVANCE_COLS.FIRST_NAME, nameUpdates.length, 2).setValues(nameUpdates);
    // Update X-AA (Email, Unit, Location, Steward)
    grievanceSheet.getRange(2, GRIEVANCE_COLS.MEMBER_EMAIL, infoUpdates.length, 4).setValues(infoUpdates);
  }

  Logger.log('Synced member data to ' + nameUpdates.length + ' grievances');
}

// ============================================================================
// HIDDEN SHEET 4: _Steward_Contact_Calc
// Source: Member Directory (Y-AA) → Aggregates steward contact tracking metrics
// ============================================================================

/**
 * Setup the _Steward_Contact_Calc hidden sheet with self-healing formulas
 * Tracks and aggregates steward contact data from Member Directory
 */
function setupStewardContactCalcSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.STEWARD_CONTACT_CALC);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.STEWARD_CONTACT_CALC);
  }

  sheet.clear();

  // Headers for steward contact summary
  var headers = ['Steward Name', 'Total Contacts', 'Contacts This Month', 'Contacts Last 7 Days', 'Last Contact Date', 'Avg Days Between Contacts'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  // Get column letters for formulas
  var mContactStewardCol = getColumnLetter(MEMBER_COLS.CONTACT_STEWARD);
  var mContactDateCol = getColumnLetter(MEMBER_COLS.RECENT_CONTACT_DATE);

  // Column A: Unique steward names who have made contacts
  sheet.getRange('A2').setFormula('=IFERROR(SORT(UNIQUE(FILTER(\'' + SHEETS.MEMBER_DIR + '\'!' + mContactStewardCol + ':' + mContactStewardCol + ',\'' + SHEETS.MEMBER_DIR + '\'!' + mContactStewardCol + ':' + mContactStewardCol + '<>""))),)');

  // Column B: Total contacts per steward
  sheet.getRange('B2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIF(\'' + SHEETS.MEMBER_DIR + '\'!' + mContactStewardCol + ':' + mContactStewardCol + ',A2:A)))');

  // Column C: Contacts this month
  sheet.getRange('C2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(\'' + SHEETS.MEMBER_DIR + '\'!' + mContactStewardCol + ':' + mContactStewardCol + ',A2:A,\'' + SHEETS.MEMBER_DIR + '\'!' + mContactDateCol + ':' + mContactDateCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1))))');

  // Column D: Contacts last 7 days
  sheet.getRange('D2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(\'' + SHEETS.MEMBER_DIR + '\'!' + mContactStewardCol + ':' + mContactStewardCol + ',A2:A,\'' + SHEETS.MEMBER_DIR + '\'!' + mContactDateCol + ':' + mContactDateCol + ',">="&(TODAY()-7))))');

  // Column E: Most recent contact date for this steward
  sheet.getRange('E2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(TEXT(MAXIFS(\'' + SHEETS.MEMBER_DIR + '\'!' + mContactDateCol + ':' + mContactDateCol + ',\'' + SHEETS.MEMBER_DIR + '\'!' + mContactStewardCol + ':' + mContactStewardCol + ',A2:A),"MM/dd/yyyy"),"-")))');

  // Column F: Placeholder for avg days between contacts (complex calculation)
  sheet.getRange('F2').setFormula('=ARRAYFORMULA(IF(A2:A="","","-"))');

  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);

  sheet.hideSheet();
  Logger.log('_Steward_Contact_Calc sheet setup complete with live formulas');
}

// ============================================================================
// HIDDEN SHEET 6: _Steward_Performance_Calc
// Source: Grievance Log → Steward Performance Metrics
// ============================================================================

/**
 * Setup the _Steward_Performance_Calc hidden sheet
 * Calculates detailed steward performance metrics
 */
function setupStewardPerformanceCalcSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.STEWARD_PERFORMANCE_CALC);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.STEWARD_PERFORMANCE_CALC);
  }

  sheet.clear();

  // Headers
  var headers = ['Steward', 'Total Cases', 'Active', 'Closed', 'Won', 'Win Rate %', 'Avg Days', 'Overdue', 'Due This Week', 'Performance Score'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  var gStewardCol = getColumnLetter(GRIEVANCE_COLS.STEWARD);
  var gStatusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);
  var gResolutionCol = getColumnLetter(GRIEVANCE_COLS.RESOLUTION);
  var gDaysOpenCol = getColumnLetter(GRIEVANCE_COLS.DAYS_OPEN);
  var gDaysToDeadlineCol = getColumnLetter(GRIEVANCE_COLS.DAYS_TO_DEADLINE);

  // Get unique stewards
  sheet.getRange('A2').setFormula(
    '=IFERROR(UNIQUE(FILTER(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',' +
    '\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + '<>"Assigned Steward",' +
    '\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + '<>"")),"")'
  );

  // Total Cases
  sheet.getRange('B2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',A2:A)))');

  // Active Cases (Open + Pending Info)
  sheet.getRange('C2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Open")+COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Pending Info")))');

  // Closed Cases
  sheet.getRange('D2').setFormula('=ARRAYFORMULA(IF(A2:A="","",B2:B-C2:C))');

  // Won Cases
  sheet.getRange('E2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gResolutionCol + ':' + gResolutionCol + ',"*Won*")))');

  // Win Rate
  sheet.getRange('F2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(ROUND(E2:E/D2:D*100,1),0)))');

  // Avg Days
  sheet.getRange('G2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(ROUND(AVERAGEIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysOpenCol + ':' + gDaysOpenCol + '),1),0)))');

  // Overdue
  sheet.getRange('H2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysToDeadlineCol + ':' + gDaysToDeadlineCol + ',"Overdue")))');

  // Due This Week
  sheet.getRange('I2').setFormula('=ARRAYFORMULA(IF(A2:A="","",COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStewardCol + ':' + gStewardCol + ',A2:A,\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysToDeadlineCol + ':' + gDaysToDeadlineCol + ',">=0",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysToDeadlineCol + ':' + gDaysToDeadlineCol + ',"<=7")))');

  // Performance Score (weighted: Win Rate * 0.4 + (100 - Overdue%) * 0.3 + (100 - AvgDays/60*100) * 0.3)
  sheet.getRange('J2').setFormula('=ARRAYFORMULA(IF(A2:A="","",ROUND(F2:F*0.4 + (100-IFERROR(H2:H/C2:C*100,0))*0.3 + MAX(0,100-G2:G/60*100)*0.3,1)))');

  sheet.hideSheet();
  Logger.log('_Steward_Performance_Calc sheet setup complete');
}

/**
 * Sync steward performance to visible sheet
 */
function syncStewardPerformance() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calcSheet = ss.getSheetByName(SHEETS.STEWARD_PERFORMANCE_CALC);

  if (!calcSheet) return;

  var data = calcSheet.getDataRange().getValues();
  Logger.log('Steward performance data available: ' + (data.length - 1) + ' stewards');
}

// ============================================================================
// AUTO-SYNC TRIGGERS
// ============================================================================

/**
 * Sync new values from Member Directory to Config (bidirectional sync)
 * When a user enters a new value in a job metadata field, add it to Config
 * @param {Object} e - The edit event object
 */
function syncNewValueToConfig(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  if (sheet.getName() !== SHEETS.MEMBER_DIR) return;

  var col = e.range.getColumn();
  var newValue = e.range.getValue();

  // Skip if empty or header row
  if (!newValue || e.range.getRow() === 1) return;

  // Check if this column is a job metadata field (includes Committees and Home Town)
  var fieldConfig = getJobMetadataByMemberCol(col);
  if (!fieldConfig) return; // Not a synced column

  // Get current Config values for this column
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!configSheet) return;

  var existingValues = getConfigValues(configSheet, fieldConfig.configCol);

  // Handle multi-value fields (comma-separated)
  var valuesToCheck = newValue.toString().split(',').map(function(v) { return v.trim(); });

  var valuesToAdd = [];
  for (var j = 0; j < valuesToCheck.length; j++) {
    var val = valuesToCheck[j];
    if (val && existingValues.indexOf(val) === -1) {
      valuesToAdd.push(val);
    }
  }

  // Add new values to Config
  if (valuesToAdd.length > 0) {
    var lastRow = configSheet.getLastRow();
    var dataStartRow = Math.max(lastRow + 1, 3); // Start at row 3 minimum

    for (var k = 0; k < valuesToAdd.length; k++) {
      configSheet.getRange(dataStartRow + k, fieldConfig.configCol).setValue(valuesToAdd[k]);
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Added "' + valuesToAdd.join(', ') + '" to ' + fieldConfig.configName,
      '🔄 Config Updated', 3
    );
  }
}

/**
 * Master onEdit trigger - routes to appropriate sync function
 * Install this as an installable trigger
 */
function onEditAutoSync(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();

  // Debounce - use cache to prevent rapid re-syncs
  var cache = CacheService.getScriptCache();
  var cacheKey = 'lastSync_' + sheetName;
  var lastSync = cache.get(cacheKey);

  if (lastSync) {
    return; // Skip if synced within last 2 seconds
  }

  cache.put(cacheKey, 'true', 2); // 2 second debounce

  try {
    if (sheetName === SHEETS.GRIEVANCE_LOG) {
      // Grievance Log changed - sync formulas and update Member Directory
      syncGrievanceFormulasToLog();
      syncGrievanceToMemberDirectory();
      // Auto-sort by status priority (active cases first, then by deadline urgency)
      sortGrievanceLogByStatus();
    } else if (sheetName === SHEETS.MEMBER_DIR) {
      // Member Directory changed - sync to Grievance Log and Config
      syncNewValueToConfig(e);  // Bidirectional: add new values to Config
      syncGrievanceFormulasToLog();
      syncMemberToGrievanceLog();
    }
  } catch (error) {
    Logger.log('Auto-sync error: ' + error.message);
  }
}

/**
 * Install the auto-sync trigger with options dialog
 * Users can customize the sync behavior
 */
function installAutoSyncTrigger() {
  var ui = SpreadsheetApp.getUi();
  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body{font-family:Arial;padding:20px;background:#f5f5f5}' +
    '.container{background:white;padding:25px;border-radius:8px}' +
    'h2{color:#1a73e8;margin-top:0}' +
    '.section{background:#f8f9fa;padding:15px;margin:15px 0;border-radius:8px}' +
    '.section h4{margin:0 0 10px;color:#333}' +
    '.option{display:flex;align-items:center;margin:8px 0}' +
    '.option input[type="checkbox"]{margin-right:10px}' +
    '.option label{font-size:14px}' +
    '.info{background:#e8f4fd;padding:12px;border-radius:8px;font-size:13px;margin-bottom:15px}' +
    '.buttons{display:flex;gap:10px;margin-top:20px}' +
    'button{padding:12px 20px;border:none;border-radius:4px;cursor:pointer}' +
    '.primary{background:#1a73e8;color:white;flex:1}' +
    '.secondary{background:#e0e0e0;flex:1}' +
    '.warning{background:#fff3cd;padding:10px;border-radius:4px;font-size:12px;color:#856404}' +
    '</style></head><body><div class="container">' +
    '<h2>⚡ Auto-Sync Settings</h2>' +
    '<div class="info">Auto-sync automatically updates cross-sheet data when you edit cells in Member Directory or Grievance Log.</div>' +

    '<div class="section"><h4>Sync Options</h4>' +
    '<div class="option"><input type="checkbox" id="syncGrievances" checked><label>Sync Grievance data to Member Directory</label></div>' +
    '<div class="option"><input type="checkbox" id="syncMembers" checked><label>Sync Member data to Grievance Log</label></div>' +
    '<div class="option"><input type="checkbox" id="autoSort" checked><label>Auto-sort Grievance Log by status/deadline</label></div>' +
    '<div class="option"><input type="checkbox" id="repairCheckboxes" checked><label>Auto-repair checkboxes after sync</label></div>' +
    '</div>' +

    '<div class="section"><h4>Performance</h4>' +
    '<div class="option"><input type="checkbox" id="showToasts" checked><label>Show sync notifications (toasts)</label></div>' +
    '<div class="warning">💡 Disabling notifications improves performance but you won\'t see sync status.</div>' +
    '</div>' +

    '<div class="buttons">' +
    '<button class="secondary" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="primary" onclick="install()">Install Trigger</button>' +
    '</div></div>' +
    '<script>' +
    'function install(){' +
    'var opts={syncGrievances:document.getElementById("syncGrievances").checked,syncMembers:document.getElementById("syncMembers").checked,autoSort:document.getElementById("autoSort").checked,repairCheckboxes:document.getElementById("repairCheckboxes").checked,showToasts:document.getElementById("showToasts").checked};' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close()}).installAutoSyncTriggerWithOptions(opts)}' +
    '</script></body></html>'
  ).setWidth(450).setHeight(480);
  ui.showModalDialog(html, '⚡ Auto-Sync Settings');
}

/**
 * Install auto-sync trigger with saved options
 * @param {Object} options - Sync configuration options
 */
function installAutoSyncTriggerWithOptions(options) {
  // Save options to script properties
  var props = PropertiesService.getScriptProperties();
  props.setProperty('autoSyncOptions', JSON.stringify(options));

  // Remove existing triggers first
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onEditAutoSync') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Install new trigger
  ScriptApp.newTrigger('onEditAutoSync')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  Logger.log('Auto-sync trigger installed with options: ' + JSON.stringify(options));
  SpreadsheetApp.getActiveSpreadsheet().toast('Auto-sync trigger installed!', '✅ Success', 3);
}

/**
 * Get auto-sync options (with defaults)
 */
function getAutoSyncOptions() {
  var props = PropertiesService.getScriptProperties();
  var optionsJSON = props.getProperty('autoSyncOptions');
  if (optionsJSON) {
    return JSON.parse(optionsJSON);
  }
  // Default options
  return {
    syncGrievances: true,
    syncMembers: true,
    autoSort: true,
    repairCheckboxes: true,
    showToasts: true
  };
}

/**
 * Quick install (no dialog) - used by repair functions
 */
function installAutoSyncTriggerQuick() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onEditAutoSync') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('onEditAutoSync')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  Logger.log('Auto-sync trigger installed (quick mode)');
}

/**
 * Remove the auto-sync trigger
 */
function removeAutoSyncTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;

  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onEditAutoSync') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });

  Logger.log('Removed ' + removed + ' auto-sync triggers');
  SpreadsheetApp.getActiveSpreadsheet().toast('Auto-sync trigger removed', 'Info', 3);
}

// ============================================================================
// HIDDEN SHEET 5: _Dashboard_Calc
// Source: Member Directory + Grievance Log → Dashboard Summary Statistics
// ============================================================================

/**
 * Setup the _Dashboard_Calc hidden sheet with self-healing formulas
 * Calculates key dashboard metrics that auto-update
 */
function setupDashboardCalcSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.DASHBOARD_CALC);

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.DASHBOARD_CALC);
  }

  sheet.clear();

  // Headers
  var headers = ['Metric', 'Value', 'Description'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.LIGHT_GRAY);

  // Column references
  var mIdCol = getColumnLetter(MEMBER_COLS.MEMBER_ID);
  var mStewardCol = getColumnLetter(MEMBER_COLS.IS_STEWARD);
  var gIdCol = getColumnLetter(GRIEVANCE_COLS.GRIEVANCE_ID);
  var gStatusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);
  var gResolutionCol = getColumnLetter(GRIEVANCE_COLS.RESOLUTION);
  var gDaysOpenCol = getColumnLetter(GRIEVANCE_COLS.DAYS_OPEN);
  var gDaysToDeadlineCol = getColumnLetter(GRIEVANCE_COLS.DAYS_TO_DEADLINE);
  var gDateFiledCol = getColumnLetter(GRIEVANCE_COLS.DATE_FILED);
  var gDateClosedCol = getColumnLetter(GRIEVANCE_COLS.DATE_CLOSED);

  // Metrics with formulas (15 key metrics)
  var metrics = [
    ['Total Members', '=COUNTA(\'' + SHEETS.MEMBER_DIR + '\'!' + mIdCol + ':' + mIdCol + ')-1', 'Total union members in directory'],
    ['Active Stewards', '=COUNTIF(\'' + SHEETS.MEMBER_DIR + '\'!' + mStewardCol + ':' + mStewardCol + ',"Yes")', 'Members marked as stewards'],
    ['Total Grievances', '=COUNTA(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gIdCol + ':' + gIdCol + ')-1', 'All grievances filed'],
    ['Open Grievances', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Open")', 'Currently open cases'],
    ['Pending Info', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Pending Info")', 'Cases awaiting information'],
    ['Settled', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Settled")', 'Cases settled'],
    ['Won', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gResolutionCol + ':' + gResolutionCol + ',"*Won*")', 'Cases won (full or partial)'],
    ['Denied', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Denied")', 'Cases denied'],
    ['Withdrawn', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Withdrawn")', 'Cases withdrawn'],
    ['Win Rate %', '=IFERROR(ROUND(COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gResolutionCol + ':' + gResolutionCol + ',"*Won*")/(COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Settled")+COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gStatusCol + ':' + gStatusCol + ',"Denied")+COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gResolutionCol + ':' + gResolutionCol + ',"*Won*"))*100,1),0)', 'Wins / (Wins + Settled + Denied)'],
    ['Avg Days to Resolution', '=IFERROR(ROUND(AVERAGEIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',"<>",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysOpenCol + ':' + gDaysOpenCol + '),1),0)', 'Average days for closed cases'],
    ['Overdue Cases', '=COUNTIF(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysToDeadlineCol + ':' + gDaysToDeadlineCol + ',"Overdue")', 'Cases past deadline'],
    ['Due This Week', '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysToDeadlineCol + ':' + gDaysToDeadlineCol + ',">=0",\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDaysToDeadlineCol + ':' + gDaysToDeadlineCol + ',"<=7")', 'Cases due in next 7 days'],
    ['Filed This Month', '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateFiledCol + ':' + gDateFiledCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateFiledCol + ':' + gDateFiledCol + ',"<="&TODAY())', 'Grievances filed this month'],
    ['Closed This Month', '=COUNTIFS(\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),\'' + SHEETS.GRIEVANCE_LOG + '\'!' + gDateClosedCol + ':' + gDateClosedCol + ',"<="&TODAY())', 'Grievances closed this month']
  ];

  for (var i = 0; i < metrics.length; i++) {
    sheet.getRange(i + 2, 1).setValue(metrics[i][0]);
    sheet.getRange(i + 2, 2).setFormula(metrics[i][1]);
    sheet.getRange(i + 2, 3).setValue(metrics[i][2]);
  }

  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 300);

  sheet.hideSheet();
  Logger.log('_Dashboard_Calc sheet setup complete');
}

// ============================================================================
// MASTER SETUP & REPAIR FUNCTIONS
// ============================================================================

/**
 * Setup all hidden calculation sheets
 */
function setupAllHiddenSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Setting up hidden calculation sheets...', '🔧 Setup', 3);

  // Core grievance/member calculation sheets (6 total)
  setupGrievanceCalcSheet();
  setupGrievanceFormulasSheet();
  setupMemberLookupSheet();
  setupStewardContactCalcSheet();
  setupDashboardCalcSheet();
  setupStewardPerformanceCalcSheet();

  ss.toast('All 6 hidden sheets created!', '✅ Success', 3);
}

/**
 * Repair all hidden sheets - recreates formulas and syncs data
 */
function repairAllHiddenSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  ss.toast('Repairing hidden sheets...', '🔧 Repair', 3);

  // Recreate all hidden sheets with formulas
  setupAllHiddenSheets();

  // Install trigger (quick mode - no dialog)
  installAutoSyncTriggerQuick();

  // Run initial sync
  ss.toast('Running initial data sync...', '🔧 Sync', 3);
  syncGrievanceFormulasToLog();
  syncGrievanceToMemberDirectory();
  syncMemberToGrievanceLog();

  // Repair checkboxes
  repairGrievanceCheckboxes();
  repairMemberCheckboxes();

  ss.toast('Hidden sheets repaired and synced!', '✅ Success', 5);
  ui.alert('✅ Repair Complete',
    'Hidden calculation sheets have been repaired:\n\n' +
    '• 5 hidden sheets recreated with self-healing formulas\n' +
    '• Auto-sync trigger installed\n' +
    '• All data synced (grievances, members, dashboard)\n' +
    '• Checkboxes repaired in Grievance Log and Member Directory\n\n' +
    'Data will now auto-sync when you edit Member Directory or Grievance Log.\n' +
    'Formulas cannot be accidentally erased - they are stored in hidden sheets.',
    ui.ButtonSet.OK);
}

/**
 * Verify all hidden sheets and triggers
 */
function verifyHiddenSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var report = [];

  report.push('🔍 HIDDEN SHEET VERIFICATION');
  report.push('============================');
  report.push('');

  // Check each hidden sheet (5 hidden sheets)
  var hiddenSheets = [
    {name: SHEETS.GRIEVANCE_CALC, purpose: 'Grievance → Member Directory'},
    {name: SHEETS.GRIEVANCE_FORMULAS, purpose: 'Self-healing Grievance formulas'},
    {name: SHEETS.MEMBER_LOOKUP, purpose: 'Member → Grievance Log'},
    {name: SHEETS.STEWARD_CONTACT_CALC, purpose: 'Steward contact tracking'},
    {name: SHEETS.DASHBOARD_CALC, purpose: 'Dashboard summary metrics'}
  ];

  report.push('📋 HIDDEN SHEETS:');
  hiddenSheets.forEach(function(hs) {
    var sheet = ss.getSheetByName(hs.name);
    if (sheet) {
      var isHidden = sheet.isSheetHidden();
      var hasData = sheet.getLastRow() > 1;
      var status = isHidden && hasData ? '✅' : (sheet ? '⚠️' : '❌');
      report.push('  ' + status + ' ' + hs.name);
      report.push('      Hidden: ' + (isHidden ? 'Yes' : 'NO - Should be hidden'));
      report.push('      Has formulas: ' + (hasData ? 'Yes' : 'No'));
    } else {
      report.push('  ❌ ' + hs.name + ' - NOT FOUND');
    }
  });

  report.push('');

  // Check triggers
  report.push('⚡ AUTO-SYNC TRIGGER:');
  var triggers = ScriptApp.getProjectTriggers();
  var hasAutoSync = false;
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onEditAutoSync') {
      hasAutoSync = true;
      report.push('  ✅ onEditAutoSync trigger installed');
    }
  });
  if (!hasAutoSync) {
    report.push('  ❌ onEditAutoSync trigger NOT installed');
    report.push('     Run: installAutoSyncTrigger()');
  }

  report.push('');
  report.push('============================');

  ui.alert('Hidden Sheet Verification', report.join('\n'), ui.ButtonSet.OK);
  Logger.log(report.join('\n'));
}

/**
 * Manual sync all data with data quality validation
 */
function syncAllData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ss.toast('Syncing all data...', '🔄 Sync', 3);

  syncGrievanceFormulasToLog();
  syncGrievanceToMemberDirectory();
  syncMemberToGrievanceLog();

  // Repair checkboxes after sync
  repairGrievanceCheckboxes();
  repairMemberCheckboxes();

  // Run data quality check
  var issues = checkDataQuality();

  if (issues.length > 0) {
    var issueMsg = issues.slice(0, 5).join('\n');
    if (issues.length > 5) {
      issueMsg += '\n... and ' + (issues.length - 5) + ' more issues';
    }

    ui.alert('⚠️ Sync Complete with Data Issues',
      'Data synced successfully, but some issues were found:\n\n' + issueMsg + '\n\n' +
      'Use "Fix Data Issues" from Administrator menu to resolve.',
      ui.ButtonSet.OK);
  } else {
    ss.toast('All data synced! No issues found.', '✅ Success', 3);
  }
}

/**
 * Check data quality and return list of issues
 * @return {Array} List of issue descriptions
 */
function checkDataQuality() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var issues = [];

  // Check Grievance Log
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!grievanceSheet || !memberSheet) return issues;

  var lastGRow = grievanceSheet.getLastRow();
  var lastMRow = memberSheet.getLastRow();

  if (lastGRow <= 1) return issues;

  // Get all member IDs for lookup
  var memberIds = {};
  if (lastMRow > 1) {
    var memberData = memberSheet.getRange(2, MEMBER_COLS.MEMBER_ID, lastMRow - 1, 1).getValues();
    memberData.forEach(function(row) {
      if (row[0]) memberIds[row[0]] = true;
    });
  }

  // Check grievances for missing/invalid member IDs
  var grievanceData = grievanceSheet.getRange(2, 1, lastGRow - 1, GRIEVANCE_COLS.MEMBER_ID).getValues();
  var missingMemberIds = 0;
  var invalidMemberIds = 0;

  grievanceData.forEach(function(row) {
    var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    var memberId = row[GRIEVANCE_COLS.MEMBER_ID - 1];

    if (!memberId || memberId === '') {
      missingMemberIds++;
    } else if (!memberIds[memberId]) {
      invalidMemberIds++;
    }
  });

  if (missingMemberIds > 0) {
    issues.push('⚠️ ' + missingMemberIds + ' grievance(s) have no Member ID');
  }
  if (invalidMemberIds > 0) {
    issues.push('⚠️ ' + invalidMemberIds + ' grievance(s) have Member IDs not found in Member Directory');
  }

  return issues;
}

/**
 * Fix data quality issues with interactive dialog
 */
function fixDataQualityIssues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var issues = checkDataQuality();

  if (issues.length === 0) {
    ui.alert('✅ No Data Issues',
      'All data passes quality checks!\n\n' +
      '• All grievances have valid Member IDs\n' +
      '• All Member IDs exist in Member Directory',
      ui.ButtonSet.OK);
    return;
  }

  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body{font-family:Arial;padding:20px;background:#f5f5f5}' +
    '.container{background:white;padding:25px;border-radius:8px}' +
    'h2{color:#DC2626;margin-top:0}' +
    '.issue{background:#fff5f5;padding:15px;margin:10px 0;border-radius:8px;border-left:4px solid #DC2626}' +
    '.issue-title{font-weight:bold;margin-bottom:5px}' +
    '.issue-desc{font-size:13px;color:#666}' +
    '.fix-option{background:#f8f9fa;padding:12px;margin:8px 0;border-radius:4px;display:flex;align-items:center}' +
    '.fix-option input{margin-right:10px}' +
    'button{padding:12px 24px;border:none;border-radius:4px;cursor:pointer;margin:5px}' +
    '.primary{background:#1a73e8;color:white}' +
    '.secondary{background:#e0e0e0}' +
    '</style></head><body><div class="container">' +
    '<h2>⚠️ Data Quality Issues</h2>' +
    '<p>The following issues were found:</p>' +
    issues.map(function(i) { return '<div class="issue">' + i + '</div>'; }).join('') +
    '<h3>How to Fix:</h3>' +
    '<div class="fix-option"><strong>Option 1:</strong> Manually update Member IDs in Grievance Log</div>' +
    '<div class="fix-option"><strong>Option 2:</strong> Use "Setup Member ID Dropdown" to add validation</div>' +
    '<div class="fix-option"><strong>Option 3:</strong> Add missing members to Member Directory first</div>' +
    '<p style="margin-top:20px"><button class="primary" onclick="google.script.run.showGrievancesWithMissingMemberIds();google.script.host.close()">📋 View Affected Rows</button>' +
    '<button class="secondary" onclick="google.script.host.close()">Close</button></p>' +
    '</div></body></html>'
  ).setWidth(500).setHeight(450);
  ui.showModalDialog(html, '⚠️ Data Quality Issues');
}

/**
 * Show grievances that have missing or invalid Member IDs
 */
function showGrievancesWithMissingMemberIds() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!grievanceSheet) {
    ui.alert('Grievance Log not found');
    return;
  }

  var lastGRow = grievanceSheet.getLastRow();
  if (lastGRow <= 1) {
    ui.alert('No grievances found');
    return;
  }

  // Get all member IDs
  var memberIds = {};
  var lastMRow = memberSheet ? memberSheet.getLastRow() : 1;
  if (lastMRow > 1) {
    var memberData = memberSheet.getRange(2, MEMBER_COLS.MEMBER_ID, lastMRow - 1, 1).getValues();
    memberData.forEach(function(row) {
      if (row[0]) memberIds[row[0]] = true;
    });
  }

  // Find problematic rows
  var grievanceData = grievanceSheet.getRange(2, 1, lastGRow - 1, GRIEVANCE_COLS.MEMBER_ID).getValues();
  var problemRows = [];

  grievanceData.forEach(function(row, index) {
    var grievanceId = row[GRIEVANCE_COLS.GRIEVANCE_ID - 1];
    var memberId = row[GRIEVANCE_COLS.MEMBER_ID - 1];
    var rowNum = index + 2;

    if (!memberId || memberId === '') {
      problemRows.push('Row ' + rowNum + ': ' + grievanceId + ' - NO MEMBER ID');
    } else if (!memberIds[memberId]) {
      problemRows.push('Row ' + rowNum + ': ' + grievanceId + ' - Invalid ID: "' + memberId + '"');
    }
  });

  if (problemRows.length === 0) {
    ui.alert('✅ All Good', 'All grievances have valid Member IDs!', ui.ButtonSet.OK);
    return;
  }

  // Show first 20 rows
  var displayRows = problemRows.slice(0, 20);
  var msg = displayRows.join('\n');
  if (problemRows.length > 20) {
    msg += '\n\n... and ' + (problemRows.length - 20) + ' more rows with issues';
  }

  ui.alert('📋 Grievances with Member ID Issues (' + problemRows.length + ' total)',
    msg + '\n\n' +
    'To fix: Open Grievance Log and update the Member ID column (B) for these rows.',
    ui.ButtonSet.OK);

  // Activate Grievance Log sheet
  ss.setActiveSheet(grievanceSheet);
}

/**
 * Refresh all formulas (force recalculation)
 */
function refreshAllHiddenFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Touch each hidden sheet to force recalc
  var hiddenSheetNames = [
    SHEETS.GRIEVANCE_CALC,
    SHEETS.GRIEVANCE_FORMULAS,
    SHEETS.MEMBER_LOOKUP,
    SHEETS.STEWARD_WORKLOAD_CALC,
    SHEETS.INTERACTIVE_CALC
  ];

  hiddenSheetNames.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      // Force recalc by getting values
      sheet.getDataRange().getValues();
    }
  });

  // Then sync
  syncAllData();

  // Repair checkboxes
  repairGrievanceCheckboxes();

  ss.toast('Formulas refreshed and data synced!', '✅ Success', 3);
}

/**
 * Repair checkboxes in Grievance Log (Message Alert column AC)
 * Call this after any bulk data operations that might overwrite checkboxes
 */
function repairGrievanceCheckboxes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var grievanceSheet = ss.getSheetByName(SHEETS.GRIEVANCE_LOG);

  if (!grievanceSheet) return;

  var lastRow = grievanceSheet.getLastRow();
  if (lastRow < 2) return;

  // Re-apply checkboxes to Message Alert column (AC = column 29)
  grievanceSheet.getRange(2, GRIEVANCE_COLS.MESSAGE_ALERT, lastRow - 1, 1).insertCheckboxes();

  Logger.log('Repaired checkboxes for ' + (lastRow - 1) + ' grievance rows');
}

/**
 * Repair checkboxes in Member Directory (Start Grievance column AE)
 */
function repairMemberCheckboxes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

  if (!memberSheet) return;

  var lastRow = memberSheet.getLastRow();
  if (lastRow < 2) return;

  // Re-apply checkboxes to Start Grievance column (AE = column 31)
  memberSheet.getRange(2, MEMBER_COLS.START_GRIEVANCE, lastRow - 1, 1).insertCheckboxes();

  Logger.log('Repaired checkboxes for ' + (lastRow - 1) + ' member rows');
}

/**
 * Repair all checkboxes in both sheets
 */
function repairAllCheckboxes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Repairing checkboxes...', '🔧 Repair', 2);

  repairGrievanceCheckboxes();
  repairMemberCheckboxes();

  ss.toast('All checkboxes repaired!', '✅ Success', 3);
}
