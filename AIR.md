# 509 Dashboard - Architecture & Implementation Reference

**Version:** 1.9.0 (Grievance Form Workflow + Visual Enhancements)
**Last Updated:** 2026-01-11
**Purpose:** Union grievance tracking and member engagement system for SEIU Local 509

---

## Creator & License

**Creator & Owner:** Wardis N. Vizcaino
**Role:** Steward at SEIU Local 509
**Contact:** wardis@pm.me

**License:** Free for use by non-profit collective bargaining groups and unions. No license required.

---

## Quick Start

> ⚠️ **IMPORTANT: Deploy ONLY `ConsolidatedDashboard.gs`**
> The modular `.gs` files are source files used to generate ConsolidatedDashboard.gs.
> Deploying multiple files will cause function conflicts and trigger errors.

1. Copy **only** `ConsolidatedDashboard.gs` to Google Apps Script
2. Run `CREATE_509_DASHBOARD()` to create 5 sheets + 6 hidden calculation sheets
3. Use `Demo > Seed All Sample Data` to populate test data
4. Customize Config sheet with your organization's values

---

## ⚠️ Protected Code - DO NOT MODIFY

The following code sections are **USER APPROVED** and should **NOT be modified or removed**:

### Custom View Modal Popup

**Location:** `ConsolidatedDashboard.gs` (lines 7536-8160) and `MobileQuickActions.gs` (lines 540-1164)

**Protected Functions:**
| Function | Purpose |
|----------|---------|
| `showInteractiveDashboardTab()` | Opens the modal dialog popup |
| `getInteractiveDashboardHtml()` | Returns the HTML/CSS/JS for the tabbed UI |
| `getInteractiveOverviewData()` | Fetches overview statistics |
| `getInteractiveGrievanceData()` | Fetches grievance list data |
| `getInteractiveMemberData()` | Fetches member list data |

**Why Protected:** This modal provides essential dashboard functionality that users rely on for quick access to data.

### Member Satisfaction Dashboard Modal

**Location:** `Code.gs` (lines 4740-5668)

**Functions:**
| Function | Purpose |
|----------|---------|
| `showSatisfactionDashboard()` | Opens the 900x750 modal dialog |
| `getSatisfactionDashboardHtml()` | Returns HTML with 4-tab interface |
| `getSatisfactionOverviewData()` | Fetches overview stats, NPS score, insights |
| `getSatisfactionResponseData()` | Fetches individual survey responses |
| `getSatisfactionSectionData()` | Fetches scores by survey section |
| `getSatisfactionAnalyticsData()` | Fetches worksite/role analysis, priorities |

**Tabs:**
1. **Overview** - Key metrics (total responses, avg satisfaction, NPS, response rate), gauge charts, auto-generated insights
2. **Responses** - Searchable list with satisfaction level filtering (High/Medium/Needs Attention)
3. **By Section** - Bar chart of all 11 survey sections ranked by score, section detail cards
4. **Insights** - Worksite/role breakdowns, steward contact impact analysis, top member priorities

**Why Added:** Provides interactive survey analysis without leaving the spreadsheet, matching the Custom View pattern.

---

## File Architecture

### Project Structure (9 Source Files)

```
509-dashboard/
├── Constants.gs           # Configuration constants (SHEETS, COLORS, MEMBER_COLS, GRIEVANCE_COLS)
├── Code.gs                # Main entry point, setup, Drive/Calendar/Email, Audit Log
├── SeedNuke.gs            # Demo data seeding and clearing functions
├── HiddenSheets.gs        # Self-healing hidden calculation sheets with auto-sync
├── ADHDFeatures.gs        # ADHD accessibility & theming (focus mode, themes, pomodoro)
├── TestingValidation.gs   # Test framework & data validation
├── PerformanceUndo.gs     # Caching layer & undo/redo system
├── MobileQuickActions.gs  # Mobile interface & quick actions menu
├── WebApp.gs              # Standalone web app for mobile phone access
└── AIR.md                 # This document
```

### File Descriptions

**Constants.gs** (~607 lines)
- `SHEETS` - Sheet name constants (3 data + 2 dashboard + 6 hidden)
- `COLORS` - Brand color scheme
- `MEMBER_COLS` - 31 Member Directory column positions
- `GRIEVANCE_COLS` - 34 Grievance Log column positions
- `CONFIG_COLS` - Config sheet column positions
- `DEFAULT_CONFIG` - Default dropdown values
- `MULTI_SELECT_COLS` - Configuration for multi-select columns
- `getMultiSelectConfig()` - Get multi-select config for a column
- `generateNameBasedId(prefix, firstName, lastName, existingIds)` - Generate unique ID
  - **Format:** `PREFIX + first 2 chars of firstName + first 2 chars of lastName + 3 random digits`
  - **Member ID:** `MJOSM123` (M + "John Smith" → JO + SM + 123)
  - **Grievance ID:** `GJOSM456` (G + "John Smith" → JO + SM + 456)
  - Includes collision detection to ensure uniqueness
- `getColumnLetter()` - Convert column number to letter
- `getColumnNumber()` - Convert column letter to number
- `mapMemberRow()` - Map row array to member object
- `mapGrievanceRow()` - Map row array to grievance object
- `getMemberHeaders()` - Get all 31 member column headers
- `getGrievanceHeaders()` - Get all 34 grievance column headers

**Code.gs** (~3900 lines) - Main entry point with Drive/Calendar/Email/Audit
- `onOpen()` - Creates menu system
- `CREATE_509_DASHBOARD()` - Main setup function (creates 5 sheets + 6 hidden)
- `DIAGNOSE_SETUP()` - System health check
- `REPAIR_DASHBOARD()` - Repair hidden sheets and triggers
- `setupDataValidations()` - Apply dropdown validations
- `setupHiddenSheets()` - Create hidden calculation sheets
- `setDropdownValidation()` - Helper: apply single-select dropdown
- `setMultiSelectValidation()` - Helper: apply multi-select dropdown (allows comma-separated)
- `showMultiSelectDialog()` - Opens multi-select checkbox dialog
- `applyMultiSelectValue()` - Saves multi-select values to cell
- `onSelectionChangeMultiSelect()` - Auto-opens dialog on cell selection
- `installMultiSelectTrigger()` - Enables auto-open mode
- `removeMultiSelectTrigger()` - Disables auto-open mode
- `getOrCreateSheet()` - Helper: get or create sheet
- `rebuildDashboard()` - Refresh data and validations
- `refreshAllFormulas()` - Refresh all formulas and sync
- `recalcAllGrievancesBatched()` - Refresh grievance formulas
- `refreshMemberDirectoryFormulas()` - Refresh member directory
- `searchMembers()` - Desktop search dialog
- `startNewGrievance()` - Opens pre-filled Google Form for new grievance
- `viewActiveGrievances()` - Navigate to Grievance Log
- Sheet creation (7 functions): `createConfigSheet()`, `createMemberDirectory()`, `createGrievanceLog()`, `createDashboard()`, `createInteractiveDashboard()`, `createSatisfactionSheet()`, `createFeedbackSheet()`
- Grievance Form Workflow:
  - `GRIEVANCE_FORM_CONFIG` - Form URL and field entry ID configuration
  - `startNewGrievance()` - Opens pre-filled form with member data from Member Directory
  - `getCurrentStewardInfo_()` - Get current user's steward info
  - `buildGrievanceFormUrl_()` - Build pre-filled form URL with all parameters
  - `onGrievanceFormSubmit(e)` - Form submission trigger handler (adds to Grievance Log + creates folder)
  - `setupGrievanceFormTrigger()` - Menu-driven trigger setup for form submissions
  - `createGrievanceFolderFromData_()` - Create folder with subfolders (Documents, Correspondence, Notes)
  - `shareWithCoordinators_()` - Share folder with coordinators from Config
  - `testGrievanceFormSubmission()` - Test function with sample data
- Contact Info Form Workflow:
  - `CONTACT_FORM_CONFIG` - Form URL and field entry ID configuration for member contact updates
  - `sendContactInfoForm()` - Opens pre-filled contact form for selected member (with copy link option)
  - `buildContactFormUrl_()` - Build pre-filled form URL with member's current data
  - `onContactFormSubmit(e)` - Form submission trigger handler (updates Member Directory)
  - `setupContactFormTrigger()` - Menu-driven trigger setup for contact form submissions
  - `getFormMultiValue_()` - Helper for multi-select checkbox responses
- Google Drive Integration:
  - `setupDriveFolderForGrievance()` - Create folder for grievance
  - `getOrCreateDashboardFolder_()` - Get/create root folder
  - `showGrievanceFiles()` - View files for grievance
  - `batchCreateGrievanceFolders()` - Bulk folder creation
- Calendar Integration:
  - `syncDeadlinesToCalendar()` - Sync deadlines to Google Calendar
  - `showUpcomingDeadlinesFromCalendar()` - View calendar deadlines
- Email Notifications:
  - `showNotificationSettings()` - UI to enable/disable daily notifications
  - `testDeadlineNotifications()` - Test notification system
- Audit Log:
  - `setupAuditLogSheet()` - Create hidden audit log
  - `logAuditEvent()` - Record audit entry
  - `onEditAudit()` - Audit trigger handler
  - `viewAuditLog()` - View audit entries
  - `getAuditHistory()` - Get history for a record
- Member Satisfaction Dashboard:
  - `showSatisfactionDashboard()` - Opens 900x750 modal with 4-tab interface
  - `getSatisfactionDashboardHtml()` - Returns HTML/CSS/JS for the satisfaction dashboard
  - `getSatisfactionOverviewData()` - Fetches overview stats, NPS score, response rate, insights
  - `getSatisfactionResponseData()` - Fetches individual survey responses with filtering support
  - `getSatisfactionSectionData()` - Fetches scores for all 11 survey sections
  - `getSatisfactionAnalyticsData()` - Fetches worksite/role analysis, steward impact, priorities

**SeedNuke.gs** (~1400 lines)
- `SEED_SAMPLE_DATA()` - Seeds Config + 1,000 members + 300 grievances (30%) + 50 survey responses + 3 feedback entries
- `seedConfigData()` - Populate Config dropdowns
- `seedSatisfactionData()` - Seed 50 sample survey responses with realistic branching logic
- `seedFeedbackData()` - Seed 3 sample feedback entries (bug, feature request, improvement)
- `SEED_MEMBERS(count, grievancePercent)` - Seed N members with optional grievances (max 2000 members)
  - Default grievancePercent is 30% if not specified
  - Example: `SEED_MEMBERS(100)` seeds 100 members + ~30 grievances
  - Example: `SEED_MEMBERS(100, 50)` seeds 100 members + ~50 grievances
- `SEED_GRIEVANCES(count)` - Seed N grievances for existing members (max 300)
- `SEED_MEMBERS_DIALOG()` - Prompt for member count (uses 30% grievances)
- `SEED_MEMBERS_ADVANCED_DIALOG()` - Prompt for member count AND grievance percentage
- `SEED_GRIEVANCES_DIALOG()` - Prompt for grievance count
- `seed50Members()` - Shortcut: seed 50 members + 15 grievances (30%)
- `seed100MembersWithGrievances()` - Shortcut: seed 100 members + 50 grievances (50%)
- `generateSingleMemberRow()` - Generate one member row (31 columns)
- `generateSingleGrievanceRow()` - Generate one grievance row (34 columns)
- `NUKE_SEEDED_DATA()` - Clear seeded data with confirmation (preserves manual entries), clears survey responses, deletes Feedback & Menu Checklist sheets
- `NUKE_CONFIG_DROPDOWNS()` - Clear only Config dropdowns
- `getConfigValues()` - Helper: get values from Config column
- `randomChoice()` - Helper: pick random array element
- `randomDate()` - Helper: generate random date
- `addDays()` - Helper: add days to date

**HiddenSheets.gs** (~1500 lines)
- `setupAllHiddenSheets()` - Create all 6 hidden calculation sheets
- Hidden Sheet Setup Functions (6 total):
  - `setupGrievanceCalcSheet()` - Grievance timeline formulas (auto-calc deadlines)
  - `setupGrievanceFormulasSheet()` - Member lookup formulas (First Name, Last Name, Email, etc.)
  - `setupMemberLookupSheet()` - Member → Grievance Log sync
  - `setupStewardContactCalcSheet()` - Steward contact tracking
  - `setupDashboardCalcSheet()` - Dashboard summary metrics (15 key metrics)
  - `setupStewardPerformanceCalcSheet()` - Steward performance scores with weighted formula
- Sync Functions:
  - `syncAllData()` - Sync all cross-sheet data
  - `syncGrievanceToMemberDirectory()` - Sync grievance data to members (AB-AD)
  - `syncMemberToGrievanceLog()` - Sync member data to grievances
  - `syncGrievanceFormulasToLog()` - Sync timeline formulas to Grievance Log
  - `sortGrievanceLogByStatus()` - Auto-sort by status priority and deadline urgency
- Trigger & Repair Functions:
  - `onEditAutoSync()` - Auto-sync trigger handler
  - `installAutoSyncTrigger()` - Install the onEdit trigger
  - `removeAutoSyncTrigger()` - Remove the onEdit trigger
  - `repairAllHiddenSheets()` - Self-healing repair function
  - `repairGrievanceCheckboxes()` - Re-apply checkboxes to Grievance Log AC column
  - `repairMemberCheckboxes()` - Re-apply checkboxes to Member Directory AE column
  - `verifyHiddenSheets()` - Verification and diagnostics
  - `refreshAllHiddenFormulas()` - Force recalculation and sync

**ADHDFeatures.gs** (~400 lines) - ADHD Accessibility & Theming
- `showADHDControlPanel()` - Main ADHD settings panel
- `getADHDSettings()`, `saveADHDSettings()`, `resetADHDSettings()` - Settings management
- `applyADHDSettings()` - Apply visual settings
- `activateFocusMode()`, `deactivateFocusMode()` - Focus mode (hide non-essential sheets)
- `toggleZebraStripes()`, `applyZebraStripes()`, `removeZebraStripes()` - Row banding
- `toggleGridlinesADHD()`, `hideAllGridlines()`, `showAllGridlines()` - Gridline control
- `toggleReducedMotion()` - Animation preferences
- `showQuickCaptureNotepad()` - Quick note-taking dialog
- `startPomodoroTimer()` - Built-in pomodoro timer
- `setBreakReminders()`, `showBreakReminder()` - Break notification system
- `showThemeManager()` - Theme selection UI
- `applyTheme()`, `applyThemeToSheet()`, `previewTheme()` - Theme application
- `getCurrentTheme()`, `resetToDefaultTheme()`, `quickToggleDarkMode()` - Theme utilities
- `setupADHDDefaults()` - Initialize ADHD-friendly defaults

**TestingValidation.gs** (~474 lines) - Testing Framework & Data Validation
- Testing Framework:
  - `Assert` - Assertion library (assertEquals, assertTrue, assertFalse, etc.)
  - `runAllTests()` - Run complete test suite
  - `runQuickTests()` - Run fast unit tests only
  - `generateTestReport()` - Create test results sheet
  - `getTestFunctionRegistry()` - Test function registry
  - Unit tests: `testMemberColsConstants()`, `testGrievanceColsConstants()`, `testColumnLetterConversion()`, etc.
- Validation Framework:
  - `VALIDATION_PATTERNS` - Regex patterns:
    - `MEMBER_ID`: `/^M[A-Z]{4}\d{3}$/` (e.g., MJOSM123)
    - `GRIEVANCE_ID`: `/^G[A-Z]{4}\d{3}$/` (e.g., GJOSM456)
    - `EMAIL`, `PHONE_US` patterns
  - `validateEmailAddress()` - Email format validation with typo detection
  - `validatePhoneNumber()` - Phone validation with auto-formatting
  - `formatUSPhone()` - Format phone to (XXX) XXX-XXXX
  - `validateRequired()` - Required field validation
  - `checkDuplicateMemberID()`, `checkDuplicateGrievanceID()` - Duplicate detection
  - `runBulkValidation()` - Validate all data
  - `showValidationReport()` - Display validation issues
  - `showValidationSettings()` - Validation configuration UI
  - `installValidationTrigger()` - Real-time validation on edit
  - `onEditValidation()` - Validation trigger handler

**PerformanceUndo.gs** (~294 lines) - Caching Layer & Undo/Redo
- Caching:
  - `getCachedData()` - Get data from cache or load
  - `setCachedData()` - Store data in cache
  - `invalidateCache()` - Clear specific cache
  - `invalidateAllCaches()` - Clear all caches
  - `warmUpCaches()` - Pre-populate caches
  - `getCachedGrievances()`, `getCachedMembers()`, `getCachedStewards()` - Cached data getters
  - `getCachedDashboardMetrics()` - Cached dashboard metrics
  - `showCacheStatusDashboard()` - Cache status UI
- Undo/Redo:
  - `getUndoHistory()`, `saveUndoHistory()` - History management
  - `recordAction()` - Record an action for undo
  - `recordCellEdit()`, `recordRowAddition()`, `recordRowDeletion()` - Specific action recording
  - `undoLastAction()`, `redoLastAction()` - Undo/redo operations
  - `undoToIndex()`, `redoToIndex()` - Jump to specific point
  - `applyState()` - Apply state snapshot
  - `clearUndoHistory()` - Clear all history
  - `showUndoRedoPanel()` - Undo/redo UI
  - `exportUndoHistoryToSheet()` - Export history to sheet
  - `createGrievanceSnapshot()`, `restoreFromSnapshot()` - Full snapshot backup

**MobileQuickActions.gs** (~1164 lines) - Mobile Interface & Quick Actions
- Mobile Interface:
  - `showMobileDashboard()` - Touch-optimized dashboard
  - `getMobileDashboardStats()` - Dashboard statistics
  - `getRecentGrievancesForMobile()` - Recent grievances (limit)
  - `showMobileGrievanceList()` - Mobile grievance list
  - `showMobileUnifiedSearch()` - Mobile search UI
  - `getMobileSearchData()` - Search handler
  - `showMyAssignedGrievances()` - View user's assigned cases
- Quick Actions:
  - `showQuickActionsMenu()` - Context-aware quick actions
  - `showMemberQuickActions()` - Quick actions for member row
  - `showGrievanceQuickActions()` - Quick actions for grievance row
  - `quickUpdateGrievanceStatus()` - One-click status update
  - `composeEmailForMember()` - Email composition dialog
  - `sendQuickEmail()` - Send email via MailApp
  - `showMemberGrievanceHistory()` - Member's grievance history
  - `openGrievanceFormForMember()` - Start grievance for member
  - `syncSingleGrievanceToCalendar()` - Sync single grievance to calendar

**WebApp.gs** (~510 lines) - Standalone Web App for Mobile Access
- Web App Entry Point:
  - `doGet(e)` - Web app entry point, routes to dashboard/search/grievances pages
- Dashboard Page:
  - `getWebAppDashboardHtml()` - Main dashboard with stats cards and quick actions
- Search Page:
  - `getWebAppSearchHtml()` - Full-text search across members and grievances
  - `getWebAppSearchResults(query, tab)` - API endpoint for search (reuses getMobileSearchData)
- Grievance List Page:
  - `getWebAppGrievanceListHtml()` - Filterable grievance list by status
  - `getWebAppGrievanceList()` - API endpoint for grievance data
- Utility:
  - `showWebAppUrl()` - Menu function to display web app URL after deployment

**Deployment:** Extensions → Apps Script → Deploy → Web app
**Access:** Bookmark URL on mobile device or add to home screen

---

## Column Mapping System

**CRITICAL: ALL column references must use these constants, never hardcoded letters.**

### MEMBER_COLS (31 columns: A-AE)

```javascript
var MEMBER_COLS = {
  // Section 1: Identity & Core Info (A-D)
  MEMBER_ID: 1,           // A
  FIRST_NAME: 2,          // B
  LAST_NAME: 3,           // C
  JOB_TITLE: 4,           // D

  // Section 2: Location & Work (E-G)
  WORK_LOCATION: 5,       // E
  UNIT: 6,                // F
  OFFICE_DAYS: 7,         // G - Multi-select

  // Section 3: Contact Information (H-K)
  EMAIL: 8,               // H
  PHONE: 9,               // I
  PREFERRED_COMM: 10,     // J - Multi-select
  BEST_TIME: 11,          // K - Multi-select

  // Section 4: Organizational Structure (L-P)
  SUPERVISOR: 12,         // L
  MANAGER: 13,            // M
  IS_STEWARD: 14,         // N
  COMMITTEES: 15,         // O - Multi-select
  ASSIGNED_STEWARD: 16,   // P - Multi-select

  // Section 5: Engagement Metrics (Q-T)
  LAST_VIRTUAL_MTG: 17,   // Q
  LAST_INPERSON_MTG: 18,  // R
  OPEN_RATE: 19,          // S
  VOLUNTEER_HOURS: 20,    // T

  // Section 6: Member Interests (U-X)
  INTEREST_LOCAL: 21,     // U
  INTEREST_CHAPTER: 22,   // V
  INTEREST_ALLIED: 23,    // W
  HOME_TOWN: 24,          // X

  // Section 7: Steward Contact Tracking (Y-AA)
  RECENT_CONTACT_DATE: 25, // Y
  CONTACT_STEWARD: 26,    // Z
  CONTACT_NOTES: 27,      // AA

  // Section 8: Grievance Management (AB-AE)
  HAS_OPEN_GRIEVANCE: 28, // AB - auto-sync: "Yes"/"No" from Grievance Log
  GRIEVANCE_STATUS: 29,   // AC - auto-sync: Status from Grievance Log
  NEXT_DEADLINE: 30,      // AD - auto-sync: Days to Deadline (number or "Overdue")
  START_GRIEVANCE: 31     // AE (checkbox)
};
```

### GRIEVANCE_COLS (34 columns: A-AH)

```javascript
var GRIEVANCE_COLS = {
  // Section 1: Identity (A-D)
  GRIEVANCE_ID: 1,        // A
  MEMBER_ID: 2,           // B
  FIRST_NAME: 3,          // C
  LAST_NAME: 4,           // D

  // Section 2: Status & Assignment (E-F)
  STATUS: 5,              // E
  CURRENT_STEP: 6,        // F

  // Section 3: Timeline - Filing (G-I)
  INCIDENT_DATE: 7,       // G
  FILING_DEADLINE: 8,     // H (auto-calc: +21 days)
  DATE_FILED: 9,          // I

  // Section 4: Timeline - Step I (J-K)
  STEP1_DUE: 10,          // J (auto-calc: +30 days)
  STEP1_RCVD: 11,         // K

  // Section 5: Timeline - Step II (L-O)
  STEP2_APPEAL_DUE: 12,   // L (auto-calc: +10 days)
  STEP2_APPEAL_FILED: 13, // M
  STEP2_DUE: 14,          // N (auto-calc: +30 days)
  STEP2_RCVD: 15,         // O

  // Section 6: Timeline - Step III (P-R)
  STEP3_APPEAL_DUE: 16,   // P (auto-calc: +30 days)
  STEP3_APPEAL_FILED: 17, // Q
  DATE_CLOSED: 18,        // R

  // Section 7: Calculated Metrics (S-U)
  DAYS_OPEN: 19,          // S (auto-calc)
  NEXT_ACTION_DUE: 20,    // T (auto-calc)
  DAYS_TO_DEADLINE: 21,   // U (auto-calc)

  // Section 8: Case Details (V-W)
  ARTICLES: 22,           // V
  ISSUE_CATEGORY: 23,     // W

  // Section 9: Contact & Location (X-AA)
  MEMBER_EMAIL: 24,       // X
  UNIT: 25,               // Y
  LOCATION: 26,           // Z
  STEWARD: 27,            // AA

  // Section 10: Resolution (AB)
  RESOLUTION: 28,         // AB

  // Section 11: Coordinator Notifications (AC-AF)
  MESSAGE_ALERT: 29,      // AC (checkbox)
  COORDINATOR_MESSAGE: 30, // AD
  ACKNOWLEDGED_BY: 31,    // AE
  ACKNOWLEDGED_DATE: 32,  // AF

  // Section 12: Drive Integration (AG-AH)
  DRIVE_FOLDER_ID: 33,    // AG
  DRIVE_FOLDER_URL: 34    // AH
};
```

### SATISFACTION_COLS (68-Question Google Form Response + Summary)

**Form Response Area (A-BP):** Auto-populated by linked Google Form (68 questions + timestamp)
**Summary/Chart Data Area (BT-CD):** Section averages for charts and dashboards

```javascript
var SATISFACTION_COLS = {
  // Form Response Columns (auto-created by Google Form)
  TIMESTAMP: 1,                   // A - Auto by Google Forms

  // Work Context (Q1-5)
  Q1_WORKSITE: 2,                 // B - Worksite/Program/Region
  Q2_ROLE: 3,                     // C - Role/Job Group
  Q3_SHIFT: 4,                    // D - Day/Evening/Night/Rotating
  Q4_TIME_IN_ROLE: 5,             // E - Tenure range
  Q5_STEWARD_CONTACT: 6,          // F - Yes/No (branching)

  // Overall Satisfaction Q6-9, Steward Ratings Q10-17, etc.
  // ... (68 total questions, see Constants.gs for full list)

  // Summary/Chart Data Area (BT onwards)
  SUMMARY_START: 72,              // BT
  AVG_OVERALL_SAT: 72,            // BT - Avg of Q6-Q9
  AVG_STEWARD_RATING: 73,         // BU - Avg of Q10-Q16
  AVG_STEWARD_ACCESS: 74,         // BV - Avg of Q18-Q20
  AVG_CHAPTER: 75,                // BW - Avg of Q21-Q25
  AVG_LEADERSHIP: 76,             // BX - Avg of Q26-Q31
  AVG_CONTRACT: 77,               // BY - Avg of Q32-Q35
  AVG_REPRESENTATION: 78,         // BZ - Avg of Q37-Q40
  AVG_COMMUNICATION: 79,          // CA - Avg of Q41-Q45
  AVG_MEMBER_VOICE: 80,           // CB - Avg of Q46-Q50
  AVG_VALUE_ACTION: 81,           // CC - Avg of Q51-Q55
  AVG_SCHEDULING: 82              // CD - Avg of Q56-Q62
};

// Survey Section Definitions for Analysis
var SATISFACTION_SECTIONS = {
  WORK_CONTEXT: { name: 'Work Context', questions: [2,3,4,5,6], scale: false },
  OVERALL_SAT: { name: 'Overall Satisfaction', questions: [7,8,9,10], scale: true },
  STEWARD_3A: { name: 'Steward Ratings', questions: [11-17], scale: true },
  STEWARD_3B: { name: 'Steward Access', questions: [19,20,21], scale: true },
  CHAPTER: { name: 'Chapter Effectiveness', questions: [22-26], scale: true },
  LEADERSHIP: { name: 'Local Leadership', questions: [27-32], scale: true },
  CONTRACT: { name: 'Contract Enforcement', questions: [33-36], scale: true },
  REPRESENTATION: { name: 'Representation Process', questions: [38-41], scale: true },
  COMMUNICATION: { name: 'Communication Quality', questions: [42-46], scale: true },
  MEMBER_VOICE: { name: 'Member Voice & Culture', questions: [47-51], scale: true },
  VALUE_ACTION: { name: 'Value & Collective Action', questions: [52-56], scale: true },
  SCHEDULING: { name: 'Scheduling/Office Days', questions: [57-63], scale: true },
  PRIORITIES: { name: 'Priorities & Close', questions: [65-68], scale: false }
};
```

### FEEDBACK_COLS (11 columns: A-K)

```javascript
var FEEDBACK_COLS = {
  TIMESTAMP: 1,                // A - Auto-generated timestamp
  SUBMITTED_BY: 2,             // B - Who submitted the feedback
  CATEGORY: 3,                 // C - Area of the system
  TYPE: 4,                     // D - Bug, Feature Request, Improvement
  PRIORITY: 5,                 // E - Low, Medium, High, Critical
  TITLE: 6,                    // F - Short title
  DESCRIPTION: 7,              // G - Detailed description
  STATUS: 8,                   // H - New, In Progress, Resolved, Won't Fix
  ASSIGNED_TO: 9,              // I - Who is working on it
  RESOLUTION: 10,              // J - How it was resolved
  NOTES: 11                    // K - Additional notes
};
```

---

## Sheet Structure (8 Visible + 6 Hidden)

### Core Data Sheets

| # | Sheet Name | Type | Purpose |
|---|------------|------|---------|
| 1 | Config | Data | Master dropdown lists for validation (43 columns) |
| 2 | Member Directory | Data | All member data (31 columns) |
| 3 | Grievance Log | Data | All grievance cases (34 columns) |

### Dashboard Sheets

| # | Sheet Name | Type | Purpose |
|---|------------|------|---------|
| 4 | 💼 Dashboard | View | Executive metrics dashboard with 12 analytics sections |
| 5 | 🎯 Custom View | View | Customizable metrics with dropdowns |

### Tracking Sheets

| # | Sheet Name | Type | Purpose |
|---|------------|------|---------|
| 6 | 📊 Member Satisfaction | Data | 68-question Google Form survey with dashboard, charts (82 cols + dashboard) |
| 7 | 💡 Feedback & Development | Data | Bug/feature tracking with priority (11 columns) |
| 8 | ✅ Menu Checklist | Reference | Function reference guide organized by 13 phases |

#### 💼 Dashboard - 12 Live Analytics Sections

| # | Section | Color | Metrics | Data Source |
|---|---------|-------|---------|-------------|
| 1 | QUICK STATS | 🟢 Green | Total Members, Active Stewards, Active Grievances, Win Rate, Overdue, Due This Week | `_Dashboard_Calc` |
| 2 | MEMBER METRICS | 🔵 Blue | Total Members, Active Stewards, Avg Open Rate, YTD Vol Hours | Member Directory |
| 3 | GRIEVANCE METRICS | 🟠 Orange | Open, Pending Info, Settled, Won, Denied, Withdrawn | Grievance Log |
| 4 | TIMELINE & PERFORMANCE | 🟣 Purple | Avg Days Open, Filed This Month, Closed This Month, Avg Resolution | Grievance Log |
| 5 | TYPE ANALYSIS | 🔷 Indigo | 5 issue categories × (Total, Open, Resolved, Win Rate, Avg Days) | Grievance Log |
| 6 | LOCATION BREAKDOWN | 🔵 Cyan | 5 locations × (Members, Grievances, Open Cases, Win Rate) | Config + Member Dir + Grievance Log |
| 7 | MONTH-OVER-MONTH TRENDS | 🔴 Red | Filed/Closed/Won × (This Month, Last Month, Change, % Change, Trend) | Grievance Log |
| 8 | STATUS LEGEND | ⬜ Gray | Color/icon reference guide | Static |
| 9 | STEWARD PERFORMANCE SUMMARY | 🟣 Purple | Total Stewards, Active w/Cases, Avg Cases, Vol Hours, Contacts | Member Dir + Grievance Log |
| 10 | TOP 30 BUSIEST STEWARDS | 🔴 Dark Red | Rank, Steward Name, Active Cases, Open, Pending Info, Total Ever | Grievance Log |
| 11 | TOP 10 PERFORMERS BY SCORE | 🟢 Green | Rank, Steward, Score, Win Rate %, Avg Days, Overdue | `_Steward_Performance_Calc` |
| 12 | STEWARDS NEEDING SUPPORT | 🔴 Red | Rank, Steward, Score, Win Rate %, Avg Days, Overdue | `_Steward_Performance_Calc` |

**All sections use live COUNTIF/COUNTIFS/AVERAGEIFS formulas that auto-update when source data changes.**

### Hidden Calculation Sheets

| # | Sheet Name | Purpose |
|---|------------|---------|
| 1 | `_Grievance_Calc` | Grievance → Member Directory sync (AB-AD) |
| 2 | `_Grievance_Formulas` | Member → Grievance Log sync (C-D, X-AA) |
| 3 | `_Member_Lookup` | Member data lookup formulas |
| 4 | `_Steward_Contact_Calc` | Steward contact tracking (Y-AA) |
| 5 | `_Dashboard_Calc` | Dashboard summary metrics (15 key KPIs) |
| 6 | `_Steward_Performance_Calc` | Per-steward performance scores with weighted formula |

---

## Config Sheet & Dropdown Validations

### Config Columns (15 columns + HOME_TOWNS at AF)

| Column | Name | Used By |
|--------|------|---------|
| A | Job Titles | Member Directory (D) |
| B | Office Locations | Member Directory (E), Grievance Log (Z) |
| C | Units | Member Directory (F), Grievance Log (Y) |
| D | Office Days | Member Directory (G) |
| E | Yes/No | Member Directory (N, U, V, W) |
| F | Supervisors | Member Directory (L) |
| G | Managers | Member Directory (M) |
| H | Stewards | Member Directory (P, Z), Grievance Log (AA) |
| I | Grievance Status | Grievance Log (E) |
| J | Grievance Step | Grievance Log (F) |
| K | Issue Category | Grievance Log (W) |
| L | Articles Violated | Grievance Log (V) |
| M | Communication Methods | Member Directory (J) |
| N | (blank) | - |
| O | Grievance Coordinators | Admin use |
| AF | Home Towns | Member Directory (X) |

### Member Directory Dropdowns (16 columns)

| Column | Field | Config Source | Multi-Select |
|--------|-------|---------------|--------------|
| D | Job Title | JOB_TITLES (A) | No |
| E | Work Location | OFFICE_LOCATIONS (B) | No |
| F | Unit | UNITS (C) | No |
| G | Office Days | OFFICE_DAYS (D) | **Yes** |
| J | Preferred Communication | COMM_METHODS (N) | **Yes** |
| K | Best Time to Contact | BEST_TIMES (AE) | **Yes** |
| L | Supervisor | SUPERVISORS (F) | No |
| M | Manager | MANAGERS (G) | No |
| N | Is Steward | YES_NO (E) | No |
| O | Committees | STEWARD_COMMITTEES (I) | **Yes** |
| P | Assigned Steward | STEWARDS (H) | **Yes** |
| U | Interest: Local | YES_NO (E) | No |
| V | Interest: Chapter | YES_NO (E) | No |
| W | Interest: Allied | YES_NO (E) | No |
| X | Home Town | HOME_TOWNS (AF) | No |
| Z | Contact Steward | STEWARDS (H) | No |

### Grievance Log Dropdowns (5 columns)

| Column | Field | Config Source |
|--------|-------|---------------|
| B | Member ID | Member Directory (A) - dynamic |
| E | Status | GRIEVANCE_STATUS (I) |
| F | Current Step | GRIEVANCE_STEP (J) |
| V | Articles Violated | ARTICLES (L) |
| W | Issue Category | ISSUE_CATEGORY (K) |

### Multi-Select Functionality

Columns marked as **Multi-Select** support comma-separated values for multiple selections.

**Auto-Open Mode (Recommended):**
1. Go to **🔧 Tools > ☑️ Multi-Select > ⚡ Enable Auto-Open**
2. Now clicking any multi-select cell automatically opens the dialog!
3. To disable: **🔧 Tools > ☑️ Multi-Select > 🚫 Disable Auto-Open**

**Manual Mode:**
1. Select a cell in a multi-select column (G, J, K, O, or P)
2. Go to **🔧 Tools > ☑️ Multi-Select > 📝 Open Editor**
3. Check multiple options in the dialog
4. Click **Save** to apply

**Storage format:** Values are stored as comma-separated text (e.g., "Monday, Wednesday, Friday")

**Validation:** Multi-select columns show a dropdown for convenience but accept any text value to allow multiple selections.

---

## Menu System

```
👤 Dashboard
├── 📊 Smart Dashboard (Auto-Detect)
├── 🎯 Custom View
├── 📊 Member Satisfaction          ← NEW: Interactive survey dashboard modal
├── 🔍 Search Members
├── 📋 View Active Grievances
├── 📱 Mobile Dashboard
├── 📱 Get Mobile App URL        ← NEW: Shows web app URL for mobile access
├── ⚡ Quick Actions
└── Grievance Tools
    ├── Start New Grievance
    ├── Refresh Grievance Formulas
    └── Refresh Member Directory Data

📊 Sheet Manager
├── Rebuild Dashboard
└── Refresh All Formulas

🔧 Tools
├── ADHD & Accessibility (submenu)
├── Theming (submenu)
├── ☑️ Multi-Select (submenu)
│   ├── 📝 Open Editor
│   ├── ⚡ Enable Auto-Open
│   └── 🚫 Disable Auto-Open
├── Undo/Redo (submenu)
├── Cache & Performance (submenu)
└── Validation (submenu)

🏗️ Setup
├── CREATE 509 DASHBOARD
├── REPAIR DASHBOARD
└── Setup Data Validations

🎭 Demo
├── 🚀 Seed All Sample Data (1,000 members + 300 grievances + 50 surveys + 3 feedback)
└── 🗑️ Nuke Data (submenu)
    ├── ☢️ NUKE SEEDED DATA
    ├── 🧹 Clear Config Dropdowns Only
    └── 🔄 Restore Config & Dropdowns

⚙️ Administrator
├── DIAGNOSE SETUP
├── Verify Hidden Sheets
└── Setup & Triggers (submenu)
```

---

## Config Sheet Columns

| Column | Name | Content |
|--------|------|---------|
| A | Job Titles | User populates |
| B | Office Locations | User populates |
| C | Units | User populates |
| D | Office Days | Monday-Sunday (preset) |
| E | Yes/No | Yes, No (preset) |
| F | Supervisors | User populates |
| G | Managers | User populates |
| H | Stewards | User populates |
| I | Grievance Status | Open, Pending Info, Settled, etc. (preset) |
| J | Grievance Step | Informal, Step I, Step II, etc. (preset) |
| K | Issue Category | Discipline, Workload, etc. (preset) |
| L | Articles Violated | Art. 1 - Art. 26 (preset) |
| M | Communication Methods | Email, Phone, Text, In Person (preset) |
| O | Grievance Coordinators | User populates |
| AF | Home Towns | User populates |

---

## Color Scheme

```javascript
var COLORS = {
  PRIMARY_PURPLE: '#7C3AED',  // Main brand
  UNION_GREEN: '#059669',     // Success
  SOLIDARITY_RED: '#DC2626',  // Alert/urgent
  PRIMARY_BLUE: '#7EC8E3',    // Light blue
  ACCENT_ORANGE: '#F97316',   // Warnings
  LIGHT_GRAY: '#F3F4F6',      // Backgrounds
  TEXT_DARK: '#1F2937',       // Primary text
  WHITE: '#FFFFFF'
};
```

---

## Usage Patterns

### Dynamic Column References

```javascript
// CORRECT - Dynamic column reference
var statusCol = getColumnLetter(GRIEVANCE_COLS.STATUS);
var formula = '=COUNTIF(\'Grievance Log\'!' + statusCol + ':' + statusCol + ',"Open")';

// WRONG - Hardcoded column (NEVER DO THIS)
var formula = '=COUNTIF(\'Grievance Log\'!E:E,"Open")';
```

### Safe Row Writes

```javascript
// CORRECT - Protects header row
var startRow = Math.max(sheet.getLastRow() + 1, 2);

// WRONG - May overwrite headers
var startRow = sheet.getLastRow() + 1;
```

### Dynamic Sheet Names

```javascript
// CORRECT
var sheet = ss.getSheetByName(SHEETS.MEMBER_DIR);

// WRONG
var sheet = ss.getSheetByName('Member Directory');
```

---

## Seed Data Limits

| Function | Max Count | Batch Size |
|----------|-----------|------------|
| SEED_MEMBERS() | 2,000 | 50 rows |
| SEED_GRIEVANCES() | 300 | 25 rows |

---

## Hidden Sheet Architecture (Self-Healing)

The system uses 6 hidden calculation sheets with auto-sync triggers for cross-sheet data population. Formulas are stored in hidden sheets and synced to visible sheets, making them **self-healing** - if formulas are accidentally deleted, running REPAIR_DASHBOARD() restores them.

### Hidden Sheets (6 total)

| Sheet | Source | Destination | Purpose |
|-------|--------|-------------|---------|
| `_Grievance_Calc` | Grievance Log | Member Directory | AB-AD (Has Open Grievance?, Status, Days to Deadline) |
| `_Grievance_Formulas` | Member Directory | Grievance Log | C-D (Name), H-P (Timeline), S-U (Days Open, Next Action, Days to Deadline), X-AA (Contact) |
| `_Member_Lookup` | Member Directory | Grievance Log | Member data lookup |
| `_Steward_Contact_Calc` | Member Directory | Contact Reports | Y-AA (Contact tracking) |
| `_Dashboard_Calc` | Both | 💼 Dashboard | 15 summary metrics (Win Rate, Overdue, Due This Week, etc.) |
| `_Steward_Performance_Calc` | Grievance Log | 💼 Dashboard | Per-steward performance scores with weighted formula |

### Auto-Sync Trigger

The `onEditAutoSync` trigger automatically syncs data when:
- Grievance Log is edited → Updates Member Directory columns AB-AD, then auto-sorts by status
- Member Directory is edited → Updates Grievance Log columns C-D, X-AA

### Key Functions (HiddenSheets.gs)

| Function | Purpose |
|----------|---------|
| `setupAllHiddenSheets()` | Create all 6 hidden sheets with formulas |
| `repairAllHiddenSheets()` | Recreate sheets, install trigger, sync data |
| `installAutoSyncTrigger()` | Install the onEdit auto-sync trigger |
| `verifyHiddenSheets()` | Verify all sheets and triggers are working |
| `syncAllData()` | Manual sync of all cross-sheet data |
| `syncGrievanceToMemberDirectory()` | Sync grievance data to members |
| `syncMemberToGrievanceLog()` | Sync member data to grievances |
| `sortGrievanceLogByStatus()` | Auto-sort by status priority and deadline |
| `setupDashboardCalcSheet()` | Create dashboard metrics calculation sheet |

### Self-Healing

If hidden sheets get corrupted or deleted:
1. Run `REPAIR_DASHBOARD()` from Setup menu
2. Or run `repairAllHiddenSheets()` from Administrator menu

This recreates all formulas and reinstalls the auto-sync trigger.

---

## Known Issues / Later TODO

### BUG: Days to Deadline Shows Duplicate Values

**Status:** FIXED (2025-12-16)
**Discovered:** 2025-12-16
**Severity:** Medium

**Issue:**
The "Days to Deadline" column (U) in the Grievance Log displayed identical values for multiple rows (e.g., `17.71857539` repeated for all grievances).

**Root Cause:**
The hidden sheet `_Grievance_Formulas` used ARRAYFORMULA with a FILTER-based row index. ARRAYFORMULA doesn't expand correctly when its source column is a FILTER result, causing all rows to receive the same calculated value.

**Fix Applied:**
Changed `syncGrievanceFormulasToLog()` in `HiddenSheets.gs` to calculate Days Open, Next Action Due, and Days to Deadline directly in JavaScript from the grievance row data, bypassing the problematic hidden sheet formulas.

**Calculations now performed directly:**
- **Days Open**: `(Date Closed or Today) - Date Filed` (in whole days)
- **Next Action Due**: Based on Current Step (Informal→Filing Deadline, Step I→Step I Due, etc.)
- **Days to Deadline**: `Next Action Due - Today` (in whole days)
- All deadline dates (Filing Deadline, Step I Due, etc.) also calculated directly

---

## Changelog

### Version 1.9.0 (2026-01-11) - Visual Enhancements, Progress Tracking & Grievance Form Workflow

**New Features: Data Validation, Heatmaps, Progress Bar & Automated Grievance Workflow**

Added visual data quality indicators, deadline heatmaps, grievance progress tracking, and automated grievance form workflow with Drive folder creation.

---

#### Member Directory Enhancements

**1. Empty Field Validation (Red Background)**
- Email field (Column H): Red background (`#ffcdd2`) when empty but Member ID exists
- Phone field (Column I): Red background (`#ffcdd2`) when empty but Member ID exists
- Formula: `=AND($A2<>"",ISBLANK($H2))` ensures only rows with members are highlighted

**2. Days to Deadline Heatmap (Column AD)**
| Days Remaining | Background | Text Color | Style |
|----------------|------------|------------|-------|
| Overdue or ≤0  | `#ffebee` (red) | `#c62828` | Bold |
| 1-3 days       | `#fff3e0` (orange) | `#e65100` | Bold |
| 4-7 days       | `#fffde7` (yellow) | `#f57f17` | Normal |
| 8+ days        | `#e8f5e9` (green) | `#2e7d32` | Normal |

**3. Column Sorting via Filter**
- Added filter row to Member Directory header
- All columns now sortable via dropdown (A-Z, Z-A)

---

#### Grievance Log Enhancements

**4. Days to Deadline Heatmap (Column U)**
Same color scheme as Member Directory - automatically applied when sheet is created.

**5. Grievance Progress Bar (Columns J-R)**
Visual progress indicator showing grievance stage via colored backgrounds:

| Current Step | Columns Highlighted | Color |
|--------------|---------------------|-------|
| Informal | None | Gray `#fafafa` |
| Step I | J-K | Soft blue `#e3f2fd` |
| Step II | J-O | Soft blue `#e3f2fd` |
| Step III | J-Q | Soft blue `#e3f2fd` |
| Closed/Won/Denied/Settled/Withdrawn | J-R | Soft green `#e8f5e9` |

- Progress bar spans: Step I Due → Step I Rcvd → Step II Appeal Due → Step II Appeal Filed → Step II Due → Step II Rcvd → Step III Appeal Due → Step III Appeal Filed → Date Closed
- Columns not yet reached remain light gray
- Completed grievances show all columns in green

**6. Auto-Sort Confirmation**
Grievance Log entries automatically sort by status priority (active cases first) and deadline urgency when edited.

---

#### Grievance Form Workflow (New)

**7. Pre-filled Google Form Integration**
- `startNewGrievance()`: Opens pre-filled Google Form with member data
- Form fields auto-populated from Member Directory (Member ID, name, job title, location, email, etc.)
- Steward info auto-populated from current user's session (if they're a steward in Member Directory)
- Default values: Date Filed = today, Step = I

**8. Automatic Form Submission Processing**
- `onGrievanceFormSubmit(e)`: Trigger handler for form submissions
- Generates unique Grievance ID (format: GXXXX123 based on member name)
- Adds grievance to Grievance Log with all form data
- Calculates deadlines via hidden sheet formulas
- Updates Member Directory grievance status

**9. Automatic Drive Folder Creation**
- Creates folder in "509 Dashboard - Grievance Files" root folder
- Folder name format: `GXXXX123 - FirstName LastName (MemberID)`
- Creates subfolders: 📄 Documents, 📧 Correspondence, 📝 Notes
- Automatically shares with Grievance Coordinators from Config (column O)
- Stores folder ID and URL in Grievance Log (columns AG, AH)

**10. Easy Trigger Setup**
- New menu: 👤 Dashboard > 📋 Grievance Tools > 📋 Setup Grievance Form Trigger
- Prompts for Google Form edit URL
- Creates installable trigger for form submissions

---

#### Personal Contact Info Form (New)

**11. Pre-filled Contact Update Form**
- `sendContactInfoForm()`: Opens pre-filled Google Form with member's current data
- Stewards can send form link to members to update their contact info
- Form fields: First/Last Name, Job Title, Unit, Work Location, Office Days, Communication Preferences, Best Time, Supervisor, Manager, Email, Phone, Interest levels
- Option to open form or copy link to send to member

**12. Automatic Member Directory Updates**
- `onContactFormSubmit(e)`: Trigger handler for contact form submissions
- Matches member by First Name + Last Name
- Updates all submitted fields in Member Directory
- Handles multi-select fields (Office Days, Preferred Communication, Best Time)

**13. Easy Trigger Setup**
- New menu: 👤 Dashboard > 👤 Member Tools > 📋 Setup Contact Form Trigger
- Prompts for Google Form edit URL
- Creates installable trigger for contact form submissions

---

**Code Changes:**

*Member Directory (`createMemberDirectory()` lines 497-576):*
- Lines 497-515: Empty Email/Phone validation rules
- Lines 517-554: Days to Deadline heatmap rules
- Lines 560-576: Filter for column sorting

*Grievance Log (`createGrievanceLog()` lines 645-739):*
- Lines 645-682: Days to Deadline heatmap rules
- Lines 684-731: Progress bar conditional formatting rules
- Lines 733-739: Apply all rules

*Grievance Form Workflow (Code.gs lines 3254-3813):*
- Lines 3258-3283: `GRIEVANCE_FORM_CONFIG` with form URL and 18 field entry IDs
- Lines 3290-3383: `startNewGrievance()` - opens pre-filled form
- Lines 3389-3412: `getCurrentStewardInfo_()` - get steward from session
- Lines 3419-3449: `buildGrievanceFormUrl_()` - build pre-filled URL
- Lines 3465-3561: `onGrievanceFormSubmit(e)` - form submission handler
- Lines 3568-3609: Helper functions (getFormValue_, parseFormDate_, getExistingGrievanceIds_)
- Lines 3615-3653: `createGrievanceFolderFromData_()` - create folder with subfolders
- Lines 3660-3683: `shareWithCoordinators_()` - share folder with coordinators
- Lines 3690-3780: `setupGrievanceFormTrigger()` - menu-driven trigger setup
- Lines 3787-3813: `testGrievanceFormSubmission()` - test function

*Contact Info Form Workflow (Code.gs lines 3844-4100):*
- Lines 3286-3311: `CONTACT_FORM_CONFIG` with form URL and 15 field entry IDs
- Lines 3854-3945: `sendContactInfoForm()` - opens pre-filled form with copy link option
- Lines 3951-3995: `buildContactFormUrl_()` - build pre-filled URL with multi-select support
- Lines 4001-4076: `onContactFormSubmit(e)` - form submission handler (updates Member Directory)
- Lines 4082-4091: `getFormMultiValue_()` - helper for checkbox responses
- Lines 4097-4150: `setupContactFormTrigger()` - menu-driven trigger setup

*Menu Update (Code.gs lines 35-47):*
- Added "📋 Setup Grievance Form Trigger" to Grievance Tools submenu
- Added new "👤 Member Tools" submenu with:
  - "📋 Send Contact Info Form"
  - "📋 Setup Contact Form Trigger"

**Build Process:** Run `node build.js` to regenerate ConsolidatedDashboard.gs

---

### Version 1.8.0 (2026-01-06) - Member Satisfaction Dashboard

**New Feature: Interactive Satisfaction Dashboard Modal**

Added a new modal popup dashboard for analyzing member satisfaction survey data, accessible via **Dashboard > 📊 Member Satisfaction**.

**Functions Added (Code.gs):**
- `showSatisfactionDashboard()` - Opens 900x750 modal dialog
- `getSatisfactionDashboardHtml()` - Returns HTML with 4-tab interface
- `getSatisfactionOverviewData()` - Calculates overview stats, NPS score, response rate
- `getSatisfactionResponseData()` - Fetches individual responses with filtering support
- `getSatisfactionSectionData()` - Fetches scores for all 11 survey sections
- `getSatisfactionAnalyticsData()` - Generates worksite/role analysis, steward impact, priorities

**Dashboard Tabs:**
1. **Overview** - 6 stat cards (responses, avg satisfaction, NPS, response rate, steward rating, leadership), circular gauge charts for key metrics, auto-generated insights
2. **Responses** - Searchable list with satisfaction level filtering (High ≥7, Medium 5-7, Low <5)
3. **By Section** - Bar chart showing all 11 survey sections ranked by score, section detail cards with progress bars
4. **Insights** - Key findings, satisfaction by worksite, by role, steward contact impact analysis, top member priorities

**Design:**
- Green gradient header theme (differentiates from purple Custom View)
- Mobile-responsive with 44px touch targets
- Client-side search and filtering for instant results
- Color-coded scores: green (7+), yellow (5-7), red (<5)
- Lazy-loads data per tab for performance

**Menu Location:** Dashboard > 📊 Member Satisfaction

---

### Version 1.7.0 (2026-01-03) - Dashboard Restructure & Bug Fixes

**Major Features:**

1. **Dashboard Restructure**
   - Removed unused columns G onwards - Dashboard now uses only columns A-F
   - Moved Steward Performance section to near bottom for better visual hierarchy
   - Added new **Top 30 Busiest Stewards** section showing stewards with most active cases
   - Compact Status Legend now fits in single row (A-F)
   - All numbers now formatted with comma separators (1,000)

2. **Overdue Cases Bug Fix**
   - Fixed Dashboard showing 0 for Overdue Cases
   - Root cause: Formula was looking for `<0` but Days to Deadline shows "Overdue" text
   - Changed formula to count cells containing "Overdue" text
   - Also fixed in hidden sheet Steward Workload, Location Analytics calculations

3. **Member Directory Enhancements**
   - Added collapsible column groups for Engagement Metrics (Q-T) and Member Interests (U-X)
   - Both groups are collapsed/hidden by default for cleaner view
   - Added conditional formatting: Has Open Grievance = "Yes" shows red background
   - Added comma formatting for numeric columns (Open Rate, Volunteer Hours)

4. **Tab Rename: Interactive → Custom View**
   - Renamed "🎯 Interactive" tab to "🎯 Custom View" for clarity
   - Updated SHEETS constant and all string references

5. **Code Cleanup**
   - Removed orphaned `setupInteractiveDashboardCalcSheet()` function (defined but never called)
   - Removed orphaned `setupEngagementCalcSheet()` function (used undefined constant)
   - Fixed hidden sheet number comments for accuracy

**Code Changes:**
- Code.gs: ~150 lines added for Dashboard restructure and Member Directory enhancements
- HiddenSheets.gs: Fixed 5 Overdue Cases formulas (changed `"<0"` to `"Overdue"`)
- HiddenSheets.gs: Removed 2 orphaned functions (~60 lines)
- Constants.gs: Changed `INTERACTIVE: '🎯 Interactive'` to `INTERACTIVE: '🎯 Custom View'`
- MobileQuickActions.gs: Updated "Interactive Dashboard" text references

---

### Version 1.6.0 (2026-01-03) - Desktop Search & Unified Seeding

**Major Features:**

1. **Desktop Search with Advanced Filtering**
   - Comprehensive search across Members and Grievances in one interface
   - Accessible via **Dashboard → 🔍 Search Members**
   - Tabbed interface: All, Members, Grievances
   - Advanced filters: Status (grievances), Location (all), Is Steward (members)
   - Searchable fields: Name, ID, Email, Job Title, Location, Issue Type, Steward
   - Click-to-navigate: Jump directly to any result row in the spreadsheet
   - Desktop optimized: 900x700 modal with responsive grid layout
   - Debounced search (300ms) for performance

2. **Dashboard STATUS Column Fix**
   - Fixed Dashboard metrics to use STATUS column for Win/Settled/Denied counts
   - Previously incorrectly referenced RESOLUTION column for some metrics
   - Status now includes both workflow states AND outcomes (single column design)
   - Ensures accurate grievance outcome tracking

3. **Unified Seed Function**
   - SEED_GRIEVANCES merged into SEED_MEMBERS function
   - Use `SEED_MEMBERS(count, grievancePercent)` to seed members with optional grievances
   - All seeded grievances are directly linked to members (no orphaned data)
   - Example: `SEED_MEMBERS(100)` seeds 100 members + ~30 grievances (30% default)
   - Example: `SEED_MEMBERS(100, 50)` seeds 100 members + ~50 grievances (50%)
   - Example: `SEED_MEMBERS(100, 0)` seeds members only, no grievances

4. **Automatic Timeline Column Grouping**
   - Grievance Log now automatically sets up collapsible column groups during creation
   - Step I columns (J-K) grouped together
   - Step II columns (L-O) grouped together
   - Step III columns (P-Q) grouped together
   - Click +/- controls to expand/collapse step details
   - Previously required manual setup via View menu

5. **Simplified Demo Menu**
   - Removed redundant "Seed Data" submenu
   - Single "🚀 Seed All Sample Data" option for seeding
   - Cleaner menu structure

6. **Member Directory Days to Deadline Fix**
   - Fixed Member Directory column AD not populating from Grievance Log
   - Root cause: MINIFS formula in _Grievance_Calc ignored "Overdue" text values
   - Solution: `syncGrievanceToMemberDirectory()` now calculates directly from Grievance Log
   - Properly handles both numeric deadlines and "Overdue" text
   - Shows minimum deadline when member has multiple open grievances

7. **SEED_SAMPLE_DATA Corrected**
   - Fixed to seed 1,000 members + 300 grievances (was incorrectly 50/25)
   - Uses merged approach: `SEED_MEMBERS(1000, 30)` for linked data
   - Automatically installs auto-sync trigger for live updates
   - Matches specification from SeedNuke.gs

**New Functions:**
- `showDesktopSearch()` - Main desktop search dialog (~300 lines HTML/JS)
- `getDesktopSearchLocations()` - Get unique locations for filter dropdown
- `getDesktopSearchData(query, tab, filters)` - Backend search handler
- `navigateToSearchResult(type, id, row)` - Navigate to search result row

**Code Changes:**
- `Code.gs`: Updated `searchMembers()` to call `showDesktopSearch()`
- `ConsolidatedDashboard.gs`: Added desktop search functions
- `ConsolidatedDashboard.gs`: `createGrievanceLog()` now auto-creates column groups
- `ConsolidatedDashboard.gs`: Rewrote `syncGrievanceToMemberDirectory()` to calculate directly
- `ConsolidatedDashboard.gs`: Fixed `SEED_SAMPLE_DATA()` to seed 1000 members + 300 grievances
- `Constants.gs`: Updated `GRIEVANCE_STATUS` comment for clarity
- `HiddenSheets.gs`: Fixed Dashboard formulas to use STATUS column for outcome counts
- `SeedNuke.gs`: Merged grievance seeding into SEED_MEMBERS function

**Desktop vs Mobile Search Comparison:**
| Aspect | Mobile | Desktop |
|--------|--------|---------|
| Search Fields | ID, Name, Email, Status | ID, Name, Email, Job Title, Location, Issue Type, Steward |
| Filters | Tab only | Status, Location, Is Steward |
| Result Limit | 20 | 50 |
| Navigation | No | Yes - click to jump to row |

---

### Version 1.5.1 (2026-01-03) - Mobile Web App for Phone Access

**Problem Solved:**
Google Sheets mobile app does not support Apps Script custom menus, making the dashboard and search features inaccessible on phones.

**Solution:**
Added standalone web app deployment that can be accessed via URL on any mobile browser.

**New File:**
- `WebApp.gs` (~510 lines) - Standalone web app with doGet() entry point

**Features:**
- **Dashboard Page**: Stats cards (Total, Active, Pending, Overdue grievances) + quick action buttons
- **Search Page**: Full-text search across members and grievances with tabbed filtering
- **Grievance List Page**: Filterable list by status (All, Open, Pending, Resolved)
- **Bottom Navigation**: Touch-friendly navigation bar
- **iOS Home Screen Support**: apple-mobile-web-app meta tags for app-like experience

**New Menu Item:**
- Dashboard → 📱 Get Mobile App URL - Shows deployment URL after web app is deployed

**Deployment Instructions:**
1. Extensions → Apps Script → Deploy → New deployment
2. Select "Web app"
3. Set access permissions
4. Copy URL and bookmark on mobile device

**Code Changes:**
- `WebApp.gs`: New file with doGet(), getWebAppDashboardHtml(), getWebAppSearchHtml(), getWebAppGrievanceListHtml(), getWebAppSearchResults(), getWebAppGrievanceList(), showWebAppUrl()
- `Code.gs`: Added menu item "📱 Get Mobile App URL" calling showWebAppUrl()

---

### Version 1.5.0 (2025-12-18) - Enhanced Analytics Dashboard

**Major Updates:**

- Enhanced 💼 Dashboard from 4 sections to 9 comprehensive analytics sections
- Added Operations Analytics-style metrics inspired by 509-dashboard
- All sections now use live auto-updating formulas (no manual refresh needed)

**New Dashboard Sections (5 added):**

| Section | Description |
|---------|-------------|
| TYPE ANALYSIS | Breakdown by issue category (Contract Violation, Discipline, Workload, Safety, Discrimination) with Total/Open/Resolved/Win Rate/Avg Days per category |
| LOCATION BREAKDOWN | Metrics per work location from Config - Members, Grievances, Open Cases, Win Rate |
| STEWARD PERFORMANCE | Total Stewards, Active w/Cases, Avg Cases/Steward, Vol Hours, Contacts This Month |
| MONTH-OVER-MONTH TRENDS | Filed/Closed/Won comparisons with 📈📉➡️ trend indicators and % change |
| STATUS LEGEND | Moved from section 5 to section 9 |

**Formula Types Used:**

- `COUNTIF` / `COUNTIFS` - Count by criteria
- `AVERAGEIFS` - Average by criteria
- `IFERROR` / `IF` - Error handling
- `TEXT` - Percentage formatting
- Date math with `DATE()`, `YEAR()`, `MONTH()`, `TODAY()`

**Live Data Flow:**

```
Config (Office Locations) ──┐
                            ├──► 💼 Dashboard (auto-updates)
Member Directory ───────────┤
                            │
Grievance Log ──────────────┘
```

**Code Changes:**

- Code.gs: Enhanced `createDashboard()` from ~170 lines to ~320 lines
- Code.gs: Fixed `rebuildDashboard()` to actually recreate dashboard sheets
- AIR.md: Updated documentation with 9-section dashboard architecture

---

### Version 1.4.5 (2025-12-18) - Auto-Sort & Seed Improvements

**New Feature: Auto-Sort Grievance Log by Status**
- Grievance Log now automatically sorts when edited
- Primary sort: Status priority (active cases first, resolved cases last)
- Secondary sort: Days to deadline (most urgent first within each status)

**Status Priority Order:**
| Priority | Status | Type |
|----------|--------|------|
| 1 | Open | Active |
| 2 | Pending Info | Active |
| 3 | In Arbitration | Active |
| 4 | Appealed | Active |
| 5 | Settled | Resolved |
| 6 | Won | Resolved |
| 7 | Denied | Resolved |
| 8 | Withdrawn | Resolved |
| 9 | Closed | Resolved |

**Seed Data Improvements:**
- Expanded name pools from 20 to 120 names each (14,400+ unique combinations)
- Significantly reduced repetition of names in seeded member and grievance data

**Code Changes:**
- `Constants.gs`: Added `GRIEVANCE_STATUS_PRIORITY` constant
- `HiddenSheets.gs`: Added `sortGrievanceLogByStatus()` function
- `HiddenSheets.gs`: Hooked sort into `onEditAutoSync()` trigger
- `ConsolidatedDashboard.gs`: Added duplicate `sortGrievanceLogByStatus()` function and trigger hook
- `SeedNuke.gs`: Expanded `firstNames` and `lastNames` arrays (20 → 120 each)
- `ConsolidatedDashboard.gs`: Expanded name arrays in `SEED_MEMBERS()` function

---

### Version 1.4.4 (2025-12-18) - Grievance Log Member Lookup Fix

**Grievance Log now auto-populates member data directly from Member Directory:**

| Column | Header | Auto-Synced From |
|--------|--------|------------------|
| C | First Name | Member Directory (by Member ID) |
| D | Last Name | Member Directory (by Member ID) |
| X | Member Email | Member Directory (by Member ID) |
| Y | Unit | Member Directory (by Member ID) |
| Z | Work Location | Member Directory (by Member ID) |
| AA | Steward | Member Directory (by Member ID) |

**Member ID Validation:**
- Column B (Member ID) uses dropdown validation
- Only Member IDs that exist in Member Directory are allowed
- Prevents orphan grievances with invalid member references

**Code Changes:**
- `HiddenSheets.gs`: Rewrote `syncGrievanceFormulasToLog()` to lookup member data directly from Member Directory instead of using hidden sheet formulas
- Bypassed the ARRAYFORMULA/FILTER issue that caused empty lookups
- Member lookup now uses `MEMBER_COLS` constants for reliable column access

---

### Version 1.4.3 (2025-12-18) - Member Directory Auto-Sync

**Member Directory Columns AB-AD now auto-sync from Grievance Log:**

| Column | Header | Auto-Synced Value |
|--------|--------|-------------------|
| AB | Has Open Grievance? | "Yes" or "No" based on active grievances |
| AC | Grievance Status | Status from most recent grievance |
| AD | Days to Deadline | Countdown number or "Overdue" |

**Code Changes:**
- `HiddenSheets.gs`: Changed `_Grievance_Calc` hidden sheet to pull `DAYS_TO_DEADLINE` instead of `NEXT_ACTION_DUE`
- `Constants.gs`: Updated header from "Next Deadline" to "Days to Deadline"
- Auto-sync trigger updates Member Directory when Grievance Log is edited

---

### Version 1.4.2 (2025-12-18) - Date Formatting & Overdue Display

**Formatting Changes:**
- Date format changed from `yyyy-mm-dd` to `dd-mm-yyyy` throughout
- Days Open (S) and Days to Deadline (U) now display as whole numbers (no decimals)
- Days to Deadline shows "Overdue" for past-due cases instead of negative numbers

**Display Values for Days to Deadline:**
| Value | Meaning |
|-------|---------|
| `18` | 18 days remaining |
| `0` | Due today |
| `Overdue` | Past deadline |
| *(blank)* | Case is closed |

**Code Changes:**
- `Code.gs`: Added `setNumberFormat('0')` for Days Open and Days to Deadline in `createGrievanceLog()`
- `Code.gs`: Changed date format to `dd-mm-yyyy` in `createGrievanceLog()`
- `HiddenSheets.gs`: Changed date format to `dd-mm-yyyy` in all sync functions
- `HiddenSheets.gs`: Days to Deadline now returns "Overdue" when `days < 0`
- `Code.gs`: Updated setup success message to confirm auto-sync trigger installation

---

### Version 1.4.1 (2025-12-16) - Days to Deadline Fix

**Bug Fix:**
- Fixed "Days to Deadline" and "Days Open" showing duplicate/incorrect values for all grievances
- Root cause: ARRAYFORMULA with FILTER-based row index in hidden sheet didn't expand correctly
- Solution: Calculate Days Open, Next Action Due, Days to Deadline, and all deadline dates directly in JavaScript within `syncGrievanceFormulasToLog()` function

**Code Changes:**
- `HiddenSheets.gs`: Rewrote metrics calculation in `syncGrievanceFormulasToLog()` (~60 lines added)
  - Now calculates Filing Deadline, Step I/II/III Due dates from source dates
  - Days Open = (Date Closed or Today) - Date Filed
  - Next Action Due = Based on Current Step status
  - Days to Deadline = Next Action Due - Today
  - All values now calculated per-row from actual grievance data

---

### Version 1.4.0 (2025-12-16) - Dashboard Views Added

**Major Updates:**

- Re-added dashboard sheets from original 509dashboard project
- Created unified 💼 Dashboard (merged Executive Dashboard + Dashboard themes)
- Added 🎯 Custom View with customizable metric selection
- Added `_Dashboard_Calc` hidden sheet with 15 self-healing metric formulas

**New Sheets (2):**

- `💼 Dashboard` - Executive-style metrics view with:
  - QUICK STATS section (green Union theme)
  - MEMBER METRICS section (blue theme)
  - GRIEVANCE METRICS section (orange theme)
  - TIMELINE & PERFORMANCE section (purple theme)
  - Real-time formulas linked to Member Directory and Grievance Log

- `🎯 Interactive` - Customizable dashboard with:
  - Dropdown metric selection (8 available metrics)
  - Time range filtering
  - Theme selection
  - Live-updating values

**New Hidden Sheet:**

- `_Dashboard_Calc` - 15 key metrics with self-healing formulas:
  - Total Members, Active Stewards
  - Total/Open/Pending/Settled/Won/Denied/Withdrawn Grievances
  - Win Rate %, Avg Days to Resolution
  - Overdue Cases, Due This Week
  - Filed This Month, Closed This Month

**Code Changes:**

- Constants.gs: Added DASHBOARD, INTERACTIVE, DASHBOARD_CALC to SHEETS
- Code.gs: Added `createDashboard()`, `createInteractiveDashboard()` (~300 lines)
- HiddenSheets.gs: Added `setupDashboardCalcSheet()` (~70 lines)
- Updated CREATE_509_DASHBOARD, DIAGNOSE_SETUP, REPAIR_DASHBOARD for 5 sheets

---

### Version 1.1.0 (2025-12-14) - Hidden Sheet Architecture

Added self-healing hidden formula system:

- HiddenSheets.gs: Full implementation of 6 hidden calculation sheets
- Auto-sync onEdit trigger for automatic cross-sheet data population
- Self-healing repair functions
- Manual sync options in Administrator menu

### Version 1.0.0 (2025-12-14) - Fresh Start

Complete rebuild from AIR.md specification.

- Removed 77 legacy .gs files (~118,000 lines)
- Created 3 clean files (~1,500 lines)
- Constants.gs: Column mappings and helper functions
- Code.gs: Setup, menus, and sheet creation
- SeedNuke.gs: Demo data management
- 98.8% code reduction while maintaining core functionality
