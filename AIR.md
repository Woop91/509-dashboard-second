# 509 Dashboard - Architecture & Implementation Reference

**Version:** 1.5.6 (View Controls, Steward Alerts & Audit Logging)
**Last Updated:** 2026-01-02
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
2. Run `CREATE_509_DASHBOARD()` to create 5 sheets + 5 hidden calculation sheets
3. Use `Demo > Seed All Sample Data` to populate test data
4. Customize Config sheet with your organization's values

---

## File Architecture

### Project Structure (10 Files)

```
509-dashboard/
├── Constants.gs           # Configuration constants (SHEETS, COLORS, MEMBER_COLS, GRIEVANCE_COLS)
├── Code.gs                # Main entry point, setup functions, sheet creation
├── ConsolidatedDashboard.gs # Complete standalone version (all functionality in one file)
├── SeedNuke.gs            # Demo data seeding and clearing functions
├── HiddenSheets.gs        # Self-healing hidden calculation sheets with auto-sync
├── ADHDFeatures.gs        # ADHD accessibility & theming (focus mode, themes, pomodoro)
├── TestingValidation.gs   # Test framework & data validation
├── PerformanceUndo.gs     # Caching layer & undo/redo system
├── MobileQuickActions.gs  # Mobile interface & quick actions menu
├── WebApp.gs              # Web app deployment for mobile URL access
└── AIR.md                 # This document
```

### File Descriptions

**Constants.gs** (~600 lines)
- `SHEETS` - Sheet name constants (3 data + 2 dashboard + 5 hidden + 2 optional + 2 auto-generated)
- `COLORS` - Brand color scheme
- `MEMBER_COLS` - 31 Member Directory column positions
- `GRIEVANCE_COLS` - 34 Grievance Log column positions
- `CONFIG_COLS` - Config sheet column positions
- `DEFAULT_CONFIG` - Default dropdown values
- `MULTI_SELECT_COLS` - Configuration for multi-select columns
- `JOB_METADATA_FIELDS` - Maps Member Directory fields to Config columns for bidirectional sync:
  | Member Directory | Config Column |
  |------------------|---------------|
  | Job Title | Job Titles |
  | Work Location | Office Locations |
  | Unit | Units |
  | Supervisor | Supervisors |
  | Manager | Managers |
  | Assigned Steward | Stewards |
  | Committees | Steward Committees |
  | Home Town | Home Towns |
- `getJobMetadataField(label)` - Get field config by label
- `getJobMetadataByMemberCol(col)` - Get field config by column number
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

**Code.gs** (~3700 lines)
- `onOpen()` - Creates menu system (10 menus including View)
- `CREATE_509_DASHBOARD()` - Main setup function (creates 5 sheets + 5 hidden)
- `DIAGNOSE_SETUP()` - System health check
- `REPAIR_DASHBOARD()` - Repair hidden sheets, triggers, and auto-create Menu Checklist
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
- `searchMembers()` - Search members dialog
- `startNewGrievance()` - Start grievance dialog
- `viewActiveGrievances()` - Navigate to Grievance Log
- `createMenuChecklistSheet_()` - Auto-create Menu Checklist with 57 items in 13 testing phases
- Sheet creation (5 functions): `createConfigSheet()`, `createMemberDirectory()`, `createGrievanceLog()`, `createDashboard()`, `createInteractiveDashboard()`
- View Controls: `simplifyTimelineView()`, `showFullTimelineView()`, `setupTimelineColumnGroups()`, `applyStepHighlighting()`, `freezeKeyColumns()`, `unfreezeAllColumns()`
- Steward Alerts: `sendStewardDeadlineAlerts()`, `sendStewardAlertsNow()`, `configureAlertSettings()`
- Audit Logging: `setupAuditLogSheet()`, `logAuditEvent()`, `onEditAudit()`, `installAuditTrigger()`, `removeAuditTrigger()`, `viewAuditLog()`, `clearOldAuditEntries()`, `getAuditHistory()`

**SeedNuke.gs** (~1200 lines)
- `SEED_SAMPLE_DATA()` - Seeds Config + 50 members + 15 grievances (30%)
- `seedConfigData()` - Populate Config dropdowns
- `SEED_MEMBERS(count, grievancePercent)` - Seed N members with optional grievances (default 30%)
- `SEED_GRIEVANCES(count)` - Seed N grievances for existing members (max 300)
- `SEED_MEMBERS_DIALOG()` - Prompt for count (30% grievances auto-created)
- `SEED_MEMBERS_ADVANCED_DIALOG()` - Prompt for count AND grievance percentage
- `SEED_GRIEVANCES_DIALOG()` - Prompt for grievance count (existing members)
- `seed50Members()` - Shortcut: 50 members with 30% grievances
- `seed100MembersWithGrievances()` - Shortcut: 100 members with 50% grievances
- `seed25Grievances()` - Shortcut: seed 25 grievances for existing members
- `generateSingleMemberRow()` - Generate one member row (31 columns)
- `generateSingleGrievanceRow()` - Generate one grievance row (34 columns)
- `NUKE_ALL_DATA()` - Clear all data with confirmation
- `NUKE_CONFIG_DROPDOWNS()` - Clear only Config dropdowns
- `getConfigValues()` - Helper: get values from Config column
- `randomChoice()` - Helper: pick random array element
- `randomDate()` - Helper: generate random date
- `addDays()` - Helper: add days to date

**HiddenSheets.gs** (~2000 lines)
- `setupAllHiddenSheets()` - Create all 5 hidden calculation sheets
- Hidden Sheet Setup Functions (5 total):
  - `setupGrievanceCalcSheet()` - Grievance timeline formulas (auto-calc deadlines)
  - `setupGrievanceFormulasSheet()` - Member lookup formulas (First Name, Last Name, Email, etc.)
  - `setupMemberLookupSheet()` - Member → Grievance Log sync
  - `setupStewardContactCalcSheet()` - Steward contact tracking
  - `setupDashboardCalcSheet()` - Dashboard summary metrics (15 key metrics)
- Sync Functions:
  - `syncAllData()` - Sync all cross-sheet data with data quality validation
  - `syncGrievanceToMemberDirectory()` - Sync grievance data to members (AB-AD)
  - `syncMemberToGrievanceLog()` - Sync member data to grievances
  - `syncGrievanceFormulasToLog()` - Sync timeline formulas to Grievance Log
  - `syncNewValueToConfig(e)` - Bidirectional sync: adds new values from Member Directory to Config
  - `sortGrievanceLogByStatus()` - Auto-sort by status priority and deadline urgency
- Trigger & Repair Functions:
  - `onEditAutoSync()` - Auto-sync trigger handler
  - `installAutoSyncTrigger()` - Interactive dialog with sync frequency and sheet options
  - `installAutoSyncTriggerWithOptions(options)` - Install with specific options
  - `getAutoSyncOptions()` - Get current auto-sync configuration
  - `installAutoSyncTriggerQuick()` - Quick install with default settings
  - `removeAutoSyncTrigger()` - Remove the onEdit trigger
  - `repairAllHiddenSheets()` - Self-healing repair function
  - `repairGrievanceCheckboxes()` - Re-apply checkboxes to Grievance Log AC column
  - `repairMemberCheckboxes()` - Re-apply checkboxes to Member Directory AE column
  - `verifyHiddenSheets()` - Verification and diagnostics
  - `refreshAllHiddenFormulas()` - Force recalculation and sync
- Data Quality Functions:
  - `checkDataQuality()` - Validate data integrity (missing member IDs, orphan grievances)
  - `fixDataQualityIssues()` - Auto-fix common data quality issues
  - `showGrievancesWithMissingMemberIds()` - Navigate to and highlight grievances missing member IDs

**ADHDFeatures.gs** (~350 lines) - ADHD Accessibility & Theming
- `showADHDControlPanel()` - Main ADHD settings panel
- `getADHDSettings()`, `saveADHDSettings()`, `resetADHDSettings()` - Settings management
- `applyADHDSettings()` - Apply visual settings
- `activateFocusMode()`, `deactivateFocusMode()` - Focus mode (hide non-essential sheets, shows comprehensive documentation)
- `toggleZebraStripes()`, `applyZebraStripes()`, `removeZebraStripes()` - Row banding
- `toggleGridlinesADHD()`, `hideAllGridlines()`, `showAllGridlines()` - Gridline control
- `toggleReducedMotion()` - Animation preferences
- `showQuickCaptureNotepad()` - Quick note-taking dialog
- `startPomodoroTimer()` - Built-in pomodoro timer
- `setBreakReminders()`, `showBreakReminder()` - Break notification system
- `showThemeManager()` - Theme selection UI
- `applyTheme()`, `applyThemeToSheet()`, `previewTheme()` - Theme application
- `getCurrentTheme()`, `resetToDefaultTheme()`, `quickToggleDarkMode()` - Theme utilities
- `setupADHDDefaults()` - Interactive dialog with customizable options (gridlines, zebra, font size, focus mode)
- `applyADHDDefaultsWithOptions(options)` - Apply ADHD defaults with specific options
- `undoADHDDefaults()` - Revert ADHD-friendly settings to defaults

**ConsolidatedDashboard.gs** (~7000 lines) - Complete Standalone Version
- Contains ALL functionality from Code.gs, SeedNuke.gs, HiddenSheets.gs, etc. in one file
- Intended for users who want to deploy without multiple file dependencies
- **Enhanced Seeding**: `SEED_MEMBERS(count, grievancePercent)` - combined member + grievance seeding
  - `SEED_MEMBERS_DIALOG()` - Prompt for count (30% grievances auto-created)
  - `SEED_MEMBERS_ADVANCED_DIALOG()` - Prompt for count AND grievance percentage
  - `seed50Members()` - 50 members with 30% grievances
  - `seed100MembersWithGrievances()` - 100 members with 50% grievances
  - `SEED_GRIEVANCES(count)` - Seed grievances for existing members only
- Includes `createMenuChecklistSheet_()` for auto-creating Menu Checklist on REPAIR_DASHBOARD

**WebApp.gs** (~500 lines) - Web App Deployment for Mobile Access
- `doGet(e)` - Web app entry point, serves mobile dashboard
- `getWebAppDashboardHtml()` - Main dashboard HTML
- `getWebAppSearchHtml()` - Search interface HTML
- `getWebAppGrievanceListHtml()` - Grievance list HTML
- `showWebAppUrl()` - Display the deployed web app URL
- Enables direct mobile access via URL without opening the spreadsheet

**Drive/Calendar/Notifications** (Code.gs & ConsolidatedDashboard.gs)
- **Google Drive Integration**:
  - `setupDriveFolderForGrievance()` - Create folder for selected grievance (shows row selection message if no row selected)
  - `showGrievanceFiles()` - View files in grievance folder (shows row selection message if no row selected)
  - `batchCreateGrievanceFolders()` - Create folders for all grievances (shows confirmation dialog)
  - `getOrCreateDashboardFolder_()` - Get/create root "509 Dashboard - Grievance Files" folder
- **Calendar Sync**:
  - `syncDeadlinesToCalendar()` - Create all-day events with rate limiting (100ms throttle, batched processing)
  - `showUpcomingDeadlinesFromCalendar()` - Show next 7 days with member names (e.g., "Step I Due (John Smith)")
  - `clearAllCalendarEvents()` - Remove all grievance calendar events
  - `buildGrievanceMemberLookup()` - Helper function for member name lookup
- **Email Notifications**:
  - `showNotificationSettings()` - Enable/disable daily 8 AM deadline emails
  - `testDeadlineNotifications()` - Send test email to verify setup
  - `checkDeadlinesAndNotify_()` - Daily trigger function (notifies if due within 3 days)
  - `installDailyTrigger_()`, `removeDailyTrigger_()` - Trigger management

**TestingValidation.gs** (~500 lines) - Testing Framework & Data Validation
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

**PerformanceUndo.gs** (~300 lines) - Caching Layer
- Caching (exposed in menu):
  - `getCachedData()` - Get data from cache or load
  - `setCachedData()` - Store data in cache
  - `invalidateCache()` - Clear specific cache
  - `invalidateAllCaches()` - Clear all caches
  - `warmUpCaches()` - Pre-populate caches (metadata only, see note below)
  - `getCachedGrievances()`, `getCachedMembers()`, `getCachedStewards()` - Cached data getters
  - `getCachedDashboardMetrics()` - Cached dashboard metrics
  - `showCacheStatusDashboard()` - Cache status UI
  - **Cache Keys**: `memberMeta`, `grievanceMeta`, `configData`
  - **Note**: CacheService has 100KB limit per item. Full data caching is avoided; only metadata (row/column counts, timestamps) is cached.
- Undo/Redo (NOT in menu - use Google Sheets built-in Ctrl+Z/Ctrl+Y):
  - Functions exist but are not exposed in menu since Google Sheets has robust built-in undo/redo

**MobileQuickActions.gs** (~1200 lines) - Mobile Interface & Quick Actions
- Mobile Interface:
  - `showMobileDashboard()` - Touch-optimized dashboard
  - `getMobileDashboardStats()` - Dashboard statistics
  - `getRecentGrievancesForMobile()` - Recent grievances (limit)
  - `showMobileGrievanceList()` - Mobile grievance list
  - `showMobileUnifiedSearch()` - Mobile search UI
  - `getMobileSearchData()` - Search handler
  - `showMyAssignedGrievances()` - View user's assigned cases
- Quick Actions:
  - `showQuickActionsMenu()` - Context-aware quick actions with detailed help (explains supported sheets and available actions)
  - `showMemberQuickActions()` - Quick actions for member row
  - `showGrievanceQuickActions()` - Quick actions for grievance row
  - `quickUpdateGrievanceStatus()` - One-click status update
  - `composeEmailForMember()` - Email composition dialog
  - `sendQuickEmail()` - Send email via MailApp
  - `showMemberGrievanceHistory()` - Member's grievance history
  - `openGrievanceFormForMember()` - Start grievance for member
  - `syncSingleGrievanceToCalendar()` - Sync single grievance to calendar

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

---

## Sheet Structure (5+ Visible + 5 Hidden)

### Core Data Sheets

| # | Sheet Name | Type | Purpose |
|---|------------|------|---------|
| 1 | Config | Data | Master dropdown lists for validation (43 columns) |
| 2 | Member Directory | Data | All member data (31 columns) |
| 3 | Grievance Log | Data | All grievance cases (34 columns) |

### Dashboard Sheets

| # | Sheet Name | Type | Purpose |
|---|------------|------|---------|
| 4 | 💼 Dashboard | View | Executive metrics dashboard with 9 analytics sections |
| 5 | 🎯 Interactive | View | Customizable metrics with dropdowns |

### Auto-Generated Sheets

| # | Sheet Name | Type | Purpose |
|---|------------|------|---------|
| 6 | Menu Checklist | Reference | Checklist of all 57 menu items organized by testing phase (auto-created on REPAIR_DASHBOARD) |
| 7 | Test Results | Output | Test framework results from runAllTests() |

### Optional Source Sheets (not auto-created)

| # | Sheet Name | Purpose |
|---|------------|---------|
| - | 📅 Meeting Attendance | Track member meeting attendance |
| - | 🤝 Volunteer Hours | Track volunteer hour contributions |

#### 💼 Dashboard - 9 Live Analytics Sections

| # | Section | Color | Metrics | Data Source |
|---|---------|-------|---------|-------------|
| 1 | QUICK STATS | 🟢 Green | Total Members, Active Stewards, Active Grievances, Win Rate, Overdue, Due This Week | `_Dashboard_Calc` |
| 2 | MEMBER METRICS | 🔵 Blue | Total Members, Active Stewards, Avg Open Rate, YTD Vol Hours | Member Directory |
| 3 | GRIEVANCE METRICS | 🟠 Orange | Open, Pending Info, Settled, Won, Denied, Withdrawn | Grievance Log |
| 4 | TIMELINE & PERFORMANCE | 🟣 Purple | Avg Days Open, Filed This Month, Closed This Month, Avg Resolution | Grievance Log |
| 5 | TYPE ANALYSIS | 🔷 Indigo | 5 issue categories × (Total, Open, Resolved, Win Rate, Avg Days) | Grievance Log |
| 6 | LOCATION BREAKDOWN | 🔵 Cyan | 5 locations × (Members, Grievances, Open Cases, Win Rate) | Config + Member Dir + Grievance Log |
| 7 | STEWARD PERFORMANCE | 🟣 Purple | Total Stewards, Active w/Cases, Avg Cases, Vol Hours, Contacts | Member Dir + Grievance Log |
| 8 | MONTH-OVER-MONTH TRENDS | 🔴 Red | Filed/Closed/Won × (This Month, Last Month, Change, % Change, Trend) | Grievance Log |
| 9 | STATUS LEGEND | ⬜ Gray | Color/icon reference guide | Static |

**All sections use live COUNTIF/COUNTIFS/AVERAGEIFS formulas that auto-update when source data changes.**

### Hidden Calculation Sheets

| # | Sheet Name | Purpose |
|---|------------|---------|
| 1 | `_Grievance_Calc` | Grievance → Member Directory sync (AB-AD) |
| 2 | `_Grievance_Formulas` | Member → Grievance Log sync (C-D, X-AA) |
| 3 | `_Member_Lookup` | Member data lookup formulas |
| 4 | `_Steward_Contact_Calc` | Steward contact tracking (Y-AA) |
| 5 | `_Dashboard_Calc` | Dashboard summary metrics (15 key KPIs) |

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
├── 🎯 Interactive Dashboard
├── ────────────────
├── 📋 View Active Grievances
├── 📱 Mobile Dashboard
├── 📱 Get Mobile App URL
├── ⚡ Quick Actions
├── ────────────────
└── 📋 Grievance Tools
    ├── ➕ Start New Grievance
    ├── 🔄 Refresh Grievance Formulas
    ├── 🔄 Refresh Member Directory Data
    ├── ────────────────
    ├── 🔗 Setup Live Grievance Links
    ├── 👤 Setup Member ID Dropdown
    └── 🔧 Fix Overdue Text Data

🔍 Search
└── 🔍 Search Members

👁️ View
├── 📅 Simplify Timeline (Hide Steps)
├── 📅 Show Full Timeline
├── ────────────────
├── 🎨 Apply Step Highlighting
├── 🔲 Setup Column Groups
├── ────────────────
├── ❄️ Freeze Key Columns
└── 🔓 Unfreeze All Columns

📊 Sheet Manager
├── 📊 Rebuild Dashboard
├── 📈 Refresh Interactive Charts
├── 🔄 Refresh All Formulas
├── ────────────────
├── 📁 Google Drive
│   ├── 📁 Setup Folder for Grievance
│   ├── 📁 View Grievance Files
│   └── 📁 Batch Create Folders
├── 📅 Calendar
│   ├── 📅 Sync Deadlines to Calendar
│   ├── 📅 View Upcoming Deadlines
│   └── 🗑️ Clear Calendar Events
└── 📬 Notifications
    ├── ⚙️ Notification Settings
    ├── ⚙️ Alert Settings
    ├── ────────────────
    ├── 📧 Send Steward Alerts Now
    └── 🧪 Test Notifications

🔧 Tools
├── ♿ ADHD & Accessibility
│   ├── ♿ ADHD Control Panel
│   ├── 🎯 Focus Mode
│   ├── 🔲 Toggle Zebra Stripes
│   ├── 📝 Quick Capture
│   └── 🍅 Pomodoro Timer
├── 🎨 Theming
│   ├── 🎨 Theme Manager
│   ├── 🌙 Toggle Dark Mode
│   └── 🔄 Reset Theme
├── ────────────────
├── ☑️ Multi-Select
│   ├── 📝 Open Editor
│   ├── ⚡ Enable Auto-Open
│   └── 🚫 Disable Auto-Open
├── ────────────────
├── 🗄️ Cache & Performance
│   ├── 🗄️ Cache Status
│   ├── 🔥 Warm Up Caches
│   └── 🗑️ Clear All Caches
├── ────────────────
└── ✅ Validation
    ├── 🔍 Run Bulk Validation
    ├── ⚙️ Validation Settings
    ├── 🧹 Clear Indicators
    └── ⚡ Install Validation Trigger

🏗️ Setup
├── 🔧 REPAIR DASHBOARD
├── ────────────────
├── ⚙️ Setup Data Validations
├── 🎨 Setup ADHD Defaults
└── ↩️ Undo ADHD Defaults

🎭 Demo (hidden if demo mode disabled)
├── 🚀 Seed All Sample Data
├── ────────────────
├── 🌱 Seed Data
│   ├── ⚙️ Seed Config Dropdowns Only
│   ├── 👥 Seed Members & Grievances (Custom)
│   ├── 👥 Seed Members (Advanced - Set % Grievances)
│   ├── 👥 Seed 50 Members (30% Grievances)
│   ├── 👥 Seed 100 Members (50% Grievances)
│   └── 📋 Seed Grievances Only (existing members)
├── ────────────────
└── 🗑️ Nuke Data
    ├── ☢️ NUKE SEEDED DATA
    ├── 🧹 Clear Config Dropdowns Only
    └── 🔄 Restore Config & Dropdowns

🧪 Testing
├── 🧪 Run All Tests
├── ⚡ Run Quick Tests
├── ────────────────
└── 📊 View Test Results

⚙️ Administrator
├── 🔍 DIAGNOSE SETUP
├── 🔍 Verify Hidden Sheets
├── ────────────────
├── 🔧 Setup & Triggers
│   ├── 🔧 Setup All Hidden Sheets
│   ├── 🔧 Repair All Hidden Sheets
│   ├── ⚡ Install Auto-Sync Trigger
│   └── 🚫 Remove Auto-Sync Trigger
├── ────────────────
├── 🩺 Data Quality
│   ├── 🔍 Check Data Quality
│   └── 📋 View Missing Member IDs
├── ────────────────
├── 🔄 Manual Sync
│   ├── 🔄 Sync All Data Now
│   ├── 🔄 Sync Grievance → Members
│   └── 🔄 Sync Members → Grievances
├── ────────────────
└── 📋 Audit Log
    ├── 📋 View Audit Log
    ├── 🔧 Setup Audit Log
    ├── ────────────────
    ├── ⚡ Enable Audit Tracking
    ├── 🚫 Disable Audit Tracking
    ├── ────────────────
    └── 🗑️ Clear Old Entries (30+ days)
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

## Live Formula Architecture

### Member Directory Grievance Columns (Static Values)

The Member Directory grievance columns (AB-AD) are populated with **STATIC VALUES** by the sync function - **NO FORMULAS in visible sheets**:

| Member Column | Header | Source | Updated By |
|---------------|--------|--------|------------|
| `MEMBER_COLS.HAS_OPEN_GRIEVANCE` | Has Open Grievance? | `_Grievance_Calc` hidden sheet | `syncGrievanceToMemberDirectory()` |
| `MEMBER_COLS.GRIEVANCE_STATUS` | Grievance Status | `_Grievance_Calc` hidden sheet | `syncGrievanceToMemberDirectory()` |
| `MEMBER_COLS.DAYS_TO_DEADLINE` | Days to Deadline | `_Grievance_Calc` hidden sheet | `syncGrievanceToMemberDirectory()` |

**Architecture:** Formulas live in hidden `_Grievance_Calc` sheet. The sync function reads calculated values and writes them as static values to Member Directory. This prevents #REF! errors and keeps visible sheets formula-free.

### Syncing Grievance Data

Run `syncGrievanceToMemberDirectory()` or use **Administrator > Manual Sync > Sync Grievance → Members** to update columns AB-AD.

---

## Hidden Sheet Architecture (Self-Healing)

The system uses 5 hidden calculation sheets with auto-sync triggers for cross-sheet data population. Formulas are stored in hidden sheets and synced to visible sheets, making them **self-healing** - if formulas are accidentally deleted, running REPAIR_DASHBOARD() restores them.

### Hidden Sheets (5 total)

| Sheet | Source | Destination | Purpose |
|-------|--------|-------------|---------|
| `_Grievance_Calc` | Grievance Log | Member Directory | Backup calculations for AB-AD |
| `_Grievance_Formulas` | Member Directory | Grievance Log | C-D (Name), H-P (Timeline), S-U (Days Open, Next Action, Days to Deadline), X-AA (Contact) |
| `_Member_Lookup` | Member Directory | Grievance Log | Member data lookup |
| `_Steward_Contact_Calc` | Member Directory | Contact Reports | Y-AA (Contact tracking) |
| `_Dashboard_Calc` | Both | 💼 Dashboard | 15 summary metrics (Win Rate, Overdue, Due This Week, etc.) |

### Auto-Sync Trigger

The `onEditAutoSync` trigger automatically syncs data when:
- Grievance Log is edited → Updates Member Directory columns AB-AD, then auto-sorts by status
- Member Directory is edited → Updates Grievance Log columns C-D, X-AA

### Key Functions (HiddenSheets.gs)

| Function | Purpose |
|----------|---------|
| `setupAllHiddenSheets()` | Create all 5 hidden sheets with formulas |
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

### Version 1.5.6 (2026-01-02) - View Controls, Steward Alerts & Audit Logging

**View Menu (New):**

- `simplifyTimelineView()` - Hide intermediate step columns (H, J-Q) to reduce visual clutter
- `showFullTimelineView()` - Restore all timeline columns
- `setupTimelineColumnGroups()` - Create collapsible column groups for steps
- `applyStepHighlighting()` - Conditional formatting: gray for inactive steps, orange/red for urgent deadlines
- `freezeKeyColumns()` / `unfreezeAllColumns()` - Column freezing controls

**Per-Steward Email Alerts (New):**

- `sendStewardDeadlineAlerts()` - Daily digest grouped by steward with overdue/urgent/upcoming categories
- `sendStewardAlertsNow()` - Manual trigger to send alerts immediately
- `configureAlertSettings()` - Configure alert window (1-30 days) and enable/disable per-steward mode

**Audit Logging for Multi-Steward Accountability (New):**

- `setupAuditLogSheet()` - Creates hidden `_Audit_Log` sheet
- `logAuditEvent()` - Logs changes with timestamp, user email, field, old/new values
- `onEditAudit()` - Trigger to automatically log edits to Member Directory and Grievance Log
- `installAuditTrigger()` / `removeAuditTrigger()` - Enable/disable audit tracking
- `viewAuditLog()` - View and sort audit entries
- `clearOldAuditEntries()` - Clean entries older than 30 days
- `getAuditHistory(recordId)` - Get change history for a specific record

**Code Quality:**

- Removed duplicate function definitions across modules
- Regenerated ConsolidatedDashboard.gs via build.js for clean build

### Version 1.5.5 (2026-01-02) - Enhanced Customization & Data Quality

**ADHDFeatures.gs Enhancements:**

- `setupADHDDefaults()` - Now shows interactive dialog with customizable options:
  - Checkbox options for gridlines, zebra stripes, font size, and focus mode
  - Users can select which features to apply
- `applyADHDDefaultsWithOptions(options)` - New function to apply specific ADHD settings
- `undoADHDDefaults()` - New function to revert ADHD-friendly settings to defaults
- `activateFocusMode()` - Enhanced with comprehensive documentation explaining what it does

**HiddenSheets.gs Enhancements:**

- `installAutoSyncTrigger()` - Now shows interactive dialog with customization options:
  - Sync frequency selection (Instant, 5 min, 15 min, Manual only)
  - Sheet selection (sync Grievance Log, Member Directory, or both)
- `installAutoSyncTriggerWithOptions(options)` - New function to install with specific options
- `getAutoSyncOptions()` - New function to get current auto-sync configuration
- `installAutoSyncTriggerQuick()` - New function for quick install with default settings
- `syncAllData()` - Now includes data quality validation and shows warnings for issues

**Data Quality System (New):**

- `checkDataQuality()` - Validates data integrity (missing member IDs, orphan grievances)
- `fixDataQualityIssues()` - Auto-fix common data quality issues
- `showGrievancesWithMissingMemberIds()` - Navigate to and highlight grievances without member IDs
- Added "🩺 Data Quality" submenu to Administrator menu

**Calendar Integration Improvements:**

- `syncDeadlinesToCalendar()` - Added rate limiting (100ms throttle between events)
  - Batch processing for large datasets
  - Graceful handling of API rate limit errors
- `showUpcomingDeadlinesFromCalendar()` - Now shows member names with each deadline
  - Format: "• 01/15: Grievance GR-001 - Step I Due (John Smith)"
- `buildGrievanceMemberLookup()` - New helper function for member name lookup

**MobileQuickActions.gs Enhancements:**

- `showQuickActionsMenu()` - Enhanced with detailed help when no row is selected
  - Explains which sheets support quick actions (Member Directory, Grievance Log)
  - Lists available actions for each sheet type

**Menu Updates:**

- Added "↩️ Undo ADHD Defaults" to Setup menu
- Added "🩺 Data Quality" submenu to Administrator menu with:
  - Check Data Quality
  - Fix Missing Member IDs
  - Show Grievances Missing IDs

**Build System:**

- Fixed build.js to only include existing modules (was listing 80+ non-existent files)
- CORE_MODULES now correctly lists the 9 actual .gs files

---

### Version 1.5.4 (2026-01-01) - Unified Seeding & AIR Sync

**Unified Seeding Architecture:**

All seeding now uses combined member + grievance approach across all files:

- `SEED_MEMBERS(count, grievancePercent)` - Seeds members with optional grievances (default 30%)
- `SEED_MEMBERS_DIALOG()` - Prompts for count, auto-creates 30% grievances
- `SEED_MEMBERS_ADVANCED_DIALOG()` - Prompts for both count AND grievance percentage
- `seed50Members()` - 50 members with 30% grievances (~15 grievances)
- `seed100MembersWithGrievances()` - 100 members with 50% grievances (~50 grievances)
- `SEED_GRIEVANCES_DIALOG()` - Seeds grievances for existing members only

**Demo Menu Updated:**

```
🎭 Demo > 🌱 Seed Data
├── ⚙️ Seed Config Dropdowns Only
├── 👥 Seed Members & Grievances (Custom)
├── 👥 Seed Members (Advanced - Set % Grievances)
├── 👥 Seed 50 Members (30% Grievances)
├── 👥 Seed 100 Members (50% Grievances)
└── 📋 Seed Grievances Only (existing members)
```

**Parity Achieved:**

- Code.gs, SeedNuke.gs, and ConsolidatedDashboard.gs now have identical seeding functionality
- All three files use the same Demo menu structure
- ConsolidatedDashboard.gs remains a complete standalone copy

**AIR.md Updates:**

- Updated all file line counts to current values
- Fixed Quick Start to show 10 .gs files
- Documented unified seeding architecture
- Date format confirmed as MM/dd/yyyy

---

### Version 1.5.3 (2025-12-31) - Drive/Calendar/Notifications Implemented

**New Features:**

1. **Google Drive Integration** (3 functions):
   - `setupDriveFolderForGrievance()` - Creates folder for selected grievance
   - `showGrievanceFiles()` - Lists files in grievance folder
   - `batchCreateGrievanceFolders()` - Batch creates folders for all grievances
   - Root folder: "509 Dashboard - Grievance Files"

2. **Calendar Integration** (3 functions):
   - `syncDeadlinesToCalendar()` - Creates all-day events for upcoming deadlines
   - `showUpcomingDeadlinesFromCalendar()` - Shows next 7 days of grievance events
   - `clearAllCalendarEvents()` - Removes all grievance-related calendar events

3. **Email Notifications** (2 functions):
   - `showNotificationSettings()` - Enable/disable daily deadline reminders at 8 AM
   - `testDeadlineNotifications()` - Send test email to verify setup
   - Daily trigger notifies when grievances are due within 3 days

4. **Navigation Functions** (3 functions):
   - `showInteractiveDashboardTab()` - Navigate to Interactive Dashboard
   - `refreshInteractiveCharts()` - Refresh Interactive Dashboard data
   - `showWebAppUrl()` - Show instructions for mobile web app setup

**Fixes:**

1. **installMultiSelectTrigger** - Changed to show manual setup instructions (onSelectionChange triggers can't be created programmatically)

2. **Removed Undo/Redo from menu** - Was only showing help text stubs; use Google Sheets built-in Ctrl+Z/Ctrl+Y

**Menu Changes:**
- Menu Checklist now has **57 items** (removed 4 Undo/Redo items)
- Drive/Calendar/Notifications now fully functional

---

### Version 1.5.2 (2025-12-31) - Menu Checklist & Static Values Fix

**New Features:**

1. **Menu Checklist Sheet**: Auto-created on REPAIR_DASHBOARD
   - 57 menu items organized into 13 testing phases
   - Includes Phase, Menu, Item, Function Name, and Tested? checkbox columns
   - Optimal order for systematic testing of all dashboard functionality

2. **WebApp.gs**: New file for mobile web app deployment
   - Enables direct URL access to dashboard on mobile devices
   - Three views: Dashboard, Search, Grievances

**Bug Fixes:**

1. **#REF! errors in Member Directory columns AB-AD**
   - Root cause: ARRAYFORMULA was inserted directly into visible sheets
   - Fix: Removed formula insertions, now uses `syncGrievanceToMemberDirectory()` with static values only
   - Architecture: Formulas in hidden `_Grievance_Calc` sheet → sync function reads values → writes static values to visible sheet

**Documentation Updates:**

- Updated file list from 9 to 10 files (added WebApp.gs, ConsolidatedDashboard.gs)
- Removed DriveCalendarEmail.gs reference (file doesn't exist)
- Updated Menu System documentation to match current menu structure
- Added Testing menu documentation
- Added auto-generated sheets (Menu Checklist, Test Results) to sheet structure

---

### Version 1.5.1 (2025-12-30) - Bidirectional Config Sync & Bug Fixes

**New Feature: Bidirectional Sync between Member Directory and Config**

When users enter new values in Member Directory job metadata fields, those values are automatically added to the Config sheet dropdown lists. This works for both seeded data and manual user edits.

**JOB_METADATA_FIELDS Mapping (8 fields):**

| Member Directory Column | Config Column | Auto-Syncs |
|------------------------|---------------|------------|
| Job Title (D) | Job Titles (A) | ✓ |
| Work Location (E) | Office Locations (B) | ✓ |
| Unit (F) | Units (C) | ✓ |
| Supervisor (L) | Supervisors (F) | ✓ |
| Manager (M) | Managers (G) | ✓ |
| Assigned Steward (P) | Stewards (H) | ✓ |
| Committees (O) | Steward Committees (I) | ✓ |
| Home Town (X) | Home Towns (AF) | ✓ |

**Bug Fixes:**

1. **Days to Deadline showing 0** - Fixed sync order: `syncGrievanceFormulasToLog()` now runs BEFORE `syncGrievanceToMemberDirectory()` so calculated values exist before being read
2. **isClosed logic incomplete** - Fixed to include all closed statuses (Settled, Withdrawn, Denied, Won, Closed) instead of just 'Closed'
3. **Date format consistency** - Applied `MM/dd/yyyy` format to ALL tabs including Member Directory, Dashboard displays, Mobile views, and Search results

**Menu Changes:**

- Removed 'CREATE 509 DASHBOARD' from Setup menu
- Moved 'Search Members' to standalone '🔍 Search' menu for quick access
- Added 'Restore Config Dropdowns' option to Nuke Data submenu

**Code Changes:**

- `Constants.gs`: Added `JOB_METADATA_FIELDS` array with 8 field mappings, added helper functions `getJobMetadataField()` and `getJobMetadataByMemberCol()`
- `HiddenSheets.gs`: Added `syncNewValueToConfig()` function for bidirectional sync, simplified to use `getJobMetadataByMemberCol()` helper
- `ConsolidatedDashboard.gs`: Fixed sync order in `SEED_MEMBERS()`, fixed `isClosed` logic in `generateSingleGrievanceRow()`
- `Code.gs` / `ConsolidatedDashboard.gs`: Added MM/dd/yyyy format to Member Directory date columns, updated menu structure
- `MobileQuickActions.gs`: Changed date format to MM/dd/yyyy

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
- Added 🎯 Interactive Dashboard with customizable metric selection
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

### Version 1.3.0 (2025-12-16) - Simplified Core Architecture

**Major Updates:**

- Simplified from 22 sheets to 3 core sheets (Config, Member Directory, Grievance Log)
- Reduced hidden calculation sheets from 13 to 4 essential sheets
- Removed 18 non-essential sheets to focus on core grievance tracking functionality

**Sheets Removed:**

- Dashboard sheets: Dashboard, Executive Dashboard, KPI Performance Dashboard, Interactive
- Analytics sheets: Steward Workload, Trends & Timeline, Location Analytics, Type Analysis, Member Engagement, Cost Impact
- Utility sheets: Audit_Log, Diagnostics, Archive, User Settings, Member Satisfaction, Feedback & Development
- Help sheets: FAQ, Getting Started

**Hidden Sheets Retained (4):**

- `_Grievance_Calc` - Grievance → Member Directory sync (AB-AD)
- `_Grievance_Formulas` - Member → Grievance Log sync (C-D, X-AA)
- `_Member_Lookup` - Member data lookup formulas
- `_Steward_Contact_Calc` - Steward contact tracking (Y-AA)

**Code Changes:**

- Code.gs: Removed 19 sheet creation functions, reduced from ~1000 to ~600 lines
- HiddenSheets.gs: Reduced from ~1600 to ~800 lines, removed 9 hidden sheet setup functions
- Constants.gs: SHEETS object reduced from 22+ entries to 7 (3 core + 4 hidden)
- Menu updated to remove references to deleted sheets

---

### Version 1.2.0 (2025-12-16) - Complete Self-Healing Architecture

**Major Updates:**

- Expanded hidden sheets from 6 to 13 total calculation sheets
- Added complete dropdown validation for all Config-referenced columns
- Added Home Towns header to Config sheet (column AF)
- Checkbox preservation after bulk setValues() operations

**New Hidden Calculation Sheets (7 added):**

- `_Grievance_Formulas` - Member lookup formulas for Grievance Log
- `_Dashboard_Summary_Calc` - 15 dashboard summary metrics
- `_Trends_Calc` - 12-month time-series analytics
- `_Location_Analytics_Calc` - Location-based breakdown
- `_Type_Analysis_Calc` - Issue category analysis
- `_Steward_Performance_Calc` - Steward performance scores with Performance Score formula
- `_Cost_Impact_Calc` - Financial impact estimates

**Dropdown Validations Added:**

- Member Directory: PREFERRED_COMM (J) → COMM_METHODS config
- Member Directory: HOME_TOWN (X) → HOME_TOWNS config

**Checkbox Repair Functions:**

- `repairGrievanceCheckboxes()` - Re-apply checkboxes to Message Alert (AC)
- `repairMemberCheckboxes()` - Re-apply checkboxes to Start Grievance (AE)

**Config Sheet Improvements:**

- Added "Home Towns" header at column AF (32) with consistent styling
- Auto-resize for Home Towns column

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
