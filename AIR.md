# 509 Dashboard - Architecture & Implementation Reference

**Version:** 1.6.0 (Merged Seed Functions & Status/Resolution Separation)
**Last Updated:** 2026-01-03
**Purpose:** Union grievance tracking and member engagement system for SEIU Local 509

---

## Creator & License

**Creator & Owner:** Wardis N. Vizcaino
**Role:** Steward at SEIU Local 509
**Contact:** wardis@pm.me

**License:** Free for use by non-profit collective bargaining groups and unions. No license required.

---

## Quick Start

1. Deploy the 9 `.gs` files to Google Apps Script
2. Run `CREATE_509_DASHBOARD()` to create 5 sheets + 5 hidden calculation sheets
3. Use `Demo > Seed All Sample Data` to populate test data
4. Customize Config sheet with your organization's values

---

## File Architecture

### Project Structure (9 Files)

```
509-dashboard/
├── Constants.gs           # Configuration constants (SHEETS, COLORS, MEMBER_COLS, GRIEVANCE_COLS)
├── Code.gs                # Main entry point, setup functions, sheet creation
├── SeedNuke.gs            # Demo data seeding and clearing functions
├── HiddenSheets.gs        # Self-healing hidden calculation sheets with auto-sync
├── ADHDFeatures.gs        # ADHD accessibility & theming (focus mode, themes, pomodoro)
├── DriveCalendarEmail.gs  # Google Drive, Calendar, Email notifications
├── TestingValidation.gs   # Test framework & data validation
├── PerformanceUndo.gs     # Caching layer & undo/redo system
├── MobileQuickActions.gs  # Mobile interface & quick actions menu
└── AIR.md                 # This document
```

### File Descriptions

**Constants.gs** (~500 lines)
- `SHEETS` - Sheet name constants (3 data + 2 dashboard + 5 hidden)
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

**Code.gs** (~900 lines)
- `onOpen()` - Creates menu system
- `CREATE_509_DASHBOARD()` - Main setup function (creates 5 sheets + 5 hidden)
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
- `searchMembers()` - Search members (stub)
- `startNewGrievance()` - Start grievance (stub)
- `viewActiveGrievances()` - Navigate to Grievance Log
- Sheet creation (5 functions): `createConfigSheet()`, `createMemberDirectory()`, `createGrievanceLog()`, `createDashboard()`, `createInteractiveDashboard()`

**SeedNuke.gs** (~500 lines)
- `SEED_SAMPLE_DATA()` - Seeds Config + 50 members + ~15 grievances
- `seedConfigData()` - Populate Config dropdowns
- `SEED_MEMBERS(count, grievancePercent)` - Seed N members with optional grievances (max 2000 members)
  - `count` - Number of members to seed
  - `grievancePercent` - Percentage of members to give grievances (0-100, default 30%)
  - Grievances are directly linked to member data (Member ID, Name, Email, etc.)
- `SEED_MEMBERS_DIALOG()` - Prompt for member count (30% get grievances)
- `SEED_MEMBERS_ADVANCED_DIALOG()` - Prompt for member count AND grievance percentage
- `seed50Members()` - Shortcut: seed 50 members (30% grievances)
- `seed100MembersWithGrievances()` - Shortcut: seed 100 members (50% grievances)
- `generateSingleMemberRow()` - Generate one member row (31 columns)
- `generateSingleGrievanceRow()` - Generate one grievance row (34 columns)
- `NUKE_ALL_DATA()` - Clear all data with confirmation
- `NUKE_CONFIG_DROPDOWNS()` - Clear only Config dropdowns
- `getConfigValues()` - Helper: get values from Config column
- `randomChoice()` - Helper: pick random array element
- `randomDate()` - Helper: generate random date
- `addDays()` - Helper: add days to date

**Note:** SEED_GRIEVANCES was merged into SEED_MEMBERS in v1.6.0. This ensures all seeded grievances are directly linked to actual seeded members with correct data.

**HiddenSheets.gs** (~1500 lines)
- `setupAllHiddenSheets()` - Create all 5 hidden calculation sheets
- Hidden Sheet Setup Functions (5 total):
  - `setupGrievanceCalcSheet()` - Grievance timeline formulas (auto-calc deadlines)
  - `setupGrievanceFormulasSheet()` - Member lookup formulas (First Name, Last Name, Email, etc.)
  - `setupMemberLookupSheet()` - Member → Grievance Log sync
  - `setupStewardContactCalcSheet()` - Steward contact tracking
  - `setupDashboardCalcSheet()` - Dashboard summary metrics (15 key metrics)
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

**DriveCalendarEmail.gs** (~500 lines) - Google Drive, Calendar & Email
- Google Drive:
  - `createRootFolder()` - Create base folder for grievance files
  - `createGrievanceFolder()` - Create folder for specific grievance
  - `linkFolderToGrievance()` - Link folder ID to grievance row
  - `setupDriveFolderForGrievance()` - Menu handler for folder creation
  - `listFolderFiles()` - List files in a folder
  - `showGrievanceFiles()` - Show files for selected grievance
  - `batchCreateGrievanceFolders()` - Create folders for all grievances
- Calendar:
  - `syncDeadlinesToCalendar()` - Sync all deadlines to Google Calendar
  - `checkCalendarEventExists()` - Check if event already exists
  - `clearAllCalendarEvents()` - Remove all dashboard calendar events
  - `showUpcomingDeadlinesFromCalendar()` - View upcoming deadlines
- Email Notifications:
  - `setupDailyDeadlineNotifications()` - Enable daily email reminders
  - `disableDailyDeadlineNotifications()` - Disable notifications
  - `checkDeadlinesAndNotify()` - Main notification check (runs daily)
  - `sendDeadlineNotification()` - Send individual notification
  - `showNotificationSettings()` - Notification configuration UI
  - `testDeadlineNotifications()` - Test notification system

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

**PerformanceUndo.gs** (~500 lines) - Caching Layer & Undo/Redo
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

**MobileQuickActions.gs** (~600 lines) - Mobile Interface & Quick Actions
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

## Sheet Structure (5 Visible + 5 Hidden)

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
├── Search Members
├── View Active Grievances
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
├── Seed All Sample Data
├── Seed Data (submenu)
│   ├── ⚙️ Seed Config Dropdowns Only
│   ├── 👥 Seed Members & Grievances (Custom)
│   ├── 👥 Seed Members (Advanced - Set % Grievances)
│   ├── 👥 Seed 50 Members (30% Grievances)
│   └── 👥 Seed 100 Members (50% Grievances)
└── Nuke Data (submenu)
    ├── ☢️ NUKE SEEDED DATA
    └── Clear Config Dropdowns Only

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
| I | Grievance Status | Open, Pending Info, In Arbitration, Appealed, Closed (preset) - **Workflow states only** |
| J | Grievance Step | Informal, Step I, Step II, etc. (preset) |
| K | Issue Category | Discipline, Workload, etc. (preset) |
| L | Articles Violated | Art. 1 - Art. 26 (preset) |
| M | Communication Methods | Email, Phone, Text, In Person (preset) |
| O | Grievance Coordinators | User populates |
| AF | Home Towns | User populates |

**Note on Status vs Resolution (v1.6.0):**
- **Status** = Workflow state (where the case is in the process): Open, Pending Info, In Arbitration, Appealed, Closed
- **Resolution** = Outcome (how the case ended, only when Status = Closed): Won - Full, Won - Partial, Settled - Favorable, Settled - Neutral, Denied - Appealing, Denied - Final, Withdrawn

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

| Function | Max Count | Batch Size | Notes |
|----------|-----------|------------|-------|
| SEED_MEMBERS(count, grievancePercent) | 2,000 members | 50 rows | grievancePercent default 30% |

**Example Usage:**
- `SEED_MEMBERS(100)` - Seeds 100 members, ~30 grievances
- `SEED_MEMBERS(100, 50)` - Seeds 100 members, ~50 grievances
- `SEED_MEMBERS(100, 0)` - Seeds 100 members, 0 grievances

---

## Hidden Sheet Architecture (Self-Healing)

The system uses 5 hidden calculation sheets with auto-sync triggers for cross-sheet data population. Formulas are stored in hidden sheets and synced to visible sheets, making them **self-healing** - if formulas are accidentally deleted, running REPAIR_DASHBOARD() restores them.

### Hidden Sheets (5 total)

| Sheet | Source | Destination | Purpose |
|-------|--------|-------------|---------|
| `_Grievance_Calc` | Grievance Log | Member Directory | AB-AD (Has Open Grievance?, Status, Days to Deadline) |
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

### Version 1.6.0 (2026-01-03) - Merged Seed Functions & Status/Resolution Separation

**Major Changes:**

1. **Status/Resolution Separation:**
   - **Status** now represents only workflow states: Open, Pending Info, In Arbitration, Appealed, Closed
   - **Resolution** now represents outcomes (only when Status = Closed): Won - Full, Won - Partial, Settled - Favorable, Settled - Neutral, Denied - Appealing, Denied - Final, Withdrawn
   - Removed duplicated values (Won, Denied, Settled, Withdrawn) from Status
   - Updated `GRIEVANCE_STATUS_PRIORITY` constant (now 5 values instead of 9)
   - Added new `GRIEVANCE_RESOLUTION_PRIORITY` constant for closed case sorting

2. **Merged Seed Functions:**
   - `SEED_GRIEVANCES` merged into `SEED_MEMBERS`
   - New signature: `SEED_MEMBERS(count, grievancePercent)`
   - `grievancePercent` parameter controls what percentage of members get grievances (default 30%)
   - All seeded grievances are now directly linked to member data (Member ID, Name, Email, etc.)
   - Menu updated with new options:
     - "Seed Members & Grievances (Custom)" - prompts for count
     - "Seed Members (Advanced - Set % Grievances)" - prompts for count AND percentage
     - "Seed 50 Members (30% Grievances)" - quick seed
     - "Seed 100 Members (50% Grievances)" - quick seed with more grievances

3. **Updated Sort Logic:**
   - Active cases (Open, Pending Info, In Arbitration, Appealed) sorted by Days to Deadline
   - Closed cases sorted by Resolution priority (wins first, denials last)
   - `sortGrievanceLogByStatus()` updated with new logic

4. **Testing Framework:**
   - Added testing framework directly to ConsolidatedDashboard.gs
   - Includes Assert library, test functions, and validation framework
   - No longer requires separate TestingValidation.gs file

5. **Interactive Dashboard Enhancements:**
   - Dropdown cells (A6:F6) now have yellow background with orange border highlighting
   - Unused columns (I onwards) are automatically hidden

**Code Changes:**
- `ConsolidatedDashboard.gs`: Updated DEFAULT_CONFIG, GRIEVANCE_STATUS_PRIORITY, added GRIEVANCE_RESOLUTION_PRIORITY
- `ConsolidatedDashboard.gs`: Rewrote `SEED_MEMBERS()` to include grievance seeding
- `ConsolidatedDashboard.gs`: Removed separate `SEED_GRIEVANCES()` function
- `ConsolidatedDashboard.gs`: Updated `sortGrievanceLogByStatus()` with new sorting logic
- `ConsolidatedDashboard.gs`: Added testing framework functions

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

**Status Priority Order (Workflow States):**
| Priority | Status | Type |
|----------|--------|------|
| 1 | Open | Active |
| 2 | Pending Info | Active |
| 3 | In Arbitration | Active |
| 4 | Appealed | Active |
| 5 | Closed | Resolved |

**Resolution Priority Order (Outcomes for Closed Cases):**
| Priority | Resolution | Category |
|----------|------------|----------|
| 1 | Won - Full | Win |
| 2 | Won - Partial | Win |
| 3 | Won | Win |
| 4 | Settled - Favorable | Settled |
| 5 | Settled - Neutral | Settled |
| 6 | Settled | Settled |
| 7 | Denied - Appealing | Denied |
| 8 | Denied - Final | Denied |
| 9 | Denied | Denied |
| 10 | Withdrawn | Other |

**Sort Logic:**
1. Primary sort: Status priority (active cases first, closed cases last)
2. For active cases: Sort by Days to Deadline (most urgent first)
3. For closed cases: Sort by Resolution priority (wins first, denials last)

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
