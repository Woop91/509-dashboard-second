# Menu Items Checklist

## 1. Dashboard Menu (👤 Dashboard)

- [ ] 📊 Smart Dashboard (Auto-Detect) - `showSmartDashboard`
- [ ] 🎯 Interactive Dashboard - `showInteractiveDashboardTab`
- [ ] 📋 View Active Grievances - `viewActiveGrievances`
- [ ] 📱 Mobile Dashboard - `showMobileDashboard`
- [ ] 📱 Get Mobile App URL - `showWebAppUrl`
- [ ] ⚡ Quick Actions - `showQuickActionsMenu`

### Dashboard > Grievance Tools (Submenu)

- [ ] ➕ Start New Grievance - `startNewGrievance`
- [ ] 🔄 Refresh Grievance Formulas - `recalcAllGrievancesBatched`
- [ ] 🔄 Refresh Member Directory Data - `refreshMemberDirectoryFormulas`
- [ ] 🔗 Setup Live Grievance Links - `setupLiveGrievanceFormulas`
- [ ] 👤 Setup Member ID Dropdown - `setupGrievanceMemberDropdown`
- [ ] 🔧 Fix Overdue Text Data - `fixOverdueTextToNumbers`

---

## 2. Search Menu (🔍 Search)

- [ ] 🔍 Search Members - `searchMembers`

---

## 3. Sheet Manager Menu (📊 Sheet Manager)

- [ ] 📊 Rebuild Dashboard - `rebuildDashboard`
- [ ] 📈 Refresh Interactive Charts - `refreshInteractiveCharts`
- [ ] 🔄 Refresh All Formulas - `refreshAllFormulas`

### Sheet Manager > Google Drive (Submenu)

- [ ] 📁 Setup Folder for Grievance - `setupDriveFolderForGrievance`
- [ ] 📁 View Grievance Files - `showGrievanceFiles`
- [ ] 📁 Batch Create Folders - `batchCreateGrievanceFolders`

### Sheet Manager > Calendar (Submenu)

- [ ] 📅 Sync Deadlines to Calendar - `syncDeadlinesToCalendar`
- [ ] 📅 View Upcoming Deadlines - `showUpcomingDeadlinesFromCalendar`
- [ ] 🗑️ Clear Calendar Events - `clearAllCalendarEvents`

### Sheet Manager > Notifications (Submenu)

- [ ] ⚙️ Notification Settings - `showNotificationSettings`
- [ ] 🧪 Test Notifications - `testDeadlineNotifications`

---

## 4. Tools Menu (🔧 Tools)

### Tools > ADHD & Accessibility (Submenu)

- [ ] ♿ ADHD Control Panel - `showADHDControlPanel`
- [ ] 🎯 Focus Mode - `activateFocusMode`
- [ ] 🔲 Toggle Zebra Stripes - `toggleZebraStripes`
- [ ] 📝 Quick Capture - `showQuickCaptureNotepad`
- [ ] 🍅 Pomodoro Timer - `startPomodoroTimer`

### Tools > Theming (Submenu)

- [ ] 🎨 Theme Manager - `showThemeManager`
- [ ] 🌙 Toggle Dark Mode - `quickToggleDarkMode`
- [ ] 🔄 Reset Theme - `resetToDefaultTheme`

### Tools > Multi-Select (Submenu)

- [ ] 📝 Open Editor - `showMultiSelectDialog`
- [ ] ⚡ Enable Auto-Open - `installMultiSelectTrigger`
- [ ] 🚫 Disable Auto-Open - `removeMultiSelectTrigger`

### Tools > Undo/Redo (Submenu)

- [ ] ↩️ Undo Last Action - `undoLastAction`
- [ ] ↪️ Redo Action - `redoLastAction`
- [ ] 📋 View History - `showUndoRedoPanel`
- [ ] 🗑️ Clear History - `clearUndoHistory`

### Tools > Cache & Performance (Submenu)

- [ ] 🗄️ Cache Status - `showCacheStatusDashboard`
- [ ] 🔥 Warm Up Caches - `warmUpCaches`
- [ ] 🗑️ Clear All Caches - `invalidateAllCaches`

### Tools > Validation (Submenu)

- [ ] 🔍 Run Bulk Validation - `runBulkValidation`
- [ ] ⚙️ Validation Settings - `showValidationSettings`
- [ ] 🧹 Clear Indicators - `clearValidationIndicators`
- [ ] ⚡ Install Validation Trigger - `installValidationTrigger`

---

## 5. Setup Menu (🏗️ Setup)

- [ ] 🔧 REPAIR DASHBOARD - `REPAIR_DASHBOARD`
- [ ] ⚙️ Setup Data Validations - `setupDataValidations`
- [ ] 🎨 Setup ADHD Defaults - `setupADHDDefaults`

---

## 6. Demo Menu (🎭 Demo) - *Conditional: shown if demo mode enabled*

### Demo > Seed Data (Submenu)

- [ ] ⚙️ Seed Config Dropdowns Only - `seedConfigData`
- [ ] 👥 Seed Members (Custom Count) - `SEED_MEMBERS_DIALOG`
- [ ] 📋 Seed Grievances (Custom Count) - `SEED_GRIEVANCES_DIALOG`
- [ ] 👥 Seed Members Advanced - `SEED_MEMBERS_ADVANCED_DIALOG`
- [ ] 👥 Seed 50 Members - `seed50Members`
- [ ] 👥 Seed 100 Members with Grievances - `seed100MembersWithGrievances`
- [ ] 📋 Seed 25 Grievances - `seed25Grievances`

### Demo > Nuke Data (Submenu)

- [ ] ☢️ NUKE SEEDED DATA - `NUKE_SEEDED_DATA`
- [ ] 🧹 Clear Config Dropdowns Only - `NUKE_CONFIG_DROPDOWNS`
- [ ] 🔄 Restore Config & Dropdowns - `restoreConfigAndDropdowns`

---

## 7. Testing Menu (🧪 Testing)

- [ ] 🧪 Run All Tests - `runAllTests`
- [ ] ⚡ Run Quick Tests - `runQuickTests`
- [ ] 📊 View Test Results - `viewTestResults`

---

## 8. Administrator Menu (⚙️ Administrator)

- [ ] 🔍 DIAGNOSE SETUP - `DIAGNOSE_SETUP`
- [ ] 🔍 Verify Hidden Sheets - `verifyHiddenSheets`

### Administrator > Setup & Triggers (Submenu)

- [ ] 🔧 Setup All Hidden Sheets - `setupAllHiddenSheets`
- [ ] 🔧 Repair All Hidden Sheets - `repairAllHiddenSheets`
- [ ] ⚡ Install Auto-Sync Trigger - `installAutoSyncTrigger`
- [ ] 🚫 Remove Auto-Sync Trigger - `removeAutoSyncTrigger`

### Administrator > Manual Sync (Submenu)

- [ ] 🔄 Sync All Data Now - `syncAllData`
- [ ] 🔄 Sync Grievance → Members - `syncGrievanceToMemberDirectory`
- [ ] 🔄 Sync Members → Grievances - `syncMemberToGrievanceLog`

---

## Summary

| Menu | Item Count |
|------|------------|
| Dashboard | 12 |
| Search | 1 |
| Sheet Manager | 11 |
| Tools | 19 |
| Setup | 3 |
| Demo | 10 |
| Testing | 3 |
| Administrator | 9 |
| **Total** | **68** |
