# 🚨 Seed Nuke Feature - Exit Demo Mode

## Overview

The **Seed Nuke** feature allows you to remove all test/seeded data from your dashboard and transition from demo mode to production mode. This is a one-way operation that prepares your dashboard for real-world use with actual member and grievance data.

---

## ⚠️ What Does "Nuke Seed Data" Do?

When you execute the **Nuke Seed Data** function, the system will:

### Data Removal
1. **Remove ALL Members**: Delete all test members from Member Directory
2. **Remove ALL Grievances**: Delete all test grievances from Grievance Log
3. **Delete Feedback Sheet**: Completely removes the Feedback & Development sheet
4. **Delete Menu Checklist**: Completely removes the Menu Checklist sheet
5. **Clear Steward Workload**: Remove all test steward assignments
6. **Clear Config Demo Data**: Remove demo entries from Config tab (if any exist):
   - Job Titles (Column A)
   - Office Locations (Column B)
   - Units (Column C)
   - Supervisors (Column F)
   - Managers (Column G)
   - Stewards (Column H)
   - Grievance Coordinators (Column O)
   - Home Towns (Column AF)
   - Office Addresses (Column AN)

   > **NOTE (v3.11+):** These fields are now LEFT EMPTY during CREATE_509_DASHBOARD. Users populate them with their own data. If no user data was added, there's nothing to clear.

### Demo Mode Disabling
7. **Disable Demo Menu**: Sets a flag (`DEMO_MODE_DISABLED`) to hide the "🎭 Demo" menu on next refresh
8. **Clear Tracking**: Removes the tracked seeded ID lists from Script Properties

### Preserved Items
9. **Preserve Organization Info**: Keep your real organization settings:
   - Organization Name, Local Number, Main Address, Phone
   - Union Parent, State/Region, Website
   - Main Fax, Toll Free numbers
   - All deadline and contract reference columns
10. **Preserve Structure**: Keep all headers, formulas, and sheet structure intact
11. **Preserve Manually Entered Data**: Only rows with seeded ID patterns (M/G + 4 letters + 3 digits) are deleted

> **🔴 IMPORTANT**: The nuke operation only deletes **seeded data rows** (matching the pattern `MJOSM123` or `GJOSM456`) and completely removes the **Feedback & Development** and **Menu Checklist** sheets. Manually entered data with different ID formats is preserved. The Demo menu is hidden but the seed functions remain in the code.

---

## 🎯 When Should You Nuke Seed Data?

**Nuke the seed data when:**
- You're ready to start using the dashboard with real member data
- You've completed all testing and training with the demo data
- You understand how the system works and are ready for production
- You want to start fresh without test data cluttering your sheets

**DON'T nuke if:**
- You're still learning how to use the system
- You want to keep testing features
- You haven't backed up the spreadsheet yet
- You're using this for training or demonstration purposes

---

## 📋 Pre-Nuke Checklist

Before nuking seed data, make sure you:

- [ ] **Understand the system**: You know how to add members and grievances
- [ ] **Have a backup**: Copy the spreadsheet if you want to keep a demo version
- [ ] **Review Config settings**: Ensure dropdown values match your needs
- [ ] **Prepare real data**: Have member and grievance data ready to import
- [ ] **Train your team**: Everyone knows how to use the system
- [ ] **Document processes**: You have procedures for data entry and workflow

---

## 🚀 How to Nuke Seed Data

### Step 1: Access the Nuke Function

**Menu**: `🎭 Demo > 🗑️ Nuke Data > ☢️ NUKE SEEDED DATA`

*Note: This menu item only appears if seed data hasn't been nuked yet.*

### Step 2: Confirm the Action

You'll see **TWO confirmation dialogs**:

**First Confirmation:**
```
⚠️ WARNING: Remove All Seeded Data & Functions

This will PERMANENTLY remove:
• All test data from Member Directory, Grievance Log, Steward Workload
• Config Tab Demo Entries (Job Titles, Locations, etc.)
• ALL seed functions from the script code
• ALL seed menu items
• THIS NUKE FUNCTION ITSELF (complete self-deletion)

After this operation, there will be NO trace of seed OR nuke functionality.

This action CANNOT be undone!

Are you sure you want to proceed?
```

**Second Confirmation (Final):**
```
🚨 FINAL CONFIRMATION

This is your last chance!

ALL test data, seed code, AND this nuke function will be permanently deleted.
The SeedNuke.gs file will be completely removed from the project.

Click YES to proceed.
```

### Step 3: Wait for Processing

After confirming:
- You'll see a "⏳ Removing seeded data..." message
- The process takes a few moments
- **Do not close the spreadsheet** while processing

### Step 4: Verify Clean State

After nuking, verify the system is ready:
- **Setup checklist**: Steps to configure your production environment
- **Steward contact info reminder**: Enter your steward details
- **Next actions**: What to do first

---

## 📊 What Happens After Nuking

### Immediate Changes

1. **Empty Sheets** (with headers intact):
   - Member Directory: 0 members
   - Grievance Log: 0 grievances
   - Steward Workload: Empty

2. **Dashboards Reset**:
   - All metrics show zero
   - Charts will be empty
   - No overdue grievances

3. **Menu Changes**:
   - "🎭 Demo" menu completely removed
   - Cleaner menu structure
   - Focus on production tools

4. **Code Changes** (if Apps Script API enabled):
   - All SEED_* functions removed from Code.gs (~600 lines deleted)
   - **SeedNuke.gs completely deleted** (file removed from project)
   - Seed menu items removed from ReorganizedMenu.gs
   - **Nuke menu item also removed** from ReorganizedMenu.gs
   - **Zero evidence** that seed OR nuke functionality ever existed

### What Remains Intact

✅ All sheet headers and column structure
✅ Config tab dropdown lists
✅ Timeline rules table
✅ Steward contact info section (if already filled)
✅ All analytical tabs and dashboard layouts
✅ Color schemes and formatting
✅ All custom menu items (except seed options)
✅ Triggers and automation

---

## 🎬 Post-Nuke Setup Steps

After nuking, follow these steps to set up for production:

### 1. Configure Steward Contact Info

**Location**: `⚙️ Config > Column U`

Enter:
- **Row 2**: Steward Name
- **Row 3**: Steward Email
- **Row 4**: Steward Phone

Or use: `👤 Dashboard > 📋 Grievance Tools > ➕ Start New Grievance` (then configure steward info in Config sheet)

### 2. Review Config Dropdown Lists

**Location**: `⚙️ Config > Columns A-N`

Verify these match your organization:
- Job Titles (Column A)
- Work Locations (Column B)
- Units (Column C)
- Grievance Types (Column G)
- Committee Types (Column L)

Add, remove, or modify as needed.

### 3. Add Real Members

**Location**: `👥 Member Directory`

You can:
- **Manual Entry**: Type directly into the sheet
- **CSV Import**: Use File > Import
- **Copy/Paste**: From another spreadsheet

Required fields:
- Member ID
- First Name
- Last Name
- Is Steward (Yes/No)

### 4. Set Up Triggers

**Menu**: `⚙️ Administrator > 🔧 Setup & Triggers > ⚡ Install Auto-Sync Trigger`

This enables:
- Automatic deadline calculations
- Real-time dashboard updates
- Member snapshot updates

### 5. Customize Interactive Dashboard

**Menu**: `👤 Dashboard > 🎯 Interactive Dashboard`

Create custom views for:
- Key metrics you track
- Charts you need most
- Your preferred layout

### 6. Test Grievance Workflow (Optional)

If using the grievance workflow:
1. Set up Google Form (see [Grievance Workflow Guide](GRIEVANCE_WORKFLOW_GUIDE.md))
2. Configure form submission trigger
3. Test with a sample grievance

---

## 🔄 Can I Undo the Nuke?

**No, the data deletion is permanent and irreversible.**

After nuking:
- Seeded data rows are permanently deleted from Member Directory and Grievance Log
- The Demo menu is hidden (based on a Script Property flag)
- Config dropdowns are cleared

**Note**: The seed functions remain in the code - only the data is deleted. To re-seed:
1. Clear the `DEMO_MODE_DISABLED` Script Property manually, or
2. Redeploy from fresh source code

**Recovery options:**
- **Re-enable Demo menu**: Script Properties → Delete `DEMO_MODE_DISABLED`
- **Restore from backup**: If you made a copy of the spreadsheet before nuking
- **Fresh deployment**: Deploy a new copy from the original source code
- **Import data**: Add your real data to start fresh (recommended approach)

---

## 🆘 Troubleshooting

### Problem: Dashboards Still Show Data

**Solution**:
- Go to `📊 Sheet Manager > 📊 Rebuild Dashboard`
- This recalculates all metrics

### Problem: Demo Menu Still Visible

**Solution**:
- Close and reopen the spreadsheet
- The menu is rebuilt on open
- The 🎭 Demo menu only shows when demo mode is enabled

### Problem: Need to Re-Seed for Training

**Solution**:
Since seed functions are permanently deleted after nuking, you cannot re-seed:
1. Deploy a fresh copy from the original source code
2. Or restore from a backup made BEFORE the nuke
3. The `resetNukeFlag()` function no longer restores seed capabilities

### Problem: Accidentally Nuked Too Soon

**Solution**:
- If you have a backup copy, restore from there
- If no backup, you'll need to import your data manually
- For future: Always make a backup first!

---

## 🔧 Apps Script API Requirement

For the **automatic code removal** feature to work, the Apps Script API must be enabled:

### Enabling the Apps Script API

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Select your project (or create one linked to your script)
3. Go to **APIs & Services** > **Library**
4. Search for "Apps Script API"
5. Click **Enable**

### Required OAuth Scope

The script needs this OAuth scope for automatic code removal:
```
https://www.googleapis.com/auth/script.projects
```

### What Happens Without the API?

If the Apps Script API is not enabled:
- All **data** will still be cleared (members, grievances, config demo data)
- Seed **code** will NOT be automatically removed
- You'll see a message with manual cleanup instructions
- You can manually delete seed functions from the Apps Script editor

---

## 💡 Best Practices

### Before Nuking

1. **Make a backup copy**: File > Make a copy
2. **Document your needs**: List required config changes
3. **Prepare import data**: Have member/grievance CSVs ready
4. **Train your team**: Everyone understands the workflow

### After Nuking

1. **Start small**: Add a few test members first
2. **Verify calculations**: Check that auto-calculations work
3. **Test workflows**: Try the grievance workflow with test data
4. **Gradual rollout**: Add real data in batches
5. **Monitor performance**: Watch how dashboards update

---

## 📚 Related Documentation

- [Main README](README.md) - Complete dashboard documentation
- [Grievance Workflow Guide](GRIEVANCE_WORKFLOW_GUIDE.md) - How to use grievance features
- [Interactive Dashboard Guide](INTERACTIVE_DASHBOARD_GUIDE.md) - Customization options
- [ADHD-Friendly Guide](ADHD_FRIENDLY_GUIDE.md) - Accessibility features

---

## 🎯 Quick Reference

### Menu Location (Before Nuke)
```
🎭 Demo > 🗑️ Nuke Data > ☢️ NUKE SEEDED DATA
```

### What Gets Deleted
- ❌ All members from Member Directory
- ❌ All grievances from Grievance Log
- ❌ All steward workload data
- ❌ Feedback & Development sheet (entire sheet)
- ❌ Menu Checklist sheet (entire sheet)
- ❌ Config demo data (job titles, locations, units, supervisors, managers, stewards, coordinators, home towns, office addresses)

### What Gets Preserved
- ✅ Headers and structure
- ✅ Organization info (name, local number, address, phone, fax, toll-free, website)
- ✅ Deadline settings and contract references
- ✅ Dashboards and charts
- ✅ All formulas and formatting
- ✅ Menu system

### Post-Nuke Priorities
1. Enter steward contact info
2. Review config settings
3. Add real members
4. Set up triggers
5. Test grievance workflow

---

## 🎉 Welcome to Production!

After nuking seed data, you're ready to use the 509 Dashboard with real member and grievance data.

The system is now configured for:
- **Real member tracking**
- **Actual grievance management**
- **Live deadline monitoring**
- **Authentic reporting and analytics**

Your dashboard is **production-ready**! 🚀

---

## See Also

- **`NUKE_SEEDED_DATA()`** - Smart nuke that removes only seeded data (pattern-matched IDs)
  - Menu: `🎭 Demo > 🗑️ Nuke Data > ☢️ NUKE SEEDED DATA`
- **`NUKE_CONFIG_DROPDOWNS()`** - Clears only Config dropdown values
  - Menu: `🎭 Demo > 🗑️ Nuke Data > 🧹 Clear Config Dropdowns Only`

---

**Last Updated**: 2026-01-03
**Version**: 1.6.0

---

## Version 1.6.0 Notes

In version 1.6.0, the seed functionality was simplified:

- **SEED_GRIEVANCES was merged into SEED_MEMBERS** - Use `SEED_MEMBERS(count, grievancePercent)` to seed members with optional grievances
- **All seeded grievances are directly linked to members** - No orphaned grievances with missing member info
- **Separate SEED_GRIEVANCES function removed** - Use the merged approach instead

**Example:**
- `SEED_MEMBERS(100)` - Seeds 100 members, ~30 grievances (30% default)
- `SEED_MEMBERS(100, 50)` - Seeds 100 members, ~50 grievances (50%)
- `SEED_MEMBERS(100, 0)` - Seeds 100 members only, no grievances
