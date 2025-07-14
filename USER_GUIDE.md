# User Guide - Deacon Visitation Rotation System v1.1

> **Complete operational guide for daily use, troubleshooting, and understanding your deacon visitation system**

---

## 🎯 Quick Start - Your First Schedule

**Already completed setup?** Great! Here's how to create your first schedule:

1. **📅 Generate Schedule** - Click this to create your rotation
2. **🚨 Full Calendar Regeneration** - Export to Google Calendar  
3. **💬 Send Weekly Chat Summary** - Test your first notification

**That's it!** Your system is now operational. Read on for detailed guidance on each function.

---

## 📋 Main Menu Functions

### **📅 Generate Schedule**
**What it does:** Creates the complete visitation rotation using your deacon and household lists

**When to use:**
- Creating your first schedule
- When deacon or household lists change
- Starting a new year or quarter
- After significant roster updates

**What happens:**
- Reads your deacon names (Column L) and household names (Column M)
- Creates mathematically fair rotation in Columns A-E
- Generates individual deacon reports in Columns G-I
- Shows success message with schedule statistics

**⚠️ Common Issues:**
- **"No deacons found"** → Add deacon names starting in cell L2
- **"No households found"** → Add household names starting in cell M2
- **"Invalid start date"** → Check cell K2 has a valid Monday date
- **"Visit frequency too high"** → Check cell K4 (usually 2-4 weeks)

---

### **🔗 Generate Shortened URLs**
**What it does:** Creates mobile-friendly short URLs for Breeze profiles and visit notes

**When to use:**
- After adding new Breeze numbers (Column P) or Notes links (Column Q)
- When calendar events show long, hard-to-click URLs
- Before exporting to calendar for better mobile experience

**What happens:**
- Converts long URLs to tinyurl.com links
- Stores shortened URLs in Columns R and S
- Makes calendar events much more mobile-friendly

**⚠️ Note:** This can take 30+ seconds for many households due to API rate limiting

---

## 📆 Calendar Functions

### **📞 Update Contact Info Only**
**What it does:** Updates phone numbers and addresses in existing calendar events **without changing any scheduling**

**When to use:** ⭐ **SAFEST UPDATE OPTION**
- Household phone numbers or addresses change
- Breeze numbers or Notes links are updated
- You want to update contact info without affecting any custom scheduling

**What it preserves:**
- All custom dates and times deacons have set
- Any personal calendar entries copied by deacons
- Current week's scheduling (no disruption)

**Best practice:** Use this for routine contact updates

---

### **🔄 Update Future Events Only**
**What it does:** Rebuilds calendar events starting next Monday, preserving this week's events

**When to use:**
- Deacon or household roster changes (people added/removed)
- Major contact information overhaul needed
- Current week has custom scheduling you want to keep

**What it preserves:**
- This week's events (protects current week planning)
- Any custom scheduling deacons made for this week

**What it rebuilds:**
- All future events with latest contact info and roster changes

---

### **🚨 Full Calendar Regeneration**
**What it does:** ⚠️ **NUCLEAR OPTION** - Deletes ALL calendar events and rebuilds everything

**When to use:** **Only when necessary!**
- First-time calendar creation
- Major system problems requiring fresh start
- Calendar corruption or significant data issues

**⚠️ WARNING: You will lose:**
- All custom dates and times deacons have set
- Any scheduling adjustments made directly in calendar
- Current week's planning and arrangements

**The system will warn you** and suggest safer alternatives before proceeding

---

## 📢 Notification Functions

### **💬 Send Weekly Chat Summary**
**What it does:** Immediately sends the weekly 2-week lookahead to your Google Chat space

**When to use:**
- Testing that notifications work
- Manual notification when automation isn't set up
- Sending extra reminders before busy periods
- One-time notifications for schedule changes

**What gets sent:**
- Next 2 calendar weeks of visits
- Contact information for each household
- Clickable links to Breeze profiles and visit notes
- Links to calendar, guide, and summary (if configured)

---

### **🔄 Enable Weekly Auto-Send**
**What it does:** Sets up automatic weekly notifications based on your settings in K11 and K13

**When to use:**
- First-time automation setup
- After changing notification day or time
- When automation stops working

**Configuration required:**
- **K11:** Day of week (dropdown: Sunday through Saturday)
- **K13:** Hour of day (dropdown: 0-23, where 18 = 6 PM)
- **Webhook URL:** Must be configured first

**⚠️ Google Apps Script Reality:**
- Triggers may be delayed 15-20+ minutes
- New triggers take 24-48 hours to stabilize
- This is normal Google Apps Script behavior

---

### **📅 Show Auto-Send Schedule**
**What it does:** Shows current automation status and next scheduled notification

**When to use:**
- Checking if automation is working
- Verifying notification timing
- Troubleshooting when notifications don't arrive

**Shows you:**
- Whether automation is active
- Next scheduled notification time
- Current configuration settings
- Troubleshooting guidance

---

### **🛑 Disable Weekly Auto-Send**
**What it does:** Turns off automatic notifications (manual sending still works)

**When to use:**
- Vacation periods when notifications aren't needed
- Changing notification timing (disable, then re-enable)
- Troubleshooting notification issues
- End of visitation season

---

### **🔧 Configure Chat Webhook**
**What it does:** Sets up the Google Chat connection for notifications

**When to use:**
- Initial system setup
- Moving to a new chat space
- Webhook URL changes
- Test vs. production mode switching

**Setup process:**
1. Create Google Chat space for deacons
2. Add webhook to the space
3. Copy webhook URL
4. Paste URL when prompted
5. System auto-detects test vs. production mode

---

## 🛠️ Testing & Diagnostic Functions

### **📋 Test Notification System**
**What it does:** Sends a simple connectivity test message to verify chat integration

**When to use:**
- Initial setup verification
- Troubleshooting notification failures
- Testing after webhook configuration
- Confirming chat space is working

**What it tests:**
- Webhook URL is valid
- Chat space is accessible
- Basic message delivery
- Test vs. production mode detection

---

### **🧪 Test Notification Now**
**What it does:** Immediately sends a full weekly notification, bypassing trigger timing

**When to use:**
- Testing complete notification content and formatting
- Diagnosing trigger timing vs. content issues
- Verifying full system works before setting up automation
- Emergency manual notification

**Different from simple test:** This sends actual visit data, not just a connectivity test

---

### **🔍 Inspect All Triggers**
**What it does:** Shows detailed information about all automated functions in your system

**When to use:**
- Troubleshooting automation problems
- Understanding why notifications aren't sending
- Checking for duplicate or broken triggers
- System diagnosis

**What it shows:**
- All active triggers and their schedules
- Trigger creation dates and settings
- Whether weekly notifications are properly configured
- Warnings about missing or duplicate triggers

---

### **🔄 Force Recreate Trigger**
**What it does:** Deletes and recreates the weekly notification trigger

**When to use:** **Emergency troubleshooting only**
- Automation exists but notifications aren't sending
- Trigger appears broken or corrupted
- After major Google Apps Script updates
- When inspection shows trigger problems

**⚠️ Use carefully:** This is the "reboot" option for automation

---

### **🧪 Test Calendar Link Config**
**What it does:** Verifies that your calendar URL in K19 is properly configured

**When to use:**
- After configuring the calendar URL in K19
- Troubleshooting missing calendar links in notifications
- Verifying calendar setup before going live
- Testing new calendar URL configurations

**⚠️ Note:** This function only tests the calendar URL in K19. It doesn't check the guide URL (K22) or summary URL (K25) - those are tested when you send actual notifications.

---

## 📊 Data Management Functions

### **📊 Export Individual Schedules**
**What it does:** Creates a new spreadsheet with separate sheets for each deacon showing only their visits

**When to use:**
- Giving each deacon their personal schedule
- Creating printable individual schedules
- Deacon board reporting
- Historical record keeping

**What you get:**
- New spreadsheet with one sheet per deacon
- Visit dates, households, and contact information
- Downloadable/printable format for each deacon

---

### **📁 Archive Current Schedule**
**What it does:** Saves the current schedule to a new sheet before making changes

**When to use:**
- Before generating a new schedule
- Creating historical records
- Backup before major changes
- End-of-year record keeping

**Creates:**
- New sheet named "Archive_YYYY-MM-DD"
- Complete copy of current schedule and deacon reports
- Permanent record you can reference later

---

## 🔧 System Maintenance Functions

### **🔧 Validate Setup**
**What it does:** Comprehensive check of your system configuration

**When to use:**
- After initial setup
- Before generating first schedule
- Troubleshooting system problems
- Periodic system health checks

**What it checks:**
- Deacon and household lists
- Configuration settings (dates, frequencies)
- Contact information completeness
- Calendar permissions
- Notification setup

---

### **🧪 Run Tests**
**What it does:** Complete system functionality test suite

**When to use:**
- Verifying system is working correctly
- After making configuration changes
- Troubleshooting multiple issues
- Before important schedule generations

**What it tests:**
- Configuration loading
- URL shortening functionality
- Breeze URL construction
- Calendar access permissions
- Script properties access

---

### **Show Current Mode & Setup Instructions**
**What it displays:**
- Whether you're in Test Mode (🧪) or Production Mode (✅)
- Basic setup guidance
- Current system status

---

## 🚨 Understanding System Behavior

### **Test Mode vs. Production Mode**
The system automatically detects your operating mode:

**🧪 Test Mode** (when you have sample/test data):
- Creates "TEST - Deacon Visitation Schedule" calendar
- Uses test chat webhook (if configured)
- All notifications prefixed with "🧪 TEST:"
- Calendar events appear red
- Safe for learning and testing

**✅ Production Mode** (when you have real deacon/household names):
- Creates "Deacon Visitation Schedule" calendar
- Uses main deacon chat webhook
- Clean notifications without test prefixes
- Calendar events appear blue
- Ready for operational use

**You don't need to do anything** - the system detects this automatically based on your data patterns.

---

## 🔧 Common Issues & Solutions

### **"Notifications aren't being sent automatically"**

**Check these in order:**

1. **📢 Notifications → 📅 Show Auto-Send Schedule**
   - Is automation enabled? If not, use "🔄 Enable Weekly Auto-Send"

2. **📢 Notifications → 🔍 Inspect All Triggers**
   - Do you have exactly one weekly trigger? 
   - If none: Use "🔄 Enable Weekly Auto-Send"
   - If multiple: Use "🛑 Disable" then "🔄 Enable"

3. **📢 Notifications → 📋 Test Notification System**
   - Does manual testing work? If not, check webhook configuration

4. **Google Apps Script Reality Check:**
   - Triggers can be delayed 15-20+ minutes (this is normal)
   - New triggers take 24-48 hours to stabilize
   - Check execution logs in Apps Script for errors

---

### **"Error: No deacons found" or "No households found"**

**Solution:**
- Add names starting in **cell L2** (deacons) and **cell M2** (households)
- Names must be in consecutive cells with no blank rows
- Each name should be in its own cell

**The system looks for:**
- Column L: Deacon names starting row 2
- Column M: Household names starting row 2
- Names should be simple text (avoid special characters)

---

### **"Calendar events don't have contact information"**

**Usually means:**
- Contact info wasn't in the right columns when calendar was created
- Phone numbers should be in Column N
- Addresses should be in Column O  
- Breeze numbers in Column P, Notes links in Column Q

**Solution:**
1. Add missing contact information in correct columns
2. Use "📞 Update Contact Info Only" to refresh calendar
3. Or use "🔗 Generate Shortened URLs" first if you have Breeze/Notes links

---

### **"Schedule looks unfair - some deacons have way more visits"**

**This happens when:**
- Deacon and household counts create mathematical harmonics
- The system detects this and uses advanced algorithms to fix it

**The system handles this automatically** by:
- Detecting harmonic patterns (when counts share common factors)
- Using prime number offsets to ensure fair distribution
- Guaranteeing every deacon visits every household over time

**If it still looks unfair:**
- Check that you have the right number of weeks in K6
- Longer schedules create better distribution
- Generate a longer schedule (e.g., 52 weeks instead of 26)

---

### **"Webhook URL returned 404 - Chat space may be deleted"**

**This means:**
- The Google Chat space was deleted or modified
- The webhook URL is no longer valid
- You may need to create a new webhook

**Solution:**
1. Check that the Google Chat space still exists
2. Create a new webhook in the chat space if needed
3. Use "📢 Notifications → 🔧 Configure Chat Webhook" with new URL

---

### **"Cannot access calendar - permissions may have been revoked"**

**Solution:**
1. Go to Google Apps Script editor
2. Run any function to re-trigger authorization
3. Grant calendar permissions when prompted
4. Try the calendar operation again

---

### **Understanding Error Messages**

The system provides specific, actionable error messages:

- **"⚠️ Missing phone numbers: X households have no phone"** → Add phone numbers in Column N
- **"❌ Invalid start date in K2"** → Enter a Monday date in K2
- **"📊 Schedule Duration: Check K6"** → Enter number of weeks in K6
- **"🔧 No webhook configured for production mode"** → Configure webhook using menu
- **"⏰ WARNING: Multiple triggers found"** → Use disable/enable to clean up automation

**Every error message tells you exactly what to fix and where to find it.**

---

## 📈 Best Practices

### **Before Making Changes:**
1. **📁 Archive Current Schedule** to save your current state
2. Make your changes to deacon/household lists
3. **🔧 Validate Setup** to check everything looks good
4. **📅 Generate Schedule** to create new rotation

### **For Contact Updates:**
1. **📞 Update Contact Info Only** for routine updates (safest)
2. **🔄 Update Future Events Only** if roster changes
3. **🚨 Full Calendar Regeneration** only when absolutely necessary

### **For Notifications:**
1. **📋 Test Notification System** first to verify connectivity
2. **💬 Send Weekly Chat Summary** to test full content
3. **🔄 Enable Weekly Auto-Send** only after manual testing works
4. **📅 Show Auto-Send Schedule** to verify automation status

### **Regular Maintenance:**
- **Monthly:** Run "🧪 Run Tests" to verify system health
- **Quarterly:** Use "📁 Archive Current Schedule" for records
- **Yearly:** Generate new schedule with updated rosters

---

## 🎯 Quick Reference

**Need to send a notification now?** → **💬 Send Weekly Chat Summary**

**Contact info changed?** → **📞 Update Contact Info Only**

**Roster changed (people added/removed)?** → **🔄 Update Future Events Only**

**Automation not working?** → **📅 Show Auto-Send Schedule** → **🔍 Inspect All Triggers**

**System acting weird?** → **🧪 Run Tests** → **🔧 Validate Setup**

**Starting fresh?** → **📁 Archive Current Schedule** → **📅 Generate Schedule**

**Setting up for first time?** → **🔧 Configure Chat Webhook** → **📋 Test Notification System**

---

The Deacon Visitation Rotation System is designed to handle most situations gracefully and provide clear guidance when things go wrong. When in doubt, start with the testing functions to diagnose the issue, then use the safest fix option available.

**Remember:** The system automatically adapts to your data and usage patterns. Trust the automation, but verify with the testing tools when needed!