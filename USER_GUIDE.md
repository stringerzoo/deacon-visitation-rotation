# User Guide v1.1 - Deacon Visitation Rotation System

> **Complete operational reference for day-to-day system management**

This guide covers everything you need to know for successfully operating the Deacon Visitation Rotation System on an ongoing basis. For initial setup, see [SETUP.md](SETUP.md).

---

## 📋 Daily Operations

### **Checking System Health**
**Frequency**: Daily (takes 30 seconds)

1. **Check for notifications**: Did Saturday's automated notification arrive?
2. **Quick calendar review**: Are this week's events still accurate?
3. **Monitor chat space**: Any deacon questions or issues?

**🚨 If something seems wrong**: Use **🧪 Run Tests** to diagnose issues

### **Handling Deacon Questions**
**Common questions and responses:**

**"When exactly should I visit?"**
- ✅ **Any time during your assigned week**
- ✅ **Call ahead to schedule specific time**  
- ✅ **±3 days flexibility is generally acceptable**

**"I can't visit this week, what do I do?"**
- 📞 **Contact the household** to reschedule (±1 week)
- 💬 **Notify other deacons** in chat if you need coverage
- 📅 **Update calendar event** with actual visit date when confirmed

**"How do I access the visit notes?"**
- 📋 **Click "Visit Notes" link** in the calendar event description
- ✏️ **Add visit summary** to the shared Google Doc
- 🙏 **Note any prayer requests** or concerns

---

## 📅 Weekly Management

### **Monday: Week Start Review**
1. **Review upcoming visits** for the week
2. **Check for any deacon scheduling conflicts**
3. **Send manual reminder if needed**: **💬 Send Weekly Chat Summary**

### **Wednesday: Mid-Week Check**
1. **Follow up on any pending visits**
2. **Address any reported issues**
3. **Update contact info if needed**: **📞 Update Contact Info Only**

### **Saturday: Automation Check**
1. **Verify automated notification arrived** (usually 6:00 PM)
2. **Review next 2 weeks** for accuracy
3. **Handle any weekend questions**

---

## 🔄 Monthly Maintenance

### **First Monday of Month**
**Run system health check:**
- **🧪 Run Tests** - Complete system validation
- **🔍 Inspect All Triggers** - Verify automation health
- **📋 Test Notification System** - Confirm chat connectivity

### **Contact Information Updates**
**When phone numbers, addresses, or links change:**

1. **Update spreadsheet data** (columns N, O, P, Q)
2. **Use smart update**: **📞 Update Contact Info Only**
   - ✅ **Preserves all custom scheduling**
   - 🔄 **Updates contact info in all events**
   - 🛡️ **Safest option - always use this first**

### **Schedule Adjustments**
**When deacon or household lists change:**

1. **Update spreadsheet data** (columns L, M)
2. **Choose appropriate update method**:
   - **🔄 Update Future Events Only** - Preserves current week
   - **🚨 Full Calendar Regeneration** - Nuclear option (use sparingly)

---

## 🛠️ Complete Menu Reference

### **📅 Generate Schedule**
**Purpose**: Creates fair rotation schedule using mathematical algorithm
**When to use**: 
- Initial setup
- New year planning
- Major roster changes
**Expected outcome**: Schedule appears in columns A-E
**⚠️ Note**: Run this before calendar export

### **🔗 Generate Shortened URLs**
**Purpose**: Creates mobile-friendly short URLs for Breeze profiles and visit notes
**When to use**:
- After adding new households
- When original URLs change
- Initial setup
**Expected outcome**: 
- ✅ **Preserves existing short URLs**
- 🆕 **Creates new URLs only for empty cells**
- 📊 **Shows detailed summary** of what was preserved vs. created

### **🔄 Force Regenerate All URLs**
**Purpose**: Intentionally replaces ALL existing short URLs with fresh ones
**When to use**:
- Existing short URLs have expired
- Need to refresh all URLs for some reason
- TinyURL service issues
**Expected outcome**: All URLs in columns R and S are replaced
**⚠️ Warning**: This destroys existing short URLs

---

## 📆 Calendar Functions

### **📞 Update Contact Info Only**
**Purpose**: Updates contact information in calendar events while preserving ALL custom scheduling
**When to use**:
- Phone numbers changed
- Addresses updated
- Breeze links modified
- Visit notes links changed
**What it preserves**:
- ✅ Custom event times
- ✅ Modified dates
- ✅ Added guests
- ✅ Location changes
- ✅ Any deacon customizations
**Expected outcome**: Contact info refreshed, scheduling intact
**⏱️ Time**: 30-60 seconds for typical schedule

### **🔄 Update Future Events Only**
**Purpose**: Updates events starting from next Monday while preserving current week
**When to use**:
- Deacon roster changes
- Household list changes
- Major contact updates
- Mid-week adjustments needed
**What it preserves**:
- ✅ Current week events (completely untouched)
- ✅ Any customizations to this week's visits
**What it updates**:
- 🔄 Future week assignments
- 🔄 Contact information
- 🔄 Event descriptions
**Expected outcome**: Current week safe, future weeks refreshed
**⏱️ Time**: 1-2 minutes for typical schedule

### **🚨 Full Calendar Regeneration** ⚠️ **Nuclear Option**
**Purpose**: Deletes ALL existing events and creates completely fresh calendar
**When to use**:
- Initial setup
- Calendar corrupted
- Major system problems
- Complete roster overhaul
**What it destroys**:
- ❌ All custom scheduling
- ❌ Modified times and dates
- ❌ Added guests
- ❌ Location customizations
**Expected outcome**: Brand new calendar, all customizations lost
**⏱️ Time**: 2-5 minutes for typical schedule
**💡 Tip**: Try other options first!

---

## 📢 Notifications

### **💬 Send Weekly Chat Summary** 
**Purpose**: Manually send 2-week lookahead to deacon chat space
**When to use**:
- Automated notification failed
- Need immediate update
- Testing notification format
**Expected outcome**: Rich-formatted message with contact info and links
**⏱️ Time**: 5-10 seconds

### **🔄 Enable Weekly Auto-Send**
**Purpose**: Set up automated weekly notifications
**When to use**: Initial setup or after trigger deletion
**Configuration**: Uses K11 (day) and K13 (time) settings
**Expected outcome**: Weekly notifications start (may take 24-48 hours to stabilize)

### **📅 Show Auto-Send Schedule**
**Purpose**: View current trigger configuration
**When to use**: Verify automation timing or troubleshoot issues
**Expected outcome**: Shows day, time, and status of automated notifications

### **🛑 Disable Weekly Auto-Send**
**Purpose**: Stop automated notifications
**When to use**: 
- Vacation periods
- System maintenance
- Changing notification timing
**Expected outcome**: Automated notifications stop

### **🔧 Configure Chat Webhook**
**Purpose**: Set up or change Google Chat integration
**When to use**:
- Initial setup
- Moving to new chat space
- Webhook URL changed
**Expected outcome**: Notifications will go to specified chat space

### **📋 Test Notification System**
**Purpose**: Verify chat connectivity and message formatting
**When to use**:
- After webhook changes
- Troubleshooting delivery issues
- Verifying test vs production separation
**Expected outcome**: Test message appears in chat space

---

## 📊 Reports and Exports

### **📊 Export Individual Schedules**
**Purpose**: Create separate spreadsheet with individual deacon schedules
**When to use**:
- Deacon wants personal copy
- Offline reference needed
- Delegation to spouses
**Expected outcome**: New spreadsheet with tab for each deacon
**⏱️ Time**: 30-45 seconds

### **📁 Archive Current Schedule**
**Purpose**: Save current schedule before making major changes
**When to use**:
- Before regenerating schedule
- Historical record keeping
- Backup before major roster changes
**Expected outcome**: Schedule saved to new sheet in same workbook

---

## 🔧 Diagnostics and Testing

### **🧪 Run Tests**
**Purpose**: Comprehensive system health check
**When to use**:
- Monthly maintenance
- Troubleshooting problems
- After making changes
**What it checks**:
- Configuration validity
- Data integrity
- URL functionality
- Calendar access
- Notification connectivity
**Expected outcome**: Detailed report of system status

### **🔍 Inspect All Triggers**
**Purpose**: View all automated schedules and their status
**When to use**:
- Troubleshooting automation
- Understanding timing issues
- Verifying trigger health
**Expected outcome**: List of all triggers with timing and status

### **🔄 Force Recreate Trigger**
**Purpose**: Delete and recreate weekly notification trigger
**When to use**:
- Trigger appears broken
- Notifications stopped working
- Major timing changes needed
**Expected outcome**: Fresh trigger created (may take 24-48 hours to stabilize)

### **🧪 Test Calendar Link Config**
**Purpose**: Verify resource URLs in K22, K25 are accessible
**When to use**:
- After changing resource links
- Troubleshooting link access
- Verifying mobile compatibility
**Expected outcome**: Report on link accessibility

---

## 🚨 Troubleshooting Guide

### **Notifications Not Arriving**
**Diagnosis Steps**:
1. Check **📅 Show Auto-Send Schedule** - Is trigger active?
2. Run **📋 Test Notification System** - Does manual test work?
3. Verify chat space webhook URL hasn't changed
4. Check **🔍 Inspect All Triggers** for errors

**Solutions**:
- **No trigger**: Use **🔄 Enable Weekly Auto-Send**
- **Webhook broken**: Use **🔧 Configure Chat Webhook**
- **Timing issues**: Wait 24-48 hours for new triggers to stabilize
- **Immediate fix**: Use **💬 Send Weekly Chat Summary** manually

### **Calendar Updates Failing**
**Diagnosis Steps**:
1. Run **🧪 Run Tests** to check permissions
2. Verify calendar exists and is accessible
3. Check for data synchronization issues
4. Try progressively safer update methods

**Solutions**:
- **Permission denied**: Reauthorize script with calendar access
- **Calendar not found**: Use **🚨 Full Calendar Regeneration** to create new
- **Data sync issues**: Regenerate schedule first
- **Partial failures**: Use **📞 Update Contact Info Only** (safest option)

### **URL Generation Problems**
**Diagnosis Steps**:
1. Check source data in columns P and Q
2. Verify internet connectivity
3. Test with small subset of households
4. Check for API rate limiting

**Solutions**:
- **Empty source data**: Add Breeze numbers and Notes links
- **Rate limiting**: Wait and try again
- **TinyURL issues**: Use **🔄 Force Regenerate All URLs**
- **Partial success**: Use **🔗 Generate Shortened URLs** (preserves working URLs)

### **Schedule Generation Issues**
**Diagnosis Steps**:
1. Verify deacon list (column L) has data
2. Verify household list (column M) has data
3. Check configuration settings (K1-K8)
4. Run **🧪 Run Tests** for detailed validation

**Solutions**:
- **Empty lists**: Add deacon and household names
- **Invalid dates**: Fix start date in K2
- **Math errors**: Adjust visit frequency in K4
- **Character issues**: Remove special characters from names

---

## 💡 Best Practices

### **Change Management**
**Always follow this order for safety:**
1. **Archive current schedule** if making major changes
2. **Start with safest update method** (Contact Info Only)
3. **Test with small changes** before major modifications
4. **Communicate changes** to deacons in advance

### **Routine Maintenance Schedule**
**Daily** (30 seconds):
- Check for obvious issues
- Respond to deacon questions

**Weekly** (5 minutes):
- Verify automation working
- Review upcoming assignments
- Handle contact updates

**Monthly** (15 minutes):
- Run system health check
- Update any changed contact information
- Review and archive completed schedules

**Quarterly** (30 minutes):
- Full system review
- Document any customizations
- Plan for roster changes
- Update backup procedures

### **Emergency Procedures**
**If system appears completely broken:**
1. **Don't panic** - Most issues are easily fixable
2. **Run diagnostics first**: **🧪 Run Tests**
3. **Try safest fixes first**: Contact info updates only
4. **Manual backup**: Use **💬 Send Weekly Chat Summary**
5. **Document the issue** for future prevention

**If you need to hand off to someone else urgently:**
1. **Send manual notification** immediately
2. **Document the current issue** clearly
3. **Share this User Guide** with emergency contact
4. **Provide spreadsheet access** with editing permissions

---

## 📞 Support and Resources

### **When You Need Help**
1. **Check this User Guide** first
2. **Run built-in diagnostics** (🧪 Run Tests)
3. **Review error messages** carefully
4. **Document the specific problem** with screenshots
5. **Check project documentation** in repository

### **Key Resources**
- **User Guide**: This document - operational reference
- **Setup Guide**: [SETUP.md](SETUP.md) - installation instructions
- **Technical Documentation**: [FEATURES.md](FEATURES.md) - system capabilities
- **Handoff Guide**: [docs-handoff_document_comprehensive.md] - transition planning

### **Understanding Test vs Production**
The system **automatically detects** which mode you're in:
- **🧪 Test Mode**: Uses "TEST" calendars and webhooks
- **✅ Production Mode**: Uses normal calendars and webhooks

**Test Mode Indicators**:
- Household names like "Alan Adams", "Barbara Baker"
- Phone numbers starting with "555"
- Breeze numbers like "12345"
- Spreadsheet name contains "test"

---

*This User Guide covers the complete operational use of the Deacon Visitation Rotation System v1.1. Keep this document handy for quick reference during day-to-day system management.*
