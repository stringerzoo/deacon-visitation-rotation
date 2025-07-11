# Setup Guide - Deacon Visitation Rotation System v1.1

> **Complete installation guide for the Deacon Visitation Rotation System with Google Chat notifications**

**v1.1 Updates**: Menu cleanup with unused functions removed, streamlined interface, enhanced user documentation.

## 📋 Prerequisites

- **Google Workspace account** (free Gmail accounts work for basic features)
- **Google Chat space** for deacon communication
- **Breeze Church Management System** account (optional but recommended)
- **Administrator access** to Google Sheets, Apps Script, and Calendar
- **Basic familiarity** with Google Workspace tools

---

## 🏗️ Step 1: Apps Script Project Setup

### **Create New Apps Script Project**
1. Go to [Google Apps Script](https://script.google.com)
2. Click **"New Project"**
3. Delete the default `Code.gs` file
4. **Rename the project** to "Deacon Visitation Rotation v1.1"

### **Create 5 Module Files**
Create these files by clicking the `+` next to "Files":

```
📁 Your Apps Script Project
├── Module1_Core_Config.gs           (~500 lines)
├── Module2_Algorithm.gs             (~470 lines)
├── Module3_Smart_Calendar.gs        (~200 lines)
├── Module4_Export_Menu.gs           (~400 lines) ⭐ UPDATED v1.1
└── Module5_Notifications.gs         (~1288 lines) ⭐ NEW IN v1.0
```

### **Copy Module Content**
1. **Module1_Core_Config.gs**: Copy from `src/module1-core-config.js`
2. **Module2_Algorithm.gs**: Copy from `src/module2-algorithm.js`
3. **Module3_Smart_Calendar.gs**: Copy from `src/module3-smart-calendar.js`
4. **Module4_Export_Menu.gs**: Copy from `src/module4-export-menu.js` ⭐ **Updated v1.1**
5. **Module5_Notifications.gs**: Copy from `src/module5-notifications.js`

### **v1.1 Menu Changes**
**Removed Functions** (no longer in menu):
- ❌ **"🗓️ Generate Next Year"** - Removed as unlikely to be used (assumes same roster)

**Clarified Functions** (both kept, serve different purposes):
- ✅ **"📋 Test Notification System"** - Tests webhook connectivity with simple message
- ✅ **"🧪 Test Notification Now"** - Tests full notification system with actual weekly summary

**Result**: Cleaner, more focused menu with only essential functions.

### **Save and Authorize**
1. **Save all modules** (Ctrl/Cmd + S)
2. **Run any function** to trigger authorization
3. **Grant required permissions** when prompted
4. **Verify no syntax errors** in any module

---

## 📊 Step 2: Enhanced Spreadsheet Setup (v1.1)

### **Create New Google Spreadsheet**
1. Go to [Google Sheets](https://sheets.google.com)
2. Create a **new blank spreadsheet**
3. **Rename** to "Deacon Visitation Rotation Generator"
4. **Share** with appropriate deacon leadership

### **Connect Apps Script to Spreadsheet**
1. In your spreadsheet: **Extensions → Apps Script**
2. **Delete** the default Code.gs file
3. **Copy the project URL** from your standalone Apps Script project
4. Or copy all 5 modules into this bound project
5. **Save** and authorize when prompted

### **Basic Data Entry**
Add your church information:

#### **Column L: Deacon Names (starting L2)**
```
L2: John Smith
L3: Jane Doe
L4: Mike Johnson
L5: Sarah Wilson
(continue for all deacons)
```

#### **Column M: Household Names (starting M2)**
```
M2: The Anderson Family
M3: Bob & Carol Stevens
M4: Margaret Thompson
M5: The Rodriguez Family
(continue for all households)
```

#### **Column N: Phone Numbers (starting N2)**
```
N2: (502) 555-0123
N3: (502) 555-0456
N4: (502) 555-0789
(matching household order)
```

#### **Column O: Addresses (starting O2)**
```
O2: 123 Main St, Louisville, KY 40203
O3: 456 Oak Ave, Louisville, KY 40205
O4: 789 Elm Dr, Louisville, KY 40207
(matching household order)
```

---

## 🔧 Step 3: Configuration Settings

### **Column K: System Configuration**
The system will create these automatically, but you can customize:

```
K2:  Start Date (Monday of first week)
K4:  Visit Frequency (2-4 weeks typical)
K6:  Schedule Length (26-52 weeks)
K8:  Calendar Instructions (customizable)
K11: Notification Day (dropdown)
K13: Notification Time (0-23 hour)
```

### **Test vs Production Mode**
The system **automatically detects** your environment:

**🧪 Test Mode** (sample data patterns):
- Creates "TEST - Deacon Visitation Schedule" calendar
- Uses test chat webhook
- Red calendar events with "🧪 TEST:" prefixes

**✅ Production Mode** (real names):
- Creates "Deacon Visitation Schedule" calendar  
- Uses production chat webhook
- Blue calendar events, clean formatting

---

## 🔔 Step 4: Google Chat Integration (v1.0)

### **Create Deacon Chat Space**
1. Open **Google Chat**
2. **Create new space** for deacon visitations
3. **Add all deacons** who need notifications
4. **Name the space** (e.g., "Deacon Visitations")

### **Generate Webhook URL**
1. In your chat space: **Click space name**
2. **Manage webhooks → Add webhook**
3. **Name**: "Visitation Notifications"
4. **Copy the webhook URL** (starts with `https://chat.googleapis.com`)

### **Configure in Spreadsheet**
1. **📢 Notifications → 🔧 Configure Chat Webhook**
2. **Paste webhook URL** when prompted
3. **Test**: Use "📋 Test Notification System"

### **Set Up Automation**
1. Configure **K11** (notification day): Sunday-Saturday
2. Configure **K13** (notification time): 0-23 (18 = 6 PM)
3. **📢 Notifications → 🔄 Enable Weekly Auto-Send**
4. **Verify**: Use "📅 Show Auto-Send Schedule"

---

## 📅 Step 5: Calendar Integration

### **Generate Initial Schedule**
1. **📅 Generate Schedule** (creates the rotation)
2. **Review** the schedule in columns A-E
3. **Check** individual deacon reports in columns G-I

### **Export to Google Calendar**
1. **🚨 Full Calendar Regeneration** (first time only)
2. **Grant calendar permissions** when prompted
3. **Wait** for export completion (30-60 seconds)
4. **Verify** calendar creation and events

### **Configure Calendar URL (Optional)**
1. **Open your Google Calendar**
2. **Settings → Calendar settings**
3. **Copy the public/embed URL**
4. **Paste in cell K19** for notification links

---

## 🔗 Step 6: Optional Integrations

### **Breeze CMS Integration**
If you use Breeze Church Management:

1. **Column P**: Add 8-digit Breeze profile numbers
```
P2: 12345678
P3: 23456789
P4: 34567890
(matching household order)
```

2. **🔗 Generate Shortened URLs**: Creates mobile-friendly links
3. **Verify**: Check columns R for shortened Breeze links

### **Visit Notes Integration**  
For Google Docs visit notes:

1. **Create Google Docs** for each household
2. **Column Q**: Add Google Docs URLs
```
Q2: https://docs.google.com/document/d/abc123.../edit
Q3: https://docs.google.com/document/d/def456.../edit
(matching household order)
```

3. **🔗 Generate Shortened URLs**: Creates mobile-friendly links
4. **Verify**: Check columns S for shortened notes links

### **Resource Links (v1.0)**
Configure additional resources in notifications:

- **K19**: Google Calendar embed URL
- **K22**: Visitation Guide URL (procedures, guidelines)
- **K25**: Schedule Summary URL (archived schedules)

---

## 🧪 Step 7: Testing & Validation

### **System Health Check**
1. **🔧 Validate Setup**: Comprehensive configuration check
2. **🧪 Run Tests**: Complete system functionality test
3. **Review results**: Address any reported issues

### **Notification Testing (v1.1 Menu)**
The v1.1 menu includes two distinct notification test functions:

1. **📋 Test Notification System**: 
   - **Purpose**: Simple connectivity test with basic message
   - **Use**: Verify chat webhook is working
   - **Message**: "Notification System Test" with basic info

2. **🧪 Test Notification Now**: 
   - **Purpose**: Full notification content test with actual visit data
   - **Use**: Test complete weekly summary format and content
   - **Message**: Real 2-week visitation summary

3. **💬 Send Weekly Chat Summary**: Manual summary (same content as automation)

**Testing Sequence**:
1. **📋 Test Notification System** first (verify connectivity)
2. **🧪 Test Notification Now** second (verify full content)  
3. **💬 Send Weekly Chat Summary** for manual sends
4. **Verify delivery** in your chat space

**Note**: The "🗓️ Generate Next Year" function was removed in v1.1 as it made unrealistic assumptions about roster continuity.

### **Calendar Testing**
1. **Check calendar events** have proper contact information
2. **Test Breeze links** (if configured)
3. **Test Notes links** (if configured)
4. **Verify mobile compatibility**

---

## 🚀 Step 8: Go Live

### **Final Configuration Review**
- ✅ **Deacon and household lists** complete and accurate
- ✅ **Contact information** current and formatted properly
- ✅ **Notification day/time** set appropriately
- ✅ **Chat webhook** configured and tested
- ✅ **Calendar integration** working properly

### **Enable Automation**
1. **📢 Notifications → 🔄 Enable Weekly Auto-Send**
2. **Confirm settings** in the dialog
3. **📅 Show Auto-Send Schedule** to verify

### **User Training**
1. **Share the [User Guide](USER_GUIDE.md)** with deacon leadership
2. **Demonstrate key functions**: manual notifications, calendar updates
3. **Review troubleshooting**: common issues and solutions
4. **Establish contacts**: Who to call for technical issues

---

## 🔧 Troubleshooting Setup Issues

### **Apps Script Authorization**
**Issue**: "Permission denied" or authorization failures
**Solution**:
1. Go to **script.google.com**
2. **Run any function** to retrigger authorization
3. **Grant all permissions** (Calendar, Sheets, Properties)
4. **Try the operation again**

### **Menu Not Appearing**
**Issue**: Custom menu doesn't show in spreadsheet
**Solution**:
1. **Refresh the spreadsheet** (F5 or Ctrl+R)
2. **Wait 30-60 seconds** for menu to appear
3. **Check Apps Script logs** for errors
4. **Re-run onOpen function** manually if needed

### **Notification Configuration**
**Issue**: Webhook configuration failing
**Solution**:
1. **Verify webhook URL** contains `chat.googleapis.com`
2. **Check chat space permissions** (bot can post)
3. **Test with simple message** first
4. **Review error messages** for specific guidance

### **Calendar Permission Issues**
**Issue**: Cannot create or access calendar
**Solution**:
1. **Google Apps Script permissions**: Ensure calendar access granted
2. **Calendar sharing**: Verify calendar is accessible
3. **API quotas**: Check for Google API rate limiting
4. **Try manual calendar creation** first

---

## 📋 Post-Setup Checklist

### **Week 1: Initial Deployment**
- [ ] System generates schedules successfully
- [ ] Calendar export works properly  
- [ ] Notifications deliver to chat space
- [ ] Deacons can access calendar and links
- [ ] Contact information is accurate

### **Week 2-4: Monitoring**
- [ ] Automated notifications arrive on schedule
- [ ] Deacons are using the system effectively
- [ ] No significant errors or issues
- [ ] Feedback collection and system refinement

### **Monthly: Maintenance**
- [ ] Contact information updates using "📞 Update Contact Info Only"
- [ ] System health check with "🧪 Run Tests"
- [ ] Review notification timing and effectiveness
- [ ] Archive completed schedules as needed

---

## 💡 Pro Tips

### **Deployment Strategy**
1. **Start in test mode** with sample data
2. **Perfect the configuration** before adding real data
3. **Train a backup administrator** before going live
4. **Document any customizations** you make

### **Change Management**
- **Contact updates**: Use "📞 Update Contact Info Only" (safest)
- **Roster changes**: Use "🔄 Update Future Events Only"
- **Major issues**: Use "🚨 Full Calendar Regeneration" (last resort)

### **Backup Strategy**
- **📁 Archive Current Schedule** before major changes
- **Export individual schedules** for external backup
- **Document configuration settings** (K column values)
- **Keep copy of webhook URLs** in secure location

---

*Setup complete! Your Deacon Visitation Rotation System v1.1 is now ready for operation. Refer to the [User Guide](USER_GUIDE.md) for daily operations and troubleshooting guidance.*
