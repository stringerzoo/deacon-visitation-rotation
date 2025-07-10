# Setup Guide - Deacon Visitation Rotation System v25.0

> **Complete installation guide for the Deacon Visitation Rotation System with Google Chat notifications**

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
4. **Rename the project** to "Deacon Visitation Rotation v25.0"

### **Create 5 Module Files**
Create these files by clicking the `+` next to "Files":

```
📁 Your Apps Script Project
├── Module1_Core_Config.gs           (~500 lines)
├── Module2_Algorithm.gs             (~470 lines)
├── Module3_Smart_Calendar.gs        (~200 lines)
├── Module4_Export_Menu.gs           (~400 lines)
└── Module5_Notifications.gs         (~1288 lines) ⭐ NEW
```

### **Copy Module Content**
1. **Module1_Core_Config.gs**: Copy from `src/module1-core-config.js`
2. **Module2_Algorithm.gs**: Copy from `src/module2-algorithm.js`
3. **Module3_Smart_Calendar.gs**: Copy from `src/module3-smart-calendar.js`
4. **Module4_Export_Menu.gs**: Copy from `src/module4-export-menu.js`
5. **Module5_Notifications.gs**: Copy from `src/module5-notifications.js` ⭐

### **Save and Authorize**
1. **Save all modules** (Ctrl/Cmd + S)
2. **Run any function** to trigger authorization
3. **Grant required permissions** when prompted
4. **Verify no syntax errors** in any module

---

## 📊 Step 2: Enhanced Spreadsheet Setup (v25.0)

### **Create New Google Spreadsheet**
1. Go to [Google Sheets](https://sheets.google.com)
2. Create a **new blank spreadsheet**
3. **Rename** to "Deacon Visitation Rotation Generator"
4. **Link to Apps Script** project created in Step 1

### **Enhanced Column Configuration (v25.0)**

#### **Core Data (L-S)**: Start with sample data for testing
```
L1: Deacons                    M1: Households
L2: Andy A                     M2: Alan & Alexa Adams
L3: Brian B                    M3: Barbara Baker
L4: Chris C                    M4: Chloe Cooper
L5: Darell D                   M5: Delilah Danvers
... (continue with your data)

N1: Phone Number               O1: Address
N2: (555) 123-4567            O2: 123 Main St, City, ST 12345
N3: (555) 234-5678            O3: 456 Oak Ave, City, ST 12345
... (continue with contact info)

P1: Breeze Link                Q1: Notes Pg Link
P2: 29760588                   Q2: https://docs.google.com/document/d/abc123/edit
P3: 29760589                   Q3: https://docs.google.com/document/d/def456/edit
... (continue with Breeze numbers and Notes URLs)
```

#### **Configuration (Column K)**: 
```
K1: Start Date                        K2: 7/21/2025
K3: Visits every x weeks              K4: 2
K5: Length of schedule in weeks       K6: 52
K7: Calendar Event Instructions       K8: Please call to confirm...
K9: [Buffer Space]
K10: Weekly Notification Day          K11: Tuesday ⭐ NEW (dropdown)
K12: Weekly Notification Time (0-23)  K13: 16 ⭐ NEW (dropdown)
K14: [Buffer Space]
K15: Current Mode                     K16: [Auto-detected]
K17: [Buffer Space]
K18: Google Calendar URL              K19: [Paste URL here]
K20: [Buffer Space]
K21: Visitation Guide URL             K22: [Paste URL here] (optional)
K23: [Buffer Space]
K24: Schedule Summary Sheet URL       K25: [Paste URL here] (optional) ⭐ NEW
```

💡 **Setup Tip**: Use sample data pattern during initial setup and testing. Replace with real member information only when ready for production use.

---

## 🔔 Step 3: Notifications Setup

### **Create Google Chat Webhook**

#### **1. Set Up Chat Space**
1. **Create or identify** your deacon Google Chat space
2. **Open the space** in Google Chat
3. **Click space name** at the top
4. **Select "Apps & integrations"**
5. **Click "Add webhooks"**
6. **Name the webhook** (e.g., "Deacon Visitation Notifications")
7. **Copy the webhook URL**

#### **2. Configure Webhook in Spreadsheet**
1. **Open your rotation spreadsheet**
2. **Go to menu**: 🔄 Deacon Rotation → 📢 Notifications → 🔧 Configure Chat Webhook
3. **Paste the webhook URL** when prompted
4. **Test the configuration**: 📢 Notifications → 📋 Test Notification System

#### **3. Set Notification Timing**
1. **K11 (Day)**: Select day of week from dropdown (Sunday-Saturday)
2. **K13 (Time)**: Select hour in 24-hour format (0-23)
   - **Examples**: 8 = 8:00 AM, 16 = 4:00 PM, 18 = 6:00 PM

#### **4. Test Mode vs Production**
- **Test Mode**: Automatically detected when using sample data patterns
- **Test Webhook**: Configure separate webhook for testing
- **Production Mode**: Uses real member data and production webhook
- **Mode Switching**: Change data to switch modes automatically

---

## 🗓️ Step 4: Calendar Links Setup ⭐ NEW

### **Configure Google Calendar URL (K19)**

#### **1. Get Your Calendar URL**
1. **Open Google Calendar** in your browser
2. **Navigate to** your deacon visitation calendar
3. **Copy the calendar URL** from the address bar
   - Format: `https://calendar.google.com/calendar/u/0/r/week/...`

#### **2. Configure K19 Cell**
1. **Click on cell K19** in your spreadsheet
2. **Paste the calendar URL**
3. **Press Enter** to save

#### **3. Test/Production Switching**
- **For testing**: Use test calendar URL in K19
- **For production**: Use production calendar URL in K19
- **Easy switching**: Simply change the URL in K19
- **Chat integration**: "📅 View Visitation Calendar" link appears in all notifications

#### **4. Verify Calendar Link**
1. **Test calendar configuration**: 📢 Notifications → Test functions
2. **Send test notification**: Check that calendar link appears
3. **Click the link**: Verify it opens correct calendar

---

## 🔗 Step 5: Additional Resource Links Setup

### **Configure Visitation Guide URL (K21-K22)**
1. **Click on cell K22** in your spreadsheet
2. **Paste your Visitation Guide URL** (Google Doc, website, etc.)
3. **Press Enter** to save
4. **Link appears** as "📋 Visitation Guide" in all chat notifications

### **Configure Schedule Summary URL (K24-K25)**  
1. **Click on cell K25** in your spreadsheet
2. **Paste your Schedule Summary Sheet URL** (archived schedule, summary document, etc.)
3. **Press Enter** to save
4. **Link appears** as "📊 Schedule Summary" in all chat notifications

### **Test/Production Switching**
- **For testing**: Use test document URLs in K22, K25
- **For production**: Use production document URLs
- **Easy switching**: Simply change URLs in K22/K25
- **Optional configuration**: These links are optional and won't appear if not configured

---

## 🎯 Step 6: Generate First Schedule

### **Initial Configuration**
1. **Run header setup**: 🔄 Deacon Rotation → ❓ Setup Instructions
2. **Validate configuration**: 🔄 Deacon Rotation → 🔧 Validate Setup
3. **Generate schedule**: 🔄 Deacon Rotation → 📅 Generate Schedule
4. **Create shortened URLs**: 🔄 Deacon Rotation → 🔗 Generate Shortened URLs

### **Test the System**
1. **Run system tests**: 🔄 Deacon Rotation → 🧪 Run Tests
2. **Send test notification**: 📢 Notifications → 💬 Send Weekly Chat Summary
3. **Verify all features**: Check schedule output, URLs, notifications

---

## 📅 Step 7: Calendar Export & Automation

### **Export to Google Calendar**
1. **Choose calendar function**: 📆 Calendar Functions → 🚨 Full Calendar Regeneration
2. **Monitor progress**: Watch for rate limiting pauses (normal behavior)
3. **Verify events**: Check calendar for proper event creation
4. **Test event details**: Confirm contact info, Breeze links, Notes links

### **Enable Weekly Automation**
1. **Configure timing**: Ensure K11 (day) and K13 (time) are set
2. **Enable auto-send**: 📢 Notifications → 🔄 Enable Weekly Auto-Send
3. **Confirm schedule**: 📢 Notifications → 📅 Show Auto-Send Schedule
4. **Monitor delivery**: Google Apps Script triggers may have 15-20 minute delays

### **Production Best Practices**
1. **Start with sample data** for all initial testing
2. **Test all features** with fake information
3. **Verify notifications work** in test chat space
4. **Replace with real data** when ready for production
5. **Confirm production mode** activation
6. **Update K19** with production calendar URL

---

## 🔧 Step 8: Advanced Configuration

### **Notification Customization**
- **Timing flexibility**: Any day of week, any hour (0-23)
- **Content customization**: Modify message templates in Module 5
- **Multiple environments**: Separate test and production webhooks
- **Error handling**: Built-in diagnostics and retry mechanisms

### **Calendar Enhancement**
- **Breeze integration**: Add 8-digit Breeze numbers in column P
- **Notes integration**: Add Google Docs URLs in column Q
- **Mobile optimization**: Automatic URL shortening for field access
- **Custom instructions**: Personalized messaging in calendar events
- **Calendar links**: Direct access via K19 configuration in notifications

### **Performance Optimization**
- **Large schedules**: System handles 300+ events efficiently
- **Rate limiting**: Built-in protections for Google API limits
- **Batch processing**: Optimized for churches with many deacons/households
- **Error recovery**: Graceful handling of temporary service issues

---

## 🎛️ Step 9: Master the Enhanced Menu System

### **Menu Structure (v25.0):**
```
🔄 Deacon Rotation
├── 📅 Generate Schedule
├── 🔗 Generate Shortened URLs
├── 📆 Calendar Functions                    
│   ├── 📞 Update Contact Info Only         (Safest updates)
│   ├── 🔄 Update Future Events Only        (Current week safe)
│   └── 🚨 Full Calendar Regeneration       (Complete rebuild)
├── 📢 Notifications                         ⭐ NEW SUBMENU
│   ├── 💬 Send Weekly Chat Summary         
│   ├── ⏰ Send Tomorrow's Reminders        
│   ├── 🔧 Configure Chat Webhook           
│   ├── 📋 Test Notification System         
│   ├── 🔄 Enable Weekly Auto-Send          
│   ├── 📅 Show Auto-Send Schedule          
│   ├── 🛑 Disable Weekly Auto-Send         
│   └── [Additional diagnostic tools...]    
├── 📊 Export Individual Schedules
├── 📁 Archive Current Schedule
├── 🗓️ Generate Next Year
├── 🔧 Validate Setup
├── 🧪 Run Tests
├── [🧪/✅] Show Current Mode               
└── ❓ Setup Instructions
```

### **Recommended Usage Flow:**
1. **Initial Setup**: Full Calendar Regeneration
2. **Weekly Management**: Automated notifications via chat
3. **Contact Updates**: Update Contact Info Only  
4. **Planning Ahead**: Update Future Events Only (if needed)
5. **Major Changes**: Full Calendar Regeneration (with caution)
6. **Troubleshooting**: Use diagnostic tools in Notifications submenu

---

## 🆘 Troubleshooting

### **Common Notification Issues**

#### **Notifications Not Sending**
1. **Check webhook URL**: 📢 Notifications → 🔧 Configure Chat Webhook
2. **Verify permissions**: Apps Script needs external URL access
3. **Test manually**: 📢 Notifications → 💬 Send Weekly Chat Summary
4. **Check Google Chat space**: Ensure webhook is active

#### **Wrong Timing**
1. **Verify K11 and K13 values**: Check day name and hour format
2. **Google Apps Script delays**: Triggers can be 15-20 minutes late
3. **Timezone settings**: Check Apps Script project timezone
4. **Recreate triggers**: 🛑 Disable then 🔄 Enable weekly auto-send

#### **Test Mode Issues**
1. **Check data patterns**: Sample data triggers test mode
2. **Verify mode indicator**: K16 should show current mode
3. **Manual mode refresh**: Run "Show Current Mode" from menu
4. **Test webhook separation**: Ensure test/production webhooks configured

#### **Calendar Link Issues**
1. **Check K19, K22, K25 values**: Ensure valid URLs are pasted
2. **URL format**: Should start with `https://` for web resources
3. **Test resource access**: Click links to verify they open correct documents
4. **Mode switching**: Different URLs for test vs production resources

### **Google Apps Script Limitations**
- **Trigger reliability**: New triggers may not fire immediately (24-48 hour "settling")
- **Execution timing**: Can be 15-20+ minutes late
- **Service dependencies**: Occasional Google service interruptions
- **Quota limits**: Daily execution limits for complex operations

### **Calendar Integration Issues**
1. **Permission errors**: Re-authorize Apps Script permissions
2. **Calendar access**: Verify shared calendar permissions
3. **Event formatting**: Check Breeze numbers and Notes URLs
4. **Test mode separation**: Confirm using correct calendar (test vs. production)

---

## 🎉 You're Ready!

Your Deacon Visitation Rotation System v25.0 is now fully configured with:

✅ **Automated scheduling** with mathematical fairness  
✅ **Google Chat notifications** with rich content and resource links  
✅ **Smart calendar updates** preserving deacon customizations  
✅ **Test/production separation** for safe development  
✅ **Comprehensive resource integration** with calendar, guide, and summary access  

**Next Steps**: Monitor your first week of automated notifications and fine-tune timing as needed!

---

*For additional help, consult the [Features Guide](FEATURES.md) or review the [Changelog](CHANGELOG.md) for version-specific details.*
