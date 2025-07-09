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
└── Module5_Notifications.gs         (~300 lines) ⭐ NEW
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
4. **Connect Apps Script** project:
   - Extensions → Apps Script
   - Select your project or paste the script ID

### **Enhanced Column Layout**
The v25.0 system uses this organized layout:

#### **Main Schedule (A-E)** - Auto-generated
- **A**: Cycle, **B**: Week, **C**: Week of, **D**: Household, **E**: Deacon

#### **Reports & Buffers (F-J)**
- **F**: Buffer column (leave empty)
- **G**: Deacon, **H**: Week of, **I**: Household (auto-generated reports)
- **J**: Buffer column (leave empty)

#### **Configuration (K)** - ⭐ Enhanced for v25.0
```
K1:  Start Date                        K2:  [Your start date]
K3:  Visits every x weeks (1,2,3,4)    K4:  2
K5:  Length of schedule in weeks       K6:  52
K7:  Calendar Event Instructions:      K8:  [Custom instructions]
K9:  [Empty - Buffer Space]
K10: Weekly Notification Day:          K11: Sunday        ⭐ NEW
K12: Weekly Notification Time (0-23):  K13: 18           ⭐ NEW
K14: [Empty - Buffer Space]
K15: Current Mode:                     K16: [Auto-detected] ⭐ MOVED
```

#### **Contact Data (L-S)**
- **L**: Deacons, **M**: Households, **N**: Phone Numbers, **O**: Addresses
- **P**: Breeze Links, **Q**: Notes Page Links
- **R**: Breeze Links (short), **S**: Notes Page Links (short)

### **Initial Data Entry**
**Start with sample data for testing:**

#### **Configuration (K2, K4, K6, K8, K11, K13)**
```
K2:  Next Monday's date
K4:  2 (bi-weekly visits)
K6:  52 (one year)
K8:  Please call to confirm visit time. Contact family 1-2 days before.
K11: Sunday (or your preferred notification day)
K13: 18 (6 PM, or your preferred hour in 24-hour format)
```

#### **Sample Deacons (L2-L6)**
```
Andy B
Mike B  
Kris B
Cody G
Jim H
```

#### **Sample Households (M2-M6)**
```
Alan & Alexa Adams
Barbara Baker
Chloe Cooper
David Danvers
Emma Evans
```

#### **Sample Contact Info (N2-O6)**
Add sample phone numbers and addresses for testing.

---

## 🔔 Step 3: Google Chat Notifications Setup ⭐ NEW

### **Create Google Chat Webhook**

#### **For Google Workspace (Recommended)**
1. **Open Google Chat** in your browser
2. **Go to your deacon chat space**
3. **Click the space name** at the top
4. **Select "Apps & integrations"**
5. **Click "Add webhooks"**
6. **Name the webhook**: "Deacon Visitation Notifications"
7. **Click "Save"** and **copy the webhook URL**

#### **For Personal Gmail Accounts**
1. **Create a Google Chat space** for deacon communication
2. **Follow the same webhook creation process**
3. **Invite all deacons** to the chat space
4. **Test webhook access** with a simple message

### **Configure Webhook in Apps Script**
1. **Go to your Apps Script project**
2. **Run the menu**: 🔄 Deacon Rotation → 📢 Notifications → 🔧 Configure Chat Webhook
3. **Paste your webhook URL** when prompted
4. **Set up test webhook** (optional but recommended):
   - Create a separate test chat space
   - Generate another webhook URL
   - Configure as test webhook for development

### **Test Notification System**
1. **Run**: 📢 Notifications → 📋 Test Notification System
2. **Verify message appears** in your Google Chat space
3. **Check formatting and links** work correctly
4. **Confirm test/production mode** separation if using both

---

## 📅 Step 4: Calendar Integration Setup

### **Smart Calendar Export**
1. **Generate your first schedule**: 🔄 Deacon Rotation → 📅 Generate Schedule
2. **Create shortened URLs**: 🔄 Deacon Rotation → 🔗 Generate Shortened URLs
3. **Export to calendar**: 📆 Calendar Functions → 🚨 Full Calendar Regeneration

### **Test Smart Calendar Features**
1. **Update Contact Info Only**: Preserves all scheduling details
2. **Update Future Events Only**: Protects current week planning
3. **Verify test mode detection**: Check if using test calendar vs. production

---

## ⚙️ Step 5: Notification Automation Setup

### **Configure Weekly Automation**
1. **Set notification timing** in spreadsheet:
   - **K11**: Day of week (Sunday, Monday, etc.)
   - **K13**: Hour in 24-hour format (18 = 6 PM)

2. **Enable automation**: 📢 Notifications → 🔄 Enable Weekly Auto-Send
   - **Review timing confirmation**
   - **Verify webhook destination**
   - **Confirm automation setup**

3. **Monitor automation**: 📢 Notifications → 📅 Show Auto-Send Schedule
   - **Check trigger status**
   - **Verify next execution time**
   - **Review configuration details**

### **Test Automation (Optional)**
1. **Set test time** (e.g., 15 minutes from now)
2. **Enable automation** with test timing
3. **Wait for notification** to verify delivery
4. **Reset to production timing**

---

## 🧪 Step 6: Test Mode vs Production

### **Understanding Test Mode**
The system automatically detects test mode based on data patterns:
- **Test indicators**: Sample names (Alan Adams, Barbara Baker), test phone numbers (555-xxx-xxxx)
- **Automatic switching**: Changes to production mode with real member data
- **Visual indicators**: K16 shows current mode (🧪 TEST MODE / ✅ PRODUCTION)

### **Safe Testing Workflow**
1. **Start with sample data** for all initial testing
2. **Test all features** with fake information
3. **Verify notifications work** in test chat space
4. **Replace with real data** when ready for production
5. **Confirm production mode** activation

---

## 🔧 Step 7: Advanced Configuration

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

### **Performance Optimization**
- **Large schedules**: System handles 300+ events efficiently
- **Rate limiting**: Built-in protections for Google API limits
- **Batch processing**: Optimized for churches with many deacons/households
- **Error recovery**: Graceful handling of temporary service issues

---

## 🎛️ Step 8: Master the Enhanced Menu System

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

## 📱 Mobile Access

### **Google Chat on Mobile**
- **Install Google Chat app** for real-time notifications
- **Enable push notifications** for the deacon chat space
- **Bookmark key links** (Breeze, Notes folders) for quick access
- **Test mobile access** to ensure links work on phones/tablets

### **Calendar Access**
- **Google Calendar app** for viewing and updating visit times
- **Offline sync** for accessing contact info without internet
- **Personal calendar integration** for private reminders
- **Quick editing** of visit times from mobile devices

---

## 🎯 Production Deployment Checklist

### **Before Going Live:**
- [ ] **All sample data replaced** with real member information
- [ ] **Production webhook configured** and tested
- [ ] **Calendar permissions verified** for all deacons
- [ ] **Notification timing confirmed** (K11, K13 values)
- [ ] **Test mode indicator shows** ✅ PRODUCTION in K16
- [ ] **Weekly automation enabled** with correct timing
- [ ] **Deacons informed** about new notification system
- [ ] **Backup schedule exported** (📁 Archive Current Schedule)

### **First Week Monitoring:**
- [ ] **Verify notification delivery** on scheduled day/time
- [ ] **Check message formatting** and link functionality
- [ ] **Monitor Google Chat space** for deacon questions
- [ ] **Test calendar event access** from mobile devices
- [ ] **Confirm automation continues** working week 2
- [ ] **Document any issues** for future troubleshooting

---

## 🔄 Development and Maintenance Workflow

### **Module-Based Updates:**
When you need to make changes:
1. **Identify the module** containing the functionality you want to modify
2. **Edit in GitHub** (or directly in Apps Script for quick fixes)
3. **Copy just that module** to the corresponding .gs file
4. **Test the specific functionality**
5. **Update other modules** only if needed

### **Module Responsibilities Quick Reference:**
- **Module 1**: Configuration, validation, header setup
- **Module 2**: Core algorithm, schedule generation, deacon reports  
- **Module 3**: Smart calendar updates, contact preservation, test mode detection
- **Module 4**: Full calendar export, individual schedules, menu system
- **Module 5**: Google Chat notifications, triggers, webhook management ⭐

### **Regular Maintenance Tasks:**
- **Weekly**: Monitor notification delivery and deacon feedback
- **Monthly**: Review shortened URLs in columns R-S for any failures
- **Quarterly**: Run system tests to verify all integrations
- **Yearly**: Archive previous schedules and generate next year
- **As needed**: Update webhook URLs if chat spaces change

---

## 🔍 Advanced Diagnostic Tools

### **Notification System Diagnostics**
Available in 📢 Notifications submenu:

- **🧪 Test Notification Now**: Immediate test message
- **🔍 Inspect All Triggers**: Shows all active triggers
- **🔄 Force Recreate Trigger**: Rebuilds weekly automation
- **🐛 Debug Trigger Creation**: Detailed troubleshooting
- **🕐 Check Timezone Settings**: Verifies Apps Script timezone

### **System Health Monitoring**
- **Configuration validation**: Automatic checks for common issues
- **Webhook connectivity**: Real-time testing of chat integration
- **Calendar permissions**: Verification of access rights
- **URL shortening**: Service availability and quota monitoring

---

## 💡 Tips for Success

### **Best Practices:**
1. **Start small**: Begin with 3-5 deacons and households for initial testing
2. **Use test mode**: Always test new features with sample data first
3. **Monitor timing**: Google's trigger delays are normal (15-20+ minutes)
4. **Keep webhooks secure**: Don't share webhook URLs outside your team
5. **Regular backups**: Use Archive function before major changes

### **Deacon Training:**
1. **Show notification format**: Walk through a sample chat message
2. **Explain timing**: When notifications arrive and what they contain
3. **Demonstrate links**: How to access Breeze profiles and Notes
4. **Mobile setup**: Help with Google Chat app installation
5. **Feedback collection**: Gather input for system improvements

### **Common Gotchas:**
- **Trigger delays**: New weekly triggers need 24-48 hours to stabilize
- **Test data**: Using sample names keeps system in test mode
- **Webhook changes**: Moving chat spaces requires webhook reconfiguration
- **Timezone confusion**: Apps Script uses project timezone, not local time
- **Calendar permissions**: Shared calendars need explicit access for all deacons

---

## 🚀 Next Steps After Setup

### **Week 1: Initial Deployment**
1. **Generate first schedule** with real member data
2. **Export to calendar** and verify all deacons can access
3. **Send first notification** manually to verify delivery
4. **Enable weekly automation** with appropriate timing
5. **Monitor deacon feedback** and adjust as needed

### **Week 2-4: Optimization**
1. **Fine-tune notification timing** based on deacon preferences
2. **Add Breeze numbers** and Notes links for enhanced functionality
3. **Train deacons** on smart calendar features
4. **Test smart update functions** with contact information changes
5. **Document local customizations** for future reference

### **Ongoing: Maintenance Mode**
1. **Monitor weekly notifications** for consistent delivery
2. **Update contact information** using smart calendar functions
3. **Archive completed schedules** quarterly or yearly
4. **Review and update** deacon and household lists as needed
5. **Share improvements** with the broader church technology community

---

**🎉 Congratulations! Your Deacon Visitation Rotation System v25.0 is now ready to transform your church's pastoral care coordination with intelligent scheduling and automated notifications!**

> **Remember**: The system prioritizes reliability and member privacy while providing modern automation tools. The built-in delays and smart update options ensure your calendar operations complete successfully while preserving the personal touch that makes deacon ministry meaningful.

For additional support:
- 📖 **[Features Documentation](FEATURES.md)** - Technical implementation details
- 📝 **[Changelog](CHANGELOG.md)** - Version history and updates  
- 🆘 **[GitHub Issues](https://github.com/your-repo/issues)** - Community support and bug reports
