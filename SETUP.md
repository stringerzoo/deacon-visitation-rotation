# Setup Guide - Deacon Visitation Rotation System v2.0

> **Complete installation and configuration guide with variable frequency support**

---

## 📋 Table of Contents

1. [Initial Setup](#initial-setup)
2. [Basic Configuration](#basic-configuration)
3. [Variable Frequency Setup (v2.0)](#variable-frequency-setup-v20)
4. [Contact Information](#contact-information)
5. [Integration Links](#integration-links)
6. [Google Chat Notifications](#google-chat-notifications)
7. [Calendar Setup](#calendar-setup)
8. [Testing & Validation](#testing--validation)
9. [Production Deployment](#production-deployment)
10. [Troubleshooting](#troubleshooting)

---

## 🎯 Initial Setup

### **Step 1: Create Your Spreadsheet**

1. Open [Google Sheets](https://sheets.google.com)
2. Create a new blank spreadsheet
3. Name it: `Deacon Visitation Rotation Generator` (or similar)
4. Note the folder location for later use

### **Step 2: Install Apps Script Code**

1. In your spreadsheet, click **Extensions → Apps Script**
2. Delete any default code in the editor
3. Create 5 script files (File → New → Script file):
   - `Module1_Core_Config.gs`
   - `Module2_Algorithm.gs`
   - `Module3_Smart_Calendar.gs`
   - `Module4_Export_Menu.gs`
   - `Module5_Notifications.gs`

4. Copy the code from each module file into the corresponding script file
5. Click **Save** (disk icon)
6. Return to your spreadsheet
7. Refresh the page - you should see a new menu: **🔄 Deacon Rotation**

---

## ⚙️ Basic Configuration

### **Step 3: Column K - Core Settings**

Configure these settings in Column K:

| Cell | Label | Cell | Your Value |
|------|-------|------|------------|
| K1 | Start Date | K2 | `01/06/2025` (Monday preferred) |
| K3 | Visits every x weeks (1,2,3,4) **(default)** 🆕 | K4 | `2` (or 1, 3, 4) |
| K5 | Length of schedule in weeks | K6 | `26` (or your preference) |
| K7 | Calendar Event Instructions: | K8 | Custom instructions text |

**Notes:**
- **K2 (Start Date)**: Should be a **Monday** for best results
- **K4 (Default Frequency)**: This is the default for ALL households unless overridden in Column T
- **K3 label change**: Now says "**(default)**" to clarify this is the baseline frequency
- **K6 (Schedule Length)**: Typically 26 or 52 weeks
- **K8 (Instructions)**: Appears in every calendar event

### **Step 4: Notification Settings**

| Cell | Label | Cell | Your Value |
|------|-------|------|------------|
| K10 | Weekly Notification Day: | K11 | Select day from dropdown |
| K12 | Weekly Notification Time (0-23): | K13 | `18` (for 6 PM) |

**Note**: Dropdowns are auto-created when you first run "Generate Schedule"

### **Step 5: Resource URLs**

| Cell | Label | Cell | Your Value |
|------|-------|------|------------|
| K18 | Google Calendar URL | K19 | Auto-detected or paste URL |
| K21 | Visitation Guide URL | K22 | Paste Google Doc URL |
| K24 | Schedule Summary URL | K25 | Auto-populated by summary generator 🆕 |

**Notes:**
- **K19**: Auto-detected from your calendar, or manually paste
- **K22**: Link to your visitation guide document
- **K25**: Automatically updated when you run "Generate Schedule Summary Sheet" 🆕

---

## 🆕 Variable Frequency Setup (v2.0)

### **Step 6: Column T - Custom Frequencies**

**This is the major new feature in v2.0!**

#### **Column T Header**
The system auto-creates Column T with:
- **Header (T1)**: "Custom visit frequency (every 1, 2, 3, or 4 weeks)"
- **Data validation**: Dropdown with values 1, 2, 3, 4
- **Light yellow background**: Makes it visually distinct

#### **How to Use Column T**

**Default Behavior (Leave Blank):**
```
Column M (Households)    Column T (Custom Frequency)
Alan & Alexa Adams       [blank]              ← Uses K4 default (e.g., 2 weeks)
Barbara & Bob Baker      [blank]              ← Uses K4 default
Chloe & Charles Copper   [blank]              ← Uses K4 default
```

**Custom Frequencies (Override Default):**
```
Column M (Households)    Column T (Custom Frequency)
Alan & Alexa Adams       [blank]              ← Uses default (2 weeks)
Barbara & Bob Baker      4                    ← Monthly visits
Chloe & Charles Copper   [blank]              ← Uses default (2 weeks)
Delilah & David Danvers  3                    ← Every 3 weeks
Emma & Edward Evans      1                    ← Weekly visits
```

#### **When to Use Custom Frequencies**

**Use 1 week (weekly) for:**
- Households with urgent or intensive care needs
- Recent hospital discharge or major life events
- Temporary increased support situations

**Use 2 weeks (bi-weekly) - TYPICAL DEFAULT:**
- Standard pastoral care schedule
- Most healthy, stable households
- Balanced contact frequency

**Use 3 weeks for:**
- Households preferring slightly less frequent contact
- Transitioning from intensive to standard care
- Balancing workload with other commitments

**Use 4 weeks (monthly) for:**
- Households preferring minimal contact
- Very independent members
- Geographical distance considerations
- Seasonal schedules (snowbirds, travelers)

#### **Important Notes**

✅ **Column T is completely optional** - leave it empty for v1.1 behavior
✅ **System auto-creates Column T** - no manual setup needed
✅ **Dropdown validation prevents errors** - can't enter invalid values
✅ **Visual indicators** - custom frequency households are highlighted in light yellow
✅ **Backward compatible** - old spreadsheets work without any changes

---

## 👥 Step 7: Lists Setup

### **Deacons List (Column L)**
```
L1: Deacons
L2: Andy A
L3: Brian B
L4: Chris C
L5: David D
...
```

**Tips:**
- Use full names or recognizable initials
- Order doesn't matter - algorithm handles distribution
- Can add/remove deacons anytime

### **Households List (Column M)**
```
M1: Households
M2: Alan & Alexa Adams
M3: Barbara & Bob Baker
M4: Chloe & Charles Copper
M5: Delilah & David Danvers
...
```

**Tips:**
- Use consistent naming format
- Include both names for couples
- Match exactly with Breeze if using integration

---

## 📞 Step 8: Contact Information

### **Phone Numbers (Column N)**
```
N1: Phone Number
N2: (555) 123-1001
N3: (555) 123-1002
N4: (555) 123-1003
...
```

### **Addresses (Column O)**
```
O1: Address
O2: 123 Maple Street, Louisville, KY 40202
O3: 456 Oak Avenue, Louisville, KY 40203
...
```

**Note**: These appear in calendar events and chat notifications

---

## 🔗 Step 9: Integration Links (Optional but Recommended)

### **Breeze Profile Numbers (Column P)**
```
P1: Breeze Profile (8-digit number)
P2: 12345001
P3: 12345002
P4: 12345003
...
```

**Breeze URL Format**: `https://immanuelky.breezechms.com/people/view/[number]`

### **Visit Notes Links (Column Q)**
```
Q1: Notes Page Link
Q2: [Google Doc URL for Adams family notes]
Q3: [Google Doc URL for Baker family notes]
Q4: [Google Doc URL for Copper family notes]
...
```

**Tip**: Create a Google Doc for each household's visit notes, then paste the sharing link

### **Auto-Generated Short URLs (Columns R-S)**

The system automatically creates shortened URLs:
- **Column R**: Shortened Breeze profile links
- **Column S**: Shortened visit notes links

Run **🔗 Generate Shortened URLs** from menu to create these.

**Note**: Only generates for empty cells - preserves existing URLs

---

## 📢 Step 10: Google Chat Notifications

### **Creating a Chat Webhook**

1. **Create Google Chat Space:**
   - Open Google Chat
   - Click **+ Create space**
   - Name it: "Deacon Visitation Notifications"
   - Add all deacons as members

2. **Generate Webhook URL:**
   - In the space, click the space name → **Apps & integrations**
   - Click **Add webhooks**
   - Name: "Visitation Rotation Bot"
   - Click **Save**
   - **Copy the webhook URL** (starts with `https://chat.googleapis.com/...`)

3. **Configure in Spreadsheet:**
   - Menu: **📢 Notifications → 🔧 Configure Chat Webhook**
   - Paste the webhook URL
   - Click OK

4. **Test the Connection:**
   - Menu: **📢 Notifications → 📋 Test Notification System**
   - Verify message appears in chat space

5. **Enable Automation:**
   - Menu: **📢 Notifications → 🔄 Enable Weekly Auto-Send**
   - Confirms using K11 (day) and K13 (time) settings
   - Creates automated trigger

### **Separate Test & Production Webhooks**

For safer testing, create two chat spaces:
- **Test Space**: "TEST - Deacon Notifications" (smaller group)
- **Production Space**: "Deacon Visitation Notifications" (all deacons)

Configure different webhooks using **Script Properties** in Apps Script.

---

## 📅 Step 11: Calendar Setup

### **Option 1: Auto-Detection (Recommended)**

The system automatically:
- Detects if you're in test mode (sample data)
- Creates/finds the appropriate calendar
- Populates K19 with the calendar URL

### **Option 2: Manual Setup**

1. **Create Calendar:**
   - Google Calendar → Settings → Add calendar → Create new calendar
   - Name: "Deacon Visitation Schedule"
   - Save

2. **Get Calendar URL:**
   - Click calendar name → Settings
   - Scroll to "Integrate calendar"
   - Copy the **Public URL**
   - Paste into K19

### **Export Schedule to Calendar**

1. Generate your schedule first
2. Menu: **📆 Calendar Functions → 🚨 Full Calendar Regeneration**
3. Confirm the operation
4. Events appear in calendar with:
   - Title: `Deacon ➡︎ Household`
   - Visit Notes link at top
   - Frequency information
   - Contact details
   - Instructions

---

## 📊 Step 12: Generate Schedule Summary Sheet

### **Creating Shareable Schedule**

1. **Generate Main Schedule:**
   - Menu: **📅 Generate Schedule**
   - Review the output

2. **Create Summary Sheet:**
   - Menu: **📊 Generate Schedule Summary Sheet** 🆕
   - Confirm creation

3. **Auto-Generated Files:**
   - Standalone spreadsheet (columns A-I only)
   - QR code PNG image (500×500px)
   - Both saved in same folder

4. **Update K25:**
   - Dialog asks: "Update K25 with new URL?"
   - Select **YES** to auto-populate K25
   - URL used in notification messages

5. **Share with Elders/Deacons:**
   - Open the summary spreadsheet
   - Click **Share**
   - Add viewers (read-only recommended)
   - Use QR code for print materials

**File Naming:**
```
Visitation Schedule - 2025-01-06 (2025-01-06_1445)
Visitation Schedule - 2025-01-06 (2025-01-06_1445) - QR Code.png
```

**Uses:**
- Insert QR code in monthly bulletins
- Share spreadsheet link in group emails
- Post to church website or app
- Print for physical distribution

---

## ✅ Step 13: Testing & Validation

### **Test Mode Features**

System automatically detects test mode when:
- Household names contain "Adams", "Baker", "Copper" (test names)
- Phone numbers start with "555"
- Breeze numbers are "12345xxx"
- Spreadsheet name contains "test" or "sample"

**Test Mode Indicators:**
- Calendar: "TEST - Deacon Visitation Schedule" (red)
- Events: "TEST:" prefix
- K16: Shows "🧪 TEST MODE"

### **Validation Checklist**

Run through this checklist before production:

- [ ] **Configuration validated**: Menu → 🔧 Validate Setup
- [ ] **Schedule generates**: No errors in console
- [ ] **Deacon reports appear**: Columns G-I populated
- [ ] **Household reports appear**: Below deacon reports 🆕
- [ ] **Custom frequencies work**: Check Column T highlighting 🆕
- [ ] **Calendar exports**: Events created correctly
- [ ] **Notifications send**: Test message received
- [ ] **Summary sheet creates**: Standalone file + QR code 🆕
- [ ] **K25 populates**: URL appears after summary generation 🆕
- [ ] **Quality validation passes**: No warnings in console 🆕

### **Run System Tests**

Menu: **🧪 Run Tests**
- Validates all configuration
- Checks data integrity
- Reports any issues

---

## 🚀 Step 14: Production Deployment

### **Migration Checklist**

1. **Backup Current System:**
   - Download existing schedules
   - Archive current configuration
   - Document any customizations

2. **Replace Test Data:**
   - Update Column M with real household names
   - Update Column N-O with real contact info
   - Update Column P-Q with real Breeze/notes links
   - **Add Column T frequencies** if needed 🆕

3. **Update Configuration:**
   - K2: Real start date
   - K4: Desired default frequency
   - K11/K13: Real notification schedule
   - K19/K22/K25: Production URLs

4. **Configure Production Webhook:**
   - Create production chat space
   - Generate production webhook
   - Test with production webhook

5. **Generate Production Schedule:**
   - Run: **📅 Generate Schedule**
   - Review output carefully
   - Check quality validation results 🆕

6. **Export to Calendar:**
   - Run: **🚨 Full Calendar Regeneration**
   - Verify events in production calendar

7. **Create Schedule Summary:**
   - Run: **📊 Generate Schedule Summary Sheet**
   - Update K25: Yes
   - Share with appropriate people
   - Use QR code in materials

8. **Enable Automation:**
   - Run: **🔄 Enable Weekly Auto-Send**
   - Verify trigger creation

9. **Monitor First Week:**
   - Confirm notification sends
   - Check for any errors
   - Gather feedback from deacons

---

## 🔧 Troubleshooting

### **Common Issues**

#### **"Script function not found" Error**
- **Cause**: Code not properly saved or named incorrectly
- **Fix**: Verify all 5 modules are present and saved

#### **Calendar Events Missing**
- **Cause**: Permission issues or wrong calendar
- **Fix**: 
  - Check K19 has correct calendar URL
  - Run: **📆 Update Future Events Only** to refresh

#### **Year Rollover Issues**
- **Cause**: Date format in spreadsheet
- **Fix**: v2.0 includes automatic fix - dates show full year (MM/DD/YYYY) 🆕

#### **Notifications Not Sending**
- **Cause**: Webhook configuration or trigger issues
- **Fix**:
  - Menu: **📋 Test Notification System**
  - Menu: **🔍 Inspect All Triggers**
  - Recreate webhook if needed

#### **Custom Frequencies Not Working**
- **Cause**: Column T not properly configured
- **Fix**: 
  - Delete Column T entirely
  - Run: **📅 Generate Schedule** (auto-recreates Column T)
  - Add custom frequencies again

#### **QR Code Generation Fails**
- **Cause**: Network connectivity or API issue
- **Fix**:
  - Check internet connection
  - Retry generation
  - QR code can be generated manually if needed

#### **Phantom Calendar Events**
- **Cause**: Legend rows being read as schedule data
- **Fix**: v2.0 includes improved filtering - should not occur 🆕

### **Getting Help**

1. **Check Console Logs:**
   - Apps Script Editor → Executions
   - Look for error details

2. **Run Diagnostics:**
   - Menu: **🧪 Run Tests**
   - Menu: **🔍 Inspect All Triggers**

3. **Review Documentation:**
   - [README](README.md) - Feature overview
   - [CHANGELOG](CHANGELOG.md) - Version history
   - [FEATURES](FEATURES.md) - Technical details

---

## 📝 Quick Reference

### **Essential Menu Items**

| Menu Item | Purpose | Frequency |
|-----------|---------|-----------|
| 📅 Generate Schedule | Create rotation | When updating schedule |
| 📊 Generate Schedule Summary Sheet 🆕 | Create shareable file + QR | After schedule changes |
| 🚨 Full Calendar Regeneration | Rebuild all events | Major changes only |
| 📞 Update Contact Info Only | Refresh contact data | After contact updates |
| 💬 Send Weekly Chat Summary | Manual notification | Testing or missed auto-send |
| 🔄 Enable Weekly Auto-Send | Start automation | Initial setup |

### **File Locations**

- **Main Spreadsheet**: Your Google Drive folder
- **Schedule Summary**: Same folder as main spreadsheet 🆕
- **QR Code**: Same folder as schedule summary 🆕
- **Calendar**: Google Calendar (separate service)
- **Chat Space**: Google Chat (separate service)

---

## 🎉 You're Ready!

Your v2.0 Deacon Visitation Rotation System is now configured with:

✅ Variable frequency scheduling per household
✅ Enhanced anti-repetition algorithm  
✅ Automatic quality validation
✅ Household-centric reports
✅ Schedule summary generator with QR codes
✅ Automated notifications
✅ Complete calendar integration

**Next Steps:**
- Generate your first production schedule
- Create and share your schedule summary
- Enable weekly notifications
- Monitor and adjust as needed

---

*For additional support, refer to the [User Guide](USER_GUIDE.md) or [Lead Deacon Handoff Guide](docs-handoff_document_comprehensive.md).*
