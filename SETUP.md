# Setup Guide - Deacon Visitation Rotation System v24.2

This comprehensive guide covers setup, smart calendar features, and detailed calendar export timing information for the enhanced deacon visitation rotation system.

## 📋 Prerequisites

Before starting, ensure you have:
- Google Workspace account (free Gmail works)
- Access to Google Sheets and Google Apps Script
- List of deacon names/initials
- List of household names to visit
- Contact information for households (phone numbers and addresses)
- **Breeze Church Management System access** (for profile integration)
- **Google Docs for visit notes** (optional but recommended)
- **Internet connectivity** (for URL shortening service)

## 🧪 Understanding Intelligent Test Mode Detection (NEW v24.2)

The system now automatically detects whether you're working with test data or live member data:

### **How It Works:**
- **Analyzes your spreadsheet data** for test patterns
- **Automatically switches** between test and production calendars
- **No manual configuration** required

### **Test Mode Triggers:**
- ✅ Household names like "Alan Adams", "Barbara Baker", "Chloe Cooper"
- ✅ Phone numbers with "555" prefix  
- ✅ Breeze numbers starting with "12345"
- ✅ Spreadsheet name containing "test" or "sample"

### **Visual Indicators:**
- **Cell K11**: "Current Mode:" label
- **Cell K12**: "🧪 TEST MODE" (yellow) or "✅ PRODUCTION" (green)
- **Menu**: Shows 🧪 or ✅ icon next to "Show Current Mode"
- **Dialogs**: All confirmations indicate current mode

### **Benefits:**
- **Safe testing** - prevents accidental live calendar modifications
- **Automatic transition** - switches to production when you use real data
- **Clear communication** - always know which environment you're in

> 💡 **Tip**: Start with the sample data provided in this guide. The system will automatically use test mode, then seamlessly switch to production when you replace with real member information.

## 🔒 Data Privacy & Sample Data Setup

### **Start with Sample Data**
**Important**: Always begin setup and testing with fake sample data to protect member privacy.

#### **Recommended Sample Data Pattern:**
```
Deacons: Andy A, Brian B, Chris C, Darell D
Households: Alan & Alexa Adams, Barbara Baker, Chloe Cooper, Delilah Danvers
Phone: (555) 123-XXXX format
Addresses: Use generic examples or your church's public address
```
#### **Why Sample Data First?**
- **Test all features safely** without exposing real member information
- **Verify smart calendar functions** before adding sensitive data
- **Share for troubleshooting** without privacy concerns
- **Train users** on the new update options with non-sensitive examples

> ⚠️ **Privacy Note**: Complete all testing with sample data before adding real member information. This protects your congregation's privacy during setup and troubleshooting.

## 🏗️ Step 1: Set Up Native Modular Architecture (v24.2)

### **Google Apps Script Project Setup**

1. **Go to Google Apps Script**: [script.google.com](https://script.google.com)
2. **Click "New Project"**
3. **Delete the default `Code.gs` file**
4. **Create 4 separate files** by clicking the `+` next to "Files":

#### **File Structure:**
```
📁 Your Apps Script Project
├── Module1_Core_Config.gs          (~500 lines)
├── Module2_Algorithm.gs             (~470 lines)
├── Module3_Smart_Calendar.gs        (~200 lines)
└── Module4_Export_Menu.gs           (~400 lines)
```

### **Copy Module Content**
1. **Module1_Core_Config.gs**: Copy content from `src/module1-core-config.js`
2. **Module2_Algorithm.gs**: Copy content from `src/module2-algorithm.js`
3. **Module3_Smart_Calendar.gs**: Copy content from `src/module3-smart-calendar.js`
4. **Module4_Export_Menu.gs**: Copy content from `src/module4-export-menu.js`

### **Configuration Note**
The system will automatically detect test mode based on your data - no manual configuration needed!

## 📊 Step 2: Create Enhanced Google Spreadsheet Layout

### **Enhanced Column Structure (v24.2):**

#### **A-E: Main Rotation Schedule** (Auto-generated)
- **A**: Cycle, **B**: Week, **C**: Week of, **D**: Household, **E**: Deacon

#### **F: Buffer Column** (Leave empty)

#### **G-I: Individual Deacon Reports** (Auto-generated)
- **G**: Deacon, **H**: Week of, **I**: Household

#### **J: Buffer Column** (Leave empty)

#### **K: Configuration Settings** (Your input required)
```
K1: Start Date
K2: 7/7/2025 (or your preferred start date)
K3: Visits every x weeks (1,2,3,4)
K4: 2 (for bi-weekly visits)
K5: Length of schedule in weeks  
K6: 52 (for one year)
K7: Calendar Event Instructions:
K8: Please call to confirm visit time. Contact family 1-2 days before scheduled date to arrange convenient time.
```

#### **L-M: Core Data Lists** (Start with sample data for testing)
```
L1: Deacons                    M1: Households
L2: Andy A                     M2: Alan & Alexa Adams
L3: Brian B                    M3: Barbara Baker
L4: Chris C                    M4: Chloe Cooper
L5: Darell D                   M5: Delilah Danvers
... (continue with your data)
```
💡 **Setup Tip**: Use this sample data pattern during initial setup and testing. Replace with real member information only when ready for production use.

#### **N-O: Basic Contact Information** (Your input required)
```
N1: Phone Number              O1: Address
N2: (555) 123-1748            O2: 123 E. Kentucky St., Louisville, KY 40202
N3: (555) 123-8359            O3: 6301 Bass Trail, Louisville, KY 40218
N4: (555) 123-5723            O4: 1801 Lynn Way, Louisville, KY 40205
... (continue with contact info)
```

#### **P-S: Church Management Integration**

##### **Column P: Breeze Link Numbers** (Your input required)
```
P1: Breeze Link
P2: 29760588
P3: 41827365
P4: 15936847
... (8-digit numbers from Breeze CMS people profiles)
```

**How to find Breeze numbers:**
1. **Go to your Breeze CMS** → People
2. **Click on a household/person** profile
3. **Look at the URL**: `https://[yourchurch].breezechms.com/people/view/29760588`
4. **Copy the number** at the end (e.g., 29760588)
5. **Paste into column P** for that household

##### **Column Q: Notes Page Links** (Your input required)
```
Q1: Notes Pg Link
Q2: https://docs.google.com/document/d/1abc123def456ghi789/edit
Q3: https://docs.google.com/document/d/1xyz789abc123def456/edit
Q4: https://docs.google.com/document/d/1def456ghi789xyz123/edit
... (Full Google Doc URLs for visit notes)
```

**How to set up Notes Pages:**
1. **Create a Google Doc** for each household
2. **Title**: "[Household Name] - Visit Notes"
3. **Header section** with:
   ```
   Household: Alan & Alexa Adams
   Phone: (555) 123-1748
   Address: 123 E. Kentucky St., Louisville, KY 40202
   Visiting Pastor: [Pastor Name]
   ```
4. **Three-column table** with headers: Date | Deacon | Visit Notes
5. **Copy the sharing URL** and paste into column Q

##### **Columns R-S: Auto-Generated Shortened URLs** (System generated)
```
R1: Breeze Link (short)       S1: Notes Pg Link (short)
R2: http://tinyurl.com/abc123 S2: http://tinyurl.com/def456
R3: http://tinyurl.com/ghi789 S3: http://tinyurl.com/jkl012
... (Generated automatically by the system)
```

## 🔧 Step 3: Test the Native Modular System

1. **Refresh your spreadsheet** (F5 or Ctrl+R)
2. **You should see the enhanced menu**: "🔄 Deacon Rotation"
3. **If the menu doesn't appear**:
   - Go back to Apps Script
   - Click **Run** → **onOpen** from any module
   - **Authorize permissions** (external URL access for shortening)
   - Refresh the spreadsheet

## ✅ Step 4: Generate Schedule and Test Smart Calendar Features

### **4.1: Validate Enhanced Setup**
1. **Click "🔄 Deacon Rotation" → "🔧 Validate Setup"**
   - Reports Breeze links and Notes links counts
   - Validates new column structure
   - Checks all module integration

### **4.2: Generate Schedule**
1. **Click "🔄 Deacon Rotation" → "📅 Generate Schedule"**
   - Creates rotation in columns A-E
   - Generates deacon reports in columns G-I

### **4.3: Generate Shortened URLs**
1. **Click "🔄 Deacon Rotation" → "🔗 Generate Shortened URLs"**
   - Builds full Breeze URLs from numbers in column P
   - Shortens both Breeze and Notes URLs
   - Stores results in columns R and S
   - **Process time**: ~0.5 seconds per household (includes API delays)

### **4.4: Test Smart Calendar Functions** ⭐ **NEW**

#### **Initial Calendar Setup**
1. **Click "🔄 Deacon Rotation" → "📆 Calendar Functions" → "🚨 Full Calendar Regeneration"**
   - Creates initial calendar with all events
   - **Important**: This step is required before smart updates work

#### **Test Contact Info Updates**
1. **Change a phone number** in column N (e.g., Adams from (555) 123-1748 to (555) 999-0001)
2. **Click "📆 Calendar Functions" → "📞 Update Contact Info Only"**
3. **Verify**: Calendar event keeps original time but shows new phone number

#### **Test Future Events Updates**
1. **Manually modify current week's calendar event** (change time from 2 PM to 4 PM)
2. **Change an address** in column O  
3. **Click "📆 Calendar Functions" → "🔄 Update Future Events Only"**
4. **Verify**: Current week event unchanged, future events have new address

## 📅 Calendar Export Process & Timing

### **What Happens During Calendar Export**

The enhanced calendar export process includes several steps to ensure reliable operation:

1. **Configuration Loading**: Reads all household data, contact info, and links
2. **URL Preparation**: Uses shortened URLs if available, falls back to full URLs
3. **Calendar Setup**: Creates or accesses calendar (automatically named based on test mode)
4. **Optional Event Deletion**: If you choose to clear existing events (with delays)
5. **Event Creation**: Creates enhanced events with rich descriptions (with rate limiting)
6. **Progress Monitoring**: Logs progress and handles any individual failures

### **Runtime Expectations**

#### **Simple Heuristic:**
**"Expect approximately 30-45 seconds per 100 events"**

#### **Detailed Timing Formula:**
```
Total Runtime = (Number of Events × 0.35 seconds) + (Number of Events ÷ 25 seconds)
                    ↑                                ↑
            Google Calendar API processing    Rate limiting delays
```

#### **Real-World Examples:**
| Events | Expected Time | What You'll See |
|--------|---------------|-----------------|
| 25 | 10-15 seconds | Very quick |
| 50 | 20-30 seconds | Quick completion |
| 100 | 30-45 seconds | Standard processing |
| 150 | 45-60 seconds | Longer but reliable |
| 200 | 60-90 seconds | Extended processing |
| 300+ | 2-3 minutes | Large schedule handling |

#### **Factors That Affect Runtime:**
- **Internet connection speed** to Google's servers
- **Google Calendar API load** (varies throughout the day)
- **Event deletion** (if clearing existing events)
- **Time of day** (Google services faster during off-peak hours)
- **Update method**: Contact Info Only is fastest, Full Regeneration takes longest

### **Rate Limiting Protection**

The system includes intelligent delays to prevent "too many operations" errors:

#### **During Event Creation:**
- **Pause every 25 events**: 1-second delay
- **Progress logging**: "Created X events so far..."
- **Individual error isolation**: One failed event won't stop the process

#### **During Event Deletion (if clearing existing):**
- **Pause every 10 deletions**: 0.5-second delay
- **Cooldown period**: 2-second wait before creating new events
- **Batch processing**: Handles large numbers of existing events safely

#### **During Smart Updates:**
- **Contact Info Only**: Fastest option - only updates descriptions
- **Future Events Only**: Moderate speed - deletes and recreates future events only
- **Full Regeneration**: Slowest - complete rebuild with all safety delays

### **What You'll Experience**

#### **Normal Operation:**
- **Progress appears steady** with occasional brief pauses
- **Console logs show** event creation progress
- **No error messages** during successful operation
- **Final success dialog** appears with event count
- **Calendar events** appear with rich descriptions and working links

#### **If You See Long Delays:**
- **This is normal** for larger schedules (150+ events)
- **Don't interrupt the process** - let the rate limiting work
- **Watch for progress messages** in the console
- **Success dialog** will appear when complete

### **Troubleshooting Long Runtimes**

#### **If Export Takes Much Longer Than Expected:**

**"Script running over 5 minutes for 100 events"**
- **Check internet connection** - slow connection affects API calls
- **Try during off-peak hours** (early morning or late evening)
- **Reduce schedule size** to test with smaller batches first
- **Use Contact Info Only** instead of Full Regeneration when possible

**"Script seems frozen with no progress"**
- **Check browser console** for error messages
- **Don't refresh** - this will interrupt the process
- **Wait for timeout** - script will eventually show error or success

**"Getting 'too many operations' errors despite rate limiting"**
- **Wait 10-15 minutes** for Google's rate limit to reset
- **Try smaller batch** of events first
- **Use Contact Info Only** instead of clearing all events
- **Check if test mode** is properly detected to avoid conflicts

### **Best Practices for Large Schedules**

#### **For 200+ Events:**
1. **Schedule during off-peak hours** (early morning recommended)
2. **Test with smaller subset first** (4-6 weeks)
3. **Ensure stable internet connection**
4. **Don't use computer for other intensive tasks** during export
5. **Be patient** - the system is designed for reliability over speed
6. **Use smart updates** when possible instead of full regeneration

#### **For Very Large Schedules (400+ Events):**
- **Consider breaking into quarters** (export 13 weeks at a time)
- **Use incremental approach** rather than clearing all events
- **Plan for 5-10 minute processing time**
- **Monitor system resources** during processing
- **Prefer Contact Info Only updates** for contact changes

### **Performance Optimization Tips**

#### **To Reduce Runtime:**
- **Generate shortened URLs first** (avoids real-time URL shortening)
- **Use smart update options** instead of full regeneration:
  - **Contact Info Only** - fastest, preserves all scheduling
  - **Future Events Only** - moderate speed, protects current week
- **Export during Google's off-peak hours**
- **Ensure good internet connection**

#### **To Improve Reliability:**
- **Start with validation** to catch configuration errors early
- **Test with small schedule first** before full year export
- **Don't interrupt the process** once it starts
- **Let rate limiting delays work** - they prevent failures
- **Use test mode** for initial testing to avoid production calendar conflicts

## 🎯 Step 5: Understanding Smart Calendar Update Options

### **📞 Update Contact Info Only** ⭐ **SAFEST OPTION**
**When to use**: Contact information changes (phone, address, Breeze numbers, notes links)

**What it preserves**:
- ✅ Custom event times and dates
- ✅ Guest lists and invitations  
- ✅ Location details
- ✅ All scheduling customizations

**What it updates**:
- 🔄 Phone numbers
- 🔄 Addresses  
- 🔄 Breeze profile links
- 🔄 Notes page links
- 🔄 Calendar instructions

**Process**: Updates event descriptions only, preserving all other event details.
**Speed**: Fastest option - typically 30-60 seconds for 100+ events.

### **🔄 Update Future Events Only**
**When to use**: Contact changes during the week without affecting current week scheduling

**What it preserves**:
- ✅ This week's scheduling details and customizations
- ✅ Any current week modifications deacons have made

**What it updates**:
- 🔄 Events starting next week with current contact info
- 🔄 Future deacon assignments if schedule changed

**Process**: Deletes and recreates events starting next week only.
**Speed**: Moderate - depends on number of future events.

### **🚨 Full Calendar Regeneration**
**When to use**: Major structural changes (new deacons, household changes, schedule rebuild)

**What it preserves**:
- ❌ Nothing - complete rebuild

**What it updates**:
- 🔄 Everything with enhanced warnings

**Process**: Deletes ALL events and recreates from current spreadsheet data.
**Speed**: Slowest - full timing expectations apply.

### **Success Indicators**

#### **Your Export is Working Well When:**
✅ **Steady progress** with occasional planned pauses  
✅ **Console shows** "Created X events so far..." or "Updated X events so far..." messages  
✅ **No error alerts** during processing  
✅ **Final success dialog** appears with event count  
✅ **Calendar events** appear with rich descriptions and working links  
✅ **Test mode detection** shows correct environment

#### **Example Success Messages:**

**Contact Info Only:**
```
✅ Updated contact information in 135 calendar events!

📞 Refreshed: Phone numbers, addresses, Breeze links, Notes links
🔒 Preserved: All custom scheduling details, times, dates, guests

All events now have current contact information.
```

**Full Regeneration:**
```
✅ Created 135 calendar events with enhanced information!

📅 Calendar: "Deacon Visitation Schedule" (or "TEST - Deacon Visitation Schedule")
🕐 Default time: 2:00 PM - 3:00 PM
🔗 Includes: Breeze profiles and visit notes links
📞 Contact info: Phone numbers and addresses
📝 Instructions: Custom visit coordination text

Each event title: "[Deacon] visits [Household]"
View and modify these events in Google Calendar.
```

## 🧪 Step 6: Test with Sample Data Scenarios

### **Scenario 1: Mid-Week Contact Update**
```
Situation: It's Wednesday, Adams family got new phone number
Solution: Use "📞 Update Contact Info Only"
Result: All events keep custom scheduling, phone number updated
Speed: ~30 seconds for 100+ events
```

### **Scenario 2: Future Planning Changes**
```
Situation: Next month's assignments need updating, this week is planned
Solution: Use "🔄 Update Future Events Only"  
Result: Current week untouched, future events refreshed
Speed: ~60 seconds for 100+ future events
```

### **Scenario 3: Major Schedule Overhaul**
```
Situation: New deacon added, household list changed significantly
Solution: Use "🚨 Full Calendar Regeneration" (with warnings)
Result: Complete rebuild with all current data
Speed: ~2-3 minutes for 200+ events
```

## 🚨 Troubleshooting Smart Calendar Features

### **Common Issues:**

**"Menu doesn't show Calendar Functions submenu"**
- Verify all 4 modules copied correctly to Apps Script
- Re-authorize script permissions 
- Run onOpen function manually from Apps Script
- Refresh spreadsheet after module setup

**"Update Contact Info Only doesn't find events"**
- Ensure calendar exists (run Full Calendar Regeneration first)
- Check event titles match "[Deacon] visits [Household]" format
- Verify household names in spreadsheet match calendar events exactly
- Check if test mode is properly detected

**"Future Events Only deletes too many events"**
- Function preserves current week (7 days from today)
- Check dates to ensure proper week boundary
- Use Contact Info Only for safer updates

**"Smart updates missing information"**
- Run "Generate Shortened URLs" before calendar updates
- Verify contact information in columns N-O
- Check Breeze numbers and Notes links in columns P-Q

**"Calendar export fails with 'too many operations'"**
- **Wait 5-10 minutes** for Google's rate limit to reset
- **Use smaller schedule** for initial testing
- **Don't clear existing events** unless necessary
- **Try Contact Info Only** instead of full regeneration
- System includes automatic rate limiting - let it work

**"Calendar events appear in wrong calendar (test vs production)"**
- Check mode indicator in cells K11-K12
- Use "Show Current Mode" menu option to verify detection
- Ensure test data patterns are properly recognized
- Mode detection happens automatically based on spreadsheet data

### **Enhanced System Tests:**
Run **"🔄 Deacon Rotation" → "🧪 Run Tests"** to verify:
- All module integration working correctly
- URL shortening functionality
- Breeze URL construction
- Calendar access permissions
- Script permissions for external services
- Test mode detection accuracy

## 🎛️ Step 7: Master the Enhanced Menu System

### **Menu Structure (v24.2):**
```
🔄 Deacon Rotation
├── 📅 Generate Schedule
├── 🔗 Generate Shortened URLs
├── 📆 Calendar Functions                    ⭐ NEW SUBMENU
│   ├── 📞 Update Contact Info Only         (Safest updates)
│   ├── 🔄 Update Future Events Only        (Current week safe)
│   └── 🚨 Full Calendar Regeneration       (Complete rebuild)
├── 📊 Export Individual Schedules
├── 📁 Archive Current Schedule
├── 🗓️ Generate Next Year
├── 🔧 Validate Setup
├── 🧪 Run Tests
├── [🧪/✅] Show Current Mode               ⭐ NEW
└── ❓ Setup Instructions
```

### **Recommended Usage Flow:**
1. **Initial Setup**: Full Calendar Regeneration
2. **Regular Updates**: Update Contact Info Only  
3. **Weekly Planning**: Update Future Events Only (if needed)
4. **Major Changes**: Full Calendar Regeneration (with caution)
5. **Mode Checking**: Use "Show Current Mode" anytime

## 🔄 Step 8: Development and Maintenance Workflow

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

### **Regular Maintenance Tasks:**
- **Weekly**: Test smart calendar updates with sample data
- **Monthly**: Review shortened URLs in columns R-S for any failures
- **Quarterly**: Run system tests to verify all integrations
- **Yearly**: Archive previous schedules and generate next year

---

**Remember: The system prioritizes reliability over speed. The built-in delays and smart update options ensure your calendar operations complete successfully while preserving the scheduling customizations that deacons depend on for effective pastoral care!** 🎯

## 🔐 Pre-Production Security Checklist

Before transitioning from sample data to real member information:

✅ **All smart calendar features tested** with sample data  
✅ **Google Sheets permissions** properly configured  
✅ **Only authorized deacons** have edit access  
✅ **Breeze CMS access** limited to appropriate users  
✅ **Visit notes documents** have proper sharing settings  
✅ **No sample data** mixed with real member information  
✅ **Backup procedures** established for member data  
✅ **Smart update options** understood by all users
✅ **Test mode detection** verified with sample data patterns

## 🎉 Success Checklist

✅ **Native modular architecture** deployed with 4 separate .gs files  
✅ **Spreadsheet configured** with P-S columns for integration  
✅ **Breeze numbers added** for household profile access  
✅ **Google Docs created** for visit documentation  
✅ **Shortened URLs generated** for clean presentation  
✅ **Smart calendar functions tested** with sample data  
✅ **Contact info only updates** verified to preserve scheduling  
✅ **Future events only updates** confirmed to protect current week  
✅ **Test mode detection** working automatically
✅ **Deacon training completed** on new update options  
✅ **Mobile access verified** for field use  

**Ready for comprehensive pastoral care coordination with smart calendar management!** 🎯
