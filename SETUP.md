# Setup Guide - Deacon Visitation Rotation System

This guide will walk you through setting up the deacon visitation rotation system for your church with the **enhanced Breeze Church Management System integration**, **Google Docs visit notes connectivity**, and **automatic URL shortening**.

## 📋 Prerequisites

Before starting, ensure you have:
- Google Workspace account (free Gmail works)
- Access to Google Sheets and Google Apps Script
- List of deacon names/initials
- List of household names to visit
- Contact information for households (phone numbers and addresses)
- **NEW**: Breeze Church Management System access (for profile integration)
- **NEW**: Google Docs for visit notes (optional but recommended)
- **NEW**: Internet connectivity (for URL shortening service)

## 🔧 Understanding the Enhanced Architecture

### **New Integration Features in v24.1**
The system now integrates with external church management tools to provide comprehensive pastoral care support:

- **Breeze CMS Integration**: Direct access to household profiles during visits
- **Google Docs Notes**: Centralized visit documentation with easy access
- **URL Shortening**: Clean, mobile-friendly links in calendar events
- **Enhanced Calendar Events**: Rich descriptions with clickable management links

### **Why These Enhancements?**
- **Field accessibility**: Deacons can quickly access household info during visits
- **Centralized documentation**: All visit notes stored in accessible Google Docs
- **Professional presentation**: Shortened URLs provide clean calendar events
- **Comprehensive information**: Everything needed for effective pastoral care in one place

## 🚀 Step 1: Choose Your Development Approach

### **Option A: Simple Deployment (Most Users)**
**Use the pre-combined file for quick setup with full integration**

1. **Go to Google Apps Script**: [script.google.com](https://script.google.com)
2. **Click "New Project"**
3. **Delete the default code** in the editor
4. **Copy the entire code** from `deacon-rotation-combined.js` (v24.1)
5. **Paste into the script editor**
6. **Save the project** (Ctrl+S or Cmd+S)
7. **Name your project**: "Deacon Visitation Rotation"
8. **Authorize permissions** (now includes external URL access for shortening)

### **Option B: Modular Development (Advanced Users)**
**Work with enhanced individual modules**

#### **1. Set Up Enhanced Modular Environment**
1. **Go to Google Apps Script**: [script.google.com](https://script.google.com)
2. **Click "New Project"**
3. **Create separate files** for each enhanced module:
   - `module1-core-config` (enhanced with Breeze/Notes configuration)
   - `module2-algorithm` (unchanged)
   - `module3-export-utils` (enhanced with URL shortening)

#### **2. Copy Enhanced Module Content**
- **Module 1**: Enhanced configuration loading with Breeze and Notes data
- **Module 2**: Core algorithm (unchanged from v24.0)
- **Module 3**: Enhanced with URL shortening and improved calendar export

#### **3. Create Enhanced Combined File**
```javascript
// Enhanced v24.1 with Breeze CMS and Google Docs integration
// Copy entire contents of enhanced module1-core-config.js
// Copy entire contents of module2-algorithm.js  
// Copy entire contents of enhanced module3-export-utils.js
```

## 📊 Step 2: Create Enhanced Google Spreadsheet Layout

### **Enhanced Column Structure (v24.1):**

#### **A-E: Main Rotation Schedule** (Auto-generated)
- **A**: Cycle, **B**: Week, **C**: Week of, **D**: Household, **E**: Deacon

#### **F: Buffer Column** (Leave empty)

#### **G-I: Individual Deacon Reports** (Auto-generated)
- **G**: Deacon, **H**: Week of, **I**: Household

#### **J: Buffer Column** (Leave empty)

#### **K: Configuration Settings** (Your input required)
```
K1: Start Date
K2: 6/9/2025 (or your preferred start date)
K3: Visits every x weeks (1,2,3,4)
K4: 2 (for bi-weekly visits)
K5: Length of schedule in weeks  
K6: 52 (for one year)
K7: Calendar Event Instructions:
K8: Please call to confirm visit time. Contact family 1-2 days before scheduled date to arrange convenient time.
```

#### **L-M: Core Data Lists** (Your input required)
```
L1: Deacons                    M1: Households
L2: Andy A                     M2: Alan & Alexa Adams
L3: Brian B                    M3: Barbara Baker
L4: Chris C                    M4: Chloe Cooper
L5: Darell D                   M5: Delilah Danvers
... (continue with your data)
```

#### **N-O: Basic Contact Information** (Your input required)
```
N1: Phone Number              O1: Address
N2: (555) 123-1748            O2: 123 E. Kentucky St., Louisville, KY 40202
N3: (555) 123-8359            O3: 6301 Bass Trail, Louisville, KY 40218
N4: (555) 123-5723            O4: 1801 Lynn Way, Louisville, KY 40205
... (continue with contact info)
```

#### **P-S: NEW - Church Management Integration**

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

## 🔧 Step 3: Run and Configure the Enhanced System

1. **Refresh your spreadsheet** (F5 or Ctrl+R)
2. **You should see a new menu**: "🔄 Deacon Rotation"
3. **If the menu doesn't appear**:
   - Go back to Apps Script
   - Click **Run** → **onOpen** 
   - **Authorize new permissions** (external URL access for shortening)
   - Refresh the spreadsheet

## ✅ Step 4: Generate Enhanced Schedule and Links

### **4.1: Validate Enhanced Setup**
1. **Click "🔄 Deacon Rotation" → "🔧 Validate Setup"**
   - Now reports Breeze links and Notes links counts
   - Validates new column structure
   - Checks Breeze number formats

### **4.2: Generate Schedule**
1. **Click "🔄 Deacon Rotation" → "📅 Generate Schedule"**
   - Creates rotation in columns A-E
   - Generates deacon reports in columns G-I

### **4.3: Generate Shortened URLs** (NEW!)
1. **Click "🔄 Deacon Rotation" → "🔗 Generate Shortened URLs"**
   - Builds full Breeze URLs from numbers in column P
   - Shortens both Breeze and Notes URLs
   - Stores results in columns R and S
   - **Process time**: ~0.5 seconds per household (includes API delays)

### **4.4: Export Enhanced Calendar**
1. **Click "🔄 Deacon Rotation" → "📆 Export to Google Calendar"**
   - **Option 1**: Generate shortened URLs first (recommended)
   - **Option 2**: Use full URLs if shortening skipped
   - Creates rich calendar events with clickable links
   - **Rate limiting**: Handles Google Calendar API limits automatically (you can read more about the export process and built-in rate limits [here](SETUP-CAL.md))

## 🧪 Step 5: Test the Enhanced System

### **Test Shortened URLs:**
- **Check columns R and S** for generated short URLs
- **Click test links** to verify they work
- **Breeze links** should open to correct profiles
- **Notes links** should open to Google Docs

### **Test Enhanced Calendar Events:**
- **Export a small schedule** (4-6 weeks) first
- **Check calendar event descriptions** for:
  - Household name and Breeze profile link
  - Contact information (phone and address)
  - Visit notes page link
  - Custom instructions
- **Verify links work** from calendar events

### **Test Rate Limiting:**
- **Large schedules**: System pauses every 25 events
- **Event deletion**: System pauses every 10 deletions
- **Progress tracking**: Watch console logs for updates

## 🚨 Troubleshooting Enhanced Features

### **Common Issues:**

**"Menu doesn't show URL shortening option"**
- Re-authorize script permissions (external URL access required)
- Run onOpen function manually from Apps Script
- Refresh spreadsheet after authorization

**"URL shortening failed"**
- Check internet connectivity
- Verify URLs in column Q are valid (start with https://)
- Try again later (TinyURL may have temporary limits)
- System falls back to full URLs if shortening fails

**"Calendar export fails with 'too many operations'"**
- **Wait 5-10 minutes** for Google's rate limit to reset
- **Use smaller schedule** for initial testing
- **Don't clear existing events** unless necessary
- System now includes automatic rate limiting

**"Breeze links don't work"**
- Verify 8-digit numbers in column P are correct
- Check Breeze CMS access and permissions
- Ensure church domain matches in buildBreezeUrl function

**"Enhanced calendar events missing information"**
- Run "Generate Shortened URLs" before calendar export
- Verify contact information in columns N-O
- Check Breeze numbers and Notes links in columns P-Q

### **Enhanced System Tests:**
Run **"🔄 Deacon Rotation" → "🧪 Run Tests"** to verify:
- Configuration loading (including new columns)
- URL shortening functionality
- Breeze URL construction
- Script permissions for external access

## 🔍 Step 6: Customize Enhanced Features

### **Breeze Integration Customization:**
- **Different church**: Update domain in Module 3's `buildBreezeUrl` function
- **Profile format**: Modify URL structure if your Breeze setup differs
- **Additional fields**: Add more Breeze data columns as needed

### **Notes Page Customization:**
- **Template standardization**: Create consistent Google Doc templates
- **Folder organization**: Store all household docs in shared church folder
- **Access permissions**: Ensure deacons have edit access to notes docs

### **URL Shortening Customization:**
- **Alternative services**: Replace TinyURL with is.gd or v.gd in Module 3
- **Custom domains**: Some URL shorteners support custom domains
- **Batch processing**: Adjust delays in `generateAndStoreShortUrls` function

## 🎯 Best Practices for Enhanced System

### **Breeze Integration Best Practices:**
1. **Test with sample households** before full deployment
2. **Verify Breeze permissions** for all deacons who will use links
3. **Keep Breeze numbers updated** when profiles change
4. **Document your Breeze URL format** for future reference
5. **Train deacons** on accessing Breeze profiles from calendar events

### **Google Docs Notes Best Practices:**
1. **Create consistent templates** for all household notes documents
2. **Use shared church folder** for organized storage
3. **Set proper sharing permissions** (deacons can edit, pastors can view)
4. **Include contact header** in each document for offline reference
5. **Regular backup** of important visit notes

### **URL Shortening Best Practices:**
1. **Generate URLs in batches** rather than individually
2. **Re-use existing short URLs** when possible (stored in columns R-S)
3. **Test shortened links** before major calendar exports
4. **Have backup plan** (system automatically falls back to full URLs)
5. **Monitor TinyURL service** status if issues arise

### **Calendar Export Best Practices:**
1. **Start with small schedules** (4-6 weeks) for testing
2. **Use "Add to existing"** rather than clearing all events frequently
3. **Let rate limiting work** - don't interrupt the process
4. **Test calendar events** on mobile devices for field use
5. **Train deacons** on accessing links from calendar apps

## 🆕 What's New in v24.1

### **Enhanced Integration Features**
- **Breeze CMS connectivity**: Direct household profile access
- **Google Docs integration**: Centralized visit documentation
- **Automatic URL shortening**: Clean, mobile-friendly links
- **Rich calendar events**: Comprehensive information in structured format

### **Improved Reliability**
- **Rate limiting protection**: Handles Google Calendar API limits gracefully
- **Batch processing**: Efficient handling of large schedules
- **Fallback systems**: Works even if external services are unavailable
- **Enhanced error handling**: Better feedback and recovery options

### **User Experience Improvements**
- **Adjacent information**: All related data logically grouped
- **One-click access**: Direct links to management systems
- **Mobile-friendly**: Shortened URLs work well on phones
- **Professional presentation**: Clean, organized calendar events

## 🔄 Step 7: Regular Maintenance

### **Weekly Tasks:**
- **Review shortened URLs** in columns R-S for any failures
- **Update Breeze numbers** for new households
- **Create Google Docs** for new household notes pages
- **Test calendar links** occasionally to ensure connectivity

### **Monthly Tasks:**
- **Audit Breeze integration** for accuracy
- **Review visit notes** documentation consistency
- **Update deacon access permissions** as needed
- **Export individual schedules** with enhanced information

### **Yearly Tasks:**
- **Archive previous year's** notes and schedules
- **Update Breeze profile links** for any household changes
- **Review URL shortening** performance and alternatives
- **Train new deacons** on enhanced system features

## 🤝 Getting Help with Enhanced Features

### **Breeze Integration Support:**
1. **Breeze CMS documentation** for profile URL formats
2. **Church IT administrator** for access and permissions
3. **Breeze support** for technical issues with their system

### **Google Docs Issues:**
1. **Google Workspace admin** for sharing and permissions
2. **Document template** questions - check church standards
3. **Access problems** - verify deacon Google accounts

### **URL Shortening Problems:**
1. **TinyURL status page** for service availability
2. **Alternative services** (is.gd, v.gd) if TinyURL fails
3. **Network connectivity** issues with your internet provider

### **General System Support:**
1. **Run enhanced system tests** for comprehensive diagnostics
2. **Check this setup guide** for troubleshooting steps
3. **Review the [Enhanced Features Documentation](FEATURES.md)**
4. **Open an issue** on the GitHub repository

## 🔄 Migration from v24.0

If you're upgrading from v24.0 to v24.1:

### **1. Update Your Codebase**
- **Replace Module 1** with enhanced version (Breeze/Notes support)
- **Keep Module 2** unchanged (core algorithm)
- **Replace Module 3** with enhanced version (URL shortening)
- **Create new combined file** with all enhancements

### **2. Enhance Your Spreadsheet**
- **Add column headers** P, Q, R, S if not present
- **Populate Breeze numbers** in column P
- **Add Google Doc URLs** in column Q
- **Leave columns R and S** empty (auto-generated)

### **3. Test Enhanced Features**
- **Validate setup** to check new columns
- **Generate shortened URLs** for existing data
- **Export test calendar** to verify enhanced events
- **Train deacons** on new link features

### **4. Deploy Gradually**
- **Start with subset** of households for testing
- **Verify Breeze integration** works correctly
- **Test Notes access** from calendar events
- **Roll out to full schedule** once confident

## 📱 Mobile Usage Tips

### **For Deacons in the Field:**
- **Calendar apps**: Links work directly from mobile calendar events
- **Breeze mobile**: Access household profiles on phones/tablets
- **Google Docs mobile**: Edit visit notes from mobile devices
- **Offline access**: Take screenshots of key info before visits

### **Recommended Mobile Workflow:**
1. **Check calendar event** for household information
2. **Click Breeze link** to review household profile
3. **Conduct pastoral visit** with full context
4. **Click Notes link** to document visit immediately
5. **Update contact info** if needed during visit

---

**Ready to transform your church's deacon visitation coordination with comprehensive church management integration!** 🎯

## 🎉 Success Checklist

✅ **Enhanced codebase deployed** with v24.1 features  
✅ **Spreadsheet configured** with P-S columns for integration  
✅ **Breeze numbers added** for household profile access  
✅ **Google Docs created** for visit documentation  
✅ **Shortened URLs generated** for clean presentation  
✅ **Calendar export tested** with enhanced event descriptions  
✅ **Deacon training completed** on new link features  
✅ **Mobile access verified** for field use  

**Your enhanced deacon visitation system is now ready for comprehensive pastoral care coordination!** 🚀
