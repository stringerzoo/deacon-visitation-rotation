# Deacon Visitation Rotation System v2.0

> **Automated scheduling with household-specific visit frequencies, smart calendar integration, and Google Chat notifications**

[![Version](https://img.shields.io/badge/version-2.0.0-blue.svg)](CHANGELOG.md)
[![Status](https://img.shields.io/badge/status-production-green.svg)]()

---

## 🎯 What's New in v2.0

### **🆕 Variable Frequency Scheduling**
The game-changing feature: **household-specific visit frequencies**. Now you can:
- Set **default frequency** for all households (K4)
- Override with **custom frequencies** for individual households (Column T)
- Support **1, 2, 3, or 4 week** intervals per household
- **Automatic workload balancing** across all deacons regardless of frequency mix

**Example Use Cases:**
- Weekly visits for households with urgent needs
- Monthly visits for households preferring less frequent contact  
- Standard bi-weekly visits for most households
- Flexible scheduling as household needs change

---

## ✨ Core Features

### 📅 **Intelligent Scheduling**
- **Variable frequency support (v2.0)**: Different households, different schedules
- **Harmonic resonance elimination**: Guarantees fair distribution even with challenging deacon/household ratios
- **Enhanced anti-repetition logic**: Prevents same deacon visiting same household too frequently
- **Automatic quality validation**: Built-in checks for scheduling issues
- **Individual & household reports**: Everyone sees their complete schedule

### 📊 **Comprehensive Reporting**
- **Main schedule (A-E)**: Complete rotation with visual frequency indicators
- **Deacon reports (G-I)**: Individual schedules for each deacon
- **Household reports (G-I)**: NEW in v2.0 - Complete visit schedule by household 🆕
- **Schedule Summary Generator**: NEW - Creates shareable spreadsheet with QR code 🆕
- **Frequency markers**: Clear indicators for custom frequency households
- **Archive support**: Historical record keeping

### 📢 **Google Chat Notifications**
- **Weekly automation**: Configurable day/time for automatic summaries
- **2-week lookahead**: Current + next calendar week visibility
- **Rich formatting**: Contact info, Breeze links, visit notes
- **Resource links**: Calendar, guide, and schedule summary access
- **Test mode separation**: Safe testing environment

### 📱 **Mobile-Friendly Integration**
- **Shortened URLs**: TinyURL integration for mobile compatibility
- **Breeze CMS**: Direct profile links from 8-digit numbers
- **Google Docs**: Visit notes integration
- **Smart calendar events**: Arrow format (Deacon ➡︎ Household) with visit notes at top
- **QR codes**: Auto-generated for schedule summary sharing 🆕

### 🛠️ **Smart Management**
- **Three-tier calendar updates**: Contact only, future only, or full regeneration
- **Automatic test/production detection**: Zero-configuration mode switching
- **Comprehensive diagnostics**: Built-in troubleshooting tools
- **Enhanced error handling**: Year rollover fixes, date validation

---

## 🚀 Quick Start

1. **[Complete Setup Guide](SETUP.md)** - Follow comprehensive installation instructions
2. **Configure Column T** - Set custom frequencies for households needing special scheduling
3. **Generate Schedule** - Use "📅 Generate Schedule" menu option
4. **Review Reports** - Check deacon and household reports for accuracy
5. **Export to Calendar** - Use "🚨 Full Calendar Regeneration"
6. **Generate Summary** - Use "📊 Generate Schedule Summary Sheet" (creates shareable file + QR code)
7. **Configure Notifications** - Use "📢 Notifications → 🔧 Configure Chat Webhook"
8. **Enable Automation** - Use "🔄 Enable Weekly Auto-Send"

**Read the [User Guide](USER_GUIDE.md)** for detailed operational instructions.

---

## 📋 System Requirements

- **Google Workspace** account (free Gmail works for basic features)
- **Google Apps Script** access
- **Google Chat** space (for notifications)
- **Breeze CMS** account (optional, for profile integration)

---

## 🗂️ Architecture Overview

### **Five-Module Design (v2.0)**
```
📦 Deacon Visitation Rotation System
├── Module 1: Core Configuration & Validation (v2.0 enhanced)
├── Module 2: Algorithm & Generation (v2.0 major rewrite)
├── Module 3: Smart Calendar & Mode Detection (v2.0 updated)
├── Module 4: Export, Menu & Utilities (v2.0 enhanced)
└── Module 5: Google Chat Notifications
```

### **Data Organization (v2.0)**
```
📊 Spreadsheet Layout
├── A-E: Generated schedule (Cycle, Week, Date, Household, Deacon)
├── G-I: Individual deacon reports
├── G-I: Household visit reports (below deacon reports) 🆕
├── K: Configuration settings (enhanced with default frequency)
├── L-M: Deacon and household lists
├── N-O: Contact information (phone, address)
├── P-Q: Integration links (Breeze, Notes)
├── R-S: Shortened URLs (auto-generated)
└── T: Custom visit frequency (1, 2, 3, or 4 weeks) 🆕
```

---

## 🔧 Configuration Highlights

### **Column K Settings (v2.0 Enhanced)**
- **K2**: Start date (Monday of first week)
- **K3**: "Visits every x weeks (1,2,3,4) **(default)**" 🆕
- **K4**: Default visit frequency (1-4 weeks)
- **K6**: Schedule length in weeks
- **K8**: Calendar event instructions
- **K11**: Notification day (dropdown)
- **K13**: Notification time (0-23 hour)
- **K19**: Google Calendar URL (auto-detected)
- **K22**: Visitation Guide URL
- **K25**: Schedule Summary URL (auto-updated by summary generator) 🆕

### **Column T - Custom Frequencies 🆕**
- **Header**: "Custom visit frequency (every 1, 2, 3, or 4 weeks)"
- **Values**: 1, 2, 3, 4, or blank (uses default from K4)
- **Data validation**: Dropdown prevents invalid entries
- **Visual indicators**: Light yellow highlighting in main schedule
- **Backward compatible**: Empty Column T = v1.1 behavior

### **Smart Features**
- **Automatic algorithm selection**: Variable vs uniform frequency detection
- **Dual-mode compatibility**: v2.0 features with v1.1 fallback
- **Enhanced quality validation**: Automatic detection of scheduling issues
- **Household-centric reports**: Complete visibility for each household
- **QR code generation**: Automatic QR codes for schedule summary sharing

---

## 📊 v2.0 Algorithm Overview

### **Enhanced Scheduling Intelligence**
The v2.0 algorithm uses a sophisticated **7-factor scoring system** to assign deacons:

1. **Workload Balance** (×100): Distribute visits evenly across all deacons
2. **Recent Activity** (×10): Prefer deacons who haven't worked recently  
3. **Household Variety** (×500): Strong penalty against repeat pairings 🆕
4. **Recent Household Visits** (×300-1000): Prevent visiting same household too soon 🆕
5. **Consecutive Week Prevention** (×200): Avoid same deacon working back-to-back 🆕
6. **Frequency Preference** (×5): Slight bonus for high-frequency households
7. **High-Frequency Variety** (×100): Extra variety for weekly/bi-weekly households 🆕

**Result**: Fair distribution, excellent variety, no repetitive patterns, even with mixed frequencies.

### **Quality Validation**
Every generated schedule is automatically checked for:
- Same deacon visiting same household within 3 weeks
- Deacon assignments in consecutive weeks
- Excessive repetition of deacon-household pairings
- Workload imbalances

Issues are reported in the console logs with specific details.

---

## 🎨 Visual Design Updates

### **Calendar Events (v2.0)**
- **Title format**: `Deacon ➡︎ Household` (arrow instead of "visits")
- **Description order**: Visit Notes link moved to top for immediate visibility
- **Frequency information**: Shows "2 weeks (default)" or "3 weeks (custom)"
- **No location field**: Gives deacons flexibility to add custom locations
- **Test mode**: Clear "TEST:" prefix and warning message

### **Spreadsheet Reports**
- **Main schedule**: Subtle light yellow highlighting for custom frequency households
- **Legend**: Explains highlighting at bottom of schedule
- **Deacon reports**: Clean v1.1 format with frequency markers only for custom households
- **Household reports**: NEW - Complete visit schedule grouped by household 🆕
- **Date format**: Full year display (MM/DD/YYYY) for proper year rollover

### **Schedule Summary (v2.0)**
- **Standalone spreadsheet**: Shareable file separate from configuration
- **QR code**: Auto-generated PNG for easy mobile access
- **K25 integration**: URL auto-populated for notification messages
- **Print-friendly**: Optimized layout for distribution

---

## 🔄 Migration from v1.1

### **Backward Compatible**
v2.0 is **fully backward compatible** with v1.1 spreadsheets:
- Leave Column T empty → System operates exactly like v1.1
- No custom frequencies → Uniform frequency algorithm (proven v1.1 math)
- All existing features preserved → Zero breaking changes

### **Upgrade Path**
1. **Test environment first**: Copy production spreadsheet to test
2. **Add Column T**: System auto-creates with data validation
3. **Set custom frequencies**: Only for households needing special scheduling
4. **Generate test schedule**: Verify quality and distribution
5. **Deploy to production**: When confident in results

---

## 📖 Documentation

- **[Setup Guide](SETUP.md)** - Complete v2.0 installation with Column T configuration
- **[Changelog](CHANGELOG.md)** - Complete v2.0 release notes
- **[Features Overview](FEATURES.md)** - Technical deep-dive (v1.1, update pending)
- **[User Guide](USER_GUIDE.md)** - Operational instructions (update pending)

---

## 🎯 The Math Behind It

### **v1.1 Harmonic Resonance (Uniform Frequency)**
When all households visit on the same frequency, the system uses proven mathematical rotation patterns that eliminate "harmonic locks" where deacons would only visit the same households repeatedly.

### **v2.0 Dynamic Scoring (Variable Frequency)**
When households have different frequencies, traditional rotation patterns don't work. The v2.0 algorithm uses dynamic scoring to:
- Calculate individual household visit schedules
- Merge into a master timeline
- Assign deacons using multi-factor optimization
- Prevent micro-resonance loops in mixed frequency scenarios

**Both algorithms guarantee fair distribution and excellent variety.**

---

## 🆚 Version Comparison

| Feature | v1.1 | v2.0 |
|---------|------|------|
| **Uniform frequency scheduling** | ✅ | ✅ |
| **Variable frequency scheduling** | ❌ | ✅ 🆕 |
| **Household-specific frequencies** | ❌ | ✅ 🆕 |
| **Enhanced anti-repetition** | ⚠️ Basic | ✅ Advanced 🆕 |
| **Quality validation** | ❌ | ✅ Automatic 🆕 |
| **Household reports** | ❌ | ✅ Complete 🆕 |
| **Schedule summary generator** | ❌ | ✅ With QR code 🆕 |
| **Backward compatibility** | N/A | ✅ Full 🆕 |
| **Calendar arrow format** | ❌ | ✅ 🆕 |
| **Year rollover fix** | ⚠️ Bug | ✅ Fixed 🆕 |

---

## 🚀 Future Roadmap

### **v2.1** - Enhanced User Experience
- Complete USER_GUIDE.md update for v2.0 features
- Enhanced FEATURES.md with variable frequency deep-dive
- Additional documentation for lead deacon transitions

### **v2.2** - Advanced Features
- Breeze API integration for automated contact sync
- Historical reporting and trend analysis
- Enhanced calendar synchronization

### **v3.0** - Platform Evolution
- Web-based interface option
- Advanced notification channels (SMS, email)
- Multi-church deployment support

---

## 💡 Key Innovations in v2.0

1. **Household-centric flexibility**: First version to support per-household scheduling
2. **Enhanced algorithm**: Prevents all forms of repetitive patterns
3. **Automatic quality validation**: Built-in issue detection
4. **Complete reporting**: Both deacon and household perspectives
5. **Seamless backward compatibility**: Zero risk upgrade path
6. **QR code integration**: Modern sharing and mobile access

---

## 📞 Support & Resources

- **GitHub Repository**: [Your repository link]
- **Documentation**: Complete guides in repository
- **Lead Deacon Handoff Guide**: Complete system transfer documentation

---

*Deacon Visitation Rotation System v2.0 - Transforming church pastoral care through intelligent automation*
