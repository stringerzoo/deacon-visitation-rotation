# Deacon Visitation Rotation System

A comprehensive Google Apps Script solution for automatically scheduling and managing church deacon household visitations with intelligent rotation algorithms, **Breeze Church Management System integration**, **Google Docs visit notes connectivity**, and **smart calendar update capabilities**. Features a **native modular architecture** for enhanced maintainability and real-world pastoral care coordination.

## 🎯 Key Features

- **Intelligent Rotation Algorithm**: Eliminates harmonic locks that cause uneven distribution
- **Smart Calendar Updates**: Preserve custom scheduling while updating contact information
- **Intelligent Test Mode Detection**: Automatic switching between test and production environments
- **Breeze CMS Integration**: Direct links to household profiles with automatic URL shortening
- **Google Docs Integration**: Seamless access to visit notes and documentation
- **Enhanced Calendar Events**: Rich descriptions with clickable Breeze profiles and notes pages
- **Native Modular Architecture**: Four maintainable Google Apps Script files
- **Flexible Visit Frequency**: Support for weekly, bi-weekly, monthly, or custom intervals  
- **Contact Information Management**: Built-in phone numbers and addresses for each household
- **Individual Deacon Reports**: Personalized schedules positioned adjacent to main schedule
- **Robust Error Handling**: Comprehensive validation and Google API rate limiting protection

## 📊 Perfect for Churches That Need

- **Smart contact updates** without losing deacon scheduling customizations
- **Fair workload distribution** among deacons with mathematical precision
- **Variety in household assignments** (no one stuck visiting the same family forever)
- **Flexible scheduling** with preservation of custom times and dates
- **Professional coordination** with integrated church management system access
- **Quick access to Breeze profiles** during pastoral visits
- **Centralized visit documentation** with Google Docs integration
- **Real-world pastoral care** that adapts to scheduling changes mid-week

## 🔄 Smart Calendar Update System (NEW v24.2)

### **The Problem Solved:**
When contact information changes, traditional systems require complete calendar regeneration, **losing all custom deacon scheduling details** (modified times, dates, guests, locations).

### **Smart Update Options:**

#### **📞 Update Contact Info Only** ⭐ **SAFEST**
- **Preserves**: All custom scheduling (times, dates, guests, locations)
- **Updates**: Phone numbers, addresses, Breeze links, Notes links, instructions
- **Use case**: Mid-week contact changes without affecting current scheduling

#### **🔄 Update Future Events Only** 
- **Preserves**: This week's scheduling details and customizations
- **Updates**: Events starting next week with current contact info
- **Use case**: Contact updates without losing current week's deacon scheduling

#### **🚨 Full Calendar Regeneration**
- **Preserves**: Nothing - complete rebuild
- **Updates**: Everything with enhanced warnings
- **Use case**: Major structural changes requiring complete rebuild

## 🧪 Intelligent Test Mode Detection (NEW v24.2)

### **Automatic Environment Switching**
The system automatically detects whether you're working with test data or production data:

#### **Test Mode Triggers:**
- **Test household names** (Alan Adams, Barbara Baker, Chloe Cooper patterns)
- **Test phone numbers** (555 prefix)
- **Test Breeze numbers** (12345 prefix)
- **Spreadsheet title** contains "test" or "sample"

#### **Visual Indicators:**
- **Spreadsheet cells K11-K12** show current mode with color coding
- **Menu icon** displays 🧪 for test mode, ✅ for production
- **Calendar names** automatically include "TEST -" prefix when in test mode
- **Confirmation dialogs** clearly indicate current mode

#### **Benefits:**
- **No manual configuration** - works automatically based on your data
- **Safe testing** - prevents accidental production calendar modifications
- **Clear communication** - always know which mode you're in
- **Seamless transition** - automatically switches as you move from test to live data

## 🔗 Church Management Integration

### **Breeze Church Management System**
- **8-digit profile integration**: Direct links to household profiles in Breeze CMS
- **Automatic URL construction**: `https://yourchurch.breezechms.com/people/view/[number]`
- **URL shortening**: Clean, clickable links in calendar events
- **Fallback support**: Uses full URLs if shortening fails

### **Google Docs Visit Notes**
- **Dedicated documentation**: Individual Google Docs for each household
- **Structured format**: Contact info header with three-column visit log (Date, Deacon, Notes)
- **Direct access**: Shortened links in calendar events for field use
- **Centralized storage**: All visit history in one accessible location

### **Enhanced Calendar Events**
```
Household: Alan & Alexa Adams
Breeze Profile: http://tinyurl.com/abc123

Contact Information:
Phone: (555) 123-1748
Address: 123 E. Kentucky St., Louisville, KY 40202

Visit Notes: http://tinyurl.com/def456

Instructions:
Please call to confirm visit time...
```

## 🏗️ Native Modular Architecture (v24.2)

### **Google Apps Script File Structure:**
```
📁 Your Apps Script Project
├── Module1_Core_Config.gs          (~500 lines)
├── Module2_Algorithm.gs             (~470 lines)
├── Module3_Smart_Calendar.gs        (~200 lines)
└── Module4_Export_Menu.gs           (~400 lines)
```

### **Module Responsibilities:**

#### **Module 1: Core Functions & Configuration**
- Main entry point and schedule generation
- Configuration loading and validation (including Breeze/Notes data)
- Header setup and UI initialization
- Enhanced validation with link reporting

#### **Module 2: Algorithm & Schedule Generation**
- Core optimal rotation pattern generator
- Harmonic resonance detection and mitigation
- Schedule writing and formatting
- Individual deacon report generation
- Quality analysis and logging

#### **Module 3: Smart Calendar Updates** ⭐ **NEW**
- Contact info only updates with scheduling preservation
- Future events only updates with current week protection
- Smart event parsing and household matching
- Rate limiting protection for reliable operation

#### **Module 4: Export, Menu & Utility Functions**
- Full calendar regeneration with enhanced warnings
- Individual schedule exports
- URL shortening with TinyURL integration
- Archive and year generation functions
- Enhanced menu system and testing utilities

## 📋 Enhanced Column Layout

### **A-E**: Main Schedule
- Cycle, Week, Week of, Household, Deacon

### **F**: Buffer Column
- Visual separation for better readability

### **G-I**: Individual Deacon Reports
- Positioned adjacent to main schedule for efficient workflow
- Deacon, Week of, Household

### **J**: Buffer Column
- Visual separation between reports and configuration

### **K**: Configuration Settings
- Start Date (K2), Visit Frequency (K4), Schedule Length (K6), Calendar Instructions (K8)

### **L-M**: Core Data Lists
- L: Deacon names, M: Household names

### **N-O**: Basic Contact Information
- N: Phone numbers, O: Addresses

### **P-S**: Church Management Integration
- **P**: Breeze Link (8-digit numbers from Breeze CMS)
- **Q**: Notes Pg Link (Google Doc URLs)
- **R**: Breeze Link (short) - Auto-generated shortened URLs
- **S**: Notes Pg Link (short) - Auto-generated shortened URLs

## 🚀 Quick Start

### **Installation Process**
1. **Create Google Apps Script project**
2. **Set up 4 separate .gs files** with modular architecture
3. **Configure spreadsheet** with enhanced column structure (P-S for integration)
4. **Add church data** including Breeze numbers and Google Doc links
5. **Generate shortened URLs** and test smart calendar functions

### **Development Workflow**
```bash
# GitHub Repository Structure
📁 Your GitHub Repo
├── src/
│   ├── module1-core-config.js
│   ├── module2-algorithm.js
│   ├── module3-smart-calendar.js
│   └── module4-export-menu.js
├── docs/
│   ├── README.md
│   ├── SETUP.md
│   └── CHANGELOG.md
└── examples/
```

### **Deployment Process**
1. **Edit modules** in GitHub web interface
2. **Copy each file** directly to corresponding .gs file in Apps Script
3. **Test functionality** with new smart calendar features
4. **Commit changes** to GitHub when verified

## 🎛️ Enhanced Menu System (v24.2)

```
🔄 Deacon Rotation
├── 📅 Generate Schedule
├── 🔗 Generate Shortened URLs
├── 📆 Calendar Functions                    ⭐ NEW SUBMENU
│   ├── 📞 Update Contact Info Only         ⭐ NEW
│   ├── 🔄 Update Future Events Only        ⭐ NEW
│   └── 🚨 Full Calendar Regeneration
├── 📊 Export Individual Schedules
├── 📁 Archive Current Schedule
├── 🗓️ Generate Next Year
├── 🔧 Validate Setup
├── 🧪 Run Tests
├── [🧪/✅] Show Current Mode               ⭐ NEW
└── ❓ Setup Instructions
```

## 🔐 Data Security Best Practices

### **Development & Testing**
- **Always test with sample data** before adding real member information
- **Use obviously fake examples** (alphabetical patterns work well: Alan Adams, Barbara Baker, etc.)
- **Keep development spreadsheets separate** from production data

### **Production Deployment**
- **Restrict spreadsheet access** to authorized deacons only
- **Use Google Workspace permissions** to control data visibility
- **Regular security review** of who has access to church member data

### **Contributing with Sample Data**
When reporting issues or requesting features:
- **Replace real member data** with sample information
- **Use the established pattern**: Andy A, Barbara Baker, Chloe Cooper
- **Test with fake data** before sharing spreadsheets for troubleshooting

For detailed setup instructions, see [SETUP.md](SETUP.md).

## 🔧 New Functionality in v24.2

### **Smart Calendar Updates** ⭐ **MAJOR FEATURE**
- **Contact info only updates** that preserve ALL scheduling customizations
- **Future events only updates** that protect current week planning
- **Smart event parsing** with household matching and graceful error handling
- **Enhanced user guidance** with clear preservation vs. update messaging

### **Improved Architecture**
- **Native modular design** using separate Google Apps Script files
- **Streamlined development** - no more file combination required
- **Better organization** - easier debugging and maintenance
- **Cleaner git history** - changes show exactly which module was modified

### **Enhanced Menu Structure**
- **Calendar Functions submenu** with organized smart update options
- **Safety-first ordering** - safer options listed first in submenus
- **Enhanced warnings** for destructive operations with suggested alternatives

### **Real-World Use Cases Addressed**
- **Mid-week contact updates** without losing deacon scheduling details
- **Breeze profile changes** with preserved custom event timing
- **Notes document updates** while maintaining all scheduling customizations
- **Monthly planning cycles** with targeted updates

## 📋 System Requirements

- Google Workspace account (free Gmail accounts work)
- Google Sheets access
- Google Apps Script permissions
- Google Calendar (optional, for calendar export feature)
- **Internet access** for URL shortening service
- **Breeze Church Management System** (for profile integration)
- **Processing time**: Approximately 30-45 seconds per 100 calendar events
- **Large schedules**: Plan for 2-3 minutes for 300+ events

## 📖 Documentation

- **[Setup Guide](SETUP.md)** - Complete installation including Breeze and Notes integration
- **[Changelog](CHANGELOG.md)** - Version history including v24.2 enhancements
- **[Test Data](test/test_data_table.md)** - Safe sample data for testing

## 🎯 The Math Behind It

This system solves a complex mathematical problem in rotation scheduling called "harmonic resonance." When the ratio of deacons to households creates mathematical harmonics (like 12:6 or 14:7), simple rotation algorithms cause deacons to visit only the same household(s) repeatedly.

Our **optimal pattern generator** uses intelligent scoring to:
- Prioritize deacons with fewer total visits
- Favor new deacon-household pairings  
- Prevent same-cycle double assignments
- Maximize variety while maintaining fairness

## 🔄 Development Workflow Benefits

### **Native Modular Architecture Advantages:**
- ✅ **Direct editing** in Google Apps Script without file combination
- ✅ **Individual module testing** for faster debugging
- ✅ **Cleaner version control** with module-specific changes
- ✅ **Better collaboration** - different developers can work on different modules
- ✅ **Easier maintenance** - edit just the functionality you need

### **Smart Calendar Updates Benefits:**
- ✅ **Real-world compatibility** - preserves deacon scheduling customizations
- ✅ **Mid-week flexibility** - update contact info without disrupting current week
- ✅ **Risk reduction** - multiple update options from safest to complete rebuild
- ✅ **User guidance** - clear warnings and suggestions for appropriate update method

## 📞 Support

This project was developed for Immanuel Baptist Church but is shared freely for other churches to benefit from. 

For technical issues:
- Check the [Setup Guide](SETUP.md) for common solutions including smart calendar setup
- Open an issue on GitHub for bugs or enhancement requests

## 🤝 Contributing

Contributions welcome! Whether you're fixing bugs, improving documentation, adding church management integrations, or enhancing features:

1. Fork the repository
2. Create a feature branch
3. Work on individual modules for easier review
4. **Test smart calendar features with sample data**
5. **Use fake member information** when creating examples or reporting issues
6. Submit a pull request

### **Contributing to Specific Modules**
- **Module 1**: Configuration and validation for new church management systems
- **Module 2**: Algorithm improvements and optimization
- **Module 3**: Smart calendar update enhancements and new update methods
- **Module 4**: Export features, API integrations, menu improvements

## 📜 License

MIT License - See [LICENSE](LICENSE) for details.

This project is free for any church or organization to use, modify, and distribute.

## 🏛️ About

Developed for church pastoral care coordination. Solving real scheduling problems with mathematical precision while providing seamless integration with church management systems and smart calendar updates that preserve the human touch of deacon ministry scheduling.

### **Version 24.2 Highlights**
- 🔄 **Smart calendar updates** that preserve custom scheduling details
- 🏗️ **Native modular architecture** using separate Google Apps Script files
- 📞 **Contact info only updates** - the safest way to refresh information
- 🔄 **Future events only updates** - protect current week planning
- 🧪 **Intelligent test mode detection** - automatic environment switching
- 🎛️ **Enhanced menu system** with organized calendar functions
- 🛡️ **Real-world compatibility** for active pastoral care coordination

---

**Built with ❤️ for church ministry, now with smart calendar updates that preserve the real-world scheduling customizations deacons need for effective pastoral care.**
