# Deacon Visitation Rotation System

A comprehensive Google Apps Script solution for automatically scheduling and managing church deacon household visitations with intelligent rotation algorithms, **Breeze Church Management System integration**, and **Google Docs visit notes connectivity**. Features a **modular architecture** for enhanced maintainability and an **optimized column layout** for improved user experience.

## 🎯 Key Features

- **Intelligent Rotation Algorithm**: Eliminates harmonic locks that cause uneven distribution
- **Breeze CMS Integration**: Direct links to household profiles with automatic URL shortening
- **Google Docs Integration**: Seamless access to visit notes and documentation
- **Enhanced Calendar Events**: Rich descriptions with clickable Breeze profiles and notes pages
- **Modular Architecture**: Three maintainable modules for better code organization
- **Optimized Column Layout**: Streamlined interface with logical data grouping
- **Flexible Visit Frequency**: Support for weekly, bi-weekly, monthly, or custom intervals  
- **Contact Information Management**: Built-in phone numbers and addresses for each household
- **Individual Deacon Reports**: Personalized schedules positioned adjacent to main schedule
- **Robust Error Handling**: Comprehensive validation and Google API rate limiting protection
- **Export Capabilities**: Individual spreadsheets and enhanced calendar integration
- **Archive Function**: Historical schedule preservation
- **Anti-Harmonic Mathematics**: Prevents deacons from being locked to the same households

## 📊 Perfect for Churches That Need

- **Fair workload distribution** among deacons
- **Variety in household assignments** (no one stuck visiting the same family forever)
- **Flexible scheduling** (weekly to monthly visit frequencies)
- **Professional coordination** with contact information and management system integration
- **Quick access to Breeze profiles** during pastoral visits
- **Centralized visit documentation** with Google Docs integration
- **Mathematical optimization** that actually works regardless of deacon/household ratios
- **Easy maintenance** with modular code architecture
- **Streamlined data entry** with logical column organization

## 🔗 Church Management Integration

### **Breeze Church Management System**
- **8-digit profile integration**: Direct links to household profiles in Breeze CMS
- **Automatic URL construction**: `https://immanuelky.breezechms.com/people/view/[number]`
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

## 🏗️ Modular Architecture

### **Module 1: Core Functions & Configuration** (~500 lines)
- Main entry point and schedule generation
- Configuration loading and validation (including Breeze/Notes data)
- Header setup and UI initialization
- Enhanced validation with link reporting

### **Module 2: Algorithm & Schedule Generation** (~470 lines) 
- Core optimal rotation pattern generator
- Harmonic resonance detection and mitigation
- Schedule writing and formatting
- Individual deacon report generation
- Quality analysis and logging

### **Module 3: Export, Menu & Utility Functions** (~530 lines)
- **NEW**: URL shortening functionality with TinyURL integration
- **NEW**: Breeze URL construction and management
- **Enhanced**: Calendar export with rich descriptions and rate limiting
- Individual schedule exports
- Archive and year generation functions
- Menu system and event handlers
- Testing and diagnostic utilities

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

### **P-S**: NEW - Church Management Integration
- **P**: Breeze Link (8-digit numbers from Breeze CMS)
- **Q**: Notes Pg Link (Google Doc URLs)
- **R**: Breeze Link (short) - Auto-generated shortened URLs
- **S**: Notes Pg Link (short) - Auto-generated shortened URLs

## 🚀 Quick Start

### **Enhanced Setup Process**
1. **Basic Setup**: Follow standard installation for deacons, households, and contact info
2. **Breeze Integration**: Add 8-digit Breeze numbers in column P
3. **Notes Setup**: Add Google Doc URLs for visit notes in column Q
4. **Generate Short URLs**: Use "🔗 Generate Shortened URLs" menu option
5. **Export to Calendar**: Enhanced events with clickable Breeze and notes links

### **Installation Options**

#### **Option A: Simple Deployment (Most Users)**
1. **Copy the combined script** from `deacon-rotation-combined.js`
2. **Create Google Spreadsheet** with enhanced column structure (P-S for integration)
3. **Add church data** including Breeze numbers and Google Doc links
4. **Generate shortened URLs** and export to calendar

#### **Option B: Modular Development (Advanced Users)**
1. **Work with individual modules** for easier maintenance
2. **Modify integration features** in Module 1 and Module 3
3. **Test and combine** for deployment
4. **Benefits**: Better organization, easier debugging, team collaboration

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

## 🔧 New Functionality in v24.1

### **URL Shortening**
- **TinyURL integration**: Free API without account requirements
- **Automatic processing**: Batch generation with rate limiting protection
- **Storage system**: Shortened URLs saved in columns R and S for reuse
- **Fallback handling**: Uses full URLs if shortening fails

### **Enhanced Menu System**
- **🔗 Generate Shortened URLs**: New dedicated function for link processing
- **📆 Export to Google Calendar**: Enhanced with rate limiting and progress tracking
- **🔧 Validate Setup**: Now reports Breeze and Notes link counts

### **Rate Limiting Protection**
- **Google Calendar API**: Intelligent delays to prevent "too many operations" errors
- **Batch processing**: Pauses every 10 deletions, every 25 creations
- **Progress tracking**: Real-time logging for large operations
- **Error isolation**: Individual failures won't stop entire process

## 📋 System Requirements

- Google Workspace account (free Gmail accounts work)
- Google Sheets access
- Google Apps Script permissions
- Google Calendar (optional, for calendar export feature)
- **NEW**: Internet access for URL shortening service
- **NEW**: Breeze Church Management System (for profile integration)
- **Processing time**: Approximately 30-45 seconds per 100 calendar events
- **Large schedules**: Plan for 2-3 minutes for 300+ events

## 📖 Documentation

- **[Setup Guide](SETUP.md)** - Complete installation including Breeze and Notes integration
- **[Features Overview](FEATURES.md)** - Detailed feature explanations  
- **[Changelog](CHANGELOG.md)** - Version history including v24.1 enhancements
- **[Examples](examples/)** - Sample configurations and use cases

## 🎯 The Math Behind It

This system solves a complex mathematical problem in rotation scheduling called "harmonic resonance." When the ratio of deacons to households creates mathematical harmonics (like 12:6 or 14:7), simple rotation algorithms cause deacons to visit only the same household(s) repeatedly.

Our **optimal pattern generator** uses intelligent scoring to:
- Prioritize deacons with fewer total visits
- Favor new deacon-household pairings  
- Prevent same-cycle double assignments
- Maximize variety while maintaining fairness

## 🔄 Development Workflow

### **Working with Enhanced Modules**
```bash
# Development in separate modules
src/
├── module1-core-config.js      (Enhanced with Breeze/Notes)
├── module2-algorithm.js        (Unchanged)
└── module3-export-utils.js     (Enhanced with URL shortening)

# Deployment (concatenated)
deacon-rotation-combined.js
```

### **Benefits of Modular + Integration Approach**
- ✅ **Easier maintenance** - Church management features isolated in specific modules
- ✅ **Better debugging** - URL shortening and API issues contained
- ✅ **Team collaboration** - Multiple developers can work on integration features
- ✅ **External API management** - Rate limiting and error handling centralized
- ✅ **Future expansion** - Ready for additional church management integrations

## 📞 Support

This project was developed for Immanuel Baptist Church but is shared freely for other churches to benefit from. 

For technical issues:
- Check the [Setup Guide](SETUP.md) for common solutions including Breeze setup
- Review the [Features Documentation](FEATURES.md) for usage questions
- Open an issue on GitHub for bugs or enhancement requests

## 🤝 Contributing

Contributions welcome! Whether you're fixing bugs, improving documentation, adding church management integrations, or enhancing features:

1. Fork the repository
2. Create a feature branch
3. Work on individual modules for easier review
4. **Test integration features with sample data**
5. **Use fake member information** when creating examples or reporting issues
6. Submit a pull request

### **Contributing to Integration Features**
- **Module 1**: Configuration and validation for new church management systems
- **Module 3**: Export enhancements, API integrations, URL shortening improvements
- **Documentation**: Setup guides for different church management systems

## 📜 License

MIT License - See [LICENSE](LICENSE) for details.

This project is free for any church or organization to use, modify, and distribute.

## 🏛️ About

Developed for church pastoral care coordination. Solving real scheduling problems with mathematical precision while providing seamless integration with church management systems and maintaining the human touch that makes deacon ministry meaningful.

### **Version 24.1 Highlights**
- 🔗 **Breeze CMS integration** with automatic URL shortening
- 📝 **Google Docs visit notes** with enhanced calendar events
- 🛡️ **Rate limiting protection** for reliable Google Calendar operations
- 🎯 **Enhanced user experience** with clickable links and rich information
- 🔧 **Improved reliability** with intelligent API handling

---

**Built with ❤️ for church ministry, now with comprehensive church management system integration**
