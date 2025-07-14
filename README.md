# Deacon Visitation Rotation System v1.1

[![Version](https://img.shields.io/badge/version-1.1-blue.svg)](CHANGELOG.md)
[![Google Apps Script](https://img.shields.io/badge/platform-Google%20Apps%20Script-green.svg)](https://script.google.com)
[![License](https://img.shields.io/badge/license-MIT-orange.svg)](LICENSE)

> **Automated, fair rotation scheduling for church deacon household visitations with Google Chat notifications**

A sophisticated Google Apps Script-based system that generates mathematically fair visitation schedules, exports to Google Calendar, and sends automated notifications via Google Chat. Designed for churches seeking to organize deacon ministry efficiently and transparently.

---

## 🎯 What This System Does

**Creates Fair Schedules**: Uses advanced algorithms to ensure every deacon visits every household over time, with optimal spacing between visits.

**Google Calendar Integration**: Exports complete schedules with contact information, Breeze profile links, and visit notes access.

**Automated Notifications**: Sends weekly 2-week lookahead summaries to your deacon Google Chat space.

**Smart Updates**: Safely update contact information or future events without losing custom scheduling.

**Test & Production Modes**: Automatically detects environment and uses appropriate calendars and chat spaces.

---

## ✨ Key Features (v1.1)

### 📅 **Advanced Scheduling**
- **Harmonic resonance elimination**: Guarantees fair distribution even with challenging deacon/household ratios
- **Flexible timing**: Configure visit frequency (every 2-4 weeks typical)
- **Individual reports**: Each deacon gets their personal schedule
- **Archive support**: Historical record keeping

### 🔔 **Google Chat Notifications (v1.0)**
- **Weekly automation**: Configurable day/time for automatic summaries
- **Rich formatting**: Contact info, Breeze links, visit notes
- **Resource links**: Calendar, guide, and summary access
- **Test mode separation**: Safe testing environment

### 📱 **Mobile-Friendly Integration**
- **Shortened URLs**: TinyURL integration for mobile compatibility
- **Breeze CMS**: Direct profile links from 8-digit numbers
- **Google Docs**: Visit notes integration
- **Calendar access**: Mobile-optimized calendar events

### 🛠️ **Smart Management (v1.1)**
- **Three-tier updates**: Contact only, future only, or full regeneration
- **Automatic detection**: Test vs production mode
- **Comprehensive diagnostics**: Built-in troubleshooting tools
- **Menu cleanup**: Streamlined interface, no dead functions

### 📅 **Automatic Calendar Links**
- **Auto-detected calendar URLs** - System automatically generates calendar links for notifications
- **K22 Guide field** - Optional Visitation Guide URL for ministry procedures  
- **K25 Summary field** - Optional Schedule Summary URL for archived schedules
- **Zero configuration** - Calendar links work immediately without manual setup
- **Mode-aware** - Automatically switches between test and production calendar URLs
- **One-click access** - "📅 Calendar", "📋 Guide", "📊 Summary" in every chat message

---

## 🚀 Quick Start

1. **[Complete Setup](docs/SETUP.md)** - Follow the comprehensive installation guide
2. **Generate Schedule** - Use "📅 Generate Schedule" menu option
3. **Export to Calendar** - Use "🚨 Full Calendar Regeneration"
4. **Configure Notifications** - Use "📢 Notifications → 🔧 Configure Chat Webhook"
5. **Test System** - Use "📋 Test Notification System"
6. **Enable Automation** - Use "🔄 Enable Weekly Auto-Send"

**Read the [User Guide](docs/USER_GUIDE.md)** for detailed operational instructions.

---

## 📋 System Requirements

- **Google Workspace** account (free Gmail works for basic features)
- **Google Apps Script** access
- **Google Chat** space (for notifications)
- **Breeze CMS** account (optional, for profile integration)

---

## 🏗️ Architecture Overview

### **Five-Module Design (v1.0)**
```
📁 Deacon Visitation Rotation System
├── Module 1: Core Configuration & Validation
├── Module 2: Mathematical Algorithm & Generation  
├── Module 3: Smart Calendar & Mode Detection
├── Module 4: Export, Menu & Utilities
└── Module 5: Google Chat Notifications ⭐
```

### **Data Organization**
```
📊 Spreadsheet Layout
├── A-E: Generated schedule (Date, Deacon, Household)
├── G-I: Individual deacon reports  
├── K: Configuration settings
├── L-M: Deacon and household lists
├── N-O: Contact information (phone, address)
├── P-Q: Integration links (Breeze, Notes)
└── R-S: Shortened URLs (auto-generated)
```

---

## 🔧 Configuration Highlights

### **Column K Settings**
- **K2**: Start date (Monday of first week)
- **K4**: Visit frequency in weeks  
- **K6**: Schedule length in weeks
- **K8**: Calendar event instructions
- **K11**: Notification day (dropdown)
- **K13**: Notification time (0-23 hour)
- **K22**: Visitation Guide URL
- **K25**: Schedule Summary URL

### **Smart Features**
- **Automatic test mode**: Detects sample vs. real data
- **Dropdown validation**: Prevents configuration errors
- **Visual indicators**: Clear mode and status display
- **Resource links**: Configurable calendar and guide access

---

## 📢 Notification System (v1.0)

### **Weekly Chat Summaries**
Automated messages include:
- **2-week lookahead**: Current + next calendar week
- **Contact information**: Phone numbers and addresses  
- **Direct links**: Breeze profiles and visit notes
- **Resource access**: Calendar, guide, and summary links
- **Mobile optimization**: Short URLs for field use

### **Sample Notification**
```
📅 Weekly Visitation Update

Week 1 (Jul 14-20, 2025)
👤 Andy B visits Stephen & Barbara OBryan
📞 (502) 415-1748 | 📍 620 E. Kentucky St.
🔗 Breeze Profile | 📝 Visit Notes

Week 2 (Jul 21-27, 2025)  
👤 Cody G visits CoCo Reparee
📞 (502) 618-4963 | 📍 1547 Trevilian Way
🔗 Breeze Profile | 📝 Visit Notes

📅 View Visitation Calendar
📋 Visitation Guide  
📊 Schedule Summary
```

---

## 🔍 Troubleshooting

### **Common Issues**
- **"No deacons found"**: Add names in column L starting row 2
- **"Notifications not sending"**: Check webhook configuration and trigger status
- **"Calendar access denied"**: Re-authorize Google Apps Script permissions
- **"Test mode unexpected"**: Review data patterns in deacon/household lists

### **Diagnostic Tools**
- **🧪 Run Tests**: Complete system validation
- **🔍 Inspect All Triggers**: Automation status check
- **📋 Test Notification System**: Chat connectivity test
- **🔧 Validate Setup**: Configuration verification

**See [User Guide](docs/USER_GUIDE.md) for comprehensive troubleshooting.**

---

## 📚 Documentation

- **[Setup Guide](docs/SETUP.md)**: Complete installation instructions
- **[User Guide](docs/USER_GUIDE.md)**: Daily operations and troubleshooting  
- **[Features Guide](docs/FEATURES.md)**: Technical deep-dive
- **[Changelog](docs/CHANGELOG.md)**: Version history and updates

---

## 🎓 Advanced Features

### **Mathematical Excellence**
- **Harmonic resonance elimination**: Solves complex rotation mathematics
- **Prime factor analysis**: Ensures fair distribution patterns
- **Performance optimization**: Handles large schedules efficiently

### **Production-Ready**
- **Error recovery**: Graceful handling of all failure scenarios
- **Rate limiting**: Respectful Google API usage
- **Security**: No external data transmission except Google services
- **Scalability**: Tested with 20+ deacons, 50+ households

### **Development Quality**
- **Modular architecture**: Five separate, maintainable modules
- **Comprehensive testing**: Built-in validation and diagnostics
- **Zero dead code**: 100% function coverage and utilization
- **Clean interface**: Streamlined menus, no confusing options

---

## 🔄 Version History

| Version | Release | Status | Key Features |
|---------|---------|--------|--------------|
| **v1.1** | 2025-07-11 | **Current** | Menu cleanup, user guide |
| **v1.0** | 2025-07-09 | Stable | First operational release |
| v0.24.x | 2025-06-07 | Legacy | Modular architecture |
| v0.23.x | 2025-05-xx | Legacy | Pre-modular experimental |

**Semantic Versioning**: v1.0 marks the first truly operational system suitable for church production use.

---

## 🤝 Support & Development

### **System Capabilities**
- **Church-tested**: Successfully deployed in operational church environment
- **AI-assisted development**: Built with Claude.ai for rapid iteration
- **Citizen development**: Designed for church administrators, not programmers
- **Future-ready**: Architecture supports planned enhancements

### **Getting Help**
1. **User Guide**: Comprehensive operational documentation
2. **Built-in diagnostics**: System health and troubleshooting tools
3. **Error messages**: Specific, actionable guidance for common issues
4. **GitHub Issues**: Bug reports and feature requests

---

*The Deacon Visitation Rotation System v1.1 represents a mature, production-ready solution built through iterative development and real-world church testing. The system successfully combines mathematical sophistication with user-friendly operation, making advanced scheduling accessible to church administrators without technical backgrounds.*
