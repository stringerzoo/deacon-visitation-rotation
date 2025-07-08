# Deacon Visitation Rotation System v25.0

> **Comprehensive church deacon household visitation scheduling with automated Google Chat notifications**

[![Version](https://img.shields.io/badge/version-25.0-blue.svg)](CHANGELOG.md)
[![Google Apps Script](https://img.shields.io/badge/platform-Google%20Apps%20Script-green.svg)](https://script.google.com)
[![License](https://img.shields.io/badge/license-MIT-orange.svg)](LICENSE)

## 🎯 Overview

This system solves complex mathematical scheduling problems to create optimal deacon visitation rotations while integrating with church management systems and providing automated Google Chat notifications. Designed for churches using Google Workspace for communication and Breeze CMS for member management.

## ⭐ What's New in v25.0

### 🔔 **Google Chat Notifications** ⭐ **MAJOR FEATURE**
- **Automated weekly summaries** sent via Google Chat webhooks
- **2-week lookahead format** supporting bi-weekly visitation rhythm
- **Configurable scheduling** using spreadsheet cells (day and time)
- **Test/production separation** with automatic mode detection
- **Rich message content** including contact info and direct links

### 🏗️ **Enhanced Architecture**
- **5-module system** with dedicated notifications module
- **Modular webhook integration** that doesn't affect core scheduling
- **Configurable notification timing** stored in spreadsheet
- **Robust error handling** with Google Apps Script API limitations

### 🎛️ **Advanced Menu System**
- **Notifications submenu** with full automation control
- **Trigger management** (enable/disable/schedule weekly notifications)
- **Diagnostic tools** for troubleshooting chat integration
- **Configuration management** for webhook setup

## 📋 System Requirements

- **Google Workspace** account (Gmail accounts work for basic features)
- **Google Sheets, Apps Script, Calendar** access
- **Google Chat space** with webhook capability for notifications
- **Internet access** for URL shortening service
- **Breeze Church Management System** for profile integration
- **Processing time**: ~30-45 seconds per 100 calendar events
- **Notification delivery**: Real-time via Google Chat webhooks

## 🚀 Quick Start

1. **[Set up the spreadsheet](SETUP.md#spreadsheet-setup)** with your deacons and households
2. **[Deploy the Apps Script modules](SETUP.md#apps-script-deployment)** (5 separate files)
3. **[Configure Google Chat webhook](SETUP.md#notifications-setup)** for automated notifications
4. **Generate your first schedule** and export to calendar
5. **Test notifications** and configure weekly automation

## 🎛️ Enhanced Menu System (v25.0)

```
🔄 Deacon Rotation
├── 📅 Generate Schedule
├── 🔗 Generate Shortened URLs
├── 📆 Calendar Functions
│   ├── 📞 Update Contact Info Only
│   ├── 🔄 Update Future Events Only
│   └── 🚨 Full Calendar Regeneration
├── 📢 Notifications                         ⭐ NEW SUBMENU
│   ├── 💬 Send Weekly Chat Summary
│   ├── ⏰ Send Tomorrow's Reminders
│   ├── 🔧 Configure Chat Webhook
│   ├── 📋 Test Notification System
│   ├── 🔄 Enable Weekly Auto-Send
│   ├── 📅 Show Auto-Send Schedule
│   └── 🛑 Disable Weekly Auto-Send
├── 📊 Export Individual Schedules
├── 📁 Archive Current Schedule
├── 🗓️ Generate Next Year
├── 🔧 Validate Setup
├── 🧪 Run Tests
├── [🧪/✅] Show Current Mode
└── ❓ Setup Instructions
```

## 📊 Enhanced Spreadsheet Layout (v25.0)

### **Core Schedule (A-E)**
- **A**: Cycle, **B**: Week, **C**: Week of, **D**: Household, **E**: Deacon

### **Reports & Configuration (F-K)**
- **F**: Buffer, **G-I**: Individual deacon reports, **J**: Buffer
- **K1-K8**: Basic configuration (start date, frequency, instructions)
- **K10-K13**: Notification settings (day, time) ⭐ **NEW**
- **K15-K16**: Test mode indicators

### **Contact Data (L-S)**
- **L**: Deacons, **M**: Households, **N**: Phones, **O**: Addresses
- **P**: Breeze numbers, **Q**: Notes links, **R-S**: Shortened URLs

## 🔔 Notification System Features

### **Automated Weekly Summaries**
- **Configurable timing** - Set day of week and hour in spreadsheet
- **2-week lookahead** - Current week assignments + next week preview
- **Rich content** - Contact info, Breeze links, Notes access
- **Test/production modes** - Separate chat spaces for development

### **Manual Notification Tools**
- **Instant summaries** - Send notifications on-demand
- **Tomorrow's reminders** - Day-before visit notifications
- **Test functions** - Verify webhook configuration
- **Diagnostic tools** - Troubleshoot delivery issues

### **Smart Integration**
- **No data validation required** - Uses existing spreadsheet structure
- **Automatic mode detection** - Switches between test/production
- **Error resilience** - Graceful handling of API limitations
- **Scalable design** - Handles growing deacon/household lists

## 📖 Documentation

- **[Setup Guide](SETUP.md)** - Complete installation with notifications
- **[Features Overview](FEATURES.md)** - Detailed feature explanations
- **[Changelog](CHANGELOG.md)** - Version history including v25.0
- **[Notification Setup](SETUP.md#notifications-setup)** - Google Chat webhook configuration

## 🎯 The Math Behind It

This system solves a complex mathematical problem in rotation scheduling called "harmonic resonance." When the ratio of deacons to households creates mathematical harmonics (like 12:6 or 14:7), simple rotation algorithms cause deacons to visit only the same household(s) repeatedly.

Our algorithm uses **modular arithmetic with prime factorization** to ensure:
- Every deacon visits every household over time
- Visits are distributed evenly across the schedule
- No mathematical "locks" that prevent fair rotation
- Optimal spacing between repeat visits

## 🔧 New Functionality in v25.0

### **Google Chat Integration** ⭐ **MAJOR FEATURE**
- **Webhook-based notifications** with rich formatting and direct links
- **Configurable automation** with spreadsheet-based scheduling
- **Test mode separation** for safe development and production use
- **Comprehensive diagnostics** for troubleshooting delivery issues

### **Enhanced Architecture**
- **5-module system** - Dedicated Module 5 for notification functionality
- **Clean separation** - Notifications don't affect core scheduling logic
- **Independent deployment** - Add notifications to existing v24.2 systems
- **Future expansion** - Ready for email, SMS, or other notification types

### **Advanced Configuration**
- **Data validation** - Dropdown menus prevent configuration errors
- **Flexible timing** - Any day of week, any hour (24-hour format)
- **Multiple environments** - Test webhooks separate from production
- **Visual feedback** - Clear status indicators and confirmation dialogs

### **Improved User Experience**
- **Streamlined setup** - Step-by-step webhook configuration
- **Clear feedback** - Success/failure notifications with detailed logging
- **Self-diagnostic** - Built-in tools for identifying and fixing issues
- **Scalable design** - Performance tested with 12 deacons, 5+ households

## 📋 Development Workflow

### **GitHub Repository Structure**
```
📁 deacon-visitation-rotation
├── src/
│   ├── module1-core-config.js
│   ├── module2-algorithm.js
│   ├── module3-smart-calendar.js
│   ├── module4-export-menu.js
│   └── module5-notifications.js          ⭐ NEW
├── docs/
│   ├── README.md
│   ├── SETUP.md
│   ├── FEATURES.md
│   └── CHANGELOG.md
└── examples/
    └── webhook-setup-guide.md             ⭐ NEW
```

### **Deployment Process**
1. **Edit modules** individually in GitHub
2. **Copy each file** to corresponding .gs file in Apps Script
3. **Test functionality** including notification features
4. **Configure webhooks** for your Google Chat space
5. **Commit changes** to GitHub when verified working

## 🛡️ Security & Best Practices

### **Data Protection**
- **Webhook URLs** stored securely in Apps Script properties
- **Test/production separation** prevents accidental notifications
- **Member data privacy** maintained in all notification content
- **No external dependencies** beyond Google services and TinyURL

### **Development Safety**
- **Sample data patterns** for safe testing (Alan Adams, Barbara Baker, etc.)
- **Test mode detection** prevents production notifications during development
- **Graceful error handling** with informative user feedback
- **Rate limiting protection** for all Google APIs

## 🆘 Support & Troubleshooting

### **Common Issues**
- **Notifications not sending**: Check webhook configuration and test mode
- **Google Apps Script triggers**: May have 15-20 minute delays
- **Calendar access**: Verify permissions for shared calendars
- **Link errors**: Ensure Breeze numbers and Notes URLs are correct

### **Getting Help**
- Review the **[Setup Guide](SETUP.md)** for configuration details
- Check **[Features documentation](FEATURES.md)** for functionality explanations
- Use built-in **diagnostic tools** in the Notifications menu
- Consult **[Changelog](CHANGELOG.md)** for version-specific information

---

**Ready to revolutionize your church's visitation ministry with automated coordination?** 🚀

*For churches seeking to maintain meaningful pastoral care while leveraging technology for better coordination and communication.*
