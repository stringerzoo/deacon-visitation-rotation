# Changelog - Deacon Visitation Rotation System

Code was generated by Claude.ai based on requirements, testing, and feedback initiated by Scott.

All notable changes to this project will be documented in this file.

---

## 🏷️ Version Numbering Transition

**Note**: Starting with v1.1, we're adopting semantic versioning. Previous versions were experimental iterations:
- v0.1-v0.22: Rapid experimental iterations in single chat thread
- v0.23.x: Monolithic code, pre-modular architecture  
- v0.24.x: Introduction of modular architecture, major code restructuring
- v1.0: First truly operational release with Google Chat notifications
- v1.1: Menu cleanup and refinements (this release)

---

## [v1.1.1] - 2025-07-14 ⭐ **CURRENT RELEASE**

### 🐛 Critical Bug Fixes
#### **Data Structure Consistency**
- **Fixed object/array data handling inconsistency** causing "not iterable" errors
- **Corrected `updateContactInfoOnly()`** - now properly handles schedule objects
- **Fixed `updateFutureEventsOnly()`** - eliminated array destructuring on objects
- **Fixed `exportIndividualSchedules()`** - consistent object property access

#### **Enhanced URL Management**
- **Smart URL preservation** - `generateShortUrls()` now only processes empty cells
- **Existing URL protection** - preserves working shortened URLs
- **Force regeneration option** - new `forceRegenerateAllShortUrls()` function
- **Improved user dialogs** - clear explanation of preserve vs. overwrite behavior

### 🔧 Technical Improvements
- **Architectural consistency** - all schedule data handled as objects throughout
- **Error prevention** - added data consistency guidelines for future development
- **Performance optimization** - fewer unnecessary API calls for URL shortening
- **Enhanced reporting** - detailed summaries of preserved vs. generated URLs

### 📋 Breaking Changes
**None** - All changes are backward compatible improvements

### 🛠️ Developer Guidelines
- **Added comprehensive data handling guidelines** to prevent future array/object confusion
- **Established architectural principles** for schedule data consistency
- **Enhanced error messages** with proper object structure expectations

---

## [v1.1] - 2025-07-13

### 🔗 **Automatic Calendar URL Detection**
- **Eliminated K19 manual configuration** - Calendar URLs now auto-generated from calendar system
- **Direct URL generation** - `generateCalendarUrlDirect()` function creates URLs on-demand
- **Backwards compatibility** - Existing systems continue to work without changes
- **Zero configuration** - Calendar links appear in notifications automatically

### 🎨 **Improved Column K Layout**
- **K18-K19 repurposed** - Now serves as section header for notification links
- **Clear separation** - Configuration (K1-K16) vs Notification Content (K18-K25)
- **Space optimization** - Eliminated redundant calendar URL field
- **Future-ready** - Available space for additional features

### 🧹 **Code Cleanup**
- **Removed K19 dependency** - `getResourceLinksFromSpreadsheet()` updated for direct generation
- **Enhanced Module 5** - Auto-detection integrated into notification system
- **Simplified setup** - One less manual configuration step for users

### 🌍 **Automatic Timezone Detection**
- **Auto-detected timezones** - Calendar embed URLs use `Session.getScriptTimeZone()` instead of hardcoded Eastern
- **Global compatibility** - System now works correctly for users in any timezone
- **No configuration needed** - Timezone automatically matches user's location

### 🧹 **Menu System Cleanup**
- **Removed**: `generateNextYearSchedule` function and menu item
  - **Reason**: Assumes same deacon/household roster year-to-year (unrealistic)
  - **Alternative**: Manual regeneration with updated lists is more practical
- **Clarified**: Distinct purposes of test notification functions
  - `testNotificationSystem`: Tests webhook connectivity with simple message
  - `testNotificationNow`: Tests full notification system with actual weekly summary
- **Streamlined**: Menu organization for better user experience

### 📖 **Documentation Enhancements**
- **User Guide**: Comprehensive operational guide for daily use and troubleshooting
- **Version Transition**: Updated all documentation to reflect semantic versioning
- **Error Message Guide**: Clear explanations of what error messages mean and how to fix them

### 🔧 **Code Quality Improvements**
- **Orphaned Function Analysis**: Confirmed no dead code in system
- **Function Coverage**: 100% of functions properly connected and used
- **Menu Coverage**: 100% of menu items map to working functions

---

## [v1.0] - 2025-07-09 **Previous Release**

### 🔔 **Added - Google Chat Notifications (Major Feature)**
- **Module 5 (Notifications)**: New 1288-line module for comprehensive chat integration
- **Automated weekly summaries** via Google Chat webhooks with 2-week lookahead format
- **Configurable scheduling** using spreadsheet cells (K11: day, K13: time)
- **Test/production separation** with automatic mode detection and separate webhooks
- **Rich message formatting** with contact info, Breeze links, and Notes access
- **Manual notification tools**: On-demand summaries and tomorrow's reminders
- **Comprehensive diagnostics**: Test functions and troubleshooting tools
- **Trigger management**: Enable/disable/schedule weekly automation via menu

### 🗓️ **Added - Configurable Resource Links**
- **K19 Calendar URL**: "📅 View Visitation Calendar" links in chat notifications
- **K22 Guide URL**: "📋 Visitation Guide" links for ministry procedures and guidelines
- **K25 Summary URL**: "📊 Schedule Summary" links for archived schedules and summaries
- **Test/production switching**: Easy environment changes by updating URLs in K19/K22/K25
- **Graceful degradation**: Links only appear when respective fields are configured
- **Mobile optimization**: All resource links work on mobile devices for field access
- **Calendar integration**: Direct access to visitation schedule from notifications

### 🏗️ **Enhanced - System Architecture**
- **5-module structure**: Dedicated Module 5 for notification functionality
- **Modular webhook integration**: Notifications don't affect core scheduling logic
- **Independent deployment**: Add notifications to existing v0.24.2 systems
- **Cross-module communication**: Shared constants and function availability
- **Enhanced error handling**: Module-specific error identification and recovery

### 🎛️ **Enhanced - Menu System**
- **Notifications submenu**: Complete automation control (📢 Notifications)
  - 💬 Send Weekly Chat Summary
  - ⏰ Send Tomorrow's Reminders
  - 🔧 Configure Chat Webhook
  - 📋 Test Notification System
  - 🔄 Enable Weekly Auto-Send
  - 📅 Show Auto-Send Schedule
  - 🛑 Disable Weekly Auto-Send
- **Enhanced menu structure**: Logical grouping with submenus for better organization
- **Smart calendar functions**: Contact info only, future events only, full regeneration
- **Advanced diagnostic tools**: Trigger inspection, system tests, setup validation

### 📊 **Enhanced - Smart Calendar Updates (v0.24.2 → v1.0)**
**Previous Challenge**: Contact information updates required complete calendar regeneration, losing all custom scheduling details.

**Additional Challenge**: Manual test mode configuration buried in code caused confusion and required technical knowledge to manage.

**Solution**: Smart calendar update system with flexible options and intelligent test mode detection:
- **Contact info only** - preserves ALL scheduling details
- **Future events only** - protects current week planning  
- **Full regeneration** - enhanced warnings with safer alternatives suggested
- **Native modular architecture** - eliminates complex file combination workflow
- **Automatic test mode detection** - no manual configuration required, seamless transition from testing to production

### 🔄 **Development Workflow Improvements**
#### **Previous Workflow:**
```
Edit modules in GitHub → Combine into single file → Copy to Apps Script → Test
```

#### **New Workflow (v0.24.2 → v1.0):**
```
Edit modules in GitHub → Copy individual files to Apps Script → Test
```

### 📈 **Architecture Benefits**
- **50% faster development** - no file combination step required
- **Better debugging** - errors show exact file and function location
- **Easier maintenance** - edit only the module you need to change
- **Cleaner git history** - changes show specific functionality modified
- **Team collaboration** - multiple developers can work on different modules

### 📋 **File Structure Changes**
```
📁 deacon-visitation-rotation
├── src/
│   ├── module1-core-config.js           (~500 lines)
│   ├── module2-algorithm.js             (~470 lines)
│   ├── module3-smart-calendar.js        (~200 lines)
│   ├── module4-export-menu.js           (~400 lines)
│   └── module5-notifications.js         (~1288 lines) ⭐ NEW
├── docs/
│   ├── README.md                        (Updated for v1.0)
│   ├── SETUP.md                         (Updated with notifications)
│   ├── FEATURES.md                      (Updated with chat integration)
│   └── CHANGELOG.md                     (This file)
└── test/
    └── test_data-table.md               ⭐ NEW
```

### 🆕 **New Functions (Module 5)**
- `sendWeeklyVisitationChat()` - Main weekly notification function
- `sendTomorrowReminders()` - Day-before visit notifications
- `buildChatMessage()` - Rich message formatting with resource links
- `sendToChatSpace()` - Webhook communication handler
- `getUpcomingVisits()` - Smart visit filtering with calendar week logic
- `configureNotifications()` - Interactive webhook setup
- `testNotificationSystem()` - Comprehensive testing suite
- `createWeeklyNotificationTrigger()` - Automated scheduling setup
- `getResourceLinksFromSpreadsheet()` - K19/K22/K25 URL retrieval
- `validateScheduleDataSync()` - Data integrity checking

### 🔧 **Modified Functions**
- `setupHeaders()` (Module 1): Added K10-K13, K18-K19, K21-K22, K24-K25 configuration
- `createMenuItems()` (Module 4): Enhanced with notifications submenu
- `addModeIndicatorToSheet()` (Module 3): Updated to use K15-K16
- `getCurrentTestMode()` (Module 3): Enhanced detection with new patterns

### 📚 **Documentation Updates**
- **README.md**: Added Google Chat features and K19 configuration
- **SETUP.md**: Comprehensive notifications setup guide with calendar links
- **FEATURES.md**: Technical deep-dive into notification system architecture
- **New webhook setup guide**: Step-by-step Google Chat integration

### ⚠️ **Known Limitations**
- **Google Apps Script triggers**: 15-20+ minute delays are normal
- **New triggers**: Take 24-48 hours to stabilize (Google platform limitation)
- **Manual webhook setup**: Breeze API integration planned for future release
- **Chat space dependencies**: System requires active Google Chat space

### 🔧 **Breaking Changes from v0.24.x**
- **K11-K12 moved to K15-K16**: Test mode indicators relocated for notification config
- **New dependencies**: Module 5 requires Google Chat webhook configuration
- **Menu structure**: Notification functions moved to dedicated submenu

---

## [v0.24.2] - 2025-06-07 **FINAL EXPERIMENTAL RELEASE**

### 🔄 **Major Modular Refactoring**
- **Modular code architecture** with four separate, maintainable modules
- **Optimized column layout** for improved user experience and data organization
- **Enhanced deacon reports positioning** for better workflow integration
- **Streamlined configuration placement** for easier setup and maintenance

### 🎯 **Column Layout (Enhanced in v0.24.1, Maintained in v0.24.2)**
- **A-E**: Main rotation schedule
- **F**: Buffer column for visual separation
- **G-I**: Individual deacon reports (moved from L-N for better adjacency)
- **J**: Buffer column for visual separation
- **K**: Configuration settings (consolidated)
- **L**: Deacon names list
- **M**: Household names list
- **N-O**: Contact information (phone numbers and addresses)
- **P-S**: Breeze and Notes integration (added in v0.24.1)

### 📁 **Modular Architecture Evolution**
- **v0.24.0**: Three combined modules for easier deployment
- **v0.24.1**: Enhanced modules with Breeze/Notes integration
- **v0.24.2**: Native Google Apps Script files for direct development

---

## [v0.24.1] - 2025-06-07

### 🔗 **Breeze & Notes Integration**
- **Breeze Church Management System integration** with profile links
- **Google Docs visit notes integration** with direct access
- **Automatic URL shortening** using TinyURL's free API
- **Enhanced calendar events** with clickable Breeze profiles and notes pages
- **Intelligent link handling** with fallback support

### 🎯 **New Column Layout**
- **Columns P-S**: Breeze and Notes integration
  - **P**: Breeze Link (8-digit numbers)
  - **Q**: Notes Pg Link (Google Doc URLs)
  - **R**: Breeze Link (short) - auto-generated
  - **S**: Notes Pg Link (short) - auto-generated

### 🔧 **Enhanced Functions**
- `generateAndStoreShortUrls()` - Batch URL shortening with rate limiting
- `buildBreezeUrl()` - Construct full Breeze URLs from 8-digit numbers
- `shortenUrl()` - TinyURL API integration with fallback handling
- Enhanced calendar export with clickable links

---

## [v0.24.0] - 2025-06-07

### 🔄 **Major Modular Refactoring**
- **Modular code architecture** with three separate, maintainable modules
- **Optimized column layout** for improved user experience and data organization
- **Enhanced deacon reports positioning** for better workflow integration
- **Streamlined configuration placement** for easier setup and maintenance

---

## Development History

### [v0.23.1] - Previous Stable Version
- **Enhanced system testing** with detailed logging and diagnostics
- **Robust error handling** throughout all functions
- **Contact information integration** (phone numbers and addresses)
- **Enhanced calendar export** with contact details and custom instructions
- **Archive function fixes** for proper data preservation

### [v0.23.0] - Major Algorithm Overhaul
- **Pre-calculated optimal rotation patterns** replacing real-time generation
- **Intelligent scoring system** for deacon assignments
- **Complete elimination** of harmonic resonance issues

### [v0.20.0-v0.22.0] - Contact Information & Anti-Harmonic Development
- **Contact information storage** in dedicated columns
- **Harmonic resonance detection** and mitigation
- **Enhanced calendar events** with phone numbers and addresses

### [v0.15.0-v0.19.0] - Robustness and Safety
- **Comprehensive input validation** for all configuration fields
- **Workload feasibility checking** with specific recommendations
- **Safety limits** for large schedules and execution timeouts

### [v0.10.0-v0.14.0] - Export and Integration Features
- **Google Calendar export** with automated event creation
- **Individual deacon schedules** in separate spreadsheets
- **Archive functionality** for historical preservation

### [v0.5.0-v0.9.0] - Core Algorithm Development
- **Basic rotation algorithms** with round-robin distribution
- **Date calculation and validation** systems
- **Google Sheets integration** and data reading

### [v0.1.0-v0.4.0] - Foundation and Setup
- **Basic project structure** and Google Apps Script setup
- **Simple rotation logic** and schedule generation

---

## 🐛 Major Issues Resolved

### Development Workflow Efficiency (v0.24.2)
**Issue**: Complex file combination workflow slowed development and made debugging difficult.

**Solution**: Native Google Apps Script modular architecture with:
- Four separate .gs files for direct editing
- Elimination of file combination step
- Individual module testing capabilities
- Better error location identification
- Cleaner version control with module-specific changes

### Smart Calendar Update Challenge (v0.24.2)
**Issue**: Contact information changes required complete calendar regeneration, losing all deacon scheduling customizations (modified times, dates, guests, locations).

**Solution**: Implemented flexible calendar update system with:
- Contact info only updates that preserve ALL scheduling details
- Future events only updates that protect current week's planning
- Enhanced warnings for destructive full regeneration
- Smart event parsing and data matching
- Comprehensive user guidance and confirmation dialogs

### Google Calendar API Rate Limiting (v0.24.1)
**Issue**: Rapid deletion and creation of calendar events triggered Google's API rate limits, causing "too many calendar operations" errors.

**Solution**: Intelligent rate limiting system:
- Small delays between API operations (0.5-2 seconds)
- Batch processing with progress indicators
- Graceful degradation and retry mechanisms
- User feedback during long operations
- Optimized event creation/update patterns

### Harmonic Resonance Algorithm Issue (v0.23.0)
**Issue**: When deacon count and household count shared common factors, simple rotation algorithms created unfair distributions where some deacons only visited specific households.

**Example**: 12 deacons, 6 households (ratio 2:1) caused each deacon to visit only 2 specific households repeatedly.

**Solution**: Mathematical breakthrough using modular arithmetic:
- Prime factor analysis to detect harmonic patterns
- Offset-based rotation patterns that guarantee fair distribution
- Pre-calculated optimal rotation matrices
- Complete elimination of mathematical "locks"
- Verification algorithms to ensure fairness

### Memory and Performance Issues (v0.15.0-v0.19.0)
**Issue**: Large schedules (52+ weeks, 10+ deacons) caused script timeouts and memory limitations.

**Solution**: Performance optimization package:
- Execution time monitoring with safety limits
- Efficient data structures and algorithms
- Batch processing for large datasets
- Progress indicators for long operations
- Memory-conscious array handling

---

## 📊 **System Evolution Metrics**

| Version | Lines of Code | Modules | Functions | Features |
|---------|---------------|---------|-----------|----------|
| v0.1    | ~50          | 1       | 3         | Basic rotation |
| v0.10   | ~300         | 1       | 8         | Calendar export |
| v0.20   | ~800         | 1       | 15        | Contact integration |
| v0.23   | ~1200        | 1       | 20        | Harmonic resolution |
| v0.24   | ~1500        | 4       | 25        | Modular architecture |
| **v1.0** | **~2800**   | **5**   | **50+**   | **Full operational** |
| **v1.1** | **~2800**   | **5**   | **48**    | **Refined operational** |

---

## 🚀 **Future Roadmap**

### **v1.2** - Enhanced Documentation & Automation
- Schedule summary export automation
- Enhanced user documentation
- Advanced troubleshooting guides

### **v1.3** - Data Integration
- Breeze API integration for automated contact sync
- Enhanced calendar synchronization
- Historical reporting features

### **v2.0** - Major Architecture Evolution
- Web-based interface option
- Advanced notification channels (SMS, email)
- Multi-church deployment support
- Enhanced security and permissions model

---

*This changelog reflects the evolution from experimental prototype (v0.x) to production-ready system (v1.0+). The transition to semantic versioning at v1.0 marks the first truly operational release suitable for church production use.*
