# Features Overview - Deacon Visitation Rotation System

This document provides detailed information about each feature in the deacon visitation rotation system.

## 🎯 Core Rotation Engine

### Optimal Pattern Generator
**What it does**: Creates mathematically optimal rotation schedules that eliminate harmonic resonance issues.

**The Problem**: Simple rotation algorithms fail when the ratio of deacons to households creates "harmonic locks" (e.g., 12:6 ratio causes each deacon to only visit the same household repeatedly).

**Our Solution**: 
- Intelligent scoring system that prioritizes deacons with fewer visits
- Variety bonuses for new deacon-household pairings
- Prevention of same-cycle double assignments
- Dynamic adaptation to any deacon/household ratio

**Benefits**:
- Every deacon visits every household over time
- Fair workload distribution (visit counts within 1-2 of each other)
- Eliminates boring repetition in assignments

### Flexible Visit Frequencies
**Supported intervals**:
- Weekly (every 1 week)
- Bi-weekly (every 2 weeks) - *default*
- Every 3 weeks
- Monthly (every 4 weeks)
- Custom intervals up to 8 weeks

**How it works**: The system calculates optimal cycle timing while maintaining the visit frequency for each household.

## 📊 Schedule Generation

### Main Rotation Display
**Columns A-E** contain:
- **Cycle**: Which rotation cycle (1, 2, 3...)
- **Week**: Week number in the overall schedule  
- **Week of**: Starting date for the visit week
- **Household**: Which household to visit
- **Deacon**: Assigned deacon for the visit

### Individual Deacon Reports
**Columns G-I** show personalized schedules:
- Each deacon's complete visit list
- Chronologically sorted dates
- Visit count summaries
- Easy-to-read format for distribution

## 🏠 Household Management

### Contact Information Storage
**Columns N-Q** (expandable) support:
- **Phone Numbers**: Primary contact numbers
- **Addresses**: Full household addresses  
- **Custom Fields**: Breeze links, Google Docs, special notes
- **Future expansion**: Add columns as needed without affecting other features

### Data Integration
- Automatically includes contact info in calendar events
- Used in individual deacon schedule exports
- Preserved in archived schedules
- Supports special characters (ampersands, apostrophes, etc.)

## 📅 Google Calendar Integration

### Enhanced Calendar Events
**Event titles**: "[Deacon Name] visits [Household Name]"
- Example: "Andy A visits Barbara Baker"

**Event descriptions include**:
- Deacon assignment
- Household name
- Phone number
- Address
- Custom coordination instructions

### Calendar Management
- Creates separate "Deacon Visitation Schedule" calendar
- Avoids cluttering personal calendars
- 1-hour default duration (2:00-3:00 PM)
- Option to clear existing events before adding new ones
- Maximum 500 events (prevents system overload)

### Coordination Instructions
Customizable text in cell K8 that appears in every calendar event. Default:
> "Please call to confirm visit time. Contact family 1-2 days before scheduled date to arrange convenient time."

## 📈 Export Capabilities

### Individual Deacon Schedules
**Creates separate spreadsheets** with:
- Each deacon's complete schedule
- **Week of** column (clarifies this is target week, not fixed date)
- **Household** column with visit assignments  
- **Notes** column (250px wide, text-wrapped) for visit documentation
- Professional formatting and proper column sizing

### CSV Export
- Standard CSV format for import into other systems
- All schedule data included
- Compatible with Excel and other spreadsheet programs

## 🔧 System Management

### Validation and Error Handling
**Setup Validation**:
- Checks for minimum requirements (2+ deacons, 1+ household)
- Validates data types and ranges
- Detects duplicate names
- Identifies problematic characters
- Warns about workload feasibility

**Workload Analysis**:
- Calculates visits per deacon
- Warns if workload too high (< 3 weeks between visits)
- Alerts if workload too low (> 12 weeks between visits)
- Provides specific suggestions for optimization

### System Tests
**Built-in diagnostics**:
- Configuration loading tests
- Date calculation validation
- Script permissions verification
- Pattern generation testing
- Detailed console logging for troubleshooting

## 📁 Archive and History

### Schedule Archiving
**Archive function** creates:
- Complete copy of current schedule
- Includes main rotation and deacon reports
- Timestamped sheet names
- Preserves historical records

### Version Management
- Tracks configuration changes
- Maintains schedule continuity across years
- Supports rollback to previous versions

## 🔄 Advanced Features

### Next Year Generation
**Seamless year-over-year transition**:
- Automatically sets next year's start date
- Continues rotation pattern from where previous year ended
- Maintains fairness across multiple years
- Preserves all configuration settings

### Auto-Update Prevention
- Manual control over schedule regeneration
- Prevents accidental overwrites
- Allows multiple configuration changes before regeneration

### Harmonic Detection and Mitigation
**Automatic identification** of problematic ratios:
- 12:6, 12:4, 12:3 (simple harmonics)
- 14:7, 15:6 (complex harmonics)
- Real-time pattern analysis and correction

## ⚙️ Configuration Options

### Flexible Schedule Length
- 1 to 520 weeks (up to 10 years)
- Common options: 26 weeks, 52 weeks, 78 weeks
- Automatic cycle calculation

### Safety Limits
- Maximum 2000 total visits (prevents execution timeouts)
- 4-minute execution time limit with progress monitoring
- Batch processing for large schedules (1000 visits per batch)

### Header and Label Customization
- Automatic header generation
- Color-coded sections (blue for data, yellow for configuration)
- Professional formatting with borders and fonts

## 🛡️ Reliability Features

### Error Recovery
- Graceful handling of invalid dates
- Fallback options for configuration failures
- Detailed error messages with specific solutions
- Automatic retry logic for temporary failures

### Data Integrity
- Input validation and sanitization
- Prevention of circular references
- Consistent data formatting
- Backup and recovery suggestions

### Performance Optimization
- Efficient algorithm implementation
- Memory management for large datasets
- Batch processing to avoid API limits
- Progress tracking for long operations

## 📝 User Interface

### Menu System
Organized menu with logical groupings:
- **Generate Schedule**: Main functionality
- **Export options**: Individual schedules, calendar integration  
- **Management**: Archive, next year generation
- **Utilities**: Validation, testing, instructions

### User Feedback
- Progress indicators for long operations
- Success/failure notifications
- Detailed status messages
- Quality assessments (Excellent/Good/Needs Improvement)

---

This feature set represents a comprehensive solution for church deacon visitation management, combining mathematical optimization with practical usability. 🎯
