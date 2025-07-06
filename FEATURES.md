# Technical Features Reference - Deacon Visitation Rotation System

This document provides deep technical explanations of the system's core algorithms, implementation details, and advanced features. For setup instructions, see [SETUP.md](SETUP.md). For project overview, see [README.md](README.md).

## 🎯 Core Rotation Engine

### Optimal Pattern Generator Algorithm
**The Problem**: Simple rotation algorithms fail when the ratio of deacons to households creates "harmonic locks" - mathematical resonance patterns where deacons only visit the same household(s) repeatedly.

**Examples of Problematic Ratios:**
- **12:6 ratio** → Each deacon visits only 2 specific households forever
- **14:7 ratio** → Complex harmonic creates uneven distribution
- **15:6 ratio** → Some deacons get locked out of certain households

**Our Solution - Intelligent Scoring System:**
```javascript
// Simplified scoring algorithm
for (let deaconIndex = 0; deaconIndex < deaconCount; deaconIndex++) {
  const totalVisits = deaconUsageCount[deaconIndex];
  const householdVisits = deaconHouseholdPairs.get(`${deaconIndex}-${householdIndex}`);
  
  // Lower score is better
  let score = totalVisits * 100 + householdVisits * 10;
  
  // Variety bonus for new pairings
  if (householdVisits === 0) score -= 50;
}
```

**Algorithm Benefits:**
- **Pre-calculated patterns** eliminate real-time harmonic issues
- **Intelligent scoring** prioritizes deacons with fewer visits
- **Variety bonuses** favor new deacon-household pairings
- **Cycle prevention** ensures no double assignments per cycle
- **Universal compatibility** works with any deacon/household ratio

### Harmonic Resonance Analysis
**Mathematical Foundation:**
```javascript
const gcd = (a, b) => b === 0 ? a : gcd(b, a % b);
const lcm = (a, b) => (a * b) / gcd(a, b);

const commonFactor = gcd(deaconCount, householdCount);
const isSimpleHarmonic = deaconCount % householdCount === 0;
const isComplexHarmonic = commonFactor > 1 && !isSimpleHarmonic;
```

**Harmonic Types Detected:**
- **Simple Harmonics**: Perfect divisors (12:6, 12:4, 12:3)
- **Complex Harmonics**: Common factors (14:7, 15:6, 21:9)
- **Prime Ratios**: No harmonics (13:7, 11:5) - naturally optimal

**Mitigation Strategies:**
- **Pattern pre-generation** bypasses harmonic calculations entirely
- **Dynamic scoring** adapts to any ratio automatically
- **Quality metrics** validate final rotation patterns
- **Fallback mechanisms** ensure schedule generation never fails

## 🔄 Smart Calendar Update System (v24.2)

### Event Preservation Architecture
**Core Challenge**: Google Calendar events contain both system data (contact info) and user customizations (times, guests, locations). Traditional approaches lose user customizations.

**Smart Parsing Implementation:**
```javascript
// Extract household from event title: "[Deacon] visits [Household]"
const visitMatch = title.match(/(.+) visits (.+)/);
if (visitMatch) {
  const household = visitMatch[2].trim();
  // Match to current spreadsheet data for updates
}
```

**Preservation Strategies:**
- **Description-only updates** preserve all event metadata
- **Selective event targeting** by date ranges
- **Smart data matching** handles household name changes
- **Error isolation** prevents cascade failures

### Dynamic Mode Detection Algorithm
**Intelligence Triggers:**
```javascript
const testIndicators = [
  () => households.some(h => testPatterns.some(p => h.toLowerCase().includes(p))),
  () => phones.some(phone => String(phone).includes('555')),
  () => breezeNumbers.some(num => String(num).startsWith('12345')),
  () => spreadsheetName.includes('test') || spreadsheetName.includes('sample')
];

const isTestMode = testIndicators.some(check => check());
```

**Real-Time Updates:**
- **Dynamic re-detection** on each calendar operation
- **Visual indicator refresh** without spreadsheet reload
- **Mode-aware dialogs** with current state information
- **Calendar name adaptation** based on detected mode

### Rate Limiting Implementation
**Google Calendar API Protection:**
```javascript
// Deletion rate limiting
if ((index + 1) % 10 === 0) {
  Utilities.sleep(500); // 0.5 second pause every 10 deletions
}

// Creation rate limiting  
if ((index + 1) % 25 === 0) {
  Utilities.sleep(1000); // 1 second pause every 25 creations
  console.log(`Created ${index + 1} events so far...`);
}

// Phase transition cooldown
Utilities.sleep(2000); // 2 second wait between deletion and creation
```

**Performance Optimization:**
- **Batch processing** with intelligent delays
- **Progress tracking** for transparency
- **Error recovery** with individual event isolation
- **API quota management** prevents service disruption

## 🔗 Integration Systems

### Church Management API Integration
**Breeze CMS Architecture:**
```javascript
function buildBreezeUrl(breezeNumber) {
  const cleanNumber = breezeNumber.trim();
  return `https://immanuelky.breezechms.com/people/view/${cleanNumber}`;
}
```

**URL Shortening Service:**
```javascript
// TinyURL integration with fallback
const apiUrl = 'http://tinyurl.com/api-create.php?url=' + encodeURIComponent(longUrl);
const response = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });

if (response.getResponseCode() === 200) {
  return shortUrl; // Success
} else {
  return longUrl; // Fallback to original
}
```

**Integration Benefits:**
- **No account requirements** for URL shortening
- **Automatic fallback** to full URLs if shortening fails
- **Batch processing** with rate limiting respect
- **Persistent storage** in dedicated columns for reuse

### Modular File System Design (v24.2)
**Native Google Apps Script Architecture:**
```
Module1_Core_Config.gs     - Configuration, validation, header setup
Module2_Algorithm.gs       - Core rotation algorithm, pattern generation  
Module3_Smart_Calendar.gs  - Calendar updates, mode detection
Module4_Export_Menu.gs     - Full export, individual schedules, menu system
```

**Cross-Module Communication:**
- **Shared constants** accessible across all modules
- **Function availability** throughout execution context
- **Global variables** for configuration state
- **Error propagation** with module-specific identification

## ⚙️ Advanced Configuration

### Algorithm Tuning Parameters
**Scoring System Weights:**
```javascript
let score = totalVisits * 100 + householdVisits * 10;
//           ↑ Primary weight    ↑ Secondary weight

if (householdVisits === 0) score -= 50; // Variety bonus
//                                  ↑ Incentive weight
```

**Pattern Quality Metrics:**
- **Visit imbalance**: Difference between max and min visits per deacon
- **Coverage percentage**: Deacons visiting all households
- **Variety index**: Average households per deacon

### Performance Optimization Settings
**Execution Time Management:**
```javascript
const maxExecutionTime = 4 * 60 * 1000; // 4 minutes safety buffer
const batchSize = 1000; // Events per batch write
const maxEvents = 500; // Calendar event limit
```

**Memory Management:**
- **Batch processing** for large datasets
- **Progressive updates** with status logging
- **Resource monitoring** with time limit enforcement
- **Graceful degradation** when approaching limits

### Error Recovery Mechanisms
**Validation Layers:**
1. **Input validation** - data types, ranges, duplicates
2. **Workload validation** - feasibility checking with recommendations
3. **Runtime validation** - date calculations, API responses
4. **Output validation** - schedule completeness verification

**Recovery Strategies:**
```javascript
try {
  // Primary operation
} catch (eventError) {
  console.error(`Failed operation: ${context}`, eventError);
  // Continue with remaining operations - don't fail entire process
}
```

**Safety Mechanisms:**
- **Individual error isolation** prevents cascade failures
- **Automatic retries** for transient failures
- **Fallback options** for service unavailability
- **User guidance** with specific error resolution steps

## 🧪 Quality Analysis and Logging

### Pattern Validation
**Real-Time Analysis:**
```javascript
// Calculate balance metrics
const avgVisitsPerDeacon = schedule.length / config.deacons.length;
const visitImbalance = maxVisits - minVisits;
const coveragePercentage = (deaconsWithFullCoverage / config.deacons.length) * 100;

// Quality assessment
if (visitImbalance <= 1 && coveragePercentage >= 80) {
  console.log('✅ EXCELLENT: Well-balanced rotation with good variety');
}
```

**Quality Grades:**
- **EXCELLENT**: Visit imbalance ≤ 1, coverage ≥ 80%
- **GOOD**: Visit imbalance ≤ 2, coverage ≥ 60%
- **NEEDS IMPROVEMENT**: Higher imbalance or limited variety

### Diagnostic System
**Comprehensive Testing Framework:**
- **Configuration loading** validation
- **URL shortening** service connectivity
- **Breeze URL construction** accuracy
- **Calendar access** permission verification
- **Script permissions** for external services

**Performance Monitoring:**
```javascript
const startTime = new Date().getTime();
// ... operation ...
console.log(`Operation completed in ${new Date().getTime() - startTime}ms`);
```

## 🔧 Implementation Details

### Data Structure Optimization
**Efficient Pattern Storage:**
```javascript
const deaconHouseholdPairs = new Map();
// Key: "deaconIndex-householdIndex"
// Value: visit count
```

**Memory-Efficient Processing:**
- **Map-based tracking** for O(1) lookup performance
- **Incremental updates** rather than full recalculation
- **Garbage collection** awareness in loop structures

### API Integration Best Practices
**External Service Reliability:**
- **Timeout handling** for network requests
- **Rate limiting** respect for service quotas
- **Graceful degradation** when services unavailable
- **Caching strategies** for repeated operations

**Google Apps Script Optimization:**
- **Batch operations** to minimize API calls
- **Efficient data structures** for large datasets
- **Memory management** for long-running operations
- **Execution time monitoring** with safety limits

---

**This technical reference provides the deep implementation details needed for system customization, troubleshooting, and contribution. For user-facing documentation, refer to the README and SETUP guides.** 🎯
