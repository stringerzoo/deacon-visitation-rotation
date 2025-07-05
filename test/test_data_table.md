# Test Data for Deacon Rotation System
**Safe sample data for testing v24.2 smart calendar updates**

## 📊 Spreadsheet Test Data (Columns M-S)

| Row | M: Households | N: Phone Number | O: Address | P: Breeze Link | Q: Notes Pg Link | R: Breeze Link (short) | S: Notes Pg Link (short) |
|-----|---------------|-----------------|------------|----------------|------------------|------------------------|--------------------------|
| **1** | **Households** | **Phone Number** | **Address** | **Breeze Link** | **Notes Pg Link** | **Breeze Link (short)** | **Notes Pg Link (short)** |
| **2** | Alan & Alexa Adams | (555) 123-1001 | 123 Maple Street, Louisville, KY 40202 | 12345001 | https://docs.google.com/document/d/1SampleDoc001Test/edit | *(leave empty - auto-generated)* | *(leave empty - auto-generated)* |
| **3** | Barbara & Bob Baker | (555) 123-1002 | 456 Oak Avenue, Louisville, KY 40205 | 12345002 | https://docs.google.com/document/d/1SampleDoc002Test/edit | *(leave empty - auto-generated)* | *(leave empty - auto-generated)* |
| **4** | Chloe & Charles Cooper | (555) 123-1003 | 789 Pine Boulevard, Louisville, KY 40218 | 12345003 | https://docs.google.com/document/d/1SampleDoc003Test/edit | *(leave empty - auto-generated)* | *(leave empty - auto-generated)* |
| **5** | Delilah & David Danvers | (555) 123-1004 | 321 Cedar Lane, Louisville, KY 40223 | 12345004 | https://docs.google.com/document/d/1SampleDoc004Test/edit | *(leave empty - auto-generated)* | *(leave empty - auto-generated)* |
| **6** | Emma & Edward Evans | (555) 123-1005 | 654 Birch Court, Louisville, KY 40214 | 12345005 | https://docs.google.com/document/d/1SampleDoc005Test/edit | *(leave empty - auto-generated)* | *(leave empty - auto-generated)* |

## 📋 Copy-Paste Format

For easy spreadsheet entry, here's the data in copy-paste format:

### **Column M (Households):**
```
Households
Alan & Alexa Adams
Barbara & Bob Baker
Chloe & Charles Cooper
Delilah & David Danvers
Emma & Edward Evans
```

### **Column N (Phone Numbers):**
```
Phone Number
(555) 123-1001
(555) 123-1002
(555) 123-1003
(555) 123-1004
(555) 123-1005
```

### **Column O (Addresses):**
```
Address
123 Maple Street, Louisville, KY 40202
456 Oak Avenue, Louisville, KY 40205
789 Pine Boulevard, Louisville, KY 40218
321 Cedar Lane, Louisville, KY 40223
654 Birch Court, Louisville, KY 40214
```

### **Column P (Breeze Link Numbers):**
```
Breeze Link
12345001
12345002
12345003
12345004
12345005
```

### **Column Q (Notes Page Links):**
```
Notes Pg Link
https://docs.google.com/document/d/1SampleDoc001Test/edit
https://docs.google.com/document/d/1SampleDoc002Test/edit
https://docs.google.com/document/d/1SampleDoc003Test/edit
https://docs.google.com/document/d/1SampleDoc004Test/edit
https://docs.google.com/document/d/1SampleDoc005Test/edit
```

### **Columns R & S (Auto-Generated):**
Leave these empty - they will be populated when you run "🔗 Generate Shortened URLs"

## 🧪 Testing Scenarios with This Data

### **Test 1: Contact Info Update**
1. Generate schedule and export to TEST calendar
2. Manually change Adams' event time from 2 PM to 4 PM in Google Calendar
3. In spreadsheet, change Adams' phone to `(555) 999-0001`
4. Run "📞 Update Contact Info Only"
5. **Expected Result**: Calendar event keeps 4 PM time but shows new phone number

### **Test 2: Future Events Update** 
1. Modify current week's Baker event in calendar (add location "Test Church")
2. In spreadsheet, change Baker's address to `999 Test Street, Louisville, KY 40299`
3. Run "🔄 Update Future Events Only"
4. **Expected Result**: Current week event untouched, future events have new address

### **Test 3: Monthly Update**
1. Change Cooper's Breeze number to `99999003`
2. Run "🗓️ Update This Month"
3. **Expected Result**: Only this month's Cooper events updated with new Breeze link

## 🎯 Deacon Test Data (Column L)

Don't forget your test deacons:
```
Deacons
Andy A
Brian B
Chris C
```

## 🔧 Configuration Test Data (Column K)

```
K1: Start Date
K2: (Next Monday's date)
K3: Visits every x weeks (1,2,3,4) 
K4: 2
K5: Length of schedule in weeks
K6: 8
K7: Calendar Event Instructions:
K8: TEST: Please call to confirm visit time. This is a test event.
```

## 🛡️ Safety Notes

- **Phone numbers**: All use (555) prefix - clearly fake
- **Addresses**: Generic street names in Louisville area
- **Breeze numbers**: Sequential test numbers starting with 12345
- **Google Docs**: Sample URLs that won't work (safe)
- **Names**: Alphabetical pattern makes it obvious this is test data

This data setup will let you thoroughly test all the smart calendar update functions without any risk to real member information! 🎯
