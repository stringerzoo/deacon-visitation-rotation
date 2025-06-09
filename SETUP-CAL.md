## 📅 Calendar Export Process & Timing

### **What Happens During Calendar Export**

The enhanced calendar export process includes several steps to ensure reliable operation:

1. **Configuration Loading**: Reads all household data, contact info, and links
2. **URL Preparation**: Uses shortened URLs if available, falls back to full URLs
3. **Calendar Setup**: Creates or accesses "Deacon Visitation Schedule" calendar
4. **Optional Event Deletion**: If you choose to clear existing events (with delays)
5. **Event Creation**: Creates enhanced events with rich descriptions (with rate limiting)
6. **Progress Monitoring**: Logs progress and handles any individual failures

### **Runtime Expectations**

#### **Simple Heuristic:**
**"Expect approximately 30-45 seconds per 100 events"**

#### **Detailed Timing Formula:**
```
Total Runtime = (Number of Events × 0.35 seconds) + (Number of Events ÷ 25 seconds)
                    ↑                                ↑
            Google Calendar API processing    Rate limiting delays
```

#### **Real-World Examples:**
| Events | Expected Time | What You'll See |
|--------|---------------|-----------------|
| 25 | 10-15 seconds | Very quick |
| 50 | 20-30 seconds | Quick completion |
| 100 | 30-45 seconds | Standard processing |
| 150 | 45-60 seconds | Longer but reliable |
| 200 | 60-90 seconds | Extended processing |
| 300+ | 2-3 minutes | Large schedule handling |

#### **Factors That Affect Runtime:**
- **Internet connection speed** to Google's servers
- **Google Calendar API load** (varies throughout the day)
- **Event deletion** (if clearing existing events)
- **Time of day** (Google services faster during off-peak hours)

### **Rate Limiting Protection**

The system includes intelligent delays to prevent "too many operations" errors:

#### **During Event Creation:**
- **Pause every 25 events**: 1-second delay
- **Progress logging**: "Created X events so far..."
- **Individual error isolation**: One failed event won't stop the process

#### **During Event Deletion (if clearing existing):**
- **Pause every 10 deletions**: 0.5-second delay
- **Cooldown period**: 2-second wait before creating new events
- **Batch processing**: Handles large numbers of existing events safely

### **What You'll Experience**

#### **Normal Operation:**
- **Progress appears steady** with occasional brief pauses
- **Console logs show** event creation progress
- **No error messages** during successful operation
- **Final success dialog** with count of events created

#### **If You See Long Delays:**
- **This is normal** for larger schedules (150+ events)
- **Don't interrupt the process** - let the rate limiting work
- **Watch for progress messages** in the console
- **Success dialog** will appear when complete

### **Troubleshooting Long Runtimes**

#### **If Export Takes Much Longer Than Expected:**

**"Script running over 5 minutes for 100 events"**
- **Check internet connection** - slow connection affects API calls
- **Try during off-peak hours** (early morning or late evening)
- **Reduce schedule size** to test with smaller batches first

**"Script seems frozen with no progress"**
- **Check browser console** for error messages
- **Don't refresh** - this will interrupt the process
- **Wait for timeout** - script will eventually show error or success

**"Getting 'too many operations' errors despite rate limiting"**
- **Wait 10-15 minutes** for Google's rate limit to reset
- **Try smaller batch** of events first
- **Use "Add to existing"** instead of clearing all events

### **Best Practices for Large Schedules**

#### **For 200+ Events:**
1. **Schedule during off-peak hours** (early morning recommended)
2. **Test with smaller subset first** (4-6 weeks)
3. **Ensure stable internet connection**
4. **Don't use computer for other intensive tasks** during export
5. **Be patient** - the system is designed for reliability over speed

#### **For Very Large Schedules (400+ Events):**
- **Consider breaking into quarters** (export 13 weeks at a time)
- **Use incremental approach** rather than clearing all events
- **Plan for 5-10 minute processing time**
- **Monitor system resources** during processing

### **Performance Optimization Tips**

#### **To Reduce Runtime:**
- **Generate shortened URLs first** (avoids real-time URL shortening)
- **Use "Add to existing events"** instead of clearing/recreating
- **Export during Google's off-peak hours**
- **Ensure good internet connection**

#### **To Improve Reliability:**
- **Start with validation** to catch configuration errors early
- **Test with small schedule first** before full year export
- **Don't interrupt the process** once it starts
- **Let rate limiting delays work** - they prevent failures

### **Success Indicators**

#### **Your Export is Working Well When:**
✅ **Steady progress** with occasional planned pauses  
✅ **Console shows** "Created X events so far..." messages  
✅ **No error alerts** during processing  
✅ **Final success dialog** appears with event count  
✅ **Calendar events** appear with rich descriptions and working links  

#### **Example Success Message:**
```
✅ Created 135 calendar events with enhanced information!

📅 Calendar: "Deacon Visitation Schedule"
🕐 Default time: 2:00 PM - 3:00 PM
🔗 Includes: Breeze profiles and visit notes links
📞 Contact info: Phone numbers and addresses
📝 Instructions: Custom visit coordination text

Each event title: "[Deacon] visits [Household]"
View and modify these events in Google Calendar.
```

---

**Remember: The system prioritizes reliability over speed. The built-in delays ensure your calendar export completes successfully, even with large schedules!** 🎯

[Return to Setup](SETUP.md)
