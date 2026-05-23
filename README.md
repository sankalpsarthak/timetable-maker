# Custom Timetable Maker

A comprehensive Python tool to create professional timetables, planners, and habit trackers in Excel.

## 🚀 Features

### Timetables
- **Daily Timetable** - Hour-by-hour schedule with customizable time intervals
- **Weekly Timetable** - Full week planning with daily activity slots

### Planners (7 Duration Options)
- ✅ Daily Planner
- ✅ Weekly Planner  
- ✅ Monthly Planner
- ✅ 3-Monthly (Quarterly) Planner
- ✅ 6-Monthly Planner
- ✅ 9-Monthly Planner
- ✅ 12-Monthly (Yearly) Planner

### Tracking
- **Habit Tracker** - Track habits for all days of the current month with checkbox interface

## 📦 Installation

```bash
pip install -r requirements.txt
```

### Dependencies
- `openpyxl==3.10.1` - Excel file manipulation
- `calendars==0.6.0` - Calendar utilities

## 🎯 Quick Start

### Option 1: Generate Everything at Once

```bash
python quick_start.py
```

This creates `my_first_timetable.xlsx` with all features pre-configured.

### Option 2: Custom Usage

```python
from timetable_maker import TimeTableMaker

# Create instance
maker = TimeTableMaker("my_timetable.xlsx")

# Define your daily activities
activities = {
    "09:00": "Work Session",
    "12:00": "Lunch Break",
    "14:00": "Meeting",
    "17:00": "Personal Project"
}

# Add components
maker.create_daily_timetable(activities=activities, time_interval=30)
maker.create_weekly_timetable()
maker.create_planner(period="monthly")
maker.create_habit_tracker()

# Save
maker.save()
```

## 📚 Usage Examples

Run pre-built examples:

```bash
python example_usage.py
```

This creates 4 different timetables:
1. **student_timetable.xlsx** - Student schedule
2. **professional_timetable.xlsx** - Professional schedule
3. **fitness_timetable.xlsx** - Fitness routine
4. **comprehensive_planner.xlsx** - All features combined

## 🔧 API Reference

### TimeTableMaker Class

#### Constructor
```python
TimeTableMaker(filename="timetable.xlsx")
```

#### Methods

##### `create_daily_timetable(activities=None, time_interval=30)`
Create a daily schedule sheet.

**Parameters:**
- `activities` (dict): Time keys with activity values. Example: `{"09:00": "Work"}`
- `time_interval` (int): Interval in minutes (15, 30, or 60). Default: 30

```python
maker.create_daily_timetable(
    activities={"09:00": "Meeting", "14:00": "Work"},
    time_interval=30
)
```

##### `create_weekly_timetable(activities=None)`
Create a weekly schedule sheet (7 days × hourly slots).

**Parameters:**
- `activities` (dict): Optional pre-filled activities

```python
maker.create_weekly_timetable()
```

##### `create_planner(period="monthly")`
Create a planner for different time durations.

**Parameters:**
- `period` (str): One of:
  - `"daily"` - Daily planner
  - `"weekly"` - Weekly planner
  - `"monthly"` - Monthly planner (30 days)
  - `"3monthly"` - Quarterly planner (90 days)
  - `"6monthly"` - Semi-annual planner (180 days)
  - `"9monthly"` - 9-month planner (270 days)
  - `"12monthly"` - Yearly planner (365 days)

```python
maker.create_planner(period="monthly")
maker.create_planner(period="12monthly")
```

##### `create_habit_tracker()`
Create a habit tracker for the current month.

Includes 8 pre-defined habits:
- Exercise
- Reading
- Meditation
- Healthy Eating
- Sleep 8 Hours
- Work Goals
- Family Time
- Hydration

Plus 3 blank rows for custom habits.

```python
maker.create_habit_tracker()
```

##### `save()`
Save the workbook to the specified filename.

```python
maker.save()
```

##### `save_as(filename)`
Save the workbook with a different filename.

```python
maker.save_as("new_timetable.xlsx")
```

## 🎨 Styling Features

- **Professional color scheme** - Blue headers, light backgrounds
- **Auto-adjusted columns** - Optimal width for readability
- **Cell borders** - Clean grid layout
- **Wrapped text** - Multi-line content support
- **Center alignment** - Professional appearance
- **Color-coded sections** - Easy visual distinction

## 📊 Output Files

Each Excel workbook contains:

| Sheet | Content | Rows |
|-------|---------|------|
| Daily Timetable | Hourly schedule | 20-30 |
| Weekly Timetable | 7 days × hourly | 20-30 |
| Daily Planner | Tasks by time | 15+ |
| Weekly Planner | 5 focus areas | 8 |
| Monthly Planner | 30 days | 32 |
| 3-Monthly Planner | Quarterly goals | 20 |
| 6-Monthly Planner | Semi-annual goals | 35 |
| 9-Monthly Planner | 9-month goals | 50 |
| Yearly Planner | Quarterly breakdown | 20 |
| Habit Tracker | All month days | 12 |

## 💡 Use Cases

### Student
- Track class schedule
- Plan study sessions
- Monitor habits (exercise, sleep, reading)
- Organize projects by deadline

### Professional
- Daily work schedule
- Weekly meetings and deliverables
- Long-term project planning
- Habit tracking for wellness

### Fitness Enthusiast
- Workout schedule
- Weekly training plan
- 12-month fitness goals
- Daily habit tracking

### Freelancer
- Client schedules
- Project timelines
- Milestone planning
- Productivity tracking

## 🔄 Workflow Example

```python
from timetable_maker import TimeTableMaker

# Create
maker = TimeTableMaker("my_schedule.xlsx")

# Define routine
my_routine = {
    "06:30": "Morning Exercise",
    "08:00": "Work",
    "12:00": "Lunch",
    "13:00": "Work",
    "17:30": "Personal Time",
    "21:00": "Sleep Prep"
}

# Add all components
maker.create_daily_timetable(activities=my_routine)
maker.create_weekly_timetable()
maker.create_planner(period="monthly")
maker.create_planner(period="12monthly")
maker.create_habit_tracker()

# Save
maker.save()
print("✅ Your timetable is ready!")
```

## 🐛 Troubleshooting

### Excel shows strange formatting
- Ensure openpyxl is up to date: `pip install --upgrade openpyxl`

### File doesn't open
- Check that the filename has `.xlsx` extension
- Ensure you have Microsoft Excel or compatible software

### Memory issues with large files
- Create separate files for different time periods
- Use the `save_as()` method to manage multiple files

## 📝 License

MIT License - Use freely for personal and commercial projects

## 🤝 Contributing

Contributions welcome! Feel free to:
- Report bugs
- Suggest features
- Submit pull requests

## 📧 Support

For issues or questions, please open a GitHub issue.

---

**Made with ❤️ for better time management**