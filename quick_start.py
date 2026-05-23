#!/usr/bin/env python3
"""
Quick start script to create a complete timetable with all features
"""

from timetable_maker import TimeTableMaker

def main():
    print("🎯 Creating Custom Timetable Maker...")
    print("-" * 50)
    
    # Create maker instance
    maker = TimeTableMaker("my_first_timetable.xlsx")
    
    # Sample daily activities
    print("✓ Creating Daily Timetable...")
    daily_activities = {
        "06:00": "Wake up & Exercise",
        "07:00": "Breakfast",
        "08:00": "Work/Study Session 1",
        "10:00": "Coffee Break",
        "10:30": "Work/Study Session 2",
        "12:30": "Lunch",
        "14:00": "Work/Study Session 3",
        "17:00": "Break & Snack",
        "18:00": "Personal Projects/Hobbies",
        "20:00": "Dinner",
        "21:00": "Relaxation/Reading",
        "22:30": "Sleep Preparation"
    }
    maker.create_daily_timetable(activities=daily_activities, time_interval=30)
    
    # Create weekly timetable
    print("✓ Creating Weekly Timetable...")
    maker.create_weekly_timetable()
    
    # Create planners for different periods
    print("✓ Creating Daily Planner...")
    maker.create_planner(period="daily")
    
    print("✓ Creating Weekly Planner...")
    maker.create_planner(period="weekly")
    
    print("✓ Creating Monthly Planner...")
    maker.create_planner(period="monthly")
    
    print("✓ Creating 3-Monthly Planner...")
    maker.create_planner(period="3monthly")
    
    print("✓ Creating 6-Monthly Planner...")
    maker.create_planner(period="6monthly")
    
    print("✓ Creating 9-Monthly Planner...")
    maker.create_planner(period="9monthly")
    
    print("✓ Creating Yearly Planner...")
    maker.create_planner(period="12monthly")
    
    # Create habit tracker
    print("✓ Creating Habit Tracker...")
    maker.create_habit_tracker()
    
    # Save all
    maker.save()
    
    print("-" * 50)
    print("✅ Complete! File: my_first_timetable.xlsx")
    print("📊 Created 10 sheets:")
    print("   1. Daily Timetable")
    print("   2. Weekly Timetable")
    print("   3. Daily Planner")
    print("   4. Weekly Planner")
    print("   5. Monthly Planner")
    print("   6. 3-Monthly Planner")
    print("   7. 6-Monthly Planner")
    print("   8. 9-Monthly Planner")
    print("   9. Yearly Planner")
    print("   10. Habit Tracker")

if __name__ == "__main__":
    main()