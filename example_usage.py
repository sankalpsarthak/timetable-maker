#!/usr/bin/env python3
"""
Example usage of TimeTableMaker with different scenarios
"""

from timetable_maker import TimeTableMaker

def example_1_student_schedule():
    """Create a timetable for a student"""
    print("\n📚 Example 1: Student Schedule")
    print("-" * 50)
    
    maker = TimeTableMaker("student_timetable.xlsx")
    
    # Daily schedule for student
    student_activities = {
        "06:00": "Wake up & Shower",
        "07:00": "Breakfast",
        "08:00": "Classes",
        "12:30": "Lunch",
        "13:30": "Classes",
        "16:00": "Sports/Club",
        "17:30": "Study Time",
        "19:00": "Dinner",
        "20:00": "Assignments",
        "22:00": "Sleep"
    }
    
    maker.create_daily_timetable(activities=student_activities, time_interval=30)
    maker.create_weekly_timetable()
    maker.create_planner(period="monthly")
    maker.create_habit_tracker()
    maker.save()
    
    print("✅ Created: student_timetable.xlsx")

def example_2_professional_schedule():
    """Create a timetable for a professional"""
    print("\n💼 Example 2: Professional Schedule")
    print("-" * 50)
    
    maker = TimeTableMaker("professional_timetable.xlsx")
    
    # Daily schedule for professional
    professional_activities = {
        "06:00": "Morning Walk/Gym",
        "07:00": "Breakfast",
        "08:00": "Commute",
        "09:00": "Office Work - Deep Focus",
        "12:00": "Lunch Meeting",
        "13:00": "Afternoon Tasks",
        "15:00": "Coffee & Meetings",
        "16:00": "Project Work",
        "18:00": "Leave Office",
        "19:00": "Dinner",
        "20:00": "Personal Development",
        "22:00": "Sleep"
    }
    
    maker.create_daily_timetable(activities=professional_activities, time_interval=30)
    maker.create_weekly_timetable()
    maker.create_planner(period="monthly")
    maker.create_planner(period="12monthly")
    maker.create_habit_tracker()
    maker.save()
    
    print("✅ Created: professional_timetable.xlsx")

def example_3_fitness_routine():
    """Create a timetable for fitness training"""
    print("\n🏋️ Example 3: Fitness Training Schedule")
    print("-" * 50)
    
    maker = TimeTableMaker("fitness_timetable.xlsx")
    
    # Daily fitness activities
    fitness_activities = {
        "05:00": "Wake up",
        "05:30": "Warm-up Stretching",
        "06:00": "Cardio Session",
        "07:00": "Breakfast",
        "08:00": "Work/Study",
        "12:00": "Lunch - Protein Focus",
        "14:00": "Work/Study",
        "17:00": "Strength Training",
        "18:30": "Cool-down & Stretching",
        "19:00": "Dinner",
        "20:00": "Rest & Recovery",
        "22:00": "Sleep"
    }
    
    maker.create_daily_timetable(activities=fitness_activities, time_interval=60)
    maker.create_weekly_timetable()
    maker.create_planner(period="weekly")
    maker.create_habit_tracker()
    maker.save()
    
    print("✅ Created: fitness_timetable.xlsx")

def example_4_multi_planner():
    """Create all planners in one file"""
    print("\n🎯 Example 4: Comprehensive Planner")
    print("-" * 50)
    
    maker = TimeTableMaker("comprehensive_planner.xlsx")
    
    # Create all planning options
    maker.create_daily_timetable(time_interval=30)
    maker.create_weekly_timetable()
    maker.create_planner(period="daily")
    maker.create_planner(period="weekly")
    maker.create_planner(period="monthly")
    maker.create_planner(period="3monthly")
    maker.create_planner(period="6monthly")
    maker.create_planner(period="9monthly")
    maker.create_planner(period="12monthly")
    maker.create_habit_tracker()
    maker.save()
    
    print("✅ Created: comprehensive_planner.xlsx")
    print("📊 10 different planning sheets created!")

if __name__ == "__main__":
    print("\n" + "="*50)
    print("🎨 TimeTableMaker - Example Usage")
    print("="*50)
    
    example_1_student_schedule()
    example_2_professional_schedule()
    example_3_fitness_routine()
    example_4_multi_planner()
    
    print("\n" + "="*50)
    print("✅ All examples completed!")
    print("="*50)