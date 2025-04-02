import pandas as pd
import random
from datetime import datetime, time, timedelta

# Constants
DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
START_TIME = time(9, 0)
END_TIME = time(18, 30)
TIME_SLOT_DURATION = 30  # minutes

def generate_time_slots():
    slots = []
    current_time = datetime.combine(datetime.today(), START_TIME)
    end_time = datetime.combine(datetime.today(), END_TIME)
    
    while current_time < end_time:
        current = current_time.time()
        # Skip break times (10:30-11:00 and 12:30-14:30)
        is_morning_break = time(10, 30) <= current < time(11, 0)
        is_lunch_break = time(12, 30) <= current < time(14, 30)
        if not is_morning_break and not is_lunch_break:
            next_time = current_time + timedelta(minutes=TIME_SLOT_DURATION)
            slots.append((current, next_time.time()))
        current_time += timedelta(minutes=TIME_SLOT_DURATION)
    return slots

# Load data from Excel
try:
    df = pd.read_excel('combined.xlsx', sheet_name='Sheet1')
except FileNotFoundError:
    print("Error: File 'combined.xlsx' not found in the current directory")
    exit()

def generate_timetable(department, semester):
    courses = df[(df['Department'] == department) & (df['Semester'] == semester)].copy()
    
    if courses.empty:
        print(f"No courses found for {department} - Semester {semester}")
        return
    
    TIME_SLOTS = generate_time_slots()
    timetable = {day: {slot: [] for slot in range(len(TIME_SLOTS))} for day in range(len(DAYS))}
    
    for _, course in courses.iterrows():
        code = str(course['Course Code'])
        name = str(course['Course Name'])
        faculty = str(course['Faculty'])
        classroom = str(course['Classroom'])
        l = int(course['L']) if pd.notna(course['L']) else 0
        t = int(course['T']) if pd.notna(course['T']) else 0
        
        # Schedule lectures (1 hour)
        for _ in range(l):
            scheduled = False
            attempts = 0
            while not scheduled and attempts < 100:
                day = random.randint(0, len(DAYS)-1)
                slot = random.randint(0, len(TIME_SLOTS)-2)
                
                if (not any(c[0] == code for c in timetable[day][slot]) and 
                    not any(c[0] == code for c in timetable[day][slot+1])):
                    timetable[day][slot].append((code, name, 'Lecture', faculty, classroom))
                    timetable[day][slot+1].append(('', '', 'cont.', '', ''))
                    scheduled = True
                attempts += 1
        
        # Schedule tutorials (30 mins)
        for _ in range(t):
            scheduled = False
            attempts = 0
            while not scheduled and attempts < 100:
                day = random.randint(0, len(DAYS)-1)
                slot = random.randint(0, len(TIME_SLOTS)-1)
                
                if not any(c[0] == code for c in timetable[day][slot]):
                    timetable[day][slot].append((code, name, 'Tutorial', faculty, classroom))
                    scheduled = True
                attempts += 1
    
    # Print timetable
    print(f"\nTimetable for {department} - Semester {semester}")
    print("\n" + "="*(15 + 12*len(TIME_SLOTS)))
    print(f"{'Day':<15}", end="")
    for slot in TIME_SLOTS:
        print(f"{slot[0].strftime('%H:%M'):<12}", end="")
    print()
    
    for day_idx, day in enumerate(DAYS):
        print(f"{day:<15}", end="")
        for slot_idx in range(len(TIME_SLOTS)):
            if timetable[day_idx][slot_idx]:
                course = timetable[day_idx][slot_idx][0]
                print(f"{course[0]}:{course[2][0]:<11}" if course[0] else f"{'':<12}", end="")
            else:
                print(f"{'':<12}", end="")
        print()

def main():
    departments = df['Department'].unique()
    for department in departments:
        semesters = df[df['Department'] == department]['Semester'].unique()
        for semester in semesters:
            generate_timetable(department, semester)

if __name__ == "__main__":
    main()