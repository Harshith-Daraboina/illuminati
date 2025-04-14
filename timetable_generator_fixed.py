import pandas as pd
import random
from datetime import datetime, time, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter
import os
import logging
import re
import time as time_module

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("timetable_generation.log"),
        logging.StreamHandler()
    ]
)

# Constants
DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
START_TIME = time(9, 0)
END_TIME = time(18, 30)
LECTURE_DURATION = 3  # 1.5 hours = 3 slots (30 mins each)
LAB_DURATION = 4      # 2 hours = 4 slots (30 mins each)
TUTORIAL_DURATION = 2  # 1 hour = 2 slots (30 mins each)
MAX_SCHEDULING_ATTEMPTS = 5000
MAX_RUNTIME = 300  # 5 minutes maximum runtime for scheduling attempts

# Define patterns for identifying elective courses
ELECTIVE_PATTERNS = [
    r'B\d+',  # Matches B1, B2, B3, etc.
    r'E\d+',  # Matches E1, E2, E3, etc.
    r'ELECTIVE',  # Matches "ELECTIVE" 
    r'OE\d*',  # Matches OE, OE1, OE2, etc.
    r'PE\d*'   # Matches PE, PE1, PE2, etc.
]

def is_elective_course(course_code):
    """Check if a course is an elective based on its code"""
    for pattern in ELECTIVE_PATTERNS:
        if re.search(pattern, str(course_code), re.IGNORECASE):
            return True
    return False

def validate_classroom(classroom):
    """Validate and normalize classroom assignment"""
    classroom = str(classroom).strip()
    if "Will be scheduled" in classroom or not classroom:
        return None
    return classroom

def validate_course_data(course):
    """Validate required course data fields"""
    required_fields = ['L', 'T', 'P', 'Faculty', 'Classroom', 'Course Code', 'Course Name']
    for field in required_fields:
        if field not in course or pd.isna(course.get(field)):
            return False
    return True

def generate_time_slots():
    """Generate time slots for the day"""
    slots = []
    current_time = datetime.combine(datetime.today(), START_TIME)
    end_time = datetime.combine(datetime.today(), END_TIME)
    
    while current_time < end_time:
        current = current_time.time()
        next_time = (current_time + timedelta(minutes=30)).time()
        
        slots.append((current, next_time))
        current_time = current_time + timedelta(minutes=30)
    
    return slots

def load_and_clean_data():
    """Load and clean course data"""
    try:
        # Try excel file first
        df = pd.read_excel('combined.xlsx', sheet_name='Sheet1')
        logging.info("Successfully loaded data from combined.xlsx")
    except (FileNotFoundError, Exception) as e:
        try:
            df = pd.read_csv('combined.csv')
            logging.info("Successfully loaded data from combined.csv")
        except FileNotFoundError:
            logging.error("Error: Neither 'combined.xlsx' nor 'combined.csv' found")
            exit()
        except Exception as e:
            logging.error(f"Error loading data: {str(e)}")
            exit()
    
    # Clean up data
    for col in ['L', 'T', 'P', 'S', 'C']:
        if col in df.columns:
            df[col] = df[col].fillna(0).astype(int)
    
    # Ensure required columns exist
    required_cols = ['Department', 'Semester', 'Course Code', 'Course Name', 'L', 'T', 'P', 'Faculty', 'Classroom']
    for col in required_cols:
        if col not in df.columns:
            logging.error(f"Required column '{col}' not found")
            exit()
    
    # Fix missing course codes
    for idx, row in df.iterrows():
        if row['Course Code'] == "-" or pd.isna(row['Course Code']):
            dept = str(row['Department']).strip()
            if pd.isna(dept) or dept == "":
                continue
            semester = str(row['Semester']).strip()
            course_name = str(row['Course Name']).strip()
            if course_name:
                words = course_name.split()
                if words:
                    code_part = ''.join([word[0] for word in words if word])[:3].upper()
                    new_code = f"{dept[:2].upper()}{semester}{code_part}"
                    df.at[idx, 'Course Code'] = new_code
    
    # Fix classroom assignments
    for idx, row in df.iterrows():
        classroom = str(row['Classroom']).strip()
        if "Will be scheduled" in classroom or pd.isna(classroom) or classroom == "":
            dept = str(row['Department']).strip()
            semester = str(row['Semester']).strip()
            df.at[idx, 'Classroom'] = f"TEMP_{dept}_{semester}"
    
    # Identify elective courses
    df['is_elective'] = df['Course Code'].apply(is_elective_course)
    df['elective_group'] = None
    
    for pattern in ELECTIVE_PATTERNS:
        mask = df['Course Code'].astype(str).str.contains(pattern, case=False, regex=True)
        for idx in df[mask].index:
            code = df.at[idx, 'Course Code']
            matches = re.search(pattern, code, re.IGNORECASE)
            if matches:
                df.at[idx, 'elective_group'] = matches.group(0).upper()
    
    # Clean data
    df = df.dropna(how='all')
    df = df[(df['Department'].notna()) & (df['Department'] != "") & 
            (df['Semester'].notna()) & (df['Semester'] != "")]
    
    return df

def is_morning_break(slot):
    """Check if time slot is morning break"""
    start, end = slot
    return (time(10, 30) <= start < time(11, 0))

def is_lunch_time(slot):
    """Check if time slot is lunch time"""
    start, end = slot
    return (time(12, 30) <= start < time(14, 30))

def check_scheduling_possibility(faculty, classroom, day, start_slot, duration, professor_schedule, classroom_schedule, timetable, TIME_SLOTS):
    """Check if slots are available for scheduling"""
    faculty_flexible = '/' in str(faculty) or ',' in str(faculty)
    classroom_flexible = str(classroom).startswith('TEMP_') or "Will be scheduled" in str(classroom)
    
    for i in range(duration):
        current_slot = start_slot + i
        if current_slot >= len(TIME_SLOTS):
            return False
            
        if timetable[day][current_slot]['type'] is not None:
            return False
            
        if is_morning_break(TIME_SLOTS[current_slot]):
            return False
            
        if not faculty_flexible and faculty in professor_schedule:
            if current_slot in professor_schedule[faculty][day]:
                return False
                
        if not classroom_flexible and classroom in classroom_schedule:
            if current_slot in classroom_schedule[classroom][day]:
                return False
    
    return True

def update_schedule(faculty, classroom, day, start_slot, duration, session_type, code, name, professor_schedule, classroom_schedule, timetable):
    """Update all schedules with new session"""
    faculty_list = [faculty]
    if '/' in str(faculty) or ',' in str(faculty):
        faculty_list = [f.strip() for f in re.split(r'[/,]', str(faculty))]
    
    for i in range(duration):
        for single_faculty in faculty_list:
            if single_faculty not in professor_schedule:
                professor_schedule[single_faculty] = {day: set() for day in range(len(DAYS))}
            professor_schedule[single_faculty][day].add(start_slot+i)
        
        if classroom not in classroom_schedule:
            classroom_schedule[classroom] = {day: set() for day in range(len(DAYS))}
        classroom_schedule[classroom][day].add(start_slot+i)
        
        timetable[day][start_slot+i]['type'] = session_type
        timetable[day][start_slot+i]['code'] = code if i == 0 else ''
        timetable[day][start_slot+i]['name'] = name if i == 0 else ''
        timetable[day][start_slot+i]['faculty'] = faculty if i == 0 else ''
        timetable[day][start_slot+i]['classroom'] = classroom if i == 0 else ''
        timetable[day][start_slot+i]['duration'] = duration if i == 0 else 0
        timetable[day][start_slot+i]['is_first'] = (i == 0)
        timetable[day][start_slot+i]['position'] = i

def schedule_session(department, semester, course, session_type, professor_schedule, classroom_schedule, timetable, TIME_SLOTS, summary_ws, attempt_limit, elective_schedule=None):
    """Schedule a specific session"""
    if not validate_course_data(course):
        logging.warning(f"Invalid course data: {course.get('Course Code', 'Unknown')}")
        return False

    code = str(course['Course Code'])
    name = str(course['Course Name'])
    faculty = str(course['Faculty'])
    classroom = validate_classroom(course['Classroom'])
    
    if classroom is None:
        classroom = f"TEMP_{department}_{semester}"
    
    is_elective = course.get('is_elective', False)
    elective_group = course.get('elective_group', None)
    
    if session_type == 'LAB':
        duration = LAB_DURATION
    elif 'LEC' in session_type:
        duration = LECTURE_DURATION
    else:
        duration = TUTORIAL_DURATION
    
    scheduled = False
    attempts = 0
    start_time = time_module.time()
    
    if is_elective and elective_group and elective_group in elective_schedule:
        day, start_slot = elective_schedule[elective_group][session_type]
        
        if check_scheduling_possibility(faculty, classroom, day, start_slot, duration, 
                                     professor_schedule, classroom_schedule, timetable, TIME_SLOTS):
            update_schedule(faculty, classroom, day, start_slot, duration, session_type, 
                         code, name, professor_schedule, classroom_schedule, timetable)
            scheduled = True
            summary_ws.append([department, semester, code, name, session_type, faculty, classroom, "Scheduled", 
                              f"{DAYS[day]} {TIME_SLOTS[start_slot][0].strftime('%H:%M')}", "ELECTIVE"])
    else:
        day_load = {day: sum(1 for slot in timetable[day].values() if slot['type'] is not None) 
                   for day in range(len(DAYS))}
        sorted_days = sorted(day_load.keys(), key=lambda d: day_load[d])
        
        for day in sorted_days:
            if scheduled:
                break
                
            for start_slot in range(len(TIME_SLOTS)-duration+1):
                if check_scheduling_possibility(faculty, classroom, day, start_slot, duration, 
                                             professor_schedule, classroom_schedule, timetable, TIME_SLOTS):
                    update_schedule(faculty, classroom, day, start_slot, duration, session_type, 
                                 code, name, professor_schedule, classroom_schedule, timetable)
                    scheduled = True
                    
                    if is_elective and elective_group and elective_schedule is not None:
                        if elective_group not in elective_schedule:
                            elective_schedule[elective_group] = {}
                        elective_schedule[elective_group][session_type] = (day, start_slot)
                            
                    summary_ws.append([department, semester, code, name, session_type, faculty, classroom, "Scheduled", 
                                      f"{DAYS[day]} {TIME_SLOTS[start_slot][0].strftime('%H:%M')}", 
                                      "ELECTIVE" if is_elective else ""])
                    break
        
        while not scheduled and attempts < attempt_limit:
            if time_module.time() - start_time > MAX_RUNTIME:
                logging.warning("Max runtime exceeded for scheduling attempt")
                break
                
            day = random.randint(0, len(DAYS)-1)
            start_slot = random.randint(0, len(TIME_SLOTS)-duration)
            
            if check_scheduling_possibility(faculty, classroom, day, start_slot, duration, 
                                         professor_schedule, classroom_schedule, timetable, TIME_SLOTS):
                update_schedule(faculty, classroom, day, start_slot, duration, session_type, 
                             code, name, professor_schedule, classroom_schedule, timetable)
                scheduled = True
                
                if is_elective and elective_group and elective_schedule is not None:
                    if elective_group not in elective_schedule:
                        elective_schedule[elective_group] = {}
                    elective_schedule[elective_group][session_type] = (day, start_slot)
                
                summary_ws.append([department, semester, code, name, session_type, faculty, classroom, "Scheduled", 
                                  f"{DAYS[day]} {TIME_SLOTS[start_slot][0].strftime('%H:%M')}",
                                  "ELECTIVE" if is_elective else ""])
            attempts += 1
    
    if not scheduled:
        logging.warning(f"Failed to schedule {session_type} for {code}: {name}")
        summary_ws.append([department, semester, code, name, session_type, faculty, classroom, "Failed", "N/A", 
                          "ELECTIVE" if is_elective else ""])
        
    return scheduled

def handle_lectures(department, semester, course, professor_schedule, classroom_schedule, timetable, TIME_SLOTS, summary_ws, attempt_limit, elective_schedule=None):
    """Handle scheduling lectures based on L value"""
    l = int(course['L'])
    total_scheduled = 0
    failed = 0
    
    if l == 3:
        for lecture_idx in range(2):
            current_
