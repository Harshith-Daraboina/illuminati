import pandas as pd
import random
from datetime import datetime, time, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter
import os
import logging

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
MAX_SCHEDULING_ATTEMPTS = 5000  # Increased from 2000

def generate_time_slots():
    """Generate time slots for the day"""
    slots = []
    current_time = datetime.combine(datetime.today(), START_TIME)
    end_time = datetime.combine(datetime.today(), END_TIME)
    
    while current_time < end_time:
        current = current_time.time()
        next_time = (current_time + timedelta(minutes=30)).time()
        
        # Keep all time slots but we'll mark break times later
        slots.append((current, next_time))
        current_time = current_time + timedelta(minutes=30)
    
    return slots

def load_and_clean_data():
    """Load course data from Excel or CSV file and clean it"""
    try:
        # Try excel file first
        df = pd.read_excel('combined.xlsx', sheet_name='Sheet1')
        logging.info("Successfully loaded data from combined.xlsx")
    except (FileNotFoundError, Exception) as e:
        # Fall back to CSV if excel file not found or other Excel-related error
        try:
            df = pd.read_csv('combined.csv')
            logging.info("Successfully loaded data from combined.csv")
        except FileNotFoundError:
            logging.error("Error: Neither 'combined.xlsx' nor 'combined.csv' found in the current directory")
            exit()
        except Exception as e:
            logging.error(f"Error loading CSV: {str(e)}")
            exit()
    
    # Clean up data: Replace NaN with 0 for numerical columns
    for col in ['L', 'T', 'P', 'S', 'C']:
        if col in df.columns:
            df[col] = df[col].fillna(0).astype(int)
    
    # Ensure required columns exist
    required_cols = ['Department', 'Semester', 'Course Code', 'Course Name', 'L', 'T', 'P', 'Faculty', 'Classroom']
    for col in required_cols:
        if col not in df.columns:
            logging.error(f"Required column '{col}' not found in data")
            exit()
    
    # Fix missing course codes: Replace "-" with a generated code
    for idx, row in df.iterrows():
        if row['Course Code'] == "-" or pd.isna(row['Course Code']):
            # Generate a course code based on department and a counter
            dept = str(row['Department']).strip()
            if pd.isna(dept) or dept == "":
                continue
            semester = str(row['Semester']).strip()
            course_name = str(row['Course Name']).strip()
            if course_name:
                # Create a code from first letters of course name words
                words = course_name.split()
                if words:
                    code_part = ''.join([word[0] for word in words if word])[:3].upper()
                    new_code = f"{dept[:2].upper()}{semester}{code_part}"
                    df.at[idx, 'Course Code'] = new_code
                    logging.info(f"Generated course code {new_code} for {course_name}")
    
    # Fix classroom assignments that are "Will be scheduled Post MidSem"
    for idx, row in df.iterrows():
        classroom = str(row['Classroom']).strip()
        if "Will be scheduled" in classroom or pd.isna(classroom) or classroom == "":
            # Assign a temporary classroom based on department and semester
            dept = str(row['Department']).strip()
            semester = str(row['Semester']).strip()
            df.at[idx, 'Classroom'] = f"TBD_{dept}_{semester}"
            logging.info(f"Assigned temporary classroom TBD_{dept}_{semester} for {row['Course Name']}")
    
    # Remove completely empty rows
    df = df.dropna(how='all')
    
    # Remove rows with no department or semester
    df = df[(df['Department'].notna()) & (df['Department'] != "") & 
            (df['Semester'].notna()) & (df['Semester'] != "")]
    
    return df

def is_break_time(slot):
    """Check if a time slot falls within break times"""
    start, end = slot
    # Morning break: 10:30-11:00
    morning_break = (time(10, 30) <= start < time(11, 0))
    
    # Lunch break: 12:30-14:30 - any two consecutive half hours
    lunch_break = False
    if time(12, 30) <= start < time(14, 0):  # Check if this slot starts a potential lunch break
        # Check if next slot is also marked as lunch time
        next_start = datetime.combine(datetime.today(), start) + timedelta(minutes=30)
        next_start_time = next_start.time()
        if time(12, 30) <= next_start_time < time(14, 30):
            lunch_break = True
    
    return morning_break or lunch_break

def check_scheduling_possibility(faculty, classroom, day, start_slot, duration, professor_schedule, classroom_schedule, timetable, TIME_SLOTS):
    """Check if the given slots are available for scheduling"""
    # 1. If faculty is marked multiple (with slashes or commas), consider it as available
    faculty_flexible = '/' in str(faculty) or ',' in str(faculty)
    
    # 2. If classroom is TBD or marked for post midsem, consider it as available
    classroom_flexible = str(classroom).startswith('TBD_') or "Will be scheduled" in str(classroom)
    
    # 3. Check if all required slots are free and not in break time
    slots_free = True
    for i in range(duration):
        current_slot = start_slot + i
        if current_slot >= len(TIME_SLOTS):  # Check if we're exceeding time slots
            return False
            
        # Check if slot is already scheduled
        if timetable[day][current_slot]['type'] is not None:
            return False
            
        # Check if slot is during break
        if is_break_time(TIME_SLOTS[current_slot]):
            return False
            
        # Check faculty availability (only if not flexible)
        if not faculty_flexible and faculty in professor_schedule:
            if current_slot in professor_schedule[faculty][day]:
                return False
                
        # Check classroom availability (only if not flexible)
        if not classroom_flexible and classroom in classroom_schedule:
            if current_slot in classroom_schedule[classroom][day]:
                return False
    
    return True

def update_schedule(faculty, classroom, day, start_slot, duration, session_type, code, name, professor_schedule, classroom_schedule, timetable):
    """Update all schedules with the new session"""
    # Handle multiple faculty members (marked with slashes or commas)
    faculty_list = [faculty]
    if '/' in str(faculty) or ',' in str(faculty):
        if '/' in str(faculty):
            faculty_list = [f.strip() for f in str(faculty).split('/')]
        else:
            faculty_list = [f.strip() for f in str(faculty).split(',')]
    
    # Update schedules
    for i in range(duration):
        # For each faculty in the list, mark them as busy
        for single_faculty in faculty_list:
            if single_faculty not in professor_schedule:
                professor_schedule[single_faculty] = {day: set() for day in range(len(DAYS))}
            professor_schedule[single_faculty][day].add(start_slot+i)
        
        # Mark classroom as busy
        if classroom not in classroom_schedule:
            classroom_schedule[classroom] = {day: set() for day in range(len(DAYS))}
        classroom_schedule[classroom][day].add(start_slot+i)
        
        # Update timetable entry
        timetable[day][start_slot+i]['type'] = session_type
        timetable[day][start_slot+i]['code'] = code if i == 0 else ''
        timetable[day][start_slot+i]['name'] = name if i == 0 else ''
        timetable[day][start_slot+i]['faculty'] = faculty if i == 0 else ''
        timetable[day][start_slot+i]['classroom'] = classroom if i == 0 else ''
        timetable[day][start_slot+i]['duration'] = duration if i == 0 else 0
        timetable[day][start_slot+i]['is_first'] = (i == 0)
        timetable[day][start_slot+i]['position'] = i

def schedule_session(department, semester, course, session_type, professor_schedule, classroom_schedule, timetable, TIME_SLOTS, summary_ws, attempt_limit):
    """Schedule a specific session (lab, lecture, or tutorial)"""
    code = str(course['Course Code'])
    name = str(course['Course Name'])
    faculty = str(course['Faculty'])
    classroom = str(course['Classroom'])
    
    # Determine duration based on session type
    if session_type == 'LAB':
        duration = LAB_DURATION
    elif 'LEC' in session_type:
        duration = LECTURE_DURATION
    else:  # TUT
        duration = TUTORIAL_DURATION
    
    # Track scheduling status
    scheduled = False
    attempts = 0
    
    # First try a more intelligent approach: schedule on days with fewer classes
    day_load = {day: sum(1 for slot in timetable[day].values() if slot['type'] is not None) 
               for day in range(len(DAYS))}
    sorted_days = sorted(day_load.keys(), key=lambda d: day_load[d])
    
    # Try days in order from least busy to most busy
    for day in sorted_days:
        if scheduled:
            break
            
        # Start with earlier time slots to avoid fragmentation
        for start_slot in range(len(TIME_SLOTS)-duration+1):
            # Check if all required slots are free for scheduling
            if check_scheduling_possibility(faculty, classroom, day, start_slot, duration, 
                                         professor_schedule, classroom_schedule, timetable, TIME_SLOTS):
                # Update all schedules
                update_schedule(faculty, classroom, day, start_slot, duration, session_type, 
                             code, name, professor_schedule, classroom_schedule, timetable)
                scheduled = True
                summary_ws.append([department, semester, code, name, session_type, faculty, classroom, "Scheduled", 
                                  f"{DAYS[day]} {TIME_SLOTS[start_slot][0].strftime('%H:%M')}"])
                break
    
    # Fall back to random attempts if structured approach fails
    while not scheduled and attempts < attempt_limit:
        day = random.randint(0, len(DAYS)-1)
        start_slot = random.randint(0, len(TIME_SLOTS)-duration)
        
        # Check if all required slots are free for scheduling
        if check_scheduling_possibility(faculty, classroom, day, start_slot, duration, 
                                     professor_schedule, classroom_schedule, timetable, TIME_SLOTS):
            # Update all schedules
            update_schedule(faculty, classroom, day, start_slot, duration, session_type, 
                         code, name, professor_schedule, classroom_schedule, timetable)
            scheduled = True
            summary_ws.append([department, semester, code, name, session_type, faculty, classroom, "Scheduled", 
                              f"{DAYS[day]} {TIME_SLOTS[start_slot][0].strftime('%H:%M')}"])
        attempts += 1
    
    if not scheduled:
        logging.warning(f"Failed to schedule {session_type} for {code}: {name} - Faculty: {faculty}, Classroom: {classroom}")
        summary_ws.append([department, semester, code, name, session_type, faculty, classroom, "Failed", "N/A"])
        
    return scheduled

def handle_lectures(department, semester, course, professor_schedule, classroom_schedule, timetable, TIME_SLOTS, summary_ws, attempt_limit):
    """Handle scheduling lectures based on L value"""
    l = int(course['L'])
    total_scheduled = 0
    failed = 0
    
    # If L=3, schedule two 1.5-hour lectures
    if l == 3:
        for lecture_idx in range(2):  # Two lectures
            lec_scheduled = schedule_session(
                department, semester, course, f'LEC {lecture_idx+1}', 
                professor_schedule, classroom_schedule, 
                timetable, TIME_SLOTS, summary_ws, attempt_limit
            )
            if lec_scheduled:
                total_scheduled += 1
            else:
                failed += 1
    else:
        # For other L values, schedule L lectures
        for lecture_idx in range(l):
            lec_scheduled = schedule_session(
                department, semester, course, f'LEC {lecture_idx+1}', 
                professor_schedule, classroom_schedule, 
                timetable, TIME_SLOTS, summary_ws, attempt_limit
            )
            if lec_scheduled:
                total_scheduled += 1
            else:
                failed += 1
    
    return total_scheduled, failed

def generate_all_timetables():
    """Main function to generate timetables for all departments and semesters"""
    TIME_SLOTS = generate_time_slots()
    logging.info(f"Generated {len(TIME_SLOTS)} time slots from {START_TIME} to {END_TIME}")
    
    wb = Workbook()
    wb.remove(wb.active)  # Remove default sheet
    
    # Create a summary sheet for failed schedules
    summary_ws = wb.create_sheet(title="Scheduling_Summary")
    summary_ws.append(["Department", "Semester", "Course Code", "Course Name", "Activity Type", 
                      "Faculty", "Classroom", "Scheduling Status", "Time"])
    
    professor_schedule = {}   # Track professor assignments
    classroom_schedule = {}   # Track classroom assignments
    
    # Load and clean course data
    df = load_and_clean_data()
    
    # Count total number of courses to schedule
    total_courses = 0
    scheduled_courses = 0
    failed_courses = 0
    
    # Process all departments and semesters
    departments = df['Department'].unique()
    logging.info(f"Found {len(departments)} departments to process")
    
    for department in departments:
        if pd.isna(department) or department == "":
            continue  # Skip empty department entries
            
        semester_groups = df[df['Department'] == department]['Semester'].unique()
        logging.info(f"Processing department: {department} with {len(semester_groups)} semester groups")
        
        for semester in semester_groups:
            if pd.isna(semester) or semester == "":
                continue  # Skip empty semester entries
                
            courses = df[(df['Department'] == department) & (df['Semester'] == str(semester))].copy()
            
            if courses.empty:
                logging.warning(f"No courses found for {department} semester {semester}")
                continue
            
            logging.info(f"Processing {department} semester {semester} with {len(courses)} courses")
            
            # Create worksheet for this department-semester
            ws_title = f"{department}_{semester}"
            ws_title = ws_title[:31]  # Excel worksheet names are limited to 31 characters
            ws = wb.create_sheet(title=ws_title)
            
            # Initialize timetable structure with expanded metadata
            timetable = {
                day: {
                    slot: {
                        'type': None, 
                        'code': '', 
                        'name': '', 
                        'faculty': '', 
                        'classroom': '',
                        'duration': 0,
                        'is_first': False,
                        'position': 0
                    } 
                    for slot in range(len(TIME_SLOTS))
                } 
                for day in range(len(DAYS))
            }
            
            # Special handling for DSAI and ECE departments
            priority_multiplier = 1.5 if department in ['DSAI', 'ECE'] else 1
            attempt_limit = int(MAX_SCHEDULING_ATTEMPTS * priority_multiplier)
            
            # First handle courses with both labs and lectures/tutorials
            combined_courses = courses[(courses['P'] > 0) & ((courses['L'] > 0) | (courses['T'] > 0))]
            for _, course in combined_courses.iterrows():
                course_scheduled = True
                
                # Schedule lab first
                p = int(course['P'])
                if p > 0:
                    total_courses += 1
                    lab_scheduled = schedule_session(
                        department, semester, course, 'LAB', 
                        professor_schedule, classroom_schedule, 
                        timetable, TIME_SLOTS, summary_ws, attempt_limit
                    )
                    if lab_scheduled:
                        scheduled_courses += 1
                    else:
                        failed_courses += 1
                        course_scheduled = False
                
                # Schedule lectures
                l = int(course['L'])
                if l > 0:
                    # If lab failed and this is DSAI or ECE, try harder
                    current_attempt_limit = attempt_limit * 2 if not course_scheduled and department in ['DSAI', 'ECE'] else attempt_limit
                    
                    lectures_scheduled, lectures_failed = handle_lectures(
                        department, semester, course, 
                        professor_schedule, classroom_schedule, 
                        timetable, TIME_SLOTS, summary_ws, 
                        current_attempt_limit
                    )
                    
                    # If L=3, we count it as 2 courses for statistics (since we're scheduling 2 lectures)
                    if l == 3:
                        total_courses += 2
                    else:
                        total_courses += l
                        
                    scheduled_courses += lectures_scheduled
                    failed_courses += lectures_failed
                
                # Schedule tutorials
                t = int(course['T'])
                for tutorial_idx in range(t):
                    total_courses += 1
                    # If lab failed and this is DSAI or ECE, try harder
                    current_attempt_limit = attempt_limit * 2 if not course_scheduled and department in ['DSAI', 'ECE'] else attempt_limit
                    
                    tut_scheduled = schedule_session(
                        department, semester, course, f'TUT {tutorial_idx+1}', 
                        professor_schedule, classroom_schedule, 
                        timetable, TIME_SLOTS, summary_ws, current_attempt_limit
                    )
                    if tut_scheduled:
                        scheduled_courses += 1
                    else:
                        failed_courses += 1
            
            # Process remaining labs
            lab_courses = courses[(courses['P'] > 0) & ~((courses['L'] > 0) | (courses['T'] > 0))]
            for _, course in lab_courses.iterrows():
                total_courses += 1
                lab_scheduled = schedule_session(
                    department, semester, course, 'LAB', 
                    professor_schedule, classroom_schedule, 
                    timetable, TIME_SLOTS, summary_ws, attempt_limit
                )
                if lab_scheduled:
                    scheduled_courses += 1
                else:
                    failed_courses += 1
            
            # Process remaining lectures and tutorials
            other_courses = courses[courses['P'] == 0]
            for _, course in other_courses.iterrows():
                l = int(course['L'])
                if l > 0:
                    lectures_scheduled, lectures_failed = handle_lectures(
                        department, semester, course, 
                        professor_schedule, classroom_schedule, 
                        timetable, TIME_SLOTS, summary_ws, 
                        attempt_limit
                    )
                    
                    # If L=3, we count it as 2 courses for statistics (since we're scheduling 2 lectures)
                    if l == 3:
                        total_courses += 2
                    else:
                        total_courses += l
                        
                    scheduled_courses += lectures_scheduled
                    failed_courses += lectures_failed
                
                # Schedule tutorials
                t = int(course['T'])
                for tutorial_idx in range(t):
                    total_courses += 1
                    tut_scheduled = schedule_session(
                        department, semester, course, f'TUT {tutorial_idx+1}', 
                        professor_schedule, classroom_schedule, 
                        timetable, TIME_SLOTS, summary_ws, attempt_limit
                    )
                    if tut_scheduled:
                        scheduled_courses += 1
                    else:
                        failed_courses += 1
            
            # Write timetable to worksheet with improved merged cells handling
            # Create header
            header = ['Day'] + [f"{slot[0].strftime('%H:%M')}-{slot[1].strftime('%H:%M')}" for slot in TIME_SLOTS]
            ws.append(header)
            
            # Apply header formatting
            header_fill = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")
            header_font = Font(bold=True)
            header_alignment = Alignment(horizontal='center', vertical='center')
            
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_alignment
            
            # Define fill colors for different session types
            lec_fill = PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid")  # Lavender
            lab_fill = PatternFill(start_color="98FB98", end_color="98FB98", fill_type="solid")  # Pale Green
            tut_fill = PatternFill(start_color="FFE4E1", end_color="FFE4E1", fill_type="solid")  # Misty Rose
            break_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid") # Light Gray
            conflict_fill = PatternFill(start_color="FF6347", end_color="FF6347", fill_type="solid") # Tomato
            border = Border(left=Side(style='thin'), right=Side(style='thin'),
                          top=Side(style='thin'), bottom=Side(style='thin'))
            
            # IMPROVED WRITING ALGORITHM: Process each day and prepare all cell values and merges
            for day_idx, day in enumerate(DAYS):
                row_num = day_idx + 2  # +1 for header, +1 because rows start at 1
                ws.append([day])
                
                # First, mark all occupied cells (this will help us avoid merge conflicts)
                occupied_cells = [False] * len(TIME_SLOTS)
                
                # Mark break times
                for slot_idx in range(len(TIME_SLOTS)):
                    if is_break_time(TIME_SLOTS[slot_idx]):
                        occupied_cells[slot_idx] = True
                
                # Track which cells need to be merged and their merge ranges
                merges = {}  # key: start slot index, value: (end slot index, activity details)
                
                # First pass - identify merges 
                for slot_idx in range(len(TIME_SLOTS)):
                    slot_info = timetable[day_idx][slot_idx]
                    
                    # If this is the first slot of a multi-slot activity
                    if slot_info['is_first'] and slot_info['duration'] > 1:
                        end_slot = slot_idx + slot_info['duration'] - 1
                        # Check if any of these slots are already marked as occupied
                        conflict = False
                        for i in range(slot_idx, end_slot + 1):
                            if i >= len(TIME_SLOTS) or occupied_cells[i]:
                                conflict = True
                                break
                                
                        if not conflict:
                            # Mark all these slots as occupied
                            for i in range(slot_idx, end_slot + 1):
                                occupied_cells[i] = True
                            # Store the merge information
                            merges[slot_idx] = (end_slot, {
                                'type': slot_info['type'],
                                'code': slot_info['code'],
                                'name': slot_info['name'],
                                'faculty': slot_info['faculty'],
                                'classroom': slot_info['classroom']
                            })
                        else:
                            # Mark only this cell as occupied - it's a conflict
                            occupied_cells[slot_idx] = True
                    
                    # For single-slot activities (like break times or conflict indicators)
                    elif slot_info['type'] is not None and not occupied_cells[slot_idx]:
                        occupied_cells[slot_idx] = True
                
                # Second pass - write cells and perform merges
                for slot_idx in range(len(TIME_SLOTS)):
                    cell_content = ''
                    cell_fill = None
                    
                    # First priority: breaks
                    if is_break_time(TIME_SLOTS[slot_idx]):
                        cell_content = "BREAK"
                        cell_fill = break_fill
                    
                    # Second priority: merge start points
                    elif slot_idx in merges:
                        end_slot, activity = merges[slot_idx]
                        activity_type = activity['type']
                        
                        # Set fill color based on activity type
                        if 'LEC' in activity_type:
                            cell_fill = lec_fill
                        elif activity_type == 'LAB':
                            cell_fill = lab_fill
                        else:  # TUT
                            cell_fill = tut_fill
                        
                        # Create cell content
                        cell_content = f"{activity['code']} {activity_type}\n{activity['name']}\n{activity['faculty']}\n{activity['classroom']}"
                        
                        # Merge cells
                        start_col = get_column_letter(slot_idx + 2)  # +1 for day column
                        end_col = get_column_letter(end_slot + 2)
                        try:
                            ws.merge_cells(f"{start_col}{row_num}:{end_col}{row_num}")
                            logging.info(f"Successfully merged cells for {activity['code']} {activity_type} on {day}")
                        except Exception as e:
                            logging.warning(f"Failed to merge cells for {activity['code']} {activity_type} on {day}: {str(e)}")
                    
                    # Third priority: cells that are part of a merged range (skip them)
                    elif any(slot_idx > start and slot_idx <= end for start, (end, _) in merges.items()):
                        continue
                    
                    # Fourth priority: individual activities or conflict markers
                    elif timetable[day_idx][slot_idx]['type'] is not None:
                        code = timetable[day_idx][slot_idx]['code']
                        activity_type = timetable[day_idx][slot_idx]['type']
                        
                        # Check if this should be a merged cell but couldn't be merged
                        if timetable[day_idx][slot_idx]['is_first'] and timetable[day_idx][slot_idx]['duration'] > 1:
                            cell_content = f"{code} {activity_type} - CONFLICT"
                            cell_fill = conflict_fill
                        else:
                            # Regular single-slot activity
                            name = timetable[day_idx][slot_idx]['name']
                            faculty = timetable[day_idx][slot_idx]['faculty']
                            classroom = timetable[day_idx][slot_idx]['classroom']
                            cell_content = f"{code} {activity_type}\n{name}\n{faculty}\n{classroom}"
                            
                            # Set fill color based on activity type
                            if 'LEC' in activity_type:
                                cell_fill = lec_fill
                            elif activity_type == 'LAB':
                                cell_fill = lab_fill
                            else:  # TUT
                                cell_fill = tut_fill
                    
                    # Write the cell content and apply formatting
                    cell = ws.cell(row=row_num, column=slot_idx+2, value=cell_content)
                    if cell_fill:
                        cell.fill = cell_fill
                    cell.border = border
                    cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
            
            # Adjust column widths and row heights
            for col_idx in range(1, len(TIME_SLOTS)+2):
                col_letter = get_column_letter(col_idx)
                ws.column_dimensions[col_letter].width = 18  # Slightly wider columns for better readability
            
            for row in ws.iter_rows(min_row=2, max_row=len(DAYS)+1):
                ws.row_dimensions[row[0].row].height = 80  # Taller rows for better readability
    
    # Format summary worksheet
    for col_idx in range(1, 10):  # One more column for time
        col_letter = get_column_letter(col_idx)
        summary_ws.column_dimensions[col_letter].width = 20
    
    # Apply styles to summary sheet
    for row in summary_ws.iter_rows(min_row=1, max_row=1):
        for cell in row:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
    
    # Add summary statistics at the top
    stats_ws = wb.create_sheet(title="Statistics", index=0)
    stats_ws.append(["Timetable Generation Statistics"])
    stats_ws.append(["Total courses processed:", total_courses])
    stats_ws.append(["Successfully scheduled:", scheduled_courses])
    stats_ws.append(["Failed to schedule:", failed_courses])
    stats_ws.append(["Success rate:", f"{scheduled_courses/total_courses*100:.2f}%" if total_courses > 0 else "N/A"])
    
    # Department-wise statistics
    stats_ws.append([])
    stats_ws.append(["Department-wise Statistics:"])
    stats_ws.append(["Department", "Scheduled", "Failed", "Success Rate"])
    
    # Calculate department-wise statistics
    dept_stats = {}
    for row in summary_ws.iter_rows(min_row=2, values_only=True):
        dept = row[0]
        status = row[7]
        if dept not in dept_stats:
            dept_stats[dept] = {'scheduled': 0, 'failed': 0}
        
        if status == 'Scheduled':
            dept_stats[dept]['scheduled'] += 1
        else:           
            dept_stats[dept]['failed'] += 1
    
    # Add department statistics to worksheet
    for dept, stats in dept_stats.items():
        total = stats['scheduled'] + stats['failed']
        success_rate = (stats['scheduled'] / total * 100) if total > 0 else 0
        stats_ws.append([dept, stats['scheduled'], stats['failed'], f"{success_rate:.2f}%"])
    
    # Apply formatting to stats sheet
    stats_ws.column_dimensions['A'].width = 25
    stats_ws.column_dimensions['B'].width = 15
    stats_ws.row_dimensions[1].height = 30
    
    stats_ws['A1'].font = Font(bold=True, size=14)
    stats_ws.merge_cells('A1:B1')
    stats_ws['A1'].alignment = Alignment(horizontal='center')
    
    # Save the workbook
    output_file = "improved_timetables.xlsx"
    try:
        wb.save(output_file)
        logging.info(f"Final timetables saved to {output_file}")
        logging.info(f"Successfully scheduled {scheduled_courses} out of {total_courses} courses ({scheduled_courses/total_courses*100:.2f}%)")
        print(f"Final timetables saved to {output_file}")
        print(f"Successfully scheduled {scheduled_courses} out of {total_courses} courses ({scheduled_courses/total_courses*100:.2f}%)")
    except PermissionError:
        alt_file = f"improved_timetables_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        wb.save(alt_file)
        logging.warning(f"Could not save to {output_file} (file may be open). Saved to {alt_file} instead")
        print(f"Could not save to {output_file} (file may be open). Saved to {alt_file} instead")

if __name__ == "__main__":
    generate_all_timetables()