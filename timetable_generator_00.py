def schedule_base_semester_electives(df, base_sem, professor_schedule, classroom_schedule, timetable, TIME_SLOTS, summary_ws, attempt_limit):
    """Schedule all electives in the same base semester across departments at the same time"""
    logging.info(f"\n{'='*50}\nProcessing base semester: {base_sem}\n{'='*50}")
    
    # Get all electives for this base semester across all departments
    elective_group = df[(df['is_elective']) & (df['base_semester'] == str(base_sem))]
    
    if elective_group.empty:
        logging.info(f"No electives found for base semester {base_sem}")
        return

    # Schedule lectures, tutorials, and labs together
    for session_type, duration in [('LEC', LECTURE_DURATION), 
                                  ('TUT', TUTORIAL_DURATION), 
                                  ('LAB', LAB_DURATION)]:
        
        # Collect all required sessions across courses
        session_counts = []
        for _, course in elective_group.iterrows():
            count = 0
            if session_type == 'LEC':
                count = 2 if course['L'] == 3 else int(course['L'])
            elif session_type == 'TUT':
                count = int(course['T'])
            elif session_type == 'LAB':
                count = int(course['P'])
            
            for i in range(count):
                session_counts.append((course, i+1))

        # Process each session instance
        for session_idx in range(max([s[1] for s in session_counts] if session_counts else 0)):
            # Find common slot for this session instance across all courses
            best_slot = find_base_sem_common_slot(
                elective_group, session_type, session_idx+1, duration,
                professor_schedule, classroom_schedule, timetable, TIME_SLOTS
            )

            if best_slot:
                day, start_slot = best_slot
                logging.info(f"Found common slot for {session_type} {session_idx+1} at {DAYS[day]} {TIME_SLOTS[start_slot][0].strftime('%H:%M')}")

                # Schedule all courses in this slot
                for _, course in elective_group.iterrows():
                    # Verify this course needs this session
                    if session_type == 'LEC' and course['L'] <= 0:
                        continue
                    if session_type == 'TUT' and course['T'] <= 0:
                        continue
                    if session_type == 'LAB' and course['P'] <= 0:
                        continue

                    # Check if already scheduled
                    if check_scheduling_possibility(
                        course['Faculty'], course['Classroom'], day, start_slot, duration,
                        professor_schedule, classroom_schedule, timetable, TIME_SLOTS
                    ):
                        update_schedule(
                            course['Faculty'], course['Classroom'], day, start_slot, duration,
                            f"{session_type} {session_idx+1}", course['Course Code'],
                            course['Course Name'], professor_schedule, classroom_schedule, timetable
                        )
                        summary_ws.append([
                            course['Department'], course['Semester'], course['Course Code'],
                            course['Course Name'], f"{session_type} {session_idx+1}",
                            course['Faculty'], course['Classroom'], "Scheduled (Cross-Department)",
                            f"{DAYS[day]} {TIME_SLOTS[start_slot][0].strftime('%H:%M')}"
                        ])
                    else:
                        summary_ws.append([
                            course['Department'], course['Semester'], course['Course Code'],
                            course['Course Name'], f"{session_type} {session_idx+1}",
                            course['Faculty'], course['Classroom'], "Failed (Cross-Department Conflict)",
                            "N/A"
                        ])

def find_base_sem_common_slot(elective_group, session_type, session_num, duration, professor_schedule, classroom_schedule, timetable, TIME_SLOTS):
    """Find time slot that works for all courses in the base semester group"""
    best_day = -1
    best_start = -1
    max_score = 0

    # Try all possible time slots
    for day in range(len(DAYS)):
        for start_slot in range(len(TIME_SLOTS) - duration + 1):
            score = 0
            
            # Check availability for all courses in group
            for _, course in elective_group.iterrows():
                # Check if course needs this session type
                if session_type == 'LEC' and (course['L'] <= 0 or (course['L'] == 3 and session_num > 2)):
                    continue
                if session_type == 'TUT' and (course['T'] < session_num):
                    continue
                if session_type == 'LAB' and (course['P'] <= 0):
                    continue

                if check_scheduling_possibility(
                    course['Faculty'], course['Classroom'], day, start_slot, duration,
                    professor_schedule, classroom_schedule, timetable, TIME_SLOTS
                ):
                    score += 1

            # Update best slot if this is better
            if score > max_score:
                max_score = score
                best_day = day
                best_start = start_slot
                if max_score == len(elective_group):
                    return (best_day, best_start)  # Found perfect slot

    return (best_day, best_start) if best_day != -1 else None

# In identify_electives function add:
df['base_semester'] = df['Semester'].str.extract('(\d+)').fillna(df['Semester'])

# In generate_all_timetables function, after identifying electives:
base_semesters = df['base_semester'].unique()
for base_sem in base_semesters:
    schedule_base_semester_electives(
        df, base_sem, professor_schedule, classroom_schedule,
        timetable, TIME_SLOTS, summary_ws, MAX_SCHEDULING_ATTEMPTS
    )
