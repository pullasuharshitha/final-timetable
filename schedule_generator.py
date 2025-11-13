"""Core scheduling logic for generating timetables from Excel data."""
import pandas as pd
import random
from config import DAYS, TEACHING_SLOTS, LECTURE_DURATION, TUTORIAL_DURATION, LAB_DURATION, MINOR_DURATION
from config import PRE_MID, POST_MID, MINOR_SUBJECT, MINOR_CLASSES_PER_WEEK, DEPARTMENTS
from config import MINOR_SLOTS, LUNCH_SLOTS
from excel_loader import ExcelLoader

class ScheduleGenerator:
    """Generates weekly class schedules for semesters and departments from Excel data."""
    
    def __init__(self, data_frames):
        """Initialize ScheduleGenerator with data frames."""
        self.dfs = data_frames
        # Track global slots per semester to avoid clashes between departments in same semester
        self.semester_global_slots = {}
        # Track room occupancy per (semester, day, slot)
        self.room_occupancy = {}
        # Track detailed room bookings per (sem_key, day, slot) for conflict validation
        self.room_bookings = {}
        # Load classrooms (room_name, capacity) and classify by type
        self.classrooms = []
        self.lab_rooms = []
        self.software_lab_rooms = []
        self.hardware_lab_rooms = []
        self.nonlab_rooms = []
        try:
            cls_df = self.dfs.get('classroom')
            if cls_df is not None and not cls_df.empty:
                name_col = None
                cap_col = None
                type_col = None
                for col in cls_df.columns:
                    cl = str(col).lower()
                    if name_col is None and any(k in cl for k in ['room', 'class', 'hall', 'name']):
                        name_col = col
                    if cap_col is None and any(k in cl for k in ['cap', 'seats', 'capacity']):
                        cap_col = col
                    if type_col is None and any(k in cl for k in ['type', 'category', 'room type']):
                        type_col = col
                if name_col is None:
                    name_col = cls_df.columns[0]
                for _, row in cls_df.iterrows():
                    room_name = str(row.get(name_col, '')).strip()
                    try:
                        capacity = int(float(row.get(cap_col, 0))) if cap_col is not None else 0
                    except Exception:
                        capacity = 0
                    room_type = str(row.get(type_col, '')).strip().lower() if type_col is not None else ''
                    if room_name:
                        self.classrooms.append((room_name, capacity))
                        if 'lab' in room_type:
                            self.lab_rooms.append((room_name, capacity))
                            if 'software' in room_type or 'soft' in room_type:
                                self.software_lab_rooms.append((room_name, capacity))
                            if 'hardware' in room_type or 'hard' in room_type:
                                self.hardware_lab_rooms.append((room_name, capacity))
                        else:
                            self.nonlab_rooms.append((room_name, capacity))
                self.classrooms.sort(key=lambda x: (x[1], x[0]))
                self.lab_rooms.sort(key=lambda x: (x[1], x[0]))
                self.software_lab_rooms.sort(key=lambda x: (x[1], x[0]))
                self.hardware_lab_rooms.sort(key=lambda x: (x[1], x[0]))
                self.nonlab_rooms.sort(key=lambda x: (x[1], x[0]))
        except Exception:
            self.classrooms = []
            self.lab_rooms = []
            self.software_lab_rooms = []
            self.hardware_lab_rooms = []
            self.nonlab_rooms = []
        # Store minor slots per semester
        self.semester_minor_slots = {}
        # Store elective slots per semester, keyed by (semester_id, elective_code)
        self.semester_elective_slots = {}
        # 240-seater combined-class capacity per semester: set of (day, slot)
        self.semester_combined_capacity = {}
        # Combined class assigned slots per course and component:
        # key=(semester_id, course_code, component['Lecture'|'Tutorial'|'Lab']) -> list[(day, slot)]
        self.semester_combined_course_slots = {}
        # Global combined course slots shared across semesters but per allowed pairing group:
        # key=('GLOBAL', group_key, course_code, component) -> list[(day, start_slot)]
        self.global_combined_course_slots = {}
        self.scheduled_slots = {}  # Track all scheduled slots by semester+department
        self.scheduled_courses = {}  # Track when each course is scheduled
        self.actual_allocations = {}  # Track actual allocated counts: key=(semester_id, dept, session, course_code), value={'lectures': X, 'tutorials': Y, 'labs': Z}
        self.assigned_rooms = {}
        self.assigned_lab_rooms = {}
        
    def _initialize_schedule(self):
        """Initialize an empty schedule with Days as rows and Time Slots as columns."""
        schedule = pd.DataFrame(index=DAYS, columns=TEACHING_SLOTS)
        
        # Initialize with 'Free'
        for day in DAYS:
            for slot in TEACHING_SLOTS:
                schedule.loc[day, slot] = 'Free'
        
        # Mark lunch break (now possibly multiple 30-min slots)
        for day in DAYS:
            for lunch_slot in LUNCH_SLOTS:
                if lunch_slot in schedule.columns:
                    schedule.loc[day, lunch_slot] = 'LUNCH BREAK'
        
        return schedule
    
    def _get_consecutive_slots(self, start_slot, duration):
        """Get consecutive time slots for a given duration."""
        try:
            start_index = TEACHING_SLOTS.index(start_slot)
            end_index = start_index + duration
            if end_index <= len(TEACHING_SLOTS):
                return TEACHING_SLOTS[start_index:end_index]
        except ValueError:
            pass
        return []
    
    def _get_dept_from_global_key(self, dept_key):
        """Extract department label from a global slot key (e.g., 'CSE-A' from 'CSE-A_Pre-Mid')."""
        return dept_key.split('_')[0] if dept_key else ''

    def _departments_can_share_slots(self, dept_a, dept_b):
        """Return True if two departments are allowed to share the same time slots."""
        if not dept_a or not dept_b:
            return False

        share_groups = [
            {"CSE-A", "CSE-B"},
        ]

        for group in share_groups:
            if dept_a in group and dept_b in group:
                return True
        return False

    def _is_time_slot_available_global(self, day, slots, department, session, semester_id):
        """Enhanced slot availability check to prevent conflicts.
        Rules:
        - Same department + same session = conflict (same students can't be in two classes)
        - Same department + different session = OK (different students)
        - Different departments = OK (different students, can share slots)
        - CSE-A and CSE-B can share slots for the same courses."""
        semester_key = f"sem_{semester_id}"
        
        # Use the same tracking system as _mark_slots_busy_global
        if semester_key not in self.semester_global_slots:
            return True  # No slots booked yet for this semester
        
        # Check for conflicts
        for slot in slots:
            # Check semester-wide conflicts
            for dept_key, used_slots in self.semester_global_slots[semester_key].items():
                if (day, slot) in used_slots:
                    # Extract department from dept_key (format: "DEPT_SESSION")
                    dept_in_slot = dept_key.split('_')[0] if '_' in dept_key else dept_key
                    session_in_slot = dept_key.split('_')[1] if '_' in dept_key else ''
                    
                    # Allow CSE-A and CSE-B to share slots (they can have same courses at same time)
                    if self._departments_can_share_slots(department, dept_in_slot):
                        continue  # Allow sharing between CSE-A and CSE-B
                    
                    # Allow same department different sessions (different students)
                    if department == dept_in_slot and session != session_in_slot:
                        continue
                    
                    # Allow different departments (different students, can share slots)
                    if department != dept_in_slot:
                        continue
                    
                    # Block: same department + same session = conflict
                    # (This means department == dept_in_slot and session == session_in_slot)
                    return False
        return True

    def _mark_slots_busy_global(self, day, slots, department, session, semester_id):
        """Mark time slots as busy in global tracker."""
        key = f"{department}_{session}"
        semester_key = f"sem_{semester_id}"
        
        if semester_key not in self.semester_global_slots:
            self.semester_global_slots[semester_key] = {}
        
        if key not in self.semester_global_slots[semester_key]:
            self.semester_global_slots[semester_key][key] = set()
        
        for slot in slots:
            self.semester_global_slots[semester_key][key].add((day, slot))
            # prepare room occupancy tracker
            occ_key = (semester_key, day, slot)
            if occ_key not in self.room_occupancy:
                self.room_occupancy[occ_key] = set()
    
    def _is_time_slot_available_local(self, schedule, day, slots):
        """Check if time slots are available in local schedule."""
        for slot in slots:
            if schedule.loc[day, slot] != 'Free':
                return False
        return True
    
    def _mark_slots_busy_local(self, schedule, day, slots, course_code, class_type):
        """Mark time slots as busy in local schedule."""
        suffix = ''
        if class_type == 'Lab':
            suffix = ' (Lab)'
        elif class_type == 'Tutorial':
            suffix = ' (Tut)'
        elif class_type == 'Minor':
            suffix = ' (Minor)'
        
        for slot in slots:
            schedule.loc[day, slot] = f"{course_code}{suffix}"
    
    def _schedule_minor_classes(self, schedule, department, session, semester_id):
        """Schedule Minor subject classes ONLY in configured MINOR_SLOTS (e.g., 07:30-08:30 split).
        All departments/sections in a semester get the same minor slots."""
        # Skip minor scheduling entirely for semester 1 as per requirement
        if int(semester_id) == 1:
            return
        scheduled = 0
        attempts = 0
        max_attempts = 200
        
        # Compute valid minor start slots (so MINOR_DURATION consecutive slots are within MINOR_SLOTS)
        minor_starts = []
        for s in MINOR_SLOTS:
            seq = self._get_consecutive_slots(s, MINOR_DURATION)
            if len(seq) == MINOR_DURATION and all(x in MINOR_SLOTS for x in seq):
                minor_starts.append(s)
        
        if not minor_starts:
            # nothing to schedule if config mismatched
            return
        
        semester_key = f"sem_{semester_id}"
        # If already assigned for this semester, use the same slots
        if semester_key in self.semester_minor_slots:
            assigned = self.semester_minor_slots[semester_key]
            for day, start in assigned:
                slots = self._get_consecutive_slots(start, MINOR_DURATION)
                self._mark_slots_busy_local(schedule, day, slots, MINOR_SUBJECT, 'Minor')
                self._mark_slots_busy_global(day, slots, department, session, semester_id)
            return

        assigned = []
        while scheduled < MINOR_CLASSES_PER_WEEK and attempts < max_attempts:
            attempts += 1
            
            day = random.choice(DAYS)
            start = random.choice(minor_starts)
            slots = self._get_consecutive_slots(start, MINOR_DURATION)
            
            if (len(slots) == MINOR_DURATION and
                self._is_time_slot_available_local(schedule, day, slots) and
                self._is_time_slot_available_global(day, slots, department, session, semester_id)):
                
                self._mark_slots_busy_local(schedule, day, slots, MINOR_SUBJECT, 'Minor')
                self._mark_slots_busy_global(day, slots, department, session, semester_id)
                assigned.append((day, start))
                scheduled += 1
        # Save assigned slots for all departments in this semester
        if assigned:
            self.semester_minor_slots[semester_key] = assigned

    def _schedule_lectures(self, schedule, course_code, lectures_per_week, department, session, semester_id, avoid_days=None):
        """Schedule lecture sessions, returns list of (day, slot) tuples."""
        if lectures_per_week == 0:
            return []
        scheduled_slots = []
        attempts = 0
        max_attempts = 2000  # Increased attempts for better allocation
        used_days = set()
        avoid_days = set(avoid_days or [])
        regular_slots = [slot for slot in TEACHING_SLOTS if slot not in (MINOR_SLOTS + LUNCH_SLOTS)]

        # Build list of all possible (day, start_slot) combinations
        all_combinations = []
        for day in DAYS:
            for start_idx in range(len(regular_slots) - LECTURE_DURATION + 1):
                start_slot = regular_slots[start_idx]
                slots = self._get_consecutive_slots(start_slot, LECTURE_DURATION)
                slots = [slot for slot in slots if slot in regular_slots]
                if len(slots) == LECTURE_DURATION:
                    all_combinations.append((day, start_slot, slots))
        
        # Shuffle combinations for randomness
        random.shuffle(all_combinations)

        while len(scheduled_slots) < lectures_per_week * LECTURE_DURATION and attempts < max_attempts:
            attempts += 1
            
            # Try to use days where this course isn't already scheduled
            available_combos = [combo for combo in all_combinations if combo[0] not in used_days and combo[0] not in avoid_days]
            if not available_combos:
                available_combos = all_combinations
            
            if not available_combos:
                break
            
            day, start_slot, slots = random.choice(available_combos)
            
            # Check all slots are available (both local and global)
            slots_available = True
            for slot in slots:
                if not (self._is_time_slot_available_local(schedule, day, [slot]) and
                       self._is_time_slot_available_global(day, [slot], department, session, semester_id)):
                    slots_available = False
                    break
            
            if slots_available:
                # Mark all slots as scheduled
                for slot in slots:
                    self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Lecture')
                    self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                    scheduled_slots.append((day, slot))
                used_days.add(day)
                avoid_days.add(day)
                # Remove this combination from future consideration
                all_combinations = [c for c in all_combinations if c != (day, start_slot, slots)]

        scheduled_count = len(scheduled_slots) // LECTURE_DURATION
        if scheduled_count < lectures_per_week:
            print(f"      WARNING: {course_code} - Only scheduled {scheduled_count}/{lectures_per_week} lectures (attempts: {attempts})")
        return scheduled_slots

    def _schedule_lectures_tracked(self, schedule, course_code, lectures_per_week, department, session, semester_id):
        """Schedule lectures and return the set of days used. Returns tuple (used_days, actual_scheduled_count)."""
        if lectures_per_week == 0:
            return set(), 0
        scheduled = 0
        attempts = 0
        max_attempts = 500  # Increased for better success rate
        used_days = set()
        while scheduled < lectures_per_week and attempts < max_attempts:
            attempts += 1
            available_days = [d for d in DAYS if d not in used_days]
            if not available_days:
                break
            day = random.choice(available_days)
            regular_slots = [slot for slot in TEACHING_SLOTS if slot not in (MINOR_SLOTS + LUNCH_SLOTS)]
            if not regular_slots:
                break
            start_slot = random.choice(regular_slots)
            slots = self._get_consecutive_slots(start_slot, LECTURE_DURATION)
            slots = [slot for slot in slots if slot in regular_slots]
            if (len(slots) == LECTURE_DURATION and 
                self._is_time_slot_available_local(schedule, day, slots) and
                self._is_time_slot_available_global(day, slots, department, session, semester_id)):
                self._mark_slots_busy_local(schedule, day, slots, course_code, 'Lecture')
                self._mark_slots_busy_global(day, slots, department, session, semester_id)
                used_days.add(day)
                scheduled += 1
        
        if scheduled < lectures_per_week:
            print(f"      WARNING: {course_code} - Only scheduled {scheduled}/{lectures_per_week} lectures (attempts: {attempts})")
        return used_days, scheduled

    def _schedule_elective_classes(self, schedule, course_code, elective_per_week, department, session, semester_id, avoid_days=None):
        """Schedule elective classes at the same slot/day for all departments/sections in a semester.
        ALL electives in a semester must use the same time slots across ALL departments (CSE, DSAI, ECE).
        Returns list of (day, slot) tuples."""
        if elective_per_week == 0:
            return []
        scheduled_slots = []
        attempts = 0
        max_attempts = 1000  # Increased for better success rate
        semester_key = f"sem_{semester_id}"
        avoid_days = set(avoid_days or [])
        
        # Use a common key for ALL electives in a semester to ensure same time slots
        # This ensures CSE, DSAI, and ECE all have electives at the same time
        common_elective_key = (semester_key, 'ALL_ELECTIVES')
        elective_key = (semester_key, course_code)
        used_days = set()
        regular_slots = [slot for slot in TEACHING_SLOTS if slot not in (MINOR_SLOTS + LUNCH_SLOTS)]
        
        # First, check if common elective slots have been assigned for this semester
        # If yes, ALL electives must use those same slots
        if common_elective_key in self.semester_elective_slots:
            assigned = self.semester_elective_slots[common_elective_key]
            print(f"      Using common elective slots for {course_code} (already assigned for semester {semester_id})")
            for day, start in assigned:
                slots = self._get_consecutive_slots(start, LECTURE_DURATION)
                if day in avoid_days:
                    continue
                for slot in slots:
                    self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Lecture')
                    self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                    scheduled_slots.append((day, slot))
            # Also store under course-specific key for backward compatibility
            self.semester_elective_slots[elective_key] = assigned
            return scheduled_slots
        
        # If no common slots assigned yet, check if this specific course has slots assigned
        # (This handles legacy cases where course-specific slots were assigned first)
        if elective_key in self.semester_elective_slots:
            assigned = self.semester_elective_slots[elective_key]
            # Promote to common slots so all future electives use the same slots
            self.semester_elective_slots[common_elective_key] = assigned
            print(f"      Using existing elective slots for {course_code} (promoted to common slots for semester {semester_id})")
            for day, start in assigned:
                slots = self._get_consecutive_slots(start, LECTURE_DURATION)
                if day in avoid_days:
                    continue
                for slot in slots:
                    self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Lecture')
                    self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                    scheduled_slots.append((day, slot))
            return scheduled_slots

        # Build list of all possible (day, start_slot) combinations
        all_combinations = []
        for day in DAYS:
            for start_idx in range(len(regular_slots) - LECTURE_DURATION + 1):
                start_slot = regular_slots[start_idx]
                slots = self._get_consecutive_slots(start_slot, LECTURE_DURATION)
                slots = [slot for slot in slots if slot in regular_slots]
                if len(slots) == LECTURE_DURATION:
                    all_combinations.append((day, start_slot, slots))
        
        # Shuffle combinations for randomness
        random.shuffle(all_combinations)

        assigned = []
        scheduled = 0
        while scheduled < elective_per_week and attempts < max_attempts:
            attempts += 1
            available_combos = [combo for combo in all_combinations if combo[0] not in used_days and combo[0] not in avoid_days]
            if not available_combos:
                available_combos = all_combinations
            if not available_combos:
                break
            
            day, start_slot, slots = random.choice(available_combos) if available_combos else (None, None, None)
            if day is None:
                break
            
            if (self._is_time_slot_available_local(schedule, day, slots) and
                self._is_time_slot_available_global(day, slots, department, session, semester_id)):
                for slot in slots:
                    self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Lecture')
                    self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                    scheduled_slots.append((day, slot))
                assigned.append((day, start_slot))
                used_days.add(day)
                avoid_days.add(day)
                scheduled += 1
                # Remove this combination from future consideration
                all_combinations = [c for c in all_combinations if c != (day, start_slot, slots)]
        
        if assigned:
            # Store under common key so ALL electives in this semester use the same slots
            self.semester_elective_slots[common_elective_key] = assigned
            # Also store under course-specific key for backward compatibility
            self.semester_elective_slots[elective_key] = assigned
            print(f"      Assigned common elective slots for semester {semester_id}: {assigned}")
        
        scheduled_count = len(scheduled_slots) // LECTURE_DURATION
        if scheduled_count < elective_per_week:
            print(f"      WARNING: {course_code} - Only scheduled {scheduled_count}/{elective_per_week} elective classes (attempts: {attempts})")
        return scheduled_slots
    
    def _schedule_elective_classes_tracked(self, schedule, course_code, elective_per_week, department, session, semester_id):
        """Schedule elective classes and return tuple (used_days, actual_scheduled_count).
        Uses common elective slots for ALL departments in a semester."""
        if elective_per_week == 0:
            return set(), 0
        scheduled = 0
        attempts = 0
        max_attempts = 500  # Increased for better success rate
        semester_key = f"sem_{semester_id}"
        common_elective_key = (semester_key, 'ALL_ELECTIVES')
        elective_key = (semester_key, course_code)
        used_days = set()
        
        # First check for common elective slots (all departments use same slots)
        if common_elective_key in self.semester_elective_slots:
            assigned = self.semester_elective_slots[common_elective_key]
            for day, start in assigned:
                slots = self._get_consecutive_slots(start, LECTURE_DURATION)
                self._mark_slots_busy_local(schedule, day, slots, course_code, 'Lecture')
                self._mark_slots_busy_global(day, slots, department, session, semester_id)
                used_days.add(day)
            scheduled = len(assigned)
            # Also store under course-specific key for backward compatibility
            self.semester_elective_slots[elective_key] = assigned
            return used_days, scheduled
        
        # Check if this specific course has slots assigned (legacy support)
        if elective_key in self.semester_elective_slots:
            assigned = self.semester_elective_slots[elective_key]
            # Promote to common slots
            self.semester_elective_slots[common_elective_key] = assigned
            for day, start in assigned:
                slots = self._get_consecutive_slots(start, LECTURE_DURATION)
                self._mark_slots_busy_local(schedule, day, slots, course_code, 'Lecture')
                self._mark_slots_busy_global(day, slots, department, session, semester_id)
                used_days.add(day)
            scheduled = len(assigned)
            return used_days, scheduled

        assigned = []
        while scheduled < elective_per_week and attempts < max_attempts:
            attempts += 1
            available_days = [d for d in DAYS if d not in used_days]
            if not available_days:
                break
            day = random.choice(available_days)
            regular_slots = [slot for slot in TEACHING_SLOTS if slot not in (MINOR_SLOTS + LUNCH_SLOTS)]
            if not regular_slots:
                break
            start_slot = random.choice(regular_slots)
            slots = self._get_consecutive_slots(start_slot, LECTURE_DURATION)
            slots = [slot for slot in slots if slot in regular_slots]
            if (len(slots) == LECTURE_DURATION and
                self._is_time_slot_available_local(schedule, day, slots) and
                self._is_time_slot_available_global(day, slots, department, session, semester_id)):
                self._mark_slots_busy_local(schedule, day, slots, course_code, 'Lecture')
                self._mark_slots_busy_global(day, slots, department, session, semester_id)
                assigned.append((day, start_slot))
                used_days.add(day)
                scheduled += 1
        if assigned:
            # Store under common key so ALL electives in this semester use the same slots
            self.semester_elective_slots[common_elective_key] = assigned
            # Also store under course-specific key for backward compatibility
            self.semester_elective_slots[elective_key] = assigned
        
        if scheduled < elective_per_week:
            print(f"      WARNING: {course_code} - Only scheduled {scheduled}/{elective_per_week} elective classes (attempts: {attempts})")
        return used_days, scheduled

    def _schedule_tutorials(self, schedule, course_code, tutorials_per_week, department, session, semester_id, avoid_days=None):
        """Schedule tutorial sessions, returns list of (day, slot) tuples."""
        if tutorials_per_week == 0:
            return []

        scheduled_slots = []
        attempts = 0
        max_attempts = 1000
        used_days = set()
        avoid_days = set(avoid_days or [])
        regular_slots = [slot for slot in TEACHING_SLOTS if slot not in (MINOR_SLOTS + LUNCH_SLOTS)]

        # Build list of all possible (day, start_slot) combinations
        all_combinations = []
        for day in DAYS:
            for start_idx in range(len(regular_slots) - TUTORIAL_DURATION + 1):
                start_slot = regular_slots[start_idx]
                slots = self._get_consecutive_slots(start_slot, TUTORIAL_DURATION)
                slots = [slot for slot in slots if slot in regular_slots]
                if len(slots) == TUTORIAL_DURATION:
                    all_combinations.append((day, start_slot, slots))
        
        # Shuffle combinations for randomness
        random.shuffle(all_combinations)

        while len(scheduled_slots) < tutorials_per_week * TUTORIAL_DURATION and attempts < max_attempts:
            attempts += 1
            available_combos = [combo for combo in all_combinations if combo[0] not in used_days and combo[0] not in avoid_days]
            if not available_combos:
                available_combos = all_combinations
            if not available_combos:
                break

            day, start_slot, slots = random.choice(available_combos)
            
            if (self._is_time_slot_available_local(schedule, day, slots) and
                self._is_time_slot_available_global(day, slots, department, session, semester_id)):
                
                for slot in slots:
                    self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Tutorial')
                    self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                    scheduled_slots.append((day, slot))
                used_days.add(day)
                avoid_days.add(day)
                # Remove this combination from future consideration
                all_combinations = [c for c in all_combinations if c != (day, start_slot, slots)]

        scheduled_count = len(scheduled_slots) // TUTORIAL_DURATION
        if scheduled_count < tutorials_per_week:
            print(f"      WARNING: {course_code} - Only scheduled {scheduled_count}/{tutorials_per_week} tutorials (attempts: {attempts})")
        
        return scheduled_slots

    def _schedule_labs(self, schedule, course_code, labs_per_week, department, session, semester_id, avoid_days=None):
        """Schedule lab sessions in regular time slots (multi-slot labs).
        Returns list of (day, slot) tuples."""
        if labs_per_week == 0:
            return []
        scheduled_slots = []
        attempts = 0
        max_attempts = 1000  # Increased for better success rate
        used_days = set()
        avoid_days = set(avoid_days or [])
        regular_slots = [slot for slot in TEACHING_SLOTS if slot not in (MINOR_SLOTS + LUNCH_SLOTS)]
        
        # Build list of all possible (day, start_slot) combinations for labs
        all_combinations = []
        for day in DAYS:
            for i in range(0, len(regular_slots) - LAB_DURATION + 1):
                start = regular_slots[i]
                seq = self._get_consecutive_slots(start, LAB_DURATION)
                if len(seq) == LAB_DURATION and all(s in regular_slots for s in seq):
                    all_combinations.append((day, start, seq))
        
        # Shuffle combinations for randomness
        random.shuffle(all_combinations)
        
        while len(scheduled_slots) < labs_per_week * LAB_DURATION and attempts < max_attempts:
            attempts += 1
            available_combos = [combo for combo in all_combinations if combo[0] not in used_days and combo[0] not in avoid_days]
            if not available_combos:
                available_combos = all_combinations
            if not available_combos:
                break
            
            day, start_slot, slots = random.choice(available_combos)
            
            # Check if all slots are available
            all_available = True
            for slot in slots:
                if not (self._is_time_slot_available_local(schedule, day, [slot]) and
                       self._is_time_slot_available_global(day, [slot], department, session, semester_id)):
                    all_available = False
                    break
            
            if all_available:
                for slot in slots:
                    self._mark_slots_busy_local(schedule, day, [slot], course_code, 'Lab')
                    self._mark_slots_busy_global(day, [slot], department, session, semester_id)
                    scheduled_slots.append((day, slot))
                used_days.add(day)
                avoid_days.add(day)
                # Remove this combination from future consideration
                all_combinations = [c for c in all_combinations if c != (day, start_slot, slots)]
        
        scheduled_count = len(scheduled_slots) // LAB_DURATION
        if scheduled_count < labs_per_week:
            print(f"      WARNING: {course_code} - Only scheduled {scheduled_count}/{labs_per_week} labs (attempts: {attempts})")
        return scheduled_slots

    def _schedule_course(self, schedule, course, department, session, semester_id):
        """Schedule all components of a course based on LTPSC."""
        course_code = course['Course Code']
        lectures_per_week = course['Lectures_Per_Week']
        tutorials_per_week = course['Tutorials_Per_Week']
        labs_per_week = course['Labs_Per_Week']
        
        # Robust elective detection: check multiple column name variants + pattern overrides
        elective_flag = False
        for colname in ['Elective (Yes/No)', 'Elective', 'Is Elective', 'Is_Elective']:
            if colname in course.index:
                elective_flag = str(course.get(colname, '')).upper() == 'YES'
                if elective_flag:
                    break
        # Pattern overrides: force ELEC as elective; force HSS as not elective
        course_code_str = str(course.get('Course Code', '')).upper()
        course_name_str = str(course.get('Course Name', '')).upper()
        if 'ELEC' in course_code_str or 'ELEC' in course_name_str:
            elective_flag = True
        if 'HSS' in course_code_str or 'HSS' in course_name_str:
            elective_flag = False
        # Determine if HSS
        is_hss = ('HSS' in course_code_str) or ('HSS' in course_name_str)
        
        # IMPORTANT: Follow LTPSC strictly for ALL courses (including electives)
        # Do not override parsed weekly counts; use values from Excel (via parse_ltpsc)
        elective_status = " [ELECTIVE]" if elective_flag else ""
        print(f"      Scheduling {course_code}{elective_status}: L={lectures_per_week}, T={tutorials_per_week}, P={labs_per_week}")
        skip_room_allocation = elective_flag or is_hss
        
        # Track used days for this course scoped to semester+department+session
        scoped_key = (semester_id, department, session, course_code)
        if scoped_key not in self.scheduled_courses:
            self.scheduled_courses[scoped_key] = set()

        success_counts = {'lectures': 0, 'tutorials': 0, 'labs': 0}
        scheduled_slots = []

        # Before normal scheduling, if combined class slots already exist for this course, apply them first
        def _get_combined_group(dept_label):
            if dept_label in {'CSE-A', 'CSE-B'}:
                return 'CSE'
            if dept_label in {'DSAI', 'ECE'}:
                return 'DSAI_ECE'
            return None
        group_key = _get_combined_group(department)

        def _apply_existing_combined(component_name, duration, required_count):
            if required_count == 0:
                return 0, []
            course_key = str(course_code).strip()
            # Prefer global combined slots (shared across semesters), else semester-specific
            global_key = ('GLOBAL', group_key, course_key, component_name) if group_key else None
            sem_key = (semester_id, course_key, component_name)
            assigned = []
            if global_key is not None:
                assigned = self.global_combined_course_slots.get(global_key, [])
            if not assigned:
                assigned = self.semester_combined_course_slots.get(sem_key, [])
            applied = 0
            applied_slots = []
            for day, start_slot in assigned:
                slots = self._get_consecutive_slots(start_slot, duration)
                # Respect local/global availability to avoid overwriting
                if (len(slots) == duration and
                    self._is_time_slot_available_local(schedule, day, slots) and
                    self._is_time_slot_available_global(day, slots, department, session, semester_id)):
                    self._mark_slots_busy_local(schedule, day, slots, course_code, 'Lecture' if component_name=='Lecture' else ('Tutorial' if component_name=='Tutorial' else 'Lab'))
                    self._mark_slots_busy_global(day, slots, department, session, semester_id)
                    # Count each component instance, not each slot unit
                    applied += 1
                    applied_slots.extend([(day, s) for s in slots])
                    if applied >= required_count:
                        break
            return applied, applied_slots

        # Apply existing combined for labs, lectures, tutorials respectively
        applied_lab_count, applied_lab_slots = _apply_existing_combined('Lab', LAB_DURATION, labs_per_week if labs_per_week else 0)
        applied_lec_count, applied_lec_slots = _apply_existing_combined('Lecture', LECTURE_DURATION, lectures_per_week if lectures_per_week else 0)
        applied_tut_count, applied_tut_slots = _apply_existing_combined('Tutorial', TUTORIAL_DURATION, tutorials_per_week if tutorials_per_week else 0)

        labs_remaining = max(0, (labs_per_week or 0) - applied_lab_count)
        lectures_remaining = max(0, (lectures_per_week or 0) - applied_lec_count)
        tutorials_remaining = max(0, (tutorials_per_week or 0) - applied_tut_count)

        # Schedule components in priority order (labs first, then lectures, then tutorials), for remaining
        # Avoid scheduling multiple events of same course on the same day
        avoid_days = set(self.scheduled_courses.get(scoped_key, set()))

        if labs_per_week > 0:
            lab_slots = []
            if labs_remaining > 0:
                lab_slots = self._schedule_labs(schedule, course_code, labs_remaining, department, session, semester_id, avoid_days=avoid_days)
            success_counts['labs'] = applied_lab_count + (len(lab_slots) // LAB_DURATION)
            scheduled_slots.extend(applied_lab_slots + lab_slots)

        if lectures_per_week > 0:
            if elective_flag:
                print(f"      -> Scheduling ELECTIVE lectures for {course_code}: {lectures_per_week} per week")
                lec_slots = []
                if lectures_remaining > 0:
                    lec_slots = self._schedule_elective_classes(schedule, course_code, lectures_remaining, department, session, semester_id, avoid_days=avoid_days)
                success_counts['lectures'] = applied_lec_count + (len(lec_slots) // LECTURE_DURATION)
                if success_counts['lectures'] > 0:
                    print(f"      -> Successfully scheduled {success_counts['lectures']} elective lecture(s) for {course_code}")
                else:
                    print(f"      -> WARNING: Failed to schedule elective lectures for {course_code}")
            else:
                lec_slots = []
                if lectures_remaining > 0:
                    lec_slots = self._schedule_lectures(schedule, course_code, lectures_remaining, department, session, semester_id, avoid_days=avoid_days)
                success_counts['lectures'] = applied_lec_count + (len(lec_slots) // LECTURE_DURATION)
            scheduled_slots.extend(applied_lec_slots + lec_slots)
        elif elective_flag and lectures_per_week == 0:
            # Log elective even if it has no lectures (might have labs or tutorials)
            print(f"      -> NOTE: Elective {course_code} has 0 lectures (may have labs/tutorials only)")

        if tutorials_per_week > 0:
            tut_slots = []
            if tutorials_remaining > 0:
                tut_slots = self._schedule_tutorials(schedule, course_code, tutorials_remaining, department, session, semester_id, avoid_days=avoid_days)
            success_counts['tutorials'] = applied_tut_count + (len(tut_slots) // TUTORIAL_DURATION)
            scheduled_slots.extend(applied_tut_slots + tut_slots)

        fully_scheduled = (
            success_counts['lectures'] == lectures_per_week and
            success_counts['tutorials'] == tutorials_per_week and
            success_counts['labs'] == labs_per_week
        )
        partially_scheduled = any(x > 0 for x in success_counts.values())
        
        # Combined fallback disabled
        
        # After all scheduling (including combined), persist slot usage and scoped days
        key = f"sem_{semester_id}"
        slot_key = f"{department}_{session}"
        if key not in self.scheduled_slots:
            self.scheduled_slots[key] = {}
        if slot_key not in self.scheduled_slots[key]:
            self.scheduled_slots[key][slot_key] = set()
        for day, slot in scheduled_slots:
            self.scheduled_slots[key][slot_key].add((day, slot))
            self.scheduled_courses[scoped_key].add(day)
        # Assign rooms: labs in lab rooms; lectures/tutorials in non-lab rooms
        # Extract component meetings from local variables if present
        lec_meetings = []
        tut_meetings = []
        lab_meetings = []
        try:
            lec_meetings = (applied_lec_slots or []) + (lec_slots or [])
        except Exception:
            pass
        try:
            tut_meetings = (applied_tut_slots or []) + (tut_slots or [])
        except Exception:
            pass
        try:
            lab_meetings = (applied_lab_slots or []) + (lab_slots or [])
        except Exception:
            pass
        course_code_str = str(course_code).strip()
        alloc_key = (semester_id, department, session, course_code_str)

        # Allocate lab rooms first
        if lab_meetings:
            if skip_room_allocation:
                if hasattr(self, 'assigned_lab_rooms') and alloc_key in self.assigned_lab_rooms:
                    self.assigned_lab_rooms.pop(alloc_key, None)
            else:
                self._allocate_lab_room_for_course(semester_id, department, session, course_code, lab_meetings)
        # Allocate lecture/tutorial rooms (can be shared)
        lt_meetings = []
        if lec_meetings:
            lt_meetings.extend(lec_meetings)
        if tut_meetings:
            lt_meetings.extend(tut_meetings)
        if lt_meetings:
            if skip_room_allocation:
                if hasattr(self, 'assigned_rooms') and alloc_key in self.assigned_rooms:
                    self.assigned_rooms.pop(alloc_key, None)
            else:
                self._allocate_room_for_course(semester_id, department, session, course_code, lt_meetings)
        # Store actual allocated counts now that we are final on success_counts
        self.actual_allocations[alloc_key] = {
            'lectures': success_counts['lectures'],
            'tutorials': success_counts['tutorials'],
            'labs': success_counts['labs'],
            'room': self.assigned_rooms.get(alloc_key, ''),
            'lab_room': getattr(self, 'assigned_lab_rooms', {}).get(alloc_key, '')
        }
        # Report issues
        if not fully_scheduled:
            missing_parts = []
            if success_counts['lectures'] < lectures_per_week:
                missing_parts.append(f"{lectures_per_week - success_counts['lectures']}/{lectures_per_week} lectures")
            if success_counts['tutorials'] < tutorials_per_week:
                missing_parts.append(f"{tutorials_per_week - success_counts['tutorials']}/{tutorials_per_week} tutorials")
            if success_counts['labs'] < labs_per_week:
                missing_parts.append(f"{labs_per_week - success_counts['labs']}/{labs_per_week} labs")
            if missing_parts:
                print(f"      ERROR: {course_code} - Failed to schedule: {', '.join(missing_parts)}")
        
        return fully_scheduled, partially_scheduled, success_counts

    def _assign_combined_slots(self, course_code, semester_id, rem_lec, rem_tut, rem_lab, group_key):
        """Assign combined (240-seater) slots for remaining counts and store for all departments in the same pairing group.
        Ensures only one combined class occurs at a time per semester per group."""
        semester_key = f"sem_{semester_id}"
        capacity_key = (semester_key, group_key)
        if capacity_key not in self.semester_combined_capacity:
            self.semester_combined_capacity[capacity_key] = set()
        added = []
        added_lec = added_tut = added_lab = 0
        # If global combined slots already exist for this course, reuse them
        course_key = str(course_code).strip()
        reused = False
        global_lec_key = ('GLOBAL', group_key, course_key, 'Lecture')
        global_tut_key = ('GLOBAL', group_key, course_key, 'Tutorial')
        global_lab_key = ('GLOBAL', group_key, course_key, 'Lab')
        if ((rem_lec and global_lec_key in self.global_combined_course_slots) or
            (rem_tut and global_tut_key in self.global_combined_course_slots) or
            (rem_lab and global_lab_key in self.global_combined_course_slots)):
            reused = True
            def reuse(comp_name, needed, duration, gkey):
                nonlocal added, added_lec, added_tut, added_lab
                if not needed:
                    return 0
                assigned = self.global_combined_course_slots.get(gkey, [])
                placed = 0
                for day, start_slot in assigned:
                    # Store under semester-specific map as well (for exporter lookups)
                    self.semester_combined_course_slots.setdefault((semester_id, course_key, comp_name), []).append((day, start_slot))
                    added.append((comp_name, day, start_slot))
                    placed += 1
                    if placed >= needed:
                        break
                return placed
            if rem_lab:
                added_lab += reuse('Lab', rem_lab, LAB_DURATION, global_lab_key)
            if rem_lec:
                added_lec += reuse('Lecture', rem_lec, LECTURE_DURATION, global_lec_key)
            if rem_tut:
                added_tut += reuse('Tutorial', rem_tut, TUTORIAL_DURATION, global_tut_key)
            return added_lec, added_tut, added_lab, added
        regular_slots = [slot for slot in TEACHING_SLOTS if slot not in (MINOR_SLOTS + LUNCH_SLOTS)]
        # Helper to place N instances of a component with given duration
        def place(component_name, n, duration):
            nonlocal added
            placed = 0
            attempts = 0
            max_attempts = 1500
            while placed < n and attempts < max_attempts:
                attempts += 1
                day = random.choice(DAYS)
                start_slot = random.choice(regular_slots) if regular_slots else None
                if start_slot is None:
                    break
                slots = self._get_consecutive_slots(start_slot, duration)
                slots = [s for s in slots if s in regular_slots]
                if len(slots) != duration:
                    continue
                # Combined capacity: none of the unit slots may be used by another combined class
                capacity_free = all((day, s) not in self.semester_combined_capacity[capacity_key] for s in slots)
                if not capacity_free:
                    continue
                # Reserve capacity
                for s in slots:
                    self.semester_combined_capacity[capacity_key].add((day, s))
                # Save under course-specific combined slots
                key = (semester_id, course_key, component_name)
                self.semester_combined_course_slots.setdefault(key, []).append((day, start_slot))
                # Also store globally so other semesters reuse same timing
                gkey = ('GLOBAL', group_key, course_key, component_name)
                self.global_combined_course_slots.setdefault(gkey, []).append((day, start_slot))
                added.append((component_name, day, start_slot))
                placed += 1
            return placed
        if rem_lab > 0:
            added_lab = place('Lab', rem_lab, LAB_DURATION)
        if rem_lec > 0:
            added_lec = place('Lecture', rem_lec, LECTURE_DURATION)
        if rem_tut > 0:
            added_tut = place('Tutorial', rem_tut, TUTORIAL_DURATION)
        return added_lec, added_tut, added_lab, added

    def generate_department_schedule(self, semester_id, department, session_type):
        """Generate schedule for a specific department and session."""
        print(f"    Generating {department} - {session_type}")
        
        # Initialize global slots for this semester if not exists
        semester_key = f"sem_{semester_id}"
        if semester_key not in self.semester_global_slots:
            self.semester_global_slots[semester_key] = {}
        
        # Get ALL courses for the semester (including duplicated CSE-B courses)
        sem_courses_all = ExcelLoader.get_semester_courses(self.dfs, semester_id)
        if sem_courses_all.empty:
            print(f"    ERROR: No courses found for semester {semester_id} in the input file")
            return self._initialize_schedule()
        
        # Enforce LTPSC strictly across the semester, get only LTPSC-valid courses (minors kept)
        sem_courses_parsed = ExcelLoader.parse_ltpsc(sem_courses_all)
        if sem_courses_parsed.empty:
            print(f"    ERROR: No LTPSC-valid courses available for semester {semester_id}")
            return self._initialize_schedule()
        
        # Filter courses for this department (exact match)
        if 'Department' in sem_courses_parsed.columns:
            dept_mask = sem_courses_parsed['Department'].astype(str).str.contains(f"^{department}$", na=False, regex=True)
            dept_courses = sem_courses_parsed[dept_mask].copy()
            print(f"    Found {len(dept_courses)} LTPSC-valid courses for {department}")
        else:
            dept_courses = sem_courses_parsed.copy()
            print(f"    Department column not found; using all {len(dept_courses)} semester courses")
        
        if dept_courses.empty:
            print(f"    WARNING: No courses found for {department} in semester {semester_id}")
            return self._initialize_schedule()
        
        # Divide courses by session using the parsed semester-wide view for shared-course detection
        pre_mid_courses, post_mid_courses = ExcelLoader.divide_courses_by_session(dept_courses, department, all_sem_courses=sem_courses_parsed)
        
        # Select courses for this session
        if session_type == PRE_MID:
            session_courses = pre_mid_courses
        else:
            session_courses = post_mid_courses
        
        if session_courses.empty:
            print(f"    WARNING: No courses for {department} in {session_type} session")
            return self._initialize_schedule()
        
        # Count electives in this session
        elective_count = 0
        elective_col_session = None
        for colname in ['Elective (Yes/No)', 'Elective', 'Is Elective', 'Is_Elective']:
            if colname in session_courses.columns:
                elective_col_session = colname
                break
        if elective_col_session is not None:
            elective_mask = session_courses[elective_col_session].astype(str).str.upper() == 'YES'
            elective_count = elective_mask.sum()
            if elective_count > 0:
                elective_codes_list = session_courses[elective_mask]['Course Code'].dropna().astype(str).tolist()
                print(f"    Found {elective_count} ELECTIVE course(s) in {department} {session_type}: {', '.join(elective_codes_list)}")
        
        # Initialize schedule
        schedule = self._initialize_schedule()
        
        # Schedule Minor subject first (ONLY in MINOR_SLOTS)
        self._schedule_minor_classes(schedule, department, session_type, semester_id)
        
        # Track expected vs scheduled courses
        expected_courses = set(session_courses['Course Code'].dropna().astype(str)) if 'Course Code' in session_courses.columns else set()
        
        # Schedule all department courses (in regular slots)
        scheduled_courses = set()
        fully_scheduled_courses = []
        partially_scheduled_courses = []
        failed_courses = []
        
        # Reorder courses: schedule those with existing global/semester combined slots first
        def has_existing_combined(course_row):
            code = str(course_row.get('Course Code', '')).strip()
            if not code:
                return 0
            # Check global or semester-specific combined slots for any component
            for comp in ['Lecture', 'Tutorial', 'Lab']:
                if ('GLOBAL', code, comp) in self.global_combined_course_slots and self.global_combined_course_slots[('GLOBAL', code, comp)]:
                    return 1
                if (semester_id, code, comp) in self.semester_combined_course_slots and self.semester_combined_course_slots[(semester_id, code, comp)]:
                    return 1
            return 0
        try:
            session_courses = session_courses.sort_values(
                by=session_courses.apply(has_existing_combined, axis=1),
                ascending=False
            )
        except Exception:
            pass
        
        for _, course in session_courses.iterrows():
            course_code = str(course.get('Course Code', ''))
            if not course_code or course_code == 'nan':
                continue
                
            fully_scheduled, partially_scheduled, counts = self._schedule_course(schedule, course, department, session_type, semester_id)
            scheduled_courses.add(course_code)
            
            if fully_scheduled:
                fully_scheduled_courses.append(course_code)
            elif partially_scheduled:
                partially_scheduled_courses.append(course_code)
            else:
                failed_courses.append(course_code)
        
        # Validate and report course allocation
        missing_courses = expected_courses - scheduled_courses
        if missing_courses:
            print(f"    ERROR: {len(missing_courses)} courses were not scheduled: {', '.join(sorted(missing_courses))}")
        
        # Print summary
        print(f"    Schedule Summary for {department} - {session_type}:")
        print(f"      Total courses expected: {len(expected_courses)}")
        print(f"      Fully scheduled: {len(fully_scheduled_courses)}")
        if partially_scheduled_courses:
            print(f"      Partially scheduled: {len(partially_scheduled_courses)} ({', '.join(partially_scheduled_courses)})")
        if failed_courses:
            print(f"      Failed to schedule: {len(failed_courses)} ({', '.join(failed_courses)})")
        
        print(f"    Completed {department} - {session_type}")
        
        # Store schedule in the schedule object for later retrieval
        schedule._department = department
        schedule._session = session_type
        schedule._semester_id = semester_id
        
        return schedule
    
    def get_actual_allocations(self, semester_id, department, session_type, course_code):
        """Get actual allocated counts for a specific course."""
        alloc_key = (semester_id, department, session_type, course_code)
        return self.actual_allocations.get(alloc_key, {'lectures': 0, 'tutorials': 0, 'labs': 0, 'room': '', 'lab_room': ''})

    def _get_course_enrollment(self, semester_id, department, course_code):
        """Registered students for this course in this department/semester."""
        try:
            df = self.dfs.get('course')
            if df is None or df.empty:
                return 0
            sub = df.copy()
            if 'Semester' in sub.columns:
                sub['Semester'] = pd.to_numeric(sub['Semester'], errors='coerce')
                sub = sub[sub['Semester'] == semester_id]
            if 'Department' in sub.columns:
                sub = sub[sub['Department'].astype(str).str.contains(f"^{department}$", na=False, regex=True)]
            sub = sub[sub['Course Code'].astype(str) == str(course_code)]
            if sub.empty:
                return 0
            for c in sub.columns:
                cl = str(c).lower()
                if any(k in cl for k in ['registered', 'students', 'strength', 'capacity']):
                    val = pd.to_numeric(sub.iloc[0][c], errors='coerce')
                    return int(val) if not pd.isna(val) else 0
            return 0
        except Exception:
            return 0

    def _allocate_room_for_course(self, semester_id, department, session, course_code, scheduled_slots):
        """Allocate a room using classroom_data; prefer one room for all meetings."""
        if not scheduled_slots or not self.classrooms:
            return
        semester_key = f"sem_{semester_id}"
        meetings = sorted(set(scheduled_slots))
        needed = self._get_course_enrollment(semester_id, department, course_code)
        # Try single room for all meetings
        chosen = None
        for room_name, capacity in self.classrooms:
            if capacity and needed and capacity < needed:
                continue
            ok = True
            for day, slot in meetings:
                used = self.room_occupancy.get((semester_key, day, slot), set())
                if room_name in used:
                    ok = False
                    break
            if ok:
                chosen = room_name
                break
        if chosen:
            for day, slot in meetings:
                occ_key = (semester_key, day, slot)
                self.room_occupancy.setdefault(occ_key, set()).add(chosen)
                # record booking for validation
                self.room_bookings.setdefault(occ_key, []).append(
                    (chosen, department, str(course_code).strip(), session)
                )
            alloc_key = (semester_id, department, session, str(course_code).strip())
            if not hasattr(self, 'assigned_rooms'):
                self.assigned_rooms = {}
            self.assigned_rooms[alloc_key] = chosen
            return
        # Otherwise allocate per meeting, final label VARIES
        for day, slot in meetings:
            occ_key = (semester_key, day, slot)
            used = self.room_occupancy.get(occ_key, set())
            allocated = None
            for room_name, capacity in self.classrooms:
                if capacity and needed and capacity < needed:
                    continue
                if room_name not in used:
                    allocated = room_name
                    break
            if allocated is None:
                # fallback ignoring capacity
                for room_name, _ in self.classrooms:
                    if room_name not in used:
                        allocated = room_name
                        break
            if allocated:
                self.room_occupancy.setdefault(occ_key, set()).add(allocated)
                # record booking for validation
                self.room_bookings.setdefault(occ_key, []).append(
                    (allocated, department, str(course_code).strip(), session)
                )
        alloc_key = (semester_id, department, session, str(course_code).strip())
        if not hasattr(self, 'assigned_rooms'):
            self.assigned_rooms = {}
        self.assigned_rooms[alloc_key] = 'VARIES'

    def _are_side_by_side(self, room_a, room_b):
        """Heuristic side-by-side: same alpha prefix and numeric suffix differs by 1."""
        import re
        def split_room(r):
            r = str(r).replace(' ', '')
            m = re.match(r'^([A-Za-z]+)(\d+)$', r)
            if m:
                return m.group(1), int(m.group(2))
            prefix = ''.join([c for c in r if c.isalpha()])
            digits = ''.join([c for c in r if c.isdigit()])
            try:
                return prefix, int(digits) if digits else None
            except Exception:
                return prefix, None
        pa, na = split_room(room_a)
        pb, nb = split_room(room_b)
        return bool(pa and pa == pb and na is not None and nb is not None and abs(na - nb) == 1)

    def _pick_lab_pool_for_department(self, department):
        d = str(department).upper()
        # CSE/DSAI -> software labs; ECE -> hardware labs
        if any(k in d for k in ['CSE-A', 'CSE-B', 'CSE', 'DSAI', 'DS']):
            if self.software_lab_rooms:
                return self.software_lab_rooms
        if 'ECE' in d and self.hardware_lab_rooms:
            return self.hardware_lab_rooms
        return self.lab_rooms if self.lab_rooms else self.classrooms

    def _allocate_lab_room_for_course(self, semester_id, department, session, course_code, lab_meetings):
        """Allocate labs: pick two side-by-side labs per meeting (40+40=80), specialized by department."""
        if not lab_meetings or not (self.lab_rooms or self.classrooms):
            return
        semester_key = f"sem_{semester_id}"
        meetings = sorted(set(lab_meetings))
        needed = self._get_course_enrollment(semester_id, department, course_code)
        pool = list(self._pick_lab_pool_for_department(department))
        # Build candidate pairs (prefer side-by-side)
        side_pairs = []
        any_pairs = []
        for i in range(len(pool)):
            r1, _ = pool[i]
            for j in range(i+1, len(pool)):
                r2, _ = pool[j]
                if self._are_side_by_side(r1, r2):
                    side_pairs.append((r1, r2))
                else:
                    any_pairs.append((r1, r2))
        def pairs_free_for_all(pair):
            a, b = pair
            for day, slot in meetings:
                used = self.room_occupancy.get((semester_key, day, slot), set())
                if a in used or b in used:
                    return False
            return True
        chosen_pair = None
        for p in side_pairs:
            if pairs_free_for_all(p):
                chosen_pair = p
                break
        if chosen_pair is None:
            for p in any_pairs:
                if pairs_free_for_all(p):
                    chosen_pair = p
                    break
        if chosen_pair:
            a, b = chosen_pair
            for day, slot in meetings:
                occ_key = (semester_key, day, slot)
                self.room_occupancy.setdefault(occ_key, set()).update([a, b])
                self.room_bookings.setdefault(occ_key, []).append((a, department, str(course_code).strip(), session))
                self.room_bookings.setdefault(occ_key, []).append((b, department, str(course_code).strip(), session))
            if not hasattr(self, 'assigned_lab_rooms'):
                self.assigned_lab_rooms = {}
            alloc_key = (semester_id, department, session, str(course_code).strip())
            self.assigned_lab_rooms[alloc_key] = f"{a} + {b}"
            return
        # Per-meeting allocation (pair may vary)
        used_labels = set()
        for day, slot in meetings:
            occ_key = (semester_key, day, slot)
            used = self.room_occupancy.get(occ_key, set())
            assigned_pair = None
            # side-by-side first
            for a, b in side_pairs:
                if a not in used and b not in used:
                    assigned_pair = (a, b)
                    break
            if assigned_pair is None:
                free_rooms = [r for r, _ in pool if r not in used]
                if len(free_rooms) >= 2:
                    assigned_pair = (free_rooms[0], free_rooms[1])
            if assigned_pair:
                a, b = assigned_pair
                self.room_occupancy.setdefault(occ_key, set()).update([a, b])
                self.room_bookings.setdefault(occ_key, []).append((a, department, str(course_code).strip(), session))
                self.room_bookings.setdefault(occ_key, []).append((b, department, str(course_code).strip(), session))
                used_labels.add(f"{a} + {b}")
        if not hasattr(self, 'assigned_lab_rooms'):
            self.assigned_lab_rooms = {}
        alloc_key = (semester_id, department, session, str(course_code).strip())
        self.assigned_lab_rooms[alloc_key] = next(iter(used_labels)) if len(used_labels) == 1 else 'VARIES'
    def validate_room_conflicts(self):
        """Return a list of detected room conflicts: entries where the same room is booked more than once in the same (semester, day, slot)."""
        conflicts = []
        for (sem_key, day, slot), bookings in self.room_bookings.items():
            room_to_entries = {}
            for room_name, dept, course_code, session in bookings:
                room_to_entries.setdefault(room_name, []).append((dept, course_code, session))
            for room_name, entries in room_to_entries.items():
                if len(entries) > 1:
                    conflicts.append({
                        'semester': sem_key,
                        'day': day,
                        'slot': slot,
                        'room': room_name,
                        'entries': entries
                    })
        return conflicts