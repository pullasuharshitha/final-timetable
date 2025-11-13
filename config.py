"""Configuration settings for the timetable generator."""

# Directory paths
INPUT_DIR = 'C:/Users/smile/OneDrive/Desktop/tt/sdtt_inputs'
OUTPUT_DIR = 'C:/Users/smile/OneDrive/Desktop/tt/output'

# Required Excel input files
REQUIRED_FILES = [
    'course_data.xlsx',
    'classroom_data.xlsx'
]

# Departments
DEPARTMENTS = ['CSE-A', 'CSE-B', 'DSAI', 'ECE']

# Target semesters
TARGET_SEMESTERS = [1, 3, 5]

# Session types
PRE_MID = 'Pre-Mid'
POST_MID = 'Post-Mid'

# Minor subject
MINOR_SUBJECT = "Minor"

# Time scheduling configuration
DAYS = ['MON', 'TUE', 'WED', 'THU', 'FRI']

# Teaching time slots in 30-minute increments (07:30 - 17:30)
TEACHING_SLOTS = [
    '07:30-08:00', '08:00-08:30', '08:30-09:00',
    '09:00-09:30', '09:30-10:00', '10:00-10:30',
    '10:30-11:00', '11:00-11:30', '11:30-12:00',
    '12:00-12:30', '12:30-13:00',
    '13:00-13:30', '13:30-14:00',  # Lunch slots
    '14:00-14:30', '14:30-15:00',
    '15:00-15:30', '15:30-16:00',
    '16:00-16:30', '16:30-17:00',
    '17:00-17:30'
]

# Lunch and Minor slot definitions
LUNCH_SLOTS = ['13:00-13:30', '13:30-14:00']
MINOR_SLOTS = ['07:30-08:00', '08:00-08:30']  # 07:30-08:30 represented as two 30-min slots

# Class durations (counted in 30-minute slots)
LECTURE_DURATION = 3    # 1.5 hours = 3 slots
TUTORIAL_DURATION = 2   # 1 hour = 2 slots
LAB_DURATION = 4        # 2 hours = 4 slots (consecutive slots)
MINOR_DURATION = 2      # 1 hour = 2 slots

# Weekly frequency
MINOR_CLASSES_PER_WEEK = 2