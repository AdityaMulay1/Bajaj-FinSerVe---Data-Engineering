import pandas as pd
import re

def run(path):
    # Load data from Excel file (attendance & students)
    attendance_df = pd.read_excel(path, sheet_name='attendance')
    students_df = pd.read_excel(path, sheet_name='students')

    # Convert attendance_date to proper date format
    attendance_df['attendance_date'] = pd.to_datetime(attendance_df['attendance_date'])

    # Sort these records by student ID and date for proper processing
    attendance_df = attendance_df.sort_values(by=['student_id', 'attendance_date'])

    # Identify absence streaks (grouping consecutive 'Absent' days)
    attendance_df['absent'] = (attendance_df['status'] == 'Absent').astype(int)
    attendance_df['streak_group'] = (attendance_df['absent'].diff().fillna(0) != 0).cumsum()

    # Get details of each absence streak (start date, end date, total days)
    streaks = attendance_df[attendance_df['status'] == 'Absent'].groupby(['student_id', 'streak_group']).agg(
        absence_start_date=('attendance_date', 'first'),
        absence_end_date=('attendance_date', 'last'),
        total_absent_days=('attendance_date', 'count')
    ).reset_index()

    # Keep only absence streaks longer than 3 days
    streaks = streaks[streaks['total_absent_days'] > 3]

    # Get the latest absence streak for each student
    latest_streaks = streaks.loc[streaks.groupby('student_id')['absence_end_date'].idxmax()]

    # Merge with student details (parent email, student name)
    final_df = latest_streaks.merge(students_df, on='student_id', how='left')

    # Function to check if an email is valid
    def is_valid_email(email):
        pattern = r'^[a-zA-Z_][a-zA-Z0-9_]*@[a-zA-Z]+\.(com)$'
        return bool(re.match(pattern, str(email)))

    # Apply email validation
    final_df['email'] = final_df['parent_email'].apply(lambda x: x if is_valid_email(x) else None)

    # Generate message for valid emails
    final_df['msg'] = final_df.apply(
        lambda row: f"Dear Parent, your child {row['student_name']} was absent from {row['absence_start_date']} to {row['absence_end_date']} for {row['total_absent_days']} days. Please ensure their attendance improves."
        if row['email'] else None,
        axis=1
    )

    # Select only the required columns
    return final_df[['student_id', 'absence_start_date', 'absence_end_date', 'total_absent_days', 'email', 'msg']]
