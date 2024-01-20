import pandas as pd

def analyze_timecard(file_path):
    # reading spreadsheet to turn it into dataframe
    df = pd.read_excel(file_path)

    # making sure to set datetime objects
    df['Time'] = pd.to_datetime(df['Time'])
    df['Time Out'] = pd.to_datetime(df['Time Out'])

    # sort df by time for sequential analysis
    df.sort_values(by='Time', inplace=True)

    # applying given criteria to analyze
    consecutive_days = 7
    more_than_14_hours = pd.Timedelta(hours=14)

    for _, group in df.groupby('Employee Name'):
        consecutive_days_worked = group['Time'].diff().dt.days
        total_shift_hrs = (group['Time Out'] - group['Time']).dt.total_seconds() / 3600

        # a) employees who worked for 7 consecutive days
        if (consecutive_days_worked >= consecutive_days).any():
            print(f"\nEmployee: {group['Employee Name'].iloc[0]}")
            print(f"Position: {group['Position ID'].iloc[0]}")
            print("Worked for 7 consecutive days")

        # b) employees with <10 hours between shifts but >1 hour
        if ((group['Time'].diff() < pd.Timedelta(hours=10)) & (group['Time'].diff() > pd.Timedelta(hours=1))).any():
            print(f"\nEmployee: {group['Employee Name'].iloc[0]}")
            print(f"Position: {group['Position ID'].iloc[0]}")
            print("Less than 10 hours between shifts but more than 1 hour")

        # c) Employees who worked for more than 14 hours in a single shift
        if (total_shift_hrs.max() > more_than_14_hours.total_seconds() / 3600):
            print(f"\nEmployee: {group['Employee Name'].iloc[0]}")
            print(f"Position: {group['Position ID'].iloc[0]}")
            print("Worked for more than 14 hours in a single shift")

if __name__ == "__main__":
    file_path = "D:/Bluejay/Assignment_Timecard.xlsx"
    analyze_timecard(file_path)
