import numpy as np
import pandas as pd

def perform_idle_time_analysis(df, book):
    # Define your desired time range
    start_time = pd.to_datetime('12:00 AM', format='%I:%M %p')
    end_time = pd.to_datetime('6:30 AM', format='%I:%M %p')

    # Filter rows based on the specified time range
    df = df.loc[(df['DateTime'] >= start_time) & (df['DateTime'] <= end_time)].copy()

    # Filter rows to include only specific pick types
    valid_pick_types = ['REPLENISHMENT PICK', 'PUTWALL PICKING', 'REGULAR PICK','SINGLE PICK', 'QUICK MOVE']
    df = df[df['Action'].isin(valid_pick_types)]

    # Calculate the time difference between each action for each user
    df.loc[:, 'TimeDiff'] = df.sort_values(['UserID', 'DateTime']).groupby('UserID')['DateTime'].diff()

    # For the first action of each user, calculate the time difference from 8:00 PM
    first_action_diff = (df.loc[df.groupby('UserID')['DateTime'].idxmin(), 'DateTime'] - start_time)
    
    # Replace NaT with Timedelta 0
    first_action_diff = first_action_diff.where(first_action_diff > pd.Timedelta(minutes=14), pd.Timedelta(minutes=0))
    
    df.loc[df.groupby('UserID')['DateTime'].idxmin(), 'TimeDiff'] = first_action_diff

    # Convert the 'TimeDiff' column to timedelta64
    df['TimeDiff'] = pd.to_timedelta(df['TimeDiff'])

    # Convert the time difference to minutes
    df.loc[:, 'TimeDiff'] = df['TimeDiff'].dt.total_seconds() / 60

    # Consider a user to be idle if the time difference is more than 14 minutes
    df.loc[:, 'Idle'] = np.where(df['TimeDiff'] > 14, 1, 0)

    # Calculate the total idle time for each user
    idle_time_per_user = df[df['Idle'] == 1].groupby('UserID').agg(
        TotalIdleTime=('TimeDiff', 'sum')
    ).reset_index()

    # Create the "IDLE TIME" sheet if it doesn't exist
    if 'IDLE TIME' not in book.sheetnames:
        idle_time_sheet = book.create_sheet('IDLE TIME')
    else:
        idle_time_sheet = book['IDLE TIME']

    # Write the header row
    header_row = ['UserID', 'TotalIdleTime']
    idle_time_sheet.append(header_row)

    # Convert the DataFrame to a list of lists for writing to Excel
    idle_time_data = idle_time_per_user[['UserID', 'TotalIdleTime']].values.tolist()

    # Write the data to the Excel sheet
    for row_data in idle_time_data:
        idle_time_sheet.append(row_data)

    print("IDLE TIME analysis completed.")
