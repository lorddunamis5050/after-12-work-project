import numpy as np
import pandas as pd
def perform_idle_time_analysis(df):
    # Define your desired time range
    start_time = pd.to_datetime('12:00 AM', format='%I:%M %p')
    end_time = pd.to_datetime('6:30 AM', format='%I:%M %p')



    # Filter rows based on the specified time range
    df = df.loc[(df['DateTime'] >= start_time) & (df['DateTime'] <= end_time)].copy()

    # Filter rows to include only specific pick types
    valid_pick_types = ['REPLENISHMENT PICK', 'PUTWALL PICKING', 'REGULAR PICK', 'SINGLE PICK', 'QUICK MOVE']
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

    # Create a dictionary to store detailed idle time for each user
    idle_time_details = {}

    for user_id, group in df.groupby('UserID'):
        idle_events = group[group['Idle'] == 1]
        time_diff = idle_events['TimeDiff'].tolist()
        idle_time_details[user_id] = time_diff

    return idle_time_details


