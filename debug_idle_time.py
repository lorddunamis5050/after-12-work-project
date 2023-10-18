import pandas as pd

# Read the data from the input Excel file
input_file = 'input_data.xlsx'
try:
    df = pd.read_excel(input_file)
except Exception as e:
    print(f"Error reading input Excel file: {str(e)}")
    exit(1)

# Check if the 'DateTime' column exists in the dataframe
if 'DateTime' not in df.columns:
    print("Error: 'DateTime' column not found in the input file.")
    exit(1)

# Convert 'DateTime' column to datetime objects
try:
    df['DateTime'] = pd.to_datetime(df['DateTime'], errors='coerce')
except Exception as e:
    print(f"Error converting 'DateTime' column to datetime: {str(e)}")
    exit(1)

# Sort the dataframe by 'UserID' and 'DateTime' columns
df = df.sort_values(by=['UserID', 'DateTime'])

# Initialize variables
idle_periods = []
current_user = None
idle_start_time = None

# Calculate idle periods and store them in a list
for _, row in df.iterrows():
    if current_user is None:
        current_user = row['UserID']
        idle_start_time = row['DateTime']
    elif current_user == row['UserID']:
        idle_end_time = row['DateTime']
        idle_time = idle_end_time - idle_start_time

        # Convert idle time to hours and minutes
        idle_hours, idle_minutes = divmod(idle_time.total_seconds() / 3600, 60)

        # Append the result to the list
        idle_periods.append({
            'UserID': current_user,
            'IdleTimeHours': idle_hours,
            'IdleTimeMinutes': idle_minutes
        })

        # Update the idle start time
        idle_start_time = row['DateTime']
    else:
        current_user = row['UserID']
        idle_start_time = row['DateTime']

# Create a new dataframe for idle periods
idle_df = pd.DataFrame(idle_periods)

# Summarize idle times for each user
summary_df = idle_df.groupby('UserID').agg({'IdleTimeHours': 'sum', 'IdleTimeMinutes': 'sum'}).reset_index()

# Write the idle periods and summary to an output Excel file
output_file = 'output_idle_periods.xlsx'
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    idle_df.to_excel(writer, sheet_name='Idle Periods', index=False)
    summary_df.to_excel(writer, sheet_name='Summary', index=False)

print("Idle periods and summary have been written to", output_file)
