import openpyxl
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows


def peform_quick_move_analysis(df, book):
    # Define your desired time range for quick move
    start_time_quick_move = pd.to_datetime('12:00 AM', format='%I:%M %p')
    end_time_quick_move = pd.to_datetime('6:30 AM', format='%I:%M %p')

        # Initialize a DataFrame to store PUTWALL PICKING data per user
    quick_move_per_user = pd.DataFrame(columns=['UserID', 'QuickMoveQuantity'])

    # Function to modify the 'Action' column based on 'BinLabel' for QUICK MOVE
    def modify_action_quickmove(row):
        action = row['Action']
        bin_label = row['BinLabel']
        
        if action == 'MOVE-IN' and bin_label.startswith('MW'):
            return 'QUICK MOVE'
        
        return action
    
    # Apply the function to the DataFrame for Quick Move
    df['Action'] = df.apply(modify_action_quickmove, axis = 1)

    # Filter rows based on the specified time range for PUTWALL PICKING
    filtered_df_quick_move = df[(df['DateTime'] >= start_time_quick_move) & (df['DateTime'] <= end_time_quick_move)]

    # Find the highest time worked by any user within the time range
    highest_hours_worked = (filtered_df_quick_move.groupby('UserID')['DateTime']
                             .agg(lambda x: (x.max() - x.min()).total_seconds() / 3600)
                             .max())
    
        # Group by "UserID" and calculate total Units picked
    quick_move_per_user = filtered_df_quick_move[filtered_df_quick_move['Action'] == 'QUICK MOVE'].groupby('UserID').agg(
        QuickMoveQuantity=('Quantity', 'sum')
    ).reset_index()

    # Calculate UPH for each user, using the highest time as the denominator
    quick_move_per_user['UPH'] = quick_move_per_user['QuickMoveQuantity'] / highest_hours_worked
           
            # Convert both "QuickMoveQuantity" and "UPH" values to their absolute values
    quick_move_per_user['QuickMoveQuantity'] = abs(quick_move_per_user['QuickMoveQuantity'])
    quick_move_per_user['UPH'] = abs(quick_move_per_user['UPH'])


    # Create the "PUTWALL PICKING" sheet if it doesn't exist
    if 'QUICK MOVE' not in book.sheetnames:
        quick_move_sheet = book.create_sheet('QUICK MOVE')
    else:
        quick_move_sheet = book['QUICK MOVE']

            # Write the header row
    header_row = ['UserID', 'QuickMoveQuantity', 'UPH']
    quick_move_sheet.append(header_row)

        # Convert the DataFrame to a list of lists for writing to Excel
    quick_move_data = quick_move_per_user[['UserID', 'QuickMoveQuantity', 'UPH']].values.tolist()

     # Write the data to the Excel sheet
    for row_data in quick_move_data:
        quick_move_sheet.append(row_data)

    print("QUICK MOVE analysis completed.")
    print(highest_hours_worked)








