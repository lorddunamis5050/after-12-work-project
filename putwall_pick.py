import openpyxl
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows

def perform_putwall_pick_analysis(df, book):
    # Define your desired time range for PUTWALL PICKING
    start_time_putwall = pd.to_datetime('12:00 AM', format='%I:%M %p')
    end_time_putwall = pd.to_datetime('6:30 AM', format='%I:%M %p')

    # Initialize a DataFrame to store PUTWALL PICKING data per user
    putwall_picking_per_user = pd.DataFrame(columns=['UserID', 'PutwallPickingQuantity'])

    # Function to modify the 'Action' column based on 'BinLabel' for PUTWALL PICKING
    def modify_action_putwall(row):
        action = row['Action']
        bin_label = row['BinLabel']
        
        if action == 'PICKLINE' and bin_label.startswith('MW'):
            return 'PUTWALL PICKING'
        
        return action

    # Apply the function to the DataFrame for PUTWALL PICKING
    df['Action'] = df.apply(modify_action_putwall, axis=1)

  # Filter rows based on the specified time range for PUTWALL PICKING
    filtered_df_putwall = df[(df['DateTime'] >= start_time_putwall) & (df['DateTime'] <= end_time_putwall)]

    # Find the highest time worked by any user within the time range
    highest_hours_worked = (filtered_df_putwall.groupby('UserID')['DateTime']
                             .agg(lambda x: (x.max() - x.min()).total_seconds() / 3600)
                             .max())

    # Group by "UserID" and calculate total Units picked
    putwall_picking_per_user = filtered_df_putwall[filtered_df_putwall['Action'] == 'PUTWALL PICKING'].groupby('UserID').agg(
        PutwallPickingQuantity=('Quantity', 'sum')
    ).reset_index()

    # Calculate UPH for each user, using the highest time as the denominator
    putwall_picking_per_user['UPH'] = putwall_picking_per_user['PutwallPickingQuantity'] / highest_hours_worked

        # Convert both "PutwallPickingQuantity" and "UPH" values to their absolute values
    putwall_picking_per_user['PutwallPickingQuantity'] = abs(putwall_picking_per_user['PutwallPickingQuantity'])
    putwall_picking_per_user['UPH'] = abs(putwall_picking_per_user['UPH'])


    # Create the "PUTWALL PICKING" sheet if it doesn't exist
    if 'PUTWALL PICKING' not in book.sheetnames:
        putwall_picking_sheet = book.create_sheet('PUTWALL PICKING')
    else:
        putwall_picking_sheet = book['PUTWALL PICKING']

    # Write the header row
    header_row = ['UserID', 'PutwallPickingQuantity', 'UPH']
    putwall_picking_sheet.append(header_row)

    # Convert the DataFrame to a list of lists for writing to Excel
    putwall_picking_data = putwall_picking_per_user[['UserID', 'PutwallPickingQuantity', 'UPH']].values.tolist()


    # Write the data to the Excel sheet
    for row_data in putwall_picking_data:
        putwall_picking_sheet.append(row_data)

    print("PUTWALL PICKING analysis completed.")
    print(highest_hours_worked)
