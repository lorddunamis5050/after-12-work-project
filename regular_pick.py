import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

def perform_regular_pick_analysis(df , book):
    # Define your desired time range for REGULAR PICK
    START_TIME_REGULAR = pd.to_datetime('12:00 AM', format='%I:%M %p')
    END_TIME_REGULAR = pd.to_datetime('6:30 AM', format='%I:%M %p')

    # Initialize DataFrames to store regular pick data
    regular_pick_per_user = pd.DataFrame(columns=['UserID', 'RegularPickQuantity'])

    # Function to modify the 'Action' column based on 'BinLabel' for REGULAR PICK
    def modify_action_regular(row):
        action = row['Action']
        bin_label = row['BinLabel']
        packslip = row['Packslip']

        if action == 'PICKLINE':
            if bin_label.startswith(('1H', '1G', '2E', '2H', '3F', '3H', '3R', '2R', '1Y', '1C', '1D', '2D', '3D')) and not packslip.startswith('TR'):
                return 'REGULAR PICK'

        return action

    # Apply the function to the DataFrame for REGULAR PICK
    df['Action'] = df.apply(modify_action_regular, axis=1)

    # Filter rows based on the specified time range for REGULAR PICK
    filtered_df_regular = df[(df['DateTime'] >= START_TIME_REGULAR) & (df['DateTime'] <= END_TIME_REGULAR)]

     # Find the highest time worked by any user within the time range
    highest_hours_worked = (filtered_df_regular.groupby('UserID')['DateTime']
                             .agg(lambda x: (x.max() - x.min()).total_seconds() / 3600)
                             .max())
    
        # Group by "UserID" and calculate total Units picked
    regular_pick_per_user = filtered_df_regular[filtered_df_regular['Action'] == 'REGULAR PICK'].groupby('UserID').agg(
        RegularPickQuantity=('Quantity', 'sum')
    ).reset_index()

        # Calculate UPH for each user, using the highest time as the denominator
    regular_pick_per_user['UPH'] = regular_pick_per_user['RegularPickQuantity'] / highest_hours_worked


        # Convert both "PutwallPickingQuantity" and "UPH" values to their absolute values
    regular_pick_per_user['RegularPickQuantity'] = abs(regular_pick_per_user['RegularPickQuantity'])
    regular_pick_per_user['UPH'] = abs(regular_pick_per_user['UPH'])




        # Create the "REGULAR PICKING" sheet if it doesn't exist
    if 'REGULAR PICK' not in book.sheetnames:
        regular_pick_sheet = book.create_sheet('REGULAR PICK')
    else:
        regular_pick_sheet = book['REGULAR PICK']

            # Write the header row
    header_row = ['UserID', 'RegularPickQuantity', 'UPH']
    regular_pick_sheet.append(header_row)

        # Convert the DataFrame to a list of lists for writing to Excel
    regular_picking_data = regular_pick_per_user[['UserID', 'RegularPickQuantity', 'UPH']].values.tolist()

    # Write the data to the Excel sheet
    for row_data in regular_picking_data:
        regular_pick_sheet.append(row_data)


    print("REGULAR PICKING analysis completed.")
