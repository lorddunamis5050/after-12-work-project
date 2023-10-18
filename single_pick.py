import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

def perform_single_pick_analysis(df , book):
    # Define your desired time range for PUTWALL PICKING
    START_TIME_PUTWALL = pd.to_datetime('12:00 AM', format='%I:%M %p')
    END_TIME_PUTWALL = pd.to_datetime('6:30 AM', format='%I:%M %p')

    # Initialize DataFrames to store single pick data
    single_pick_per_user = pd.DataFrame(columns=['UserID', 'SinglePickQuantity'])

    # Function to check for "SINGLE PICK"

    
    def modify_action_single_pick(row):
        action = row['Action']
        packslip = row['Packslip']
        datetime = row['DateTime']

        if action == 'REPLNISH' and len(str(packslip)) >= 7 and str(packslip)[6] == 'S' and datetime >= START_TIME_PUTWALL and datetime <= END_TIME_PUTWALL:
            return 'SINGLE PICK'

        return action


    # Apply the function to the DataFrame to identify "SINGLE PICK"
    df['Action'] = df.apply(modify_action_single_pick, axis=1)


    # Filter rows based on the criteria for "SINGLE PICK"
    

    filtered_df_single = df[(df['DateTime'] >= START_TIME_PUTWALL) & (df['DateTime'] <=END_TIME_PUTWALL)]

             # Find the highest time worked by any user within the time range
    highest_hours_worked = (filtered_df_single.groupby('UserID')['DateTime']
                             .agg(lambda x: (x.max() - x.min()).total_seconds() / 3600)
                             .max())


    # Count the sum of 'Quantity' for "SINGLE PICK" actions per user within the time range for PUTWALL PICKING
    single_pick_per_user = filtered_df_single[filtered_df_single['Action'] == 'SINGLE PICK'].groupby('UserID').agg(
        SinglePickQuantity=('Quantity', 'sum')
    ).reset_index()

    # Calculate UPH for each user, using the highest time as the denominator
    single_pick_per_user['UPH'] = single_pick_per_user['SinglePickQuantity'] / highest_hours_worked


    # Convert both "PutwallPickingQuantity" and "UPH" values to their absolute values
    single_pick_per_user['SinglePickQuantity'] = abs(single_pick_per_user['SinglePickQuantity'])
    single_pick_per_user['UPH'] = abs(single_pick_per_user['UPH'])




    # Create the "REPLENISHMENT  PICKING" sheet if it doesn't exist
    if 'SINGLE PICK' not in book.sheetnames:
        single_pick_sheet = book.create_sheet('SINGLE PICK')
    else:
        single_pick_sheet = book['SINGLE PICK']

    # Write the header row
    header_row = ['UserID', 'SinglePickQuantity', 'UPH']
    single_pick_sheet.append(header_row)

    # Convert the DataFrame to a list of lists for writing to Excel
    single_picking_data = single_pick_per_user[['UserID', 'SinglePickQuantity', 'UPH']].values.tolist()


        # Write the data to the Excel sheet
    for row_data in single_picking_data:
        single_pick_sheet.append(row_data)

    print("Effufilment PICKING analysis completed.")

