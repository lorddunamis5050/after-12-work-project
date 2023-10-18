import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

def perform_replenishment_pick_analysis(df, book):
    # Define your desired time range for PUTWALL PICKING
    START_TIME_REGULAR = pd.to_datetime('12:00 AM', format='%I:%M %p')
    END_TIME_REGULAR = pd.to_datetime('6:30 AM', format='%I:%M %p')

    # Initialize DataFrames to store replenishment pick data
    replenishment_pick_per_user = pd.DataFrame(columns=['UserID', 'ReplenishmentPickQuantity'])

    # Function to check for "REPLENISHMENT PICK"
    def modify_action(row):
        action = row['Action']
        packslip = row['Packslip']

        if action == 'REPLNISH' and len(str(packslip)) >= 7 and str(packslip)[6] == 'P':
            return 'REPLENISHMENT PICK'
        elif packslip == 'BIG':
            return 'REPLENISHMENT PICK'  # OR do something else for 'BIG'

        return action


# Apply the modification function to the 'Action' column
    df['Action'] = df.apply(modify_action, axis=1)

    filtered_df_regular = df[(df['DateTime'] >= START_TIME_REGULAR) & (df['DateTime'] <= END_TIME_REGULAR)]

    # Filter rows based on the criteria for "REPLENISHMENT PICK"
    # replenishment_pick_df = df[df['IsReplenishmentPick']]



         # Find the highest time worked by any user within the time range
    highest_hours_worked = (filtered_df_regular.groupby('UserID')['DateTime']
                             .agg(lambda x: (x.max() - x.min()).total_seconds() / 3600)
                             .max())

    # Count the sum of 'Quantity' for "REPLENISHMENT PICK" actions per user within the time range for PUTWALL PICKING
    replenishment_pick_per_user = filtered_df_regular[filtered_df_regular['Action'] == 'REPLENISHMENT PICK'].groupby('UserID').agg(
        ReplenishmentPickQuantity = ('Quantity', 'sum')
    ).reset_index()

            # Calculate UPH for each user, using the highest time as the denominator
    replenishment_pick_per_user['UPH'] = replenishment_pick_per_user['ReplenishmentPickQuantity'] / highest_hours_worked

            # Convert both "PutwallPickingQuantity" and "UPH" values to their absolute values
    replenishment_pick_per_user['ReplenishmentPickQuantity'] = abs(replenishment_pick_per_user['ReplenishmentPickQuantity'])
    replenishment_pick_per_user['UPH'] = abs(replenishment_pick_per_user['UPH'])


            # Create the "REPLENISHMENT  PICKING" sheet if it doesn't exist
    if 'REPLENISHMENT PICK' not in book.sheetnames:
        replenishment_pick_sheet = book.create_sheet('REPLENISHMENT PICK')
    else:
        replenishment_pick_sheet = book['REPLENISHMENT PICK']

                    # Write the header row
    header_row = ['UserID', 'ReplenishmentPickQuantity', 'UPH']
    replenishment_pick_sheet.append(header_row)

            # Convert the DataFrame to a list of lists for writing to Excel
    replenishment_picking_data = replenishment_pick_per_user[['UserID', 'ReplenishmentPickQuantity', 'UPH']].values.tolist()

        # Write the data to the Excel sheet
    for row_data in replenishment_picking_data:
        replenishment_pick_sheet.append(row_data)


    print("ByGroup PICKING analysis completed.")


