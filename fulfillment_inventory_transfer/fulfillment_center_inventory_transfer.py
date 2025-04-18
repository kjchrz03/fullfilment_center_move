import pandas as pd
import regex as re
import numpy as np

#rename the file from center
#delete Reference Key tab
#remove macros
#save as csv
center_inventory = pd.read_csv("current_inventory.csv")
fishbowl_inventory = pd.read_csv("Inventory Movement - center.csv")


# Find the exact column name by iterating through the column names
target_column = None
for column in center_inventory.columns:
    if 'ON HAND' in column and 'ACTL' in column:
        target_column = column
        break

# Check if the target column was found
if target_column is not None:
    # Select the desired columns and rename the 'target_column'
    center_inventory = center_inventory[['PROD#', 'PRODUCT DESCRIPTION', target_column,  'REFERENCE CODE']]
    center_inventory = center_inventory.rename(columns={target_column: 'at_center', 'PROD#': 'prod#', 'PRODUCT DESCRIPTION': 'description', 
                                                      'REFERENCE CODE': 'reference_code'})
else:
    print("The target column was not found.")


#reference codes at center
center_items = center_inventory[['reference_code']]
center_items = center_items['reference_code'].drop_duplicates()
center_items = pd.DataFrame(center_items).reset_index()
center_items2 = center_items[['reference_code']]

#part numbers in fishbowl
fishbowl_items = fishbowl_inventory[['PARTNUM']]
fishbowl_items = fishbowl_items['PARTNUM'].drop_duplicates()
fishbowl_items = pd.DataFrame(fishbowl_items).reset_index()
fishbowl_items = fishbowl_items[['PARTNUM']]

# Find non-matching values in 'part number' column
non_matching_values = set(center_items2['reference_code']).symmetric_difference(set(fishbowl_items['PARTNUM']))

# Create a DataFrame with non-matching values
non_matching_df = center_items2[center_items2['reference_code'].isin(non_matching_values)].append(
    fishbowl_items[fishbowl_items['PARTNUM'].isin(non_matching_values)])

#listed at center, not in FB:
at_center = non_matching_df[non_matching_df['reference_code'].notna()]
at_center = at_center[at_center['reference_code'].notna()]

#These are items at center that are not on the list in fishbowl
center= center_inventory[['reference_code', 'description', 'at_center']]
center_details = at_center.merge(center, on = "reference_code", how = "left")
#center_details = center_details.drop_duplicates()
#center_details = pd.DataFrame(center_details).reset_index()
center_details = center_details[['reference_code', 'PARTNUM', 'description', 'at_center']]

center_details = center_details.groupby(['reference_code', 'description'])['at_center'].sum()
center_details = pd.DataFrame(center_details).reset_index()

#listed in FB, not at center:
in_fb = non_matching_df[non_matching_df['PARTNUM'].notna()]
in_fb = in_fb[in_fb['PARTNUM'].notna()]

#These are part numbers in fishbowl inventory that are not listed in the center list
fishbowl= fishbowl_inventory[['PARTNUM', 'PARTDESC', 'tagcaseqty']]
fishbowl_details = in_fb.merge(fishbowl, on = "PARTNUM", how = "left")
#fishbowl_details = fishbowl_details.drop_duplicates()
#fishbowl_details = pd.DataFrame(fishbowl_details).reset_index()
fishbowl_details = fishbowl_details[['reference_code', 'PARTNUM', 'PARTDESC', 'tagcaseqty']]

fishbowl_details = fishbowl_details.groupby(['PARTNUM', 'PARTDESC'])['tagcaseqty'].sum()
fishbowl_details = pd.DataFrame(fishbowl_details).reset_index()

#collapsed part numbers and total quantities
fishbowl_df = fishbowl.groupby(['PARTNUM', 'PARTDESC'])['tagcaseqty'].sum()
fishbowl_df =  pd.DataFrame(fishbowl_df).reset_index()

#collapsed part numbers and total quantities at center
center_df = center_inventory[['prod#', 'at_center', 'reference_code']]
center_df = center_df.groupby(['prod#', 'reference_code'])['at_center'].sum()
center_df =  pd.DataFrame(center_df).reset_index()

#matching on reference numbers
matching_values = pd.merge(center_df, fishbowl_df, left_on='reference_code', right_on='PARTNUM', how = 'inner')
matching_values=matching_values.drop_duplicates().reset_index()
matching_values = matching_values[['prod#', 'reference_code', 'PARTDESC', 'at_center', 'tagcaseqty']]

matching_values['difference'] = matching_values['tagcaseqty']-matching_values['at_center']
matching_values.query('difference != 0')
mismatched_qty = matching_values.query('difference != 0')

# Create an ExcelWriter object and specify the Excel file path
excel_writer = pd.ExcelWriter('11_9_inventory_comparison.xlsx', engine='xlsxwriter')

# Use the to_excel method to save each DataFrame to a different sheet
center_details.to_excel(excel_writer, sheet_name='at_center_only', index=False)
fishbowl_details.to_excel(excel_writer, sheet_name='in_fishbowl_only', index=False)
mismatched_qty.to_excel(excel_writer, sheet_name='mismatches', index=False)

# Save the Excel file
excel_writer.save()

