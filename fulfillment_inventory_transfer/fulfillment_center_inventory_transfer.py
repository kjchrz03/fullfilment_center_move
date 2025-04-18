import pandas as pd
import regex as re
import numpy as np

#rename the file from Primo
#delete Reference Key tab
#remove macros
#save as csv
primo_inventory = pd.read_csv("PurpleCarrotInventory.csv")
fishbowl_inventory = pd.read_csv("Inventory Movement - Primo.csv")


# Find the exact column name by iterating through the column names
target_column = None
for column in primo_inventory.columns:
    if 'ON HAND' in column and 'ACTL' in column:
        target_column = column
        break

# Check if the target column was found
if target_column is not None:
    # Select the desired columns and rename the 'target_column'
    primo_inventory = primo_inventory[['PROD#', 'PRODUCT DESCRIPTION', target_column,  'REFERENCE CODE']]
    primo_inventory = primo_inventory.rename(columns={target_column: 'at_primo', 'PROD#': 'prod#', 'PRODUCT DESCRIPTION': 'description', 
                                                      'REFERENCE CODE': 'reference_code'})
else:
    print("The target column was not found.")


#reference codes at primo
primo_items = primo_inventory[['reference_code']]
primo_items = primo_items['reference_code'].drop_duplicates()
primo_items = pd.DataFrame(primo_items).reset_index()
primo_items2 = primo_items[['reference_code']]

#part numbers in fishbowl
fishbowl_items = fishbowl_inventory[['PARTNUM']]
fishbowl_items = fishbowl_items['PARTNUM'].drop_duplicates()
fishbowl_items = pd.DataFrame(fishbowl_items).reset_index()
fishbowl_items = fishbowl_items[['PARTNUM']]

# Find non-matching values in 'part number' column
non_matching_values = set(primo_items2['reference_code']).symmetric_difference(set(fishbowl_items['PARTNUM']))

# Create a DataFrame with non-matching values
non_matching_df = primo_items2[primo_items2['reference_code'].isin(non_matching_values)].append(
    fishbowl_items[fishbowl_items['PARTNUM'].isin(non_matching_values)])

#listed at Primo, not in FB:
at_primo = non_matching_df[non_matching_df['reference_code'].notna()]
at_primo = at_primo[at_primo['reference_code'].notna()]

#These are items at Primo that are not on the list in fishbowl
primo= primo_inventory[['reference_code', 'description', 'at_primo']]
primo_details = at_primo.merge(primo, on = "reference_code", how = "left")
#primo_details = primo_details.drop_duplicates()
#primo_details = pd.DataFrame(primo_details).reset_index()
primo_details = primo_details[['reference_code', 'PARTNUM', 'description', 'at_primo']]

primo_details = primo_details.groupby(['reference_code', 'description'])['at_primo'].sum()
primo_details = pd.DataFrame(primo_details).reset_index()

#listed in FB, not at Primo:
in_fb = non_matching_df[non_matching_df['PARTNUM'].notna()]
in_fb = in_fb[in_fb['PARTNUM'].notna()]

#These are part numbers in fishbowl inventory that are not listed in the primo list
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

#collapsed part numbers and total quantities at primo
primo_df = primo_inventory[['prod#', 'at_primo', 'reference_code']]
primo_df = primo_df.groupby(['prod#', 'reference_code'])['at_primo'].sum()
primo_df =  pd.DataFrame(primo_df).reset_index()

#matching on reference numbers
matching_values = pd.merge(primo_df, fishbowl_df, left_on='reference_code', right_on='PARTNUM', how = 'inner')
matching_values=matching_values.drop_duplicates().reset_index()
matching_values = matching_values[['prod#', 'reference_code', 'PARTDESC', 'at_primo', 'tagcaseqty']]

matching_values['difference'] = matching_values['tagcaseqty']-matching_values['at_primo']
matching_values.query('difference != 0')
mismatched_qty = matching_values.query('difference != 0')

# Create an ExcelWriter object and specify the Excel file path
excel_writer = pd.ExcelWriter('11_9_inventory_comparison.xlsx', engine='xlsxwriter')

# Use the to_excel method to save each DataFrame to a different sheet
primo_details.to_excel(excel_writer, sheet_name='at_primo_only', index=False)
fishbowl_details.to_excel(excel_writer, sheet_name='in_fishbowl_only', index=False)
mismatched_qty.to_excel(excel_writer, sheet_name='mismatches', index=False)

# Save the Excel file
excel_writer.save()

