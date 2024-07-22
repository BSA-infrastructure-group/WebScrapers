"""
The purpose of this script is to run the FINRA IDs of a select group of individuals on the IRM
"""

import openpyxl
from datetime import datetime
import win32com.client as win32
import traceback
import time
from Generate_Header_Dictionary import get_column_headers
from FINRA_Scrape import search_webpage

#load the workbook and sheet
workbook_path = r"K:\Market Maps\Interest Rates Map.xlsm"
output_workbook_path = r"C:\Users\BSA-OliverJ'22\OneDrive\Desktop\OneDrive\Programming\Projects\WebScrapers\WebScrapers\workbooks\FINRA_Check_Output_Book.xlsm"
sheet_name = "Master"
table_name = "Master"
wb = openpyxl.load_workbook(workbook_path, data_only=True)
sheet = wb[sheet_name]

# Prompt the user for search values
location_search_value = input("Enter the location search value (e.g., 'New York, NY'): ").strip()
function_search_value = input("Enter the function search value (e.g., 'Trading'): ").strip()
firm_search_value = input("Enter the firm search value (e.g., 'Goldman Sachs'): ").strip()

#group_search_value = "Interest Rate Swaps" #Add this line back in if you want to filter by group

# Get column headers
column_headers = get_column_headers(workbook_path, sheet_name, table_name)

print("Column Headers:", column_headers)

# Initialize an empty list to store the results
results = []

for row in sheet.iter_rows(min_row=2, values_only=True): #need to check this to make sure that we are starting in the correct first row of the table
    
    #Check for specified conditions
    if row[column_headers['Location']] == location_search_value and \
        row[column_headers['Firm']] == firm_search_value and \
        row[column_headers['Function']] == function_search_value:
         
        # Print confirmation of match
        print("Match found for row:", row)

        #row[column_headers['Group']] == group_search_value and \ 'add this line back in if you want to filter by group
        
        finra_id = str(row[column_headers['FINRA ID']])
        print("FINRA ID:", finra_id)

        if not finra_id.isdigit():
            output = "No ID"
            time_of_check = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        else:
            output = search_webpage(finra_id)
            time_of_check = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Create a dictionary for the row
        result_row = {
            "Map Firm": row[column_headers['Firm']],
            "Name": row[column_headers['Name']],
            "Title": row[column_headers['Title']],
            "Function": row[column_headers['Function']],
            "Group": row[column_headers['Group']],
            "FINRA ID": finra_id,
            "Output": output,
            "Time of Check": time_of_check
        }
        results.append(result_row)

#Create a new workbook for output
output_wb = openpyxl.Workbook()
output_ws = output_wb.active

# Write headers
headers = ["Map Firm", "Name", "Title", "Function", "Group", "FINRA ID", "Output", "Time of Check"]
output_ws.append(headers)

# Write data rows
for result in results:
    row = [result[header] for header in headers]
    output_ws.append(row)

# Save the output workbook
output_wb.save(output_workbook_path)
output_wb.close()
wb.close()

#Path to the macro book
Macro_wb_path = r"C:\Users\BSA-OliverJ'22\OneDrive\Desktop\OneDrive\Programming\Projects\WebScrapers\WebScrapers\workbooks\FINRA_IRM_Report_Macro_Book.xlsm"

# Macro names
AFR_Macro = 'Attach_FINRA_Report.CreateEmailFromData'
Refresh_Macro = 'RefreshConnections.RefreshAllConnections'

try:
    # Start an instance of Excel
    excel_app = win32.Dispatch('Excel.Application')
    excel_app.Visible = False  # Or set to True if you want to see what happens in Excel

    # Open the workbook
    macro_wb = excel_app.Workbooks.Open(Filename=Macro_wb_path)

    # Run the macros
    excel_app.Run(Refresh_Macro)
    time.sleep(15)
    excel_app.Run(AFR_Macro)

    # Save the workbook
    macro_wb.Save()

# If there was an error, print the traceback
except Exception as e:
    print("An error occurred:")
    traceback.print_exc()

finally:
    # Close the workbook without saving again
    macro_wb.Close(False)

    # Quit Excel
    excel_app.Quit()

    # Cleanup the COM object
    del excel_app





