

import win32com.client

try:
    excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = True
    print("Excel Application Initialized Successfully")
except Exception as e:
    print(f"Error initializing Excel Application: {e}")