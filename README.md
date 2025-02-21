# Dispatch-App-2-Excel-Data-Matcher
This Python script uses a Tkinter GUI to compare and highlight matching rows from two Excel files. It's designed for processing return reports and identifying partial matches between shipment numbers and product descriptions.
Features:

📂 Load two Excel files (Raport_zwrot_postint and Zwroty UA) via a simple graphical interface.
🔍 Automatically compare shipment numbers and product descriptions using partial string matching.
🎨 Highlight matching rows directly in the Excel file using yellow fill.
💾 Save the updated Excel file with highlighted matches.
Technologies Used:

pandas for data manipulation
openpyxl for working with Excel files and cell formatting
difflib for fuzzy string comparison
tkinter for a simple GUI file selection
This tool is ideal for logistics and return management tasks, helping to quickly identify matching shipment records.
