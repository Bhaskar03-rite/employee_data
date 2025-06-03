import pandas as pd
import re

#file path
excel_file = r"C:\Users\VenkataBhaskarReddyS\Downloads\Daily Log Report Matrix.xls"
sheet1_df = pd.read_excel(excel_file, sheet_name='Sheet1')
sheet2_df = pd.read_excel(excel_file, sheet_name='Sheet2')
df = pd.concat([sheet1_df, sheet2_df], axis=0, ignore_index=True)

# Remove the first 5 rows
df = df.iloc[5:]

all_records = []
current_date = None
data_started = False
headers = []
data_rows = []

for idx, row in df.iterrows():
    first_cell = str(row[0]) if pd.notna(row[0]) else ""

    # Detect Log Date row
    if re.search(r'Log\s*Date', first_cell, re.IGNORECASE):
        row_str = " ".join(str(x) for x in row if pd.notna(x))
        date_match = re.search(r'(\d{1,2}\s\w+\s\d{4})', row_str)
        if date_match:
            current_date = pd.to_datetime(date_match.group(1)).strftime('%Y-%m-%d')
        data_started = False
        headers = []
        data_rows = []
        continue

    # Detect header row (containing "Emp Code")
    if not data_started and any(isinstance(x, str) and 'emp code' in x.lower() for x in row if pd.notna(x)):
        headers = list(row)
        
        # Replace empty header names with generic "Unnamed_X" labels
        for i, col_name in enumerate(headers):
            if pd.isna(col_name) or (isinstance(col_name, str) and col_name.strip() == ''):
                headers[i] = f"Unnamed_{i}"
            else:
                headers[i] = str(col_name).strip()
        
        data_started = True
        continue

    # If data started, check for empty row (end of block)
    if data_started:
        if all(pd.isna(row[col]) or str(row[col]).strip() == '' for col in range(3)):
            if data_rows:
                temp_df = pd.DataFrame(data_rows, columns=headers)
                temp_df['Login Date'] = current_date
                all_records.append(temp_df)
            data_started = False
            headers = []
            data_rows = []
            continue
        
        data_rows.append(row.tolist())

if data_started and data_rows:
    temp_df = pd.DataFrame(data_rows, columns=headers)
    temp_df['Login Date'] = current_date
    all_records.append(temp_df)

# Combine all data blocks
if all_records:
    final_df = pd.concat(all_records, ignore_index=True)

    
    final_df.to_excel("combined_attendance_with_all_columns.xlsx", index=False)
    print("Processed data saved to 'combined_attendance_with_all_columns.xlsx'")
else:
    print("No attendance data found.")
