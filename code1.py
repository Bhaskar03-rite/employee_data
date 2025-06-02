import pandas as pd

excel_file =r"C:\Users\VenkataBhaskarReddyS\Downloads\Daily Attendance Report 2.xls"
sheet1_df = pd.read_excel(excel_file, sheet_name='Sheet1')
sheet2_df = pd.read_excel(excel_file, sheet_name='Sheet2')
df =pd.concat([sheet1_df, sheet2_df], axis=0, ignore_index=True)
records = []
current_date = None
current_department = None
headers = []

for idx, row in df.iterrows():
    row_values = row.dropna().tolist()

    if not row_values:
        continue

    first_cell = str(row_values[0]).strip().lower()

    if first_cell == "attendance date":
        current_date = row_values[1] if len(row_values) > 1 else None

    elif first_cell == "department":
        current_department = row_values[1] if len(row_values) > 1 else None

    elif first_cell in ("sno", "sn"):  # Header row
        headers = ['Attendance Date', 'Department'] + row.tolist()

    else:
        full_row = [current_date, current_department] + row.tolist()
        records.append(full_row)

if not headers:
    raise ValueError("No headers detected. Make sure your file has an 'SNo' row.")

# Create DataFrame
clean_df = pd.DataFrame(records, columns=headers)

# Clean up column names
clean_df.columns = clean_df.columns.str.strip()

columns = ['SNo','E. Code','Name','Shift','InTime','OutTime','Work Dur.','OT','Tot.  Dur.','Status','Remarks','Attendance Date','Department']

clean_df=clean_df[columns]
clean_df = clean_df.dropna(subset=['E. Code'])
# Save output
clean_df.to_excel("merged_output.xlsx", index=False)
print(clean_df.columns)
print(" Processed and saved successfully!")
