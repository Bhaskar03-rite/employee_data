import pandas as pd

excel_file = r"C:\Users\VenkataBhaskarReddyS\Downloads\Daily Attendance Report 2.xls"
sheet1_df = pd.read_excel(excel_file, sheet_name='Sheet1', header=None)
sheet2_df = pd.read_excel(excel_file, sheet_name='Sheet2', header=None)

merged_df = pd.concat([sheet1_df, sheet2_df], axis=0, ignore_index=True)

merged_df.to_excel("merged_output.xlsx", index=False)

# Print output
print(merged_df)