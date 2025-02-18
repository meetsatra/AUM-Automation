import pandas as pd
from datetime import timedelta

pivot_file_name = 'pivot.xlsx'

# Reading the 'old_ddr' sheet from 'pivot.xlsx'
old_ddr_df = pd.read_excel(pivot_file_name, sheet_name='old_ddr')

# Function to set 'Batches' column based on time intervals
def set_batches(row):
    if row['Time'] < '12:00:00':
        return '1. Till 12 pm'
    elif '12:00:00' <= row['Time'] < '13:00:00':
        return '2. 12 to 1 pm'
    elif '13:00:00' <= row['Time'] < '14:00:00':
        return '3. 1 to 2 pm'
    elif '14:00:00' <= row['Time'] < '15:00:00':
        return '4. 2 to 3 pm'
    elif '15:00:00' <= row['Time'] < '16:00:00':
        return '5. 3 to 4 pm'
    elif '16:00:00' <= row['Time'] < '17:00:00':
        return '6. 4 to 5 pm'
    elif '17:00:00' <= row['Time'] < '18:00:00':
        return '7. 5 to 6 pm'
    elif '18:00:00' <= row['Time'] < '19:00:00':
        return '8. 6 to 7 pm'
    elif '19:00:00' <= row['Time'] < '20:00:00':
        return '9. 7 to 8 pm'
    elif '20:00:00' <= row['Time'] < '21:00:00':
        return '10. 8 to 9 pm'
    elif '21:00:00' <= row['Time'] < '22:00:00':
        return '11. 9 to 10 pm'
    elif '22:00:00' <= row['Time'] < '23:00:00':
        return '12. 10 to 11 pm'
    elif '23:00:00' <= row['Time']:
        return '13. 11 to 12 am'
    else:
        return None

# Apply the function to create the 'Batches' column
old_ddr_df['Batches'] = old_ddr_df.apply(set_batches, axis=1)

# Save the modified 'old_ddr_df' to 'pivot.xlsx' in the 'old_ddr' sheet
with pd.ExcelWriter(pivot_file_name, engine='openpyxl', mode='a') as writer:
    # If 'old_ddr' sheet exists in 'pivot.xlsx', delete that sheet
    if 'old_ddr' in writer.book.sheetnames:
        writer.book.remove(writer.book['old_ddr'])

    # Write the modified 'old_ddr_df' to 'old_ddr' sheet in 'pivot.xlsx'
    old_ddr_df.to_excel(writer, sheet_name='old_ddr', index=False)

print(f"Modified 'old_ddr_df' saved to 'old_ddr' sheet in {pivot_file_name}.")
