import pandas as pd
from datetime import timedelta
import random


dump_file_path = 'Dumps_AUM_Mov.xlsx'
pivot_file_name = 'pivot.xlsx'

# Reading the 'test' sheet from 'Dumps.xlsx'
old_ddr_df = pd.read_excel(dump_file_path, sheet_name='DDR Old')
old_coll_df = pd.read_excel(dump_file_path, sheet_name='Coll Old')
new_ddr_df = pd.read_excel(dump_file_path, sheet_name='DDR Today')
new_coll_df = pd.read_excel(dump_file_path, sheet_name='Coll Today')

# Splitting values in Column A into three columns: Date, Time, and UTC
old_ddr_df[['Date', 'Time', 'UTC']] = old_ddr_df['Created Date'].str.split(expand=True)
old_coll_df[['Date', 'Time', 'UTC']] = old_coll_df['Created at'].str.split(expand=True)
new_ddr_df[['Date', 'Time', 'UTC']] = new_ddr_df['Created Date'].str.split(expand=True)
new_coll_df[['Date', 'Time', 'UTC']] = new_coll_df['Created at'].str.split(expand=True)

# Update the time conversion with an explicit format
old_ddr_df['Time'] = (pd.to_datetime(old_ddr_df['Time'], format='%H:%M:%S', errors='coerce') + timedelta(hours=5, minutes=30)).dt.strftime('%H:%M:%S')
old_coll_df['Time'] = (pd.to_datetime(old_coll_df['Time'], format='%H:%M:%S', errors='coerce') + timedelta(hours=5, minutes=30)).dt.strftime('%H:%M:%S')
new_ddr_df['Time'] = (pd.to_datetime(new_ddr_df['Time'], format='%H:%M:%S', errors='coerce') + timedelta(hours=5, minutes=30)).dt.strftime('%H:%M:%S')
new_coll_df['Time'] = (pd.to_datetime(new_coll_df['Time'], format='%H:%M:%S', errors='coerce') + timedelta(hours=5, minutes=30)).dt.strftime('%H:%M:%S')

# Filter rows where 'Status' column is not 'rejected'
old_ddr_df = old_ddr_df[old_ddr_df['Status'] != 'rejected']
old_coll_df = old_coll_df[old_coll_df['Status'] != 'failed']
new_ddr_df = new_ddr_df[new_ddr_df['Status'] != 'rejected']
new_coll_df = new_coll_df[new_coll_df['Status'] != 'failed']

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
old_coll_df['Batches'] = old_coll_df.apply(set_batches, axis=1)
new_ddr_df['Batches'] = new_ddr_df.apply(set_batches, axis=1)
new_coll_df['Batches'] = new_coll_df.apply(set_batches, axis=1)

# Save the modified old_ddr_df to 'pivot.xlsx' in the 'test' sheet
with pd.ExcelWriter(pivot_file_name, engine='openpyxl', mode='a') as writer:
    # If 'old_ddr' sheet exists in 'pivot.xlsx', delete that sheet
    if 'old_ddr' in writer.book.sheetnames:
        writer.book.remove(writer.book['old_ddr'])

    # Write the modified 'old_ddr_df' to 'old_ddr' sheet in 'pivot.xlsx'
    old_ddr_df.to_excel(writer, sheet_name='old_ddr', index=False)

    if 'old_coll' in writer.book.sheetnames:
        writer.book.remove(writer.book['old_coll'])

    # Write the modified 'old_coll_df' to 'old_coll' sheet in 'pivot.xlsx'
    old_coll_df.to_excel(writer, sheet_name='old_coll', index=False)
    
    if 'new_ddr' in writer.book.sheetnames:
        writer.book.remove(writer.book['new_ddr'])

    # Write the modified 'new_ddr_df' to 'new_ddr' sheet in 'pivot.xlsx'
    new_ddr_df.to_excel(writer, sheet_name='new_ddr', index=False)
    
    if 'new_coll' in writer.book.sheetnames:
        writer.book.remove(writer.book['new_coll'])

    # Write the modified 'new_coll_df' to 'new_coll' sheet in 'pivot.xlsx'
    new_coll_df.to_excel(writer, sheet_name='new_coll', index=False)
    
    # Create a pivot table for 'old_ddr_df'
    pivot_table_old_ddr = pd.pivot_table(old_ddr_df, values='Drawdown Amount', index='Batches', aggfunc='sum')
    pivot_table_old_coll = pd.pivot_table(old_coll_df, values='Total amount', index='Batches', aggfunc='sum')
    pivot_table_new_ddr = pd.pivot_table(new_ddr_df, values='Drawdown Amount', index='Batches', aggfunc='sum')
    pivot_table_new_coll = pd.pivot_table(new_coll_df, values='Total amount', index='Batches', aggfunc='sum')

    # Save the pivot table to 'old_ddr' sheet in 'pivot.xlsx'
    if 'old_ddr_pivot' in writer.book.sheetnames:
        writer.book.remove(writer.book['old_ddr_pivot'])
    pivot_table_old_ddr.to_excel(writer, sheet_name='old_ddr_pivot')

    if 'old_coll_pivot' in writer.book.sheetnames:
        writer.book.remove(writer.book['old_coll_pivot'])
    pivot_table_old_coll.to_excel(writer, sheet_name='old_coll_pivot')
    
    if 'new_ddr_pivot' in writer.book.sheetnames:
        writer.book.remove(writer.book['new_ddr_pivot'])
    pivot_table_new_ddr.to_excel(writer, sheet_name='new_ddr_pivot')

    if 'new_coll_pivot' in writer.book.sheetnames:
        writer.book.remove(writer.book['new_coll_pivot'])
    pivot_table_new_coll.to_excel(writer, sheet_name='new_coll_pivot')

print(f"Modified 'old_ddr_df' saved to 'old_ddr' sheet in {pivot_file_name}.")

# old_ddr_df = old_ddr_df[old_ddr_df['Status'] != 'rejected']
# new_ddr_df = new_ddr_df[new_ddr_df['Status'] != 'rejected']

# Calculate the total sum of 'Drawdown Amount' column for both 'old_ddr' and 'new_ddr' sheets
total_old_ddr_amount = old_ddr_df['Drawdown Amount'].sum()
total_old_coll_amount = old_coll_df['Total amount'].sum()
total_new_ddr_amount = new_ddr_df['Drawdown Amount'].sum()
total_new_coll_amount = new_coll_df['Total amount'].sum()

# Finding the newest time in 'new_ddr_df'
newest_time_in_new_ddr = new_ddr_df['Time'].max()
newest_time_in_new_coll = new_coll_df['Time'].max()
# Filtering 'old_ddr_df' based on time until the newest time in 'new_ddr_df'
filtered_old_ddr_df = old_ddr_df[old_ddr_df['Time'] <= newest_time_in_new_ddr]
filtered_old_coll_df = old_coll_df[old_coll_df['Time'] <= newest_time_in_new_coll]
# Calculating the sum of 'Drawdown Amount' until the newest time
total_drawdown_amount_until_newest_time = filtered_old_ddr_df['Drawdown Amount'].sum()
total_collection_amount_until_newest_time = filtered_old_coll_df['Total amount'].sum()

# print("\n")
# print(f"Last DDR Today: {newest_time_in_new_ddr}")
# print(f"Last Collection Today: {newest_time_in_new_coll}")
# print(f"DDR untill {newest_time_in_new_ddr} on {old_ddr_df['Date'][0]} : {total_drawdown_amount_until_newest_time/10000000:.2f} Cr")
# print(f"Collection untill {newest_time_in_new_coll} on {old_coll_df['Date'][0]} : {total_collection_amount_until_newest_time/10000000:.2f} Cr")

# print(f"Total DDR on {old_ddr_df['Date'][0]} : {total_old_ddr_amount/10000000:.2f} Cr")
# print(f"Total Collection on {old_coll_df['Date'][0]} : {total_old_coll_amount/10000000:.2f} Cr")
# old_aum =  total_old_ddr_amount - total_old_coll_amount
# print(f"AUM on {old_ddr_df['Date'][0]}: {old_aum/10000000:.2f} Cr")

# print(f"Total DDR on {new_ddr_df['Date'][0]}: {total_new_ddr_amount/10000000:.2f} Cr")
# print(f"Total Collection on {new_coll_df['Date'][0]}: {total_new_coll_amount/10000000:.2f} Cr")

# predicted_ddr = (total_new_ddr_amount/total_drawdown_amount_until_newest_time) * total_old_ddr_amount
# print(f"Predicted DDR: {predicted_ddr/10000000:.2f} Cr")

# predicted_coll = (total_new_coll_amount/total_collection_amount_until_newest_time) * total_old_coll_amount
# print(f"Predicted Collection: {predicted_coll/10000000:.2f} Cr")



# curr_aum =  total_new_ddr_amount - total_new_coll_amount
# print(f"AUM on {new_ddr_df['Date'][0]}: {curr_aum/10000000:.2f} Cr")

# predicted_aum =  predicted_ddr - predicted_coll
# print(f"Predicted AUM {new_ddr_df['Date'][0]}: {predicted_aum/10000000:.2f} Cr")

print("\n")


def print_motivational_quote():
    quotes = [
        "Think like a proton and stay positive.",
        "Embrace challenges as opportunities for growth.",
        "Your mindset determines your success.",
        "Mistakes are proof that you are trying.",
        "Success is a journey, not a destination.",
        "The only limit is your mind.",
        "Every setback is a setup for a comeback.",
        "Growth starts at the end of your comfort zone.",
        "Your attitude determines your direction.",
        "Learn, adapt, and keep growing.",
        "Progress, not perfection.",
        "Believe in your potential to learn and improve.",
        "Challenges are opportunities in disguise.",
        "Success is not final; failure is not fatal.",
        "The more you learn, the more you earn.",
        "Your mindset shapes your reality.",
        "Turn obstacles into stepping stones.",
        "Be a work in progress.",
        "Strive for progress, not perfection.",
        "The only way to do great work is to love what you do."
    ]

    selected_quote = random.choice(quotes)
    print(selected_quote + " ~ Meet Satra")

# Call the function to print a motivational quote
print_motivational_quote()

print("\n")