from openpyxl import load_workbook
from openpyxl.worksheet.hyperlink import Hyperlink
import os

excel_file = r'c:\Users\dioni\Desktop\vinilia 2\vinyl_records_list.xlsx'
temp_file = r'c:\Users\dioni\Desktop\vinilia 2\vinyl_records_list_TEMP.xlsx'
photos_dir = r'c:\Users\dioni\Desktop\vinilia 2\vinilia 2'

# Load workbook
wb = load_workbook(excel_file)
ws = wb.active

print(f"Current rows: {ws.max_row}")
print(f"Current last record ID: {ws.cell(ws.max_row, 1).value}")

# Find photo columns
photo_col_info = {}
for col_idx in range(1, ws.max_column + 1):
    header = ws.cell(1, col_idx).value
    if header == 'Front cover photo':
        photo_col_info[col_idx] = 1
    elif header == 'back cover photo ':
        photo_col_info[col_idx] = 2
    elif header == 'Sleeve photo ':
        photo_col_info[col_idx] = 3
    elif header == 'Front vinyl photo':
        photo_col_info[col_idx] = 4
    elif header == 'Back vinyl photo':
        photo_col_info[col_idx] = 5

# Find the last row
last_row = ws.max_row

# Add new records 81-104
print("\nAdding new records 81-104...")
for record_id in range(81, 105):
    new_row = last_row + (record_id - 80)
    
    # Set ID in column A
    ws.cell(new_row, 1).value = record_id
    
    # Set title in column B to just the number
    ws.cell(new_row, 2).value = str(record_id)
    
    # Add photo hyperlinks
    for col_idx, photo_num in photo_col_info.items():
        photo_filename = f"{record_id}.{photo_num}.jpg"
        photo_path = os.path.join(photos_dir, photo_filename)
        
        if os.path.exists(photo_path):
            cell = ws.cell(new_row, col_idx)
            cell.value = photo_filename
            cell.hyperlink = Hyperlink(ref=cell.coordinate, target=photo_path)
            if record_id <= 82:  # Show first few
                print(f"✓ Row {new_row}, ID {record_id}: {photo_filename}")

# Save
wb.save(temp_file)
os.remove(excel_file)
os.rename(temp_file, excel_file)

print(f"\n✓ DONE!")
print(f"Added records 81-104 (24 new records)")
print(f"Total rows now: {ws.max_row}")
