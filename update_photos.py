from openpyxl import load_workbook
import os

# Paths
excel_file = r'c:\Users\dioni\Desktop\vinilia 2\vinyl_records_list.xlsx'
photos_dir = r'c:\Users\dioni\Desktop\ninyls 2\vinilia 2\vinilia 2'

# Photo column mapping (column letters) - CORRECTED
photo_columns = {
    'E': '.1.jpg',  # Front cover photo
    'F': '.2.jpg',  # back cover photo
    'G': '.3.jpg',  # Sleeve photo
    'H': '.4.jpg',  # Front vinyl photo
    'I': '.5.jpg'   # Back vinyl photo
}

# Load workbook
wb = load_workbook(excel_file)
ws = wb.active

# Find the ID column
id_col = 'A'
count = 0

# Process each row (start from 2 to skip header)
for row_idx in range(2, ws.max_row + 1):
    record_id = ws[f'{id_col}{row_idx}'].value
    
    if record_id is None:
        continue
    
    # Convert record_id to int
    try:
        record_id_int = int(record_id) if isinstance(record_id, (int, float)) else int(str(record_id).split('.')[0])
    except:
        continue
    
    # Add photo paths for each column
    for col_letter, suffix in photo_columns.items():
        photo_path = os.path.join(photos_dir, f"{record_id_int}{suffix}")
        if os.path.exists(photo_path):
            ws[f'{col_letter}{row_idx}'].value = photo_path
            count += 1

# Save workbook
wb.save(excel_file)

print(f"✓ Excel file updated with correct photo paths!")
print(f"Total photo entries added: {count}")
print(f"File: vinyl_records_list.xlsx")
