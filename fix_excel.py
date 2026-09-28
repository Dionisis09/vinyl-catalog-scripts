from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
import os
import shutil

# Paths
excel_file = r'c:\Users\dioni\Desktop\vinilia 2\vinyl_records_list.xlsx'
temp_file = r'c:\Users\dioni\Desktop\vinilia 2\vinyl_records_list_TEMP.xlsx'
photos_dir = r'c:\Users\dioni\Desktop\ninyls 2\vinilia 2\vinilia 2'

# Load workbook
wb = load_workbook(excel_file)
ws = wb.active

# Remove Facebook column (find it first)
header_row = 1
facebook_col = None
for col_idx, cell in enumerate(ws[header_row], 1):
    if cell.value == 'Facebook':
        facebook_col = col_idx
        break

if facebook_col:
    ws.delete_cols(facebook_col)
    print(f"✓ Removed Facebook column (was column {get_column_letter(facebook_col)})")

# Photo columns mapping with their suffix
photo_columns = {
    'Front cover photo': '.1.jpg',
    'back cover photo ': '.2.jpg',
    'Sleeve photo ': '.3.jpg',
    'Front vinyl photo': '.4.jpg',
    'Back vinyl photo': '.5.jpg'
}

# Process each row to create hyperlinks to actual photo files
for row_idx in range(2, ws.max_row + 1):
    # Get record ID from column A
    id_cell = ws.cell(row_idx, 1)
    record_id = id_cell.value
    
    if record_id is None:
        continue
    
    try:
        record_id = int(record_id) if isinstance(record_id, (int, float)) else int(str(record_id).split('.')[0])
    except:
        continue
    
    # Find photo columns by header name and create hyperlinks
    for col_idx in range(1, ws.max_column + 1):
        header_cell = ws.cell(1, col_idx)
        if header_cell.value in photo_columns:
            photo_suffix = photo_columns[header_cell.value]
            photo_filename = f"{record_id}{photo_suffix}"
            photo_path = os.path.join(photos_dir, photo_filename)
            
            if os.path.exists(photo_path):
                photo_cell = ws.cell(row_idx, col_idx)
                # Create hyperlink - use file:/// protocol for local files
                file_url = 'file:///' + photo_path.replace('\\', '/')
                photo_cell.hyperlink = Hyperlink(ref=photo_cell.coordinate, target=file_url)
                # Set cell text to display (show the filename like 2.1.jpg)
                photo_cell.value = photo_filename

# Save to temp file
wb.save(temp_file)
print(f"✓ Created hyperlinks to actual photo files")

# Replace original file
os.remove(excel_file)
os.rename(temp_file, excel_file)

print(f"Total columns: {ws.max_column}")
print(f"Total rows: {ws.max_row - 1}")
print(f"\nFile updated: {excel_file}")
