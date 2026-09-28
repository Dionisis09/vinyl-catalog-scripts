from openpyxl import load_workbook
from openpyxl.worksheet.hyperlink import Hyperlink
import os

# Paths
excel_file = r'c:\Users\dioni\Desktop\vinilia 2\vinyl_records_list.xlsx'
temp_file = r'c:\Users\dioni\Desktop\vinilia 2\vinyl_records_list_TEMP.xlsx'
photos_dir = r'c:\Users\dioni\Desktop\ninyls 2\vinilia 2\vinilia 2'

# Load workbook
wb = load_workbook(excel_file)
ws = wb.active

print("Step 1: Clearing all existing hyperlinks...")
for row in ws.iter_rows():
    for cell in row:
        cell.hyperlink = None

print("Step 2: Finding photo columns...")
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

print("Step 3: Creating hyperlinks with proper references...")
total_links = 0

for row_idx in range(2, ws.max_row + 1):
    record_id = ws.cell(row_idx, 1).value
    
    if record_id is None:
        continue
    
    try:
        record_id = int(record_id) if isinstance(record_id, (int, float)) else int(str(record_id).split('.')[0])
    except:
        continue
    
    for col_idx, photo_num in photo_col_info.items():
        photo_filename = f"{record_id}.{photo_num}.jpg"
        photo_path = os.path.join(photos_dir, photo_filename)
        
        if os.path.exists(photo_path):
            cell = ws.cell(row_idx, col_idx)
            cell.value = photo_filename
            
            # Use relative path or absolute path with proper encoding
            # Try using absolute Windows path with quotes
            cell.hyperlink = Hyperlink(ref=cell.coordinate, target=photo_path)
            total_links += 1
            
            if row_idx == 3:
                print(f"Row {row_idx}, Col {col_idx}: {photo_filename} -> {photo_path}")
        else:
            print(f"File not found: {photo_path}")

print(f"\nStep 4: Saving...")
wb.save(temp_file)
os.remove(excel_file)
os.rename(temp_file, excel_file)

print(f"✓ Updated! Total links: {total_links}")
