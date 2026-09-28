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
# First, clear ALL existing hyperlinks
for row in ws.iter_rows():
    for cell in row:
        cell.hyperlink = None

print("Step 2: Identifying photo columns by header...")
# Find the correct columns
photo_col_info = {}
for col_idx in range(1, ws.max_column + 1):
    header = ws.cell(1, col_idx).value
    if header == 'Front cover photo':
        photo_col_info[col_idx] = (1, 'Front cover')
    elif header == 'back cover photo ':
        photo_col_info[col_idx] = (2, 'back cover')
    elif header == 'Sleeve photo ':
        photo_col_info[col_idx] = (3, 'Sleeve')
    elif header == 'Front vinyl photo':
        photo_col_info[col_idx] = (4, 'Front vinyl')
    elif header == 'Back vinyl photo':
        photo_col_info[col_idx] = (5, 'Back vinyl')

print(f"Photo columns found: {photo_col_info}")

print("Step 3: Creating new hyperlinks...")
total_links = 0
# Now create hyperlinks for each row
for row_idx in range(2, ws.max_row + 1):
    record_id_cell = ws.cell(row_idx, 1)
    record_id = record_id_cell.value
    
    if record_id is None:
        continue
    
    # Convert to integer
    try:
        record_id = int(record_id) if isinstance(record_id, (int, float)) else int(str(record_id).split('.')[0])
    except Exception as e:
        print(f"Error parsing ID at row {row_idx}: {record_id}, error: {e}")
        continue
    
    # For each photo column
    for col_idx, (photo_num, col_name) in photo_col_info.items():
        photo_filename = f"{record_id}.{photo_num}.jpg"
        photo_path = os.path.join(photos_dir, photo_filename)
        
        cell = ws.cell(row_idx, col_idx)
        
        if os.path.exists(photo_path):
            # Set the cell value to the filename
            cell.value = photo_filename
            
            # Create the hyperlink
            file_url = 'file:///' + photo_path.replace('\\', '/')
            cell.hyperlink = file_url
            total_links += 1
            
            if row_idx <= 5:  # Print first few for verification
                print(f"✓ Row {row_idx}, Col {col_idx} ({col_name}): {record_id} -> {photo_filename} -> {file_url}")
        else:
            cell.value = None
            print(f"✗ File not found: {photo_path}")

print(f"\nStep 4: Saving file...")
# Save to temp
wb.save(temp_file)

# Replace original
os.remove(excel_file)
os.rename(temp_file, excel_file)

print(f"✓ File updated successfully!")
print(f"Total hyperlinks created: {total_links}")
