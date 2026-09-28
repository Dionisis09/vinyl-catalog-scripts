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

# Step 1: Find columns by header name
photo_columns = {}
facebook_col = None

for col_idx in range(1, ws.max_column + 1):
    header = ws.cell(1, col_idx).value
    
    if header == 'Front cover photo':
        photo_columns[col_idx] = 1  # Maps to .1.jpg
    elif header == 'back cover photo ':
        photo_columns[col_idx] = 2  # Maps to .2.jpg
    elif header == 'Sleeve photo ':
        photo_columns[col_idx] = 3  # Maps to .3.jpg
    elif header == 'Front vinyl photo':
        photo_columns[col_idx] = 4  # Maps to .4.jpg
    elif header == 'Back vinyl photo':
        photo_columns[col_idx] = 5  # Maps to .5.jpg
    elif header == 'Facebook':
        facebook_col = col_idx

print(f"Found photo columns: {photo_columns}")
print(f"Facebook column: {facebook_col}")

# Step 2: Create hyperlinks for each photo cell
for row_idx in range(2, ws.max_row + 1):
    # Get record ID from column A (Α/Α)
    record_id = ws.cell(row_idx, 1).value
    
    if record_id is None:
        continue
    
    try:
        record_id = int(record_id) if isinstance(record_id, (int, float)) else int(str(record_id).split('.')[0])
    except:
        continue
    
    # For each photo column, create proper hyperlink
    for col_idx, photo_number in photo_columns.items():
        photo_filename = f"{record_id}.{photo_number}.jpg"
        photo_path = os.path.join(photos_dir, photo_filename)
        
        if os.path.exists(photo_path):
            photo_cell = ws.cell(row_idx, col_idx)
            # Create hyperlink with full file URL
            file_url = 'file:///' + photo_path.replace('\\', '/')
            photo_cell.hyperlink = Hyperlink(ref=photo_cell.coordinate, target=file_url)
            # Display just the filename
            photo_cell.value = photo_filename
            print(f"Row {row_idx}, Col {col_idx}: {record_id} -> {photo_filename}")
        else:
            print(f"NOT FOUND: {photo_path}")

# Step 3: Delete Facebook column
if facebook_col:
    ws.delete_cols(facebook_col)
    print(f"\n✓ Removed Facebook column")

# Save to temp file
wb.save(temp_file)
print(f"✓ Hyperlinks created")

# Replace original
os.remove(excel_file)
os.rename(temp_file, excel_file)

print(f"\n✓ File updated: {excel_file}")
