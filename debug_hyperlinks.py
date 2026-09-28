from openpyxl import load_workbook

excel_file = r'c:\Users\dioni\Desktop\vinilia 2\vinyl_records_list.xlsx'
wb = load_workbook(excel_file)
ws = wb.active

print("Checking row 3 in detail:")
for col_idx in range(4, 9):
    cell = ws.cell(3, col_idx)
    print(f"\nCol {col_idx}:")
    print(f"  Value: {cell.value}")
    print(f"  Hyperlink object: {cell.hyperlink}")
    print(f"  Hyperlink type: {type(cell.hyperlink)}")
    if cell.hyperlink:
        print(f"  Target: {cell.hyperlink.target}")
        print(f"  Target type: {type(cell.hyperlink.target)}")
