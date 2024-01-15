import openpyxl
EXCEL_COMMON_TYPES = int|float|str

WORKBOOKPATH = ""

wb = openpyxl.load_workbook(WORKBOOKPATH)
ws = wb.active

def cell_to_coords(cell):
    col, row = '', ''
    for char in cell:
        if char.isdigit():
            row += char
        else:
            col += char
    col = sum((ord(char.lower()) - ord('a') + 1) * (26 ** i) for i, char in enumerate(reversed(col)))
    return (col, int(row))

# below are basic APIs

def set_text(dest: str, value: EXCEL_COMMON_TYPES):
    ws[dest] = value

def insert_row_before(dest: str):
    ws.insert_rows(cell_to_coords(dest)[0])

def insert_column_before(dest: str):
    ws.insert_cols(cell_to_coords(dest)[1])

# TODO: deletion

def set_title(text: str):
    ws.title = text

def create_sheet(title: str):
    ws = wb.create_sheet(title)

def copy_region(dest: str, source: str):
    ws[dest] = ws[source]

# above are basic APIs
    
# below are Excel functions

def f_sum(dest: str, regions: str):
    ws[dest] = f"=SUM({regions})"

def f_average(dest: str, regions: str):
    ws[dest] = f"=AVERAGE({regions})"

def f_max(dest: str, regions: str):
    ws[dest] = f"=MAX({regions})"

def f_min(dest: str, regions: str):
    ws[dest] = f"=MIN({regions})"

def f_count(dest: str, regions: str):
    ws[dest] = f"=COUNT({regions})"

def f_median(dest: str, regions: str):
    ws[dest] = f"=MEDIAN({regions})"

def f_mode(dest: str, regions: str):
    ws[dest] = f"=MODE({regions})"

def f_stdev(dest: str, regions: str):
    ws[dest] = f"=STDEV({regions})"

def f_var(dest: str, regions: str):
    ws[dest] = f"=VAR({regions})"

def f_if(dest: str, condition: str, t_branch: EXCEL_COMMON_TYPES, f_branch: EXCEL_COMMON_TYPES):
    ws[dest] = f"=IF({condition}, {t_branch}, {f_branch})"

def f_vlookup():
    pass

# above are Excel functions

