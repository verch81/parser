from tabulate import tabulate #print to console as a table
from openpyxl import Workbook #export to .xlsx
from openpyxl.styles import Alignment, Font #add styling in Excel table

def parse_file_to_dict_list(filename):
    """
    Parses a file where each line contains parameters separated by '&'.
    Each parameter is a key=value pair.

    Args:
        filename (str): Path to the input file.

    Returns:
        list of dict: List of dictionaries, one per line, preserving order.
    """
    dicts = []
    with open(filename, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue  # skip empty lines
            params = line.split('&')
            d = {}
            for p in params:
                # Split by first '=', in case value contains '='
                key_value = p.split('=', 1)
                if len(key_value) == 2:
                    key, value = key_value
                    d[key] = value
                else:
                    # handle case where there is no '=' (optional)
                    d[key_value[0]] = ''
            dicts.append(d)
    return dicts

def dicts_to_table_data(data):
    """
    Transforms a list of dictionaries into a 2D list representing a table.
    The first row is the header: 'Key', followed by dictionary indices.
    Each subsequent row contains a unique key and its values across dictionaries.
    Missing values are replaced with 'n/a'.

    Args:
        data (list[dict]): List of dictionaries to transform.

    Returns:
        list[list]: 2D list representing the table.
    """
    # Collect all unique keys from all dictionaries
    all_keys = set()
    for d in data:
        all_keys.update(d.keys())

    # Sort keys for consistent ordering
    all_keys = sorted(all_keys)

    # Create the header row
    header = ['Key'] + [str(i) for i in range(len(data))]

    # Create the table rows
    rows = []
    for key in all_keys:
        row = [key]
        for d in data:
            row.append(d.get(key, 'n/a'))
        rows.append(row)

    # Combine header and data rows
    return [header] + rows

def save_table_to_excel(table_data, filename='output.xlsx', group_size=6): #add other value as a parameter
    """
    Save a 2D list to Excel and wrap comma-separated numbers
    into multiple lines ONLY for the row where first column is 's'.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    center_wrap_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    bold_font = Font(bold=True, name='Arial')
    regular_font = Font(name='Arial')

    for row_idx, row in enumerate(table_data, start=1):  # openpyxl is 1-based
        is_s_row = row[0] == 's'  # Check if this is the target row

        for col_idx, cell_value in enumerate(row, start=1):
            # Wrap only if it's the 's' row and not the first column
            if is_s_row and col_idx > 1 and isinstance(cell_value, str):
                parts = [p.strip() for p in cell_value.split(',')]
                grouped = [','.join(parts[i:i+group_size]) for i in range(0, len(parts), group_size)]
                cell_value = ';\n'.join(grouped) # add ';' for visual separation of rows
                # Optional: adjust row height
                ws.row_dimensions[row_idx].height = (cell_value.count('\n') + 1) * 15

            # Write the cell with alignment and font
            cell = ws.cell(row=row_idx, column=col_idx, value=cell_value)
            cell.alignment = center_wrap_alignment

            if row_idx == 1 or col_idx == 1:
                cell.font = bold_font
            else:
                cell.font = regular_font

    # Freeze first row and first column
    ws.freeze_panes = 'B2'
    wb.save(filename)

# Example usage for file parsing:
filename = 'input/input.txt'
dict_list = parse_file_to_dict_list(filename)
for d in dict_list:
    print(d)

# filename = input("Enter the filename to parse: ")
# dict_list = parse_file_to_dict_list(filename)
# for d in dict_list:
#     print(d)

# Example usage:
example_data = dict_list
# [
#     {'a': 1, 'b': 2},
#     {'a': 3, 'c': 4},
#     {'b': 5, 'c': 6, 'd': 7}
# ]

# 1. Convert to table format
table_data = dicts_to_table_data(example_data)

##### Below are output actions with table_data
# Save to Excel file
# [optional, default value = 6] enter number of items in a row after file name:
# save_table_to_excel(table_data, 'result/my_table6.xlsx',5)
save_table_to_excel(table_data, 'result/my_table6.xlsx')

# Now you can print it, convert to CSV, etc.
for row in table_data:
    print(row)

print(tabulate(table_data[1:], headers=table_data[0], tablefmt='grid')) #output to console via tabulate
print(table_data[-1])

