from gspread_formatting import *

def Formatting(worksheet):
    all_values = worksheet.get_all_values()

    # Calculate the number of filled rows and columns
    num_filled_rows = len(all_values)
    num_filled_columns = len(all_values[0]) if num_filled_rows > 0 else 0
    cell_range = f"A2:{chr(ord('A') + num_filled_columns - 1)}{num_filled_rows + 1}"

    # Define formats
    header_fmt = cellFormat(
        backgroundColor=color(1, 0.8, 1),
        textFormat=textFormat(bold=True, foregroundColor=color(0, 0, 0)),
        horizontalAlignment='CENTER',
        borders=borders(
            top=border('SOLID', color=color(0, 0, 0)),
            bottom=border('SOLID', color=color(0, 0, 0)),
            left=border('SOLID', color=color(0, 0, 0)),
            right=border('SOLID', color=color(0, 0, 0))
        )
    )

    data_fmt = cellFormat(
        horizontalAlignment='LEFT',
        borders=borders(
            top=border('SOLID', color=color(0, 0, 0)),
            bottom=border('SOLID', color=color(0, 0, 0)),
            left=border('SOLID', color=color(0, 0, 0)),
            right=border('SOLID', color=color(0, 0, 0))
        )
    )

    # Format headers
    headers = worksheet.row_values(1)  # Get the headers from the first row
    num_columns = len(headers)
    max_lengths = []

    for i in range(num_columns):
        column_values = worksheet.col_values(i + 1)  # Google Sheets columns are 1-indexed
        length = max(len(value) for value in column_values)
        max_length = 250 if length >= 40 else 140
        max_lengths.append((chr(65 + i), max_length))
    set_column_widths(worksheet, max_lengths)
    worksheet.format("A:ZZ", {"wrapStrategy": "WRAP"})
    cell_format = {
        "verticalAlignment": "TOP"
    }
    worksheet.format('A1:ZZ1000', cell_format)
    format_cell_range(worksheet, '1', header_fmt)

    # Format the data rows
    format_cell_range(worksheet, cell_range, data_fmt)

    # Freeze the header row
    set_frozen(worksheet, rows=1)

    print("Successfully made changes in the Google sheet.")
