from gspread_formatting import *
def formatting(worksheet):

    all_values = worksheet.get_all_values()

    # Calculate the number of filled rows and columns
    num_filled_rows = len(all_values)
    num_filled_columns = len(all_values[0]) if num_filled_rows > 0 else 0
    cell_range = f"A2:{chr(ord('A') + num_filled_columns - 1)}{num_filled_rows + 1}"
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


    # Format the header row
    format_cell_range(worksheet, '1', header_fmt)

    # Format the data rows
    format_cell_range(worksheet, cell_range, data_fmt)

    # Freeze the header row
    set_frozen(worksheet, rows=1)



    print("Successfully made changes in the Google sheet.")