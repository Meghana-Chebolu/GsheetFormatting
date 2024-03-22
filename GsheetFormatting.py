class GoogleSpreadsheet:
    def __init__(self, spreadsheet_id, worksheet_name):
        auth.authenticate_user()
        creds, _ = default()
        self.gc = gspread.authorize(creds)
        self.spreadsheet = self.gc.open_by_key(spreadsheet_id)
        self.worksheet = self.spreadsheet.worksheet(worksheet_name)

    def set_header(self, col_names, table_range):
        cell_list = self.worksheet.range(table_range)
        empty_range = all(cell.value == '' for cell in cell_list)
        if empty_range:
            self.worksheet.insert_row(col_names, index=1)
        self.worksheet.update(table_range, [col_names])

    def get_col_names(self):
        col_names = self.worksheet.row_values(1)
        return col_names

    def get_col_id(self, col_name):
        header_row = self.get_col_names()
        if col_name in header_row:
            col_num = header_row.index(col_name)
            return col_num
        print(f"Column '{col_name}' does not exist.")

    def append_row(self, row):
        self.worksheet.append_row(row)

    def update_cell(self, row_num, col_num, value):
        self.worksheet.update_cell(row_num,col_num, value)
        print(f"New value updated at  row {row_num}, col {col_num}.")

    def format_sheet(self):
        set_frozen(self.worksheet, rows=1)
        self.worksheet.format("A:ZZ", {"wrapStrategy": "WRAP"})
        cell_format = {
        "verticalAlignment": "TOP"
        }
        self.worksheet.format('A2:ZZ1000', cell_format)
        headers = self.worksheet.row_values(1)
        header_fmt = cellFormat(
        backgroundColor = color(1, 0.8, 0.95),
        textFormat = textFormat(bold=True, foregroundColor = color(0, 0, 0)),
        horizontalAlignment = 'CENTER',
        borders = borders(
            top = border('SOLID', color = color(0, 0, 0)),
            bottom = border('SOLID', color = color(0, 0, 0)),
            left = border('SOLID', color = color(0, 0, 0)),
            right = border('SOLID', color = color(0, 0, 0))
        )
             )
        format_cell_range(self.worksheet, '1', header_fmt)
        data_fmt = cellFormat(
        horizontalAlignment = 'LEFT',
        borders = borders(
            top = border('SOLID', color = color(0, 0, 0)),
            bottom = border('SOLID', color = color(0, 0, 0)),
            left = border('SOLID', color = color(0, 0, 0)),
            right = border('SOLID', color = color(0, 0, 0))
        )
         )
        format_cell_range(self.worksheet, 'A1:Z1000', data_fmt)

        num_columns = len(headers)
        max_lengths = []

        for i in range(num_columns):
            column_values = self.worksheet.col_values(i + 1)  # Google Sheets columns are 1-indexed
            length = max(len(value) for value in column_values)
            max_length = 370 if length>=40 else 140
            max_lengths.append((chr(65 + i), max_length))
        set_column_widths(self.worksheet,max_lengths)
        print('Gsheet successfully formatted.')

    def freeze_columns(self,columns):
        set_frozen(self.worksheet,cols=columns)

    def reset_sheet(self):
        self.worksheet.clear()
