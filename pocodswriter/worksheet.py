class Worksheet:
    def __init__(self):
        self.data = {}
        self.tables = []
        self.column_widths = {}

    def hide_gridlines(self, option):
        self.gridlines_option = option

    def write(self, row, col, data, _format, *args):
        self.data[(row, col)] = {"format": _format, "data": data}

    def write_url(self, row, col, url, _format, display_text, *args):
        self.data[(row, col)] = {"format": _format, "data": display_text, "url": url}

    def set_column(self, first_col, last_col, width, cell_format=None, options=None):
        for col in range(first_col, last_col + 1):
            self.column_widths[col] = width / 5

    def get_number_of_columns(self):
        return max(col for row, col in self.data.keys()) + 1

    def get_column_widths(self):
        column_widths = [2.258] * self.get_number_of_columns()
        for index, width in self.column_widths.items():
            column_widths[index] = width
        return column_widths

    def _cell_reference(self, row, col):
        if col < 26:
            column_letters = chr(ord('A') + col)
        else:
            column_letters = chr(ord('A') + col // 26 - 1) + chr(ord('A') + col % 26)
        return column_letters + str(row + 1)

    def add_table(self, first_row, first_col, last_row, last_col, options=None):
        self.tables.append({
            "first_cell": self._cell_reference(first_row, first_col),
            "last_cell": self._cell_reference(last_row, last_col),
            "name": options["name"]
        })

    def data_as_jagged_array(self):
        result = []
        for row, col in self.data:
            if row >= len(result):
                result += [[] for _ in range(row - len(result) + 1)]
            if col >= len(result[row]):
                result[row] += [""] * (col - len(result[row]) + 1)
            result[row][col] = self.data[(row, col)]
        return result
