from typing import Dict
from xlsxwriter import utility
from writers.styles import Styles


class Cell:
    """Stores xlsx cell data including coordinates, value, and format."""

    def __init__(self, row: int, col: int, value: str | float, format: Dict[str, str] = None):
        self.row = row
        self.col = col
        self.value = value
        self.format = format

    def __str__(self) -> str:
        return "Cell(row: {}, col: {}, value: {}, format: {}".format(self.row, self.col, self.value, self.format)


class Series:
    """Stores xlsx column data including coordinates, cells, and cell formats."""

    def __init__(self, start_row: int, col: int, category: str, styles: Styles, use_empty_header: bool = False):
        self.start_row = start_row
        self.col = col
        self.category = category
        self.formats = styles.get(category)
        header_text = '' if use_empty_header else category
        self.header_cell = Cell(start_row, col, header_text, self.get_format('header'))
        self.data: list[Cell] = []
        self.sum_cell = None
    
    def __str__(self) -> str:
        cells_str = ",".join([str(cell) for cell in self.get_cells_as_list])
        return "Series(start_row: {}, col: {}, category: {}, formats: {}, cells: {})".format(
            self.start_row, self.col, self.category, self.formats, cells_str)

    def get_format(self, format_type: str) -> Dict[str, str]:
        """Get the format `dict` for the given `format_type` from this series' formats."""
        return self.formats.get(format_type) if self.formats else None

    def get_line_styles(self) -> Dict[str, str]:
        """Get the `dict` of line formats from this series' formats."""
        return self.get_format('line')

    def get_height(self) -> int:
        """Get the column's height."""
        if self.sum_cell:
            return 1 + len(self.data) + 1
        else:
            return 1 + len(self.data)

    def get_cells_as_list(self) -> list[Cell]:
        """Get all cells in the series as a list. Includes header, data, and sum cells."""
        cell_list = [self.header_cell]
        cell_list.extend(self.data)
        if self.sum_cell:
            cell_list.append(self.sum_cell)
        return cell_list

    def append_data_cell(self, value: str | float, cell_type: str = 'data'):
        """
        Append a data cell containing `value` to this series' data.
        Optionally specify `cell_type` to use for the cell's format. Default is `data` or `alt` depending on row index.
        """
        row = self.start_row + len(self.data) + 1
        if cell_type == 'data' and row % 2 == 0:
            cell_type = 'alt'
        self.data.append(Cell(row, self.col, value, self.get_format(cell_type)))

    def append_sum_row_cell(self, start_col: int, end_col: int):
        """Append a data cell containing the SUM formula for all columns between `start_col` and `end_col`."""
        row = self.start_row + len(self.data) + 1
        xl_row = row + 1
        xl_start_col = utility.xl_col_to_name(start_col)
        xl_end_col = utility.xl_col_to_name(end_col)
        sum_formula = '=SUM({}{}:{}{})'.format(xl_start_col, xl_row, xl_end_col, xl_row)
        self.data.append(Cell(row, self.col, sum_formula, self.get_format('total')))

    def append_difference_row_cell(self, col1_index: int, col2_index: int):
        """Append a data cell containing a subtraction formula for the columns `col1_index` and `col2_index`."""
        row = self.start_row + len(self.data) + 1
        xl_row = row + 1
        xl_col1 = utility.xl_col_to_name(col1_index)
        xl_col2 = utility.xl_col_to_name(col2_index)
        diff_formula = '={}{}-{}{}'.format(xl_col1, xl_row, xl_col2, xl_row)
        self.data.append(Cell(row, self.col, diff_formula, self.get_format('total')))

    def create_sum_cell(self):
        """Create a cell that contains the SUM formula for the rows in this column."""
        row = self.start_row + len(self.data) + 1
        xl_start_sum_row = self.start_row + 2
        xl_end_sum_row = self.start_row + len(self.data) + 1
        xl_sum_col = utility.xl_col_to_name(self.col)
        sum_formula = '=SUM({}{}:{}{})'.format(xl_sum_col, xl_start_sum_row, xl_sum_col, xl_end_sum_row)
        self.sum_cell = Cell(row, self.col, sum_formula, self.get_format('total'))


class Table:
    """Stores xlsx table data including coordinates, columns and styles."""

    def __init__(self, start_row: int, start_col: int, data: Dict[str, dict], styles: Styles):
        self.start_row = start_row
        self.start_col = start_col
        self.timespans = list(data.keys())
        self.timespan_col = Series(start_row, start_col, 'timespan', styles, use_empty_header=True)
        
        for timespan in self.timespans:
            self.timespan_col.append_data_cell(timespan)

    def get_num_data_rows(self) -> int:
        """Get number of rows in the table containing data. Does not include header or sum rows."""
        return len(self.timespans)

    def get_height(self) -> int:
        """Get total height of the table."""
        return self.timespan_col.get_height()

    def get_width(self) -> int:
        """Get total width of the table."""
        pass

    def get_cols_as_lists(self) -> list[list[Cell]]:
        """Get all columns in the table represented as lists of cells."""
        pass


class ExpensesTable(Table):
    """
    Converts an expenses dictionary to a `Table`

    Expenses should take the form of `{ timespan: { category: value } }`
    """

    def __init__(self,
                 start_row: int,
                 start_col: int,
                 expenses: Dict[str, Dict[str, float]],
                 styles: Styles,
                 include_sum_row: bool = True):
        super().__init__(start_row, start_col, expenses, styles)

        self.columns: list[Series] = []
        for col_index, category in enumerate(expenses[self.timespans[0]].keys(), start=start_col + 1):
            col = Series(start_row, col_index, category, styles)
            for time in self.timespans:
                col.append_data_cell(expenses[time][category])
            self.columns.append(col)

        if include_sum_row:
            self.append_sum_cells()

    def get_width(self) -> int:
        """Get total width of the table."""
        return len(self.columns) + 1

    def get_cols_as_lists(self) -> list[list[Cell]]:
        """Get all columns in the table represented as lists of cells."""
        all_columns = [self.timespan_col.get_cells_as_list()]
        for col in self.columns:
            all_columns.append(col.get_cells_as_list())
        return all_columns

    def get_series_for_expenses_chart(self) -> list[Series]:
        """Get the series used in an expenses chart."""
        return self.columns

    def append_sum_cells(self):
        """For each data column in the table. Append a cell that sums the values in that column."""
        self.timespan_col.append_data_cell('Totals', cell_type='total')
        for col in self.columns:
            col.create_sum_cell()


class OverallTable(Table):
    """
    Converts an overall data dictionary to a `Table`

    Data should take the form of `{ timespan: { major_category: { minor_category: value } } }`
    """

    def __init__(self,
                 start_row: int,
                 start_col: int,
                 overall_data: Dict[str, Dict[str, Dict[str, float]]],
                 styles: Styles,
                 include_sum_row: bool = True):
        super().__init__(start_row, start_col, overall_data, styles)

        timespan = self.timespans[0]
        col_index = start_col + 1
        income_series_start_index = col_index
        self.income_series: list[Series] = []
        for category in overall_data[timespan]['income'].keys():
            col = Series(start_row, col_index, category, styles)
            for time in self.timespans:
                col.append_data_cell(overall_data[time]['income'][category])
            self.income_series.append(col)
            col_index += 1

        total_income_series_index = col_index
        self.total_income_series = Series(start_row, col_index, 'Total Income', styles)
        for time in self.timespans:
            self.total_income_series.append_sum_row_cell(income_series_start_index, col_index - 1)
        col_index += 1

        expenses_series_start_index = col_index
        self.expenses_series: list[Series] = []
        for category in overall_data[timespan]['expenses'].keys():
            col = Series(start_row, col_index, category, styles)
            for time in self.timespans:
                col.append_data_cell(overall_data[time]['expenses'][category])
            self.expenses_series.append(col)
            col_index += 1

        total_iexpenses_series_index = col_index
        self.total_expenses_series = Series(start_row, col_index, 'Total Expenses', styles)
        for time in self.timespans:
            self.total_expenses_series.append_sum_row_cell(expenses_series_start_index, col_index - 1)
        col_index += 1

        self.total_surplus_series = Series(start_row, col_index, 'Total Surplus', styles)
        for time in self.timespans:
            self.total_surplus_series.append_difference_row_cell(
                total_income_series_index, total_iexpenses_series_index)
        col_index += 1

        self.transfers_series: list[Series] = []
        for category in overall_data[timespan]['transfers'].keys():
            col = Series(start_row, col_index, category, styles)
            for time in self.timespans:
                col.append_data_cell(overall_data[time]['transfers'][category])
            self.transfers_series.append(col)
            col_index += 1

        self.unknown_series: list[Series] = []
        for category in overall_data[timespan]['unknown'].keys():
            col = Series(start_row, col_index, category, styles)
            for time in self.timespans:
                col.append_data_cell(overall_data[time]['unknown'][category])
            self.unknown_series.append(col)
            col_index += 1

        if include_sum_row:
            self.append_sum_cells()

    def get_width(self) -> int:
        """Get total width of the table."""
        return 1 + len(self.income_series) + len(self.expenses_series) + len(self.transfers_series) + len(self.unknown_series) + 3

    def get_cols_as_lists(self) -> list[list[Cell]]:
        """Get all columns in the table represented as lists of cells."""
        all_columns = [self.timespan_col.get_cells_as_list()]
        for col in self.income_series:
            all_columns.append(col.get_cells_as_list())
        all_columns.append(self.total_income_series.get_cells_as_list())
        for col in self.expenses_series:
            all_columns.append(col.get_cells_as_list())
        all_columns.append(self.total_expenses_series.get_cells_as_list())
        all_columns.append(self.total_surplus_series.get_cells_as_list())
        for col in self.transfers_series:
            all_columns.append(col.get_cells_as_list())
        for col in self.unknown_series:
            all_columns.append(col.get_cells_as_list())
        return all_columns

    def get_series_for_income_expenses_chart(self) -> list[Series]:
        """Get the series used in a cahrt of incomes and expenses."""
        all_series = []
        all_series.extend(self.income_series)
        all_series.extend(self.expenses_series)
        return all_series

    def get_series_for_totals_chart(self) -> list[Series]:
        """Get the series used in a chart of income and expenses totals."""
        total_series = []
        total_series.append(self.total_income_series)
        total_series.append(self.total_expenses_series)
        total_series.append(self.total_surplus_series)
        return total_series

    def append_sum_cells(self):
        """For each data column in the table. Append a cell that sums the values in that column."""
        self.timespan_col.append_data_cell('Totals', cell_type='total')
        for col in self.income_series:
            col.create_sum_cell()
        self.total_income_series.create_sum_cell()
        for col in self.expenses_series:
            col.create_sum_cell()
        self.total_expenses_series.create_sum_cell()
        self.total_surplus_series.create_sum_cell()
        for col in self.transfers_series:
            col.create_sum_cell()
        for col in self.unknown_series:
            col.create_sum_cell()
