from typing import Dict
from xlsxwriter import utility
from writers.styles import Styles


class Cell:

    def __init__(self, row: int, col: int, value: str | float, format: Dict[str, str] = None):
        self.row = row
        self.col = col
        self.value = value
        self.format = format

    def __str__(self) -> str:
        return "Cell(row: {}, col: {}, value: {}, format: {}".format(self.row, self.col, self.value, self.format)


class Series:

    def __init__(self,
                 start_row: int,
                 col: int,
                 category: str,
                 styles: Styles):
        self.start_row = start_row
        self.col = col
        self.category = category
        self.formats = styles.get(category)
        self.header_cell = Cell(start_row, col, category, self.get_format('header'))
        self.data: list[Cell] = []
        self.sum_cell = None

    def get_format(self, format_type: str) -> Dict[str, str]:
        return self.formats.get(format_type) if self.formats else None

    def get_line_styles(self) -> Dict[str, str]:
        return self.get_format('line')

    def get_cells_as_list(self) -> list[Cell]:
        cell_list = [self.header_cell]
        cell_list.extend(self.data)
        if self.sum_cell:
            cell_list.append(self.sum_cell)
        return cell_list

    def append_data_cell(self, value: float):
        row = self.start_row + len(self.data) + 1
        if row % 2 == 0:
            self.data.append(Cell(row, self.col, value, self.get_format('alt')))
        else:
            self.data.append(Cell(row, self.col, value, self.get_format('data')))

    def append_sum_row_cell(self, start_col: int, end_col: int):
        row = self.start_row + len(self.data) + 1
        xl_row = row + 1
        xl_start_col = utility.xl_col_to_name(start_col)
        xl_end_col = utility.xl_col_to_name(end_col)
        sum_formula = '=SUM({}{}:{}{})'.format(xl_start_col, xl_row, xl_end_col, xl_row)
        self.data.append(Cell(row, self.col, sum_formula, self.get_format('total')))

    def append_difference_row_cell(self, col1_index: int, col2_index: int):
        row = self.start_row + len(self.data) + 1
        xl_row = row + 1
        xl_col1 = utility.xl_col_to_name(col1_index)
        xl_col2 = utility.xl_col_to_name(col2_index)
        diff_formula = '={}{}-{}{}'.format(xl_col1, xl_row, xl_col2, xl_row)
        self.data.append(Cell(row, self.col, diff_formula, self.get_format('total')))

    def create_sum_cell(self):
        row = self.start_row + len(self.data) + 1
        xl_start_sum_row = self.start_row + 2
        xl_end_sum_row = self.start_row + len(self.data) + 1
        xl_sum_col = utility.xl_col_to_name(self.col)
        sum_formula = '=SUM({}{}:{}{})'.format(xl_sum_col, xl_start_sum_row, xl_sum_col, xl_end_sum_row)
        self.sum_cell = Cell(row, self.col, sum_formula, self.get_format('total'))


class Table:

    def __init__(self, start_row: int, start_col: int, data: dict):
        self.start_row = start_row
        self.start_col = start_col
        self.timespans = list(data.keys())
        self.timespan_col: list[Cell] = []

    def get_num_data_rows(self) -> int:
        return len(self.timespans)
    
    def get_height(self) -> int:
        return self.start_row + 1 + len(self.timespan_col)
    
    def get_width(self) -> int:
        pass

    def get_series_for_chart(self) -> list[Series]:
        pass

    def get_cols_as_lists(self) -> list[list[Cell]]:
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
        super().__init__(start_row, start_col, expenses)

        timespan_row_index = start_row + 1
        for timespan in self.timespans:
            self.timespan_col.append(Cell(timespan_row_index, start_col, timespan))
            timespan_row_index += 1
        if include_sum_row:
            self.timespan_col.append(Cell(timespan_row_index, start_col, 'Totals'))

        self.columns: list[Series] = []
        for col_index, category in enumerate(expenses[self.timespans[0]].keys(), start=start_col + 1):
            col = Series(start_row, col_index, category, styles)
            for time in self.timespans:
                col.append_data_cell(expenses[time][category])
            if include_sum_row:
                col.create_sum_cell()
            self.columns.append(col)

    def get_width(self) -> int:
        return len(self.columns) + 1

    def get_series_for_chart(self) -> list[Series]:
        return self.columns

    def get_cols_as_lists(self) -> list[list[Cell]]:
        all_columns = [self.timespan_col]
        for col in self.columns:
            all_columns.append(col.get_cells_as_list())
        return all_columns


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
                 include_sum_row: bool = False):
        super().__init__(start_row, start_col, overall_data)

        timespan_row_index = start_row + 1
        for timespan in self.timespans:
            self.timespan_col.append(Cell(timespan_row_index, start_col, timespan))
            timespan_row_index += 1

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
            self.total_surplus_series.append_difference_row_cell(total_income_series_index, total_iexpenses_series_index)
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
    
    def get_width(self) -> int:
        return 1 + len(self.income_series) + len(self.expenses_series) + len(self.transfers_series) + len(self.unknown_series) + 3

    def get_series_for_chart(self) -> list[Series]:
        all_series = []
        all_series.extend(self.income_series)
        all_series.extend(self.expenses_series)
        return all_series

    def get_cols_as_lists(self) -> list[list[Cell]]:
        all_columns = [self.timespan_col]
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
