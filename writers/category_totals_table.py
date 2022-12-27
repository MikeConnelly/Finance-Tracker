from typing import Dict
from xlsxwriter import utility


DEFAULT_STYLES = [
    {
        "header": {
            "bg_color": "e16463"
        },
        "data": {
            "bg_color": "f4cdcc"
        },
        "alt": {
            "bg_color": "d3b1af"
        },
        "total": {
            "bg_color": "e16463"
        }
    },
    {
        "header": {
            "bg_color": "92c57c"
        },
        "data": {
            "bg_color": "daf2d8"
        },
        "alt": {
            "bg_color": "acbfaa"
        },
        "total": {
            "bg_color": "92c57c"
        }
    },
    {
        "header": {
            "bg_color": "f3b369"
        },
        "data": {
            "bg_color": "fae5cd"
        },
        "alt": {
            "bg_color": "c6b6a3"
        },
        "total": {
            "bg_color": "f3b369"
        },
    },
    {
        "header": {
            "bg_color": "6f9eeb"
        },
        "data": {
            "bg_color": "c9dbf9"
        },
        "alt": {
            "bg_color": "9faec5"
        },
        "total": {
            "bg_color": "6f9eeb"
        },
    },
    {
        "header": {
            "bg_color": "fad866"
        },
        "data": {
            "bg_color": "fcf2cd"
        },
        "alt": {
            "bg_color": "c9c1a3"
        },
        "total": {
            "bg_color": "fad866"
        },
    },
    {
        "header": {
            "bg_color": "8e7cc2"
        },
        "data": {
            "bg_color": "dbd2ea"
        },
        "alt": {
            "bg_color": "aba4b6"
        },
        "total": {
            "bg_color": "8e7cc2"
        },
    }
]


class Cell:

    def __init__(self, row: int, col: int, value: str | float, format: Dict[str, str] = None):
        self.row = row
        self.col = col
        self.value = value
        self.format = format


class Series:

    def __init__(self,
                 start_row: int,
                 col: int,
                 category: str,
                 styles: Dict[str, Dict[str, Dict[str, str]]]):
        self.start_row = start_row
        self.col = col
        self.formats = styles[category] if category in styles.keys() else DEFAULT_STYLES[col % len(DEFAULT_STYLES)]
        self.header_cell = Cell(start_row, col, category, self.formats['header'])
        self.data: list[Cell] = []

    def append_data_cell(self, value: float):
        row = self.start_row + len(self.data) + 1
        if row % 2 == 0:
            self.data.append(Cell(row, self.col, value, self.formats['alt']))
        else:
            self.data.append(Cell(row, self.col, value, self.formats['data']))

    def create_sum_cell(self):
        row = self.start_row + len(self.data) + 1
        xl_start_sum_row = self.start_row + 2 
        xl_end_sum_row = self.start_row + len(self.data) + 1
        xl_sum_col = utility.xl_col_to_name(self.col)
        sum_formula = '=SUM({}{}:{}{})'.format(xl_sum_col, xl_start_sum_row, xl_sum_col, xl_end_sum_row)
        self.sum_cell = Cell(row, self.col, sum_formula, self.formats['total'])
    
    def get_cells_as_list(self) -> list[Cell]:
        cell_list = [self.header_cell]
        cell_list.extend(self.data)
        if self.sum_cell:
            cell_list.append(self.sum_cell)
        return cell_list


class CategoryTotalsTable:
    """
    Converts an expenses dictionary to a Table

    Expenses should take the form of { timespan: { category: value } }
    """

    def __init__(self,
                 start_row: int,
                 start_col: int,
                 expenses: Dict[str, Dict[str, float]],
                 styles: Dict[str, Dict[str, str]],
                 include_sum_row: bool = True):
        self.start_row = start_row
        self.start_col = start_col
        self.timespans = list(expenses.keys())
        self.timespan_col: list[Cell] = []
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

    def get_cols_as_lists(self) -> list[list[Cell]]:
        all_columns = [self.timespan_col]
        for col in self.columns:
            all_columns.append(col.get_cells_as_list())
        return all_columns

    def get_data_cols_as_lists(self) -> list[list[Cell]]:
        data_columns = []
        for col in self.columns:
            data_columns.append(col.get_cells_as_list())
        return data_columns
