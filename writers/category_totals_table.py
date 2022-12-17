from typing import Dict


DEFAULT_STYLES = [
    {
        "header_bg_color": "e16463",
        "bg_color": "f4cdcc",
        "alt_bg_color": "d3b1af"
    },
    {
        "header_bg_color": "92c57c",
        "bg_color": "daf2d8",
        "alt_bg_color": "acbfaa"
    },
    {
        "header_bg_color": "f3b369",
        "bg_color": "fae5cd",
        "alt_bg_color": "c6b6a3"
    },
    {
        "header_bg_color": "6f9eeb",
        "bg_color": "c9dbf9",
        "alt_bg_color": "9faec5"
    },
    {
        "header_bg_color": "fad866",
        "bg_color": "fcf2cd",
        "alt_bg_color": "c9c1a3"
    },
    {
        "header_bg_color": "8e7cc2",
        "bg_color": "dbd2ea",
        "alt_bg_color": "aba4b6"
    }
]


class DataColumn:

    def __init__(self, header: str):
        self.header = header
        self.data: list[float] = []
    
    def append_data(self, value: float):
        self.data.append(value)


class CategoryTotalsTable:
    """
    Converts an expenses dictionary to a Table

    Expenses should take the form of { timespan: { category: value } }
    """
    
    def __init__(self, expenses: Dict[str, Dict[str, float]]):
        self.timespan_col = list(expenses.keys())
        self.num_data_rows = len(self.timespan_col)
        self.data_cols: list[DataColumn] = []
        for category in expenses[self.timespan_col[0]].keys():
            col = DataColumn(category)
            for time in self.timespan_col:
                col.append_data(expenses[time][category])
            self.data_cols.append(col)
        self.num_data_cols = len(self.data_cols)

