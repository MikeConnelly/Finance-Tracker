import copy
import json
from calendar import monthrange
from datetime import datetime
from typing import Dict


class FinanceData:

    def __init__(self, default_values: Dict[str, Dict[str, float]]):
        self.data = {}
        self.default_values = default_values

    def __str__(self):
        return json.dumps(self.data, indent=4)

    def get_monthly_totals(self) -> Dict[str, Dict[str, Dict[str, float]]]:
        """Get the toals for each major and minor category for every month in this data."""
        monthly_totals = {}
        for year in self.data.keys():
            for month in self.data[year].keys():
                month_key = f'{year}/{month}'
                monthly_totals[month_key] = copy.deepcopy(self.default_values)
                for day in self.data[year][month]:
                    for major in self.data[year][month][day]:
                        for minor in self.data[year][month][day][major]:
                            monthly_totals[month_key][major][minor] += self.data[year][month][day][major][minor]
                            monthly_totals[month_key][major][minor] = round(monthly_totals[month_key][major][minor], 2)
        return monthly_totals

    def get_monthly_expenses_totals(self) -> Dict[str, Dict[str, float]]:
        """Get the totals for each minor expenses category for every month in this data."""
        default_expenses_values = self.default_values['expenses']
        monthly_expenses_totals = {}
        for year in self.data.keys():
            for month in self.data[year].keys():
                month_key = f'{year}/{month}'
                monthly_expenses_totals[month_key] = copy.deepcopy(default_expenses_values)
                for day in self.data[year][month]:
                    for minor in self.data[year][month][day]['expenses']:
                        monthly_expenses_totals[month_key][minor] += self.data[year][month][day]['expenses'][minor]
                        monthly_expenses_totals[month_key][minor] = round(monthly_expenses_totals[month_key][minor], 2)
        return monthly_expenses_totals

    def get_months(self):
        """Get the list of tuples containing all year, month combos present."""
        return [(year, month) for year in self.data.keys() for month in self.data[year].keys()]

    def get_daily_expenses(self, year, month) -> Dict[str, Dict[str, float]]:
        """Get the contents of the expenses category for every day of a given month and year."""
        month_map = copy.deepcopy(self.data[year][month])
        for day in month_map:
            month_map[day] = month_map[day]['expenses']
        return month_map

    def add_value(self, date: datetime, major_category: str, minor_category: str, amount: float):
        """Add an amount to the current value for the given date and categories."""
        self.add_date_if_not_exists(date)
        self.data[date.year][date.month][date.day][major_category][minor_category] += amount

    def add_date_if_not_exists(self, date: datetime):
        """Add a new date if it does not already exist in data."""
        # create new year
        if date.year not in self.data.keys():
            self.data[date.year] = {}
        # create new month
        if date.month not in self.data[date.year].keys():
            # populate each day of the month with default values
            _, num_days = monthrange(date.year, date.month)
            self.data[date.year][date.month] = {
                day: copy.deepcopy(self.default_values)
                for day in range(1, num_days + 1)
            }
