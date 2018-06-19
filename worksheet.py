#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jun 18 22:14:37 2018

@author: ahura
"""

from collections import defaultdict
from openpyxl import load_workbook, worksheet

# The cell describing the meal, e.g. "dinner saturday"
_mealcell = "B2"

# Cell describing the dish, e.g. "Spaghet"
_dishcell = "B3"

# Cell that contains the number of people partaking in the meal
_n_peoplecell = "B5"

# Row number where ingredients start appearing
first_ing_row = 13

ingredient_col = "A"
amount_col = "C"
unit_col = "D"

def convert_units(amount, unit):
    if unit == "g":
        return 0.001*amount, "kg"

    return amount, unit

def parse_spreadsheet(filename):
    '''Takes the name of a worksheet file and parses it. Returns a tuple of
    two dicts.
    First one maps ingredients to dict like {"kg": 42} (units to quantities).
    Second one maps ingredients to list of which dishes they're used in.'''

    # Initialize
    wb = load_workbook(filename=filename, data_only = True)
    data = defaultdict(lambda: defaultdict(lambda: 0.0))
    details = defaultdict(lambda: [])

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        meal = ws[_mealcell].value
        dish = ws[_dishcell].value
        print("Parsing sheet: %s. Recipe for %s (%s)"
              % (sheet_name, dish, meal))

        # Extract the number of people who'll be eating this meal
        n_people = ws[_n_peoplecell].value

        for i, row in enumerate(ws.iter_rows()):
            #  Skip rows until we hit the ingredient list
            row_num = i + 1
            if row_num < first_ing_row:
                continue

            # Extract ingredient, amount, and unit.
            col2val = {c.column: c.value for c in row}
            needed = (ingredient_col, amount_col, unit_col)
            if any(c not in col2val or col2val[c] is None for c in needed):
                continue
            ingredient = col2val[ingredient_col]
            amount_raw = n_people*col2val[amount_col]
            unit_raw = col2val[unit_col]
            amount, unit = convert_units(amount_raw, unit_raw)

            # Update data
            data[ingredient][unit] += amount
            detail = "%s (%s)" % (dish, meal)
            details[ingredient].append(detail)
        #

    return data, details

if __name__ == '__main__':
    a, b = parse_spreadsheet("recipes.xlsx")
    print(a)