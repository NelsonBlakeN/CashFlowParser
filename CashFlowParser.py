#!/usr/bin/env python3

import sys
import os
from pprint import pprint
try:
    import logging
except Exception as e:
    # Catch error with backup log file from stdout
    print("ERROR logging import failed: {}".format(e))
    sys.exit()

# Create logging utility
LOGFILE = "/home/blake/Documents/logs/CashFlowParser.log"
logging.basicConfig(filename=LOGFILE, level=logging.INFO)
logger = logging.getLogger("CashFlowParser")

try:
    # Import xlsx module
    from openpyxl import load_workbook

    # Import workbook path (private)
    from Workbook import *

    # Date related imports
    from datetime import date, datetime
    from dateutil.relativedelta import relativedelta

    from collections import defaultdict
    from tabulate import tabulate
except Exception as e:
    # logger.error("Import error: {}".format(e))
    # logger.error("Exiting.")
    print("Bad import")
    sys.exit()

class Main:
    def __init__(self):
        try:
            # Workbook related items
            self.book = load_workbook(WBPATH)
            self.sheetnames = self.book.sheetnames
            for name in self.sheetnames:
                if name == str(datetime.now().year):
                    sheet = name
            self.sheet = self.book[sheet]

            # Relevant columns
            self.DATE = 1
            self.DESC = 2
            self.VALUE = 3
            self.date_column = list(self.sheet.columns)[self.DATE]
            self.desc_column = list(self.sheet.columns)[self.DESC]
            self.value_column = list(self.sheet.columns)[self.VALUE]

        except Exception as e:
            logger.error("Failed to initialize.")
            logger.error("{}".format(e))
            logger.error("Exiting.")
            sys.exit()

    # Given: Excel formula (i.e.: =-1.25-8.24)
    # Converts to integer value, or converts neg numbers positive
    def numeric(self, formula):
        if type(formula) is int and formula > 0:
            return -formula
        formula = str(formula)
        return eval(formula[1:])

    def sixMonthExpenses(self):
        six_months_ago = date.today() - relativedelta(months=6)

    def freqExpenses(self):
        # A year ago
        one_year = datetime.today() - relativedelta(years=1)
        # 3 months ago
        three_months = datetime.today() - relativedelta(months=3)

        # To store all expenses and track frequency
        expenses = defaultdict(float)
        frequency = defaultdict(int)

        # Find row corresponding to 3 mos ago/1 yr ago
        start_row = None
        for cell in self.date_column[2:]:               # Skip first 2 cells
            if type(cell.value) is not str and start_row is None and cell.value >= three_months:
                start_row = int(cell.coordinate[1:])    # Only need the row num
            if type(cell.value) is not str and cell.value is None:
                end_row = int(cell.coordinate[1:])      # Only need the row num
                break                                   # Finish at the first blank cell

        # Collect expenses
        for cell in self.desc_column[start_row:end_row-1]:
            desc = cell.value
            i = self.desc_column.index(cell)
            val = self.value_column[i].value
            if val is not None:
                val = self.numeric(val)
                expenses[desc] += val
                frequency[desc] += 1

        # Order and display expenses
        ordered_list = sorted(expenses.items(), key=lambda kv: (kv[1]), reverse=True)
        final_list = []
        for exp in ordered_list:
            desc = exp[0]
            total = format(exp[1], '.2f')
            freq = frequency[desc]
            avg = "%.2f" % (exp[1]/freq)
            final_list.append((desc, freq, total, avg))

        most_frequent_list = sorted(final_list, key=lambda tup: float(tup[1]), reverse=True)
        gross_list = sorted(final_list, key=lambda tup: float(tup[2]), reverse=True)

        print(tabulate(most_frequent_list, headers=["Desc.", "Freq", "Total", "Avg"], tablefmt='orgtbl', floatfmt=".2f"))
        print()
        print(tabulate(gross_list, headers=["Desc.", "Freq", "Total", "Avg"], tablefmt='orgtbl', floatfmt=".2f"))

if __name__ == "__main__":
    m = Main()
    m.freqExpenses()
