#!/usr/bin/env python3

import sys
import os

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
    from datetime import date
    from dateutil.relativedelta import relativedelta

    from collections import defaultdict
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

            # Relevant columns
            self.DATE = 'B'
            self.DESC = 'C'
            self.VALUE = 'D'

        except Exception as e:
            logger.error("Failed to initialize.")
            logger.error("{}".format(e))
            logger.error("Exiting.")
            sys.exit()

    def sixMonthExpenses(self):
        six_months_ago = date.today() - relativedelta(months=6)

    def freqExpenses(self):
        # A year ago
        one_year = date.today() - relativedelta(years=1)
        # 3 months ago
        three_months = date.today() - relativedelta(months=3)

        # To store all expenses
        expenses = defaultdict(float)

        # This years sheet
        sheet = max(self.sheetnames)

        # Find row corresponding to 3 mos ago/1 yr ago,
        # based on comparing self.DATE w/ variables defined
        # in this function

        # Find last relevant row (no date, value, etc)

        # Concatenate beg/end rows with desc column for loop

        # Loop: for row in sheet['self.DESC+beg':'self.DESC+end']: for cell in row...
        # Current cell is desc, so must obtain value cell:
        # val = sheet[self.VALUE+cell.coordinate[1:]]
        # expenses[cell.value] += val

        # Order and display expenses
