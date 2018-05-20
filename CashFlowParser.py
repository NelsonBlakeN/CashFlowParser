import sys
import os
from pprint import pprint
try:
    import logging
except Exception as e:
    # Catch error with backup log file from stdout
    print("ERROR logging import failed: {}".format(e))
    sys.exit(1)

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

    from EmailUtils import *

    # Utilities
    from collections import defaultdict
    from tabulate import tabulate
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    import smtplib
except Exception as e:
    logger.error("Import error: {}".format(e))
    logger.error("Exiting.")
    sys.exit(1)

class CashFlowParser:
    def __init__(self):
        try:
            # Workbook related items
            self.book = load_workbook(WBPATH)
            self.sheetnames = self.book.sheetnames

            # Relevant columns
            self.DATE = 1
            self.DESC = 2
            self.VALUE = 3

        except Exception as e:
            logger.error("Failed to initialize.")
            logger.error("{}".format(e))
            logger.error("Exiting.")
            sys.exit()

    # Given: Excel formula (i.e.: =-1.25-8.24)
    # OR just a normal number - function will execute as normal
    def numeric(self, formula):
        return eval(formula.replace("=", ""))

    # Return all applicable worksheets (date may span
    # multiple years)
    # Date must be datetime object
    def getSheets(self, date=None):
        if date is None:
            return None

        sheets = []
        # If the date spans multiple years,
        # grab last years sheet
        if date.year != datetime.today().year:
            for name in self.sheetnames:
                if name == str(datetime.today().year):
                    sheets.append(self.book[name])

        # Grab the current years sheet
        for name in self.sheetnames:
            if name == str(date.year):
                sheets.append(self.book[name])
        return sheets

    # Takes given date, and what to order on
    # Finds all expenses within the given timeframe, sums them
    # and returns a sorted list
    def orderExpenses(self, date=None, order=None):
        if date is None or order is None:
            return None

        # To store data
        expenses = defaultdict(float)
        frequency = defaultdict(int)

        # Relevant worksheets
        # Could be > 1 (see getSheets)
        sheets = self.getSheets(date)

        # Collect data for all 1+ sheets
        for sheet in sheets:
            date_column = list(sheet.columns)[self.DATE]
            desc_column = list(sheet.columns)[self.DESC]
            value_column = list(sheet.columns)[self.VALUE]

            # Find row corresponding to given date
            start_row = None
            end_row = None
            for cell in date_column[2:]:                    # Skip first 2 cells
                if type(cell.value) is not str and start_row is None and cell.value >= date:
                    start_row = int(cell.coordinate[1:])    # Only need the row number
                if type(cell.value) is not str and cell.value is None:
                    end_row = int(cell.coordinate[1:])      # Only need the row number
                    break                                   # Finish at the first blank cell

            # Collect expenses
            for cell in desc_column[start_row:end_row-1]:
                desc = cell.value
                i = desc_column.index(cell)
                val = value_column[i].value
                if val is not None:
                    val = -self.numeric(str(val))
                    expenses[desc] += val
                    frequency[desc] += 1

        ordered_list = sorted(expenses.items(), key=lambda kv: kv[1], reverse=True)
        final_list = []
        for exp in ordered_list:
            desc = exp[0]
            total = format(exp[1], '.2f')
            freq = frequency[desc]
            avg = '%.2f' % (exp[1]/freq)
            final_list.append((desc, freq, total, avg))

        return sorted(final_list, key=lambda tup: float(tup[order]), reverse=True)

    def sixMonthExpenses(self):
        six_months = datetime.today() - relativedelta(months=6)

        # "Total" index in tuple
        TOTAL = 2

        expense_list = self.orderExpenses(date=six_months, order=TOTAL)

        expense_sum = 0
        for expense in expense_list:
            # Values were converted to strings for table formatting,
            # these have to be reverted for summing
            if float(expense[TOTAL]) > 0:
                expense_sum += float(expense[TOTAL])

        print("6 MONTH EXPENSE: {}".format(expense_sum))
        return expense_sum, tabulate(expense_list, headers=["Desc.", "Freq", "Total", "Avg"], tablefmt='orgtbl', floatfmt=".2f")

    def grossExpenses(self):
        # 3 months ago
        three_months = datetime.today() - relativedelta(months=3)

        # Tuple position for ordering
        TOTAL = 2

        gross_list = self.orderExpenses(three_months, order=TOTAL)

        return tabulate(gross_list, headers=["Desc.", "Freq", "Total", "Avg"], tablefmt='orgtbl', floatfmt=".2f")

    def freqExpenses(self):
        # 3 months ago
        three_months = datetime.today() - relativedelta(months=3)

        # Tuple positions, for ordering
        FREQUENCY = 1

        frequency_list = self.orderExpenses(date=three_months, order=FREQUENCY)

        return tabulate(frequency_list, headers=["Desc.", "Freq", "Total", "Avg"], tablefmt='orgtbl', floatfmt=".2f")

    def sendMail(self, content):
        msg = MIMEMultipart("alternative", None, [MIMEText(content)])

        msg['Subject'] = "Weekly Expense Report"
        msg['From'] = FROM
        msg['To'] = TO
        server = smtplib.SMTP(SERVER)
        server.ehlo()
        server.starttls()
        print("Signing into gmail with " + TO + ", " + PASSWORD)
        server.login(TO, PASSWORD)    # Authenticate with the actual gmail account
        server.sendmail(FROM, TO, msg.as_string())
        server.quit()

    def sendExpenseReport(self):
        frequency_table = self.freqExpenses()
        gross_expense_table = self.grossExpenses()
        six_month_sum, six_month_table = self.sixMonthExpenses()
        content = EMAILTXT.format(frequency_table, gross_expense_table, six_month_sum, six_month_table)

        self.sendMail(content=content)
