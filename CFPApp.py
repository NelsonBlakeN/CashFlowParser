#!/usr/bin/env python3

from CashFlowParser import CashFlowParser

try:
    cfp = CashFlowParser()
    cfp.sendExpenseReport()
except Exception as e:
    print("Something went wrong in app execution:")
    print(e)
    sys.exit(1)
