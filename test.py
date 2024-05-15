# importing required modules
from PyPDF2 import PdfReader
import re
import pandas as pd
from datetime import datetime


officeExpenses = ("samsclub", "sams club", "walmart",
                  'walmart.com', "wal-mart", "amzn",
                  'amazon', "best buy", "big lots",
                  "liquor", "hobby-lobby", "bath & body",
                  "staples", "joann", "sally beauty",
                  'wal sam', 'locked up', 'bestbuy', 'hobbylobby'
                  )
description = dict({"Rent": "robson",
                    "Merchant Fees": ("mthly disc direct payment", 'direct dps', 'hs group'),
                    "Bank Fees": "service charge",
                    "Cable": ("optimum", "suddenlink"),
                    "Utilities": ("ok natural gas", "city of stillwater"),
                    "Insurance": "insurance",
                    "Marketing": ("facebk", "google", "college coupon", 'metaplatfor'),
                    "Office Expenses": officeExpenses,
                    "License fees": ("secretary of state", "osbcb"),
                    "Supplies": {'supply', 'nailsjobs', 'nails plus'},
                    "Wages": 'check',
                    "Taxes": ("irs", "tax", 'oklahomataxpmts'),
                    "Remodel/Maintenance": ("lowe", "heating", 'frontier fire'),
                    "Miscellaneous": {},
                    "Depreciation": {},
                    "Amortization": {},
                    "Sales": "merch dep",
                    "Deposit": 'deposit',
                    "Non Deductible": {}
                    })


bank = 'jan-feb.csv'
book = '2024taxCat - Copy.xlsx'
cols = [' Posted Date', ' Description', ' Debit', ' Credit', ' Check No.']
transactions = pd.read_csv(bank, index_col=False, usecols=cols)
transactions.columns = ['date', 'desc', 'debit', 'credit', 'checkNum']
dfBank = pd.DataFrame(columns=['Category', 'type', 'date', 'description', 'amount', 'check#'])
newDf = []
for index, row in transactions.iterrows():
    identified = False
    # break out of loop when no more row in transactions
    if isinstance(row['desc'], float):
        break
    desc = row['desc'].lower()
    # create new row template
    newRow = {'Category': '',
              'type': row["desc"],
              'date': row["date"],
              'description': desc,
              'amount': '',
              'check#': ''
              }
    if row['debit'] != row['debit']:
        # fastest way to check if float equals 'nan' is if it DOES NOT EQUAL itself
        # in this case check to see if debit value is Nan, if true then row['credit'] has a value
        newRow['Category'] = 'Sales' if description['Sales'] in desc else 'Deposit'
        newRow['type'] = 'Credit'
        newRow['amount'] = float(row['credit'])
        identified = True
    else:
        newRow['type'] = 'Debit'
        newRow['amount'] = -abs(row['debit'])
        if 'check' in desc:
            newRow['Category'] = 'Wages'
            newRow['check#'] = '' if row['checkNum'] != row['checkNum'] else int(row['checkNum'])
            identified = True
        else:
            for category in description.keys():
                values = description[category]
                if isinstance(values, str):  # if only one category
                    if values in desc:
                        newRow['Category'] = category
                        identified = True
                        break
                else:
                    for value in values:
                        if value in desc.lower():
                            newRow['Category'] = category
                            identified = True
                            break
    if not identified:
        newRow['Category'] = 'Miscellaneous'
        if row['debit']:
            newRow['amount'] = -abs(row['debit'])
        else:
            newRow['amount'] = float(row['credit'])
    newDf.append(pd.DataFrame(newRow, index=[0]))
    print(newRow)