from datetime import datetime
import pandas as pd
import PySimpleGUI as sg
import os
from PyPDF2 import PdfReader
import re

# completed exchange parse statement
# completed exchange transactions
# completed chase transactions
# NEXT chase parse statements
# need to: automatically load all statements and either merge or create new book


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
                    "Remodel/Maintenance": ("lowe", "heating", 'frontier fire', 'pest'),
                    "Miscellaneous": {},
                    "Depreciation": {},
                    "Amortization": {},
                    "Sales": "merch dep",
                    "Deposits": 'deposit',
                    "Non Deductible": {}
                    })


def makeWindow(theme):
    sg.theme(theme)
    layout = [
        [sg.Text('Year End Reporting', size=(35, 1), justification='center', relief=sg.RELIEF_RIDGE)],
        [sg.HorizontalSeparator()],
        [sg.Radio('Exchange', 'bank', default=True, k='-exchange-'),
         sg.Radio('Chase', 'bank', default=False, k='-chase-')],
        [sg.Radio('Transaction downloads', 'type', k='-transactions-', default=True),
         sg.Radio('Bank statements', 'type', k='-statements-', )],
        [sg.Button('Bank transactions', k='-bank-')],
        [sg.Button('Excel bookkeeper', k='-book-'), sg.Column([[]], expand_x=True),
         sg.Button('Process', k='-process-'), sg.Button('Exit', k='exit')],
        [sg.StatusBar('', size=60, k='-status-')]
    ]
    window = sg.Window('Reports', layout, grab_anywhere=True, finalize=True)
    return window


def chaseParseTransactions(bank, dfBank):
    cols = ['Details', 'Posting Date', 'Description', 'Amount', 'Check or Slip #']
    transactions = pd.read_csv(bank, index_col=False, usecols=cols)
    transactions.columns = ['type', 'date', 'desc', 'amount', 'check#']
    newDf = []
    for index, row in transactions.iterrows():
        identified = False
        # break out of loop when no more row in transactions
        if isinstance(row['type'], float):
            break
        desc = row['desc'].lower()
        # create new row template
        newRow = {'Category': '',
                  'type': row['type'],
                  'date': row['date'],
                  'description': desc,
                  'amount': row['amount'],
                  'check#': row['check#']
                  }

        if row['type'] == "CREDIT":
            newRow['Category'] = 'Sales'
            identified = True
        elif row['type'] == 'CHECK':
            newRow['Category'] = 'Wages'
            identified = True
        elif row['type'] == 'DSLIP':
            newRow['type'] = 'Deposits'
            identified = True
        else:
            # now we figure out what Category expense
            for category in description.keys():
                values = description[category]
                if isinstance(values, str):
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
        newDf.append(pd.DataFrame(newRow, index=[0]))
    dfBank = pd.concat(newDf, ignore_index=True)
    return dfBank


def exchangeParseTransactions(bank, dfBank):
    cols = [' Posted Date', ' Description', ' Debit', ' Credit', ' Check No.']
    transactions = pd.read_csv(bank, index_col=False, usecols=cols)
    transactions.columns = ['date', 'desc', 'debit', 'credit', 'checkNum']
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
                  'amount': row['credit'] if row['debit'] != row['debit'] else -abs(row['debit']),
                  'check#': ''
                  }
        if row['debit'] != row['debit']:
            # fastest way to check if float equals 'nan' is if it DOES NOT EQUAL itself
            # in this case check to see if debit value is Nan,
            # if true then row['debit'] is nan and row['credit'] has a value
            # a = float('nan') if (a!=a)
            newRow['Category'] = 'Sales' if description['Sales'] in desc else 'Deposits'
            newRow['type'] = 'Credit'
            identified = True
        else:
            newRow['type'] = 'Debit'
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
        newDf.append(pd.DataFrame(newRow, index=[0]))
    dfBank = pd.concat(newDf, ignore_index=True)
    return dfBank


def exchangeParseStatements(statement, dfBank):
    # creating a pdf reader object
    reader = PdfReader(statement)
    newDf = []
    firstPage = reader.pages[0].extract_text()
    statementDate = re.search('\d+/\d+/\d+', firstPage).group(0)
    year = statementDate[-2:]
    for num in range(len(reader.pages)):
        page = reader.pages[num]
        text = page.extract_text()
        lines = re.split('\n', text)
        for lineNumber, l1 in enumerate(lines):
            if not isinstance(l1, str):
                continue
            validTransaction = re.match('^\s*\d+/\d+\s\w+', l1)
            if not validTransaction:
                continue
            l1 = l1.lower()
            l2 = ''
            date = ''
            desc = ''
            amt = ''
            checkNum = ''
            merchant = ''
            if 'check' in l1:
                l2 = re.match(
                    '\s+(?P<date>\d+/\d\d)\s(?P<desc>.*?(?=\s{4}))\s(?P<checkNum>.*?(?=\s{4}))\s*(?P<amt>\d*,?\d+\.\d\d-?)',
                    l1)
                checkNum = l2.group('checkNum')
                if len(checkNum) > 4:
                    checkNum = checkNum[-4:]
            elif 'cable' in l1:
                l2 = re.match('\s+(?P<date>\d+/\d+)\s(?P<desc>.*?(?=\s{4}))\s*(?P<amt>\d+\.\d\d-?)', l1)
            elif ' pos ' in l1 or ' dbt ' in l1:
                date = re.match('\s+(?P<date>\d+/\d+)', l1).group('date')
                amt = re.search('\d*,?\d+\.\d\d-?', l1).group(0)
                merchant = re.match('\s+(?P<merch>.*?(?=\s{2}))', lines[lineNumber + 1])
            elif 'service charge' in l1:
                l2 = re.match('\s+(?P<date>\d+/\d+)\s*(?P<desc>.*?(?=\s{4}))\s*(?P<amt>\d*\.\d\d-?)', l1)
            else:
                l2 = re.match('\s+(?P<date>\d+/\d+)\s*(?P<desc>.*?(?=\s{4}))\s*(?P<amt>\d*,?\d+\.\d\d-?)', l1)
                if not l2:
                    continue

            if ' pos ' in l1 or ' dbt ' in l1:
                desc = merchant.group('merch').lower()
            else:
                date = l2.group('date')
                desc = l2.group('desc')
                amt = l2.group('amt')
            if amt[-1] == '-':
                amt = '-' + amt[0:-1]
            # create new row template
            newRow = {'Category': '',
                      'type': 'Debit' if float(amt.replace(",", "")) < 0 else 'Credit',
                      'date': date + '/' + year,
                      'description': desc,
                      'amount': float(amt.replace(",", "")),
                      'check#': checkNum
                      }
            identified = False
            for category in description.keys():
                values = description[category]
                if isinstance(values, str):  # if only one value
                    if values in desc:
                        newRow['Category'] = category
                        identified = True
                        break
                else:  # has a list of values
                    for value in values:
                        if value in desc:
                            newRow['Category'] = category
                            identified = True
                            break
            if not identified:
                newRow['Category'] = 'Miscellaneous'
            newDf.append(pd.DataFrame(newRow, index=[0]))
    dfBank = pd.concat(newDf, ignore_index=True)
    return dfBank


def exportToExcel(outputExcel, dfBank, initial):
    with pd.ExcelWriter(outputExcel, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        dfBank.to_excel(writer, sheet_name=initial + '.bank', header=None, index=False,
                        startrow=writer.sheets[initial + '.bank'].max_row)


def main():
    window = makeWindow(sg.theme())
    bank = 'chase.csv'
    book = '2024taxCat - Copy.xlsx'
    transactions = ''
    dfBank = pd.DataFrame(columns=['Category', 'type', 'date', 'description', 'amount', 'check#'])
    while True:
        event, values = window.read()
        if event not in (sg.TIMEOUT_EVENT, sg.WIN_CLOSED, 'exit'):
            print('============ Event = ', event, ' ==============')
            print('-------- Values Dictionary (key=value) --------')
            for key in values:
                print(key, ' = ', values[key])
        if event == '-bank-':
            bank = sg.popup_get_file('Bank transactions/statements', 'Choose bank info', initial_folder=os.getcwd())
        elif event == '-book-':
            book = sg.popup_get_file('Excel book', 'Choose excel book', initial_folder=os.getcwd())
        elif event == '-process-':
            bankInitial = 't' if values['-exchange-'] else 'e'
            result = ''
            if values['-transactions-']:
                if values['-exchange-']:
                    result = exchangeParseTransactions(bank, dfBank)
                else:
                    result = chaseParseTransactions(bank, dfBank)
            elif values['-statements-']:
                if values['-exchange-']:
                    result = exchangeParseStatements(bank, dfBank)
                else:
                    # result = chaseParseStatements(bank, dfBank)
                    continue

            exportToExcel(book, result, bankInitial)
            window['-status-'].update('Process complete')
        else:
            window.close()
            exit(0)


if __name__ == '__main__':
    sg.theme('dark grey 14')
    main()

# # bank statement
# bank = "Chase7668_Activity_20240324.csv"
# # excel bookkeeper
# book = "book2024 - Copy.xlsx"
#
# # load bank sheet
# loadedBank = pd.read_excel(book, sheet_name="y.bank", )
# dfBank = pd.DataFrame(columns=loadedBank.columns)
# transactions = pd.read_csv(bank, index_col=False)
#
# bank = {'chase': False, 'exchange': False}
# bankNum = input("press 1 for chase or 2 for exchange bank\n")
# statement = False
# if bankNum == '1':
#     bank['chase'] = True
# else:
#     bank['exchange'] = True
#     ask = input('press 1 for statement or 2 for downloaded transactions:')
#     statement = True if ask == '1' else False
#
# result = ''
# if bank['chase']:
#     result = chaseParseTransactions(transactions, dfBank)
#     exportToExcel(book, result, 'y')
# else:
#     if statement:
#         exchangeParseStatements()
#     else:
#         exchangeParseTransactions(transactions)
