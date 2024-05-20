import pandas as pd
import PySimpleGUI as sg
import os
from pypdf import PdfReader
import re
import shutil



# working on comparing already imported and new transactions to filter out duplicates only import new ones
# need to: automatically load all statements and either merge or create new book


officeExpenses = ("samsclub", "sams club", "walmart",
                  'walmart.com', "wal-mart", "amzn",
                  'amazon', "best buy", "big lots",
                  "liquor", "hobby-lobby", "bath & body",
                  "staples", "joann", "sally beauty",
                  'wal sam', 'locked up', 'bestbuy', 'hobbylobby'
                  )
description = dict({"Rent": "robson",
                    "Merchant Fees": ("mthly disc direct payment", 'direct dps', 'hs group', 'mthly discsec'),
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
        [sg.Radio('Exchange', 'bank', default=False, k='-exchange-'),
         sg.Radio('Chase', 'bank', default=True, k='-chase-')],
        [sg.Radio('Transaction downloads', 'type', k='-transactions-', default=False),
         sg.Radio('Bank statements', 'type', k='-statements-', default=True)],
        [sg.Button('Bank transactions', k='-bank-'), sg.T('', key='bank')],
        [sg.Button('Excel bookkeeper', k='-book-'), sg.T('', key='book'), sg.Column([[]], expand_x=True)],
        [sg.T(expand_x=True), sg.Button('Automate', k='-auto-'), sg.Button('Process', k='-process-'),
         sg.Button('Exit', k='exit')],
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


def chaseParseStatements(statement, dfBank):
    """
        Chase statement is tricky because date and amount is not always in the same line as in exchange
        statements. Even statement dates are not always in the same line on page.
        First section of statement are deposits and additions, then checks paid, then atm & debit card withdrawals,
        then electronic withdrawals,
        :param statement: [str] path to pdf file
        :param dfBank: [dataframe] empty dataframe with pre-set column names
        :return: dfBank: [dataframe] parsed transactions and formatted into dataframe
    """
    reader = PdfReader(statement)
    newDf = []
    firstPage = reader.pages[0].extract_text()
    firstPage = re.split('\n', firstPage)
    statementDate = ''
    # test variables for debugging
    # page1 = re.split('\n', reader.pages[0].extract_text())
    # page2 = re.split('\n', reader.pages[1].extract_text())
    # page3 = re.split('\n', reader.pages[2].extract_text())
    # page4 = re.split('\n', reader.pages[3].extract_text())
    # page5 = re.split('\n', reader.pages[4].extract_text())
    # page6 = re.split('\n', reader.pages[5].extract_text())
    # page7 = re.split('\n', reader.pages[6].extract_text())
    # page8 = re.split('\n', reader.pages[7].extract_text())
    # page1a = reader.pages[0].extract_text()
    # page2a = reader.pages[1].extract_text()
    # page3a = reader.pages[2].extract_text()
    # page4a = reader.pages[3].extract_text()
    # page5a = reader.pages[4].extract_text()
    # page6a = reader.pages[5].extract_text()
    # page7a = reader.pages[6].extract_text()
    # page8a = reader.pages[7].extract_text()
    # print()
    def depositHelper(startLine, desc):
        while(True):
            # look for money amount
            line = re.search('\d*,?\d+\.\d{2}', lines[startLine])
            if line:
                filtered = re.search('(?P<txt>.*?(?P<amt>\d*,?\d+\.\d{2}))', lines[startLine])
                desc += '\n' + filtered.group('txt')
                return startLine, desc, filtered.group('amt')
            else:
                desc += '\n' + lines[startLine]
                startLine += 1

    # we have to iterate page because date is in different lines in different statements
    for line in firstPage:
        searchDate = re.search('[a-zA-Z]+\s\d+,\s\d{4}', line)
        if searchDate:
            statementDate = searchDate.group(0)
            break
    year = statementDate[-2:]

    section = ['Total Deposits and Additions', 'Total Checks Paid',
               'Total ATM & Debit Card', 'Total Electronic Withdrawals', 'Total Fees']
    sectionIndex = 0
    completedFlag = False
    for num in range(len(reader.pages)):
        if completedFlag:
            break
        page = reader.pages[num]
        text = page.extract_text()
        lines = re.split('\n', text)
        date = ''
        desc = ''
        amt = ''
        checkNum = ''
        category = ''
        i = 0
        while lines[i] != lines[-1]:
            a = lines[i]
            # if num == 6 and i == 8:
            #     print()

            if section[sectionIndex] in ['Total Deposits and Additions']:
                # this is to check to see if transaction is a checks paid transaction
                checkQuery = re.match('(?P<check>\d{4})(\s*\*)*(\s\^)\s*(?P<date>\d{2}\/\d{2})\s*(\d{2}\/\d{2})*\s*'
                                      '(?P<amt>\d*,{1}\d{3}.\d{2})', lines[i])
                # this is to catch merged transaction within this summary line
                depositSummary = re.search('(?P<flag>Total Deposits and Additions \$\d*,{1}\d{3}.\d{2})'
                                           '(?P<trans>.*)', lines[i])
                if checkQuery:
                    sectionIndex += 1
                    desc = 'check'
                    checkNum = checkQuery.group('check')
                    date = checkQuery.group('date') + '/' + year
                    amt = checkQuery.group('amt')

                elif depositSummary:
                    # check for transaction merged with summary at the end
                    transaction = re.search('(?P<date>\d{2}/\d{2}) (?P<desc>.*)', depositSummary.group('trans'))
                    if transaction:
                        # if transaction found parse info and then check if one liner w/ amt at end
                        date = transaction.group('date') + '/' + year
                        desc = transaction.group('desc')
                        amtTest = re.search('(\d{2},{1}\d{3}\.\d{2})', desc)
                        if amtTest:
                            amt = amtTest.group(0)
                            desc = desc[:-(len(amt))]
                        else:
                            i, desc, amt = depositHelper(i + 1, desc)
                else:
                    validTransaction = re.match('^\s*\d+/\d+\s\w+', lines[i])
                    if not validTransaction:
                        i += 1
                        continue
                    date = validTransaction.group(0)[:5] + '/' + year
                    # searches for everything after date and a space
                    desc = re.search('(?P<txt>(?<=\d{2}/\d{2} ).*)', lines[i]).group('txt')
                    # check if line is a one liner where amount ia at the end of same line
                    amtTest = re.search('\d*,?\d+\.\d{2}', lines[i])
                    if amtTest:
                        amt = amtTest.group(0)
                        desc = desc[:-(len(amt))]   # and remove amt from description
                    else:
                        i, desc, amt = depositHelper(i+1, desc)
            elif section[sectionIndex] == 'Total Checks Paid':
                debitSummary = re.search('(?P<flag>Total ATM & Debit Card Withdrawals \$\d*,?\d{3}.\d\d)'
                                         '(?P<transaction>.*)', lines[i])
                checksPaidSummary = re.search('(?P<flag>Total Checks Paid \$\d*,?\d{3}.\d\d)'
                                              '(?P<transaction>.*)', lines[i])
                debitTransaction = re.search('Card Purchase', lines[i])
                if debitSummary:
                    nextTransaction = debitSummary.group('transaction')
                    # nextTransaction can be either a check or atm & debit card withdrawals transaction
                    checkQuery = re.match('(?P<check>\d{4})(\s*\*)*(\s\^)\s*(?P<date>\d{2}\/\d{2})\s*(\d{2}\/\d{2})*\s*'
                                          '(?P<amt>\d*,{1}\d{3}\.\d{2})', nextTransaction)
                    if checkQuery:
                        date = checkQuery.group('date') + '/' + year
                        desc = 'check'
                        amt = checkQuery.group('amt')
                        checkNum = checkQuery.group('check')
                    else:
                        nextTransaction = re.search('(?P<extra>.*(?<!\.|\s)\s{2,})(?P<date>\d{2}\/\d{2})\s'
                                                    '(?P<desc>.*(?=\$))\$(?P<amt>.*)', nextTransaction)
                        if not nextTransaction:
                            print(f'page {num} line{i}: transaction not found')
                            i += 1
                            continue
                        date = nextTransaction.group('date') + '/' + year
                        desc = nextTransaction.group('desc')
                        amt = nextTransaction.group('amt')
                elif debitTransaction:
                    sectionIndex += 1
                    validTransaction = re.match('(?P<extra>.*(?<!\.|\s)\s{2,})'
                                                '(?P<date>\d{2}\/\d{2})\s'
                                                '(?P<desc>.*)', lines[i])
                    if not validTransaction:
                        i += 1
                        continue
                    date = validTransaction.group('date') + '/' + year
                    # check if line is a one liner where amount ia at the end of same line
                    amtTest = re.search('\d*,?\d+\.\d{2}', validTransaction.group('desc'))
                    if amtTest:
                        amt = amtTest.group(0)
                        desc = validTransaction.group('desc')[:-(len(amt))]  # and remove amt from description
                    else:
                        i, desc, amt = depositHelper(i + 1, desc)
                else:
                    if checksPaidSummary:
                        checkQuery = re.match(
                            '(?P<check>\d{4})(\s*\*)*(\s\^)\s*(?P<date>\d{2}/\d{2})\s*(\d{2}/\d{2})*\s*'
                            '(?P<amt>\d*,?\d{3}.\d{2})', checksPaidSummary.group('transaction'))
                    else:
                        checkQuery = re.match('(?P<check>\d{4})(\s*\*)*(\s\^)\s*(?P<date>\d{2}/\d{2})\s*(\d{2}/\d{2})*\s*'
                                              '(?P<amt>\d*,?\d{3}.\d{2})', lines[i])
                    if not checkQuery:
                        i += 1
                        continue
                    date = checkQuery.group('date') + '/' + year
                    desc = 'check'
                    amt = checkQuery.group('amt')
                    checkNum = checkQuery.group('check')
            elif section[sectionIndex] == 'Total ATM & Debit Card':
                totalFlag = re.search('Total Card Purchase', lines[i])
                if totalFlag:
                    i += 1
                    continue
                debitFlag = re.search('Card Purchase', lines[i])
                # the extra group is any word or space and before a large space
                validTransaction = re.match('^\s*\d+/\d+\s\w+', lines[i])
                if debitFlag:
                    debitTransaction = re.match('(?P<extra>.*)(?P<date>\d{2}\/\d{2})\s(?P<desc>.*)', lines[i])
                    if debitTransaction:
                        date = debitTransaction.group('date') + '/' + year
                        # check if line is a one liner where amount ia at the end of same line
                        amtTest = re.search('\d*,?\d+\.\d{2}', debitTransaction.group('desc'))
                        if amtTest:
                            amt = amtTest.group(0)
                            desc = debitTransaction.group('desc')[:-(len(amt))]  # and remove amt from description
                        else:
                            i, desc, amt = depositHelper(i + 1, desc)
                elif validTransaction:
                    sectionIndex += 1
                    date = validTransaction.group(0)[:5] + '/' + year
                    # searches for everything after date and a space
                    desc = re.search('(?P<txt>(?<=\d{2}/\d{2} ).*)', lines[i]).group('txt')
                    # check if line is a one liner where amount ia at the end of same line
                    amtTest = re.search('\d*,?\d+\.\d{2}', lines[i])
                    if amtTest:
                        amt = amtTest.group(0)
                        desc = desc[:-(len(amt))]  # and remove amt from description
                    else:
                        i, desc, amt = depositHelper(i + 1, desc)
                else:
                    i += 1
                    continue

            elif section[sectionIndex] == 'Total Electronic Withdrawals':
                totalSummary = re.search('(?P<flag>Total Transactions \d*)(?P<trans>(?=\d\d\/\d\d).*)', lines[i])
                feeSummary = re.search('(?P<date>\d{2}\/\d{2})\s*(?P<flag>Service Charges For The Month of \w*)', lines[i])
                if totalSummary:
                    validTransaction = re.match('^\s*\d+/\d+\s\w+', totalSummary.group('trans'))
                    if not validTransaction:
                        i += 1
                        continue
                    date = validTransaction.group(0)[:5] + '/' + year
                    trans = totalSummary.group('trans')
                    # searches for everything after date and a space
                    desc = re.search('(?P<txt>(?<=\d{2}/\d{2} ).*)', trans).group('txt')
                    # check if line is a one liner where amount ia at the end of same line
                    amtTest = re.search('\d*,?\d+\.\d{2}', desc)
                    if amtTest:
                        amt = amtTest.group(0)
                        desc = desc[:-(len(amt))]  # and remove amt from description
                    else:
                        i, desc, amt = depositHelper(i + 1, desc)
                elif feeSummary:
                    date = feeSummary.group('date') + '/' + year
                    desc = 'service charge'
                    amt = re.search('\d*,?\d+\.\d{2}', lines[i]).group(0)
                    amt = -abs(float(amt))
                    newRow = {'Category': 'Bank Fees',
                              'type': 'debit' if amt < 0 else 'credit',
                              'date': date,
                              'description': desc,
                              'amount': amt,
                              'check#': checkNum
                              }
                    newDf.append(pd.DataFrame(newRow, index=[0]))
                    completedFlag = True
                    break
                else:
                    validTransaction = re.match('^\s*\d+/\d+\s\w+', lines[i])
                    if not validTransaction:
                        i += 1
                        continue
                    date = validTransaction.group(0)[:5] + '/' + year
                    # searches for everything after date and a space
                    desc = re.search('(?P<txt>(?<=\d{2}/\d{2} ).*)', lines[i]).group('txt')
                    # check if line is a one liner where amount ia at the end of same line
                    amtTest = re.search('\d*,?\d+\.\d{2}', lines[i])
                    if amtTest:
                        amt = amtTest.group(0)
                        desc = desc[:-(len(amt))]  # and remove amt from description
                    else:
                        i, desc, amt = depositHelper(i + 1, desc)
            amt = float(amt.replace(",", ""))
            if section[sectionIndex] != 'Total Deposits and Additions':
                amt = -abs(amt)

            identified = False
            for subCategory in description.keys():
                values = description[subCategory]
                if isinstance(values, str):  # if only one value
                    if values in desc.lower():
                        category = subCategory
                        identified = True
                        break
                else:  # has a list of values
                    for value in values:
                        if value in desc.lower():
                            category = subCategory
                            identified = True
                            break
            if not identified:
                category = 'Miscellaneous'

            newRow = {'Category': category,
                      'type': 'debit' if amt < 0 else 'credit',
                      'date': date,
                      'description': desc,
                      'amount': amt,
                      'check#': checkNum
                      }
            if desc == 'check':
                newRow['type'] = 'check'
            if newRow['Category'] == "Deposits":
                newRow['type'] = 'dslip'
            i += 1
            newDf.append(pd.DataFrame(newRow, index=[0]))
            checkNum = ''
            if lines[i] == lines[-1]:
                break
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
    dfBank = pd.concat(newDf)
    return dfBank


def exportToExcel(outputExcel, dfBank, initial):
    with pd.ExcelWriter(outputExcel, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        dfBank.to_excel(writer, sheet_name=initial + '.bank', header=None, index=False,
                        startrow=writer.sheets[initial + '.bank'].max_row)


def copy_and_replace(source_path, destination_path):
    if os.path.exists(destination_path):
        os.remove(destination_path)
    shutil.copy2(source_path, destination_path)





def main():
    window = makeWindow(sg.theme())
    bank = ''
    book = '2024taxCat.xlsx'

    transactions = ''
    dfBank = pd.DataFrame(columns=['Category', 'type', 'date', 'description', 'amount', 'check#'])
    while True:
        window['bank'].update(bank)
        window['book'].update(book)
        event, values = window.read()
        if event not in (sg.TIMEOUT_EVENT, sg.WIN_CLOSED, 'exit'):
            print('============ Event = ', event, ' ==============')
            print('-------- Values Dictionary (key=value) --------')
            for key in values:
                print(key, ' = ', values[key])
        if event == '-bank-':
            bank = sg.popup_get_file('Bank transactions/statements', 'Choose bank info', initial_folder=os.getcwd())
            window['bank'].update(bank)
        elif event == '-book-':
            book = sg.popup_get_file('Excel book', 'Choose excel book', initial_folder=os.getcwd())
            window['book'].update(book)
        elif event == '-process-':
            if book == '':
                window['-status-'].update(f'Error: no book specified')
                continue
            elif bank == '':
                window['-status-'].update(f'Error: no bank specified')
                continue

            # need to catch file exists error later on
            newBook = book[:-5] + '.1.xlsx'
            copy_and_replace(book, newBook)
            book = newBook
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
                    result = chaseParseStatements(bank, dfBank)
            exportToExcel(book, result, bankInitial)
            window['-status-'].update(f'{bank} in {book} Process complete')
        elif event == '-auto-':
            result = ''
            if book == '':
                window['-status-'].update(f'Error: no book specified')
                continue

            # need to catch file exists error later on
            newBook = book[:-5] + '.1.xlsx'
            copy_and_replace(book, newBook)
            book = newBook
            directories = ['banks\\everProsper', 'banks\\thrivegen']
            for directory in directories:
                # iterate over files in
                # that directory
                for filename in os.listdir(directory):
                    f = os.path.join(directory, filename)
                    # checking if it is a file
                    if os.path.isfile(f):
                        print(f)
                        if 'everProsper' in directory:
                            if f[-3:] == 'pdf':
                                result = chaseParseStatements(f, dfBank)
                            elif f[-3:] == 'csv':
                                result = chaseParseTransactions(f, dfBank)
                            bankInitial = 'e'
                        else:
                            if f[-3:] == 'pdf':
                                result = exchangeParseStatements(f, dfBank)
                            elif f[-3:] == 'csv':
                                result = exchangeParseTransactions(f, dfBank)
                            bankInitial = 't'
                        exportToExcel(book, result, bankInitial)
                        newBook = book[:-5] + '.1.xlsx'
                        copy_and_replace(book, newBook)
            os.remove(newBook)
            window['-status-'].update(f'{bank} in {book} Auto Process complete')
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
