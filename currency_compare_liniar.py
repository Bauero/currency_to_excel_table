from openpyxl import Workbook
from helper_functions import *
from forex_python.converter import CurrencyRates


# preparation to work on the spreadsheat
workbook = Workbook()
sheet = workbook.active

def get_convresion_rate(crc_1, crc_2):
    if crc_1 == crc_2:
        return "-"
    else:
        try:
            return round(CurrencyRates().get_rate(crc_1, crc_2), 4)
        except:
            return ""

# list of compared currencies
currency_list = ["USD", 
                 "PLN", 
                 "GBP", 
                 "EUR", 
                 "CHF", 
                 "DKK", 
                 "SEK", 
                 "NOK", 
                 "CZK",
                 "HUF", 
                 "JPY"]


# current cell in which the data are put
column = 1
row = 1

# writing the header column and header row of names

for i in range(len(currency_list)):
    localization = toExclNot(column + i + 1, row)      # header row
    sheet[localization] = currency_list[i]
    localization = toExclNot(column, row + i + 1)      # header column
    sheet[localization] = currency_list[i]

# translation to the place where the data will be input
column = 2
row = 2

# actually writing the data into table
for next_column in range(len(currency_list)):
    for next_row in range(len(currency_list)):
        localization = toExclNot(column + next_column, row + next_row)
        value = get_convresion_rate(currency_list[next_column], 
                                    currency_list[next_row])
        #zaokraglenie
        sheet[localization] = value

workbook.save(filename="currency_liniar_solution.xlsx")
