from forex_python.converter import CurrencyRates
from openpyxl import Workbook
from codetiming import Timer
from threading import Thread
from helper_functions import *

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

################################  ACTUAL CODE  ################################


def ask_for_currency(curr_1, curr_2, sheet):
    crc1 = currency_list[curr_1]
    crc2 = currency_list[curr_2]
    # print(f"Asking for conversionf between {crc1} and {crc2}")
    if curr_1 != curr_2:
        try:
            var = round(
                float(CurrencyRates().get_rate(crc1, crc2)), 
                4)
        except:
            var = "0"
    else:
        var = "-"

    localization = toExclNot(curr_2+2, curr_1+2)
    sheet[localization] = var


def fill_using_threading():

    workbook = Workbook()
    sheet = workbook.active

    column = 1
    row = 1

    sheet[toExclNot(1,1)] = "One ↓ ="

    for i in range(len(currency_list)):
        localization = toExclNot(column + i + 1, row)      # header row
        sheet[localization] = currency_list[i]
        localization = toExclNot(column, row + i + 1)      # header column
        sheet[localization] = currency_list[i]

    threads = []

    for curr_1 in range(len(currency_list)):
        for curr_2 in range(len(currency_list)):
            t = Thread(target=ask_for_currency, args=[curr_1, curr_2, sheet])
            t.start()
            threads.append(t)

    working = [True]

    while any(working):
        working = []
        for t in threads:
            working.append(t.is_alive())

    return workbook


#################################  EXECUTION  #################################


if __name__ == "__main__":
    with Timer(text="\nTotal elapsed time: {:.2f}s"):
        workbook = fill_using_threading()
    workbook.save(filename="currencies_threded_solution.xlsx")
