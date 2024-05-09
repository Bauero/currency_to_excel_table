import asyncio
from forex_python.converter import CurrencyRates
from openpyxl import Workbook
from codetiming import Timer
from copy import deepcopy
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

#
# Solution using while loop and iterating over queue
#

async def convert_and_save(currency_queue, sheet) -> Workbook:
    
    currency_queue_copy = deepcopy(currency_queue)
    var = None
    row = 1

    while not currency_queue.empty():

        # Get the new value from queue and 
        crc_1 = await currency_queue.get()
        tmp_currency_queue = deepcopy(currency_queue_copy)
        row += 1
        column = 2

        while not tmp_currency_queue.empty():

            crc_2 = await tmp_currency_queue.get()
    
            if crc_1 != crc_2:
                try:
                    var = round(float(CurrencyRates().get_rate(crc_1, crc_2)) , 4)
                except:
                    var = "0"
            else:
                var = "-"

            localization = toExclNot(column, row)
            sheet[localization] = var
            column += 1

async def create_currency_table() -> Workbook:

    workbook = Workbook()
    sheet = workbook.active

    currency_queue = asyncio.Queue()
    for c in currency_list:
        await currency_queue.put(c)

    column = 1
    row = 1

    sheet[toExclNot(1,1)] = "One ↓ ="

    for i in range(len(currency_list)):
        localization = toExclNot(column + i + 1, row)      # header row
        sheet[localization] = currency_list[i]
        localization = toExclNot(column, row + i + 1)      # header column
        sheet[localization] = currency_list[i]

    with Timer(text="Total elapsed time: {:.1f}"):
        await asyncio.gather(
            asyncio.create_task(
                convert_and_save(currency_queue, sheet)
                )
            )

    return workbook

#
# Solution using for loop and iterating over list
#

async def convert_and_save(currency_list, sheet) -> Workbook:

    for curr_1 in range(len(currency_list)):
        for curr_2 in range(len(currency_list)):
            crc1 = currency_list[curr_1]
            crc2 = currency_list[curr_2]
            if curr_1 != curr_2:
                try:
                    var = await round(
                        float(CurrencyRates().get_rate(crc1, crc2)), 
                        4)
                except:
                    var = "0"
            else:
                var = "-"
            localization = toExclNot(curr_1+2, curr_2+2)
            sheet[localization] = var
    
    return sheet


async def create_currency_table() -> Workbook:

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

    with Timer(text="Total elapsed time: {:.1f}"):
        await asyncio.gather(
            asyncio.create_task(
                convert_and_save(currency_list, sheet)
                )
            )
        
    return workbook


#################################  EXECUTION  #################################


if __name__ == "__main__":
    workbook = asyncio.run(create_currency_table())
    workbook.save(filename="currencies_asyncio_solution.xlsx")
