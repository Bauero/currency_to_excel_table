# Purpose of files in this repo
This is a small example of a program I've created to allow to create an Excel file, which would contain small table to compare conversion rates between currencies. All files "currency_compare_#####.py" perform the same action, but using different approach - liniar, download every combination of conversion one by one, while asyncio and threaded solutions tries to do this more efficiently using threading (so far, there is not so much difference between liniar and threading approach in therms of time - might have to rewrite program, as it might not work in a way I expec it to)

## Currencies being converted
- USD
- PLN
- GBP
- EUR
- CHF
- DKK
- SEK
- NOK
- CZK
- HUF
- JPY

# Running the program
## Requirements
External liblaries, with sources:
- forex (https://pypi.org/project/forex-python/0.3.1/)
- openpyxl (https://pypi.org/project/openpyxl/)
- codetiming (https://pypi.org/project/codetiming/)

In therms of python version, I've tested program on Python3.11

## How to run
### On Windows
```cmd
# You have to go to the specific directory using dir command or include full path with file name
py .\"name_of_the_file_to_run_with.py"
```

### On Linux/MacOS
```bash
# You have to go to the specific directory using cd command or include full path with file name
python3.11 ./"name_of_the_file_to_run_with.py"
```
