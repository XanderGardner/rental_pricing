# rental_pricing
Tool for automatically finding current rental prices per hour 

## Input:
- Excel File named "equipment rates.xlsx" with columns:
    - B: year
    - C: description
    - E: manufacturer
    - F: model
    - W: given value
    - H: given operating rate
    - J: given standby rate
    - BB: sourced value (program finds this if it is not already present)
    - BC: sourced rental rate (program finds this if it is not already present)

## Output:
- Output is in the same excel file named "equipment rates.xlsx" with added columns:
    - BB: sourced value
    - BC: sourced rental rate (per hour)

## Creating the Executable for Windows:
https://www.zacoding.com/en/post/python-selenium-to-exe/
```
pyinstaller ./main.py --onefile --add-binary "./chromedriver_win32/chromedriver.exe;./chromedriver_win32"
```