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
    - BB: Recommendation
    - BC: Sourced Rental Rate
    - BD: Sourced Rental Rates
    - BE: Rental Rate Source
    - BF: Sourced Value
    - BG: Sourced Values
    - BH: Value Source
    - BI: Date of Program Execution

### Meaning of Recommendations
- **"Data supports pricing":** the data found online supports the given pricing
- **"Insufficient data available online":** there is not enough available data online to draw reasonable conclusions
- **"Further research required: possible slight increase/decrease in pricing":** The given price fits that data but is not near the median. Further research is required as the online prices found may be linked to simmilar but different type of equipment
- **"Consider increase/decrease in pricing":** The given price does not fit the data and is well below the median value. This suggests that the price should change to reflect closer to market values

## Creating the Executable for Windows:
```
pyinstaller ./main.py --onefile
```
