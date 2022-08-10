# rental_pricing
Tool for automatically finding current rental prices per hour 

## Table of Contents
- [Input](#input)
- [Output](#output)
- [Meaning of Recommendations](#meaning-of-recommendations)
- [Reasons for Crash](#reasons-for-crash)

## Input:
- Excel File named "equipment rates.xlsx"
- Sheet in the excel file MUST be named **General**
- Cell BA2 must have a location
- Excel File must have columns:
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
    - **BB: Recommendation:** A recommendation on how to change the operating rate based on sourced data
    - **BC: Sourced Rental Rate:** The found rental rate (median of Sourced Rental Rates)
    - **BD: Sourced Rental Rates:** A list of all rental rates found online
    - **BE: Rental Rate Source:** URL where many of the rental rates were found
    - **BF: Sourced Value:**  The found market value (median of Sourced Values)
    - **BG: Sourced Values:** A list of all market values found online
    - **BH: Value Source:** URL where many of the market values were found
    - **BI: Date of Program Execution:** Date and time the program completed webscrapping for this line item

## Meaning of Recommendations
- **"Data supports pricing":** The given operation rate lies inside the Sourced Rental Rates (ie it is not an extreme point) and the given operation rate is within 50% of the Sourced Rental Rate.
- **"Insufficient data available online":** There are less than 3 Sourced Rental Rates found online
- **"Further research required: possible slight increase/decrease in pricing":** The given operation rate lies inside the Sourced Rental Rates but is not within 50% of the Sourced Rental Rate. The online prices found may be linked to simmilar but different types of equipment
- **"Consider increase/decrease in pricing":** The given operation rate lies outside the Sourced Rental Rates (ie it is an extreme point) and is well below/above (by >50%) the Sourced Rental Rate. 

## Reasons for Crash
- There must be an excel File named "equipment rates.xlsx" in the same folder as rental_pricing.exe
- Sheet in the excel file MUST be named **General**
- "equipement rates.xlsx" must not be opened while the software is running
- Columns BB-BI must be either empty or filled in by a previous run of the program  

## Creating the Executable for Windows:
```
pyinstaller ./main.py --onefile
```
