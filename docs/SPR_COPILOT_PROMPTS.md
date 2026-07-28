# Copilot prompts to build the SPR sheets
Run ONE prompt per country in Excel's Copilot pane, on AMKAD_Current_Input.xlsm.
After each one, verify the two table names via Table Design > Table Name.

---

## AR  (6 customers)

```
Create a new worksheet named SPR_AR in this workbook.

On SPR_AR, put these column headers in row 1, cells A1 to E1, exactly:
Month | Country | Customer | Issue | Action

Below that header row, fill in the Month, Country and Customer columns only,
repeating each customer on 3 consecutive rows, leaving Issue and Action empty.
Use these Month, Country and Customer values, each repeated 3 times:
2026-07 | AR | Arrow Electronics
2026-07 | AR | Halliburton
2026-07 | AR | Hpe
2026-07 | AR | Hpi
2026-07 | AR | Lenovo/Motorola
2026-07 | AR | SLB/Cameron

Then convert the whole block including the header row into an Excel Table and
name that table exactly: SPRIssues_AR
Set column D and column E to about 60 characters wide with text wrapping on.
Set columns A, B and C to about 12 characters wide.

Next, leaving two empty rows after that table, add a second header row with
these 8 headers, exactly:
Month | Country | Customer | Payments In House Pending | Expected Payments | Forecasted Overdue EOM | Forecasted GT60 EOM | Forecasted GT90 EOM

Below it add one row per customer, filling only Month, Country and Customer:
2026-07 | AR | Arrow Electronics
2026-07 | AR | Halliburton
2026-07 | AR | Hpe
2026-07 | AR | Hpi
2026-07 | AR | Lenovo/Motorola
2026-07 | AR | SLB/Cameron

Convert that block including its header row into an Excel Table named exactly:
SPRForecast_AR
Format the five forecast amount columns as number with thousands separator and
no decimal places.

Do not add any other sample data.
```

## BM  (1 customers)

```
Create a new worksheet named SPR_BM in this workbook.

On SPR_BM, put these column headers in row 1, cells A1 to E1, exactly:
Month | Country | Customer | Issue | Action

Below that header row, fill in the Month, Country and Customer columns only,
repeating each customer on 3 consecutive rows, leaving Issue and Action empty.
Use these Month, Country and Customer values, each repeated 3 times:
2026-07 | BM | SLB/Cameron

Then convert the whole block including the header row into an Excel Table and
name that table exactly: SPRIssues_BM
Set column D and column E to about 60 characters wide with text wrapping on.
Set columns A, B and C to about 12 characters wide.

Next, leaving two empty rows after that table, add a second header row with
these 8 headers, exactly:
Month | Country | Customer | Payments In House Pending | Expected Payments | Forecasted Overdue EOM | Forecasted GT60 EOM | Forecasted GT90 EOM

Below it add one row per customer, filling only Month, Country and Customer:
2026-07 | BM | SLB/Cameron

Convert that block including its header row into an Excel Table named exactly:
SPRForecast_BM
Format the five forecast amount columns as number with thousands separator and
no decimal places.

Do not add any other sample data.
```

## BO  (2 customers)

```
Create a new worksheet named SPR_BO in this workbook.

On SPR_BO, put these column headers in row 1, cells A1 to E1, exactly:
Month | Country | Customer | Issue | Action

Below that header row, fill in the Month, Country and Customer columns only,
repeating each customer on 3 consecutive rows, leaving Issue and Action empty.
Use these Month, Country and Customer values, each repeated 3 times:
2026-07 | BO | Halliburton
2026-07 | BO | SLB/Cameron

Then convert the whole block including the header row into an Excel Table and
name that table exactly: SPRIssues_BO
Set column D and column E to about 60 characters wide with text wrapping on.
Set columns A, B and C to about 12 characters wide.

Next, leaving two empty rows after that table, add a second header row with
these 8 headers, exactly:
Month | Country | Customer | Payments In House Pending | Expected Payments | Forecasted Overdue EOM | Forecasted GT60 EOM | Forecasted GT90 EOM

Below it add one row per customer, filling only Month, Country and Customer:
2026-07 | BO | Halliburton
2026-07 | BO | SLB/Cameron

Convert that block including its header row into an Excel Table named exactly:
SPRForecast_BO
Format the five forecast amount columns as number with thousands separator and
no decimal places.

Do not add any other sample data.
```

## BR  (5 customers)

```
Create a new worksheet named SPR_BR in this workbook.

On SPR_BR, put these column headers in row 1, cells A1 to E1, exactly:
Month | Country | Customer | Issue | Action

Below that header row, fill in the Month, Country and Customer columns only,
repeating each customer on 3 consecutive rows, leaving Issue and Action empty.
Use these Month, Country and Customer values, each repeated 3 times:
2026-07 | BR | Arrow Electronics
2026-07 | BR | Halliburton
2026-07 | BR | Hpi
2026-07 | BR | Lenovo/Motorola
2026-07 | BR | SLB/Cameron

Then convert the whole block including the header row into an Excel Table and
name that table exactly: SPRIssues_BR
Set column D and column E to about 60 characters wide with text wrapping on.
Set columns A, B and C to about 12 characters wide.

Next, leaving two empty rows after that table, add a second header row with
these 8 headers, exactly:
Month | Country | Customer | Payments In House Pending | Expected Payments | Forecasted Overdue EOM | Forecasted GT60 EOM | Forecasted GT90 EOM

Below it add one row per customer, filling only Month, Country and Customer:
2026-07 | BR | Arrow Electronics
2026-07 | BR | Halliburton
2026-07 | BR | Hpi
2026-07 | BR | Lenovo/Motorola
2026-07 | BR | SLB/Cameron

Convert that block including its header row into an Excel Table named exactly:
SPRForecast_BR
Format the five forecast amount columns as number with thousands separator and
no decimal places.

Do not add any other sample data.
```

## CA  (6 customers)

```
Create a new worksheet named SPR_CA in this workbook.

On SPR_CA, put these column headers in row 1, cells A1 to E1, exactly:
Month | Country | Customer | Issue | Action

Below that header row, fill in the Month, Country and Customer columns only,
repeating each customer on 3 consecutive rows, leaving Issue and Action empty.
Use these Month, Country and Customer values, each repeated 3 times:
2026-07 | CA | Amazon
2026-07 | CA | Halliburton
2026-07 | CA | Honeywell
2026-07 | CA | SLB/Cameron
2026-07 | CA | TE Connectivity
2026-07 | CA | Texas Instruments

Then convert the whole block including the header row into an Excel Table and
name that table exactly: SPRIssues_CA
Set column D and column E to about 60 characters wide with text wrapping on.
Set columns A, B and C to about 12 characters wide.

Next, leaving two empty rows after that table, add a second header row with
these 8 headers, exactly:
Month | Country | Customer | Payments In House Pending | Expected Payments | Forecasted Overdue EOM | Forecasted GT60 EOM | Forecasted GT90 EOM

Below it add one row per customer, filling only Month, Country and Customer:
2026-07 | CA | Amazon
2026-07 | CA | Halliburton
2026-07 | CA | Honeywell
2026-07 | CA | SLB/Cameron
2026-07 | CA | TE Connectivity
2026-07 | CA | Texas Instruments

Convert that block including its header row into an Excel Table named exactly:
SPRForecast_CA
Format the five forecast amount columns as number with thousands separator and
no decimal places.

Do not add any other sample data.
```

## CL  (4 customers)

```
Create a new worksheet named SPR_CL in this workbook.

On SPR_CL, put these column headers in row 1, cells A1 to E1, exactly:
Month | Country | Customer | Issue | Action

Below that header row, fill in the Month, Country and Customer columns only,
repeating each customer on 3 consecutive rows, leaving Issue and Action empty.
Use these Month, Country and Customer values, each repeated 3 times:
2026-07 | CL | Hpe
2026-07 | CL | Hpi
2026-07 | CL | Lenovo/Motorola
2026-07 | CL | SLB/Cameron

Then convert the whole block including the header row into an Excel Table and
name that table exactly: SPRIssues_CL
Set column D and column E to about 60 characters wide with text wrapping on.
Set columns A, B and C to about 12 characters wide.

Next, leaving two empty rows after that table, add a second header row with
these 8 headers, exactly:
Month | Country | Customer | Payments In House Pending | Expected Payments | Forecasted Overdue EOM | Forecasted GT60 EOM | Forecasted GT90 EOM

Below it add one row per customer, filling only Month, Country and Customer:
2026-07 | CL | Hpe
2026-07 | CL | Hpi
2026-07 | CL | Lenovo/Motorola
2026-07 | CL | SLB/Cameron

Convert that block including its header row into an Excel Table named exactly:
SPRForecast_CL
Format the five forecast amount columns as number with thousands separator and
no decimal places.

Do not add any other sample data.
```

## CO  (5 customers)

```
Create a new worksheet named SPR_CO in this workbook.

On SPR_CO, put these column headers in row 1, cells A1 to E1, exactly:
Month | Country | Customer | Issue | Action

Below that header row, fill in the Month, Country and Customer columns only,
repeating each customer on 3 consecutive rows, leaving Issue and Action empty.
Use these Month, Country and Customer values, each repeated 3 times:
2026-07 | CO | Halliburton
2026-07 | CO | Hpe
2026-07 | CO | Hpi
2026-07 | CO | Lenovo/Motorola
2026-07 | CO | SLB/Cameron

Then convert the whole block including the header row into an Excel Table and
name that table exactly: SPRIssues_CO
Set column D and column E to about 60 characters wide with text wrapping on.
Set columns A, B and C to about 12 characters wide.

Next, leaving two empty rows after that table, add a second header row with
these 8 headers, exactly:
Month | Country | Customer | Payments In House Pending | Expected Payments | Forecasted Overdue EOM | Forecasted GT60 EOM | Forecasted GT90 EOM

Below it add one row per customer, filling only Month, Country and Customer:
2026-07 | CO | Halliburton
2026-07 | CO | Hpe
2026-07 | CO | Hpi
2026-07 | CO | Lenovo/Motorola
2026-07 | CO | SLB/Cameron

Convert that block including its header row into an Excel Table named exactly:
SPRForecast_CO
Format the five forecast amount columns as number with thousands separator and
no decimal places.

Do not add any other sample data.
```

## CR  (3 customers)

```
Create a new worksheet named SPR_CR in this workbook.

On SPR_CR, put these column headers in row 1, cells A1 to E1, exactly:
Month | Country | Customer | Issue | Action

Below that header row, fill in the Month, Country and Customer columns only,
repeating each customer on 3 consecutive rows, leaving Issue and Action empty.
Use these Month, Country and Customer values, each repeated 3 times:
2026-07 | CR | Amazon
2026-07 | CR | Hpe
2026-07 | CR | Hpi

Then convert the whole block including the header row into an Excel Table and
name that table exactly: SPRIssues_CR
Set column D and column E to about 60 characters wide with text wrapping on.
Set columns A, B and C to about 12 characters wide.

Next, leaving two empty rows after that table, add a second header row with
these 8 headers, exactly:
Month | Country | Customer | Payments In House Pending | Expected Payments | Forecasted Overdue EOM | Forecasted GT60 EOM | Forecasted GT90 EOM

Below it add one row per customer, filling only Month, Country and Customer:
2026-07 | CR | Amazon
2026-07 | CR | Hpe
2026-07 | CR | Hpi

Convert that block including its header row into an Excel Table named exactly:
SPRForecast_CR
Format the five forecast amount columns as number with thousands separator and
no decimal places.

Do not add any other sample data.
```

## EC  (2 customers)

```
Create a new worksheet named SPR_EC in this workbook.

On SPR_EC, put these column headers in row 1, cells A1 to E1, exactly:
Month | Country | Customer | Issue | Action

Below that header row, fill in the Month, Country and Customer columns only,
repeating each customer on 3 consecutive rows, leaving Issue and Action empty.
Use these Month, Country and Customer values, each repeated 3 times:
2026-07 | EC | Halliburton
2026-07 | EC | SLB/Cameron

Then convert the whole block including the header row into an Excel Table and
name that table exactly: SPRIssues_EC
Set column D and column E to about 60 characters wide with text wrapping on.
Set columns A, B and C to about 12 characters wide.

Next, leaving two empty rows after that table, add a second header row with
these 8 headers, exactly:
Month | Country | Customer | Payments In House Pending | Expected Payments | Forecasted Overdue EOM | Forecasted GT60 EOM | Forecasted GT90 EOM

Below it add one row per customer, filling only Month, Country and Customer:
2026-07 | EC | Halliburton
2026-07 | EC | SLB/Cameron

Convert that block including its header row into an Excel Table named exactly:
SPRForecast_EC
Format the five forecast amount columns as number with thousands separator and
no decimal places.

Do not add any other sample data.
```

## GT  (1 customers)

```
Create a new worksheet named SPR_GT in this workbook.

On SPR_GT, put these column headers in row 1, cells A1 to E1, exactly:
Month | Country | Customer | Issue | Action

Below that header row, fill in the Month, Country and Customer columns only,
repeating each customer on 3 consecutive rows, leaving Issue and Action empty.
Use these Month, Country and Customer values, each repeated 3 times:
2026-07 | GT | SLB/Cameron

Then convert the whole block including the header row into an Excel Table and
name that table exactly: SPRIssues_GT
Set column D and column E to about 60 characters wide with text wrapping on.
Set columns A, B and C to about 12 characters wide.

Next, leaving two empty rows after that table, add a second header row with
these 8 headers, exactly:
Month | Country | Customer | Payments In House Pending | Expected Payments | Forecasted Overdue EOM | Forecasted GT60 EOM | Forecasted GT90 EOM

Below it add one row per customer, filling only Month, Country and Customer:
2026-07 | GT | SLB/Cameron

Convert that block including its header row into an Excel Table named exactly:
SPRForecast_GT
Format the five forecast amount columns as number with thousands separator and
no decimal places.

Do not add any other sample data.
```

## MX  (9 customers)

```
Create a new worksheet named SPR_MX in this workbook.

On SPR_MX, put these column headers in row 1, cells A1 to E1, exactly:
Month | Country | Customer | Issue | Action

Below that header row, fill in the Month, Country and Customer columns only,
repeating each customer on 3 consecutive rows, leaving Issue and Action empty.
Use these Month, Country and Customer values, each repeated 3 times:
2026-07 | MX | Arrow Electronics
2026-07 | MX | Dell Technologies
2026-07 | MX | Halliburton
2026-07 | MX | Hpe
2026-07 | MX | Hpi
2026-07 | MX | Lenovo/Motorola
2026-07 | MX | Nokia
2026-07 | MX | TE Connectivity
2026-07 | MX | Texas Instruments

Then convert the whole block including the header row into an Excel Table and
name that table exactly: SPRIssues_MX
Set column D and column E to about 60 characters wide with text wrapping on.
Set columns A, B and C to about 12 characters wide.

Next, leaving two empty rows after that table, add a second header row with
these 8 headers, exactly:
Month | Country | Customer | Payments In House Pending | Expected Payments | Forecasted Overdue EOM | Forecasted GT60 EOM | Forecasted GT90 EOM

Below it add one row per customer, filling only Month, Country and Customer:
2026-07 | MX | Arrow Electronics
2026-07 | MX | Dell Technologies
2026-07 | MX | Halliburton
2026-07 | MX | Hpe
2026-07 | MX | Hpi
2026-07 | MX | Lenovo/Motorola
2026-07 | MX | Nokia
2026-07 | MX | TE Connectivity
2026-07 | MX | Texas Instruments

Convert that block including its header row into an Excel Table named exactly:
SPRForecast_MX
Format the five forecast amount columns as number with thousands separator and
no decimal places.

Do not add any other sample data.
```

## PA  (1 customers)

```
Create a new worksheet named SPR_PA in this workbook.

On SPR_PA, put these column headers in row 1, cells A1 to E1, exactly:
Month | Country | Customer | Issue | Action

Below that header row, fill in the Month, Country and Customer columns only,
repeating each customer on 3 consecutive rows, leaving Issue and Action empty.
Use these Month, Country and Customer values, each repeated 3 times:
2026-07 | PA | Halliburton

Then convert the whole block including the header row into an Excel Table and
name that table exactly: SPRIssues_PA
Set column D and column E to about 60 characters wide with text wrapping on.
Set columns A, B and C to about 12 characters wide.

Next, leaving two empty rows after that table, add a second header row with
these 8 headers, exactly:
Month | Country | Customer | Payments In House Pending | Expected Payments | Forecasted Overdue EOM | Forecasted GT60 EOM | Forecasted GT90 EOM

Below it add one row per customer, filling only Month, Country and Customer:
2026-07 | PA | Halliburton

Convert that block including its header row into an Excel Table named exactly:
SPRForecast_PA
Format the five forecast amount columns as number with thousands separator and
no decimal places.

Do not add any other sample data.
```

## PE  (4 customers)

```
Create a new worksheet named SPR_PE in this workbook.

On SPR_PE, put these column headers in row 1, cells A1 to E1, exactly:
Month | Country | Customer | Issue | Action

Below that header row, fill in the Month, Country and Customer columns only,
repeating each customer on 3 consecutive rows, leaving Issue and Action empty.
Use these Month, Country and Customer values, each repeated 3 times:
2026-07 | PE | Hpe
2026-07 | PE | Hpi
2026-07 | PE | Lenovo/Motorola
2026-07 | PE | SLB/Cameron

Then convert the whole block including the header row into an Excel Table and
name that table exactly: SPRIssues_PE
Set column D and column E to about 60 characters wide with text wrapping on.
Set columns A, B and C to about 12 characters wide.

Next, leaving two empty rows after that table, add a second header row with
these 8 headers, exactly:
Month | Country | Customer | Payments In House Pending | Expected Payments | Forecasted Overdue EOM | Forecasted GT60 EOM | Forecasted GT90 EOM

Below it add one row per customer, filling only Month, Country and Customer:
2026-07 | PE | Hpe
2026-07 | PE | Hpi
2026-07 | PE | Lenovo/Motorola
2026-07 | PE | SLB/Cameron

Convert that block including its header row into an Excel Table named exactly:
SPRForecast_PE
Format the five forecast amount columns as number with thousands separator and
no decimal places.

Do not add any other sample data.
```

## TT  (2 customers)

```
Create a new worksheet named SPR_TT in this workbook.

On SPR_TT, put these column headers in row 1, cells A1 to E1, exactly:
Month | Country | Customer | Issue | Action

Below that header row, fill in the Month, Country and Customer columns only,
repeating each customer on 3 consecutive rows, leaving Issue and Action empty.
Use these Month, Country and Customer values, each repeated 3 times:
2026-07 | TT | Halliburton
2026-07 | TT | SLB/Cameron

Then convert the whole block including the header row into an Excel Table and
name that table exactly: SPRIssues_TT
Set column D and column E to about 60 characters wide with text wrapping on.
Set columns A, B and C to about 12 characters wide.

Next, leaving two empty rows after that table, add a second header row with
these 8 headers, exactly:
Month | Country | Customer | Payments In House Pending | Expected Payments | Forecasted Overdue EOM | Forecasted GT60 EOM | Forecasted GT90 EOM

Below it add one row per customer, filling only Month, Country and Customer:
2026-07 | TT | Halliburton
2026-07 | TT | SLB/Cameron

Convert that block including its header row into an Excel Table named exactly:
SPRForecast_TT
Format the five forecast amount columns as number with thousands separator and
no decimal places.

Do not add any other sample data.
```

## US  (13 customers)

```
Create a new worksheet named SPR_US in this workbook.

On SPR_US, put these column headers in row 1, cells A1 to E1, exactly:
Month | Country | Customer | Issue | Action

Below that header row, fill in the Month, Country and Customer columns only,
repeating each customer on 3 consecutive rows, leaving Issue and Action empty.
Use these Month, Country and Customer values, each repeated 3 times:
2026-07 | US | Amazon
2026-07 | US | Arrow Electronics
2026-07 | US | Baker Hughes
2026-07 | US | Dell Technologies
2026-07 | US | Halliburton
2026-07 | US | Honeywell
2026-07 | US | Hpe
2026-07 | US | Hpi
2026-07 | US | Lenovo/Motorola
2026-07 | US | Nokia
2026-07 | US | SLB/Cameron
2026-07 | US | TE Connectivity
2026-07 | US | Texas Instruments

Then convert the whole block including the header row into an Excel Table and
name that table exactly: SPRIssues_US
Set column D and column E to about 60 characters wide with text wrapping on.
Set columns A, B and C to about 12 characters wide.

Next, leaving two empty rows after that table, add a second header row with
these 8 headers, exactly:
Month | Country | Customer | Payments In House Pending | Expected Payments | Forecasted Overdue EOM | Forecasted GT60 EOM | Forecasted GT90 EOM

Below it add one row per customer, filling only Month, Country and Customer:
2026-07 | US | Amazon
2026-07 | US | Arrow Electronics
2026-07 | US | Baker Hughes
2026-07 | US | Dell Technologies
2026-07 | US | Halliburton
2026-07 | US | Honeywell
2026-07 | US | Hpe
2026-07 | US | Hpi
2026-07 | US | Lenovo/Motorola
2026-07 | US | Nokia
2026-07 | US | SLB/Cameron
2026-07 | US | TE Connectivity
2026-07 | US | Texas Instruments

Convert that block including its header row into an Excel Table named exactly:
SPRForecast_US
Format the five forecast amount columns as number with thousands separator and
no decimal places.

Do not add any other sample data.
```

## UY  (2 customers)

```
Create a new worksheet named SPR_UY in this workbook.

On SPR_UY, put these column headers in row 1, cells A1 to E1, exactly:
Month | Country | Customer | Issue | Action

Below that header row, fill in the Month, Country and Customer columns only,
repeating each customer on 3 consecutive rows, leaving Issue and Action empty.
Use these Month, Country and Customer values, each repeated 3 times:
2026-07 | UY | Lenovo/Motorola
2026-07 | UY | SLB/Cameron

Then convert the whole block including the header row into an Excel Table and
name that table exactly: SPRIssues_UY
Set column D and column E to about 60 characters wide with text wrapping on.
Set columns A, B and C to about 12 characters wide.

Next, leaving two empty rows after that table, add a second header row with
these 8 headers, exactly:
Month | Country | Customer | Payments In House Pending | Expected Payments | Forecasted Overdue EOM | Forecasted GT60 EOM | Forecasted GT90 EOM

Below it add one row per customer, filling only Month, Country and Customer:
2026-07 | UY | Lenovo/Motorola
2026-07 | UY | SLB/Cameron

Convert that block including its header row into an Excel Table named exactly:
SPRForecast_UY
Format the five forecast amount columns as number with thousands separator and
no decimal places.

Do not add any other sample data.
```

## XC  (1 customers)

```
Create a new worksheet named SPR_XC in this workbook.

On SPR_XC, put these column headers in row 1, cells A1 to E1, exactly:
Month | Country | Customer | Issue | Action

Below that header row, fill in the Month, Country and Customer columns only,
repeating each customer on 3 consecutive rows, leaving Issue and Action empty.
Use these Month, Country and Customer values, each repeated 3 times:
2026-07 | XC | SLB/Cameron

Then convert the whole block including the header row into an Excel Table and
name that table exactly: SPRIssues_XC
Set column D and column E to about 60 characters wide with text wrapping on.
Set columns A, B and C to about 12 characters wide.

Next, leaving two empty rows after that table, add a second header row with
these 8 headers, exactly:
Month | Country | Customer | Payments In House Pending | Expected Payments | Forecasted Overdue EOM | Forecasted GT60 EOM | Forecasted GT90 EOM

Below it add one row per customer, filling only Month, Country and Customer:
2026-07 | XC | SLB/Cameron

Convert that block including its header row into an Excel Table named exactly:
SPRForecast_XC
Format the five forecast amount columns as number with thousands separator and
no decimal places.

Do not add any other sample data.
```
