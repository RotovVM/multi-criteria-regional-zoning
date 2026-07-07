# Regional Zoning Tool

Python desktop application for multi-criteria zoning of Russian regions based on selected indicators, with statistical comparison of the best and worst quartile-based groups.

## Overview

The program loads regional indicator data, assigns each subject to one of four zones using quartile-based thresholds, aggregates disadvantage ranks across selected methodologies, and compares zone-based groups using statistical tests.

## Input data

### 1. Indicators file
Excel file with:
- region names;
- one column per methodology;
- interpretation direction for each indicator.

Example: `data_examples/Template_methode_test.xlsx`

### 2. Indices / morbidity file
Excel file with:
- region names;
- additional indices or morbidity indicators for analytical comparison.

Example: `data_examples/Template_indexs_test-4.xlsx`

## Output

The generated Excel report includes:
- separate sheets for each methodology;
- aggregated ranking sheet;
- combined zone/index worksheet;
- statistical comparison of Zone 1 and Zone 4;
- summary distribution table.

Example: `output_example/tt_example-2.xlsx`

## Methods

- Quartile-based zoning
- Aggregated disadvantage scoring
- Normality testing
- Welch’s t-test
- Mann–Whitney U test

## Tech stack

- Python
- pandas
- NumPy
- SciPy
- openpyxl
- Tkinter

## How to run

1. Install dependencies.
2. Open the application.
3. Load the indicators file.
4. Load the indices file.
5. Select the methods.
6. Choose the output path.
7. Generate the Excel report.
