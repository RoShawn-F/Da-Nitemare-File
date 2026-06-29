
# Rick and Morty API Data Collector

## Overview
This project pulls data from the [Rick and Morty API](https://rickandmortyapi.com/) and exports it into a structured Excel workbook with multiple sheets.

## What It Does
- Fetches **all** characters, locations, and episodes from the Rick and Morty API
- Handles pagination to retrieve complete datasets (826 characters, 126 locations, 51 episodes)
- Exports data into a formatted Excel file with three separate worksheets
- Replaces episode URLs with actual episode names for each character (Nightmare Mode!)

## Sheets
| Sheet | Records |
|-------|---------|
| Rick and Morty Characters | 826 |
| Rick and Morty Location | 126 |
| Rick and Morty Episode | 51 |

## Technologies Used
- Python
- `requests` — API calls
- `openpyxl` — Excel file creation
- `time` — Rate limiting between API calls

## How to Run
1. Install dependencies:
```bash
pip install requests openpyxl
```
2. Create a `spreadsheets` folder in the project directory
3. Run the script:
```bash
python exercise.py
```
4. Open `spreadsheets/exercise.xlsx`

## Key Concepts Used
- Paginated API requests using `while` loops
- Nested API calls to replace URLs with actual data
- Error handling with `try/except`
- Writing structured data to Excel with multiple worksheets

## Notes
- The script makes thousands of API calls due to Nightmare Mode — expect it to take several minutes to complete
- Rate limiting is handled with `time.sleep()` between requests
