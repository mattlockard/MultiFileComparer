# FileComparer

**Author:** Matt Lockard  
**Date:** 2025-10-31  
**Theme:** Happy Halloween! 🦇

## Overview

MultiFileComparer compares columns and values across multiple CSV or Excel files, reporting unique and duplicate values per file.  
It supports `.csv`, `.xls`, `.xlsx`, and `.xlsm` files.

## Features

- Compares any number of files.
- Handles missing columns and blank values.
- Reports values unique to each file and duplicates within each file.
- Outputs results to a timestamped CSV in the `Results` folder.
- Auto-installs required Python packages (`pandas`, `openpyxl`) if missing.
- Fun Halloween ASCII bat on startup!

## Installation

1. [Download Python](https://www.python.org/downloads/) (version 3.8 or higher).
2. (Optional) Create and activate a virtual environment:

## Usage

### Command Line
python multifilecomparer.py file1.csv file2.xlsm [file3.xlsx ...] [--output results.csv]
- If no files are provided, all supported files in `./PlaceFilesHere` are processed.
- Results are saved in the `./Results` folder as `results_<datetime>.csv` by default.

### Example
python multifilecomparer.py
Processes all files in `./PlaceFilesHere`.

python multifilecomparer.py data1.csv data2.xlsm --output myresults.csv
Processes specified files and saves results as `myresults.csv` in `./Results`.

### Batch File

You can use a batch file to automate running and virtual environment setup:
@echo off if not exist venv ( python -m venv venv ) call venv\Scripts\activate python multifilecomparer.py %* pause
## Requirements

- Python 3.8+
- `pandas` and `openpyxl` (auto-installed if missing)

## Folder Structure

- `FileComparer.py` — Main script
- `PlaceFilesHere/` — Put your input files here (if not specifying files)
- `Results/` — Output results are saved here

## Notes

- The script will skip files it cannot read and mark them as "File error" in the results.
- Columns with only blank values across all files are excluded from the results.
## Run Modes

You can control how column names are matched using the `--mode` argument:

- **exact** (default): No normalization. Column names must match exactly.
- **lower**: Forces all column names to lowercase (preserves spaces and punctuation).
- **loose**: Forces lowercase and removes all non-alphanumeric characters (spaces, underscores, dashes, etc.).

### Usage Examples

Run in **exact** mode (default):

## Output Example

The results CSV will contain columns like:

| Normalized Column | Only in file1 (Exact) | Duplicates in file1 | Only in file2 (Exact) | Duplicates in file2 | ... |
|-------------------|----------------------|---------------------|----------------------|---------------------|-----|
| email             | ...                  | ...                 | ...                  | ...                 | ... |


## Troubleshooting

- **Permission Denied:** Make sure the `Results` folder is not open in another program.
- **Missing Python or pip:** Ensure Python and pip are installed and available in your system PATH.
- **Excel file errors:** Make sure your Excel files are not password protected or corrupted.

## License

This project is licensed under the MIT License.

You are free to use, modify, distribute, and sell this software, provided that the original copyright and license notice are included in all copies or substantial portions of the software.  
See [https://opensource.org/licenses/MIT](https://opensource.org/licenses/MIT) for details.


## Contributing

Pull requests and suggestions are welcome!

## Contact

For questions or feedback, contact Matt Lockard.

## Halloween ASCII Bat
  /\                 /\
 / \'._   (\_/)   _.'/ \
/_.''._'--('.')--'_.''._\
| \_ / `;=/ " \=;` \ _/ |
 \/ `\__|`\___/`|__/`  \/
         \(/|\)/    
         

Happy comparing! 🎃🦇