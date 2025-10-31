r"""
FileComparer.py
Author: Matt Lockard
Date: 2025-10-31 Happy Halloween
      /\                 /\
     / \'._   (\_/)   _.'/ \
    /_.''._'--('.')--'_.''._\
    | \_ / `;=/ " \=;` \ _/ |
     \/ `\__|`\___/`|__/`  \/
             \(/|\)/    
Compares columns and values across multiple CSV or Excel files, reporting unique and duplicate values per file.
Usage:
    python FileComparer.py file1.csv file2.csv [file3.csv ...] [--output results.csv]
"""

import sys
import os
import re
from datetime import datetime

#Bootstrappin Bat
print(r"""
      /\                 /\
     / \'._   (\_/)   _.'/ \
    /_.''._'--('.')--'_.''._\
    | \_ / `;=/ " \=;` \ _/ |
     \/ `\__|`\___/`|__/`  \/
             \(/|\)/    
""")


# Bootstrapper for pandas and openpyxl
try:
    import pandas as pd
except ImportError:
    import subprocess
    print("pandas not found. Installing pandas...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pandas"])
    import pandas as pd

try:
    import openpyxl
except ImportError:
    import subprocess
    print("openpyxl not found. Installing openpyxl...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl

def normalize_col_exact(col):
    return col

def normalize_col_lower(col):
    return col.strip().lower()

def normalize_col_loose(col):
    return re.sub(r'[^a-zA-Z0-9]+', '', col.strip().lower())

def read_table(f):
    ext = os.path.splitext(f)[1].lower()
    try:
        if ext in ['.csv']:
            return pd.read_csv(f)
        elif ext in ['.xls', '.xlsx', '.xlsm']:
            return pd.read_excel(f, engine='openpyxl')
        else:
            raise Exception(f"Unsupported file type: {f}")
    except Exception as e:
        print(f"Error reading '{f}': {e}")
        return None  # Return None if error

def compare_all_columns(file_list, output_file="Results.csv", mode="exact"):
    if mode == "loose":
        normalize = normalize_col_loose
    elif mode == "lower":
        normalize = normalize_col_lower
    else:
        normalize = normalize_col_exact
    
        # Read all files into DataFrames with error handling
    dfs = []
    file_status = []  # Track status for each file
    for f in file_list:
        df = read_table(f)
        if df is not None:
            # Normalize column names
            df.columns = [normalize(col) for col in df.columns]
            dfs.append(df)
            file_status.append("ok")
        else:
            dfs.append(None)
            file_status.append("error")

    # Get union of all columns from successfully read files
    all_columns = set()
    for df in dfs:
        if df is not None:
            all_columns.update(df.columns)
    all_columns = list(all_columns)
    print(f"All normalized columns found: {all_columns}")

    results = []

    for col in all_columns:
        col_values = []
        col_duplicates = []
        files_with_col = []
        for i, df in enumerate(dfs):
            if df is None:
                col_values.append(set())
                col_duplicates.append([])
            elif col in df.columns:
                vals = df[col].dropna().astype(str).tolist()
                col_values.append(set(vals))
                dupes = df[col][df[col].duplicated(keep=False)].dropna().unique().tolist()
                col_duplicates.append(dupes)
                files_with_col.append(i)
            else:
                col_values.append(set())
                col_duplicates.append([])

        # Skip this column if all files are blank for it or if no file contains the column
        valid_col_files = [df for df in dfs if df is not None and col in df.columns]
        if not valid_col_files or all(len(vals) == 0 for vals in col_values):
            continue

        # Find values only in each file for this column
        only_in_files = []
        for i, vals in enumerate(col_values):
            others = set.union(*(col_values[:i] + col_values[i+1:]))
            only_in = vals - others
            only_in_files.append(';'.join(only_in))

        # Prepare result row with short file names
        row = {'Normalized Column': col}
        for i, file in enumerate(file_list):
            file_short = os.path.splitext(os.path.basename(file))[0]
            if file_status[i] == "error":
                row[f'Only in {file_short} (Exact)'] = 'File error'
                row[f'Duplicates in {file_short}'] = 'File error'
            elif i in files_with_col:
                row[f'Only in {file_short} (Exact)'] = only_in_files[i]
                row[f'Duplicates in {file_short}'] = ';'.join(col_duplicates[i])
            else:
                row[f'Only in {file_short} (Exact)'] = 'Column not present'
                row[f'Duplicates in {file_short}'] = 'Column not present'
        results.append(row)

    # Convert results to DataFrame and save
    results_df = pd.DataFrame(results)
    results_df.to_csv(output_file, index=False)
    print(f"Results saved to {output_file}")
    print(f"Compared {len(all_columns)} normalized columns across {len([f for f in file_status if f == 'ok'])} files.")
    for i, status in enumerate(file_status):
        if status == "error":
            print(f"Skipped file due to error: {file_list[i]}")

if __name__ == "__main__":

    # Help option
    if "--help" in sys.argv:
        print("""
FileComparer.py - Compare columns and values across multiple CSV or Excel files.

Usage:
    python FileComparer.py [file1.csv file2.xlsm ...] [--output results.csv] [--mode exact|lower|loose]

Options:
    --output <filename>   Specify output CSV file name (default: results_<datetime>.csv in ./Results)
    --mode <mode>         Column normalization mode:
                            exact  - No normalization (default)
                            lower  - Force lowercase
                            loose  - Lowercase and remove non-alphanumeric characters
    --help                Show this help message

Behavior:
    - If no files are provided, all supported files in ./placefileshere are processed.
    - Results are saved in ./Results.
    - Skips unreadable files and marks them as "File error" in results.
    - Excludes columns with only blank values.

Examples:
    python FileComparer.py
    python FileComparer.py data1.csv data2.xlsm --mode loose --output myresults.csv
    python FileComparer.py --help

Happy Halloween! 🦇
""")
        sys.exit(0)

    
    default_dir = "./placefileshere"
    results_dir = "./results"

    mode = "exact"
    if "--mode" in sys.argv:
        idx = sys.argv.index("--mode")
        if idx + 1 < len(sys.argv):
            mode = sys.argv[idx + 1].lower()
            del sys.argv[idx:idx+2]

    if "--output" in sys.argv:
        idx = sys.argv.index("--output")
        output_file = sys.argv[idx + 1]
        file_args = [arg for i, arg in enumerate(sys.argv[1:]) if arg != "--output" and arg != output_file]
    else:
        datestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"results_{datestamp}.csv"
        file_args = sys.argv[1:]

    if not file_args:
        print(f"No files provided. Reading all supported files in directory: {default_dir}")
        supported_exts = {'.csv', '.xls', '.xlsx', '.xlsm'}
        file_args = [
            os.path.join(default_dir, f)
            for f in os.listdir(default_dir)
            if os.path.splitext(f)[1].lower() in supported_exts
        ]
        if not file_args:
            print(f"No supported files found in directory: {default_dir}")
            sys.exit(1)

    os.makedirs(results_dir, exist_ok=True)
    output_file_path = os.path.join(results_dir, output_file)

    compare_all_columns(file_args, output_file_path, mode=mode)

