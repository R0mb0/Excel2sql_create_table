# Excel to SQL CREATE TABLE PowerShell Script

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/ee9146d3911445c2b781cdd0dd537161)](https://app.codacy.com/gh/R0mb0/Excel2sql_create_table/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade)

[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/R0mb0/Excel2sql_create_table)
[![Open Source Love svg3](https://badges.frapsoft.com/os/v3/open-source.svg?v=103)](https://github.com/R0mb0/Excel2sql_create_table)
[![MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/license/mit)

[![Donate](https://img.shields.io/badge/PayPal-Donate%20to%20Author-blue.svg)](http://paypal.me/R0mb0)

## Overview

This project provides a PowerShell script (`ExcelToSqlCreateTable.ps1`) to automatically generate a SQL `CREATE TABLE` statement from the structure and content of an Excel `.xlsx` file. The script analyzes the sheet's headers and data, deduces the most appropriate SQL data types for each column, and handles common Excel issues such as duplicate headers, unnamed columns, and empty cells.

## Features

- **Automatic column type detection:** Detects `INT`, `FLOAT`, `DATETIME`, or `NVARCHAR` types based on column content.
- **Handles duplicate and unnamed columns:** Ensures all column names are unique, even if the Excel sheet contains duplicates or blanks.
- **Multi-file support:**  
  - If multiple `.xlsx` files are present in the script directory, the script lists them and prompts you to select which one to use.
  - If only one `.xlsx` file is present, it is selected automatically.
- **User-friendly prompts:** Asks for sheet name, SQL table name, output file name, and type detection threshold.
- **Output preview:** Displays the generated SQL statement in the console as well as saving it to a file.
- **Robust error handling:** Provides clear error messages for missing files, unreadable sheets, or missing columns.

## Prerequisites

- **PowerShell:** The script is designed for Windows PowerShell (tested on 5.1+) and PowerShell Core.
- **ImportExcel module:** The script uses the [ImportExcel](https://github.com/dfinke/ImportExcel) PowerShell module.
  - If not present, the script will attempt to install it automatically.

## Installation

1. **Clone or download** this repository to your local machine.
2. Place your Excel `.xlsx` files in the same directory as the script.

## Usage

1. **Open PowerShell** and navigate to the directory containing the script and your Excel file(s).

2. **Run the script:**

    ```powershell
    .\ExcelToSqlCreateTable.ps1
    ```

    > If the script doesn't run, please execute this command, then run the script again.
    > 
    > ```powershell
    >Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
    > ```

3. **If multiple `.xlsx` files are present:**
    - The script will display a numbered list of all `.xlsx` files in the directory.
    - Enter the corresponding number to select your desired file.

4. **Follow the prompts:**
    - **Sheet name:** Enter the name of the Excel worksheet (not the file or table name).
    - **SQL table name:** Enter the desired name for the SQL table.
    - **Output file name:** Enter the output file for the SQL script (press Enter to use the default).
    - **Type detection threshold:** Enter a number (press Enter to use the default, usually 500).
    - The script will display detected columns and types, and create an output file with the SQL command.

5. **Check the output:**
    - The generated SQL statement is shown in the console and saved to the output file you specified.

## Example

Suppose you have two Excel files, `data1.xlsx` and `data2.xlsx`, in the same folder as the script. When you run:

```powershell
.\ExcelToSqlCreateTable.ps1
```

You will see output like:

```
Multiple .xlsx files found in the folder:
[0] data1.xlsx
[1] data2.xlsx
Enter the number of the Excel file to use:
```
Enter the number corresponding to the file you want to process, and continue as prompted.

## Troubleshooting

- **No Excel file found:** Ensure at least one `.xlsx` file is in the script directory.
- **Module ImportExcel not found:** The script will try to install it. If you encounter permission issues, install manually:
    ```powershell
    Install-Module -Name ImportExcel -Scope CurrentUser
    ```
- **Sheet not found:** Double-check the worksheet name is correct (case-sensitive).
- **Weird column names:** If your Excel has empty or duplicate headers, the script will auto-correct them (`UnnamedColumn`, `ColumnName_2`, etc.).

## License

MIT License. See [LICENSE](LICENSE) for details.

## Credits

Based on PowerShell scripting and the ImportExcel module by Doug Finke.
