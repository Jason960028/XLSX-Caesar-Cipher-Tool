# XLSX Caesar Cipher Tool

[![Python Version](https://img.shields.io/badge/python-3.x-blue.svg)](https://www.python.org/)
[![License: WTFPL](https://img.shields.io/badge/License-WTFPL-brightgreen.svg)](http://www.wtfpl.net/about/)

## Overview

**XLSX Caesar Cipher Tool** is a simple Python project that demonstrates how to encrypt and decrypt Excel (XLSX) files using a Caesar cipher. The tool:
- **Encrypts** each cell in an XLSX file by shifting its alphabetic characters.
- **Decrypts** the encrypted XLSX file by reversing the shift.
- Generates a **sample Excel file** (`input.xlsx`) with dummy data using `pandas`.

> **Note:** The Caesar cipher is a basic encryption technique and is not secure for real-world use.

## Features

- **Encryption:** Shifts alphabetic characters in each cell using a Caesar cipher (default shift: 3).
- **Decryption:** Reverses the Caesar cipher to restore original text.
- **Sample Data Creation:** Automatically creates a sample Excel file with test data.

## File Structure
```plaintext
XLSX Caesar Cipher Tool/
├── README.md                    # Project overview and instructions
├── LICENSE                      # License information
├── main
    ├── input.xlsx               # Sample Excel file created by the script
    ├── encrypted_input.xlsx     # Encrypted Excel file (generated from input.xlsx)
    ├── decrypted_output.xlsx    # Decrypted Excel file (generated from encrypted_input.xlsx)
    └── xslxStructEx.png         # Example excel file structure
```

## How it works
    
**Sample Data Generation:** The script first creates an Excel file (input.xlsx) using sample data (Name, Age, Email, Gender).

**Encryption:** The encrypt_xlsx_file function reads input.xlsx, encrypts each cell in the data rows using the Caesar cipher (default shift of 3), and writes the encrypted data to encrypted_input.xlsx.

**Decryption:** The decrypt_xlsx_file function reads encrypted_input.xlsx, decrypts each cell by reversing the cipher (shift by -3), and saves the result as decrypted_output.xlsx.

**INCASE OF USING SAPARATE xlsx FILE**
Make sure the xlsx file structure look like this
![Screenshot of XLSX File Structure](./main/xslxStructEx.png)



## Requirements

- Python 3.x

Install the required Python libraries using pip:

```bash
pip install openpyxl pandas




