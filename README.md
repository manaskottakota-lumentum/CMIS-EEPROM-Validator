# CMIS EEPROM Validator

Python-only CMIS EEPROM validation utility for comparing an EEPROM/register dump against a CMIS-style spreadsheet.

## What It Does

- Loads an EEPROM dump from `.txt`, `.hex`, `.dump`, or `.log`.
- Loads a specification spreadsheet from `.xlsx`, `.csv`, or `.tsv`.
- Slices bytes by page/address/length and compares the module value against the expected spec value.
- Generates a CMIS-formatted specification spreadsheet from an Arista/SONiC-style IDPROM dump.
- Exports validation results for review.

## Run

Run the packaged Windows app:

```powershell
cd "C:\Users\kot107695\OneDrive - Lumentum Operations LLC\Documents\CMIS-EEPROM-Validator"
.\dist\CMIS-EEPROM-Validator.exe
```

Run the Python Tkinter app directly:

```powershell
cd "C:\Users\kot107695\OneDrive - Lumentum Operations LLC\Documents\CMIS-EEPROM-Validator"
& "C:\Users\kot107695\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" .\cmis_eeprom_validator_tk.py
```

## Quick Check

```powershell
& "C:\Users\kot107695\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" .\cmis_eeprom_validator.py --self-test
```

## Folder Layout

- `cmis_eeprom_validator_tk.py` - Tkinter GUI.
- `cmis_eeprom_validator.py` - CMIS dump parsing, spreadsheet loading, and validation logic.
- `workbook_generator.py` - generates validator-ready `.xlsx` specification spreadsheets from EEPROM dumps.
- `samples/` - bundled sample dumps and spreadsheets used by the app.
- `dist/` - standalone Windows EXE build.

## Spreadsheet Format

The app supports MSFT-style CMIS workbook tabs and simpler validator spreadsheets. For generated specification spreadsheets, the key columns are:

```text
Byte, Hex, Size(bytes), Bits, Field Name, Type, Dump Value, Converted Value, Description
```

Validation uses the page/sheet, byte/address, bit range, and `Dump Value` to determine whether each row passes. `Converted Value` is a type-aware interpretation of the raw dump bytes, such as ASCII text, decimal numbers, decoded bit values, or engineering units where the CMIS field definition is known.

## Notes

- Lower memory maps to bytes `00h-7Fh`.
- Upper pages map to bytes `80h-FFh`.
- Generated specs include CMIS-defined rows for supported pages including Lower, 00h, 01h, 02h, 03h, 10h, 11h, 13h, and 2Fh when available in the dump/template.
- `Load Arista Sample` loads the bundled Arista dump and generated spec sample.
