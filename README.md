# CMIS EEPROM Validator

Standalone CMIS EEPROM validation utility for reading a CMIS EEPROM/register dump into a CMIS Rev 5.0 memory-map spreadsheet and, when expected values are supplied, validating actual values against them.

The current frontend is a WPF/XAML/C# MVVM app. The validation backend remains Python and is exposed through a small local HTTP REST API.

## Scope

The first milestone is intentionally small: load one hex dump, load one spreadsheet, slice bytes using `Page`, `Address`, and `Length`, and validate active rows where `Check` is `YES`. Additional CMIS rows can be added by adding rows to the spreadsheet.

The UI defaults to active checks only. Use `Show all rows` to display workbook rows where `Check` is `No` or blank.

## Run

Start the WPF frontend:

```powershell
cd "C:\Users\kot107695\OneDrive - Lumentum Operations LLC\Documents\CMIS-EEPROM-Validator\CmisEepromValidator.Wpf"
dotnet run
```

No Python package install is required. The WPF app automatically starts the local Python REST backend on launch and uses the installed .NET SDK without extra NuGet packages.

If the app cannot locate Python, set `CMIS_VALIDATOR_PYTHON` to a valid `python.exe` path before launching.

The legacy Python/Tkinter prototype can still be run directly if needed:

```powershell
& "C:\Users\kot107695\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" .\cmis_eeprom_validator.py
```

## Quick Check

```powershell
python .\cmis_eeprom_validator.py --self-test
```

The self-test loads the included sample EEPROM dump and spreadsheet CSV, reads Lower Memory bytes `00h-02h`, and verifies that the populated value is `19 52 04`.

## Spreadsheet Columns

The spreadsheet can be `.xlsx`, `.csv`, or `.tsv`.

CMIS memory-map columns:

```text
Check,Description,Bank,Page,Address,Length,CMIS Type,Simple Type,Value
```

Behavior:

- The workbook reader ignores the overview `CMIS Map` sheet and processes dedicated page tabs such as `Lower`, `00H`, `01H`, `02H`, and `03H`.
- Only rows with `Check` set to `YES` are active pass/fail checks.
- Rows with `Check` set to `No` or blank remain visible as `NOT CHECKED` reference rows.
- `Page` should be `Lower` or `Upper`; the tool does not infer page from the description.
- `Address` can be a lower-memory range such as `00h-02h` or an upper-memory page/window such as `01h:80h-FFh`.
- `Length` can be a byte count such as `3 bytes` or a simple number such as `3`.
- `Description` is used as the row label and detail explanation.
- `CMIS Type` is kept as the authoritative spec type.
- `Simple Type` is kept as a programming-style helper.
- For active `YES` rows, non-empty `Value` is treated as the current reference value to compare against.

Optional comparison columns are still supported:

```text
parameter,expected,comparator,tolerance,min,max,mask,significance
```

Supported comparison rules when `expected` is supplied:

- `exact`
- `range`
- `tolerance`
- `mask`

Supported data types:

- `ASCII`
- `U8`, `U16`, `U32`, `U64`
- `S8`, `S16`, `S32`, `S64`
- `Decimal`
- `Unsigned Integer`
- `Signed Integer`
- `Float32`
- `Hex`
- `Bit Field`

## Address Notes

- `CMIS Lower Page` maps to lower memory offsets `00h-7Fh`.
- `CMIS Upper Page XXh` maps to upper memory offsets `80h-FFh`.
- For upper pages, `start` may be written either as an absolute CMIS address like `0x81` or as a page-relative offset like `0x01`.
- Page names like `00h`, `01h`, and `10h` are hex page numbers.
