# SG\_Metadata\_Export\_Migration (Excel VBA)

Prepares a **raw ShareGate metadata export** for SharePoint migration mapping. The macro normalizes the sheet, builds a clean table, enforces column names and order, clears the Destination Library values, and removes everything else.

---

## Features

* Unhides rows/columns and clears filters
* Removes **empty** rows/columns and shifts the used range to **A1**
* Builds a new **Excel Table** over the data
* Renames headers (exact matches only):

  * `ContentType` → **Content Type**
  * `SourcePath` → **Source Location**
  * `DestinationPath` → **Destination Library**
  * `Column 1` (or the column to the right of `ID`) → **Folder or Filename**
* Ensures **Destination Folder** exists (immediately after **Destination Library**)
* **Clears all values** in **Destination Library** (keeps the header)
* Keeps **only** these columns, in this order (skips missing, deletes everything else):

  1. Content Type
  2. Source Location
  3. Folder or Filename
  4. Destination Library
  5. Destination Folder
  6. Created By
  7. Created
  8. Modified By
  9. Modified

---

## Requirements

* Excel for Windows with VBA enabled
* Raw ShareGate export (no existing table on the sheet)
* Header row in **row 1**

---

## Install

1. Download `SG_Metadata_Export_Migration.bas` from this folder.
2. In Excel: `Alt + F11` → **File** → **Import File…** → select the `.bas`.
3. (Optional) Rename the module to `mod_SG_Metadata_Export_Migration`.
4. Save your workbook as `.xlsm` if you want to keep the macro.

---

## Usage

1. Open the raw export sheet and **Enable Editing** if prompted.
2. Go to **Developer → Macros** → select `SG_Metadata_Export_Migration` → **Run**.
3. The sheet is normalized and reduced to the target columns in the order above.

**Notes**

* Values in **Destination Library** are cleared intentionally.
* Any column **not** in the keep list is deleted.

---

## Troubleshooting

* **Nothing happened**: The workbook might be in Protected View. Click **Enable Editing**.
* **Columns didn’t move/rename**: Confirm headers are spelled exactly as listed.
* **Error 1004 during move**: Ensure row 1 has no merged cells and filters are cleared.

---

## Version

* **v1 – August 15, 2025**

## Authors

* Trevor Styler — [https://trevor.styler.ca](https://trevor.styler.ca)
* ChatGPT 5

**Repo folder:** `Excel/ShareGate-Metadata-Export-Migration`
**Source:** [https://github.com/TrevorStyler/M365-Office-Macros/tree/main/Excel](https://github.com/TrevorStyler/M365-Office-Macros/tree/main/Excel)
