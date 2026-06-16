# Product Matcher

Transfers sales data from a Bulgarian product table into a German product table inside an Excel file.

For each product in the Bulgarian table (columns H–K), it finds the matching product in the German table (column A) and copies the three values — Menge, EK-Preis, VK-Preis — into columns C, D, and E of that German row. Any Bulgarian product with no German counterpart is flagged with a red `<` in column L.

Matching works across languages and volume formats. The name translation lives in a separate `translations.xlsx` file that the user manually edits in Excel.

---

## Requirements

- [Node.js](https://nodejs.org)
- `translations.xlsx` — two-column Excel file: Bulgarian base name | German base name (not included in the repo for privacy reasons)

## Usage

```bash
node match.js <inputFile> <sheetName> [outputFile] [--map <translationsFile>]
```

**Write results back into the input file (default):**
```bash
node match.js "input.xlsx" "1"
```

**Write to a separate output file:**
```bash
node match.js "input.xlsx" "1" "result.xlsx"
```

**Use a translations file in a different location:**
```bash
node match.js "input.xlsx" "1" --map "C:\path\to\translations.xlsx"
```

> Close the Excel file before running — the script overwrites it in place and will fail if the file is locked by Excel.

---

## Output

The console prints a summary after every run:

```
Loaded 48 translation(s) from translations.xlsx.
Sheet "1": 125 German rows, 82 Bulgarian rows.

Transferred 77 product(s) into German C/D/E.
```

**NEW products** (flagged with `>>`) need a new row added to `translations.xlsx` — Bulgarian base name in column A, German base name in column B. Re-run afterward.

**Size gaps** (indented) are not errors — the product is known but that package size has no slot in the German table.

---

## Adding a new product

1. Open `translations.xlsx` in Excel.
2. Add a row: Bulgarian base name in column A, German base name in column B.
3. Save and re-run.