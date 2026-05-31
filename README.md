# dronereportparser

Converts drone operations report PDFs into CSV and XLSX.

## Parsers

| Script | Format |
|--------|--------|
| `parse_safetydrone_pdf.py` | SafetyDrone Operations Report (columnar table layout) |
| `parse_airdata_pdf.py` | Airdata UAV Flight Report (labeled key:value layout) |

## Usage

```
python parse_safetydrone_pdf.py [--comments] <input.pdf> [output.csv]
python parse_airdata_pdf.py [--comments] <input.pdf> [output.csv]
```

- Output defaults to the same basename as the input with a `.csv` extension.
- A matching `.xlsx` is always written alongside the CSV.
- `--comments` prepends report metadata as `#` comment lines in the CSV.

## Output

| File | Contents |
|------|----------|
| `.csv` | UTF-8 BOM, fully quoted, CRLF, Excel-compatible |
| `.xlsx` | **Flights** sheet (frozen header, auto-width) + **Metadata** sheet |

## Setup

```
poetry install
poetry run python parse_safetydrone_pdf.py <input.pdf>
poetry run python parse_airdata_pdf.py <input.pdf>
```