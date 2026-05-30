# safetydronereportparser

Converts SafetyDrone **Operations Report** PDFs into CSV and XLSX.

## Usage

```
python parse_flights.py [--comments] <input.pdf> [output.csv]
```

- Output defaults to the same basename as the input with a `.csv` extension.
- A matching `.xlsx` is always written alongside the CSV.
- `--comments` prepends report metadata as `#` comment lines in the CSV.

## Output

| File | Contents |
|------|----------|
| `.csv` | UTF-8 BOM, fully quoted, CRLF, Excel-compatible |
| `.xlsx` | **Flights** sheet (frozen header, auto-width) + **Metadata** sheet |

Each flight row captures: date/time, flight name, drone, duration, location (address + lat/lon), flight type, personnel, weather conditions, IGC/KML links, and notes.

## Setup

```
poetry install
poetry run python parse_flights.py <input.pdf>
```