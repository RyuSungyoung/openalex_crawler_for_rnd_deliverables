# OpenAlex Crawler for R&D Deliverables

A Python tool for enriching an Excel bibliographic dataset with metadata from the [OpenAlex API](https://openalex.org/). It matches works by DOI, preserves existing spreadsheet columns, appends OpenAlex metrics, and flags records that may require manual review.

## Features

- Loads an Excel workbook and processes the `DOI 찾기` sheet.
- Normalizes DOI values, including `doi:`, `doi.org`, URL, and punctuation variants.
- Queries OpenAlex in batches of up to 100 DOIs.
- Appends title, first author, journal, publication date, cited-by count, FWCI, OpenAlex ID, DOI, and update date.
- Cross-checks title, author, journal, and publication year against source data.
- Preserves existing workbook content and writes to `OpenAlex_Output.xlsx`.
- Reports invalid DOIs, unmatched records, and records that need review.
- Supports an optional `OPENALEX_API_KEY` environment variable.

## Repository Contents

| File | Description |
| --- | --- |
| [`openalex_update.py`](openalex_update.py) | Standalone Python script for processing an Excel workbook. |
| [`OpenAlex_Update_VSCode.ipynb`](OpenAlex_Update_VSCode.ipynb) | Notebook version with setup, execution, and review examples. |
| `OpenAlex_Input.xlsx` | Expected input workbook. It is not included in this repository. |
| `OpenAlex_Output.xlsx` | Generated enriched workbook. It is not included in this repository. |

## Requirements

- Python 3.10 or newer
- `requests`
- `openpyxl`

The notebook also uses `truststore` for Windows system certificate-store support.

## Installation

```bash
python -m pip install requests openpyxl
```

Optional Windows SSL support:

```bash
python -m pip install truststore
```

## Input Workbook Format

Place `OpenAlex_Input.xlsx` in the repository or project directory. The workbook must contain a sheet named `DOI 찾기` with these header names:

| Korean header | Purpose |
| --- | --- |
| `논문명` | Publication title |
| `주저자명` | Principal/first author |
| `학술지명` | Journal name |
| `게재일` | Publication date |
| `DOI` | DOI used for OpenAlex lookup |

The script preserves existing columns and appends the following columns from column J onward:

- `OA 논문명`
- `OA 주저자`
- `OA 학술지명`
- `OA 게재일`
- `Cited by`
- `FWCI`
- `OpenAlex ID`
- `OA DOI`
- `OpenAlex Updated`
- `Cross-check`

## Usage

### Python script

Run from the directory containing `openalex_update.py` and `OpenAlex_Input.xlsx`:

```bash
python openalex_update.py
```

The result is saved as `OpenAlex_Output.xlsx`.

### Jupyter notebook

Open `OpenAlex_Update_VSCode.ipynb` in VS Code or Jupyter, ensure the dependencies are installed, and run all cells. The notebook uses the same input, sheet, and output names by default.

### Optional API key

An API key is not required but can be supplied through an environment variable:

**Windows PowerShell**

```powershell
$env:OPENALEX_API_KEY="your-api-key"
python openalex_update.py
```

**macOS/Linux**

```bash
export OPENALEX_API_KEY="your-api-key"
python openalex_update.py
```

## Cross-check Statuses

| Status | Meaning |
| --- | --- |
| `OK` | No configured mismatch was detected. |
| `CHECK: ...` | A title, author, journal, or publication-year difference needs review. |
| `NOT FOUND` | No matching work was found in OpenAlex. |
| `INVALID DOI` | The DOI value could not be normalized into a valid DOI format. |

The checks use similarity thresholds and are review aids rather than definitive bibliographic validations:

- Title: 90% for `OK`; below 75% is a strong mismatch.
- Journal: 85% for `OK`; below 65% is a strong mismatch.
- Author: compares the OpenAlex first author against the source author cell.
- Year: compares the source year with OpenAlex's publication year.

## Configuration

The main settings are defined in `openalex_update.py`:

```python
INPUT_FILE = "OpenAlex_Input.xlsx"
OUTPUT_FILE = "OpenAlex_Output.xlsx"
SHEET_NAME = "DOI 찾기"
HEADER_ROW = 1
BATCH_SIZE = 100
MAX_RETRIES = 4
REQUEST_DELAY = 0.25
```

Update these constants if your filenames, sheet name, header row, or request settings differ.

## API Behavior

- Multiple DOIs are combined into one OpenAlex filter per batch.
- Requests return title, publication date, authorships, primary source, cited-by count, FWCI, ID, DOI, and updated date.
- HTTP 429 and 5xx responses trigger exponential backoff retries.
- Request failures are retried up to `MAX_RETRIES` times.

## License

No license file is currently included in this repository.
