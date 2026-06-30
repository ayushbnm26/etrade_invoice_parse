# Invoice Processing

Production-style Python project for parsing invoice PDFs into an Excel workbook.

## Project Layout

```text
input_pdfs/                  Place PDF invoices here
output/                      Generated Excel workbooks
logs/                        Runtime logs
src/invoice_processing/      Application package
tests/                       Lightweight unit tests
main.py                      Local CLI runner
```

## Setup

```bash
python -m pip install -r requirements.txt
```

## Run

```bash
python main.py
```

The default command reads PDFs from `input_pdfs/` and writes a timestamped workbook like:

```text
output/parsed_invoices_YYYYMMDD_HHMMSS.xlsx
```

It also writes one public-facing workbook per successfully parsed PDF:

```text
output/public_workbooks/<pdf_name>_public.xlsx
```

Optional arguments:

```bash
python main.py --input-dir input_pdfs --output-dir output
python main.py --output-file parsed_invoices_latest.xlsx
python main.py --log-level DEBUG
```

## Output Sheets

Normal workbook:

- `Invoice_Header`
- `Invoice_Line_Items`
- `Invoice_Summary`
- `Validation`
- `Exceptions`

PDF-level errors are captured in the `Exceptions` sheet so one bad file does not stop the whole batch.

Public workbook:

- `Invoice Details` with shipping address, receiver shipping address, invoice number, and system reference number.
- `Items` with ASIN, quantity, price/unit, net amount, separate CGST/SGST/IGST rate and amount columns, other tax fallback columns, total amount, sub total, and total quantity.

## Tests

```bash
python -m unittest discover -s tests
```
