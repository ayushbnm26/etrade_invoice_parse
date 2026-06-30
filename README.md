# Invoice Processing

Internal Streamlit utility for parsing one supported Etrade invoice PDF at a time and generating a public-facing Excel workbook for the team.

This project also keeps the local CLI workflow for batch processing PDFs from `input_pdfs/`.

## What the Streamlit App Does

- A team member opens the deployed Streamlit app.
- The app asks for a simple app password.
- The team member uploads one supported Etrade invoice PDF.
- The app validates that the upload is a non-empty PDF-looking file.
- The parser generates the normal internal workbook and the public-facing workbook.
- The team member can download only the public-facing workbook.
- The admin receives an email with the original PDF, public workbook, detailed/internal workbook, and runtime log if available.
- Every generated Excel file must be manually verified before finance, accounting, GST, reconciliation, reporting, payment, audit, or official business use.

## Important Accuracy Notice

This tool is designed specifically for Etrade invoice PDFs that follow the supported invoice layout pattern.

It works best when the uploaded PDF has the same or very similar structure, table format, field placement, and text quality as the sample Etrade invoices used during development. A major change in invoice design, column order, address placement, tax layout, scanned-image quality, or page structure can significantly reduce extraction accuracy.

This app uses PDF text and table extraction through `pdfplumber`; it is not an accounting-certified OCR system. Scanned-image PDFs or OCR-dependent documents may fail or be less reliable. The generated Excel file must be checked by a human before it is used for any official business process.

## Project Layout

```text
input_pdfs/                  Optional local CLI input folder
output/                      Optional local CLI output folder
logs/                        Optional local CLI log folder
src/invoice_processing/      Application package
tests/                       Unit tests
app.py                       Streamlit web app
main.py                      Local CLI runner
```

## Local Run

Install dependencies:

```bash
python -m pip install -r requirements.txt
```

For local Streamlit testing, create a local secrets file from the placeholder example:

```bash
copy .streamlit\secrets.example.toml .streamlit\secrets.toml
```

Edit `.streamlit/secrets.toml` and replace every placeholder. Never commit `.streamlit/secrets.toml` to GitHub.

Run the Streamlit app locally:

```bash
python -m streamlit run app.py
```

The dashboard uses an isolated temporary folder for each processing run. It does not write uploaded PDFs into `input_pdfs/` and does not rely on permanent cloud filesystem storage.

## Local CLI Run

The existing CLI behavior is still available:

```bash
python main.py
```

By default, the CLI reads PDFs from `input_pdfs/` and writes a timestamped workbook like:

```text
output/parsed_invoices_YYYYMMDD_HHMMSS.xlsx
```

It also writes one public-facing workbook per successfully parsed PDF:

```text
output/public_workbooks/<pdf_name>_public.xlsx
```

Optional CLI arguments:

```bash
python main.py --input-dir input_pdfs --output-dir output
python main.py --output-file parsed_invoices_latest.xlsx
python main.py --log-level DEBUG
```

## Required Streamlit Secrets

Use this exact shape in Streamlit Community Cloud secrets. Replace only the placeholder values.

```toml
APP_PASSWORD = "replace-with-team-password"

[SMTP]
HOST = "smtp.gmail.com"
PORT = 587
USERNAME = "ayushbnm26@gmail.com"
APP_PASSWORD = "replace-with-gmail-app-password"
FROM_EMAIL = "ayushbnm26@gmail.com"
ADMIN_EMAIL = "ayush.kumar@algorithmtrix.com"
```

Do not put real secrets in GitHub. The real `.streamlit/secrets.toml` file is ignored by `.gitignore`; keep it local only. In Streamlit Community Cloud, paste secrets into the app's secrets settings.

Official docs:

- Streamlit deployment: https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app
- Streamlit secrets management: https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/secrets-management

## Gmail App Password

Use a Gmail App Password, not the normal Gmail account password.

1. Open the Google Account for `ayushbnm26@gmail.com`.
2. Enable 2-Step Verification.
3. Create an App Password for mail/SMTP use.
4. Copy the generated 16-character app password.
5. Paste that app password into `SMTP.APP_PASSWORD` in Streamlit secrets.
6. Never commit the app password, normal Gmail password, or Streamlit app password to GitHub.

Official Google help: https://support.google.com/accounts/answer/185833

## Deploy for Free on Streamlit Community Cloud

Streamlit Community Cloud is the intended free deployment target for this app.

1. Push this repository to GitHub.
2. Open https://streamlit.io/cloud in your browser.
3. Sign in with GitHub.
4. Click the button to create or deploy a new app.
5. Choose the GitHub repository.
6. Choose branch `main`.
7. Set the main file path to `app.py`.
8. Open the app secrets/settings area.
9. Paste the full placeholder secrets block shown above, replacing the team password and Gmail app password.
10. Deploy the app.
11. Open the generated Streamlit app URL.
12. Test with one known supported Etrade invoice PDF before sharing it with the team.
13. Share the app URL and the app password only with authorized team members.

No paid hosting, database, queue, storage bucket, paid OCR API, or paid email service is required.

## How the Team Uses It

1. Open the Streamlit app URL.
2. Enter the app password.
3. Read the usage notice.
4. Upload one supported Etrade invoice PDF.
5. Click `Process Invoice`.
6. Download the public Excel workbook.
7. Manually verify every important field before using the workbook.

Normal users do not receive the detailed/internal workbook and do not see technical logs in the app.

## What the Admin Receives

After successful processing, the configured admin email receives:

- Original uploaded PDF.
- Public-facing workbook sent to the team member.
- Detailed/internal workbook generated by the processor.
- Runtime log file, if available.

If an attachment is unavailable, the email body says which artifact was unavailable and why. If email delivery fails, the team member can still download the public workbook, and the app shows only this non-sensitive warning:

```text
Workbook generated, but admin email notification failed.
```

## Output Sheets

Normal internal workbook:

- `Invoice_Header`
- `Invoice_Line_Items`
- `Invoice_Summary`
- `Validation`
- `Exceptions`

Public workbook:

- `Invoice Details` with shipping address, receiver shipping address, invoice number, and system reference number.
- `Items` with ASIN, quantity, price/unit, net amount, separate CGST/SGST/IGST rate and amount columns, other tax fallback columns, total amount, sub total, and total quantity.

## Limitations and Risks

- The parser is layout-dependent and works best on the supported Etrade invoice layout.
- Major invoice layout changes can produce wrong or incomplete extraction.
- Scanned-image PDFs may fail or be less reliable because this app uses PDF text/table extraction.
- Validation can fail even when a workbook is generated; review validation status and workbook contents manually.
- The generated workbook is not accounting-certified and is not a system of record.
- Streamlit Community Cloud free hosting is not enterprise-grade security. The password gate is a simple internal access gate, not SSO or enterprise authentication.
- Streamlit Community Cloud filesystem storage is not durable. The app uses temporary per-run files and does not permanently store uploaded PDFs.
- Multiple users can upload at the same time because each run uses an isolated temporary workspace, but there is no database, queue, or background worker.
- Large PDFs or large attachments may cause Gmail SMTP delivery failure. The public workbook download still remains available if processing succeeded.

## Tests

```bash
python -m unittest discover -s tests
```
