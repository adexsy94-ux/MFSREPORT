# Payment Test Management Report Generator

This project converts a **Flutterwave-style transaction CSV export** into a management-ready Excel report for tracking:

- successful payments;
- failed payment attempts;
- payments awaiting validation;
- payment methods;
- source banks and card issuers;
- source countries;
- gateway failure messages;
- weekly payment-test trends;
- settlement information; and
- the original raw transaction export.

The generated workbook follows the management-report format used for the payment testing review.

## Output workbook

The Python script creates an Excel workbook with these sheets:

1. **Management Summary**  
   Executive KPIs, currency-level performance, payment-method performance, observations, and charts.

2. **Weekly Trend**  
   Weekly payment attempts split into successful, failed, and pending transactions.

3. **Payment Sources**  
   Analysis by payment method, bank/card issuer, country, and instrument.

4. **Failed & Pending**  
   All exceptions requiring review, with gateway messages and tester/customer details.

5. **Test Transactions**  
   Clean transaction-level records for the selected reporting period.

6. **Historical Overview**  
   A high-level view of the complete CSV export, including transactions outside the management review period.

7. **Raw Data**  
   The source CSV preserved in the workbook for audit/reference.

## Default reporting period

By default, the script finds the **latest transaction date in the CSV** and reports on the latest **21 calendar days**, including that date.

For example, if the latest transaction date is **14 September 2026**, the default report covers:

**25 August 2026 to 14 September 2026**

Older rows remain available under **Historical Overview** and **Raw Data** but are not included in the management-period KPIs.

## Requirements

- Python 3.10 or later recommended
- `XlsxWriter`

Install the dependency with:

```bash
pip install -r requirements.txt
```

## Project files

```text
payment_report_generator/
├── generate_payment_report.py
├── requirements.txt
└── README.md
```

## Setup on Windows

Open **Command Prompt** or **PowerShell** in the project folder.

Create a virtual environment:

```bash
python -m venv .venv
```

Activate it in Command Prompt:

```bat
.venv\Scripts\activate
```

Or in PowerShell:

```powershell
.venv\Scripts\Activate.ps1
```

Install the requirements:

```bash
pip install -r requirements.txt
```

## Basic usage

Put the transaction CSV in the same folder as the script, then run:

```bash
python generate_payment_report.py "tx_download.csv"
```

The script will automatically create an Excel report beside the CSV.

Example output filename:

```text
Management_Payment_Test_Report_25Aug-14Sep2026.xlsx
```

## Specify the output filename

```bash
python generate_payment_report.py "tx_download.csv" --output "Management Payment Report.xlsx"
```

## Use a different automatic number of days

For the latest 30 days:

```bash
python generate_payment_report.py "tx_download.csv" --days 30
```

For the latest 7 days:

```bash
python generate_payment_report.py "tx_download.csv" --days 7
```

## Use an exact reporting period

```bash
python generate_payment_report.py "tx_download.csv" \
  --start-date 2026-08-25 \
  --end-date 2026-09-14
```

On Windows Command Prompt, this can also be entered on one line:

```bat
python generate_payment_report.py "tx_download.csv" --start-date 2026-08-25 --end-date 2026-09-14
```

The end date is inclusive.

## Change the report title

```bash
python generate_payment_report.py "tx_download.csv" --title "HOUSE OF FAITH PAYMENT TEST REPORT"
```

## Expected input columns

The script requires these columns:

```text
created
status
amount
currency
paymenttype
```

For a richer source analysis, the Flutterwave export can also include columns such as:

```text
txref
txid
chargemessage
pcardname
pcardcountry
pcardtype
paccountbankname
paccountfirstname
paccountlastname
custname
custemail
settlement_status
settlement_amount
originatorname
bankname
originatoraccountnumber
```

The optional columns are used when available. If an optional field is missing, the report leaves that detail blank or uses the payment channel as the source.

## Status treatment

The script standardizes transaction statuses as follows:

```text
successful                  -> Successful
failed                      -> Failed
success-pending-validation  -> Pending Validation
```

Other statuses are retained in readable title case.

## Payment-source logic

### Bank transfer

The report tries to identify the source using:

1. `bankname`;
2. `paccountbankname`;
3. `originatorname`; and
4. the available account information.

Any unmasked source account number written to the management report is automatically masked.

### Card

Card sources are taken from:

- `pcardname`;
- `pcardcountry`; and
- `pcardtype`.

This makes it possible to understand whether tests originated from cards issued in the UK, Nigeria, Canada, the USA, Kenya, or other countries available in the export.

### Apple Pay

Apple Pay is grouped as:

```text
Apple Pay / Internet Banking
```

### UK ACH

`account-ach-uk` is presented as:

```text
UK ACH / Internet Banking
```

## Currency treatment

The report **does not add different currencies together**.

NGN, GBP, USD, CAD, EUR, and other currencies are reported separately so that management totals remain financially meaningful.

For each currency the summary shows:

- total attempts;
- successful attempts;
- failed attempts;
- pending attempts;
- successful transaction value;
- failed transaction value; and
- pending transaction value.

## Weekly trend

The selected reporting period is divided into consecutive 7-day blocks starting from the report start date.

For a 21-day report this produces:

```text
Week 1
Week 2
Week 3
```

Each week shows:

- total attempts;
- successful payments;
- failed attempts;
- pending validation; and
- overall success rate.

## Failed-attempt analysis

The **Failed & Pending** sheet includes:

- transaction date/time;
- transaction reference;
- amount and currency;
- payment method;
- bank/card source;
- country;
- status;
- gateway message;
- tester/customer name; and
- customer email.

It also contains a summary of the gateway failure/pending messages, making recurring problems easier to spot.

Examples may include:

```text
Blocked Card
Restricted Card
transaction is pending
Pending user action
Cardholder browser session timed out
```

The actual report uses whatever messages are present in the input CSV.

## Data privacy

The report generator automatically masks a plain source account number before writing it to the management-facing transaction sheet.

The **Raw Data** tab intentionally retains the original CSV content for audit purposes. If the workbook will be distributed outside the authorised management team, consider removing the Raw Data sheet or sanitising the source CSV first.

## Run it against any new export

The normal workflow is:

1. Export the latest payment transaction CSV.
2. Save it to the project folder.
3. Run:

```bash
python generate_payment_report.py "your_new_export.csv"
```

4. Open the generated `.xlsx` file.
5. Review the **Management Summary** and **Failed & Pending** tabs before sharing with management.

No changes to the Python code should be required when the CSV follows the same Flutterwave-style structure.

## Troubleshooting

### `ModuleNotFoundError: No module named 'xlsxwriter'`

Run:

```bash
pip install -r requirements.txt
```

### CSV is missing required columns

The script will display the missing column names.

Confirm that the input file contains at least:

```text
created,status,amount,currency,paymenttype
```

### No transactions in the selected period

Either allow the script to use the automatic latest-21-day window:

```bash
python generate_payment_report.py "tx_download.csv"
```

or choose dates that exist in the CSV:

```bash
python generate_payment_report.py "tx_download.csv" --start-date YYYY-MM-DD --end-date YYYY-MM-DD
```

### Excel warns that the file already exists

Use another output filename:

```bash
python generate_payment_report.py "tx_download.csv" --output "report_v2.xlsx"
```

## Example used for the current report

For the supplied export, the command is:

```bash
python generate_payment_report.py "tx_download_3382492_2026_09_15T08_45_37_129Z(1).csv"
```

Because the latest transaction date in that export is 14 September 2026, the automatic 21-day report covers 25 August through 14 September 2026.

## Notes for future expansion

The script is structured so further analysis can be added later, such as:

- settlement reconciliation;
- success rate by card country;
- success rate by issuer;
- fee analysis;
- donation campaign/source analysis;
- comparison against the previous test period; or
- automatic PDF management summaries.
