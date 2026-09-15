#!/usr/bin/env python3
"""
Payment Test Management Report Generator

Reads a Flutterwave-style transaction CSV and generates a management-ready
Excel report containing:
    1. Management Summary
    2. Weekly Trend
    3. Payment Sources
    4. Failed & Pending
    5. Test Transactions
    6. Historical Overview
    7. Raw Data

Default reporting period:
    The latest 21 calendar days available in the CSV, inclusive of the
    latest transaction date.

Example:
    python generate_payment_report.py transactions.csv

Custom period:
    python generate_payment_report.py transactions.csv \
        --start-date 2026-08-25 \
        --end-date 2026-09-14 \
        --output management_report.xlsx
"""

from __future__ import annotations

import argparse
import csv
import math
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path
from typing import Any, Iterable, Optional

import xlsxwriter


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

REQUIRED_COLUMNS = {
    "created",
    "status",
    "amount",
    "currency",
    "paymenttype",
}

NUMERIC_COLUMNS = {
    "amount",
    "chargedamount",
    "appfee",
    "merchantfee",
    "merchantbearsfee",
    "amountsettledforthistransaction",
    "settlement_amount",
    "originatoramount",
    "vat",
}

SUCCESS_STATUSES = {"successful"}
FAILED_STATUSES = {"failed"}
PENDING_STATUSES = {
    "success-pending-validation",
    "pending",
    "pending-validation",
    "pending_validation",
}

METHOD_LABELS = {
    "bank_transfer": "Bank Transfer",
    "card": "Card",
    "applepay": "Apple Pay",
    "account-ach-uk": "UK ACH / Internet Banking",
    "account": "Bank Account",
    "opay": "OPay",
}

TITLE_FILL = "#17365D"
HEADER_FILL = "#1F4E78"
SECTION_FILL = "#D9EAF7"
PALE_BLUE = "#EAF2F8"
PALE_YELLOW = "#FFF2CC"
PALE_GREEN = "#E2F0D9"
PALE_RED = "#FCE4D6"

WHITE = "#FFFFFF"
NAVY = "#17365D"
GREY = "#44546A"
DARK_GREEN = "#375623"
DARK_RED = "#C00000"
DARK_YELLOW = "#7F6000"


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class Transaction:
    created: datetime
    txref: str
    txid: str
    amount: float
    currency: str
    payment_method: str
    payment_method_label: str
    source: str
    source_country: str
    instrument: str
    status: str
    gateway_message: str
    customer_name: str
    customer_email: str
    originator_name: str
    source_account: str
    settlement_status: str
    settlement_amount: Optional[float]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def clean(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def safe_float(value: Any) -> Optional[float]:
    text = clean(value).replace(",", "")
    if not text:
        return None

    try:
        number = float(text)
        if math.isnan(number) or math.isinf(number):
            return None
        return number
    except (TypeError, ValueError):
        return None


def parse_datetime(value: Any) -> Optional[datetime]:
    """
    Parse common ISO-8601 timestamps from gateway CSV exports.

    Naive timestamps are treated as UTC. A trailing Z is converted to +00:00.
    """
    text = clean(value)
    if not text:
        return None

    candidates = [
        text,
        text.replace("Z", "+00:00"),
    ]

    for candidate in candidates:
        try:
            dt = datetime.fromisoformat(candidate)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            return dt.astimezone(timezone.utc)
        except ValueError:
            pass

    # Fallbacks for a few common export formats.
    formats = (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M:%S.%f",
        "%d/%m/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M:%S",
    )
    for fmt in formats:
        try:
            return datetime.strptime(text, fmt).replace(tzinfo=timezone.utc)
        except ValueError:
            pass

    return None


def normalize_status(raw_status: str) -> str:
    status = clean(raw_status).lower()

    if status in SUCCESS_STATUSES:
        return "Successful"
    if status in FAILED_STATUSES:
        return "Failed"
    if status in PENDING_STATUSES:
        return "Pending Validation"
    if not status:
        return "Unknown"

    return status.replace("_", " ").replace("-", " ").title()


def normalize_method(raw_method: str) -> tuple[str, str]:
    method = clean(raw_method).lower()
    return method, METHOD_LABELS.get(method, method.replace("_", " ").title() or "Other")


def mask_account(value: str) -> str:
    """
    Preserve an already-masked account number. If an unmasked number is present,
    mask the middle digits before writing it to the management report.
    """
    text = clean(value)
    if not text:
        return ""

    if "*" in text:
        return text

    digits = re.sub(r"\D", "", text)
    if len(digits) <= 6:
        return text

    return f"{digits[:3]}{'*' * (len(digits) - 6)}{digits[-3:]}"


def choose_first(row: dict[str, str], *columns: str) -> str:
    for column in columns:
        value = clean(row.get(column))
        if value:
            return value
    return ""


def card_source_country(raw: str) -> str:
    return clean(raw)


def normalize_transaction(row: dict[str, str]) -> Optional[Transaction]:
    created = parse_datetime(row.get("created"))
    if created is None:
        return None

    method, method_label = normalize_method(row.get("paymenttype", ""))
    currency = clean(row.get("currency")).upper()
    amount = safe_float(row.get("amount")) or 0.0

    source = ""
    source_country = ""
    instrument = ""
    originator_name = ""
    source_account = ""

    if method == "card":
        source = choose_first(row, "pcardname") or "Card"
        source_country = card_source_country(row.get("pcardcountry", ""))
        instrument = choose_first(row, "pcardtype") or "Card"

    elif method == "bank_transfer":
        source = choose_first(row, "bankname", "paccountbankname") or "Bank Transfer"
        source_country = "Nigeria" if currency == "NGN" else ""
        instrument = "Bank Transfer"
        originator_name = choose_first(row, "originatorname")
        if not originator_name:
            originator_name = " ".join(
                part
                for part in [
                    clean(row.get("paccountfirstname")),
                    clean(row.get("paccountlastname")),
                ]
                if part
            )
        source_account = mask_account(
            choose_first(row, "originatoraccountnumber", "paccountnumber")
        )

    elif method == "applepay":
        source = "Apple Pay / Internet Banking"
        instrument = "Apple Pay"

    elif method == "account-ach-uk":
        source = "UK Internet Banking (ACH)"
        source_country = "United Kingdom"
        instrument = "ACH UK"

    else:
        source = method_label
        instrument = method_label
        source_country = clean(row.get("pcardcountry"))

    return Transaction(
        created=created,
        txref=clean(row.get("txref")),
        txid=clean(row.get("txid")),
        amount=amount,
        currency=currency,
        payment_method=method,
        payment_method_label=method_label,
        source=source,
        source_country=source_country,
        instrument=instrument,
        status=normalize_status(row.get("status", "")),
        gateway_message=clean(row.get("chargemessage")),
        customer_name=clean(row.get("custname")),
        customer_email=clean(row.get("custemail")),
        originator_name=originator_name,
        source_account=source_account,
        settlement_status=clean(row.get("settlement_status")),
        settlement_amount=safe_float(row.get("settlement_amount")),
    )


def read_csv(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    if not path.exists():
        raise FileNotFoundError(f"Input file does not exist: {path}")

    with path.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        headers = reader.fieldnames or []
        rows = list(reader)

    if not headers:
        raise ValueError("The CSV has no header row.")

    missing = sorted(REQUIRED_COLUMNS - set(headers))
    if missing:
        raise ValueError(
            "The CSV is missing required column(s): "
            + ", ".join(missing)
            + "\nRequired columns are: "
            + ", ".join(sorted(REQUIRED_COLUMNS))
        )

    if not rows:
        raise ValueError("The CSV contains no transaction rows.")

    return headers, rows


def parse_date_argument(value: Optional[str], name: str) -> Optional[date]:
    if not value:
        return None
    try:
        return date.fromisoformat(value)
    except ValueError as exc:
        raise ValueError(f"{name} must use YYYY-MM-DD format.") from exc


def determine_period(
    transactions: list[Transaction],
    days: int,
    start_date_text: Optional[str],
    end_date_text: Optional[str],
) -> tuple[datetime, datetime]:
    if not transactions:
        raise ValueError("No rows contain a valid 'created' transaction timestamp.")

    if days < 1:
        raise ValueError("--days must be at least 1.")

    explicit_start = parse_date_argument(start_date_text, "--start-date")
    explicit_end = parse_date_argument(end_date_text, "--end-date")

    latest_date = max(tx.created.date() for tx in transactions)

    if explicit_start is None and explicit_end is None:
        end_date = latest_date
        start_date = end_date - timedelta(days=days - 1)

    elif explicit_start is not None and explicit_end is not None:
        start_date = explicit_start
        end_date = explicit_end

    elif explicit_start is not None:
        start_date = explicit_start
        end_date = latest_date

    else:
        end_date = explicit_end
        start_date = end_date - timedelta(days=days - 1)

    if start_date > end_date:
        raise ValueError("--start-date cannot be after --end-date.")

    start_dt = datetime.combine(start_date, time.min, tzinfo=timezone.utc)
    end_dt_exclusive = datetime.combine(
        end_date + timedelta(days=1),
        time.min,
        tzinfo=timezone.utc,
    )
    return start_dt, end_dt_exclusive


def percentage(numerator: int, denominator: int) -> float:
    return (numerator / denominator) if denominator else 0.0


def period_label(start_dt: datetime, end_dt_exclusive: datetime) -> str:
    end_date = end_dt_exclusive.date() - timedelta(days=1)
    return f"{start_dt:%d %b %Y} - {end_date:%d %b %Y}"


def filename_period(start_dt: datetime, end_dt_exclusive: datetime) -> str:
    end_date = end_dt_exclusive.date() - timedelta(days=1)
    return f"{start_dt:%d%b}-{end_date:%d%b%Y}"


def group_by_week(
    transactions: list[Transaction],
    start_dt: datetime,
    end_dt_exclusive: datetime,
) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    cursor = start_dt
    week_no = 1

    while cursor < end_dt_exclusive:
        week_end = min(cursor + timedelta(days=7), end_dt_exclusive)
        items = [tx for tx in transactions if cursor <= tx.created < week_end]

        successful = sum(tx.status == "Successful" for tx in items)
        failed = sum(tx.status == "Failed" for tx in items)
        pending = sum(tx.status == "Pending Validation" for tx in items)

        result.append(
            {
                "week": f"Week {week_no}",
                "start": cursor,
                "end": week_end - timedelta(seconds=1),
                "attempts": len(items),
                "successful": successful,
                "failed": failed,
                "pending": pending,
                "success_rate": percentage(successful, len(items)),
            }
        )
        cursor = week_end
        week_no += 1

    return result


def build_observations(transactions: list[Transaction]) -> list[str]:
    observations: list[str] = []

    total = len(transactions)
    successful = sum(tx.status == "Successful" for tx in transactions)
    failed = sum(tx.status == "Failed" for tx in transactions)
    pending = sum(tx.status == "Pending Validation" for tx in transactions)

    observations.append(
        f"{total} payment attempt{'s' if total != 1 else ''} were recorded in the review period: "
        f"{successful} successful, {failed} failed and {pending} pending validation."
    )

    method_order = [
        "bank_transfer",
        "card",
        "applepay",
        "account-ach-uk",
    ]
    by_method: dict[str, list[Transaction]] = defaultdict(list)
    for tx in transactions:
        by_method[tx.payment_method].append(tx)

    for method in method_order:
        items = by_method.get(method, [])
        if not items:
            continue

        success_count = sum(tx.status == "Successful" for tx in items)
        failed_count = sum(tx.status == "Failed" for tx in items)
        pending_count = sum(tx.status == "Pending Validation" for tx in items)
        label = items[0].payment_method_label

        parts = [
            f"{label}: {success_count} of {len(items)} successful "
            f"({percentage(success_count, len(items)):.1%})"
        ]
        if failed_count:
            parts.append(f"{failed_count} failed")
        if pending_count:
            parts.append(f"{pending_count} pending validation")
        observations.append("; ".join(parts) + ".")

    failed_messages = Counter(
        tx.gateway_message or "No gateway message"
        for tx in transactions
        if tx.status == "Failed"
    )
    if failed_messages:
        reason_text = ", ".join(
            f"{message} ({count})"
            for message, count in failed_messages.most_common(3)
        )
        observations.append(f"Main failed-attempt messages: {reason_text}.")

    countries = sorted(
        {
            tx.source_country
            for tx in transactions
            if tx.source_country
        }
    )
    if countries:
        observations.append(
            "Payment sources/countries captured in the export include: "
            + ", ".join(countries)
            + "."
        )

    return observations


def safe_sheet_table_name(name: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_]", "", name)
    if not cleaned:
        cleaned = "ReportTable"
    if cleaned[0].isdigit():
        cleaned = "T_" + cleaned
    return cleaned[:250]


# ---------------------------------------------------------------------------
# Excel formatting helpers
# ---------------------------------------------------------------------------

def workbook_formats(workbook: xlsxwriter.Workbook) -> dict[str, Any]:
    return {
        "title": workbook.add_format(
            {
                "bold": True,
                "font_color": WHITE,
                "font_size": 18,
                "bg_color": TITLE_FILL,
                "valign": "vcenter",
            }
        ),
        "subtitle": workbook.add_format(
            {
                "italic": True,
                "font_color": GREY,
                "bg_color": PALE_BLUE,
                "text_wrap": True,
                "valign": "vcenter",
            }
        ),
        "header": workbook.add_format(
            {
                "bold": True,
                "font_color": WHITE,
                "bg_color": HEADER_FILL,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
                "border": 1,
                "border_color": WHITE,
            }
        ),
        "section": workbook.add_format(
            {
                "bold": True,
                "font_color": NAVY,
                "bg_color": SECTION_FILL,
                "valign": "vcenter",
            }
        ),
        "note": workbook.add_format(
            {
                "font_color": DARK_YELLOW,
                "bg_color": PALE_YELLOW,
                "text_wrap": True,
                "valign": "top",
                "border": 1,
                "border_color": "#E6D77A",
            }
        ),
        "body": workbook.add_format(
            {
                "valign": "top",
            }
        ),
        "body_wrap": workbook.add_format(
            {
                "text_wrap": True,
                "valign": "top",
            }
        ),
        "date": workbook.add_format(
            {
                "num_format": "dd-mmm-yyyy hh:mm",
                "valign": "top",
            }
        ),
        "money": workbook.add_format(
            {
                "num_format": "#,##0.00",
                "valign": "top",
            }
        ),
        "integer": workbook.add_format(
            {
                "num_format": "0",
                "valign": "top",
            }
        ),
        "percent": workbook.add_format(
            {
                "num_format": "0.0%",
                "valign": "top",
            }
        ),
        "success": workbook.add_format(
            {
                "font_color": DARK_GREEN,
                "bg_color": PALE_GREEN,
                "bold": True,
            }
        ),
        "failed": workbook.add_format(
            {
                "font_color": DARK_RED,
                "bg_color": PALE_RED,
                "bold": True,
            }
        ),
        "pending": workbook.add_format(
            {
                "font_color": DARK_YELLOW,
                "bg_color": PALE_YELLOW,
                "bold": True,
            }
        ),
    }


def add_title(
    worksheet: xlsxwriter.worksheet.Worksheet,
    formats: dict[str, Any],
    title: str,
    subtitle: str,
    last_col: int,
) -> None:
    worksheet.merge_range(0, 0, 0, last_col, title, formats["title"])
    worksheet.set_row(0, 30)
    worksheet.merge_range(1, 0, 1, last_col, subtitle, formats["subtitle"])
    worksheet.set_row(1, 30)


def write_table_header(
    worksheet: xlsxwriter.worksheet.Worksheet,
    row: int,
    headers: list[str],
    formats: dict[str, Any],
    start_col: int = 0,
) -> None:
    for offset, header in enumerate(headers):
        worksheet.write(row, start_col + offset, header, formats["header"])


def status_format(formats: dict[str, Any], status: str) -> Any:
    if status == "Successful":
        return formats["success"]
    if status == "Failed":
        return formats["failed"]
    if status == "Pending Validation":
        return formats["pending"]
    return formats["body"]


def add_excel_table(
    worksheet: xlsxwriter.worksheet.Worksheet,
    first_row: int,
    first_col: int,
    last_row: int,
    last_col: int,
    headers: list[str],
    name: str,
) -> None:
    if last_row <= first_row:
        return

    worksheet.add_table(
        first_row,
        first_col,
        last_row,
        last_col,
        {
            "name": safe_sheet_table_name(name),
            "style": "Table Style Medium 2",
            "columns": [{"header": header} for header in headers],
        },
    )


# ---------------------------------------------------------------------------
# Sheet writers
# ---------------------------------------------------------------------------

def write_transactions_sheet(
    workbook: xlsxwriter.Workbook,
    formats: dict[str, Any],
    transactions: list[Transaction],
) -> None:
    ws = workbook.add_worksheet("Test Transactions")

    headers = [
        "Date/Time (UTC)",
        "Transaction Reference",
        "Transaction ID",
        "Amount",
        "Currency",
        "Payment Method",
        "Source / Issuer",
        "Source Country",
        "Instrument",
        "Status",
        "Gateway Message",
        "Payer / Tester",
        "Email",
        "Originator / Account Name",
        "Masked Source Account",
        "Settlement Status",
        "Settlement Amount",
    ]

    write_table_header(ws, 0, headers, formats)

    for row_no, tx in enumerate(transactions, start=1):
        values = [
            tx.created.replace(tzinfo=None),
            tx.txref,
            tx.txid,
            tx.amount,
            tx.currency,
            tx.payment_method_label,
            tx.source,
            tx.source_country,
            tx.instrument,
            tx.status,
            tx.gateway_message,
            tx.customer_name,
            tx.customer_email,
            tx.originator_name,
            tx.source_account,
            tx.settlement_status,
            tx.settlement_amount,
        ]

        for col_no, value in enumerate(values):
            if col_no == 0:
                ws.write_datetime(row_no, col_no, value, formats["date"])
            elif col_no in {3, 16} and isinstance(value, (int, float)):
                ws.write_number(row_no, col_no, value, formats["money"])
            elif col_no == 9:
                ws.write(row_no, col_no, value, status_format(formats, tx.status))
            elif col_no in {1, 6, 10, 11, 12, 13}:
                ws.write(row_no, col_no, value, formats["body_wrap"])
            else:
                ws.write(row_no, col_no, value, formats["body"])

    add_excel_table(
        ws,
        0,
        0,
        max(len(transactions), 1),
        len(headers) - 1,
        headers,
        "TestTransactionsTable",
    )

    widths = [19, 42, 16, 13, 10, 23, 42, 22, 17, 20, 48, 25, 31, 30, 24, 19, 18]
    for col, width in enumerate(widths):
        ws.set_column(col, col, width)

    ws.freeze_panes(1, 0)


def write_management_summary(
    workbook: xlsxwriter.Workbook,
    formats: dict[str, Any],
    transactions: list[Transaction],
    start_dt: datetime,
    end_dt_exclusive: datetime,
    report_title: str,
) -> None:
    ws = workbook.add_worksheet("Management Summary")

    label = period_label(start_dt, end_dt_exclusive)
    add_title(
        ws,
        formats,
        report_title,
        f"Management review of payments, failed payment attempts and source of payments | {label}",
        7,
    )

    total = len(transactions)
    successful = sum(tx.status == "Successful" for tx in transactions)
    failed = sum(tx.status == "Failed" for tx in transactions)
    pending = sum(tx.status == "Pending Validation" for tx in transactions)
    completed = total - pending

    ws.write_row(3, 0, ["KEY METRIC", "RESULT"], formats["header"])
    metrics = [
        ("Total payment attempts", total, formats["integer"]),
        ("Successful payments", successful, formats["integer"]),
        ("Failed attempts", failed, formats["integer"]),
        ("Pending validation", pending, formats["integer"]),
        ("Completed-attempt success rate", percentage(successful, completed), formats["percent"]),
        ("Overall success rate", percentage(successful, total), formats["percent"]),
        ("Review period", label, formats["body"]),
    ]
    for row_no, (metric, value, value_format) in enumerate(metrics, start=4):
        ws.write(row_no, 0, metric, formats["body"])
        if isinstance(value, float):
            ws.write_number(row_no, 1, value, value_format)
        elif isinstance(value, int):
            ws.write_number(row_no, 1, value, value_format)
        else:
            ws.write(row_no, 1, value, value_format)

    ws.merge_range(3, 3, 3, 7, "MANAGEMENT OBSERVATIONS", formats["section"])
    observations = build_observations(transactions)
    observation_text = "\n".join(f"• {item}" for item in observations)
    ws.merge_range(4, 3, 10, 7, observation_text, formats["note"])

    # Currency summary
    currency_headers = [
        "Currency",
        "Attempts",
        "Successful",
        "Failed",
        "Pending",
        "Successful Value",
        "Failed Value",
        "Pending Value",
    ]
    write_table_header(ws, 13, currency_headers, formats)

    currencies = sorted({tx.currency for tx in transactions})
    currency_start_row = 14

    for offset, currency in enumerate(currencies):
        row_no = currency_start_row + offset
        items = [tx for tx in transactions if tx.currency == currency]

        attempts = len(items)
        success_items = [tx for tx in items if tx.status == "Successful"]
        failed_items = [tx for tx in items if tx.status == "Failed"]
        pending_items = [tx for tx in items if tx.status == "Pending Validation"]

        values = [
            currency,
            attempts,
            len(success_items),
            len(failed_items),
            len(pending_items),
            sum(tx.amount for tx in success_items),
            sum(tx.amount for tx in failed_items),
            sum(tx.amount for tx in pending_items),
        ]

        for col_no, value in enumerate(values):
            if col_no >= 5:
                ws.write_number(row_no, col_no, value, formats["money"])
            elif isinstance(value, int):
                ws.write_number(row_no, col_no, value, formats["integer"])
            else:
                ws.write(row_no, col_no, value, formats["body"])

    method_start_row = currency_start_row + len(currencies) + 2
    method_headers = [
        "Payment Method",
        "Attempts",
        "Successful",
        "Failed",
        "Pending",
        "Overall Success Rate",
    ]
    write_table_header(ws, method_start_row, method_headers, formats)

    method_order = ["bank_transfer", "card", "applepay", "account-ach-uk"]
    present_methods = {tx.payment_method for tx in transactions}

    output_methods = [m for m in method_order if m in present_methods]
    output_methods.extend(sorted(present_methods - set(output_methods)))

    for offset, method in enumerate(output_methods, start=1):
        row_no = method_start_row + offset
        items = [tx for tx in transactions if tx.payment_method == method]
        success_count = sum(tx.status == "Successful" for tx in items)
        failed_count = sum(tx.status == "Failed" for tx in items)
        pending_count = sum(tx.status == "Pending Validation" for tx in items)
        label_text = items[0].payment_method_label if items else METHOD_LABELS.get(method, method)

        values = [
            label_text,
            len(items),
            success_count,
            failed_count,
            pending_count,
            percentage(success_count, len(items)),
        ]
        for col_no, value in enumerate(values):
            if col_no == 5:
                ws.write_number(row_no, col_no, value, formats["percent"])
            elif isinstance(value, int):
                ws.write_number(row_no, col_no, value, formats["integer"])
            else:
                ws.write(row_no, col_no, value, formats["body"])

    # Status helper data for pie chart.
    chart_start_row = method_start_row + len(output_methods) + 3
    write_table_header(ws, chart_start_row, ["Status", "Attempts"], formats)
    chart_data = [
        ("Successful", successful),
        ("Failed", failed),
        ("Pending Validation", pending),
    ]
    for offset, (status, count) in enumerate(chart_data, start=1):
        ws.write(chart_start_row + offset, 0, status, formats["body"])
        ws.write_number(chart_start_row + offset, 1, count, formats["integer"])

    # Pie chart
    pie = workbook.add_chart({"type": "pie"})
    pie.add_series(
        {
            "name": "Payment Attempts by Status",
            "categories": ["Management Summary", chart_start_row + 1, 0, chart_start_row + 3, 0],
            "values": ["Management Summary", chart_start_row + 1, 1, chart_start_row + 3, 1],
            "data_labels": {"percentage": True},
        }
    )
    pie.set_title({"name": "Payment Attempts by Status"})
    pie.set_legend({"position": "bottom"})
    pie.set_style(10)
    ws.insert_chart("J4", pie, {"x_scale": 1.05, "y_scale": 1.05})

    # Column chart by payment method
    method_chart = workbook.add_chart({"type": "column"})
    if output_methods:
        method_chart.add_series(
            {
                "name": "Attempts",
                "categories": [
                    "Management Summary",
                    method_start_row + 1,
                    0,
                    method_start_row + len(output_methods),
                    0,
                ],
                "values": [
                    "Management Summary",
                    method_start_row + 1,
                    1,
                    method_start_row + len(output_methods),
                    1,
                ],
                "data_labels": {"value": True},
            }
        )
    method_chart.set_title({"name": "Attempts by Payment Method"})
    method_chart.set_legend({"none": True})
    method_chart.set_y_axis({"major_gridlines": {"visible": False}})
    method_chart.set_style(10)
    ws.insert_chart("J20", method_chart, {"x_scale": 1.05, "y_scale": 1.05})

    widths = {
        0: 29,
        1: 19,
        2: 14,
        3: 16,
        4: 14,
        5: 20,
        6: 18,
        7: 18,
    }
    for col, width in widths.items():
        ws.set_column(col, col, width)

    ws.set_column(9, 16, 12)
    ws.freeze_panes(2, 0)


def write_weekly_trend(
    workbook: xlsxwriter.Workbook,
    formats: dict[str, Any],
    transactions: list[Transaction],
    start_dt: datetime,
    end_dt_exclusive: datetime,
) -> None:
    ws = workbook.add_worksheet("Weekly Trend")
    add_title(
        ws,
        formats,
        "THREE-WEEK PAYMENT TREND",
        "Weekly tracking of successful, failed and pending payment attempts within the management review period.",
        6,
    )

    headers = [
        "Week",
        "Date Range",
        "Attempts",
        "Successful",
        "Failed",
        "Pending",
        "Overall Success Rate",
    ]
    write_table_header(ws, 3, headers, formats)

    weeks = group_by_week(transactions, start_dt, end_dt_exclusive)
    for offset, week in enumerate(weeks, start=4):
        date_range = f"{week['start']:%d %b %Y} - {week['end']:%d %b %Y}"
        values = [
            week["week"],
            date_range,
            week["attempts"],
            week["successful"],
            week["failed"],
            week["pending"],
            week["success_rate"],
        ]
        for col_no, value in enumerate(values):
            if col_no == 6:
                ws.write_number(offset, col_no, value, formats["percent"])
            elif isinstance(value, int):
                ws.write_number(offset, col_no, value, formats["integer"])
            else:
                ws.write(offset, col_no, value, formats["body"])

    # Chart uses Attempts/Successful/Failed/Pending, not success rate.
    chart = workbook.add_chart({"type": "column"})
    if weeks:
        for series_col, name in [(3, "Successful"), (4, "Failed"), (5, "Pending")]:
            chart.add_series(
                {
                    "name": name,
                    "categories": ["Weekly Trend", 4, 0, 3 + len(weeks), 0],
                    "values": ["Weekly Trend", 4, series_col, 3 + len(weeks), series_col],
                    "data_labels": {"value": True},
                }
            )
    chart.set_title({"name": "Weekly Attempts by Status"})
    chart.set_legend({"position": "bottom"})
    chart.set_style(10)
    ws.insert_chart("I4", chart, {"x_scale": 1.15, "y_scale": 1.15})

    widths = [15, 28, 14, 14, 14, 14, 21]
    for col, width in enumerate(widths):
        ws.set_column(col, col, width)
    ws.freeze_panes(4, 0)


def write_sources_sheet(
    workbook: xlsxwriter.Workbook,
    formats: dict[str, Any],
    transactions: list[Transaction],
) -> None:
    ws = workbook.add_worksheet("Payment Sources")
    add_title(
        ws,
        formats,
        "PAYMENT SOURCE ANALYSIS",
        "Source is taken from the card issuer, originating bank or payment channel available in the gateway export.",
        8,
    )

    headers = [
        "Payment Method",
        "Source / Issuer",
        "Country",
        "Instrument",
        "Attempts",
        "Successful",
        "Failed",
        "Pending",
        "Success Rate",
    ]
    write_table_header(ws, 3, headers, formats)

    groups: dict[tuple[str, str, str, str], list[Transaction]] = defaultdict(list)
    for tx in transactions:
        key = (
            tx.payment_method_label,
            tx.source,
            tx.source_country,
            tx.instrument,
        )
        groups[key].append(tx)

    source_rows = []
    for key in sorted(groups, key=lambda x: (x[0], x[1], x[2], x[3])):
        items = groups[key]
        success_count = sum(tx.status == "Successful" for tx in items)
        failed_count = sum(tx.status == "Failed" for tx in items)
        pending_count = sum(tx.status == "Pending Validation" for tx in items)

        source_rows.append(
            [
                *key,
                len(items),
                success_count,
                failed_count,
                pending_count,
                percentage(success_count, len(items)),
            ]
        )

    for row_no, row in enumerate(source_rows, start=4):
        for col_no, value in enumerate(row):
            if col_no == 8:
                ws.write_number(row_no, col_no, value, formats["percent"])
            elif isinstance(value, int):
                ws.write_number(row_no, col_no, value, formats["integer"])
            else:
                ws.write(row_no, col_no, value, formats["body_wrap"] if col_no == 1 else formats["body"])

    if source_rows:
        add_excel_table(
            ws,
            3,
            0,
            3 + len(source_rows),
            len(headers) - 1,
            headers,
            "PaymentSourcesTable",
        )

    widths = [23, 46, 23, 18, 13, 13, 13, 13, 14]
    for col, width in enumerate(widths):
        ws.set_column(col, col, width)
    ws.freeze_panes(4, 0)


def write_failures_sheet(
    workbook: xlsxwriter.Workbook,
    formats: dict[str, Any],
    transactions: list[Transaction],
) -> None:
    ws = workbook.add_worksheet("Failed & Pending")
    add_title(
        ws,
        formats,
        "FAILED & PENDING PAYMENT ATTEMPTS",
        "Detailed exceptions requiring review, including gateway messages and payment source.",
        10,
    )

    exceptions = [
        tx for tx in transactions if tx.status != "Successful"
    ]

    headers = [
        "Date/Time (UTC)",
        "Transaction Reference",
        "Amount",
        "Currency",
        "Payment Method",
        "Source / Issuer",
        "Country",
        "Status",
        "Gateway Message",
        "Payer / Tester",
        "Email",
    ]
    write_table_header(ws, 3, headers, formats)

    for row_no, tx in enumerate(exceptions, start=4):
        values = [
            tx.created.replace(tzinfo=None),
            tx.txref,
            tx.amount,
            tx.currency,
            tx.payment_method_label,
            tx.source,
            tx.source_country,
            tx.status,
            tx.gateway_message,
            tx.customer_name,
            tx.customer_email,
        ]

        for col_no, value in enumerate(values):
            if col_no == 0:
                ws.write_datetime(row_no, col_no, value, formats["date"])
            elif col_no == 2:
                ws.write_number(row_no, col_no, value, formats["money"])
            elif col_no == 7:
                ws.write(row_no, col_no, value, status_format(formats, tx.status))
            elif col_no in {1, 5, 8, 9, 10}:
                ws.write(row_no, col_no, value, formats["body_wrap"])
            else:
                ws.write(row_no, col_no, value, formats["body"])

    if exceptions:
        add_excel_table(
            ws,
            3,
            0,
            3 + len(exceptions),
            len(headers) - 1,
            headers,
            "FailedPendingTable",
        )

    # Failure reason summary.
    reason_counts = Counter(
        tx.gateway_message or "No gateway message"
        for tx in exceptions
    )
    write_table_header(ws, 3, ["Failure / Pending Reason", "Count"], formats, start_col=12)
    for offset, (message, count) in enumerate(reason_counts.most_common(), start=4):
        ws.write(offset, 12, message, formats["body_wrap"])
        ws.write_number(offset, 13, count, formats["integer"])

    widths = [19, 42, 13, 10, 23, 43, 22, 20, 50, 25, 31]
    for col, width in enumerate(widths):
        ws.set_column(col, col, width)
    ws.set_column(12, 12, 52)
    ws.set_column(13, 13, 12)
    ws.freeze_panes(4, 0)


def write_historical_sheet(
    workbook: xlsxwriter.Workbook,
    formats: dict[str, Any],
    all_transactions: list[Transaction],
    report_transactions: list[Transaction],
    start_dt: datetime,
    end_dt_exclusive: datetime,
) -> None:
    ws = workbook.add_worksheet("Historical Overview")
    add_title(
        ws,
        formats,
        "FULL EXPORT - HISTORICAL OVERVIEW",
        "The uploaded export may contain older activity outside the management review period. This sheet keeps that broader context separate.",
        10,
    )

    earliest = min(tx.created for tx in all_transactions)
    latest = max(tx.created for tx in all_transactions)

    write_table_header(ws, 3, ["Full Export Metric", "Value"], formats)

    metrics = [
        ("Total valid transactions", len(all_transactions)),
        ("Earliest transaction", earliest.replace(tzinfo=None)),
        ("Latest transaction", latest.replace(tzinfo=None)),
        ("Transactions in review period", len(report_transactions)),
        ("Transactions outside review period", len(all_transactions) - len(report_transactions)),
        ("Review period", period_label(start_dt, end_dt_exclusive)),
    ]

    for row_no, (metric, value) in enumerate(metrics, start=4):
        ws.write(row_no, 0, metric, formats["body"])
        if isinstance(value, datetime):
            ws.write_datetime(row_no, 1, value, formats["date"])
        elif isinstance(value, int):
            ws.write_number(row_no, 1, value, formats["integer"])
        else:
            ws.write(row_no, 1, value, formats["body"])

    status_counts = Counter(tx.status for tx in all_transactions)
    method_counts = Counter(tx.payment_method_label for tx in all_transactions)
    currency_counts = Counter(tx.currency for tx in all_transactions)

    write_table_header(ws, 3, ["Status", "Transactions"], formats, start_col=3)
    for row_no, (status, count) in enumerate(sorted(status_counts.items()), start=4):
        ws.write(row_no, 3, status, formats["body"])
        ws.write_number(row_no, 4, count, formats["integer"])

    write_table_header(ws, 3, ["Payment Method", "Transactions"], formats, start_col=6)
    for row_no, (method, count) in enumerate(sorted(method_counts.items()), start=4):
        ws.write(row_no, 6, method, formats["body"])
        ws.write_number(row_no, 7, count, formats["integer"])

    write_table_header(ws, 3, ["Currency", "Transactions"], formats, start_col=9)
    for row_no, (currency, count) in enumerate(sorted(currency_counts.items()), start=4):
        ws.write(row_no, 9, currency, formats["body"])
        ws.write_number(row_no, 10, count, formats["integer"])

    ws.set_column(0, 0, 31)
    ws.set_column(1, 1, 25)
    ws.set_column(3, 3, 28)
    ws.set_column(4, 4, 14)
    ws.set_column(6, 6, 26)
    ws.set_column(7, 7, 14)
    ws.set_column(9, 9, 14)
    ws.set_column(10, 10, 14)


def write_raw_sheet(
    workbook: xlsxwriter.Workbook,
    formats: dict[str, Any],
    headers: list[str],
    raw_rows: list[dict[str, str]],
) -> None:
    ws = workbook.add_worksheet("Raw Data")
    write_table_header(ws, 0, headers, formats)

    for row_no, row in enumerate(raw_rows, start=1):
        for col_no, header in enumerate(headers):
            value = row.get(header, "")

            if header == "created":
                dt = parse_datetime(value)
                if dt:
                    ws.write_datetime(row_no, col_no, dt.replace(tzinfo=None), formats["date"])
                    continue

            if header in NUMERIC_COLUMNS:
                numeric = safe_float(value)
                if numeric is not None:
                    ws.write_number(row_no, col_no, numeric, formats["money"])
                    continue

            ws.write(row_no, col_no, clean(value), formats["body"])

    if raw_rows:
        add_excel_table(
            ws,
            0,
            0,
            len(raw_rows),
            len(headers) - 1,
            headers,
            "RawGatewayData",
        )

    ws.freeze_panes(1, 0)
    ws.set_column(0, max(len(headers) - 1, 0), 15)

    # Widen commonly useful text columns.
    for header_name, width in {
        "txref": 42,
        "chargemessage": 50,
        "authurl": 45,
        "pcardname": 45,
        "custname": 25,
        "custemail": 31,
        "originatorname": 30,
        "originatoraccountnumber": 24,
    }.items():
        if header_name in headers:
            col = headers.index(header_name)
            ws.set_column(col, col, width)


# ---------------------------------------------------------------------------
# Report generation
# ---------------------------------------------------------------------------

def generate_report(
    input_csv: Path,
    output_xlsx: Path,
    days: int = 21,
    start_date_text: Optional[str] = None,
    end_date_text: Optional[str] = None,
    report_title: str = "PAYMENT TEST MONITORING REPORT",
) -> dict[str, Any]:
    headers, raw_rows = read_csv(input_csv)

    all_transactions = [
        tx
        for row in raw_rows
        if (tx := normalize_transaction(row)) is not None
    ]

    if not all_transactions:
        raise ValueError(
            "No valid transaction rows were found. Check that the 'created' column contains valid timestamps."
        )

    start_dt, end_dt_exclusive = determine_period(
        all_transactions,
        days,
        start_date_text,
        end_date_text,
    )

    report_transactions = sorted(
        [
            tx
            for tx in all_transactions
            if start_dt <= tx.created < end_dt_exclusive
        ],
        key=lambda tx: tx.created,
    )

    if not report_transactions:
        raise ValueError(
            "No transactions fall within the selected reporting period: "
            + period_label(start_dt, end_dt_exclusive)
        )

    output_xlsx.parent.mkdir(parents=True, exist_ok=True)

    workbook = xlsxwriter.Workbook(str(output_xlsx))
    workbook.set_properties(
        {
            "title": report_title.title(),
            "subject": "Payment test management report",
            "author": "Payment Report Generator",
            "comments": f"Generated from {input_csv.name}",
        }
    )
    formats = workbook_formats(workbook)

    write_management_summary(
        workbook,
        formats,
        report_transactions,
        start_dt,
        end_dt_exclusive,
        report_title,
    )
    write_weekly_trend(
        workbook,
        formats,
        report_transactions,
        start_dt,
        end_dt_exclusive,
    )
    write_sources_sheet(
        workbook,
        formats,
        report_transactions,
    )
    write_failures_sheet(
        workbook,
        formats,
        report_transactions,
    )
    write_transactions_sheet(
        workbook,
        formats,
        report_transactions,
    )
    write_historical_sheet(
        workbook,
        formats,
        all_transactions,
        report_transactions,
        start_dt,
        end_dt_exclusive,
    )
    write_raw_sheet(
        workbook,
        formats,
        headers,
        raw_rows,
    )

    workbook.close()

    status_counts = Counter(tx.status for tx in report_transactions)
    return {
        "output": output_xlsx,
        "period": period_label(start_dt, end_dt_exclusive),
        "attempts": len(report_transactions),
        "successful": status_counts["Successful"],
        "failed": status_counts["Failed"],
        "pending": status_counts["Pending Validation"],
    }


# ---------------------------------------------------------------------------
# Command-line interface
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Generate a management-ready Excel report from a Flutterwave-style transaction CSV."
        )
    )
    parser.add_argument(
        "input_csv",
        help="Path to the input transaction CSV file.",
    )
    parser.add_argument(
        "--output",
        help=(
            "Output .xlsx file path. If omitted, the file is created beside the input CSV."
        ),
    )
    parser.add_argument(
        "--days",
        type=int,
        default=21,
        help=(
            "Number of calendar days in the automatic reporting window. "
            "Default: 21. Ignored when both --start-date and --end-date are supplied."
        ),
    )
    parser.add_argument(
        "--start-date",
        help="Optional reporting start date in YYYY-MM-DD format.",
    )
    parser.add_argument(
        "--end-date",
        help="Optional reporting end date in YYYY-MM-DD format, inclusive.",
    )
    parser.add_argument(
        "--title",
        default="PAYMENT TEST MONITORING REPORT",
        help="Optional title used on the Management Summary sheet.",
    )
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    input_path = Path(args.input_csv).expanduser().resolve()

    # Determine the period once so we can also create a useful default filename.
    headers, raw_rows = read_csv(input_path)
    all_transactions = [
        tx
        for row in raw_rows
        if (tx := normalize_transaction(row)) is not None
    ]
    if not all_transactions:
        raise ValueError("No valid transaction timestamps were found in the CSV.")

    start_dt, end_dt_exclusive = determine_period(
        all_transactions,
        args.days,
        args.start_date,
        args.end_date,
    )

    if args.output:
        output_path = Path(args.output).expanduser().resolve()
    else:
        output_name = f"Management_Payment_Test_Report_{filename_period(start_dt, end_dt_exclusive)}.xlsx"
        output_path = input_path.parent / output_name

    result = generate_report(
        input_csv=input_path,
        output_xlsx=output_path,
        days=args.days,
        start_date_text=args.start_date,
        end_date_text=args.end_date,
        report_title=args.title,
    )

    print("")
    print("Report generated successfully.")
    print(f"Period:      {result['period']}")
    print(f"Attempts:    {result['attempts']}")
    print(f"Successful:  {result['successful']}")
    print(f"Failed:      {result['failed']}")
    print(f"Pending:     {result['pending']}")
    print(f"Output:      {result['output']}")


if __name__ == "__main__":
    main()
