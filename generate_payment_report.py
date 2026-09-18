import csv
import io
import math
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta, timezone
from typing import Any, Optional

import streamlit as st
import xlsxwriter


# ============================================================
# STREAMLIT PAGE
# ============================================================

st.set_page_config(
    page_title="Payment Report Generator",
    page_icon="📊",
    layout="wide",
)


# ============================================================
# CONFIGURATION
# ============================================================

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


# ============================================================
# DATA MODEL
# ============================================================

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


# ============================================================
# HELPERS
# ============================================================

def clean(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def safe_float(value: Any) -> Optional[float]:
    value = clean(value).replace(",", "")

    if not value:
        return None

    try:
        number = float(value)

        if math.isnan(number) or math.isinf(number):
            return None

        return number

    except (TypeError, ValueError):
        return None


def parse_datetime(value: Any) -> Optional[datetime]:
    value = clean(value)

    if not value:
        return None

    candidates = [
        value,
        value.replace("Z", "+00:00"),
    ]

    for candidate in candidates:
        try:
            parsed = datetime.fromisoformat(candidate)

            if parsed.tzinfo is None:
                parsed = parsed.replace(tzinfo=timezone.utc)

            return parsed.astimezone(timezone.utc)

        except ValueError:
            pass

    fallback_formats = (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M:%S.%f",
        "%d/%m/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M:%S",
    )

    for fmt in fallback_formats:
        try:
            return datetime.strptime(value, fmt).replace(tzinfo=timezone.utc)
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

    label = METHOD_LABELS.get(
        method,
        method.replace("_", " ").title() if method else "Other",
    )

    return method, label


def mask_account(value: str) -> str:
    value = clean(value)

    if not value:
        return ""

    if "*" in value:
        return value

    digits = re.sub(r"\D", "", value)

    if len(digits) <= 6:
        return value

    return f"{digits[:3]}{'*' * (len(digits) - 6)}{digits[-3:]}"


def choose_first(row: dict[str, str], *columns: str) -> str:
    for column in columns:
        value = clean(row.get(column))

        if value:
            return value

    return ""


def read_csv_bytes(file_bytes: bytes) -> tuple[list[str], list[dict[str, str]]]:
    text = file_bytes.decode("utf-8-sig", errors="replace")

    reader = csv.DictReader(io.StringIO(text))

    headers = reader.fieldnames or []

    rows = list(reader)

    if not headers:
        raise ValueError("The uploaded CSV has no header row.")

    missing = sorted(REQUIRED_COLUMNS - set(headers))

    if missing:
        raise ValueError(
            "The CSV is missing required column(s): "
            + ", ".join(missing)
        )

    if not rows:
        raise ValueError("The CSV contains no transaction rows.")

    return headers, rows


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

        source_country = clean(row.get("pcardcountry"))

        instrument = choose_first(row, "pcardtype") or "Card"

    elif method == "bank_transfer":

        source = (
            choose_first(row, "bankname", "paccountbankname")
            or "Bank Transfer"
        )

        source_country = "Nigeria" if currency == "NGN" else ""

        instrument = "Bank Transfer"

        originator_name = choose_first(row, "originatorname")

        if not originator_name:
            first_name = clean(row.get("paccountfirstname"))
            last_name = clean(row.get("paccountlastname"))

            originator_name = " ".join(
                x for x in [first_name, last_name] if x
            )

        source_account = mask_account(
            choose_first(
                row,
                "originatoraccountnumber",
                "paccountnumber",
            )
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


def percentage(numerator: int, denominator: int) -> float:
    if denominator == 0:
        return 0.0

    return numerator / denominator


def period_label(start_dt: datetime, end_dt_exclusive: datetime) -> str:
    end_date = end_dt_exclusive.date() - timedelta(days=1)

    return f"{start_dt:%d %b %Y} - {end_date:%d %b %Y}"


def filename_period(start_dt: datetime, end_dt_exclusive: datetime) -> str:
    end_date = end_dt_exclusive.date() - timedelta(days=1)

    return f"{start_dt:%d%b}-{end_date:%d%b%Y}"


def to_utc_start(day: date) -> datetime:
    return datetime.combine(
        day,
        time.min,
        tzinfo=timezone.utc,
    )


def to_utc_end_exclusive(day: date) -> datetime:
    return datetime.combine(
        day + timedelta(days=1),
        time.min,
        tzinfo=timezone.utc,
    )


def group_by_week(
    transactions: list[Transaction],
    start_dt: datetime,
    end_dt_exclusive: datetime,
) -> list[dict[str, Any]]:

    result = []

    cursor = start_dt

    week_number = 1

    while cursor < end_dt_exclusive:

        week_end = min(
            cursor + timedelta(days=7),
            end_dt_exclusive,
        )

        items = [
            tx
            for tx in transactions
            if cursor <= tx.created < week_end
        ]

        successful = sum(
            tx.status == "Successful"
            for tx in items
        )

        failed = sum(
            tx.status == "Failed"
            for tx in items
        )

        pending = sum(
            tx.status == "Pending Validation"
            for tx in items
        )

        result.append(
            {
                "week": f"Week {week_number}",
                "start": cursor,
                "end": week_end - timedelta(seconds=1),
                "attempts": len(items),
                "successful": successful,
                "failed": failed,
                "pending": pending,
                "success_rate": percentage(
                    successful,
                    len(items),
                ),
            }
        )

        cursor = week_end

        week_number += 1

    return result


def build_observations(
    transactions: list[Transaction],
) -> list[str]:

    observations = []

    total = len(transactions)

    successful = sum(
        tx.status == "Successful"
        for tx in transactions
    )

    failed = sum(
        tx.status == "Failed"
        for tx in transactions
    )

    pending = sum(
        tx.status == "Pending Validation"
        for tx in transactions
    )

    observations.append(
        f"{total} payment attempts were recorded in the review period: "
        f"{successful} successful, {failed} failed and "
        f"{pending} pending validation."
    )

    groups = defaultdict(list)

    for tx in transactions:
        groups[tx.payment_method].append(tx)

    priority = [
        "bank_transfer",
        "card",
        "applepay",
        "account-ach-uk",
    ]

    methods = [
        method
        for method in priority
        if method in groups
    ]

    methods.extend(
        sorted(
            set(groups.keys()) - set(methods)
        )
    )

    for method in methods:

        items = groups[method]

        success_count = sum(
            tx.status == "Successful"
            for tx in items
        )

        failed_count = sum(
            tx.status == "Failed"
            for tx in items
        )

        pending_count = sum(
            tx.status == "Pending Validation"
            for tx in items
        )

        label = items[0].payment_method_label

        sentence = (
            f"{label}: {success_count} of {len(items)} successful "
            f"({percentage(success_count, len(items)):.1%})"
        )

        extras = []

        if failed_count:
            extras.append(f"{failed_count} failed")

        if pending_count:
            extras.append(f"{pending_count} pending validation")

        if extras:
            sentence += "; " + ", ".join(extras)

        observations.append(sentence + ".")

    failed_messages = Counter(
        tx.gateway_message or "No gateway message"
        for tx in transactions
        if tx.status == "Failed"
    )

    if failed_messages:

        reason_text = ", ".join(
            f"{message} ({count})"
            for message, count
            in failed_messages.most_common(3)
        )

        observations.append(
            f"Main failed-attempt messages: {reason_text}."
        )

    countries = sorted(
        {
            tx.source_country
            for tx in transactions
            if tx.source_country
        }
    )

    if countries:

        observations.append(
            "Payment source countries captured in the export include: "
            + ", ".join(countries)
            + "."
        )

    return observations



# ============================================================
# DONOR / FUNDER ANALYSIS
# ============================================================

def _identity_text(value: str) -> str:
    value = clean(value).lower()
    return re.sub(r"\s+", " ", value).strip()


def donor_identity(tx: Transaction) -> tuple[str, str, str]:
    """Return a conservative donor key, display name and identity basis."""
    originator = _identity_text(tx.originator_name)
    email = _identity_text(tx.customer_email)
    customer = _identity_text(tx.customer_name)
    account = _identity_text(tx.source_account)

    display = (
        clean(tx.originator_name)
        or clean(tx.customer_name)
        or clean(tx.customer_email)
        or clean(tx.source_account)
        or 'Unidentified Donor'
    )

    # For bank transfers, analyse the person/account that actually sent the funds.
    if tx.payment_method == 'bank_transfer':
        if originator:
            return f'ORIGINATOR::{originator}', display, 'Bank transfer originator name'
        if account:
            return f'ACCOUNT::{account}', display, 'Bank transfer source account'

    # For card/other channels, email is usually the most stable customer identifier.
    if email:
        return f'EMAIL::{email}', display, 'Customer email'
    if customer:
        return f'CUSTOMER::{customer}', display, 'Customer name'
    if account:
        return f'ACCOUNT::{account}', display, 'Source account'

    fallback = clean(tx.txref) or clean(tx.txid) or tx.created.isoformat()
    return f'UNIDENTIFIED::{fallback}', display, 'Unidentified / transaction specific'


def funding_frequency_label(successful_count: int) -> str:
    """Frequency is based on successful funding events only."""
    if successful_count >= 10:
        return 'Very Frequent'
    if successful_count >= 4:
        return 'Frequent'
    if successful_count >= 2:
        return 'Repeat'
    if successful_count == 1:
        return 'One-Time'
    return 'No Successful Funding'


def build_donor_analysis(transactions: list[Transaction]) -> list[dict[str, Any]]:
    """
    Build one row per donor + currency.

    Total Amount Funded = successful transaction value only.
    Failed and pending values are reported separately.
    Currencies are deliberately not combined.
    """
    groups = defaultdict(list)
    metadata = {}

    for tx in transactions:
        donor_key, donor_name, basis = donor_identity(tx)
        currency = tx.currency or 'UNKNOWN'
        groups[(donor_key, currency)].append(tx)

        meta = metadata.setdefault(
            donor_key,
            {
                'name': donor_name,
                'basis': basis,
                'emails': set(),
                'originators': set(),
                'accounts': set(),
            },
        )
        if tx.customer_email:
            meta['emails'].add(clean(tx.customer_email))
        if tx.originator_name:
            meta['originators'].add(clean(tx.originator_name))
        if tx.source_account:
            meta['accounts'].add(clean(tx.source_account))

    rows = []
    for (donor_key, currency), items in groups.items():
        successful_items = [tx for tx in items if tx.status == 'Successful']
        failed_items = [tx for tx in items if tx.status == 'Failed']
        pending_items = [tx for tx in items if tx.status == 'Pending Validation']
        other_items = [
            tx for tx in items
            if tx.status not in {'Successful', 'Failed', 'Pending Validation'}
        ]
        meta = metadata[donor_key]
        success_count = len(successful_items)

        rows.append({
            'Donor / Funder': meta['name'],
            'Currency': currency,
            'Funding Frequency': funding_frequency_label(success_count),
            'Total Attempts': len(items),
            'Successful Funding Count': success_count,
            'Failed / Unsuccessful Count': len(failed_items),
            'Pending Validation Count': len(pending_items),
            'Other Status Count': len(other_items),
            'Total Amount Funded': sum(tx.amount for tx in successful_items),
            'Failed Attempt Value': sum(tx.amount for tx in failed_items),
            'Pending Attempt Value': sum(tx.amount for tx in pending_items),
            'Total Attempted Value': sum(tx.amount for tx in items),
            'Successful Funding Rate': percentage(success_count, len(items)),
            'First Funding Attempt': min(tx.created for tx in items),
            'Last Funding Attempt': max(tx.created for tx in items),
            'Payment Method(s)': ' | '.join(sorted({tx.payment_method_label for tx in items if tx.payment_method_label})),
            'Source / Issuer(s)': ' | '.join(sorted({tx.source for tx in items if tx.source})),
            'Originator Name(s)': ' | '.join(sorted(meta['originators'])),
            'Email(s)': ' | '.join(sorted(meta['emails'])),
            'Masked Source Account(s)': ' | '.join(sorted(meta['accounts'])),
            'Identity Basis': meta['basis'],
        })

    rows.sort(key=lambda r: (r['Currency'], -r['Total Amount Funded'], -r['Successful Funding Count'], r['Donor / Funder'].lower()))
    return rows


def donor_summary_metrics(transactions: list[Transaction]) -> dict[str, int]:
    grouped = defaultdict(list)
    for tx in transactions:
        donor_key, _, _ = donor_identity(tx)
        grouped[donor_key].append(tx)

    return {
        'unique_donors': len(grouped),
        'successful_donors': sum(any(tx.status == 'Successful' for tx in items) for items in grouped.values()),
        'repeat_donors': sum(sum(tx.status == 'Successful' for tx in items) >= 2 for items in grouped.values()),
        'donors_with_failures': sum(any(tx.status == 'Failed' for tx in items) for items in grouped.values()),
        'donors_with_pending': sum(any(tx.status == 'Pending Validation' for tx in items) for items in grouped.values()),
    }


# ============================================================
# EXCEL FORMATS
# ============================================================

def workbook_formats(
    workbook: xlsxwriter.Workbook,
) -> dict[str, Any]:

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
    worksheet,
    formats,
    title,
    subtitle,
    last_col,
):

    worksheet.merge_range(
        0,
        0,
        0,
        last_col,
        title,
        formats["title"],
    )

    worksheet.set_row(0, 30)

    worksheet.merge_range(
        1,
        0,
        1,
        last_col,
        subtitle,
        formats["subtitle"],
    )

    worksheet.set_row(1, 30)


def write_header(
    worksheet,
    row,
    headers,
    formats,
    start_col=0,
):

    for offset, header in enumerate(headers):

        worksheet.write(
            row,
            start_col + offset,
            header,
            formats["header"],
        )


def status_format(formats, status):

    if status == "Successful":
        return formats["success"]

    if status == "Failed":
        return formats["failed"]

    if status == "Pending Validation":
        return formats["pending"]

    return formats["body"]


# ============================================================
# SHEETS
# ============================================================

def write_management_summary(
    workbook,
    formats,
    transactions,
    start_dt,
    end_dt_exclusive,
    report_title,
):

    ws = workbook.add_worksheet(
        "Management Summary"
    )

    label = period_label(
        start_dt,
        end_dt_exclusive,
    )

    add_title(
        ws,
        formats,
        report_title,
        (
            "Management review of payments, failed payment attempts "
            f"and source of payments | {label}"
        ),
        7,
    )

    total = len(transactions)

    successful = sum(
        tx.status == "Successful"
        for tx in transactions
    )

    failed = sum(
        tx.status == "Failed"
        for tx in transactions
    )

    pending = sum(
        tx.status == "Pending Validation"
        for tx in transactions
    )

    completed = total - pending

    write_header(
        ws,
        3,
        ["KEY METRIC", "RESULT"],
        formats,
    )

    metrics = [
        (
            "Total payment attempts",
            total,
            formats["integer"],
        ),
        (
            "Successful payments",
            successful,
            formats["integer"],
        ),
        (
            "Failed attempts",
            failed,
            formats["integer"],
        ),
        (
            "Pending validation",
            pending,
            formats["integer"],
        ),
        (
            "Completed-attempt success rate",
            percentage(successful, completed),
            formats["percent"],
        ),
        (
            "Overall success rate",
            percentage(successful, total),
            formats["percent"],
        ),
        (
            "Review period",
            label,
            formats["body"],
        ),
    ]

    for row_number, (
        metric,
        value,
        value_format,
    ) in enumerate(metrics, start=4):

        ws.write(
            row_number,
            0,
            metric,
            formats["body"],
        )

        if isinstance(value, (int, float)):
            ws.write(
                row_number,
                1,
                value,
                value_format,
            )
        else:
            ws.write(
                row_number,
                1,
                value,
                value_format,
            )

    ws.merge_range(
        3,
        3,
        3,
        7,
        "MANAGEMENT OBSERVATIONS",
        formats["section"],
    )

    observation_text = "\n".join(
        "• " + item
        for item in build_observations(
            transactions
        )
    )

    ws.merge_range(
        4,
        3,
        10,
        7,
        observation_text,
        formats["note"],
    )

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

    write_header(
        ws,
        13,
        currency_headers,
        formats,
    )

    currencies = sorted(
        {
            tx.currency
            for tx in transactions
        }
    )

    currency_start = 14

    for offset, currency in enumerate(
        currencies
    ):

        row_number = currency_start + offset

        items = [
            tx
            for tx in transactions
            if tx.currency == currency
        ]

        successful_items = [
            tx
            for tx in items
            if tx.status == "Successful"
        ]

        failed_items = [
            tx
            for tx in items
            if tx.status == "Failed"
        ]

        pending_items = [
            tx
            for tx in items
            if tx.status == "Pending Validation"
        ]

        values = [
            currency,
            len(items),
            len(successful_items),
            len(failed_items),
            len(pending_items),
            sum(
                tx.amount
                for tx in successful_items
            ),
            sum(
                tx.amount
                for tx in failed_items
            ),
            sum(
                tx.amount
                for tx in pending_items
            ),
        ]

        for column, value in enumerate(values):

            if column >= 5:
                ws.write_number(
                    row_number,
                    column,
                    value,
                    formats["money"],
                )

            elif isinstance(value, int):
                ws.write_number(
                    row_number,
                    column,
                    value,
                    formats["integer"],
                )

            else:
                ws.write(
                    row_number,
                    column,
                    value,
                    formats["body"],
                )

    method_row = (
        currency_start
        + len(currencies)
        + 2
    )

    method_headers = [
        "Payment Method",
        "Attempts",
        "Successful",
        "Failed",
        "Pending",
        "Overall Success Rate",
    ]

    write_header(
        ws,
        method_row,
        method_headers,
        formats,
    )

    method_groups = defaultdict(list)

    for tx in transactions:
        method_groups[
            tx.payment_method
        ].append(tx)

    priority = [
        "bank_transfer",
        "card",
        "applepay",
        "account-ach-uk",
    ]

    methods = [
        method
        for method in priority
        if method in method_groups
    ]

    methods.extend(
        sorted(
            set(method_groups.keys())
            - set(methods)
        )
    )

    for offset, method in enumerate(
        methods,
        start=1,
    ):

        row_number = method_row + offset

        items = method_groups[method]

        success_count = sum(
            tx.status == "Successful"
            for tx in items
        )

        failed_count = sum(
            tx.status == "Failed"
            for tx in items
        )

        pending_count = sum(
            tx.status == "Pending Validation"
            for tx in items
        )

        label_text = (
            items[0].payment_method_label
        )

        values = [
            label_text,
            len(items),
            success_count,
            failed_count,
            pending_count,
            percentage(
                success_count,
                len(items),
            ),
        ]

        for column, value in enumerate(values):

            if column == 5:
                ws.write_number(
                    row_number,
                    column,
                    value,
                    formats["percent"],
                )

            elif isinstance(value, int):
                ws.write_number(
                    row_number,
                    column,
                    value,
                    formats["integer"],
                )

            else:
                ws.write(
                    row_number,
                    column,
                    value,
                    formats["body"],
                )

    chart_data_row = (
        method_row
        + len(methods)
        + 3
    )

    write_header(
        ws,
        chart_data_row,
        ["Status", "Attempts"],
        formats,
    )

    status_data = [
        ("Successful", successful),
        ("Failed", failed),
        ("Pending Validation", pending),
    ]

    for offset, (
        status,
        count,
    ) in enumerate(
        status_data,
        start=1,
    ):

        ws.write(
            chart_data_row + offset,
            0,
            status,
            formats["body"],
        )

        ws.write_number(
            chart_data_row + offset,
            1,
            count,
            formats["integer"],
        )

    pie = workbook.add_chart(
        {"type": "pie"}
    )

    pie.add_series(
        {
            "categories": [
                "Management Summary",
                chart_data_row + 1,
                0,
                chart_data_row + 3,
                0,
            ],
            "values": [
                "Management Summary",
                chart_data_row + 1,
                1,
                chart_data_row + 3,
                1,
            ],
            "data_labels": {
                "percentage": True,
            },
        }
    )

    pie.set_title(
        {
            "name": "Payment Attempts by Status"
        }
    )

    pie.set_legend(
        {
            "position": "bottom"
        }
    )

    pie.set_style(10)

    ws.insert_chart(
        "J4",
        pie,
        {
            "x_scale": 1.05,
            "y_scale": 1.05,
        },
    )

    column_chart = workbook.add_chart(
        {"type": "column"}
    )

    if methods:

        column_chart.add_series(
            {
                "name": "Attempts",
                "categories": [
                    "Management Summary",
                    method_row + 1,
                    0,
                    method_row + len(methods),
                    0,
                ],
                "values": [
                    "Management Summary",
                    method_row + 1,
                    1,
                    method_row + len(methods),
                    1,
                ],
                "data_labels": {
                    "value": True,
                },
            }
        )

    column_chart.set_title(
        {
            "name": "Attempts by Payment Method"
        }
    )

    column_chart.set_legend(
        {
            "none": True
        }
    )

    column_chart.set_style(10)

    ws.insert_chart(
        "J20",
        column_chart,
        {
            "x_scale": 1.05,
            "y_scale": 1.05,
        },
    )

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

    for column, width in widths.items():
        ws.set_column(
            column,
            column,
            width,
        )

    ws.freeze_panes(2, 0)


def write_weekly_trend(
    workbook,
    formats,
    transactions,
    start_dt,
    end_dt_exclusive,
):

    ws = workbook.add_worksheet(
        "Weekly Trend"
    )

    add_title(
        ws,
        formats,
        "PAYMENT TREND",
        (
            "Weekly tracking of successful, failed and pending "
            "payment attempts within the review period."
        ),
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

    write_header(
        ws,
        3,
        headers,
        formats,
    )

    weeks = group_by_week(
        transactions,
        start_dt,
        end_dt_exclusive,
    )

    for row_number, week in enumerate(
        weeks,
        start=4,
    ):

        date_range = (
            f"{week['start']:%d %b %Y} - "
            f"{week['end']:%d %b %Y}"
        )

        values = [
            week["week"],
            date_range,
            week["attempts"],
            week["successful"],
            week["failed"],
            week["pending"],
            week["success_rate"],
        ]

        for column, value in enumerate(values):

            if column == 6:
                ws.write_number(
                    row_number,
                    column,
                    value,
                    formats["percent"],
                )

            elif isinstance(value, int):
                ws.write_number(
                    row_number,
                    column,
                    value,
                    formats["integer"],
                )

            else:
                ws.write(
                    row_number,
                    column,
                    value,
                    formats["body"],
                )

    chart = workbook.add_chart(
        {"type": "column"}
    )

    if weeks:

        for column, name in [
            (3, "Successful"),
            (4, "Failed"),
            (5, "Pending"),
        ]:

            chart.add_series(
                {
                    "name": name,
                    "categories": [
                        "Weekly Trend",
                        4,
                        0,
                        3 + len(weeks),
                        0,
                    ],
                    "values": [
                        "Weekly Trend",
                        4,
                        column,
                        3 + len(weeks),
                        column,
                    ],
                }
            )

    chart.set_title(
        {
            "name": "Weekly Attempts by Status"
        }
    )

    chart.set_legend(
        {
            "position": "bottom"
        }
    )

    chart.set_style(10)

    ws.insert_chart(
        "I4",
        chart,
        {
            "x_scale": 1.15,
            "y_scale": 1.15,
        },
    )

    widths = [
        15,
        28,
        14,
        14,
        14,
        14,
        21,
    ]

    for column, width in enumerate(widths):
        ws.set_column(
            column,
            column,
            width,
        )

    ws.freeze_panes(4, 0)


def write_sources_sheet(
    workbook,
    formats,
    transactions,
):

    ws = workbook.add_worksheet(
        "Payment Sources"
    )

    add_title(
        ws,
        formats,
        "PAYMENT SOURCE ANALYSIS",
        (
            "Source is taken from the card issuer, originating bank "
            "or payment channel available in the gateway export."
        ),
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

    write_header(
        ws,
        3,
        headers,
        formats,
    )

    groups = defaultdict(list)

    for tx in transactions:

        key = (
            tx.payment_method_label,
            tx.source,
            tx.source_country,
            tx.instrument,
        )

        groups[key].append(tx)

    rows = []

    for key in sorted(groups):

        items = groups[key]

        successful = sum(
            tx.status == "Successful"
            for tx in items
        )

        failed = sum(
            tx.status == "Failed"
            for tx in items
        )

        pending = sum(
            tx.status == "Pending Validation"
            for tx in items
        )

        rows.append(
            [
                *key,
                len(items),
                successful,
                failed,
                pending,
                percentage(
                    successful,
                    len(items),
                ),
            ]
        )

    for row_number, row in enumerate(
        rows,
        start=4,
    ):

        for column, value in enumerate(row):

            if column == 8:
                ws.write_number(
                    row_number,
                    column,
                    value,
                    formats["percent"],
                )

            elif isinstance(value, int):
                ws.write_number(
                    row_number,
                    column,
                    value,
                    formats["integer"],
                )

            else:
                ws.write(
                    row_number,
                    column,
                    value,
                    (
                        formats["body_wrap"]
                        if column == 1
                        else formats["body"]
                    ),
                )

    if rows:

        ws.add_table(
            3,
            0,
            3 + len(rows),
            len(headers) - 1,
            {
                "name": "PaymentSourcesTable",
                "style": "Table Style Medium 2",
                "columns": [
                    {"header": header}
                    for header in headers
                ],
            },
        )

    widths = [
        23,
        46,
        23,
        18,
        13,
        13,
        13,
        13,
        14,
    ]

    for column, width in enumerate(widths):
        ws.set_column(
            column,
            column,
            width,
        )

    ws.freeze_panes(4, 0)



def write_donor_analysis_sheet(workbook, formats, transactions):
    ws = workbook.add_worksheet('Donor Analysis')
    add_title(
        ws,
        formats,
        'DONOR / FUNDER ANALYSIS',
        (
            'Who funded the account, how many times they funded, total successful amount funded, '
            'and successful, failed and pending transactions by donor. Amounts remain in original currency.'
        ),
        20,
    )

    rows = build_donor_analysis(transactions)
    metrics = donor_summary_metrics(transactions)

    write_header(ws, 3, ['DONOR METRIC', 'RESULT'], formats)
    donor_kpis = [
        ('Unique donors / funders', metrics['unique_donors']),
        ('Donors with successful funding', metrics['successful_donors']),
        ('Repeat donors (2+ successful fundings)', metrics['repeat_donors']),
        ('Donors with failed attempts', metrics['donors_with_failures']),
        ('Donors with pending attempts', metrics['donors_with_pending']),
    ]
    for r, (label, value) in enumerate(donor_kpis, start=4):
        ws.write(r, 0, label, formats['body'])
        ws.write_number(r, 1, value, formats['integer'])

    ws.merge_range(3, 3, 3, 8, 'HOW TO READ THIS TAB', formats['section'])
    ws.merge_range(
        4, 3, 8, 8,
        (
            '• Total Amount Funded = successful transaction value only.\n'
            '• Failed / Unsuccessful Count = failed transaction attempts.\n'
            '• Pending Validation is shown separately.\n'
            '• Funding Frequency is based on successful funding events.\n'
            '• Each donor is split by currency so unlike currencies are never added together.\n'
            '• Bank transfers prioritise the captured originator/account name.'
        ),
        formats['note'],
    )

    table_row = 11
    headers = [
        'Donor / Funder', 'Currency', 'Funding Frequency', 'Total Attempts',
        'Successful Funding Count', 'Failed / Unsuccessful Count', 'Pending Validation Count',
        'Other Status Count', 'Total Amount Funded', 'Failed Attempt Value',
        'Pending Attempt Value', 'Total Attempted Value', 'Successful Funding Rate',
        'First Funding Attempt', 'Last Funding Attempt', 'Payment Method(s)',
        'Source / Issuer(s)', 'Originator Name(s)', 'Email(s)',
        'Masked Source Account(s)', 'Identity Basis',
    ]
    write_header(ws, table_row, headers, formats)

    integer_headers = {
        'Total Attempts', 'Successful Funding Count', 'Failed / Unsuccessful Count',
        'Pending Validation Count', 'Other Status Count',
    }
    money_headers = {
        'Total Amount Funded', 'Failed Attempt Value',
        'Pending Attempt Value', 'Total Attempted Value',
    }
    date_headers = {'First Funding Attempt', 'Last Funding Attempt'}

    for row_number, row in enumerate(rows, start=table_row + 1):
        for column, header in enumerate(headers):
            value = row[header]
            if header in integer_headers:
                ws.write_number(row_number, column, int(value), formats['integer'])
            elif header in money_headers:
                ws.write_number(row_number, column, float(value), formats['money'])
            elif header == 'Successful Funding Rate':
                ws.write_number(row_number, column, float(value), formats['percent'])
            elif header in date_headers:
                ws.write_datetime(row_number, column, value.replace(tzinfo=None), formats['date'])
            else:
                ws.write(row_number, column, value, formats['body_wrap'] if column in {0, 15, 16, 17, 18, 19, 20} else formats['body'])

    if rows:
        ws.add_table(
            table_row, 0, table_row + len(rows), len(headers) - 1,
            {
                'name': 'DonorAnalysisTable',
                'style': 'Table Style Medium 2',
                'columns': [{'header': h} for h in headers],
            },
        )
        failed_col = headers.index('Failed / Unsuccessful Count')
        pending_col = headers.index('Pending Validation Count')
        ws.conditional_format(table_row + 1, failed_col, table_row + len(rows), failed_col, {
            'type': 'cell', 'criteria': '>', 'value': 0, 'format': formats['failed']
        })
        ws.conditional_format(table_row + 1, pending_col, table_row + len(rows), pending_col, {
            'type': 'cell', 'criteria': '>', 'value': 0, 'format': formats['pending']
        })

    widths = [30, 10, 18, 14, 20, 21, 20, 18, 20, 19, 20, 20, 19, 20, 20, 26, 34, 28, 31, 27, 28]
    for column, width in enumerate(widths):
        ws.set_column(column, column, width)
    ws.freeze_panes(table_row + 1, 0)


def write_failures_sheet(
    workbook,
    formats,
    transactions,
):

    ws = workbook.add_worksheet(
        "Failed & Pending"
    )

    add_title(
        ws,
        formats,
        "FAILED & PENDING PAYMENT ATTEMPTS",
        (
            "Detailed exceptions requiring review, including gateway "
            "messages and payment source."
        ),
        10,
    )

    exceptions = [
        tx
        for tx in transactions
        if tx.status != "Successful"
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

    write_header(
        ws,
        3,
        headers,
        formats,
    )

    for row_number, tx in enumerate(
        exceptions,
        start=4,
    ):

        values = [
            tx.created.replace(
                tzinfo=None
            ),
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

        for column, value in enumerate(values):

            if column == 0:
                ws.write_datetime(
                    row_number,
                    column,
                    value,
                    formats["date"],
                )

            elif column == 2:
                ws.write_number(
                    row_number,
                    column,
                    value,
                    formats["money"],
                )

            elif column == 7:
                ws.write(
                    row_number,
                    column,
                    value,
                    status_format(
                        formats,
                        tx.status,
                    ),
                )

            else:
                ws.write(
                    row_number,
                    column,
                    value,
                    (
                        formats["body_wrap"]
                        if column
                        in {1, 5, 8, 9, 10}
                        else formats["body"]
                    ),
                )

    if exceptions:

        ws.add_table(
            3,
            0,
            3 + len(exceptions),
            len(headers) - 1,
            {
                "name": "FailedPendingTable",
                "style": "Table Style Medium 2",
                "columns": [
                    {"header": header}
                    for header in headers
                ],
            },
        )

    reason_counts = Counter(
        tx.gateway_message
        or "No gateway message"
        for tx in exceptions
    )

    write_header(
        ws,
        3,
        [
            "Failure / Pending Reason",
            "Count",
        ],
        formats,
        start_col=12,
    )

    for row_number, (
        reason,
        count,
    ) in enumerate(
        reason_counts.most_common(),
        start=4,
    ):

        ws.write(
            row_number,
            12,
            reason,
            formats["body_wrap"],
        )

        ws.write_number(
            row_number,
            13,
            count,
            formats["integer"],
        )

    widths = [
        19,
        42,
        13,
        10,
        23,
        43,
        22,
        20,
        50,
        25,
        31,
    ]

    for column, width in enumerate(widths):
        ws.set_column(
            column,
            column,
            width,
        )

    ws.set_column(
        12,
        12,
        52,
    )

    ws.set_column(
        13,
        13,
        12,
    )

    ws.freeze_panes(4, 0)


def write_transactions_sheet(
    workbook,
    formats,
    transactions,
):

    ws = workbook.add_worksheet(
        "Test Transactions"
    )

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

    write_header(
        ws,
        0,
        headers,
        formats,
    )

    for row_number, tx in enumerate(
        transactions,
        start=1,
    ):

        values = [
            tx.created.replace(
                tzinfo=None
            ),
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

        for column, value in enumerate(values):

            if column == 0:
                ws.write_datetime(
                    row_number,
                    column,
                    value,
                    formats["date"],
                )

            elif column in {3, 16}:

                if value is None:
                    ws.write_blank(
                        row_number,
                        column,
                        None,
                        formats["money"],
                    )
                else:
                    ws.write_number(
                        row_number,
                        column,
                        value,
                        formats["money"],
                    )

            elif column == 9:

                ws.write(
                    row_number,
                    column,
                    value,
                    status_format(
                        formats,
                        tx.status,
                    ),
                )

            else:

                ws.write(
                    row_number,
                    column,
                    value,
                    (
                        formats["body_wrap"]
                        if column
                        in {1, 6, 10, 11, 12, 13}
                        else formats["body"]
                    ),
                )

    if transactions:

        ws.add_table(
            0,
            0,
            len(transactions),
            len(headers) - 1,
            {
                "name": "TestTransactionsTable",
                "style": "Table Style Medium 2",
                "columns": [
                    {"header": header}
                    for header in headers
                ],
            },
        )

    widths = [
        19,
        42,
        16,
        13,
        10,
        23,
        42,
        22,
        17,
        20,
        48,
        25,
        31,
        30,
        24,
        19,
        18,
    ]

    for column, width in enumerate(widths):
        ws.set_column(
            column,
            column,
            width,
        )

    ws.freeze_panes(1, 0)


def write_historical_sheet(
    workbook,
    formats,
    all_transactions,
    report_transactions,
    start_dt,
    end_dt_exclusive,
):

    ws = workbook.add_worksheet(
        "Historical Overview"
    )

    add_title(
        ws,
        formats,
        "FULL EXPORT - HISTORICAL OVERVIEW",
        (
            "Transactions outside the selected management review period "
            "are kept here for context."
        ),
        10,
    )

    earliest = min(
        tx.created
        for tx in all_transactions
    )

    latest = max(
        tx.created
        for tx in all_transactions
    )

    write_header(
        ws,
        3,
        [
            "Full Export Metric",
            "Value",
        ],
        formats,
    )

    metrics = [
        (
            "Total valid transactions",
            len(all_transactions),
        ),
        (
            "Earliest transaction",
            earliest.replace(
                tzinfo=None
            ),
        ),
        (
            "Latest transaction",
            latest.replace(
                tzinfo=None
            ),
        ),
        (
            "Transactions in review period",
            len(report_transactions),
        ),
        (
            "Transactions outside review period",
            len(all_transactions)
            - len(report_transactions),
        ),
        (
            "Review period",
            period_label(
                start_dt,
                end_dt_exclusive,
            ),
        ),
    ]

    for row_number, (
        metric,
        value,
    ) in enumerate(
        metrics,
        start=4,
    ):

        ws.write(
            row_number,
            0,
            metric,
            formats["body"],
        )

        if isinstance(
            value,
            datetime,
        ):

            ws.write_datetime(
                row_number,
                1,
                value,
                formats["date"],
            )

        elif isinstance(
            value,
            int,
        ):

            ws.write_number(
                row_number,
                1,
                value,
                formats["integer"],
            )

        else:

            ws.write(
                row_number,
                1,
                value,
                formats["body"],
            )

    status_counts = Counter(
        tx.status
        for tx in all_transactions
    )

    method_counts = Counter(
        tx.payment_method_label
        for tx in all_transactions
    )

    currency_counts = Counter(
        tx.currency
        for tx in all_transactions
    )

    write_header(
        ws,
        3,
        ["Status", "Transactions"],
        formats,
        3,
    )

    for row_number, (
        status,
        count,
    ) in enumerate(
        sorted(
            status_counts.items()
        ),
        start=4,
    ):

        ws.write(
            row_number,
            3,
            status,
            formats["body"],
        )

        ws.write_number(
            row_number,
            4,
            count,
            formats["integer"],
        )

    write_header(
        ws,
        3,
        [
            "Payment Method",
            "Transactions",
        ],
        formats,
        6,
    )

    for row_number, (
        method,
        count,
    ) in enumerate(
        sorted(
            method_counts.items()
        ),
        start=4,
    ):

        ws.write(
            row_number,
            6,
            method,
            formats["body"],
        )

        ws.write_number(
            row_number,
            7,
            count,
            formats["integer"],
        )

    write_header(
        ws,
        3,
        [
            "Currency",
            "Transactions",
        ],
        formats,
        9,
    )

    for row_number, (
        currency,
        count,
    ) in enumerate(
        sorted(
            currency_counts.items()
        ),
        start=4,
    ):

        ws.write(
            row_number,
            9,
            currency,
            formats["body"],
        )

        ws.write_number(
            row_number,
            10,
            count,
            formats["integer"],
        )

    ws.set_column(
        0,
        0,
        31,
    )

    ws.set_column(
        1,
        1,
        25,
    )

    ws.set_column(
        3,
        3,
        28,
    )

    ws.set_column(
        6,
        6,
        26,
    )

    ws.set_column(
        9,
        10,
        16,
    )


def write_raw_sheet(
    workbook,
    formats,
    headers,
    raw_rows,
):

    ws = workbook.add_worksheet(
        "Raw Data"
    )

    write_header(
        ws,
        0,
        headers,
        formats,
    )

    for row_number, row in enumerate(
        raw_rows,
        start=1,
    ):

        for column, header in enumerate(
            headers
        ):

            value = row.get(
                header,
                "",
            )

            if header == "created":

                parsed = parse_datetime(
                    value
                )

                if parsed:

                    ws.write_datetime(
                        row_number,
                        column,
                        parsed.replace(
                            tzinfo=None
                        ),
                        formats["date"],
                    )

                    continue

            if header in NUMERIC_COLUMNS:

                numeric = safe_float(
                    value
                )

                if numeric is not None:

                    ws.write_number(
                        row_number,
                        column,
                        numeric,
                        formats["money"],
                    )

                    continue

            ws.write(
                row_number,
                column,
                clean(value),
                formats["body"],
            )

    if raw_rows:

        ws.add_table(
            0,
            0,
            len(raw_rows),
            len(headers) - 1,
            {
                "name": "RawGatewayData",
                "style": "Table Style Medium 2",
                "columns": [
                    {"header": header}
                    for header in headers
                ],
            },
        )

    ws.freeze_panes(
        1,
        0,
    )

    ws.set_column(
        0,
        max(
            len(headers) - 1,
            0,
        ),
        15,
    )

    wider_columns = {
        "txref": 42,
        "chargemessage": 50,
        "authurl": 45,
        "pcardname": 45,
        "custname": 25,
        "custemail": 31,
        "originatorname": 30,
        "originatoraccountnumber": 24,
    }

    for header_name, width in (
        wider_columns.items()
    ):

        if header_name in headers:

            column = headers.index(
                header_name
            )

            ws.set_column(
                column,
                column,
                width,
            )


# ============================================================
# REPORT GENERATOR
# ============================================================

def generate_report_bytes(
    file_bytes: bytes,
    start_date: date,
    end_date: date,
    report_title: str,
) -> tuple[bytes, dict[str, Any]]:

    headers, raw_rows = read_csv_bytes(
        file_bytes
    )

    all_transactions = [
        tx
        for row in raw_rows
        if (
            tx := normalize_transaction(
                row
            )
        )
        is not None
    ]

    if not all_transactions:

        raise ValueError(
            "No valid transactions were found. "
            "Please check the created column."
        )

    start_dt = to_utc_start(
        start_date
    )

    end_dt_exclusive = (
        to_utc_end_exclusive(
            end_date
        )
    )

    selected_transactions = sorted(
        [
            tx
            for tx in all_transactions
            if (
                start_dt
                <= tx.created
                < end_dt_exclusive
            )
        ],
        key=lambda tx: tx.created,
    )

    if not selected_transactions:

        raise ValueError(
            "No transactions fall within the selected date range."
        )

    output = io.BytesIO()

    workbook = xlsxwriter.Workbook(
        output,
        {
            "in_memory": True,
        },
    )

    workbook.set_properties(
        {
            "title": report_title.title(),
            "subject": (
                "Payment test management report"
            ),
            "author": (
                "Payment Report Generator"
            ),
        }
    )

    formats = workbook_formats(
        workbook
    )

    write_management_summary(
        workbook,
        formats,
        selected_transactions,
        start_dt,
        end_dt_exclusive,
        report_title,
    )

    write_weekly_trend(
        workbook,
        formats,
        selected_transactions,
        start_dt,
        end_dt_exclusive,
    )

    write_sources_sheet(
        workbook,
        formats,
        selected_transactions,
    )

    write_donor_analysis_sheet(
        workbook,
        formats,
        selected_transactions,
    )

    write_failures_sheet(
        workbook,
        formats,
        selected_transactions,
    )

    write_transactions_sheet(
        workbook,
        formats,
        selected_transactions,
    )

    write_historical_sheet(
        workbook,
        formats,
        all_transactions,
        selected_transactions,
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

    output.seek(0)

    successful = sum(
        tx.status == "Successful"
        for tx in selected_transactions
    )

    failed = sum(
        tx.status == "Failed"
        for tx in selected_transactions
    )

    pending = sum(
        tx.status == "Pending Validation"
        for tx in selected_transactions
    )

    details = {
        "attempts": len(
            selected_transactions
        ),
        "successful": successful,
        "failed": failed,
        "pending": pending,
        "period": period_label(
            start_dt,
            end_dt_exclusive,
        ),
        "filename": (
            "Management_Payment_Test_Report_"
            + filename_period(
                start_dt,
                end_dt_exclusive,
            )
            + ".xlsx"
        ),
        "all_transactions": all_transactions,
        "selected_transactions": selected_transactions,
        "donor_analysis": build_donor_analysis(selected_transactions),
        "donor_metrics": donor_summary_metrics(selected_transactions),
    }

    return output.getvalue(), details


# ============================================================
# STREAMLIT APPLICATION
# ============================================================

def main():

    st.title(
        "📊 Payment Test Management Report Generator"
    )

    st.caption(
        "Upload a Flutterwave-style CSV export and generate a "
        "management-ready Excel report."
    )

    uploaded_file = st.file_uploader(
        "Upload transaction CSV",
        type=["csv"],
        help=(
            "The CSV must contain created, status, amount, "
            "currency and paymenttype columns."
        ),
    )

    if uploaded_file is None:

        st.info(
            "Upload a CSV file to begin."
        )

        with st.expander(
            "What the generated report contains"
        ):

            st.markdown(
                """
- Management Summary
- Weekly Trend
- Payment Sources
- Donor Analysis
- Failed & Pending
- Test Transactions
- Historical Overview
- Raw Data
"""
            )

        return

    try:

        file_bytes = (
            uploaded_file.getvalue()
        )

        headers, raw_rows = (
            read_csv_bytes(
                file_bytes
            )
        )

        all_transactions = [
            tx
            for row in raw_rows
            if (
                tx := normalize_transaction(
                    row
                )
            )
            is not None
        ]

        if not all_transactions:

            st.error(
                "No valid transaction dates were found in the CSV."
            )

            return

        earliest_date = min(
            tx.created.date()
            for tx in all_transactions
        )

        latest_date = max(
            tx.created.date()
            for tx in all_transactions
        )

        default_start = max(
            earliest_date,
            latest_date
            - timedelta(days=20),
        )

        st.success(
            f"CSV loaded successfully: {len(all_transactions)} valid transactions."
        )

        col1, col2 = st.columns(2)

        with col1:

            start_date = st.date_input(
                "Report start date",
                value=default_start,
                min_value=earliest_date,
                max_value=latest_date,
            )

        with col2:

            end_date = st.date_input(
                "Report end date",
                value=latest_date,
                min_value=earliest_date,
                max_value=latest_date,
            )

        report_title = st.text_input(
            "Report title",
            value=(
                "PAYMENT TEST MONITORING REPORT"
            ),
        )

        if start_date > end_date:

            st.error(
                "The start date cannot be after the end date."
            )

            return

        selected = [
            tx
            for tx in all_transactions
            if (
                to_utc_start(start_date)
                <= tx.created
                < to_utc_end_exclusive(end_date)
            )
        ]

        st.subheader(
            "Preview"
        )

        successful = sum(
            tx.status == "Successful"
            for tx in selected
        )

        failed = sum(
            tx.status == "Failed"
            for tx in selected
        )

        pending = sum(
            tx.status
            == "Pending Validation"
            for tx in selected
        )

        m1, m2, m3, m4 = st.columns(4)

        m1.metric(
            "Attempts",
            len(selected),
        )

        m2.metric(
            "Successful",
            successful,
        )

        m3.metric(
            "Failed",
            failed,
        )

        m4.metric(
            "Pending",
            pending,
        )

        if selected:

            donor_metrics = donor_summary_metrics(selected)
            donor_rows = build_donor_analysis(selected)

            st.subheader("Donor / Funder Preview")
            d1, d2, d3, d4 = st.columns(4)
            d1.metric("Unique donors", donor_metrics["unique_donors"])
            d2.metric("Successful donors", donor_metrics["successful_donors"])
            d3.metric("Repeat donors", donor_metrics["repeat_donors"])
            d4.metric("Donors with failed attempts", donor_metrics["donors_with_failures"])

            if donor_rows:
                preview_rows = [
                    {
                        "Donor / Funder": row["Donor / Funder"],
                        "Currency": row["Currency"],
                        "Successful Funding Count": row["Successful Funding Count"],
                        "Failed / Unsuccessful Count": row["Failed / Unsuccessful Count"],
                        "Pending Validation Count": row["Pending Validation Count"],
                        "Total Amount Funded": row["Total Amount Funded"],
                        "Successful Funding Rate": row["Successful Funding Rate"],
                    }
                    for row in donor_rows[:25]
                ]
                st.dataframe(preview_rows, use_container_width=True, hide_index=True)

            completed = (
                len(selected)
                - pending
            )

            completed_success_rate = (
                percentage(
                    successful,
                    completed,
                )
            )

            st.write(
                "Completed-attempt success rate: "
                f"**{completed_success_rate:.1%}**"
            )

        else:

            st.warning(
                "There are no transactions in the selected period."
            )

            return

        if st.button(
            "Generate Excel Report",
            type="primary",
            use_container_width=True,
        ):

            with st.spinner(
                "Generating report..."
            ):

                report_bytes, details = (
                    generate_report_bytes(
                        file_bytes=file_bytes,
                        start_date=start_date,
                        end_date=end_date,
                        report_title=report_title,
                    )
                )

            st.success(
                "Report generated successfully."
            )

            st.download_button(
                label="⬇️ Download Excel Report",
                data=report_bytes,
                file_name=details[
                    "filename"
                ],
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                type="primary",
                use_container_width=True,
            )

    except Exception as error:

        st.error(
            "The report could not be generated."
        )

        st.exception(
            error
        )


main()
