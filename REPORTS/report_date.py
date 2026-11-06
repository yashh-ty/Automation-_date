import calendar
import datetime
import re

_YESTERDAY_EXPR = re.compile(
    r"\(\s*CURRENT_DATE\s*-\s*1\s*\)\s*::\s*DATE",
    re.IGNORECASE,
)
_CURRENT_DATE_MINUS_ONE = re.compile(
    r"CURRENT_DATE\s*-\s*1",
    re.IGNORECASE,
)
_MONTH_START_EXPR = re.compile(
    r"DATE_TRUNC\(\s*'month'\s*,\s*CURRENT_DATE\s*\)\s*::\s*DATE",
    re.IGNORECASE,
)
_BETWEEN_YESTERDAY = re.compile(
    r"(between\s+(?:d\.)?month_start\s+and\s+)(?:d\.)?yesterday\b",
    re.IGNORECASE,
)


def prompt_report_date():
    """Ask for DD-MM-YYYY. Enter keeps the queries on CURRENT_DATE - 1."""
    while True:
        raw = input(
            "\nEnter report date (DD-MM-YYYY), or press Enter to keep yesterday: "
        ).strip()
        if raw == "":
            print("Using the existing yesterday date from the queries.")
            return None
        try:
            report_date = datetime.datetime.strptime(raw, "%d-%m-%Y").date()
        except ValueError:
            print("Invalid date. Use DD-MM-YYYY, for example 05-10-2026.")
            continue
        month_end = _month_end(report_date)
        print(
            f"Daily data: {report_date.strftime('%d-%m-%Y')}. "
            f"Monthly data: {report_date.strftime('%d-%m-%Y')} to {month_end.strftime('%d-%m-%Y')}."
        )
        return report_date


def sheet_headers(report_date):
    """Excel titles. A custom date uses that day through the end of its month."""
    if report_date is None:
        today = datetime.date.today()
        text = today.strftime("%d-%m-%Y")
        return text, f"MTD (1st to {text.split('-')[0]} July)"
    month_end = _month_end(report_date)
    left = report_date.strftime("%d-%m-%Y")
    right = (
        f"MTD ({report_date.day:02d} to {month_end.day:02d} {report_date.strftime('%B')})"
    )
    return left, right


def apply_report_date(sql, filename, report_date):
    """Return SQL with the entered date applied. The file on disk is not changed."""
    if report_date is None:
        return sql

    kind = _query_kind(filename)
    month_end = _month_end(report_date)
    if kind == "monthly":
        yesterday_value = month_end
        month_start_value = report_date
    else:
        yesterday_value = report_date
        month_start_value = report_date

    sql = _YESTERDAY_EXPR.sub(_date_literal(yesterday_value), sql)
    sql = _CURRENT_DATE_MINUS_ONE.sub(_date_literal(report_date), sql)
    sql = _MONTH_START_EXPR.sub(_date_literal(month_start_value), sql)

    if kind == "mixed":
        sql = _BETWEEN_YESTERDAY.sub(rf"\1{_date_literal(month_end)}", sql)
    return sql


def _query_kind(filename):
    name = filename.lower()
    if "monthly" in name:
        return "monthly"
    if "state_tax" in name or "tin_and_gst" in name:
        return "mixed"
    return "daily"


def _month_end(report_date):
    last_day = calendar.monthrange(report_date.year, report_date.month)[1]
    return datetime.date(report_date.year, report_date.month, last_day)


def _date_literal(value):
    return f"DATE '{value.isoformat()}'"
