"""Period and scenario parsing helpers."""

from __future__ import annotations

import calendar
import re
from datetime import date, datetime
from typing import Any

from .utils import normalize_label


MONTHS = {
    name.casefold(): index
    for index, name in enumerate(calendar.month_name)
    if name
}
MONTHS.update(
    {
        name.casefold(): index
        for index, name in enumerate(calendar.month_abbr)
        if name
    }
)
SCENARIOS = {
    "actual": "Actual",
    "budget": "Budget",
    "forecast": "Forecast",
    "plan": "Plan",
    "target": "Target",
    "prior year": "Prior Year",
    "previous year": "Previous Year",
}


def _expand_year(value: int) -> int:
    """Expand a two-digit reporting year."""

    if value >= 100:
        return value
    return 2000 + value if value < 70 else 1900 + value


def parse_period_label(value: Any) -> dict[str, Any]:
    """Parse a common business-period label without guessing ambiguity."""

    result = {
        "period_label_original": value,
        "period_type": None,
        "period_start": None,
        "period_end": None,
        "period_year": None,
        "period_number": None,
        "scenario": None,
        "period_parse_status": "unparsed",
        "period_parse_confidence": 0.0,
    }

    if isinstance(value, datetime):
        value = value.date()
    if isinstance(value, date):
        result.update(
            {
                "period_type": "date",
                "period_start": value.isoformat(),
                "period_end": value.isoformat(),
                "period_year": value.year,
                "period_number": value.month,
                "period_parse_status": "parsed",
                "period_parse_confidence": 1.0,
            }
        )
        return result

    if isinstance(value, int) and 1900 <= value <= 2200:
        result.update(
            {
                "period_type": "year",
                "period_start": date(value, 1, 1).isoformat(),
                "period_end": date(value, 12, 31).isoformat(),
                "period_year": value,
                "period_parse_status": "parsed",
                "period_parse_confidence": 0.98,
            }
        )
        return result

    text = normalize_label(value)
    if not text:
        return result

    for key, scenario in SCENARIOS.items():
        if text == key:
            result.update(
                {
                    "scenario": scenario,
                    "period_parse_status": "scenario_only",
                    "period_parse_confidence": 0.98,
                }
            )
            return result

    match = re.fullmatch(r"fy\s*(\d{2}|\d{4})", text)
    if match:
        year = _expand_year(int(match.group(1)))
        result.update(
            {
                "period_type": "fiscal_year",
                "period_start": date(year, 1, 1).isoformat(),
                "period_end": date(year, 12, 31).isoformat(),
                "period_year": year,
                "period_parse_status": "parsed",
                "period_parse_confidence": 0.95,
            }
        )
        return result

    match = re.fullmatch(
        r"(?:q([1-4])\s*[-/]?\s*(\d{2}|\d{4})|"
        r"(\d{2}|\d{4})\s*[-/]?\s*q([1-4]))",
        text,
    )
    if match:
        quarter = int(match.group(1) or match.group(4))
        year = _expand_year(int(match.group(2) or match.group(3)))
        start_month = 3 * (quarter - 1) + 1
        end_month = start_month + 2
        result.update(
            {
                "period_type": "quarter",
                "period_start": date(year, start_month, 1).isoformat(),
                "period_end": date(
                    year,
                    end_month,
                    calendar.monthrange(year, end_month)[1],
                ).isoformat(),
                "period_year": year,
                "period_number": quarter,
                "period_parse_status": "parsed",
                "period_parse_confidence": 0.98,
            }
        )
        return result

    match = re.fullmatch(r"(\d{4})[-/](0?[1-9]|1[0-2])", text)
    if match:
        year = int(match.group(1))
        month = int(match.group(2))
        result.update(
            {
                "period_type": "month",
                "period_start": date(year, month, 1).isoformat(),
                "period_end": date(
                    year, month, calendar.monthrange(year, month)[1]
                ).isoformat(),
                "period_year": year,
                "period_number": month,
                "period_parse_status": "parsed",
                "period_parse_confidence": 0.99,
            }
        )
        return result

    month_pattern = "|".join(sorted(MONTHS, key=len, reverse=True))
    match = re.fullmatch(
        rf"({month_pattern})[\s-]+(\d{{2}}|\d{{4}})", text
    )
    if match:
        month = MONTHS[match.group(1)]
        year = _expand_year(int(match.group(2)))
        result.update(
            {
                "period_type": "month",
                "period_start": date(year, month, 1).isoformat(),
                "period_end": date(
                    year, month, calendar.monthrange(year, month)[1]
                ).isoformat(),
                "period_year": year,
                "period_number": month,
                "period_parse_status": "parsed",
                "period_parse_confidence": 0.96,
            }
        )
        return result

    if text in MONTHS:
        result.update(
            {
                "period_type": "month_of_year",
                "period_number": MONTHS[text],
                "period_parse_status": "partial",
                "period_parse_confidence": 0.75,
            }
        )
        return result

    return result
