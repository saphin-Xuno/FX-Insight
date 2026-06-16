"""Send a formatted FX Insights Slack message after each daily run.

Reads data directly from the AppState populated by main_script.py.
Import and call send_slack_summary(state) from the run() function.
"""

import json
import os
import urllib.request
from datetime import datetime, timedelta


def _arrow(change: float | None) -> str:
    """Return + or - prefix string for a percentage change."""
    if change is None:
        return "N/A"
    sign = "+" if change >= 0 else ""
    return f"{sign}{change * 100:.2f}%"


def _price(value: float | None, decimals: int = 2) -> str:
    if value is None:
        return "N/A"
    return f"{value:,.{decimals}f}"


def _weekday(date_str: str) -> str:
    """Return 'Monday', 'Tuesday' etc. from a YYYY-MM-DD string."""
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        return dt.strftime("%A")
    except ValueError:
        return ""


def _format_date_label(date_str: str) -> str:
    """Return '15th June (Monday)' from '2025-06-15'."""
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        day = int(dt.strftime("%d"))
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(day % 10 if day not in (11, 12, 13) else 0, "th")
        return dt.strftime(f"%-d{suffix} %B (%A)").lstrip("0")
    except ValueError:
        return date_str


def build_slack_message(state) -> str:
    """Build the FX Insights Slack message matching the sample format."""
    market = state.market
    fii = state.fii

    npr = market.get("USDNPR")
    npr_open = _price(npr["open"] if npr else None)
    npr_close = _price(npr["close"] if npr else None)
    npr_date = npr["date"] if npr else ""
    npr_change = _arrow((npr["close"] - npr["open"]) / npr["open"] if npr else None)

    # Previous day label: one business day before current date
    try:
        current_dt = datetime.strptime(npr_date, "%Y-%m-%d")
        prev_dt = current_dt - timedelta(days=3 if current_dt.weekday() == 0 else 1)
        prev_label = _format_date_label(prev_dt.strftime("%Y-%m-%d"))
    except ValueError:
        prev_label = "previous session"

    current_label = _format_date_label(npr_date)

    gold = market.get("XAUUSD")
    gold_close = _price(gold["close"] if gold else None)
    gold_change = _arrow((gold["close"] - gold["open"]) / gold["open"] if gold else None)

    silver = market.get("XAGUSD")
    silver_close = _price(silver["close"] if silver else None)
    silver_change = _arrow((silver["close"] - silver["open"]) / silver["open"] if silver else None)

    oil = market.get("OIL")
    oil_close = _price(oil["close"] if oil else None)
    oil_change = _arrow((oil["close"] - oil["open"]) / oil["open"] if oil else None)

    ndx = market.get("NDX")
    ndx_close = _price(ndx["close"] if ndx else None, decimals=2)
    ndx_change = _arrow((ndx["close"] - ndx["open"]) / ndx["open"] if ndx else None)

    spx = market.get("SPX")
    spx_close = _price(spx["close"] if spx else None, decimals=2)
    spx_change = _arrow((spx["close"] - spx["open"]) / spx["open"] if spx else None)

    lines = [
        f"*FX Insights (USD/NPR):*",
        f"Opening {current_label}: NPR {npr_open}",
        f"Closing {prev_label}: NPR {npr_close}",
        f"Percentage Change in Price: {npr_change}",
        "",
        "*Commodity Market Movement:*",
        f"• Gold: ${gold_close} ({gold_change})",
        f"• Silver: ${silver_close} ({silver_change})",
        f"• Crude Oil: ${oil_close} ({oil_change})",
        "",
        "*USA Capital Market Updates:*",
        f"• NASDAQ 100: {ndx_close} points ({ndx_change})",
        f"• S&P 500: {spx_close} points ({spx_change})",
    ]

    # Optional FII/DII block if data is available
    fii_net = fii.get("fii_net")
    dii_net = fii.get("dii_net")
    if fii_net is not None or dii_net is not None:
        lines += [
            "",
            "*FII / DII Activity (India):*",
            f"• FII Net: ₹{_price(fii_net)} Cr",
            f"• DII Net: ₹{_price(dii_net)} Cr",
        ]

    return "\n".join(lines)


def send_slack_summary(state) -> None:
    """Post the daily FX Insights summary to Slack."""
    webhook_url = os.environ.get("SLACK_WEBHOOK_URL", "")
    if not webhook_url:
        print("  SKIP Slack: SLACK_WEBHOOK_URL not set")
        return

    message = build_slack_message(state)
    payload = json.dumps({"text": message}).encode("utf-8")

    req = urllib.request.Request(
        webhook_url,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            if resp.status == 200:
                print("  OK Slack message sent successfully")
            else:
                print(f"  WARN Slack responded with status {resp.status}")
    except Exception as exc:
        print(f"  WARN Slack send failed: {exc}")
