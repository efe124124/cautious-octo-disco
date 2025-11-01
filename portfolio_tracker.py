#!/usr/bin/env python3
"""Create an Excel workbook for live stock and portfolio tracking using yfinance.

The script reads a JSON configuration file that defines the base currency and the
portfolio holdings, fetches the latest market data via the Yahoo Finance API
(through the yfinance package) and generates a formatted Excel workbook with a
holdings table and a summary sheet.

Example usage:

    python portfolio_tracker.py --config portfolio_config.json --output portfolio.xlsx

The generated workbook can be refreshed by re-running the script. In Excel you
can enable "Refresh All" on open so that the latest file produced by this script
is always loaded with up-to-date quotes.
"""

from __future__ import annotations

import argparse
import json
import math
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, Iterable, List, Optional

import pandas as pd
import yfinance as yf


@dataclass
class Holding:
    symbol: str
    quantity: float
    cost_basis: float
    currency: str


@dataclass
class Quote:
    last_price: Optional[float]
    previous_close: Optional[float]
    day_high: Optional[float]
    day_low: Optional[float]
    volume: Optional[float]


def load_config(path: Path) -> tuple[str, List[Holding]]:
    with path.open("r", encoding="utf-8") as fh:
        data = json.load(fh)

    base_currency = data.get("base_currency")
    if not base_currency:
        raise ValueError("Configuration file must define a 'base_currency'.")

    holdings: List[Holding] = []
    for raw in data.get("holdings", []):
        try:
            holding = Holding(
                symbol=str(raw["symbol"]).strip(),
                quantity=float(raw["quantity"]),
                cost_basis=float(raw["cost_basis"]),
                currency=str(raw.get("currency", base_currency)).strip().upper(),
            )
        except KeyError as exc:  # missing required key
            raise ValueError(f"Holding configuration is missing field: {exc}") from exc

        if not holding.symbol:
            raise ValueError("Holding symbol cannot be empty.")
        if holding.quantity <= 0:
            raise ValueError(f"Quantity for {holding.symbol} must be > 0.")
        if holding.cost_basis < 0:
            raise ValueError(f"Cost basis for {holding.symbol} must be >= 0.")

        holdings.append(holding)

    if not holdings:
        raise ValueError("Configuration file must contain at least one holding.")

    return base_currency.upper(), holdings


def _safe_float(value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def fetch_quote(symbol: str) -> Quote:
    ticker = yf.Ticker(symbol)
    fast = getattr(ticker, "fast_info", {}) or {}

    def get_fast_value(*keys: str) -> Optional[float]:
        for key in keys:
            value = fast.get(key)
            if value is not None and not (isinstance(value, float) and math.isnan(value)):
                return _safe_float(value)
        return None

    last_price = get_fast_value("last_price", "last")
    previous_close = get_fast_value("previous_close", "last_close", "open")
    day_high = get_fast_value("day_high", "regular_market_day_high")
    day_low = get_fast_value("day_low", "regular_market_day_low")
    volume = get_fast_value("last_volume", "regular_market_volume")

    if last_price is None or previous_close is None:
        history = ticker.history(period="1d")
        if not history.empty:
            close_value = _safe_float(history["Close"].iloc[-1])
            if last_price is None:
                last_price = close_value
            if previous_close is None and history.shape[0] > 1:
                previous_close = _safe_float(history["Close"].iloc[-2])
            elif previous_close is None:
                previous_close = _safe_float(history["Open"].iloc[-1])

    return Quote(last_price=last_price, previous_close=previous_close, day_high=day_high, day_low=day_low, volume=volume)


def fetch_fx_rate(source_currency: str, target_currency: str) -> float:
    source_currency = source_currency.upper()
    target_currency = target_currency.upper()
    if source_currency == target_currency:
        return 1.0

    pair = f"{source_currency}{target_currency}=X"
    rate = _extract_price_for_symbol(pair)
    if rate is not None:
        return rate

    # Fallback: try the inverse pair and invert the rate
    inverse_pair = f"{target_currency}{source_currency}=X"
    inverse_rate = _extract_price_for_symbol(inverse_pair)
    if inverse_rate is not None and inverse_rate != 0:
        return 1.0 / inverse_rate

    raise RuntimeError(
        f"Could not fetch FX rate for {source_currency} -> {target_currency}. "
        "Check that the currency codes are valid Yahoo Finance pairs."
    )


def _extract_price_for_symbol(symbol: str) -> Optional[float]:
    ticker = yf.Ticker(symbol)
    fast = getattr(ticker, "fast_info", {}) or {}
    price = fast.get("last_price") or fast.get("previous_close")
    price = _safe_float(price)
    if price is not None:
        return price

    history = ticker.history(period="1d")
    if not history.empty:
        return _safe_float(history["Close"].iloc[-1])
    return None


def build_portfolio_frame(base_currency: str, holdings: Iterable[Holding]) -> pd.DataFrame:
    rows: List[Dict[str, Optional[float]]] = []
    fx_cache: Dict[tuple[str, str], float] = {}

    for holding in holdings:
        quote = fetch_quote(holding.symbol)
        fx_key = (holding.currency, base_currency)
        if fx_key not in fx_cache:
            fx_cache[fx_key] = fetch_fx_rate(holding.currency, base_currency)
        fx_rate = fx_cache[fx_key]

        last_price = quote.last_price
        previous_close = quote.previous_close
        price_in_base = last_price * fx_rate if last_price is not None else None
        previous_close_base = previous_close * fx_rate if previous_close is not None else None

        position_value = None
        position_cost = holding.cost_basis * holding.quantity * fx_rate
        if price_in_base is not None:
            position_value = price_in_base * holding.quantity

        unrealized_pl = position_value - position_cost if position_value is not None else None
        unrealized_pl_pct = None
        if unrealized_pl is not None and position_cost:
            unrealized_pl_pct = (unrealized_pl / position_cost) * 100

        daily_change = None
        daily_change_pct = None
        if last_price is not None and previous_close is not None:
            daily_change_pct = ((last_price - previous_close) / previous_close) * 100 if previous_close else None
            if price_in_base is not None and previous_close_base is not None:
                daily_change = (price_in_base - previous_close_base) * holding.quantity

        rows.append(
            {
                "Symbol": holding.symbol,
                "Quantity": holding.quantity,
                "Cost Basis (Asset Currency)": holding.cost_basis,
                "Asset Currency": holding.currency,
                f"FX Rate (to {base_currency})": fx_rate,
                "Price (Asset Currency)": last_price,
                f"Price ({base_currency})": price_in_base,
                f"Value ({base_currency})": position_value,
                f"Cost ({base_currency})": position_cost,
                f"Unrealized P/L ({base_currency})": unrealized_pl,
                "Unrealized P/L %": unrealized_pl_pct,
                f"Daily Change ({base_currency})": daily_change,
                "Daily Change %": daily_change_pct,
                "Previous Close (Asset Currency)": previous_close,
                "Day High (Asset Currency)": quote.day_high,
                "Day Low (Asset Currency)": quote.day_low,
                "Volume": quote.volume,
            }
        )

    df = pd.DataFrame(rows)

    # Convert selected columns to numeric (pandas may treat None as NaN automatically)
    numeric_columns = [
        col
        for col in df.columns
        if col not in {"Symbol", "Asset Currency"}
    ]
    df[numeric_columns] = df[numeric_columns].apply(pd.to_numeric, errors="coerce")

    total_value = df[f"Value ({base_currency})"].sum(skipna=True)
    if total_value:
        df["Weight %"] = (df[f"Value ({base_currency})"] / total_value) * 100
    else:
        df["Weight %"] = pd.NA

    return df


def build_summary_frame(df: pd.DataFrame, base_currency: str) -> pd.DataFrame:
    total_cost = df[f"Cost ({base_currency})"].sum(skipna=True)
    total_value = df[f"Value ({base_currency})"].sum(skipna=True)
    total_unrealized = df[f"Unrealized P/L ({base_currency})"].sum(skipna=True)
    total_daily_change = df.get(f"Daily Change ({base_currency})", pd.Series(dtype=float)).sum(skipna=True)

    pl_pct = (total_unrealized / total_cost * 100) if total_cost else None
    previous_value = total_value - total_daily_change if pd.notna(total_value) else None
    daily_pct = None
    if previous_value:
        daily_pct = (total_daily_change / previous_value) * 100

    timestamp = datetime.now(timezone.utc).astimezone().replace(microsecond=0)

    rows = [
        ("Last Updated", timestamp.isoformat()),
        (f"Total Market Value ({base_currency})", total_value),
        (f"Total Cost ({base_currency})", total_cost),
        (f"Unrealized P/L ({base_currency})", total_unrealized),
        ("Unrealized P/L %", pl_pct),
        (f"Daily Change ({base_currency})", total_daily_change),
        ("Daily Change %", daily_pct),
    ]

    summary_df = pd.DataFrame(rows, columns=["Metric", "Value"])
    return summary_df


def format_workbook(output_path: Path, df: pd.DataFrame, summary_df: pd.DataFrame, base_currency: str) -> None:
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        df_for_excel = df.copy()
        percentage_columns = [col for col in df_for_excel.columns if col.endswith("%")]
        for col in percentage_columns:
            df_for_excel[col] = df_for_excel[col] / 100.0

        df_for_excel.to_excel(writer, sheet_name="Holdings", index=False)
        summary_for_excel = summary_df.copy()
        if "Value" in summary_for_excel.columns:
            summary_for_excel.loc[summary_for_excel["Metric"].str.contains("%"), "Value"] = (
                summary_for_excel.loc[summary_for_excel["Metric"].str.contains("%"), "Value"] / 100.0
            )
        summary_for_excel.to_excel(writer, sheet_name="Summary", index=False)

        workbook = writer.book
        holdings_ws = writer.sheets["Holdings"]
        summary_ws = writer.sheets["Summary"]

        currency_format = workbook.add_format({"num_format": "#,##0.00"})
        integer_format = workbook.add_format({"num_format": "#,##0"})
        quantity_format = workbook.add_format({"num_format": "#,##0.00"})
        percent_format = workbook.add_format({"num_format": "0.00%"})

        currency_columns = [
            f"Price ({base_currency})",
            f"Value ({base_currency})",
            f"Cost ({base_currency})",
            f"Unrealized P/L ({base_currency})",
            f"Daily Change ({base_currency})",
        ]
        asset_currency_columns = [
            "Cost Basis (Asset Currency)",
            "Price (Asset Currency)",
            "Previous Close (Asset Currency)",
            "Day High (Asset Currency)",
            "Day Low (Asset Currency)",
        ]

        column_formats: Dict[str, object] = {col: currency_format for col in currency_columns}
        column_formats.update({col: currency_format for col in asset_currency_columns})
        column_formats["Quantity"] = quantity_format
        column_formats[f"FX Rate (to {base_currency})"] = currency_format
        column_formats["Volume"] = integer_format
        for col in df_for_excel.columns:
            if col.endswith("%"):
                column_formats[col] = percent_format

        for idx, column in enumerate(df_for_excel.columns):
            column_width = max(len(column) + 2, 14)
            holdings_ws.set_column(idx, idx, column_width, column_formats.get(column))

        holdings_ws.freeze_panes(1, 0)
        holdings_ws.autofilter(0, 0, len(df_for_excel), len(df_for_excel.columns) - 1)

        summary_ws.set_column(0, 0, 32)
        summary_ws.set_column(1, 1, 26, currency_format)
        for row_idx, value in enumerate(summary_for_excel["Value"], start=1):
            metric = summary_for_excel.at[row_idx - 1, "Metric"]
            if isinstance(value, str):
                summary_ws.write(row_idx, 1, value)
            elif isinstance(value, (float, int)) and " %" in metric:
                summary_ws.write(row_idx, 1, value, percent_format)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate an Excel portfolio tracker with live market data.")
    parser.add_argument("--config", required=True, type=Path, help="Path to the JSON configuration file.")
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("portfolio.xlsx"),
        help="Output Excel file path (defaults to ./portfolio.xlsx).",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    base_currency, holdings = load_config(args.config)
    df = build_portfolio_frame(base_currency, holdings)
    summary_df = build_summary_frame(df, base_currency)
    format_workbook(args.output, df, summary_df, base_currency)
    print(f"Portfolio workbook generated: {args.output}")


if __name__ == "__main__":
    main()
