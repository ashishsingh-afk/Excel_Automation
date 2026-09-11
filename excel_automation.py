from __future__ import annotations
import argparse
import re
import sys
from dataclasses import dataclass, field
from difflib import get_close_matches
from typing import Optional
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, PieChart, ScatterChart, Reference, Series
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# 1. CLEANING & TRANSFORMATION

def detect_header_row(raw: pd.DataFrame, scan_rows: int = 10) -> int:
    """Guess which of the first `scan_rows` rows is the real header row
    (handles sheets that have a title / merged cells above the table)."""
    best_row, best_score = 0, -1
    for i in range(min(scan_rows, len(raw))):
        row = raw.iloc[i]
        non_null = row.notna().sum()
        stringy = row.apply(lambda v: isinstance(v, str)).sum()
        score = non_null + stringy
        if score > best_score:
            best_score, best_row = score, i
    return best_row


def load_all_sheets(path: str) -> dict[str, pd.DataFrame]:
    """Load every sheet, auto-detecting the header row for each."""
    raw_sheets = pd.read_excel(path, sheet_name=None, header=None)
    sheets = {}
    for name, raw in raw_sheets.items():
        if raw.empty:
            sheets[name] = raw
            continue
        header_row = detect_header_row(raw)
        header = raw.iloc[header_row]
        df = raw.iloc[header_row + 1:].reset_index(drop=True)
        df.columns = [str(c).strip() if pd.notna(c) else f"col_{i}" for i, c in enumerate(header)]
        sheets[name] = df
    return sheets


def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [
        re.sub(r"[^0-9a-zA-Z]+", "_", str(c)).strip("_").lower() or f"col_{i}"
        for i, c in enumerate(df.columns)
    ]
    return df


def infer_and_coerce_types(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """Try numeric, then date, then leave as text. Returns (df, log)."""
    df = df.copy()
    log = {"numeric": [], "date": [], "text": []}
    for col in df.columns:
        series = df[col]
        if series.dtype != object:
            if pd.api.types.is_numeric_dtype(series):
                log["numeric"].append(col)
            elif pd.api.types.is_datetime64_any_dtype(series):
                log["date"].append(col)
            continue

        cleaned = series.astype(str).str.strip()
        cleaned_numeric_candidate = cleaned.str.replace(r"[,$%]", "", regex=True)
        numeric = pd.to_numeric(cleaned_numeric_candidate, errors="coerce")
        non_null = cleaned.notna() & (cleaned != "") & (cleaned.str.lower() != "nan")
        numeric_hit_rate = numeric.notna().sum() / max(non_null.sum(), 1)

        if numeric_hit_rate > 0.85:
            df[col] = numeric
            log["numeric"].append(col)
            continue

        date = pd.to_datetime(cleaned, errors="coerce", format=None)
        date_hit_rate = date.notna().sum() / max(non_null.sum(), 1)
        if date_hit_rate > 0.85:
            df[col] = date
            log["date"].append(col)
            continue

        df[col] = cleaned.replace({"nan": pd.NA, "": pd.NA, "None": pd.NA})
        log["text"].append(col)
    return df, log


def clean_text_case(df: pd.DataFrame, type_log: dict) -> pd.DataFrame:
    """Proper-case likely name/label columns, upper-case likely code/ID columns."""
    df = df.copy()
    for col in type_log["text"]:
        series = df[col]
        if series.isna().all():
            continue
        lname = col.lower()
        avg_len = series.dropna().astype(str).str.len().mean() if series.notna().any() else 0
        looks_like_code = bool(re.search(r"(id|code|sku|ref|no)$", lname)) or (avg_len and avg_len <= 6 and series.dropna().astype(str).str.isupper().mean() > 0.5)
        if looks_like_code:
            df[col] = series.astype(str).str.strip().str.upper().where(series.notna())
        else:
            df[col] = series.astype(str).str.strip().str.title().where(series.notna())
    return df


def clean_sheet(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """Full cleaning pipeline for one sheet. Returns (cleaned_df, report)."""
    report = {"rows_in": len(df)}
    if df.empty:
        report.update(rows_out=0, dup_rows_removed=0, empty_cols_removed=0, type_log={})
        return df, report

    df = df.dropna(axis=1, how="all")
    df = df.dropna(axis=0, how="all")
    empty_cols_removed = report["rows_in"] and None  # placeholder, real count below

    df = standardize_columns(df)

    before = len(df)
    df = df.drop_duplicates()
    dup_removed = before - len(df)

    df, type_log = infer_and_coerce_types(df)
    df = clean_text_case(df, type_log)

    for col in type_log["numeric"]:
        df[col] = df[col].round(4)

    report.update(
        rows_out=len(df),
        dup_rows_removed=dup_removed,
        type_log=type_log,
    )
    return df.reset_index(drop=True), report

# 2. EDA

def flag_outliers_iqr(series: pd.Series) -> int:
    s = series.dropna()
    if len(s) < 4:
        return 0
    q1, q3 = s.quantile(0.25), s.quantile(0.75)
    iqr = q3 - q1
    if iqr == 0:
        return 0
    lo, hi = q1 - 1.5 * iqr, q3 + 1.5 * iqr
    return int(((s < lo) | (s > hi)).sum())


def build_eda(df: pd.DataFrame, type_log: dict) -> dict:
    eda = {}
    eda["shape"] = df.shape
    eda["missing"] = (
        df.isna().sum().rename("missing_count").to_frame()
        .assign(missing_pct=lambda x: (x["missing_count"] / max(len(df), 1) * 100).round(1))
        .reset_index().rename(columns={"index": "column"})
    )

    numeric_cols = type_log.get("numeric", [])
    if numeric_cols:
        desc = df[numeric_cols].describe().T.reset_index().rename(columns={"index": "column"})
        desc["outliers_iqr"] = [flag_outliers_iqr(df[c]) for c in numeric_cols]
        eda["numeric_summary"] = desc
        eda["correlation"] = df[numeric_cols].corr().round(3) if len(numeric_cols) > 1 else None
    else:
        eda["numeric_summary"] = None
        eda["correlation"] = None

    text_cols = type_log.get("text", [])
    cat_summaries = {}
    for c in text_cols:
        if df[c].nunique(dropna=True) <= 30:
            cat_summaries[c] = df[c].value_counts(dropna=True).head(10)
    eda["categorical_top_values"] = cat_summaries
    return eda

# 3. RULE-BASED DASHBOARD COMMAND PARSER


AGG_KEYWORDS = {
    "sum": ["total", "sum", "overall"],
    "mean": ["average", "avg", "mean"],
    "count": ["count", "number of", "how many"],
    "max": ["max", "maximum", "highest", "peak"],
    "min": ["min", "minimum", "lowest"],
}

CHART_KEYWORDS = {
    "pie": ["pie", "share", "proportion", "breakdown"],
    "line": ["trend", "over time", "line", "growth"],
    "bar": ["bar", "compare", "comparison", "performance", "ranking"],
    "scatter": ["scatter", "relationship", "correlation"],
}

TIME_GRANULARITY_KEYWORDS = {
    "quarter": ["quarter", "quarterly", "q1", "q2", "q3", "q4"],
    "year": ["year", "yearly", "annual"],
    "month": ["month", "monthly"],
}

LAST_N_RE = re.compile(r"last\s+(\d+)\s*(year|quarter|month)s?", re.IGNORECASE)


@dataclass
class DashboardSpec:
    metric_col: Optional[str]
    agg_func: str
    dimension_col: Optional[str]
    date_col: Optional[str]
    time_granularity: Optional[str]
    chart_type: str
    last_n: Optional[tuple[int, str]]
    raw_command: str
    warnings: list = field(default_factory=list)


def find_best_column(tokens: list[str], candidates: list[str], phrase: str) -> Optional[str]:
    """Fuzzy-match column names against the command text."""
    phrase_low = phrase.lower()
    # 1. exact substring match on the underscored column name
    for col in sorted(candidates, key=len, reverse=True):
        if col.replace("_", " ") in phrase_low or col in phrase_low:
            return col
    # 2. fuzzy match against individual tokens
    for tok in tokens:
        match = get_close_matches(tok, candidates, n=1, cutoff=0.75)
        if match:
            return match[0]
    return None


def parse_command(command: str, df: pd.DataFrame, type_log: dict) -> DashboardSpec:
    phrase = command.lower()
    tokens = re.findall(r"[a-z0-9_]+", phrase)
    warnings = []

    numeric_cols = type_log.get("numeric", [])
    text_cols = type_log.get("text", [])
    date_cols = type_log.get("date", [])

    #   aggregation function  
    agg_func = "sum"
    for func, kws in AGG_KEYWORDS.items():
        if any(kw in phrase for kw in kws):
            agg_func = func
            break

    #   chart type
    chart_type = "bar"
    for ctype, kws in CHART_KEYWORDS.items():
        if any(kw in phrase for kw in kws):
            chart_type = ctype
            break

    # metric column (numeric)
    metric_col = find_best_column(tokens, numeric_cols, phrase)
    if metric_col is None and numeric_cols:
        metric_col = numeric_cols[0]
        warnings.append(f"No metric named in command — defaulting to '{metric_col}'.")

    #    dimension column (categorical, after 'by'/'per'/'of each')    
    dimension_col = None
    dim_match = re.search(r"(?:by|per|for each|of each)\s+([a-z0-9_ ]+)", phrase)
    if dim_match:
        dimension_col = find_best_column(dim_match.group(1).split(), text_cols, dim_match.group(1))
    if dimension_col is None:
        dimension_col = find_best_column(tokens, text_cols, phrase)

    #      date column + time granularity  
    date_col = date_cols[0] if date_cols else None
    time_granularity = None
    for gran, kws in TIME_GRANULARITY_KEYWORDS.items():
        if any(kw in phrase for kw in kws):
            time_granularity = gran
            break

    #   last N years/quarters/months
    last_n = None
    m = LAST_N_RE.search(phrase)
    if m:
        last_n = (int(m.group(1)), m.group(2))
        if date_col is None:
            warnings.append("Command mentions a time window but no date column was found — ignoring it.")
            last_n = None

    if metric_col is None:
        warnings.append("No numeric column found in the sheet — cannot build this chart.")

    return DashboardSpec(
        metric_col=metric_col,
        agg_func=agg_func,
        dimension_col=dimension_col,
        date_col=date_col,
        time_granularity=time_granularity,
        chart_type=chart_type,
        last_n=last_n,
        raw_command=command,
        warnings=warnings,
    )


def build_pivot(df: pd.DataFrame, spec: DashboardSpec) -> Optional[pd.DataFrame]:
    if spec.metric_col is None:
        return None
    work = df.copy()

    if spec.last_n and spec.date_col:
        n, unit = spec.last_n
        max_date = work[spec.date_col].max()
        if unit == "year":
            cutoff = max_date - pd.DateOffset(years=n)
        elif unit == "quarter":
            cutoff = max_date - pd.DateOffset(months=3 * n)
        else:
            cutoff = max_date - pd.DateOffset(months=n)
        work = work[work[spec.date_col] >= cutoff]

    time_bucket_col = None
    if spec.time_granularity and spec.date_col:
        if spec.time_granularity == "year":
            work["_time_bucket"] = work[spec.date_col].dt.year.astype(str)
        elif spec.time_granularity == "quarter":
            work["_time_bucket"] = work[spec.date_col].dt.to_period("Q").astype(str)
        else:
            work["_time_bucket"] = work[spec.date_col].dt.to_period("M").astype(str)
        time_bucket_col = "_time_bucket"

    group_cols = [c for c in [time_bucket_col, spec.dimension_col] if c]

    if not group_cols:
        result = work[[spec.metric_col]].agg(spec.agg_func).to_frame().T
        return result

    grouped = work.groupby(group_cols, dropna=True)[spec.metric_col].agg(spec.agg_func).reset_index()

    if time_bucket_col and spec.dimension_col:
        pivot = grouped.pivot(index=time_bucket_col, columns=spec.dimension_col, values=spec.metric_col)
        pivot = pivot.sort_index().reset_index().rename(columns={time_bucket_col: "period"})
        return pivot
    elif time_bucket_col:
        return grouped.sort_values(time_bucket_col).rename(columns={time_bucket_col: "period"})
    else:
        return grouped.sort_values(spec.metric_col, ascending=False).head(20)

# 4. EXCEL OUTPUT

HEADER_FILL = PatternFill("solid", fgColor="1D9E75")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def proper_header(col: str) -> str:
    """Turns an internal snake_case column name into a display-friendly
    proper-cased heading — e.g. 'employee_name' -> 'Employee Name'.
    Mirrors Excel's PROPER() function; only affects what's shown in the
    sheet, not the underlying column name used for parsing/aggregation."""
    return re.sub(r"_", " ", str(col)).strip().title()


def write_df_to_sheet(ws, df: pd.DataFrame, start_row: int = 1, title: str = None) -> int:
    """Writes a DataFrame starting at start_row. Returns the next free row."""
    row = start_row
    if title:
        ws.cell(row=row, column=1, value=title).font = Font(bold=True, size=13)
        row += 2
    if df is None or df.empty:
        ws.cell(row=row, column=1, value="(no data)")
        return row + 2

    for j, col in enumerate(df.columns, start=1):
        cell = ws.cell(row=row, column=j, value=str(col))
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
    row += 1
    for _, record in df.iterrows():
        for j, val in enumerate(record, start=1):
            if pd.isna(val):
                val = None
            elif hasattr(val, "isoformat"):
                val = val.isoformat()
            ws.cell(row=row, column=j, value=val)
        row += 1
    for j, col in enumerate(df.columns, start=1):
        ws.column_dimensions[get_column_letter(j)].width = max(12, len(str(col)) + 2)
    return row + 2


def write_cleaned_sheet(wb: Workbook, sheet_name: str, df: pd.DataFrame):
    ws = wb.create_sheet(sheet_name[:31])
    ws.append([proper_header(c) for c in df.columns])
    for r in dataframe_to_rows(df, index=False, header=False):
        ws.append(r)
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
    for j, col in enumerate(df.columns, start=1):
        ws.column_dimensions[get_column_letter(j)].width = max(12, len(str(col)) + 2)


def write_eda_sheet(wb: Workbook, sheet_label: str, eda: dict):
    ws = wb.create_sheet(f"EDA_{sheet_label}"[:31])
    row = 1
    ws.cell(row=row, column=1, value=f"EDA report: {sheet_label}").font = Font(bold=True, size=14)
    row += 2
    ws.cell(row=row, column=1, value=f"Shape: {eda['shape'][0]} rows x {eda['shape'][1]} columns")
    row += 2

    row = write_df_to_sheet(ws, eda["missing"], row, "Missing values")
    if eda["numeric_summary"] is not None:
        row = write_df_to_sheet(ws, eda["numeric_summary"], row, "Numeric summary (incl. IQR outlier count)")
    if eda["correlation"] is not None:
        row = write_df_to_sheet(ws, eda["correlation"].reset_index().rename(columns={"index": "column"}), row, "Correlation matrix")
    for col, counts in eda["categorical_top_values"].items():
        cdf = counts.rename("count").reset_index().rename(columns={"index": col})
        row = write_df_to_sheet(ws, cdf, row, f"Top values: {col}")


def add_native_chart(ws, data_start_row: int, data_end_row: int, n_cols: int, chart_type: str, title: str, anchor: str):
    if chart_type == "line":
        chart = LineChart()
    elif chart_type == "pie":
        chart = PieChart()
    elif chart_type == "scatter":
        chart = ScatterChart()
    else:
        chart = BarChart()
        chart.type = "col"

    chart.title = title
    chart.height, chart.width = 10, 20

    cats = Reference(ws, min_col=1, min_row=data_start_row, max_row=data_end_row)
    if chart_type == "pie":
        data = Reference(ws, min_col=2, min_row=data_start_row, max_row=data_end_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
    else:
        for col in range(2, n_cols + 1):
            data = Reference(ws, min_col=col, min_row=data_start_row, max_row=data_end_row)
            chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)

    ws.add_chart(chart, anchor)


def write_dashboard_entry(wb: Workbook, spec: DashboardSpec, pivot: pd.DataFrame, entry_idx: int):
    if "Dashboard" in wb.sheetnames:
        ws = wb["Dashboard"]
    else:
        ws = wb.create_sheet("Dashboard")

    start_row = 1 + entry_idx * 25
    ws.cell(row=start_row, column=1, value=f"Command: {spec.raw_command}").font = Font(bold=True, italic=True)
    if spec.warnings:
        ws.cell(row=start_row + 1, column=1, value="Note: " + " | ".join(spec.warnings)).font = Font(size=9, color="A32D2D")

    table_start = start_row + 2
    next_row = write_df_to_sheet(ws, pivot, table_start)
    data_end_row = next_row - 2

    add_native_chart(
        ws,
        data_start_row=table_start,
        data_end_row=data_end_row,
        n_cols=pivot.shape[1],
        chart_type=spec.chart_type,
        title=spec.raw_command[:60],
        anchor=f"H{start_row}",
    )

# 5. MAIN PIPELINE

def run_pipeline(input_path: str, output_path: str, commands: list[str] | None = None):
    print(f"Loading '{input_path}' ...")
    sheets = load_all_sheets(input_path)

    out_wb = Workbook()
    out_wb.remove(out_wb.active)

    cleaned_sheets = {}
    type_logs = {}

    for name, df in sheets.items():
        cleaned, report = clean_sheet(df)
        cleaned_sheets[name] = cleaned
        type_logs[name] = report.get("type_log", {})
        write_cleaned_sheet(out_wb, f"Clean_{name}", cleaned)
        print(f"  Cleaned '{name}': {report['rows_in']} -> {report['rows_out']} rows, "
              f"{report.get('dup_rows_removed', 0)} duplicates removed")

        if not cleaned.empty:
            eda = build_eda(cleaned, type_logs[name])
            write_eda_sheet(out_wb, name, eda)

    # pick the largest cleaned sheet as the default target for dashboard commands
    target_name = max(cleaned_sheets, key=lambda n: len(cleaned_sheets[n])) if cleaned_sheets else None
    target_df = cleaned_sheets.get(target_name)
    target_types = type_logs.get(target_name, {})

    if commands is None:
        # ask the user before building anything,
        # and keep asking after each chart until they say no.
        chart_idx = 0
        while True:
            try:
                answer = input("\nDo you want to build a chart with this data? (yes/no): ").strip().lower()
            except EOFError:
                break
            if answer not in ("yes", "y"):
                print("Okay — no charts will be built. Saving the cleaned data and EDA report as-is.")
                break

            if target_df is None or target_df.empty:
                print("No usable data was found to chart — stopping here.")
                break

            try:
                cmd = input(
                    "Give me a clear, correctly typed chart/statement you want to "
                    "build from this data: "
                ).strip()
            except EOFError:
                break
            if not cmd:
                print("That was empty — try again.")
                continue

            spec = parse_command(cmd, target_df, target_types)
            pivot = build_pivot(target_df, spec)
            if pivot is None or pivot.empty:
                print(f"  Could not build a chart for: '{cmd}' — try rephrasing it.")
                continue

            write_dashboard_entry(out_wb, spec, pivot, chart_idx)
            chart_idx += 1
            print(f"  Added chart for: '{cmd}' ({spec.chart_type}, {spec.agg_func} of {spec.metric_col})")
            for w in spec.warnings:
                print(f"    note: {w}")
            # loop back and ask "build another chart?" again
    else:
        # Non-interactive / scripted mode (e.g. --commands from the CLI):
        # process the given list directly, no prompting.
        for i, cmd in enumerate(commands):
            if target_df is None or target_df.empty:
                print(f"  Skipping '{cmd}': no usable data.")
                continue
            spec = parse_command(cmd, target_df, target_types)
            pivot = build_pivot(target_df, spec)
            if pivot is None or pivot.empty:
                print(f"  Could not build a chart for: '{cmd}'")
                continue
            write_dashboard_entry(out_wb, spec, pivot, i)
            print(f"  Added chart for: '{cmd}' ({spec.chart_type}, {spec.agg_func} of {spec.metric_col})")
            for w in spec.warnings:
                print(f"    note: {w}")

    out_wb.save(output_path)
    print(f"\nSaved -> {output_path}")


def main():
    parser = argparse.ArgumentParser(description="Generic Excel cleaning, EDA, and NLP-driven dashboard tool.")
    parser.add_argument("input", help="Path to the input .xlsx file")
    parser.add_argument("output", help="Path to write the output .xlsx file")
    parser.add_argument("--commands", nargs="*", default=None,
                         help="One or more dashboard commands to run non-interactively.")
    args = parser.parse_args()
    run_pipeline(args.input, args.output, args.commands)


if __name__ == "__main__":
    _launcher = sys.argv[0] if sys.argv else ""
    if "ipykernel" in _launcher or "colab_kernel_launcher" in _launcher:
        print(" Previous files are cuurently in work, delete them as new files are unable to overwite them"
        )
    else:
        main()
        