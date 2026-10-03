"""Reading Excel files, exact money/date parsing, column matching, joins and filters."""
import re
import unicodedata
from bisect import bisect_right
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from io import BytesIO

import pandas as pd

# ------------------------------------------------------------ reading files
def sheet_names(uploaded):
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        return ["(CSV file)"]
    uploaded.seek(0)
    return pd.ExcelFile(uploaded).sheet_names


def _cell_to_text(v):
    """Turn one Excel cell into exact text. Dates -> YYYY-MM-DD, 1001.0 -> 1001."""
    if v is None:
        return ""
    if isinstance(v, float):
        if pd.isna(v):
            return ""
        if v.is_integer():
            return str(int(v))
        return repr(v)  # shortest exact round-trip, e.g. 1234.56
    if isinstance(v, (pd.Timestamp, datetime)):
        if pd.isna(v):
            return ""
        return v.strftime("%Y-%m-%d")
    if isinstance(v, date):
        return v.strftime("%Y-%m-%d")
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    return str(v).strip()


def _clean_columns(cols):
    out, seen = [], {}
    for i, c in enumerate(cols):
        c = "" if c is None or (isinstance(c, float) and pd.isna(c)) else str(c).strip()
        if not c or c.lower().startswith("unnamed:"):
            c = f"Column {i + 1}"
        if c in seen:
            seen[c] += 1
            c = f"{c} ({seen[c]})"
        else:
            seen[c] = 1
        out.append(c)
    return out


def read_table(uploaded, sheet, header_row=1):
    """Read a sheet into a DataFrame of text. header_row is 1-based (as shown in Excel)."""
    uploaded.seek(0)
    if uploaded.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded, header=header_row - 1, dtype=str, keep_default_na=False)
    else:
        df = pd.read_excel(uploaded, sheet_name=sheet, header=header_row - 1, dtype=object)
    df.columns = _clean_columns(df.columns)
    df = df.apply(lambda s: s.map(_cell_to_text))
    # drop rows that are completely empty
    df = df[~(df == "").all(axis=1)].reset_index(drop=True)
    return df


# ------------------------------------------------------------ normalising
def norm_name(s):
    """'Supplier  Code ' / 'supplier_code' / 'SupplierCode' all -> 'suppliercode'."""
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def norm_key(v):
    """Normalise a matching key: trims, ignores case/extra spaces, '1001.0' == '1001'."""
    s = re.sub(r"\s+", " ", str(v).strip()).casefold()
    if re.fullmatch(r"-?\d+\.0+", s):
        s = s.split(".")[0]
    return s


def auto_match(column, names):
    """Return the name in `names` that matches `column` ignoring case/spaces/symbols, else None."""
    target = norm_name(column)
    for n in names:
        if norm_name(n) == target:
            return n
    return None


# ------------------------------------------------------------ values
_MONEY_JUNK = re.compile(r"(₹|rs\.?|inr|\$|,|\s)", re.IGNORECASE)


def parse_decimal(v):
    """Exact number from text like '₹ 1,23,456.50', '(500)', 'Rs. 12'. None if not a number."""
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    neg = s.startswith("(") and s.endswith(")")
    if neg:
        s = s[1:-1]
    s = _MONEY_JUNK.sub("", s)
    if s.endswith("-") and s.count("-") == 1:  # '500-' style negatives
        s, neg = s[:-1], True
    try:
        d = Decimal(s)
    except InvalidOperation:
        return None
    if not d.is_finite():
        return None
    return -d if neg else d


_ISO = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def parse_date(v):
    """Date from 'YYYY-MM-DD', 'DD-MM-YYYY', 'DD/MM/YYYY', '5 Jan 2026'... (day first). None if not a date."""
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    try:
        if _ISO.match(s):
            return datetime.strptime(s, "%Y-%m-%d").date()
        if not re.search(r"\d", s):
            return None
        ts = pd.to_datetime(s, dayfirst=True, errors="coerce")
        return None if pd.isna(ts) else ts.date()
    except (ValueError, OverflowError):
        return None


def clean_for_type(value, dtype):
    """Convert text to the stored form for a parameter type.
    Returns (stored_text, problem) - problem is None when the value is valid."""
    s = "" if value is None else str(value).strip()
    if not s:
        return "", None
    if dtype in ("Money", "Number"):
        d = parse_decimal(s)
        if d is None:
            return s, f"'{s}' is not a number"
        return format(d.normalize() if d == d.to_integral() else d, "f"), None
    if dtype == "Date":
        d = parse_date(s)
        if d is None:
            return s, f"'{s}' is not a date"
        return d.strftime("%Y-%m-%d"), None
    return s, None


def indian_group(int_str):
    """'1234567' -> '12,34,567'"""
    neg = int_str.startswith("-")
    s = int_str.lstrip("-")
    if len(s) > 3:
        head, tail = s[:-3], s[-3:]
        head = re.sub(r"(\d)(?=(\d{2})+$)", r"\1,", head)
        s = head + "," + tail
    return ("-" if neg else "") + s


def format_money(d):
    """Decimal -> '12,34,567.50' (always 2 decimals, exact half-up rounding)."""
    if d is None:
        return ""
    q = Decimal(d).quantize(Decimal("0.01"), rounding="ROUND_HALF_UP")
    whole, frac = format(q, "f").split(".")
    return f"{indian_group(whole)}.{frac}"


def format_value(v, dtype):
    if dtype == "Money":
        d = parse_decimal(v)
        return format_money(d) if d is not None else ("" if v is None else str(v))
    if dtype == "Date":
        d = parse_date(v)
        return d.strftime("%d-%m-%Y") if d else ("" if v is None else str(v))
    return "" if v is None else str(v)


def guess_type(name, series):
    """Guess Text/Money/Number/Date for an unknown column."""
    vals = [v for v in series.astype(str) if v.strip()]
    if not vals:
        return "Text"
    sample = vals[:300]
    n = len(sample)
    if sum(1 for v in sample if _ISO.match(v)) >= 0.9 * n:
        return "Date"
    if sum(1 for v in sample if parse_decimal(v) is not None) >= 0.9 * n:
        if re.search(r"amount|amt|value|total|balance|price|payment|paid|due|rs|inr|cost", name, re.I):
            return "Money"
        return "Number"
    return "Text"


# ------------------------------------------------------------ comparing
def _cmp_pair(a, b):
    """Return (a', b') comparable: both Decimal, both date, or both lower-case text."""
    da, db = parse_decimal(a), parse_decimal(b)
    if da is not None and db is not None:
        return da, db
    ta, tb = parse_date(a), parse_date(b)
    if ta is not None and tb is not None:
        return ta, tb
    return str(a).strip().casefold(), str(b).strip().casefold()


def _equal(a, b):
    x, y = _cmp_pair(a, b)
    return x == y


OPERATORS = [
    "equals",
    "does not equal",
    "contains",
    "does not contain",
    "is one of (comma separated)",
    "is empty",
    "is not empty",
    "greater than",
    "greater or equal",
    "less than",
    "less or equal",
    "between",
    "date is in year",
    "equals column",
    "differs from column",
    "greater than column",
    "less than column",
]
NO_VALUE_OPS = {"is empty", "is not empty"}
COLUMN_OPS = {"equals column", "differs from column", "greater than column", "less than column"}


def _test(op, a, v1, v2):
    a = "" if a is None else str(a)
    if op == "is empty":
        return a.strip() == ""
    if op == "is not empty":
        return a.strip() != ""
    if op in ("equals", "equals column"):
        return _equal(a, v1)
    if op in ("does not equal", "differs from column"):
        return not _equal(a, v1)
    if op == "contains":
        return str(v1).strip().casefold() in a.casefold()
    if op == "does not contain":
        return str(v1).strip().casefold() not in a.casefold()
    if op == "is one of (comma separated)":
        return any(_equal(a, x) for x in str(v1).split(",") if x.strip())
    if op == "date is in year":
        d = parse_date(a)
        return d is not None and str(d.year) == str(v1).strip()
    if a.strip() == "" or str(v1).strip() == "":
        return False
    x, y = _cmp_pair(a, v1)
    if type(x) is not type(y):
        return False
    if isinstance(x, str) and (parse_decimal(v1) is not None or parse_date(v1) is not None):
        return False  # comparing with a number/date: cells that are plain text never match
    if op in ("greater than", "greater than column"):
        return x > y
    if op == "greater or equal":
        return x >= y
    if op in ("less than", "less than column"):
        return x < y
    if op == "less or equal":
        return x <= y
    if op == "between":
        if str(v2).strip() == "":
            return False
        x2, y2 = _cmp_pair(a, v2)
        if type(x2) is not type(y2):
            return False
        return y <= x and x <= y2
    return True


def apply_filters(df, filters, match_all=True):
    """filters: list of dicts {col, op, value, value2, other_col}."""
    active = [f for f in filters if f.get("col") in df.columns]
    if not active or df.empty:
        return df
    masks = []
    for f in active:
        col, op = f["col"], f["op"]
        if op in COLUMN_OPS:
            other = f.get("other_col")
            if other not in df.columns:
                continue
            m = [_test(op, a, b, "") for a, b in zip(df[col], df[other])]
        else:
            m = [_test(op, a, f.get("value", ""), f.get("value2", "")) for a in df[col]]
        masks.append(pd.Series(m, index=df.index))
    if not masks:
        return df
    total = masks[0]
    for m in masks[1:]:
        total = (total & m) if match_all else (total | m)
    return df[total]


JOIN_MODES = [
    "Keep all rows (blank if not found here)",
    "Keep only rows found in this file",
    "Keep only rows NOT found in this file",
]


def join_sources(base_df, base_key, joins):
    """Combine the base table with other tables by a key column.

    joins: list of dicts {label, df, key, columns, mode}
    Returns (result_df, warnings list)."""
    warnings = []
    out = base_df.copy()
    base_keys = out[base_key].map(norm_key)

    dup = base_keys[base_keys.duplicated(keep=False) & (base_keys != "")]
    if not dup.empty:
        warnings.append(
            f"Base data has {dup.nunique()} key value(s) appearing more than once in '{base_key}': "
            + ", ".join(sorted(set(out.loc[dup.index, base_key].astype(str)))[:15])
        )
    blanks = int((base_keys == "").sum())
    if blanks:
        warnings.append(f"{blanks} row(s) in the base data have an empty '{base_key}' and cannot be matched.")

    keep = pd.Series(True, index=out.index)
    for j in joins:
        jdf, jkey, label = j["df"], j["key"], j["label"]
        jkeys = jdf[jkey].map(norm_key)
        dups = jkeys[jkeys.duplicated(keep=False) & (jkeys != "")]
        if not dups.empty:
            warnings.append(
                f"'{label}' has {dups.nunique()} key value(s) repeated in '{jkey}' - only the FIRST row "
                "is used for each: " + ", ".join(sorted(set(jdf.loc[dups.index, jkey].astype(str)))[:15])
            )
        lookup = {}
        for idx, k in zip(jdf.index, jkeys):
            if k and k not in lookup:
                lookup[k] = idx
        found = base_keys.map(lambda k: k in lookup)
        missing = int((~found & (base_keys != "")).sum())
        if missing and j["mode"] == JOIN_MODES[0]:
            warnings.append(f"{missing} row(s) were not found in '{label}' (their '{label}' columns are blank).")
        for c in j["columns"]:
            new_name = c if c not in out.columns else f"{c} [{label}]"
            n = 2
            while new_name in out.columns:
                new_name = f"{c} [{label} {n}]"
                n += 1
            out[new_name] = [
                jdf.at[lookup[k], c] if k in lookup else "" for k in base_keys
            ]
        if j["mode"] == JOIN_MODES[1]:
            keep &= found
        elif j["mode"] == JOIN_MODES[2]:
            keep &= ~found
    return out[keep].reset_index(drop=True), warnings


def sort_key(v):
    """Numbers first (by value), then dates, then text (ignoring capitals)."""
    d = parse_decimal(v)
    if d is not None:
        return (0, d, "")
    t = parse_date(v)
    if t is not None:
        return (1, Decimal(t.toordinal()), "")
    return (2, Decimal(0), str(v).casefold())


def sort_df(df, col, descending=False, keep_index=False):
    """keep_index=True keeps each row's original label (so a row can be recognised after sorting)."""
    if not col or col not in df.columns:
        return df
    out = df.sort_values(col, key=lambda s: s.map(sort_key), ascending=not descending, kind="stable")
    return out if keep_index else out.reset_index(drop=True)


# ------------------------------------------------------------ column (Excel-style) filters
def column_filter_mask(series, spec):
    """spec: {"in": [values]} / {"out": [values]} (ticked values, compared as trimmed text)
    plus optional {"op", "v1", "v2"} condition. Both parts must hold."""
    mask = pd.Series(True, index=series.index)
    text = series.astype(str).str.strip()
    if spec.get("in") is not None:
        mask &= text.isin(set(spec["in"]))
    if spec.get("out"):
        mask &= ~text.isin(set(spec["out"]))
    op = spec.get("op")
    if op and op in OPERATORS and op not in COLUMN_OPS:
        mask &= pd.Series([_test(op, a, spec.get("v1", ""), spec.get("v2", "")) for a in series],
                          index=series.index)
    return mask


def apply_column_filters(df, filters):
    """filters: {column: spec}. Columns that no longer exist are ignored."""
    mask = pd.Series(True, index=df.index)
    for col, spec in filters.items():
        if col in df.columns and spec:
            mask &= column_filter_mask(df[col], spec)
    return df[mask]


def distinct_values(df, filters, limit=1000):
    """For each column, the sorted values left after the filters on the OTHER columns
    (like Excel's filter list). Returns {column: (values, more_than_limit)}."""
    masks = {c: column_filter_mask(df[c], s) for c, s in filters.items() if c in df.columns and s}
    out = {}
    for c in df.columns:
        m = pd.Series(True, index=df.index)
        for oc, om in masks.items():
            if oc != c:
                m &= om
        vals = sorted(set(df.loc[m, c].astype(str).str.strip()), key=sort_key)
        out[c] = (vals[:limit], len(vals) > limit)
    return out


MATCH_HOWS = [
    "the values are the same",
    "the other cell contains this value",
    "either one contains the other",
]


def norm_words(v):
    """'ABC Pharma-Chem  GmbH,\\nMunich' -> 'abc pharma chem gmbh munich'
    (no capitals, accents or punctuation; single spaces)."""
    s = unicodedata.normalize("NFKD", str(v))
    s = "".join(c for c in s if not unicodedata.combining(c)).casefold()
    return re.sub(r"[\W_]+", " ", s).strip()


def match_rows(series, other_series, how=MATCH_HOWS[0]):
    """For each value of `series`, the index label of the FIRST row of `other_series` it
    matches, or None. Blank values never match.

    - same: equal after trimming, ignoring capitals and extra spaces ('1001.0' == '1001')
    - contains: this value appears as whole words inside a cell of the other column
      (e.g. 'ABC Pharma' is found in 'ABC Pharma GmbH, Munich')
    - either: as above, or a cell of the other column appears inside this value"""
    exact = {}
    for idx, v in other_series.items():
        k = norm_key(v)
        if k and k not in exact:
            exact[k] = idx
    if how == MATCH_HOWS[0]:
        return pd.Series([exact.get(norm_key(v)) for v in series], index=series.index, dtype=object)

    # all cells of the other column in one text, to find a value inside any of them quickly
    parts, starts, labels, whole = [], [], [], {}
    pos = 0
    for idx, v in other_series.items():
        w = norm_words(v)
        if not w:
            continue
        seg = f" {w} \n"
        parts.append(seg)
        starts.append(pos)
        labels.append(idx)
        pos += len(seg)
        if len(w) >= 3 and w not in whole:  # skip tiny values like 'ag' or 'k'
            whole[w] = idx
    haystack = "".join(parts)

    def one(v):
        w = norm_words(v)
        if not w:
            return None
        k = norm_key(v)
        if k in exact:
            return exact[k]
        p = haystack.find(f" {w} ")
        if p >= 0:
            return labels[bisect_right(starts, p) - 1]
        if how == MATCH_HOWS[2]:
            words = w.split()
            for i in range(len(words)):
                for j in range(len(words), i, -1):  # longest piece first
                    hit = whole.get(" ".join(words[i:j]))
                    if hit is not None:
                        return hit
        return None

    return pd.Series([one(v) for v in series], index=series.index, dtype=object)


def found_in(series, other_series, how=MATCH_HOWS[0]):
    """True for each value of `series` that is found in `other_series` (see match_rows)."""
    return match_rows(series, other_series, how).notna()


def column_total(series):
    """Exact Decimal total of a column. Returns (total, count_of_non_numeric_values)."""
    total, bad = Decimal(0), 0
    for v in series:
        if str(v).strip() == "":
            continue
        d = parse_decimal(v)
        if d is None:
            bad += 1
        else:
            total += d
    return total, bad


def to_excel_bytes(df, types, sheet_name="Accounts"):
    """Excel download. Money/Number columns are written as real numbers, dates as dates."""
    out = pd.DataFrame(index=df.index)
    for c in df.columns:
        t = types.get(c, "Text")
        if t in ("Money", "Number"):
            out[c] = [float(parse_decimal(v)) if parse_decimal(v) is not None else (v or None) for v in df[c]]
        elif t == "Date":
            out[c] = [parse_date(v) or (v or None) for v in df[c]]
        else:
            out[c] = df[c]
    sheet_name = re.sub(r"[\[\]:*?/\\]", "-", str(sheet_name))[:31].strip() or "Sheet1"
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl", date_format="DD-MM-YYYY") as xw:
        out.to_excel(xw, index=False, sheet_name=sheet_name)
        ws = xw.sheets[sheet_name]
        for i, c in enumerate(out.columns, start=1):
            width = max([len(str(c))] + [len(str(v)) for v in df[c].head(500)]) + 2
            ws.column_dimensions[ws.cell(1, i).column_letter].width = min(width, 45)
            if types.get(c) == "Money":
                for row in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                    for cell in row:
                        cell.number_format = "#,##0.00"
    return buf.getvalue()
