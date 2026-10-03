"""SQLite storage for the accounts app.

Everything is kept in one file (accounts.db) next to this script, so a backup
is just a copy of that file.

Parameters (columns of the main accounts sheet) are user-defined, so account
values are stored one row per (account, parameter). Adding or removing a
parameter never needs a schema change. All values are stored as TEXT so money
keeps its exact digits (no float rounding).
"""
import json
import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd

DB_PATH = Path(__file__).with_name("accounts.db")

PARAM_TYPES = ["Text", "Money", "Number", "Date"]

DEFAULT_PARAMETERS = [
    ("Supplier Code", "Text"),
    ("Supplier Name", "Text"),
    ("Account Start", "Date"),
    ("Account End", "Date"),
    ("Status", "Text"),
    ("Amount", "Money"),
]


def connect():
    con = sqlite3.connect(DB_PATH, check_same_thread=False)
    con.execute("PRAGMA foreign_keys = ON")
    return con


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def init_db():
    with connect() as con:
        con.executescript(
            """
            CREATE TABLE IF NOT EXISTS parameters (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL UNIQUE COLLATE NOCASE,
                dtype TEXT NOT NULL DEFAULT 'Text',
                position INTEGER NOT NULL DEFAULT 0
            );
            CREATE TABLE IF NOT EXISTS accounts (
                id INTEGER PRIMARY KEY,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS account_values (
                account_id INTEGER NOT NULL REFERENCES accounts(id) ON DELETE CASCADE,
                parameter_id INTEGER NOT NULL REFERENCES parameters(id) ON DELETE CASCADE,
                value TEXT,
                PRIMARY KEY (account_id, parameter_id)
            );
            CREATE TABLE IF NOT EXISTS sources (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL UNIQUE COLLATE NOCASE,
                file_name TEXT,
                sheet_name TEXT,
                columns_json TEXT NOT NULL,
                rows_json TEXT NOT NULL,
                uploaded_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            );
            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY,
                ts TEXT NOT NULL,
                action TEXT NOT NULL,
                detail TEXT
            );
            """
        )
        if con.execute("SELECT COUNT(*) FROM parameters").fetchone()[0] == 0:
            for i, (name, dtype) in enumerate(DEFAULT_PARAMETERS):
                con.execute(
                    "INSERT INTO parameters(name, dtype, position) VALUES (?,?,?)",
                    (name, dtype, i),
                )


def log(action, detail=""):
    with connect() as con:
        con.execute(
            "INSERT INTO audit_log(ts, action, detail) VALUES (?,?,?)",
            (now(), action, detail),
        )


def get_log(limit=200):
    with connect() as con:
        return pd.read_sql_query(
            "SELECT ts AS Time, action AS Action, detail AS Detail "
            "FROM audit_log ORDER BY id DESC LIMIT ?",
            con,
            params=(limit,),
        )


# ---------------------------------------------------------------- settings
def get_setting(key, default=""):
    with connect() as con:
        row = con.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return row[0] if row else default


def set_setting(key, value):
    with connect() as con:
        con.execute(
            "INSERT INTO settings(key, value) VALUES (?,?) "
            "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, value),
        )


# -------------------------------------------------------------- parameters
def get_parameters():
    """List of dicts: id, name, dtype, position (in display order)."""
    with connect() as con:
        rows = con.execute(
            "SELECT id, name, dtype, position FROM parameters ORDER BY position, id"
        ).fetchall()
    return [dict(id=r[0], name=r[1], dtype=r[2], position=r[3]) for r in rows]


def add_parameter(name, dtype="Text"):
    name = name.strip()
    if not name:
        raise ValueError("Parameter name cannot be empty.")
    with connect() as con:
        exists = con.execute(
            "SELECT id FROM parameters WHERE name=? COLLATE NOCASE", (name,)
        ).fetchone()
        if exists:
            raise ValueError(f"A parameter called '{name}' already exists.")
        pos = con.execute("SELECT COALESCE(MAX(position), -1) + 1 FROM parameters").fetchone()[0]
        cur = con.execute(
            "INSERT INTO parameters(name, dtype, position) VALUES (?,?,?)",
            (name, dtype, pos),
        )
        pid = cur.lastrowid
    log("Parameter added", f"{name} ({dtype})")
    return pid


def update_parameter(pid, name, dtype):
    name = name.strip()
    if not name:
        raise ValueError("Parameter name cannot be empty.")
    with connect() as con:
        clash = con.execute(
            "SELECT id FROM parameters WHERE name=? COLLATE NOCASE AND id<>?", (name, pid)
        ).fetchone()
        if clash:
            raise ValueError(f"A parameter called '{name}' already exists.")
        con.execute("UPDATE parameters SET name=?, dtype=? WHERE id=?", (name, dtype, pid))
    log("Parameter changed", f"id {pid} -> {name} ({dtype})")


def move_parameter(pid, direction):
    params = get_parameters()
    ids = [p["id"] for p in params]
    i = ids.index(pid)
    j = i + direction
    if j < 0 or j >= len(ids):
        return
    ids[i], ids[j] = ids[j], ids[i]
    with connect() as con:
        for pos, x in enumerate(ids):
            con.execute("UPDATE parameters SET position=? WHERE id=?", (pos, x))


def delete_parameter(pid):
    with connect() as con:
        name = con.execute("SELECT name FROM parameters WHERE id=?", (pid,)).fetchone()
        con.execute("DELETE FROM parameters WHERE id=?", (pid,))
    log("Parameter deleted", name[0] if name else str(pid))


# ---------------------------------------------------------------- accounts
def get_accounts_df():
    """Main accounts as a DataFrame: 'ID' column + one column per parameter (text values)."""
    params = get_parameters()
    with connect() as con:
        accounts = con.execute("SELECT id FROM accounts ORDER BY id").fetchall()
        vals = con.execute("SELECT account_id, parameter_id, value FROM account_values").fetchall()
    by_acc = {}
    for aid, pid, v in vals:
        by_acc.setdefault(aid, {})[pid] = v
    rows = []
    for (aid,) in accounts:
        rec = {"ID": aid}
        d = by_acc.get(aid, {})
        for p in params:
            rec[p["name"]] = d.get(p["id"]) or ""
        rows.append(rec)
    cols = ["ID"] + [p["name"] for p in params]
    return pd.DataFrame(rows, columns=cols)


def add_account(values_by_pid):
    """values_by_pid: {parameter_id: text value}. Returns new account id."""
    ts = now()
    with connect() as con:
        aid = con.execute(
            "INSERT INTO accounts(created_at, updated_at) VALUES (?,?)", (ts, ts)
        ).lastrowid
        for pid, v in values_by_pid.items():
            con.execute(
                "INSERT INTO account_values(account_id, parameter_id, value) VALUES (?,?,?)",
                (aid, pid, v),
            )
    return aid


def update_account(aid, values_by_pid):
    with connect() as con:
        con.execute("UPDATE accounts SET updated_at=? WHERE id=?", (now(), aid))
        for pid, v in values_by_pid.items():
            con.execute(
                "INSERT INTO account_values(account_id, parameter_id, value) VALUES (?,?,?) "
                "ON CONFLICT(account_id, parameter_id) DO UPDATE SET value=excluded.value",
                (aid, pid, v),
            )


def save_accounts_bulk(items):
    """items: list of (account_id or None, {parameter_id: value}) saved in ONE transaction,
    so an import either fully succeeds or changes nothing.
    None -> new account. Returns (added, updated)."""
    ts = now()
    added = updated = 0
    con = connect()
    try:
        with con:
            for aid, vals in items:
                if aid is None:
                    aid = con.execute(
                        "INSERT INTO accounts(created_at, updated_at) VALUES (?,?)", (ts, ts)
                    ).lastrowid
                    added += 1
                else:
                    con.execute("UPDATE accounts SET updated_at=? WHERE id=?", (ts, aid))
                    updated += 1
                con.executemany(
                    "INSERT INTO account_values(account_id, parameter_id, value) VALUES (?,?,?) "
                    "ON CONFLICT(account_id, parameter_id) DO UPDATE SET value=excluded.value",
                    [(aid, pid, v) for pid, v in vals.items()],
                )
    finally:
        con.close()
    return added, updated


def set_parameter_order(ids):
    with connect() as con:
        for pos, pid in enumerate(ids):
            con.execute("UPDATE parameters SET position=? WHERE id=?", (pos, pid))


def delete_accounts(ids):
    with connect() as con:
        con.executemany("DELETE FROM accounts WHERE id=?", [(i,) for i in ids])
    log("Accounts deleted", f"{len(ids)} record(s): IDs {', '.join(map(str, ids))}")


# ----------------------------------------------------------------- sources
def save_source(name, file_name, sheet_name, df):
    """Store an uploaded Excel sheet as-is (all cells as text). Replaces a source with the same name."""
    df = df.fillna("").astype(str)
    cols = list(df.columns)
    rows = df.values.tolist()
    with connect() as con:
        con.execute("DELETE FROM sources WHERE name=? COLLATE NOCASE", (name,))
        con.execute(
            "INSERT INTO sources(name, file_name, sheet_name, columns_json, rows_json, uploaded_at) "
            "VALUES (?,?,?,?,?,?)",
            (name, file_name, sheet_name, json.dumps(cols), json.dumps(rows), now()),
        )
    log("Excel file saved", f"{name}: {file_name} / {sheet_name}, {len(rows)} rows")


def list_sources():
    with connect() as con:
        rows = con.execute(
            "SELECT id, name, file_name, sheet_name, columns_json, rows_json, uploaded_at "
            "FROM sources ORDER BY name"
        ).fetchall()
    return [
        dict(
            id=r[0],
            name=r[1],
            file_name=r[2],
            sheet_name=r[3],
            n_cols=len(json.loads(r[4])),
            n_rows=len(json.loads(r[5])),
            uploaded_at=r[6],
        )
        for r in rows
    ]


def get_source_df(name):
    with connect() as con:
        row = con.execute(
            "SELECT columns_json, rows_json FROM sources WHERE name=? COLLATE NOCASE", (name,)
        ).fetchone()
    if not row:
        return pd.DataFrame()
    return pd.DataFrame(json.loads(row[1]), columns=json.loads(row[0]))


def delete_source(name):
    with connect() as con:
        con.execute("DELETE FROM sources WHERE name=? COLLATE NOCASE", (name,))
    log("Excel file deleted", name)
