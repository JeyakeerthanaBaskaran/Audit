"""Supplier Accounts - year-end account builder.

Run with:  streamlit run app.py
"""
from datetime import date
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

import db
import utils as U
from filter_grid import filter_grid
from pdf_export import build_pdf

st.set_page_config(page_title="Supplier Accounts", page_icon="📒", layout="wide")
db.init_db()

SKIP = "— Don't import —"
NEW = "+ Create new parameter"
MAIN = "Main accounts"


# =====================================================================
# Shared pieces
# =====================================================================
def pick_table(prefix, label="Choose an Excel file"):
    """File uploader + sheet + header row. Returns (df, file_name, sheet) or None."""
    up = st.file_uploader(label, type=["xlsx", "xlsm", "xls", "csv"], key=f"{prefix}_file")
    if not up:
        return None
    try:
        sheets = U.sheet_names(up)
    except Exception as e:  # noqa: BLE001
        st.error(f"This file could not be opened: {e}")
        return None
    c1, c2 = st.columns([3, 1])
    sheet = c1.selectbox("Sheet", sheets, key=f"{prefix}_sheet")
    header_row = c2.number_input(
        "Column names are in row", min_value=1, max_value=50, value=1, step=1,
        key=f"{prefix}_hdr", help="The Excel row number that holds the column headings.",
    )
    try:
        df = U.read_table(up, sheet, int(header_row))
    except Exception as e:  # noqa: BLE001
        st.error(f"This sheet could not be read: {e}")
        return None
    return df, up.name, sheet


def named_bytes(data, name):
    """File contents as a file-like object with a .name, like an uploaded file."""
    f = BytesIO(data)
    f.name = name
    return f


@st.cache_data(show_spinner="Reading sheet…", max_entries=100)
def read_sheet(data, file_name, sheet, header_row):
    return U.read_table(named_bytes(data, file_name), sheet, header_row)


@st.cache_data(show_spinner=False, max_entries=50)
def load_sheet(name, saved_at, n_rows):
    """A saved sheet. saved_at/n_rows make the cache reload it after every save."""
    return db.get_source_df(name)


def mapping_ui(df, prefix):
    """One row per column of `df`: choose which parameter it goes into.
    Returns {column: ("existing", name) | ("new", name, type)}."""
    params = db.get_parameters()
    names = [p["name"] for p in params]
    options = [SKIP, NEW] + names
    tag = f"{prefix}_{abs(hash(tuple(df.columns))) % 10**8}"

    h = st.columns([3, 3, 2, 2])
    h[0].markdown("**Column in the file**")
    h[1].markdown("**Goes into parameter**")
    h[2].markdown("**New parameter name**")
    h[3].markdown("**Type**")
    mapping = {}
    for i, col in enumerate(df.columns):
        c = st.columns([3, 3, 2, 2])
        sample = [v for v in df[col].head(20) if str(v).strip()][:3]
        c[0].markdown(f"{col}")
        c[0].caption("e.g. " + ", ".join(map(str, sample)) if sample else "(empty column)")
        match = U.auto_match(col, names)
        choice = c[1].selectbox(
            "Goes into", options, index=options.index(match) if match else 0,
            key=f"{tag}_map_{i}", label_visibility="collapsed",
        )
        if choice == NEW:
            nm = c[2].text_input("Name", value=col, key=f"{tag}_nm_{i}", label_visibility="collapsed")
            guessed = U.guess_type(col, df[col])
            ty = c[3].selectbox(
                "Type", db.PARAM_TYPES, index=db.PARAM_TYPES.index(guessed),
                key=f"{tag}_ty_{i}", label_visibility="collapsed",
            )
            mapping[col] = ("new", nm.strip(), ty)
        elif choice != SKIP:
            mapping[col] = ("existing", choice)
            c[3].caption(next(p["dtype"] for p in params if p["name"] == choice))
    return mapping


def check_mapping(mapping):
    """Returns a list of error messages for a mapping."""
    errors = []
    existing = {p["name"].casefold() for p in db.get_parameters()}
    targets = {}
    for col, m in mapping.items():
        name = m[1]
        if not name:
            errors.append(f"Column '{col}': give the new parameter a name.")
            continue
        if m[0] == "new" and name.casefold() in existing:
            errors.append(f"Column '{col}': a parameter called '{name}' already exists - choose it from the list instead.")
        targets.setdefault(name.casefold(), []).append(col)
    for name, cols in targets.items():
        if len(cols) > 1:
            errors.append(f"Columns {', '.join(repr(c) for c in cols)} all go into the same parameter. Pick one.")
    return errors


def prepare_import(df, mapping):
    """Validate every cell. Returns (rows, problems) where rows = [{param name: stored text}]."""
    params = {p["name"].casefold(): p for p in db.get_parameters()}
    targets = {}
    for col, m in mapping.items():
        if m[0] == "existing":
            p = params[m[1].casefold()]
            targets[col] = (p["name"], p["dtype"])
        else:
            targets[col] = (m[1], m[2])
    rows, problems = [], []
    for n, (_, r) in enumerate(df.iterrows(), start=1):
        rec = {}
        for col, (name, dtype) in targets.items():
            v, prob = U.clean_for_type(r[col], dtype)
            if prob:
                problems.append({"Data row": n, "Column": col, "Parameter": name, "Problem": prob})
            rec[name] = v
        if any(v != "" for v in rec.values()):
            rows.append(rec)
    return rows, problems


def commit_import(rows, mapping, update_key=None, source_label=""):
    """Create any new parameters, then add/update accounts in one transaction."""
    existing = {p["name"].casefold() for p in db.get_parameters()}
    for m in mapping.values():
        if m[0] == "new" and m[1].casefold() not in existing:
            db.add_parameter(m[1], m[2])
            existing.add(m[1].casefold())
    pid = {p["name"].casefold(): p["id"] for p in db.get_parameters()}

    index, dup_keys = {}, set()
    if update_key:
        acc = db.get_accounts_df()
        if update_key in acc.columns:
            for aid, v in zip(acc["ID"], acc[update_key]):
                k = U.norm_key(v)
                if not k:
                    continue
                if k in index:
                    dup_keys.add(v)
                else:
                    index[k] = aid

    items, pending_new = [], {}
    for rec in rows:
        vals = {pid[n.casefold()]: v for n, v in rec.items()}
        if update_key:
            k = U.norm_key(rec.get(update_key, ""))
            if k and k in index:
                # blanks in the file never wipe out values already stored
                items.append((index[k], {p: v for p, v in vals.items() if v != ""}))
                continue
            if k and k in pending_new:
                # same key twice in the file: merge into the row being added
                pending_new[k].update({p: v for p, v in vals.items() if v != ""})
                continue
            if k:
                pending_new[k] = vals
        items.append((None, vals))
    added, updated = db.save_accounts_bulk(items)
    db.log("Accounts imported", f"{source_label}: {added} added, {updated} updated"
           + (f" (matched on {update_key})" if update_key else ""))
    return added, updated, dup_keys


def import_panel(df, prefix, source_label):
    """Mapping + validation + import button for a DataFrame going into Main accounts."""
    st.markdown("##### Match the columns")
    st.caption("Columns whose names match an existing parameter are matched automatically "
               "(ignoring capitals, spaces and symbols). Change any that are wrong.")
    mapping = mapping_ui(df, prefix)
    if not mapping:
        st.info("Choose at least one column to import.")
        return
    errors = check_mapping(mapping)
    for e in errors:
        st.error(e)
    if errors:
        return

    st.markdown("##### How to add them")
    target_names = [m[1] for m in mapping.values()]
    mode = st.radio(
        "Import mode",
        ["Add every row as a new account",
         "Update accounts that already exist (matched by a key), add the rest"],
        key=f"{prefix}_mode", label_visibility="collapsed",
    )
    update_key = None
    if mode.startswith("Update"):
        default = next((i for i, n in enumerate(target_names) if "code" in n.lower()), 0)
        update_key = st.selectbox("Match existing accounts using", target_names, index=default, key=f"{prefix}_key")
        st.caption("Blank cells in the file never overwrite values already saved.")

    rows, problems = prepare_import(df, mapping)
    st.write(f"**{len(rows)}** row(s) ready to import.")
    allow = True
    if problems:
        st.warning(f"{len(problems)} cell(s) don't match their parameter type. "
                   "Fix them in Excel and upload again, or import them as plain text.")
        st.dataframe(pd.DataFrame(problems).head(200), hide_index=True, width="stretch")
        allow = st.checkbox("Import anyway and keep these cells exactly as written", key=f"{prefix}_allow")
    if st.button("Import into Main accounts", type="primary", disabled=not allow or not rows, key=f"{prefix}_go"):
        added, updated, dup = commit_import(rows, mapping, update_key, source_label)
        st.success(f"Done: {added} account(s) added, {updated} updated.")
        if dup:
            st.warning("These keys appear more than once in Main accounts, so only the first was updated: "
                       + ", ".join(sorted(map(str, dup))[:20]))


def col_letter(i):
    """0 -> A, 25 -> Z, 26 -> AA (like Excel)."""
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def excel_config(columns, key_col=None):
    """Column headings like Excel: 'A · Supplier Code'. The key column gets a 🔑."""
    cfg = {"_index": st.column_config.Column(label="#", width="small")}
    for i, c in enumerate(columns):
        cfg[c] = st.column_config.Column(label=f"{col_letter(i)} · {c}" + (" 🔑" if c == key_col else ""))
    return cfg


def excel_grid(df, key, key_col=None, height=320):
    """Show a sheet like Excel (row numbers 1, 2, 3… and column letters).
    Click column headings to select them. Returns the selected column names."""
    view = df.copy()
    view.index = range(1, len(view) + 1)
    event = st.dataframe(
        view,
        key=key,
        on_select="rerun",
        selection_mode="multi-column",
        column_config=excel_config(df.columns, key_col),
        width="stretch",
        height=min(height, 38 + 35 * max(len(view), 1)),
    )
    chosen = set(event.selection.columns)
    return [c for c in df.columns if c in chosen]  # keep sheet order


def money_totals_caption(df, types):
    parts = []
    for c, t in types.items():
        if t == "Money" and c in df.columns:
            total, bad = U.column_total(df[c])
            parts.append(f"{c}: **{U.format_money(total)}**" + (f" ({bad} non-number cell(s) left out)" if bad else ""))
    if parts:
        st.markdown("Totals of shown rows — " + " · ".join(parts))


# =====================================================================
# Sidebar
# =====================================================================
with st.sidebar:
    st.title("📒 Supplier Accounts")
    company = st.text_input("Company name (printed on the PDF)", value=db.get_setting("company"))
    if company != db.get_setting("company"):
        db.set_setting("company", company)
    page = st.radio(
        "Go to",
        ["Excel files", "Filter & compare", "Parameters", "Main accounts", "Year-end sheet",
         "History & backup"],
    )
    st.divider()
    srcs = db.list_sources()
    st.caption(f"{len(srcs)} Excel file(s) saved · {len(db.get_accounts_df())} account(s) · "
               f"{len(db.get_parameters())} parameter(s)")


# =====================================================================
# Excel files
# =====================================================================
if page == "Excel files":
    st.header("Excel files")
    st.write("Upload all the Excel files you work with — you can pick several at once, and every "
             "sheet inside them is listed. They are saved as-is, so you can filter and compare them "
             "on the **Filter & compare** page. Saving a sheet under a name that is already used "
             "replaces the old copy.")

    st.session_state.setdefault("up_gen", 0)
    ups = st.file_uploader("Choose Excel files", type=["xlsx", "xlsm", "xls", "csv"],
                           accept_multiple_files=True, key=f"up_files_{st.session_state.up_gen}")
    if ups:
        files = {up.name: up.getvalue() for up in ups}
        plan_rows = []
        for fname, data in files.items():
            try:
                sheets = U.sheet_names(named_bytes(data, fname))
            except Exception as e:  # noqa: BLE001
                st.error(f"'{fname}' could not be opened: {e}")
                continue
            stem = Path(fname).stem
            for sh in sheets:
                plan_rows.append({
                    "Save": True, "File": fname, "Sheet": sh,
                    "Save as": stem if len(sheets) == 1 else f"{stem} - {sh}",
                    "Headings in row": 1,
                })
        if plan_rows:
            st.caption("Untick sheets you don't need. Change **Save as** to give a sheet a shorter name, "
                       "and **Headings in row** if the column names are not in the first row.")
            plan = st.data_editor(
                pd.DataFrame(plan_rows), hide_index=True, width="stretch", disabled=["File", "Sheet"],
                column_config={
                    "Save": st.column_config.CheckboxColumn(width="small"),
                    "Headings in row": st.column_config.NumberColumn(min_value=1, max_value=50, step=1),
                },
                key=f"up_plan_{st.session_state.up_gen}_{abs(hash(tuple(files))) % 10**8}",
            )
            todo = plan[plan["Save"]]
            taken = {s["name"].casefold() for s in srcs}
            problems, ready = [], []
            names = [str(n).strip() for n in todo["Save as"]]
            for (_, r), name in zip(todo.iterrows(), names):
                where = f"{r['File']} / {r['Sheet']}"
                if not name:
                    problems.append(f"{where}: give it a name.")
                    continue
                if sum(n.casefold() == name.casefold() for n in names) > 1:
                    problems.append(f"'{name}' is used for more than one sheet — names must be different.")
                    continue
                try:
                    df = read_sheet(files[r["File"]], r["File"], r["Sheet"], int(r["Headings in row"] or 1))
                except Exception as e:  # noqa: BLE001
                    problems.append(f"{where}: could not be read ({e}).")
                    continue
                ready.append((name, r["File"], r["Sheet"], df))
            for p in sorted(set(problems)):
                st.error(p)

            if ready:
                summary = pd.DataFrame([
                    {"Save as": n, "From": f"{f} / {s}", "Rows": len(d), "Columns": len(d.columns),
                     "Note": "empty — will be skipped" if d.empty
                     else ("replaces the saved copy" if n.casefold() in taken else "")}
                    for n, f, s, d in ready
                ])
                st.dataframe(summary, hide_index=True, width="stretch")
                peek = st.selectbox("Preview a sheet", [n for n, *_ in ready], key="up_peek")
                st.dataframe(next(d for n, _, _, d in ready if n == peek).head(100),
                             hide_index=True, width="stretch")
                to_save = [x for x in ready if not x[3].empty]
                if st.button(f"Save {len(to_save)} sheet(s)", type="primary",
                             disabled=bool(problems) or not to_save):
                    for name, fname, sheet, df in to_save:
                        db.save_source(name, fname, sheet, df)
                    st.session_state.up_gen += 1  # empties the uploader
                    st.session_state.up_saved = len(to_save)
                    st.rerun()
    if st.session_state.pop("up_saved", None):
        st.success("Saved. Go to **Filter & compare** to work with them.")

    srcs = db.list_sources()
    for kind, heading, empty_note in [
        ("built", "Sheets you made", None),
        ("upload", "Uploaded sheets", "No sheets uploaded yet."),
    ]:
        group = [s for s in srcs if s.get("kind", "upload") == kind]
        if not group and not empty_note:
            continue
        st.subheader(heading)
        if not group:
            st.info(empty_note)
        for s in group:
            origin = s["sheet_name"] if kind == "built" else f"from {s['file_name']} / {s['sheet_name']}"
            with st.expander(f"**{s['name']}** — {s['n_rows']} rows, {s['n_cols']} columns · "
                             f"{origin} · saved {s['uploaded_at']}"):
                sdf = db.get_source_df(s["name"])
                if kind == "built":
                    st.caption("Change cells by typing in them. **Delete rows:** tick the boxes at the "
                               "left of the rows, then press 🗑 at the top right of the table. "
                               "**Delete columns:** choose them below. Then press **Save changes**.")
                    edited = st.data_editor(sdf, num_rows="dynamic", hide_index=True, width="stretch",
                                            key=f"edit_src_{s['id']}_{s['uploaded_at']}")
                    e1, e2 = st.columns([3, 1], vertical_alignment="bottom")
                    drop_cols = e1.multiselect("Delete columns", list(sdf.columns),
                                               key=f"drop_cols_{s['id']}_{s['uploaded_at']}")
                    new_df = edited.drop(columns=drop_cols).fillna("").astype(str)
                    changed = not new_df.reset_index(drop=True).equals(sdf.reset_index(drop=True))
                    if e2.button("Save changes", type="primary", key=f"save_src_{s['id']}",
                                 disabled=not changed or new_df.columns.empty, width="stretch"):
                        db.replace_source_df(s["name"], new_df)
                        st.rerun()
                    if changed:
                        st.caption(f"Not saved yet: {len(new_df)} row(s), {len(new_df.columns)} column(s) "
                                   f"(was {len(sdf)} and {len(sdf.columns)}).")
                else:
                    st.dataframe(sdf, width="stretch", hide_index=True)
                c1, c2, c3 = st.columns([2, 2, 1])
                c1.download_button("⬇ Download Excel", U.to_excel_bytes(sdf, {}, s["name"]),
                                   file_name=f"{s['name']}.xlsx", key=f"dl_src_{s['id']}")
                sure = c2.checkbox("Yes, delete this sheet", key=f"del_src_ok_{s['id']}")
                if c3.button("Delete", key=f"del_src_{s['id']}", disabled=not sure):
                    db.delete_source(s["name"])
                    st.rerun()


# =====================================================================
# Filter & compare
# =====================================================================
elif page == "Filter & compare":
    st.header("Filter & compare")
    st.write("**1** Pick a sheet and filter it with the **▼** on the column headings, like Excel. "
             "**2** Check a column against another sheet so only the rows found there stay. "
             "**3** Pick the columns you want, then **4** save them as a new sheet — or add them "
             "to a sheet you made before.")

    srcs = {s["name"]: s for s in db.list_sources()}
    acc = db.get_accounts_df()
    names = list(srcs) + ([MAIN] if not acc.empty else [])
    if not names:
        st.info("Upload your Excel files on the **Excel files** page first.")
        st.stop()

    def get_df(name):
        if name == MAIN:
            return acc.drop(columns="ID")
        s = srcs[name]
        return load_sheet(name, s["uploaded_at"], s["n_rows"])

    def label(name):
        if name == MAIN:
            return f"{MAIN} ({len(acc)} rows)"
        s = srcs[name]
        return f"{'⭐ ' if s['kind'] == 'built' else ''}{name} ({s['n_rows']} rows)"

    sheet = st.selectbox("Sheet to work on", names, format_func=label, key="fc_sheet",
                         help="⭐ = a sheet you made on this page.")
    df = get_df(sheet)

    # ---- 1. filter in the column headings
    st.subheader("1 · Filter")
    st.caption("Click **▼** on a column heading to sort it, tick the values to keep (with a search box), "
               "or set a condition such as *date is in year · 2026* or *greater than · 50000*. "
               "Filters on several columns all apply together.")
    filtered = filter_grid(df, key=f"fc_filters::{sheet}")
    grid_state = st.session_state[f"fc_filters::{sheet}"]

    # ---- 2. check against other sheets
    st.subheader("2 · Check against another sheet")
    st.caption("Choose a column of your filtered rows, then the sheet and column to look for it in. "
               "Every row stays visible until you choose the column to check against. "
               "Matching ignores capitals and extra spaces.")
    checks = st.session_state.setdefault("fc_checks", {}).setdefault(sheet, [1])
    others = [n for n in names if n != sheet]
    result, notes, lookups = filtered, [], []
    for uid in list(checks):
        k = f"fc_chk::{sheet}::{uid}"
        box = st.container(border=True)
        top = box.columns([4, 5, 4], vertical_alignment="bottom")
        bottom = box.columns([5, 4, 3], vertical_alignment="bottom")
        c = [top[0], top[1], top[2], bottom[0], bottom[1], bottom[2]]
        col = c[0].selectbox("Column in this sheet", list(df.columns), key=f"{k}_col")
        osheet = c[1].selectbox("Look for it in sheet", others, index=None, format_func=label,
                                placeholder="choose a sheet…", key=f"{k}_sheet")
        ocol = None
        if osheet:
            odf = get_df(osheet)
            ocol = c[2].selectbox("…in its column", list(odf.columns), index=None,
                                  placeholder="choose a column…", key=f"{k}_ocol_{osheet}")
        else:
            c[2].selectbox("…in its column", [], placeholder="choose a sheet first", disabled=True,
                           key=f"{k}_ocol_none")
        how = c[3].selectbox(
            "Match when", U.MATCH_HOWS, key=f"{k}_how",
            help="Each value of *Column in this sheet* is looked for in **every row of the one column "
                 "you chose in the other sheet** — no other columns are searched.  \n"
                 "**the values are the same** – the whole cell must be equal, e.g. codes.  \n"
                 "**the other cell contains this value** – 'ABC Pharma' is found in "
                 "'ABC Pharma GmbH, Munich'.  \n"
                 "**either one contains the other** – as above, and also 'ABC Pharma GmbH' here is "
                 "found when the other cell is just 'ABC Pharma'.  \n"
                 "Capitals, accents, punctuation and line breaks are ignored; only whole words count.")
        keep = c[4].selectbox("Keep the rows", ["found there", "NOT found there"], key=f"{k}_keep")
        if c[5].button("✕ Remove this check", key=f"{k}_rm", width="stretch"):
            checks.remove(uid)
            st.rerun()
        if ocol:
            values = result[col]
            matched = U.match_rows(values, odf[ocol], how)
            found = matched.notna()
            n_found, before = int(found.sum()), len(result)
            result = result[found] if keep == "found there" else result[~found]
            if keep == "found there":
                taken = {lk["label"] for lk in lookups} | {sheet}
                lk_label = osheet if osheet not in taken else f"{osheet} (check {uid})"
                lookups.append(dict(label=lk_label, df=odf, match=matched[found]))
            box.caption(f"✔ {n_found} of {before} row(s) have their **{col}** in **{osheet} › {ocol}** "
                       f"({how}) — keeping the {len(result)} row(s) {keep}.")
            if n_found < before:
                for h in U.MATCH_HOWS[U.MATCH_HOWS.index(how) + 1:]:
                    n = int(U.found_in(values, odf[ocol], h).sum())
                    if n > n_found:
                        box.info(f"💡 {n} of {before} row(s) are found if you set **Match when** to "
                                f"*{h}*. Useful when a cell holds more than the name, e.g. an address.")
                        break
            notes.append(f"{col} {'found' if keep == 'found there' else 'not found'} in {osheet} › {ocol}"
                         + ("" if how == U.MATCH_HOWS[0] else f" ({how})"))
    if st.button("+ Check against another sheet", disabled=not others):
        checks.append(max(checks, default=0) + 1)
        st.rerun()

    # ---- 3. result and columns
    st.subheader("3 · Result — choose the columns to keep")
    removed = st.session_state.setdefault("fc_removed", {}).setdefault(sheet, set())
    after_checks = len(result)
    result = result[~result.index.isin(list(removed))]
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Rows in the sheet", len(df))
    m2.metric("After filters", len(filtered))
    m3.metric("After checks", after_checks)
    m4.metric("Rows you removed", after_checks - len(result))
    st.caption("Choose a sheet, then the columns you want from it. Besides this sheet you can take "
               "columns from every sheet you checked against: each row gets the values of the row "
               "it matched there (the first one, if several match).")

    # what is picked per sheet; kept outside the widgets so it survives switching sheets
    picks = st.session_state.setdefault("fc_picks", {}).setdefault(sheet, {})
    sources = {sheet: result}
    for lk in lookups:
        sources[lk["label"]] = lk["df"]
    c1, c2, c3, c4 = st.columns([3, 5, 1, 1], vertical_alignment="bottom")
    src = c1.selectbox("Take columns from", list(sources),
                       format_func=lambda n: f"{n} (this sheet)" if n == sheet else f"{n} (checked against)",
                       key=f"fc_src::{sheet}")
    src_cols = list(sources[src].columns)
    wkey = f"fc_cols_w::{sheet}::{src}"
    st.session_state[wkey] = [c for c in picks.get(src, []) if c in src_cols]

    def remember(src=src, wkey=wkey):
        picks[src] = list(st.session_state[wkey])

    c2.multiselect(f"Columns from {src}", src_cols, key=wkey, on_change=remember,
                   placeholder="click here to choose columns",
                   help="They appear in your new sheet in the order you pick them. Remove one with its ✕.")
    c3.button("All", key=f"{wkey}_all", width="stretch", help="Take every column of this sheet",
              on_click=lambda: picks.__setitem__(src, src_cols))
    c4.button("None", key=f"{wkey}_none", width="stretch", help="Take no columns from this sheet",
              on_click=lambda: picks.__setitem__(src, []))

    # build the new sheet: this sheet's columns first, then each checked sheet's
    parts, origin = {}, {}
    for name, sdf in sources.items():
        chosen = [c for c in picks.get(name, []) if c in sdf.columns]
        for c in chosen:
            if name == sheet:
                values = result[c]
            else:
                lk = next(x for x in lookups if x["label"] == name)
                idx = lk["match"].reindex(result.index)
                values = pd.Series([sdf.at[i, c] if pd.notna(i) else "" for i in idx], index=result.index)
            heading = c if c not in parts else f"{c} [{name}]"
            parts[heading] = values
            origin[heading] = (name, c)
    if not parts:  # nothing chosen yet: all columns of this sheet
        parts = {c: result[c] for c in result.columns}
        origin = {c: (sheet, c) for c in result.columns}
        st.info("No columns chosen yet, so all columns of this sheet are shown.")
    out = pd.DataFrame(parts, index=result.index)
    row_labels = list(out.index)
    out = out.reset_index(drop=True)

    st.caption(f"Your new sheet: **{len(out.columns)}** column(s), **{len(out)}** row(s) — "
               "this is exactly what will be saved. **To delete rows**, tick the boxes at the left of "
               "the rows; **to delete columns**, click their headings (hold Ctrl for several). "
               "Then press the 🗑 button below the table.")
    gen_key = f"fc_preview_gen::{sheet}"
    st.session_state.setdefault(gen_key, 0)
    view = out.copy()
    view.index = range(1, len(view) + 1)
    event = st.dataframe(view, column_config=excel_config(view.columns), width="stretch",
                         height=min(420, 38 + 35 * max(len(view), 1)), on_select="rerun",
                         selection_mode=["multi-row", "multi-column"],
                         key=f"fc_preview::{sheet}::{st.session_state[gen_key]}")
    sel_rows = [p for p in event.selection.rows if p < len(row_labels)]
    sel_cols = [c for c in event.selection.columns if c in origin]
    b1, b2, b3 = st.columns(3)
    if b1.button(f"🗑 Remove {len(sel_rows)} selected row(s)", disabled=not sel_rows, width="stretch"):
        removed.update(row_labels[p] for p in sel_rows)
        st.session_state[gen_key] += 1
        st.rerun()
    if b2.button(f"🗑 Remove {len(sel_cols)} selected column(s)", disabled=not sel_cols, width="stretch"):
        if not any(picks.values()):  # was showing all columns: from now on, all but these
            picks[sheet] = [c for c in result.columns if c not in sel_cols]
        else:
            for h in sel_cols:
                name, c = origin[h]
                picks[name] = [x for x in picks.get(name, []) if x != c]
        st.session_state[gen_key] += 1
        st.rerun()
    if b3.button(f"↺ Bring back removed rows ({after_checks - len(result)})",
                 disabled=after_checks == len(result), width="stretch"):
        removed.clear()
        st.session_state[gen_key] += 1
        st.rerun()

    describe = f"Made from {sheet}"
    if grid_state["filters"]:
        describe += " · filtered on " + ", ".join(grid_state["filters"])
    if notes:
        describe += " · " + "; ".join(notes)

    # ---- 4. save
    st.subheader("4 · Save")
    if msg := st.session_state.pop("fc_msg", None):
        st.success(msg)
    built = [n for n, s in srcs.items() if s["kind"] == "built"]
    how = st.radio("What to do with these rows",
                   ["Make a new sheet", "Add them to a sheet I made before"],
                   horizontal=True, key="fc_how", label_visibility="collapsed")
    if out.empty:
        st.info("There are no rows to save. Loosen a filter or a check.")
    elif how == "Make a new sheet":
        c1, c2 = st.columns([3, 2], vertical_alignment="bottom")
        new_name = c1.text_input("Name of the new sheet", placeholder="e.g. Audits due in 2026",
                                 key="fc_new_name").strip()
        stamp = c2.checkbox("Add an 'Added on' column with today's date", key="fc_new_stamp",
                            help="Handy when you keep adding rows to this sheet later.")
        clash = new_name.casefold() in {n.casefold() for n in names}
        if clash:
            st.error(f"'{new_name}' is already used. Choose another name, or use "
                     "*Add them to a sheet I made before*.")
        if st.button(f"Save {len(out)} row(s) as a new sheet", type="primary",
                     disabled=not new_name or clash):
            to_save = out.copy()
            if stamp:
                to_save["Added on"] = date.today().strftime("%Y-%m-%d")
            db.save_source(new_name, "Filter & compare", describe, to_save, kind="built")
            st.session_state.fc_msg = (f"Saved '{new_name}' with {len(to_save)} row(s). It is listed with ⭐ "
                                       "above, so you can filter and compare it too.")
            st.rerun()
    elif not built:
        st.info("You haven't made a sheet yet. Use **Make a new sheet** first.")
    else:
        target = st.selectbox("Add to", built, format_func=label, key="fc_target")
        tdf = get_df(target)
        st.caption(f"'{target}' has these columns: " + ", ".join(tdf.columns))
        st.markdown("**Which column of the sheet does each of your columns go into?**")
        NEW_COL, SKIP_COL = "+ Add as a new column", "— Leave out —"
        mapping = {}
        for i, c in enumerate(out.columns):
            r = st.columns([2, 3])
            r[0].markdown(f"{c}")
            match = U.auto_match(c, tdf.columns)
            opts = list(tdf.columns) + [NEW_COL, SKIP_COL]
            choice = r[1].selectbox("Goes into", opts, index=opts.index(match) if match else len(tdf.columns),
                                    key=f"fc_map::{target}::{c}", label_visibility="collapsed")
            if choice != SKIP_COL:
                mapping[c] = c if choice == NEW_COL else choice
        dup_targets = {t for t in mapping.values() if list(mapping.values()).count(t) > 1}
        if dup_targets:
            st.error("Two of your columns go into the same column: " + ", ".join(dup_targets) + ". Pick one.")
        new_rows = out[list(mapping)].rename(columns=mapping)
        stamp = st.checkbox("Fill 'Added on' with today's date", value="Added on" in tdf.columns,
                            key="fc_add_stamp")
        if stamp:
            new_rows["Added on"] = date.today().strftime("%Y-%m-%d")

        skipped = 0
        key_opts = [c for c in new_rows.columns if c in tdf.columns and c != "Added on"]
        if key_opts:
            c1, c2 = st.columns([2, 3], vertical_alignment="bottom")
            no_dups = c1.checkbox("Don't add rows that are already in the sheet", value=True,
                                  key="fc_nodup")
            default = next((i for i, c in enumerate(key_opts) if "code" in c.lower() or "id" in c.lower().split()), 0)
            key_col = c2.selectbox("Same row = same value in", key_opts, index=default,
                                   key=f"fc_dupkey::{target}", disabled=not no_dups)
            if no_dups:
                keys = new_rows[key_col].map(U.norm_key)
                already = keys.isin(set(tdf[key_col].map(U.norm_key)) - {""})
                repeat = keys.duplicated() & (keys != "")
                skipped = int((already | repeat).sum())
                new_rows = new_rows[~(already | repeat)]
        st.write(f"**{len(new_rows)}** row(s) will be added"
                 + (f" · {skipped} skipped because they are already there" if skipped else "") + ".")
        with st.expander("Preview the rows to add"):
            st.dataframe(new_rows, hide_index=True, width="stretch")
        if st.button(f"Add {len(new_rows)} row(s) to '{target}'", type="primary",
                     disabled=new_rows.empty or bool(dup_targets) or not mapping):
            db.append_source_rows(target, new_rows)
            st.session_state.fc_msg = f"Added {len(new_rows)} row(s) to '{target}'."
            st.rerun()

    if not out.empty:
        st.download_button("⬇ Download these rows (Excel)", U.to_excel_bytes(out, {}, sheet),
                           file_name=f"{sheet} - filtered.xlsx")


# =====================================================================
# Parameters
# =====================================================================
elif page == "Parameters":
    st.header("Parameters")
    st.write("Parameters are the columns of your Main accounts sheet. Add a new one when a year "
             "needs it; on the Year-end sheet you choose which ones to show, so old ones can stay.")

    params = db.get_parameters()
    counts = db.get_accounts_df()
    table = pd.DataFrame(
        [{"Order": i + 1, "Name": p["name"], "Type": p["dtype"],
          "Filled in": int((counts[p["name"]] != "").sum()) if p["name"] in counts else 0}
         for i, p in enumerate(params)]
    )
    edited = st.data_editor(
        table,
        hide_index=True,
        width="stretch",
        disabled=["Filled in"],
        column_config={
            "Order": st.column_config.NumberColumn(min_value=1, step=1, help="Change the numbers to reorder"),
            "Name": st.column_config.TextColumn(required=True),
            "Type": st.column_config.SelectboxColumn(options=db.PARAM_TYPES, required=True),
            "Filled in": st.column_config.NumberColumn(help="How many accounts have a value"),
        },
        key="param_editor",
    )
    st.caption("Types: **Money** = exact amounts with 2 decimals and totals · **Number** · "
               "**Date** (enter as DD-MM-YYYY) · **Text**.")
    if st.button("Save changes", type="primary"):
        try:
            for p, (_, row) in zip(params, edited.iterrows()):
                if row["Name"] != p["name"] or row["Type"] != p["dtype"]:
                    db.update_parameter(p["id"], str(row["Name"]), row["Type"])
            order = sorted(zip(edited["Order"].fillna(999).tolist(), range(len(params))))
            db.set_parameter_order([params[i]["id"] for _, i in order])
            st.success("Saved.")
            st.rerun()
        except ValueError as e:
            st.error(str(e))

    st.subheader("Add a parameter")
    with st.form("add_param", clear_on_submit=True):
        c1, c2 = st.columns([3, 1])
        new_name = c1.text_input("Name", placeholder="e.g. GST Number, Extension Date, TDS Amount")
        new_type = c2.selectbox("Type", db.PARAM_TYPES)
        if st.form_submit_button("Add parameter"):
            try:
                db.add_parameter(new_name, new_type)
                st.success(f"Added '{new_name.strip()}'.")
                st.rerun()
            except ValueError as e:
                st.error(str(e))

    st.subheader("Delete parameters")
    to_del = st.multiselect("Parameters to delete", [p["name"] for p in params])
    if to_del:
        st.warning("Deleting a parameter also deletes its values from every account. This cannot be undone.")
        sure = st.checkbox("Yes, delete them and their values")
        if st.button("Delete", disabled=not sure):
            for p in params:
                if p["name"] in to_del:
                    db.delete_parameter(p["id"])
            st.rerun()


# =====================================================================
# Main accounts
# =====================================================================
elif page == "Main accounts":
    st.header("Main accounts")
    params = db.get_parameters()
    types = {p["name"]: p["dtype"] for p in params}
    t_view, t_add, t_import = st.tabs(["View & edit", "Add one account", "Import from Excel"])

    with t_view:
        acc = db.get_accounts_df()
        q = st.text_input("Search", placeholder="Type any part of a name, code, date…")
        view = acc
        if q:
            mask = acc.drop(columns="ID").apply(lambda s: s.str.contains(q, case=False, regex=False)).any(axis=1)
            view = acc[mask]
        st.caption(f"Showing {len(view)} of {len(acc)} account(s). Edit cells directly, then press "
                   "**Save changes**. Dates are stored as YYYY-MM-DD; you may type DD-MM-YYYY.")
        edited = st.data_editor(
            view, hide_index=True, disabled=["ID"], width="stretch", key="acc_editor",
        )
        money_totals_caption(view, types)
        if st.button("Save changes", type="primary", key="acc_save"):
            pid = {p["name"]: p["id"] for p in params}
            items, problems = [], []
            for (_, old), (_, new) in zip(view.iterrows(), edited.iterrows()):
                changed = {}
                for p in params:
                    n = p["name"]
                    if str(new[n] or "") != str(old[n] or ""):
                        v, prob = U.clean_for_type(new[n], p["dtype"])
                        if prob:
                            problems.append(f"ID {old['ID']}, {n}: {prob}")
                        changed[pid[n]] = v
                if changed:
                    items.append((int(old["ID"]), changed))
            if problems:
                st.error("Nothing was saved. Fix these first:\n\n- " + "\n- ".join(problems))
            elif items:
                db.save_accounts_bulk(items)
                db.log("Accounts edited", f"{len(items)} account(s) changed")
                st.success(f"Saved changes to {len(items)} account(s).")
                st.rerun()
            else:
                st.info("No changes to save.")

        with st.expander("Delete accounts"):
            ids = st.multiselect("Account IDs to delete", view["ID"].tolist())
            sure = st.checkbox("Yes, permanently delete these accounts", key="acc_del_ok")
            if st.button("Delete accounts", disabled=not (ids and sure)):
                db.delete_accounts(ids)
                st.rerun()

        if not acc.empty:
            st.download_button(
                "Download Main accounts (Excel)",
                U.to_excel_bytes(acc.drop(columns="ID"), types, "Main accounts"),
                file_name=f"main_accounts_{date.today():%Y%m%d}.xlsx",
            )

    with t_add:
        with st.form("add_account", clear_on_submit=True):
            vals = {}
            cols = st.columns(2)
            hints = {"Date": "DD-MM-YYYY", "Money": "e.g. 125000.50", "Number": "e.g. 12", "Text": ""}
            for i, p in enumerate(params):
                vals[p["id"]] = cols[i % 2].text_input(
                    f"{p['name']} ({p['dtype']})", placeholder=hints[p["dtype"]], key=f"new_{p['id']}"
                )
            if st.form_submit_button("Add account", type="primary"):
                clean, problems = {}, []
                for p in params:
                    v, prob = U.clean_for_type(vals[p["id"]], p["dtype"])
                    if prob:
                        problems.append(f"{p['name']}: {prob}")
                    clean[p["id"]] = v
                if problems:
                    st.error("Not added. " + "; ".join(problems))
                elif not any(clean.values()):
                    st.error("Fill in at least one field.")
                else:
                    db.save_accounts_bulk([(None, clean)])
                    db.log("Account added", "entered by hand")
                    st.success("Account added.")

    with t_import:
        st.write("Bring rows from an Excel sheet into Main accounts. Each column is matched to a parameter.")
        how = st.radio("Take the rows from", ["A new Excel file", "An Excel file saved on the Excel files page"],
                       horizontal=True, key="imp_how")
        df, label = None, ""
        if how == "A new Excel file":
            res = pick_table("imp")
            if res:
                df, fname, sheet = res
                label = f"{fname} / {sheet}"
        else:
            names = [s["name"] for s in db.list_sources()]
            if not names:
                st.info("No saved files yet.")
            else:
                label = st.selectbox("Saved file", names, key="imp_src")
                df = db.get_source_df(label)
        if df is not None:
            with st.expander(f"Preview ({len(df)} rows)"):
                st.dataframe(df.head(50), hide_index=True, width="stretch")
            import_panel(df, "imp", label)


# =====================================================================
# Year-end sheet
# =====================================================================
elif page == "Year-end sheet":
    st.header("Year-end sheet")
    st.write("Compare your files, keep only the rows and columns you need, check the totals, "
             "then download the PDF or save the rows into Main accounts.")

    params = db.get_parameters()
    param_types = {p["name"]: p["dtype"] for p in params}
    datasets = {}
    acc = db.get_accounts_df()
    if not acc.empty:
        datasets[MAIN] = acc.drop(columns="ID")
    for s in db.list_sources():
        datasets[s["name"]] = db.get_source_df(s["name"])
    if not datasets:
        st.info("Upload your Excel files on the **Excel files** page first.")
        st.stop()

    c1, c2 = st.columns([3, 1])
    year = c2.number_input("Year", min_value=2000, max_value=2100, value=date.today().year, step=1)
    title = c1.text_input("Title of the sheet", value=f"Supplier Accounts — {year}")

    # ---- 1 & 2. sheets, shown like Excel
    st.subheader("1 · Pick your sheets and columns")
    c1, c2 = st.columns(2)
    base_name = c1.selectbox("Main list of suppliers", list(datasets), key="ye_base",
                             help="Every row of the final sheet comes from this sheet.")
    base_df = datasets[base_name]
    key_default = next((i for i, c in enumerate(base_df.columns) if "code" in c.lower()), 0)
    base_key = c2.selectbox("Match rows using column", list(base_df.columns), index=key_default,
                            key=f"ye_key_{base_name}", help="Usually the supplier code. Marked 🔑 in the sheets.")
    st.caption("Each tab is one sheet. **Click a column heading to select it** (click again to unselect; "
               "hold Ctrl/Shift for several). Selected columns go into the year-end sheet.")

    others = [n for n in datasets if n != base_name]
    tabs = st.tabs([f"📗 {base_name} (main list)"] + [f"📄 {n}" for n in others])
    base_keys = set(base_df[base_key].map(U.norm_key)) - {""}

    with tabs[0]:
        st.caption(f"{len(base_df)} rows · {len(base_df.columns)} columns")
        base_sel = excel_grid(base_df, key=f"grid_base_{base_name}", key_col=base_key)
        st.markdown("**Showing:** " + (", ".join(base_sel) if base_sel else "all columns (none selected)"))

    joins = []
    for tab, n in zip(tabs[1:], others):
        odf = datasets[n]
        with tab:
            c1, c2 = st.columns(2)
            match = U.auto_match(base_key, odf.columns)
            okey = c1.selectbox(f"Matching column in {n}", list(odf.columns),
                                index=list(odf.columns).index(match) if match else 0, key=f"ye_okey_{n}")
            mode = c2.selectbox("Which rows to keep", U.JOIN_MODES, key=f"ye_mode_{n}")
            okeys = set(odf[okey].map(U.norm_key)) - {""}
            found = len(base_keys & okeys)
            st.caption(f"{len(odf)} rows · {found} of {len(base_keys)} suppliers from the main list are in this sheet"
                       + (f" · {len(okeys - base_keys)} here are NOT in the main list" if okeys - base_keys else ""))
            cols = [c for c in excel_grid(odf, key=f"grid_{n}", key_col=okey) if c != okey]
            st.markdown("**Bringing in:** " + (", ".join(cols) if cols else "nothing yet — click column headings"))
            if cols or mode != U.JOIN_MODES[0]:
                joins.append(dict(label=n, df=odf, key=okey, columns=cols, mode=mode))

    combined, warnings = U.join_sources(base_df, base_key, joins)
    shown_by_default = set(base_sel or base_df.columns) | (set(combined.columns) - set(base_df.columns))

    # ---- 3. filters
    st.subheader("2 · Filter rows")
    if "filters" not in st.session_state:
        st.session_state.filters = []
        st.session_state.filter_uid = 0
    all_cols = list(combined.columns)
    match_all = st.radio("Keep rows that match", ["ALL filters", "ANY filter"], horizontal=True,
                         key="ye_matchall") == "ALL filters"
    filters = []
    for uid in list(st.session_state.filters):
        c = st.columns([3, 3, 3, 2, 1])
        col = c[0].selectbox("Column", all_cols, key=f"f_{uid}_col", label_visibility="collapsed")
        op = c[1].selectbox("Condition", U.OPERATORS, key=f"f_{uid}_op", label_visibility="collapsed")
        f = dict(col=col, op=op, value="", value2="", other_col=None)
        if op in U.COLUMN_OPS:
            f["other_col"] = c[2].selectbox("Other column", all_cols, key=f"f_{uid}_other",
                                            label_visibility="collapsed")
        elif op not in U.NO_VALUE_OPS:
            ph = "year e.g. 2026" if op == "date is in year" else "from" if op == "between" else "value"
            f["value"] = c[2].text_input("Value", key=f"f_{uid}_v", placeholder=ph, label_visibility="collapsed")
            if op == "between":
                f["value2"] = c[3].text_input("To", key=f"f_{uid}_v2", placeholder="to",
                                              label_visibility="collapsed")
        if c[4].button("✕", key=f"f_{uid}_rm", help="Remove this filter"):
            st.session_state.filters.remove(uid)
            st.rerun()
        filters.append(f)
    if st.button("+ Add filter"):
        st.session_state.filter_uid += 1
        st.session_state.filters.append(st.session_state.filter_uid)
        st.rerun()
    st.caption("Numbers and dates are compared as numbers and dates (type dates as DD-MM-YYYY). "
               "Text is compared ignoring capitals. Example: *Account End · date is in year · 2026*.")

    result = U.apply_filters(combined, filters, match_all)

    # ---- 4. columns & totals
    st.subheader("3 · Headings, format and totals")
    base_types = param_types if base_name == MAIN else {}
    layout = pd.DataFrame([
        {
            "Show": c in shown_by_default,
            "Column": c,
            "Show as": c,
            "Format": base_types.get(c) or U.guess_type(c, combined[c]),
            "Total": (base_types.get(c) or U.guess_type(c, combined[c])) == "Money",
        }
        for c in all_cols
    ])
    layout = st.data_editor(
        layout,
        hide_index=True,
        width="stretch",
        disabled=["Column"],
        column_config={
            "Show": st.column_config.CheckboxColumn(width="small"),
            "Show as": st.column_config.TextColumn(help="Heading printed on the sheet"),
            "Format": st.column_config.SelectboxColumn(options=db.PARAM_TYPES, required=True),
            "Total": st.column_config.CheckboxColumn(help="Add a total at the bottom", width="small"),
        },
        key=f"ye_layout_{abs(hash((tuple(all_cols), tuple(base_sel)))) % 10**8}",
    )
    st.caption("Change **Show as** to rename a heading, tick **Total** to add up a column.")
    shown = layout[layout["Show"]]
    c1, c2 = st.columns([3, 1])
    sort_col = c1.selectbox("Sort rows by", ["(file order)"] + shown["Column"].tolist(), key="ye_sort")
    desc = c2.checkbox("Largest / latest first", key="ye_desc")
    if sort_col != "(file order)":
        result = U.sort_df(result, sort_col, desc)

    # final raw table with display names
    rename, seen = {}, set()
    for _, r in shown.iterrows():
        name = str(r["Show as"] or "").strip() or r["Column"]
        while name in seen:
            name += " "
        seen.add(name)
        rename[r["Column"]] = name
    final_raw = result[list(rename)].rename(columns=rename)
    final_types = {rename[r["Column"]]: r["Format"] for _, r in shown.iterrows()}
    totals, bad_totals = {}, []
    for _, r in shown.iterrows():
        if r["Total"]:
            name = rename[r["Column"]]
            t, bad = U.column_total(final_raw[name])
            totals[name] = U.format_money(t) if r["Format"] == "Money" else format(t, "f")
            if bad:
                bad_totals.append(f"'{name}' has {bad} cell(s) that are not numbers — they are NOT in the total.")
    final_view = final_raw.copy()
    for c, t in final_types.items():
        final_view[c] = final_view[c].map(lambda v, t=t: U.format_value(v, t))

    # ---- 5. result
    st.subheader("4 · Year-end sheet")
    for w in warnings + bad_totals:
        st.warning(w)
    m1, m2, m3 = st.columns(3)
    m1.metric("Rows in main list", len(base_df))
    m2.metric("Rows in year-end sheet", len(final_view))
    m3.metric("Columns", len(final_view.columns))

    grid = final_view.copy()
    grid.index = [str(i) for i in range(1, len(grid) + 1)]
    if totals and len(grid.columns):
        trow = {c: totals.get(c, "") for c in grid.columns}
        if grid.columns[0] not in totals:
            trow[grid.columns[0]] = "TOTAL"
        grid.loc["Σ"] = trow
    grid_cfg = excel_config(grid.columns)
    if totals and grid.size <= 200_000:
        grid = grid.style.apply(
            lambda r: ["font-weight: bold; background-color: #D5E3E8; color: #0B2530"] * len(r)
            if r.name == "Σ" else [""] * len(r),
            axis=1,
        )
    st.dataframe(grid, column_config=grid_cfg, width="stretch", height=min(520, 38 + 35 * (len(final_view) + 1)))

    notes = [f"Based on: {base_name} (matched on '{base_key}')"]
    for j in joins:
        notes.append(f"Compared with: {j['label']} on '{j['key']}' — {j['mode'].lower()}")
    for f in filters:
        if f["op"] in U.COLUMN_OPS:
            notes.append(f"Filter: {f['col']} {f['op']} '{f['other_col']}'")
        elif f["op"] in U.NO_VALUE_OPS:
            notes.append(f"Filter: {f['col']} {f['op']}")
        else:
            val = f"{f['value']} and {f['value2']}" if f["op"] == "between" else f["value"]
            notes.append(f"Filter: {f['col']} {f['op']} {val}")
    if len(filters) > 1:
        notes.append("Rows match " + ("all" if match_all else "any") + " of the filters.")

    c1, c2 = st.columns(2)
    if not final_view.empty and len(final_view.columns):
        c1.download_button(
            "⬇ Download PDF",
            build_pdf(final_view, final_types, title, db.get_setting("company"), notes, totals),
            file_name=f"{title}.pdf".replace("/", "-"),
            mime="application/pdf",
            type="primary",
            width="stretch",
        )
        c2.download_button(
            "⬇ Download Excel",
            U.to_excel_bytes(final_raw, final_types, str(year)),
            file_name=f"{title}.xlsx".replace("/", "-"),
            width="stretch",
        )

        with st.expander("Save these rows into Main accounts"):
            if base_name == MAIN and not joins:
                st.caption("These rows already come from Main accounts. Use this when you brought in "
                           "columns from other files and want to store them.")
            import_panel(final_raw, "ye_imp", f"Year-end sheet '{title}'")
    else:
        st.info("No rows match. Loosen a filter or change the comparison settings.")


# =====================================================================
# History & backup
# =====================================================================
elif page == "History & backup":
    st.header("History & backup")
    st.write("All data is stored in one file: `accounts.db` in the app folder. Download a copy "
             "regularly and keep it safe — copying that file back restores everything.")
    if db.DB_PATH.exists():
        st.download_button("⬇ Download backup (accounts.db)", db.DB_PATH.read_bytes(),
                           file_name=f"accounts_backup_{date.today():%Y%m%d}.db", type="primary")
    st.subheader("Change history")
    st.dataframe(db.get_log(), hide_index=True, width="stretch")
