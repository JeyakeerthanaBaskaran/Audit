"""Supplier Accounts - year-end account builder.

Run with:  streamlit run app.py
"""
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

import db
import utils as U
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
        ["Excel files", "Parameters", "Main accounts", "Year-end sheet", "History & backup"],
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
    st.write("Upload the Excel files you compare every year. They are saved as-is, "
             "so you can use them later on the **Year-end sheet** page. Uploading a file with "
             "the same name replaces the old copy.")

    res = pick_table("up")
    if res:
        df, fname, sheet = res
        st.write(f"**{len(df)}** rows · **{len(df.columns)}** columns")
        st.dataframe(df.head(100), width="stretch", hide_index=True)
        name = st.text_input("Save as", value=Path(fname).stem, key="up_name",
                             help="A short name, e.g. 'Contracts' or 'Payments 2026'.")
        taken = {s["name"].casefold() for s in srcs}
        if name.strip().casefold() in taken:
            st.warning(f"A file called '{name}' is already saved. Saving will replace it.")
        if st.button("Save file", type="primary", disabled=not name.strip() or df.empty):
            db.save_source(name.strip(), fname, sheet, df)
            st.success(f"Saved '{name.strip()}' ({len(df)} rows).")
            st.rerun()

    st.subheader("Saved files")
    srcs = db.list_sources()
    if not srcs:
        st.info("No files saved yet.")
    for s in srcs:
        with st.expander(f"**{s['name']}** — {s['n_rows']} rows, {s['n_cols']} columns · "
                         f"from {s['file_name']} / {s['sheet_name']} · saved {s['uploaded_at']}"):
            st.dataframe(db.get_source_df(s["name"]), width="stretch", hide_index=True)
            c1, c2 = st.columns([3, 1])
            sure = c1.checkbox("Yes, delete this saved file", key=f"del_src_ok_{s['id']}")
            if c2.button("Delete", key=f"del_src_{s['id']}", disabled=not sure):
                db.delete_source(s["name"])
                st.rerun()


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
