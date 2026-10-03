"""Excel-style table: every column heading has a ▼ that opens a filter menu
(sort, a condition, and a tick list of the values), like Excel's AutoFilter.

The menu only *describes* the filter; the filtering itself is done in Python
(utils.apply_column_filters) so the result is exactly what gets saved.
"""
import hashlib

import streamlit as st

import utils as U

MAX_ROWS_SHOWN = 1000

_CSS = """
:host, .fg-root { font-family: var(--st-font, "Source Sans Pro", sans-serif); font-size: 14px;
  color: var(--st-text-color, #262730); }
.fg-bar { display: flex; align-items: center; gap: 12px; margin-bottom: 6px; flex-wrap: wrap; }
.fg-count { opacity: .8; }
.fg-btn { font: inherit; font-size: 13px; padding: 3px 10px; border-radius: 6px; cursor: pointer;
  border: 1px solid var(--st-border-color, rgba(128,128,128,.4));
  background: var(--st-background-color, #fff); color: inherit; }
.fg-btn:hover { border-color: var(--st-primary-color, #ff4b4b); color: var(--st-primary-color, #ff4b4b); }
.fg-btn.primary { background: var(--st-primary-color, #ff4b4b); border-color: var(--st-primary-color, #ff4b4b);
  color: #fff; }
.fg-btn.primary:hover { color: #fff; opacity: .9; }
.fg-btn:disabled { opacity: .45; cursor: default; }
.fg-wrap { max-height: 440px; overflow: auto; border-radius: 8px;
  border: 1px solid var(--st-border-color, rgba(128,128,128,.3)); }
table { border-collapse: separate; border-spacing: 0; width: max-content; min-width: 100%; }
th, td { padding: 4px 10px; white-space: nowrap; max-width: 320px; overflow: hidden; text-overflow: ellipsis;
  border-bottom: 1px solid var(--st-border-color, rgba(128,128,128,.2));
  border-right: 1px solid var(--st-border-color, rgba(128,128,128,.12)); text-align: left; }
th { position: sticky; top: 0; z-index: 1; font-weight: 600;
  background: var(--st-secondary-background-color, #f0f2f6); }
td.num, th.num { position: sticky; left: 0; text-align: right; opacity: .75; font-weight: 400;
  background: var(--st-secondary-background-color, #f0f2f6); }
th.num { z-index: 2; }
td { background: var(--st-background-color, #fff); }
.fg-head { display: flex; align-items: center; gap: 6px; }
.fg-letter { opacity: .55; font-weight: 400; }
.fg-drop { font: inherit; font-size: 11px; line-height: 1; padding: 3px 5px; margin-left: auto; cursor: pointer;
  border-radius: 4px; border: 1px solid var(--st-border-color, rgba(128,128,128,.45));
  background: var(--st-background-color, #fff); color: inherit; }
.fg-drop.on { background: var(--st-primary-color, #ff4b4b); border-color: var(--st-primary-color, #ff4b4b);
  color: #fff; }
.fg-empty { padding: 18px; opacity: .7; }
.fg-menu { position: fixed; z-index: 1000; width: 290px; display: flex; flex-direction: column; gap: 8px;
  padding: 10px; border-radius: 8px; box-shadow: 0 6px 24px rgba(0,0,0,.25);
  background: var(--st-background-color, #fff); color: var(--st-text-color, #262730);
  border: 1px solid var(--st-border-color, rgba(128,128,128,.35)); font-size: 13px; }
.fg-menu .row { display: flex; gap: 6px; }
.fg-menu .row > * { flex: 1; }
.fg-menu label.t { font-weight: 600; font-size: 12px; opacity: .8; }
.fg-menu select, .fg-menu input[type=text] { font: inherit; width: 100%; box-sizing: border-box; padding: 4px 6px;
  border-radius: 5px; border: 1px solid var(--st-border-color, rgba(128,128,128,.45));
  background: var(--st-secondary-background-color, #f0f2f6); color: inherit; }
.fg-list { overflow: auto; min-height: 60px; border-radius: 5px; padding: 4px 6px;
  border: 1px solid var(--st-border-color, rgba(128,128,128,.3)); }
.fg-list label { display: flex; gap: 6px; align-items: center; padding: 1px 0; cursor: pointer;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.fg-list label.all { font-weight: 600; border-bottom: 1px solid var(--st-border-color, rgba(128,128,128,.25));
  padding-bottom: 3px; margin-bottom: 3px; }
.fg-note { font-size: 12px; opacity: .7; }
.fg-sort.on { border-color: var(--st-primary-color, #ff4b4b); color: var(--st-primary-color, #ff4b4b); }
"""

_JS = r"""
const scrollMemory = new Map();

function el(tag, attrs = {}, ...kids) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === "class") e.className = v;
    else if (k === "text") e.textContent = v;
    else if (k.startsWith("on")) e.addEventListener(k.slice(2), v);
    else e.setAttribute(k, v);
  }
  for (const kid of kids) if (kid) e.appendChild(kid);
  return e;
}

function colLetter(i) {
  let s = ""; i += 1;
  while (i) { const r = (i - 1) % 26; s = String.fromCharCode(65 + r) + s; i = Math.floor((i - 1) / 26); }
  return s;
}

export default function (component) {
  const { data, key, parentElement, setTriggerValue } = component;
  const { columns, rows, total, matched, filters, sort, values, more, ops, noValueOps } = data;
  const send = (newFilters, newSort) => setTriggerValue("change", { filters: newFilters, sort: newSort });

  let root = parentElement.querySelector(".fg-root");
  if (!root) { root = el("div", { class: "fg-root" }); parentElement.appendChild(root); }
  root.replaceChildren();

  // ---- top bar
  const nActive = Object.keys(filters).length;
  const bar = el("div", { class: "fg-bar" },
    el("span", { class: "fg-count",
      text: `${matched} of ${total} rows` + (nActive ? ` · filtered on ${nActive} column(s)` : "")
        + (matched > rows.length ? ` · first ${rows.length} shown here` : "") }));
  if (nActive || sort) {
    bar.appendChild(el("button", { class: "fg-btn", text: "Clear all filters & sorting",
      onclick: () => send({}, null) }));
  }
  root.appendChild(bar);

  // ---- table
  const wrap = el("div", { class: "fg-wrap" });
  const table = el("table");
  const headRow = el("tr", {}, el("th", { class: "num", text: "#" }));
  columns.forEach((c, i) => {
    const on = !!filters[c] || (sort && sort.col === c);
    const btn = el("button", { class: "fg-drop" + (on ? " on" : ""), title: filters[c] ? "Filtered — click to change" : "Filter / sort this column",
      text: "▼" + (sort && sort.col === c ? (sort.desc ? " ↓" : " ↑") : "") });
    btn.addEventListener("click", (e) => { e.stopPropagation(); openMenu(c, btn); });
    headRow.appendChild(el("th", {}, el("div", { class: "fg-head" },
      el("span", { class: "fg-letter", text: colLetter(i) }), el("span", { text: c }), btn)));
  });
  table.appendChild(el("thead", {}, headRow));
  const body = el("tbody");
  rows.forEach((r, n) => {
    const tr = el("tr", {}, el("td", { class: "num", text: String(n + 1) }));
    r.forEach((v) => tr.appendChild(el("td", { text: v, title: v })));
    body.appendChild(tr);
  });
  table.appendChild(body);
  wrap.appendChild(table);
  if (!rows.length) wrap.appendChild(el("div", { class: "fg-empty", text: "No rows match these filters." }));
  root.appendChild(wrap);

  const mem = scrollMemory.get(key);
  if (mem) { wrap.scrollLeft = mem.left; wrap.scrollTop = mem.top; }
  wrap.addEventListener("scroll", () => scrollMemory.set(key, { left: wrap.scrollLeft, top: wrap.scrollTop }));

  // ---- filter menu
  let menu = null;
  const closeMenu = () => { if (menu) { menu.remove(); menu = null; } };
  const inMenu = (e) => menu && e.composedPath().includes(menu);
  const onDown = (e) => { if (menu && !inMenu(e)) closeMenu(); };
  const onScroll = (e) => { if (menu && !inMenu(e)) closeMenu(); };
  const onKey = (e) => { if (e.key === "Escape") closeMenu(); };
  document.addEventListener("mousedown", onDown, true);
  window.addEventListener("scroll", onScroll, true);
  document.addEventListener("keydown", onKey);

  function openMenu(col, btn) {
    closeMenu();
    const spec = filters[col] || {};
    const vals = values[col] || [];
    const listIsCut = !!more[col];
    const inSet = spec.in ? new Set(spec.in) : null;
    const outSet = new Set(spec.out || []);
    const checked = new Set(vals.filter((v) => (inSet ? inSet.has(v) : !outSet.has(v))));

    menu = el("div", { class: "fg-menu" });
    menu.addEventListener("mousedown", (e) => e.stopPropagation());

    // sort
    const isSorted = (desc) => sort && sort.col === col && !!sort.desc === desc;
    const sortRow = el("div", { class: "row" },
      el("button", { class: "fg-btn fg-sort" + (isSorted(false) ? " on" : ""), text: "Sort A → Z / 1 → 9",
        onclick: () => { closeMenu(); send(filters, { col, desc: false }); } }),
      el("button", { class: "fg-btn fg-sort" + (isSorted(true) ? " on" : ""), text: "Sort Z → A / 9 → 1",
        onclick: () => { closeMenu(); send(filters, { col, desc: true }); } }));
    menu.appendChild(sortRow);
    if (sort && sort.col === col) {
      menu.appendChild(el("button", { class: "fg-btn", text: "Remove sorting",
        onclick: () => { closeMenu(); send(filters, null); } }));
    }

    // condition
    const opSel = el("select");
    opSel.appendChild(el("option", { value: "", text: "(no condition)" }));
    ops.forEach((o) => opSel.appendChild(el("option", { value: o, text: o })));
    opSel.value = spec.op || "";
    const v1 = el("input", { type: "text", placeholder: "value" });
    const v2 = el("input", { type: "text", placeholder: "to" });
    v1.value = spec.v1 || ""; v2.value = spec.v2 || "";
    const valRow = el("div", { class: "row" }, v1, v2);
    const syncOp = () => {
      const op = opSel.value;
      valRow.style.display = !op || noValueOps.includes(op) ? "none" : "flex";
      v2.style.display = op === "between" ? "" : "none";
      v1.placeholder = op === "date is in year" ? "year, e.g. 2026" : op === "between" ? "from" : "value";
    };
    opSel.addEventListener("change", syncOp);
    syncOp();
    menu.appendChild(el("label", { class: "t", text: "Condition" }));
    menu.appendChild(opSel);
    menu.appendChild(valRow);

    // tick list
    menu.appendChild(el("label", { class: "t", text: "Values" }));
    const search = el("input", { type: "text", placeholder: "Search values…" });
    menu.appendChild(search);
    const list = el("div", { class: "fg-list" });
    const allBox = el("input", { type: "checkbox" });
    list.appendChild(el("label", { class: "all" }, allBox, el("span", { text: "(Select all)" })));
    const items = vals.map((v) => {
      const box = el("input", { type: "checkbox" });
      box.checked = checked.has(v);
      box.addEventListener("change", () => { box.checked ? checked.add(v) : checked.delete(v); refresh(); });
      const lab = el("label", { title: v === "" ? "(Blanks)" : v }, box, el("span", { text: v === "" ? "(Blanks)" : v }));
      list.appendChild(lab);
      return { v, box, lab };
    });
    menu.appendChild(list);
    if (listIsCut) {
      menu.appendChild(el("div", { class: "fg-note",
        text: `Only the first ${vals.length} values are listed. Use the search box or a condition for the rest.` }));
    }

    const visible = () => items.filter((it) => it.lab.style.display !== "none");
    const okBtn = el("button", { class: "fg-btn primary", text: "OK" });
    function refresh() {
      const vis = visible();
      allBox.checked = vis.length > 0 && vis.every((it) => checked.has(it.v));
      allBox.indeterminate = !allBox.checked && vis.some((it) => checked.has(it.v));
      const anyTicked = search.value.trim() ? vis.some((it) => checked.has(it.v)) : checked.size > 0;
      okBtn.disabled = vals.length > 0 && !anyTicked;
    }
    allBox.addEventListener("change", () => {
      visible().forEach((it) => { it.box.checked = allBox.checked; allBox.checked ? checked.add(it.v) : checked.delete(it.v); });
      refresh();
    });
    search.addEventListener("input", () => {
      const q = search.value.trim().toLowerCase();
      items.forEach((it) => { it.lab.style.display = !q || it.v.toLowerCase().includes(q) ? "" : "none"; });
      refresh();
    });
    refresh();

    okBtn.addEventListener("click", () => {
      const next = {};
      if (search.value.trim()) {
        next.in = visible().filter((it) => checked.has(it.v)).map((it) => it.v);
      } else if (checked.size < vals.length) {
        const unticked = vals.filter((v) => !checked.has(v));
        if (listIsCut || unticked.length <= checked.size) next.out = unticked;
        else next.in = vals.filter((v) => checked.has(v));
      }
      if (opSel.value) { next.op = opSel.value; next.v1 = v1.value; next.v2 = v2.value; }
      const newFilters = { ...filters };
      if (Object.keys(next).length) newFilters[col] = next; else delete newFilters[col];
      closeMenu();
      send(newFilters, sort);
    });
    const clearBtn = el("button", { class: "fg-btn", text: "Clear filter", onclick: () => {
      const newFilters = { ...filters }; delete newFilters[col]; closeMenu(); send(newFilters, sort);
    } });
    clearBtn.disabled = !filters[col];
    menu.appendChild(el("div", { class: "row" }, okBtn,
      el("button", { class: "fg-btn", text: "Cancel", onclick: closeMenu }), clearBtn));

    root.appendChild(menu);
    const r = btn.getBoundingClientRect();
    const left = Math.max(8, Math.min(r.right - 290, window.innerWidth - 300));
    const top = r.bottom + 4;
    menu.style.left = left + "px";
    menu.style.top = top + "px";
    const room = Math.max(260, window.innerHeight - top - 12);
    menu.style.maxHeight = room + "px";
    list.style.maxHeight = Math.max(80, room - 260) + "px";
    if (top + menu.offsetHeight > window.innerHeight - 8) {
      menu.style.top = Math.max(8, r.top - menu.offsetHeight - 4) + "px";
    }
  }

  return () => {
    document.removeEventListener("mousedown", onDown, true);
    window.removeEventListener("scroll", onScroll, true);
    document.removeEventListener("keydown", onKey);
    closeMenu();
  };
}
"""

_component = st.components.v2.component("excel_filter_grid", css=_CSS, js=_JS)

FILTER_OPS = [o for o in U.OPERATORS if o not in U.COLUMN_OPS]


def filter_grid(df, key):
    """Show `df` with a filter ▼ on every column heading.

    The filters and sorting are kept in st.session_state[key] and survive page
    changes. Returns the filtered (and sorted) DataFrame."""
    state = st.session_state.setdefault(key, {"filters": {}, "sort": None})
    filters = {c: s for c, s in state["filters"].items() if c in df.columns}
    sort = state["sort"] if state["sort"] and state["sort"]["col"] in df.columns else None

    result = U.apply_column_filters(df, filters)
    if sort:
        result = U.sort_df(result, sort["col"], sort["desc"], keep_index=True)
    distinct = U.distinct_values(df, filters)

    shown = result.head(MAX_ROWS_SHOWN)
    out = _component(
        key="fg_" + hashlib.md5(key.encode("utf-8")).hexdigest(),  # component keys may not contain "__"
        data={
            "columns": list(df.columns),
            "rows": shown.astype(str).values.tolist(),
            "total": len(df),
            "matched": len(result),
            "filters": filters,
            "sort": sort,
            "values": {c: v for c, (v, _) in distinct.items()},
            "more": {c: m for c, (_, m) in distinct.items()},
            "ops": FILTER_OPS,
            "noValueOps": sorted(U.NO_VALUE_OPS),
        },
        on_change_change=lambda: None,
    )
    change = out.get("change") if out else None
    if change and (change.get("filters") != filters or change.get("sort") != sort):
        state["filters"] = change.get("filters") or {}
        state["sort"] = change.get("sort")
        st.rerun()
    return result
