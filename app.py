# -*- coding: utf-8 -*-
"""
OE Coder - Streamlit app for open-end coding on any study.

Run with:   streamlit run app.py
"""
from __future__ import annotations

import json
from io import BytesIO

import pandas as pd
import streamlit as st

from oecoder import engine, excel, frames as fr

st.set_page_config(page_title="OE Coder", page_icon="🗂️", layout="wide")

PATTERN_HELP = ("One pattern per line. Each is a case-insensitive regular expression — plain "
                "words work fine (`brand name`), `\\b` marks a word boundary (`\\bpay\\b` will not "
                "match *payment*), and `(le|ility)` gives alternatives. Never separate patterns "
                "with `|`: that character is regex alternation and belongs inside a pattern.")


# ----------------------------------------------------------------- state
def init_state():
    st.session_state.setdefault("raw", {})          # q_key -> DataFrame
    st.session_state.setdefault("frames", {})       # q_key -> frame
    st.session_state.setdefault("titles", {})       # q_key -> question text
    st.session_state.setdefault("overrides", {})    # q_key -> {id: [codes]}
    st.session_state.setdefault("id_col", None)
    st.session_state.setdefault("extra_cols", [])
    st.session_state.setdefault("meta", {})


init_state()


def coded_for(q: str) -> pd.DataFrame:
    df = st.session_state.raw[q]
    return engine.code_dataframe(
        df,
        id_col=st.session_state.id_col,
        verbatim_col=st.session_state.titles[q]["col"],
        frame=st.session_state.frames.get(q, []),
        overrides=st.session_state.overrides.get(q, {}),
    )


# ----------------------------------------------------------------- sidebar
with st.sidebar:
    st.title("🗂️ OE Coder")
    st.caption("Rule-based open-end coding with an analyst review step.")

    st.subheader("Project details")
    st.session_state.meta["title"] = st.text_input(
        "Deliverable title", st.session_state.meta.get("title", "Open-End Coding Deliverable"))
    job = st.text_input("Job number", st.session_state.meta.get("job", ""))
    client = st.text_input("Client", st.session_state.meta.get("client", ""))
    st.session_state.meta["job"] = job
    st.session_state.meta["client"] = client

    st.divider()
    st.subheader("Code frame library")
    lib = st.file_uploader("Load a saved frame (.json)", type=["json"], key="lib_json")
    if lib and st.button("Load frames", use_container_width=True):
        payload = json.loads(lib.getvalue().decode("utf-8"))
        loaded, meta = (payload.get("frames", payload), payload.get("meta", {})) \
            if isinstance(payload, dict) else ({"Q1": payload}, {})
        st.session_state.frames.update({q: fr.normalise(f) for q, f in loaded.items()})
        st.success(f"Loaded {len(loaded)} frame(s): {', '.join(loaded)}")

    if st.session_state.frames:
        st.download_button(
            "💾 Save frames as .json",
            data=json.dumps({"meta": st.session_state.meta,
                             "frames": {q: fr.normalise(f) for q, f in st.session_state.frames.items()}},
                            indent=2, ensure_ascii=False),
            file_name=f"{job or 'project'}_codeframes.json",
            mime="application/json", use_container_width=True)


# ----------------------------------------------------------------- tabs
tab_data, tab_frame, tab_run, tab_review, tab_export = st.tabs(
    ["1 · Data", "2 · Code frame", "3 · Run coding", "4 · Review", "5 · Export"])

# ============================================================ 1 · DATA
with tab_data:
    st.header("Load the open-end data")
    up = st.file_uploader("Excel or CSV with the verbatims", type=["xlsx", "xlsm", "csv"])

    if up:
        if up.name.lower().endswith(".csv"):
            book = {"data": pd.read_csv(up)}
        else:
            book = pd.read_excel(up, sheet_name=None)

        st.success(f"Found {len(book)} sheet(s): {', '.join(book)}")
        sheets = st.multiselect("Sheets to code (one question per sheet)",
                                list(book), default=list(book))

        if sheets:
            first = book[sheets[0]]
            st.session_state.id_col = st.selectbox(
                "Respondent ID column", list(first.columns),
                index=list(first.columns).index(st.session_state.id_col)
                if st.session_state.id_col in first.columns else 0)
            st.session_state.extra_cols = st.multiselect(
                "Other columns to carry through (password, cell, weight…)",
                [c for c in first.columns if c != st.session_state.id_col])

            st.markdown("**Pick the verbatim column on each sheet**")
            for sh in sheets:
                df = book[sh]
                guess = [c for c in df.columns
                         if c not in (st.session_state.id_col, *st.session_state.extra_cols)]
                col = st.selectbox(f"{sh}", guess, key=f"vcol_{sh}")
                st.session_state.raw[sh] = df
                st.session_state.titles[sh] = {"col": col, "text": str(col)}

            st.divider()
            st.markdown("**Question wording shown on the deliverable**")
            for sh in sheets:
                st.session_state.titles[sh]["text"] = st.text_input(
                    f"{sh} title", st.session_state.titles[sh]["text"], key=f"title_{sh}")

            st.dataframe(book[sheets[0]].head(8), use_container_width=True)

# ============================================================ 2 · FRAME
with tab_frame:
    st.header("Build or edit the code frame")
    if not st.session_state.raw:
        st.info("Load the data first.")
    else:
        q = st.selectbox("Question", list(st.session_state.raw), key="frame_q")
        st.caption(PATTERN_HELP)

        c1, c2, c3 = st.columns(3)
        with c1:
            if st.button("Start a blank frame", use_container_width=True):
                st.session_state.frames[q] = fr.blank_frame()
        with c2:
            src = st.selectbox("Copy frame from", ["—"] + [k for k in st.session_state.frames if k != q],
                               label_visibility="collapsed")
            if src != "—" and st.button("Copy frame", use_container_width=True):
                st.session_state.frames[q] = [dict(r) for r in st.session_state.frames[src]]
        with c3:
            xl = st.file_uploader("Import frame from Excel", type=["xlsx"], key=f"fx_{q}",
                                  label_visibility="collapsed")
            if xl and st.button("Import", use_container_width=True):
                got = fr.from_excel(BytesIO(xl.getvalue()))
                st.session_state.frames[q] = list(got.values())[0]

        frame = st.session_state.frames.get(q, fr.blank_frame())
        grid = pd.DataFrame([{
            "NET": r["net"], "Code": r["code"], "Code label": r["label"],
            "Patterns": "\n".join(r["patterns"]),
        } for r in frame])

        edited = st.data_editor(
            grid, num_rows="dynamic", use_container_width=True, height=460,
            column_config={
                "Code": st.column_config.NumberColumn(width="small", format="%d"),
                "Code label": st.column_config.TextColumn(width="large"),
                "Patterns": st.column_config.TextColumn(width="large", help=PATTERN_HELP),
            }, key=f"editor_{q}")

        new_frame = []
        for _, r in edited.iterrows():
            if pd.isna(r["Code"]) or not str(r["Code"]).strip():
                continue
            new_frame.append({"net": r["NET"], "code": r["Code"],
                              "label": r["Code label"], "patterns": r["Patterns"]})
        if new_frame:
            new_frame = fr.normalise(new_frame)
            st.session_state.frames[q] = new_frame

            issues = fr.validate(new_frame)
            if issues:
                st.error("Fix before running:\n\n- " + "\n- ".join(issues))
            else:
                st.success(f"{len(new_frame)} codes, {len(set(r['net'] for r in new_frame))} NETs — frame is valid.")

        if st.session_state.frames:
            buf = BytesIO()
            fr.to_excel(st.session_state.frames, buf)
            st.download_button("⬇️ Export all frames to Excel", buf.getvalue(),
                               file_name="code_frames.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ============================================================ 3 · RUN
with tab_run:
    st.header("Run the coding")
    if not st.session_state.frames:
        st.info("Build a code frame first.")
    else:
        base_n = st.number_input("Base N for percentages (0 = use row count)", min_value=0,
                                 value=0, step=1)
        if st.button("▶️ Code all questions", type="primary"):
            for q in st.session_state.raw:
                if not st.session_state.frames.get(q):
                    st.warning(f"{q}: no frame, skipped.")
                    continue
                coded = coded_for(q)
                n = base_n or len(coded)
                freq = engine.frequency_table(coded, st.session_state.frames[q], base_n=n)

                st.subheader(st.session_state.titles[q]["text"])
                unres = len(engine.unresolved(coded, st.session_state.id_col))
                blank = sum(1 for cl in coded["Codes"] if 999 in cl)
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Responses", len(coded))
                m2.metric("Answered", len(coded) - blank)
                m3.metric("Codes used", int((freq["Count"] > 0).sum()))
                m4.metric("Needs review (996/997)", unres)

                show = freq[freq["Count"] > 0].sort_values("Count", ascending=False)
                st.dataframe(
                    show.style.format({"% of Total": "{:.1%}", "% of Answered": "{:.1%}"}),
                    use_container_width=True, height=320)

# ============================================================ 4 · REVIEW
with tab_review:
    st.header("Review and override")
    st.caption("Everything the rules could not place sits in 996 / 997. Fix those here — "
               "overrides are kept per respondent and survive a re-run of the rules.")
    if not st.session_state.frames:
        st.info("Run the coding first.")
    else:
        q = st.selectbox("Question", list(st.session_state.raw), key="rev_q")
        if st.session_state.frames.get(q):
            coded = coded_for(q)
            idc = st.session_state.id_col
            only_unres = st.checkbox("Show only 996 / 997", value=True)

            view = engine.unresolved(coded, idc) if only_unres else pd.DataFrame({
                idc: coded[idc],
                "Verbatim": coded["Verbatim"],
                "Codes": coded["Codes"].map(lambda cl: "/".join(map(str, cl))),
            })
            view = view[view["Verbatim"].str.len() > 0]

            st.write(f"{len(view)} verbatim(s)")
            edited = st.data_editor(
                view, use_container_width=True, height=430, key=f"rev_{q}",
                column_config={
                    "Verbatim": st.column_config.TextColumn(width="large", disabled=True),
                    "Codes": st.column_config.TextColumn(
                        width="medium", help="Codes separated by / — e.g. 101/204/995"),
                })

            if st.button("Save overrides", type="primary", key=f"save_{q}"):
                store = st.session_state.overrides.setdefault(q, {})
                before = dict(zip(view[idc], view["Codes"]))
                saved = 0
                for _, r in edited.iterrows():
                    txt = str(r["Codes"]).strip()
                    if txt and txt != str(before.get(r[idc], "")):
                        try:
                            store[r[idc]] = [int(x) for x in txt.replace(",", "/").split("/") if x.strip()]
                            saved += 1
                        except ValueError:
                            st.error(f"{r[idc]}: '{txt}' is not a valid code list.")
                st.success(f"{saved} override(s) saved. {len(store)} in total for {q}.")

            if st.session_state.overrides.get(q):
                st.download_button(
                    "⬇️ Export overrides (.json)",
                    json.dumps({str(k): v for k, v in st.session_state.overrides[q].items()}, indent=2),
                    file_name=f"{q}_overrides.json", mime="application/json")

# ============================================================ 5 · EXPORT
with tab_export:
    st.header("Export the deliverable")
    if not st.session_state.frames:
        st.info("Nothing to export yet.")
    else:
        base_n = st.number_input("Base N (0 = row count)", min_value=0, value=0, step=1, key="exp_n")
        if st.button("📦 Build workbook", type="primary"):
            questions = {}
            for q in st.session_state.raw:
                if not st.session_state.frames.get(q):
                    continue
                questions[q] = {
                    "title": st.session_state.titles[q]["text"],
                    "frame": st.session_state.frames[q],
                    "coded": coded_for(q),
                    "id_col": st.session_state.id_col,
                    "extra_cols": st.session_state.extra_cols,
                }
            meta = {
                "title": st.session_state.meta.get("title", "Open-End Coding Deliverable"),
                "fields": {
                    "Job number": st.session_state.meta.get("job", ""),
                    "Client": st.session_state.meta.get("client", ""),
                    "Questions coded": ", ".join(questions),
                    "Total sample (N)": base_n or len(next(iter(questions.values()))["coded"]),
                },
            }
            buf = excel.build_workbook(questions, meta, base_n=base_n or None)
            st.success("Workbook ready.")
            st.download_button(
                "⬇️ Download coded data + code structure",
                buf.getvalue(),
                file_name=f"{st.session_state.meta.get('job','project')}_OE_Coding_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.caption("Formulas are written unevaluated by openpyxl. Excel recalculates them on open; "
                       "if you need cached values in the file itself, run scripts/recalc.py over it.")
