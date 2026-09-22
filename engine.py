# -*- coding: utf-8 -*-
"""
oecoder.engine
Rule-based open-end coding engine.

A code frame is a list of dicts:
    {"net": "Brand & Reputation", "code": 101, "label": "...", "patterns": ["brand", "goodwill"]}

Patterns are case-insensitive regular expressions. A verbatim is multi-coded:
every code whose pattern list has at least one hit is assigned.
"""
from __future__ import annotations

import re
from typing import Iterable, Sequence

import pandas as pd

# --------------------------------------------------------------- housekeeping
HOUSEKEEPING = [
    {"net": "Generic / Non-specific", "code": 995,
     "label": "Generic positive / non-specific mention only - no specific reason given", "patterns": []},
    {"net": "Generic / Non-specific", "code": 996,
     "label": "Any other mention", "patterns": []},
    {"net": "Generic / Non-specific", "code": 997,
     "label": "Irrelevant / illegible / does not answer the question", "patterns": []},
    {"net": "Generic / Non-specific", "code": 998,
     "label": "Don't know / Can't say / Nothing / Not applicable", "patterns": []},
    {"net": "Generic / Non-specific", "code": 999,
     "label": "No response (blank)", "patterns": []},
]
HK_CODES = {r["code"] for r in HOUSEKEEPING}

# Fallback classifiers. Override per project if a study needs different wording.
DEFAULT_GENERIC_POS = (
    r"^(it'?s |its |this is |i think it is |overall |everthing is |everything is )?"
    r"(very+\s+|really\s+|so\s+|too\s+|highly\s*,?\s*|extremely\s+|quite\s+)*"
    r"(good|nice|great|excellent|excelent|exceleent|best|super|superb|awesome|awsome|awsem|awesom|fine|"
    r"ok+|okay|okk+|wonderful|amazing|fantastic|marvellous|outstanding|cool|perfect|happy|yes+|y|"
    r"agree|agri|average|strong|strongly|high|top|better|rigorously|bravely|lovely|verry good|noce|tes|prode)"
    r"([\s\-,]*(brand|company|companey|compnay|organi[sz]ation|organaisation|experience|experiance|"
    r"feel|feeling|see|one|and better|and great|guide|very|branch name|for me|absolutely|best|good))*"
    r"[\s\.\!\,\-\u2019'\*\u00a0\U0001F300-\U0001FAFF]*$"
)

DEFAULT_NULL = (
    r"^(na|n\.?a\.?|n/a|nil|no|none|nothing|not|no comments?|no comment|don'?t know|dont know|"
    r"can'?t say|cant say|not applicable|no idea|no response|no suggestions?|unable to say)"
    r"[\s\.\!\,\-]*$"
)

DEFAULT_JUNK = (
    r"^[\s\.\,\-\_\*\?/\\]*$|^[a-z]{1,2}[\s\.]*$|^\d+[\s\.\,\-/]*\d*[\s\.\,\-/]*\d*$|"
    r"^[b-df-hj-np-tv-z]{6,}$"
)


class Rules:
    """Compiled fallback rules for one question."""

    def __init__(self, generic_pos=DEFAULT_GENERIC_POS, null=DEFAULT_NULL, junk=DEFAULT_JUNK):
        self.generic_pos = re.compile(generic_pos, re.I)
        self.null = re.compile(null, re.I)
        self.junk = re.compile(junk, re.I)


# --------------------------------------------------------------- helpers
def clean(text) -> str:
    """Normalise a raw cell into a comparable verbatim string."""
    if text is None:
        return ""
    s = str(text)
    if s.strip().lower() in ("nan", "none", "null"):
        return ""
    return re.sub(r"\s+", " ", s).strip()


def compile_frame(frame: Sequence[dict]) -> list[tuple[int, list[re.Pattern]]]:
    """Pre-compile a frame once so a 10k-row file does not recompile per row."""
    out = []
    for row in frame:
        if int(row["code"]) in HK_CODES:
            continue  # housekeeping codes are assigned by fallback, never by pattern
        pats = [p for p in (row.get("patterns") or []) if str(p).strip()]
        out.append((int(row["code"]), [re.compile(p, re.I) for p in pats]))
    return out


def code_text(text, compiled, rules: Rules | None = None) -> list[int]:
    """Return the sorted list of codes for a single verbatim."""
    rules = rules or Rules()
    t = clean(text)
    if t == "":
        return [999]
    low = t.lower()

    hits = [code for code, pats in compiled if any(p.search(low) for p in pats)]
    if hits:
        return sorted(set(hits))

    if rules.null.match(t):
        return [998]
    if rules.generic_pos.match(t):
        return [995]
    if rules.junk.match(t) or len(re.sub(r"[^a-z]", "", low)) <= 2:
        return [997]
    return [996]


def code_series(values: Iterable, frame: Sequence[dict], rules: Rules | None = None) -> list[list[int]]:
    compiled = compile_frame(frame)
    rules = rules or Rules()
    return [code_text(v, compiled, rules) for v in values]


def apply_overrides(ids: Iterable, codes: list[list[int]], overrides: dict) -> list[list[int]]:
    """overrides: {respondent_id: [codes]}. Keys are matched as-is and as strings."""
    out = []
    for rid, cl in zip(ids, codes):
        if rid in overrides:
            out.append(sorted(set(int(c) for c in overrides[rid])))
        elif str(rid) in overrides:
            out.append(sorted(set(int(c) for c in overrides[str(rid)])))
        else:
            out.append(cl)
    return out


def code_dataframe(df: pd.DataFrame, id_col: str, verbatim_col: str,
                   frame: Sequence[dict], overrides: dict | None = None,
                   rules: Rules | None = None) -> pd.DataFrame:
    """Return a copy of df with normalised Verbatim and a Codes column."""
    out = df.copy()
    out["Verbatim"] = out[verbatim_col].map(clean)
    codes = code_series(out["Verbatim"], frame, rules)
    if overrides:
        codes = apply_overrides(out[id_col], codes, overrides)
    out["Codes"] = codes
    return out


# --------------------------------------------------------------- reporting
def frequency_table(coded: pd.DataFrame, frame: Sequence[dict], base_n: int | None = None) -> pd.DataFrame:
    """Counts and percentages for every code in the frame, in frame order."""
    from collections import Counter

    counts = Counter(c for cl in coded["Codes"] for c in cl)
    n = base_n or len(coded)
    answered = sum(1 for cl in coded["Codes"] if 999 not in cl)

    rows = []
    for row in list(frame) + [r for r in HOUSEKEEPING
                              if int(r["code"]) not in {int(x["code"]) for x in frame}]:
        code = int(row["code"])
        cnt = counts.get(code, 0)
        rows.append({
            "NET": row["net"],
            "Code": code,
            "Code label": row["label"],
            "Count": cnt,
            "% of Total": cnt / n if n else 0,
            "% of Answered": cnt / answered if answered else 0,
        })
    return pd.DataFrame(rows)


def unresolved(coded: pd.DataFrame, id_col: str) -> pd.DataFrame:
    """Rows that landed in 996/997 - the analyst review queue."""
    mask = coded["Codes"].map(lambda cl: bool({996, 997} & set(cl)))
    out = coded.loc[mask, [id_col, "Verbatim", "Codes"]].copy()
    out["Codes"] = out["Codes"].map(lambda cl: "/".join(map(str, cl)))
    return out
