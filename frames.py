# -*- coding: utf-8 -*-
"""
oecoder.frames
Load, save and validate code frames. A frame lives as JSON so it can be
version-controlled and reused; it can also be round-tripped through Excel
so a non-technical coder can edit labels and patterns in a familiar tool.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import pandas as pd

from .engine import HOUSEKEEPING, HK_CODES

PATTERN_SEP = "\n"   # newline, NOT "|" - that character is regex alternation


# --------------------------------------------------------------- validation
def validate(frame: list[dict]) -> list[str]:
    """Return a list of human-readable problems. Empty list means the frame is clean."""
    problems = []
    seen = {}
    for i, row in enumerate(frame, start=1):
        for key in ("net", "code", "label"):
            if not str(row.get(key, "")).strip():
                problems.append(f"Row {i}: '{key}' is empty.")
        try:
            code = int(row["code"])
        except (KeyError, TypeError, ValueError):
            problems.append(f"Row {i}: code '{row.get('code')}' is not a whole number.")
            continue
        if code in seen:
            problems.append(f"Code {code} appears more than once (rows {seen[code]} and {i}).")
        seen[code] = i
        if code in HK_CODES:
            problems.append(f"Row {i}: {code} is a reserved housekeeping code - it is added automatically.")
        for p in row.get("patterns") or []:
            try:
                re.compile(p)
            except re.error as e:
                problems.append(f"Code {code}: pattern {p!r} is not valid regex ({e}).")
    return problems


def normalise(frame: list[dict]) -> list[dict]:
    """Coerce types and sort by code so the frame reads cleanly."""
    out = []
    for row in frame:
        pats = row.get("patterns") or []
        if isinstance(pats, str):
            pats = [p.strip() for p in pats.splitlines() if p.strip()]
        out.append({
            "net": str(row["net"]).strip(),
            "code": int(row["code"]),
            "label": str(row["label"]).strip(),
            "patterns": [str(p).strip() for p in pats if str(p).strip()],
        })
    return sorted(out, key=lambda r: r["code"])


# --------------------------------------------------------------- JSON
def save_json(frames: dict[str, list[dict]], path: str | Path, meta: dict | None = None) -> None:
    """frames: {question_key: frame}. Saved with an optional project meta block."""
    payload = {"meta": meta or {}, "frames": {q: normalise(f) for q, f in frames.items()}}
    Path(path).write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def load_json(path: str | Path) -> tuple[dict[str, list[dict]], dict]:
    payload = json.loads(Path(path).read_text(encoding="utf-8"))
    if "frames" in payload:                       # full project file
        return {q: normalise(f) for q, f in payload["frames"].items()}, payload.get("meta", {})
    if isinstance(payload, list):                 # bare single frame
        return {"Q1": normalise(payload)}, {}
    return {q: normalise(f) for q, f in payload.items()}, {}


# --------------------------------------------------------------- Excel round-trip
def to_excel(frames: dict[str, list[dict]], path: str | Path) -> None:
    """One sheet per question, editable by hand, re-readable by from_excel."""
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for q, frame in frames.items():
            rows = [{
                "NET": r["net"],
                "Code": r["code"],
                "Code label": r["label"],
                "Patterns (one per line)": PATTERN_SEP.join(r["patterns"]),
            } for r in normalise(frame)]
            pd.DataFrame(rows).to_excel(xw, sheet_name=q[:31], index=False)


def from_excel(path: str | Path) -> dict[str, list[dict]]:
    frames = {}
    book = pd.read_excel(path, sheet_name=None)
    for sheet, df in book.items():
        cols = {c.lower().strip(): c for c in df.columns}
        pat_col = next((cols[c] for c in cols if c.startswith("pattern")), None)
        rows = []
        for _, r in df.iterrows():
            if pd.isna(r.get(cols.get("code"))):
                continue
            pats = r.get(pat_col, "") if pat_col else ""
            pats = [] if pd.isna(pats) else [p.strip() for p in str(pats).splitlines() if p.strip()]
            rows.append({
                "net": r[cols["net"]],
                "code": r[cols["code"]],
                "label": r[cols.get("code label", cols.get("label"))],
                "patterns": pats,
            })
        if rows:
            frames[sheet] = normalise(rows)
    return frames


# --------------------------------------------------------------- misc
def with_housekeeping(frame: list[dict]) -> list[dict]:
    """Frame plus the five reserved codes, for display and reporting."""
    have = {int(r["code"]) for r in frame}
    return list(frame) + [r for r in HOUSEKEEPING if r["code"] not in have]


def blank_frame() -> list[dict]:
    return [{"net": "NET 1", "code": 101, "label": "First code", "patterns": ["example|sample"]}]
