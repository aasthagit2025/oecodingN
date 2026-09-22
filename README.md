# OE Coder

A small Streamlit app for coding open-ended survey responses. Built out of the
KEKT106 SBI Life employee study, generalised so the same workflow runs on any
project: load verbatims, build a code frame, run it, review what the rules
could not place, export a client-ready workbook.

## Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

Python 3.10 or later. No database, no services — everything is in-memory plus
the JSON/Excel files you save.

## Layout

```
app.py                      Streamlit front end (5 tabs)
oecoder/engine.py           coding logic and frequency tables
oecoder/frames.py           code frame load/save/validate, Excel round-trip
oecoder/excel.py            deliverable builder
codeframes/                 saved frames, reusable across studies
requirements.txt
```

`oecoder` has no dependency on Streamlit, so the same functions drive a batch
script or a notebook if you would rather not use the UI.

## Workflow

**1 · Data** — upload the OE file. One sheet per question is the expected
shape (the layout the DP team already exports). Pick the respondent ID column,
any columns to carry through (password, cell, weight), and the verbatim column
on each sheet.

**2 · Code frame** — build the frame in the grid: NET, code number, code label,
and the patterns that assign it. Start blank, copy an existing question's
frame, import from Excel, or load a saved `.json` from the sidebar. The frame
validates as you type: duplicate codes, empty labels and broken regex are
flagged before you can run anything.

**3 · Run coding** — applies the frame to every question and shows counts,
percentages and how many verbatims need a human eye.

**4 · Review** — everything the rules could not place lands in 996 (any other)
or 997 (irrelevant). Edit codes directly in the table. Overrides are stored per
respondent ID and survive a re-run, so you can keep refining patterns without
losing manual decisions.

**5 · Export** — builds the workbook: Read Me, Code Structure (the sheet for
client sign-off), and per question a coded-data sheet and a 0/1 binary matrix
for tabulation.

## Writing patterns

Patterns are case-insensitive regular expressions, **one per line**. Plain
words work fine.

| Pattern | Matches | Notes |
|---|---|---|
| `brand name` | "the brand name", "Brand Name" | substring, case-insensitive |
| `\bpay\b` | "good pay" but not "payment" | `\b` is a word boundary |
| `stab(le\|ility)` | "stable", "stability" | alternation **inside** one pattern |
| `career\s*growth` | "career growth", "careergrowth" | `\s*` absorbs spacing |
| `^growth` | "Growth" at the start only | `^` anchors to the start |

Do not separate two patterns with `|` — that character is alternation and
belongs inside a single pattern. Put each on its own line.

Verbatims are multi-coded: every code whose pattern list gets a hit is
assigned. Codes 995–999 are reserved and applied automatically when nothing
else matches:

| Code | Meaning |
|---|---|
| 995 | Generic positive / non-specific mention only |
| 996 | Any other mention (the review queue) |
| 997 | Irrelevant / illegible |
| 998 | Don't know / Can't say / NA |
| 999 | No response (blank) |

## Reusing a frame

Save frames from the sidebar as `.json` and drop them in `codeframes/`. A
shipped example is `codeframes/KEKT106_sbi_life_employee.json` — brand
perception, employer brand, career drivers and a campaign-reaction frame, all
of which transfer to other brand-tracking or employee studies with edits to the
labels rather than a rebuild.

## Using it without the UI

```python
import pandas as pd
from oecoder import engine, frames as fr, excel

frames, meta = fr.load_json("codeframes/KEKT106_sbi_life_employee.json")
df = pd.read_excel("oe_data.xlsx", sheet_name="Q11 OE Data")

coded = engine.code_dataframe(df, id_col="Resp_Num",
                              verbatim_col=df.columns[2],
                              frame=frames["Q11"],
                              overrides={12: [101, 204]})

freq = engine.frequency_table(coded, frames["Q11"], base_n=450)
review = engine.unresolved(coded, "Resp_Num")

buf = excel.build_workbook(
    {"Q11": {"title": "Q11. What influenced your decision to join?",
             "frame": frames["Q11"], "coded": coded,
             "id_col": "Resp_Num", "extra_cols": ["Password"]}},
    meta={"title": "My study", "fields": {"Job number": "ABC123"}},
    base_n=450)
open("out.xlsx", "wb").write(buf.getvalue())
```

## Two things to keep in mind

Rules give you consistency and an audit trail, not judgement. Tab 4 is the part
that does the coding; the rules just clear the easy 80% first. On a new study,
read a sample of verbatims before writing any pattern — the frame should come
out of the data, not out of a previous project.

Excel formulas are written unevaluated by openpyxl. Excel recalculates them on
open, so the percentages are correct for anyone who opens the file. If a
downstream tool reads cached values instead, run a recalc pass over the file
first.
