# Deploying OE Coder to Streamlit Cloud

## Why the ModuleNotFoundError happened

The traceback showed `/mount/src/oecodingn/app.py` failing on
`from oecoder import engine, excel, frames as fr`. Streamlit Cloud runs your
app from the repo root, so `app.py` was found but the `oecoder` package beside
it was not importable — either the folder was never pushed, or it was pushed
without its `__init__.py`, which is what makes a folder a package.

## This version fixes it by removing the problem

`app.py` here is self-contained. There is nothing to import, so there is no
package layout to get wrong.

Your repo should look exactly like this:

```
your-repo/
├── app.py              <- this file
├── requirements.txt
└── codeframes/         <- optional, saved frames
    └── KEKT106_sbi_life_employee.json
```

Push those, point Streamlit Cloud at `app.py`, and it runs.

## Two things that catch people out

**requirements.txt must be at the repo root.** Streamlit Cloud only looks
there. If it sits in a subfolder, the app boots with no pandas or openpyxl and
fails on the first import instead of the tenth line.

**Check the file actually landed in the repo.** On GitHub, open the repo page
and confirm you can see `app.py` and `requirements.txt` at the top level. A
drag-and-drop upload of a folder sometimes uploads the folder's contents but
skips files the browser treats as hidden or empty — `__init__.py` is a common
casualty, and that is very likely what happened here.

## If you would rather keep the multi-file version

It is the same code and better for reuse from scripts or notebooks. To make it
work on Streamlit Cloud, push this structure and make sure every file is
present:

```
your-repo/
├── app.py
├── requirements.txt
└── oecoder/
    ├── __init__.py     <- this one is the usual missing piece
    ├── engine.py
    ├── frames.py
    └── excel.py
```

Verify from the command line rather than by eye:

```bash
git ls-files oecoder/
```

That should list all four files. If `__init__.py` is absent, add it:

```bash
git add -f oecoder/__init__.py
git commit -m "Add missing package init"
git push
```

The `-f` matters — some default `.gitignore` templates exclude files that look
empty or match build patterns, and that exclusion is silent.

## Running locally

```bash
pip install -r requirements.txt
streamlit run app.py
```
