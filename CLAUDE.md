# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

Python package (`trapo_app`, Python 3.13 only) that automates the "Trapo" (animal transport) workflow at Ein Herz für Streuner. It is a set of interactive German-language CLI tools that process Word/Excel/PDF files (transport tables from a messenger export, PetOffice exports, Traces PDFs). User-facing text, column names and comments are German; keep new prompts and output in German.

## Setup and commands

There is no test suite (`src/test.py` is empty) and no linter config.

```bash
python3.13 -m venv .venv && source .venv/bin/activate
pip install -e .          # installs the console scripts below (requirements.txt is untracked and pins versions)
```

Each CLI command is a function in `src/trapo_app/app.py`, registered in `setup.cfg` under `[options.entry_points]`. Run one during development with e.g. `python -c "from trapo_app.app import compare; compare()"` or the installed script (`trapo-vergleich`, `trapo-extrakt`, `trapo-traces-vergleich`, `trapo-traces`, `trapo-km`, `trapo-kombi`, `trapo-komplett`, `trapo-ro`, `trapo-split`, `trapo-sort`).

Adding a command means: add the function to `app.py`, add the entry point in `setup.cfg`, and list it in `main()`'s help text. The version lives in `src/trapo_app/__init__.py` (`__version__`, read by `setup.cfg`); commits bump it (`... | 1.0.22`).

## Architecture

`app.py` only orchestrates: each command prompts the user (via `input()` and Tk file dialogs), calls into the helper modules, and writes an output file (usually `.xlsx` into the current directory).

- `io_helpers.py`: all file/UI I/O. Tk file pickers (`get_file_ui`, `get_several_files_ui`), the Tk drag-to-reorder `ReorderableListApp` (used by `trapo-sort`), reading `.docx`/`.xlsx` into DataFrames (`read_file`, `read_docx`, `read_docx_with_images` for photo columns), file renaming/moving for Traces, and the Excel writers (`save_distance_sheets`, `save_ro_excel`) plus `sort_word_table`, which rewrites a Word table in place by a column and a user-defined order.
- `table_helpers.py`: DataFrame logic. Cleaning/normalizing names, DOBs, contacts, chip numbers; fuzzy matching (thefuzz/rapidfuzz) for `compare` (messenger vs PetOffice) and `compare_traces`; building Traces file names; license-plate extraction and matching (`add_plates`); distance columns (`add_distance`); Romanian header translation (`translate_headers`).
- `pdf_helpers.py`: extracts table data from Traces PDFs using camelot.
- `math_helpers.py`: address cleaning, geocoding via geopy/Photon (a custom User-Agent is required, see commit history), and driving distance via OSM (attribution required, see README).

Typical pipeline: `trapo-vergleich` → `Trapo_Vergleich.xlsx`; `trapo-extrakt` → `Traces_Extrakt.xlsx`; `trapo-traces-vergleich` combines both → `Trapo_Traces_Vergleich.xlsx`; `trapo-traces` renames/moves the Traces PDFs from that table. `trapo-komplett` is meant to chain all steps (`do_all`).

## Gotchas

- The address column in input tables must be named `KONTAKT`; the code depends on it.
- `entry_points` in `setup.cfg` reference `app.do_all`, `app.translate`, `app.split`, `app.sort_by_tp`, but `do_all` is not defined in `app.py` (so `trapo-komplett` is broken) and `split` is an unfinished `# TODO` stub.
- `trapo-sort` hardcodes `sort_column=19` (the `Treffpunkt` column index in the Word table).
- Input and output data files (`*.xlsx`, `*.docx`, `data/`) sit in the repo root but are not source; don't commit them.
