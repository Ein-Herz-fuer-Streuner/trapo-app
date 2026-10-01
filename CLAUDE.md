# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

Python package (`trapo_app`, Python 3.13+) that automates the "Trapo" (animal transport) workflow at Ein Herz für Streuner. It is a set of interactive German-language CLI tools that process Word/Excel/PDF files (transport tables from a messenger export, PetOffice exports, Traces PDFs). User-facing text, column names and comments are German; keep new prompts and output in German.

## Setup and commands

```bash
python3.13 -m venv .venv && source .venv/bin/activate
pip install -e ".[dev]"   # installs the console scripts below plus pytest
pytest                    # all tests; single test: pytest tests/test_table_helpers.py::TestCompareContact
```

Tests need no Tk and no network (`tests/conftest.py` stubs `tkinter` when it is missing, HTTP calls are monkeypatched). Sample data (`*.xlsx`/`*.docx` in the repo root) is not used by the tests.

Each CLI command is a function in `src/trapo_app/app.py`, registered in `setup.cfg` under `[options.entry_points]` (`trapo-vergleich`, `trapo-extrakt`, `trapo-traces-vergleich`, `trapo-traces`, `trapo-km`, `trapo-kombi`, `trapo-ro`, `trapo-sort`, `trapo-split`). Run one during development with e.g. `python -c "from trapo_app.app import compare; compare()"`.

Adding a command means: add the function to `app.py` (decorate with `@cli_command`), add the entry point in `setup.cfg`, and list it in `main()`'s help text. The version lives in `src/trapo_app/__init__.py` (`__version__`, read by `setup.cfg` and used in the geocoding User-Agent); commits bump it (`... | 1.0.22`).

## Architecture

`app.py` only orchestrates: each command prompts the user (via `print` and Tk file dialogs), calls into the helper modules, and writes an output file (usually `.xlsx` into the current directory). Helpers never exit the program; they raise `errors.TrapoError` with a German user-facing message, and the `@cli_command` decorator prints it and exits with code 1.

- `gui.py`: everything Tk. File pickers (`get_file_ui`, `get_several_files_ui`) and the drag-to-reorder `ReorderableListApp` used by `trapo-sort`, plus `PartNamesApp`/`AssignmentApp` used by `trapo-split`. Only `app.py` imports it, so the rest of the package works without Tk.
- `io_helpers.py`: reading `.docx`/`.xlsx`/`.csv` into DataFrames (`read_file`, `read_docx_with_images` for photo columns), renaming/moving Traces files, and `sort_word_table`, which reorders a Word table's rows in place by a column and a user-defined order, and `split_word_table`, which writes one copy of the document per part (`<name>_<Teil>.docx`) keeping only that part's rows.
- `excel_writers.py`: all Excel output through xlsxwriter (`write_df_to_excel`, `save_distance_sheets` for `trapo-km`, `save_ro_excel` for `trapo-ro`).
- `table_helpers.py`: DataFrame logic. Cleaning names, DOBs, contacts and chips; `compare` (messenger vs PetOffice) and `compare_traces`, both built on `compare_contact`/`match_pet`; Traces file name building; license-plate extraction and fuzzy matching (`add_plates`, rapidfuzz); distance columns and sorting (`add_distance`, `insert_headers`); Romanian header translation (`translate_headers`).
- `pdf_helpers.py`: extracts table data from Traces PDFs using camelot.
- `math_helpers.py`: address cleaning, geocoding via Photon and driving distance via OSRM. Requests are throttled to one per second and cached per run; a custom User-Agent is required, and OSM attribution is required (see README).

Typical pipeline: `trapo-vergleich` → `Trapo_Vergleich.xlsx`; `trapo-extrakt` → `Traces_Extrakt.xlsx`; `trapo-traces-vergleich` combines both → `Trapo_Traces_Vergleich.xlsx`; `trapo-traces` renames/moves the Traces PDFs from that table.

## Gotchas

- The address column in input tables must be named `KONTAKT`; the code depends on it.
- `find_stopp_for_plate` accumulates its 5-match threshold (`MIN_NAME_MATCHES`) across all documents of a plate; this looks odd but is unchanged behaviour.
- Input and output data files (`*.xlsx`, `*.docx`, `data/`) sit in the repo root but are not source; don't commit them.
