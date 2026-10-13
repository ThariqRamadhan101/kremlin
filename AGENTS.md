# AGENTS.md — kremlin (KGB pipeline)

## What this is

Three standalone Python scripts forming a KGB (Kenaikan Gaji Berkala / periodic salary
increase) pipeline against the `ehrm.pu.go.id` API. No tests, no lint, no CI, no
pyproject/tox/pre-commit. Each script runs from **its own directory** against its own
`.venv`; they share no code and are copy-edited independently.

`inquiry_kgb/` (fetch + filter + compute → Excel + `documents/`) →
`naming_kgb/` (split one PDF into per-person files by unit kerja → `kgb/`) →
`upload_kgb/` (Selenium upload into EHRM).

| Dir | Entrypoint | Real deps (what the code actually imports) |
|-----|-----------|------------------------------------------|
| `inquiry_kgb/` | `kgb_inquiry.py` + `config.py` | requests, pandas, openpyxl, **python-dateutil** (undeclared) |
| `naming_kgb/` | `naming_kgb.py` (top-level script, no functions) | **pandas, PyPDF2** (neither declared) |
| `upload_kgb/` | `upload_kgb.py` | selenium, python-dotenv, pandas |

Declared but never imported: `xlrd` (inquiry_kgb) and `webdriver-manager` (upload_kgb). Don't
rely on them; `upload_kgb` needs no webdriver-manager because it attaches to a live Chrome.

## Run

Run each from inside its own subdir — every input path is `Path.cwd()`-relative, and
`import config` resolves as a sibling. Running from the repo root silently finds nothing.

```
cd inquiry_kgb && python kgb_inquiry.py
cd naming_kgb  && python naming_kgb.py
cd upload_kgb  && python upload_kgb.py
```

### inquiry_kgb
- All tunables (token, base URL, search payload, input filenames, output prefix,
  even-year golongan list) live in `inquiry_kgb/config.py`. Change config, not the script.
- Inputs required in-dir: `Daftar Gaji PNS.xlsx`, `Daftar Gaji PPPK.xlsx`,
  `pejabat_penandatangan.json`, `pejabat.json`, `tembusan_mapping.json`,
  `urutan_unit_kerja.json`.
- Output: `kgb_processed_data_YYYYMMDD_HHMMSS.xlsx` (sheets `PNS`, `PPPK`,
  `Summary_by_Unit_Status_Pejabat`) + `documents/NNN_Name_Kenaikan {Pangkat,Gaji Berkala}.pdf`.
- `kgb_inquiry.py` is one 1698-line `class KGBInquiry`; `main()` only parses args and calls
  `run_all_phases()` (search → process → filter → kgb_data_processing → save).

### naming_kgb
- Picks the **first** `*.xlsx` and **first** `*.pdf` via unsorted `glob` — the directory must
  contain exactly the intended pair (`KGB Print.xlsx` + the scan PDF), or you split the wrong
  documents. Reads only sheet `PNS`.
- Page→person mapping is **Excel row order** (`all_records.sort(key=idx)` re-flattens the
  direktorat grouping), at `PAGES_PER_PERSON = 1`. Row order in the `PNS` sheet must match the
  PDF's physical page order.
- Output: `kgb/<unit_kerja>/<tahun>_<Bulan>_<nama>_kenaikan_gaji_berkala.pdf`. The filename
  must contain the employee's full name — `upload_kgb` globs on it.

### upload_kgb
- **Semi-automatic by design.** It attaches to a Chrome you already started and logged into:
  `chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\temp\chrome_debug"`.
  `open_website()` sets `debuggerAddress` and returns — it never navigates, and the login
  functions (`click_enter_button`, `fill_username`, …) are **not called by `main()`**, so
  `USERNAME_EHRM`/`PASSWORD_EHRM` in `.env` are read but unused. You must leave the browser
  sitting on the EHRM KGB page.
- **`pause()` per employee**: one before the confirm click
  (`Press Enter to continue to next employee anyway...`), one on any failure, one before
  closing. Never run it unattended; it also calls `driver.quit()` at the end.
- **Every page interaction is wrapped in `@manual_step`**, which catches any exception,
  prints, pauses, and returns `False`. Add that decorator to new step functions — a bare
  selector failure would otherwise abort the run. `process_employee()` checks each step's
  return and skips the row, so a failed field never gets submitted.
- **All selectors live in the `X_*` constants block** near the top; the KGB form fields
  share the `FORM` prefix. They are absolute and verified against the live site — when a
  step starts failing, fix the constant there, not inside the step.
- Excel input is the first `*.xlsx` in `os.listdir('.')` with `kgb` in the name, now
  `sorted()` for determinism, so it is not guaranteed to be `kgb_processed_data_*.xlsx`
  (e.g. `KGB Print.xlsx` also matches). Sheet is `PNS`.
- `verify_submission_success()` is unreliable by design — the EHRM table re-sorts
  asynchronously, so it can report failure for a row that saved. Confirm visually.
- `upload_kgb/README.md` still documents the pre-refactor function list and the (unused)
  webdriver-manager install step. Trust the code over it.
- `.env` is gitignored (`upload_kgb/.gitignore`) — never commit it.

## Gotchas

- **`config.py` contains a hardcoded live JWT as the `KGB_AUTH_TOKEN` fallback.** Token
  resolution is `args.token or config.AUTH_TOKEN` (kgb_inquiry.py:1657), so the script
  silently authenticates with that embedded token. It also decodes to a real NIP/email —
  treat it as a live credential, keep `config.py` untracked, and prefer passing `--token`.
- **`config.py` and `naming_kgb/requirements.txt` are untracked.** A fresh clone has
  `kgb_inquiry.py` doing `import config` with no `config.py` → `ModuleNotFoundError`.
  Conversely `inquiry_kgb/__pycache__/kgb_inquiry.cpython-313.pyc` *is* tracked and shows as
  modified on every run.
- **`naming_kgb/requirements.txt` is wrong.** It lists `pikepdf` + `PyMuPDF`; the script
  imports neither. It actually needs `pandas` + `PyPDF2`, and `pip install -r
  requirements.txt` will not install either.
- **Do not "fix" the EHRM search payload.** `SEARCH_PAGE_SIZE = 100000000000` (fetch-all in
  one call) and `SEARCH_JABATAN = "jenderal prasarana strategis"` encode the intended scope.
- **PPPK +3 years is a salary computation, not a filter.** In Phase 3, `baru_masa_kerja_tahun`
  and `acuan_terakhir_masa_kerja_tahun` get `+3` (on top of the base `+1`) only for PPPK with
  golongan II/III/IV/VI/VII/VIII (kgb_inquiry.py:1171). The KGB-eligibility filter is pure
  year-parity and has no PPPK case (kgb_inquiry.py:288). The README claims otherwise.
- **`--nip` is `type=int`**, so leading-zero 18-digit NIPs are unsupported. NIP also feeds the
  API payload directly (`"nip": f"{self.nip}"`).
- **Token expiry**: EHRM session token from DevTools → Network → request headers →
  `authorization`. Raw value, **no `Bearer` prefix** — the script sets the header directly.
  On 401/403 mint a fresh one.
- **PII**: `kgb_processed_data_*.xlsx`, `documents/`, `kgb/`, and the `Daftar Gaji`/`KGB Print`
  workbooks all hold real names, NIPs, and salaries. None are gitignored. Do not commit
  generated runs, and stage files deliberately.
- **No tests.** Verify changes by running the affected script end-to-end in its own subdir.
