# Mitchell Lab Project — master source package and web starter

**Created 2026-11-25** from files available in this conversation.

This ZIP contains a copy of essential desktop source code and analysis files,
plus a new, independently deployable **web-app starter**. It does **not**
contain the Mac `.venv-external` environment, installed R/Python package
binaries, a PyInstaller-built `.app`, user datasets, any production user
accounts, or a running hosted site. Those are machine/hosting-specific.

## Directory guide

- `desktop_snapshot/` — uploaded `app.py`, `mylauncher`, `.spec`, original
  desktop `requirements.txt`, notebook, QBA Rmd, COSINOR helpers, and the latest
  four `(2)` COSINOR templates **renamed to the canonical filenames** the app
  loads. It also contains extracted Venn/Manhattan source modules and their
  R assets. `reference_working/ADORA1_E3_E4.Rmd` is included only as a
  reference example. No Excel test data is included.
- `web_app/` — separate Streamlit-based upload/analysis/download portal,
  R processing assets, Dockerfile, Python requirements, tests, and web deployment
  instructions in `README_WEB.md`.
- `scripts/backup_current_mac.py` — optional source-only backup tool for the **actual latest** Seagate app.
- `SHA256_MANIFEST.txt` — per-file checksums for audit/provenance.

## IMPORTANT: actual Seagate app may be newer

The `desktop_snapshot/app.py` file in this package is the latest **uploaded
snapshot available in the conversation**, not a verified copy of the macOS
`/Volumes/Seagate/App/app.py` currently being used. Since the upload, the
Seagate app has received multiple Terminal patches (range picker, settings,
Group 1/2, template handling, helper paths and other corrections).

**DO NOT overwrite your current `/Volumes/Seagate/App/app.py` with the copy
in this ZIP.** To create a complete up-to-date disaster-recovery archive,
copy your actual Seagate `app.py`, locally installed template files,
`lab_analysis_programs.py`, `analysis_assets/`, identifier registry, protocol
profiles, and any other added code into your own secured backup separately.

An included optional helper creates that current-source backup safely:

```bash
python3 scripts/backup_current_mac.py "/Volumes/Seagate/App" "$HOME/Desktop/Mitchell_Current_Source.zip"
```

Run it from the extracted master folder. It does not change the app or include
workbooks, outputs, binaries, or virtual environments.


## Scientific status

| Program | Web exposure | Status |
| --- | --- | --- |
| Convert Txt | On | IOX notebook logic adapted; schema-specific |
| Venn | On | R functions from original desktop extension; needs validation |
| Manhattan | On | R functions from original desktop extension; needs validation |
| Quiet Breathing | On | Original Rmd adapted for web; Baseline/Challenge/Excluded controls, needs output comparison |
| COSINOR | Disabled | Templates/helpers archived; continue desktop development |

QBA is **enabled** in the web UI with a three-column Baseline/Challenge/Excluded protocol classifier and isolated per-workbook R execution. No QBA feature flag is needed; only COSINOR is paused.

The web version has independent code paths and does not modify your Mac app.
Do not assume web and desktop outputs are scientifically equivalent until
compared with known reference workbooks. Do not publish research/participant
information on an unapproved host.

## Quick start

From this extracted ZIP:

```bash
cd web_app
docker build -t mitchell-lab-web .
docker run --rm -p 8501:8501 -e APP_REQUIRE_PASSWORD=1 \
  -e APP_PASSWORD='CHOOSE_A_UNIQUE_LONG_SECRET' mitchell-lab-web
```

Browse to http://localhost:8501. See `web_app/README_WEB.md` for a cloud
hosting workflow (Docker web service) and access-control/security caveats.

No COSINOR changes or patches are applied to the Seagate drive by extracting
this ZIP. No local installation or account linking is performed automatically.
