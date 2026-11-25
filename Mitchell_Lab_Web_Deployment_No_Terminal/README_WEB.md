# Mitchell Lab Analysis Portal — deployable starter

The web UI is a **separate application** from the Tkinter desktop interface.
The desktop `app.py` is deliberately NOT imported by the web app; it depends
on macOS/Tk widgets and a notebook subprocess workflow.

## Current capabilities

| Program | Web state | Scientific/core code |
| --- | --- | --- |
| Convert Txt | Available; IOX schema must match | Python adapter based on supplied `app.py` and `edu_iox_dataclean_notebook.ipynb` |
| Venn Diagram Analysis | Available; requires R and actual-data testing | Supplied `venn_functions.R` + `analysis_runner.R` |
| Manhattan Plot Analysis | Available; requires R and actual-data testing | Supplied `manhattan_functions.R` + `analysis_runner.R` |
| Quiet Breathing Analysis | **Available in the web interface** | Supplied QBA Rmd with noninteractive input and protocol controls, no threshold changes |
| COSINOR Modeling | Disabled | Four supplied templates stored separately in `desktop_snapshot/`; no web execution |

**This is not a scientifically validated release.** Python compilation and adapter
checks can be run locally. R-based Venn/Manhattan/QBA need a real end-to-end
run with known validated output before being shared for research decisions.

## Quiet Breathing web workflow (enabled)

1. Open **Quiet Breathing Analysis**; upload 1–5 cleaned `.xlsx` workbooks.
2. The app reads the **Protocol Type** labels from the first worksheet of **every** selected workbook and displays three editable protocol categories side by side: **Baseline**, **Challenge**, **Excluded**.
3. Each label must belong to exactly one category. The app initially maps `Quiet Breathing` / `Normoxia Baseline` to Baseline and `Acclimation` to Excluded; other labels default to Challenge. Review and adjust before running. This selection applies to the whole batch.
4. Enter the rat light/output prefix and run. The server executes the included original `QB_Breathing_Tag_Updated_Fin.Rmd` noninteractively for each workbook, then provides a ZIP of the generated Excel/PNG outputs, generated R scripts, and logs. Individual user workbooks are omitted from the output ZIP.

No `WEB_ENABLE_QBA` environment variable is necessary. The original QBA breath classification, overlap priority and 120-second challenge-block selection are retained. The latest **Seagate** app protocol-profile persistence and recent unuploaded desktop patches are **not** assumed to be present in this web version. Scientifically compare web results to a validated standalone Rmd run before publication.

## Launch locally with Docker (recommended)

Docker keeps the R and Python dependencies together. From the unzipped master archive:

```bash
cd web_app
docker build -t mitchell-lab-web .
docker run --rm -p 8501:8501 -e APP_REQUIRE_PASSWORD=1 \
  -e APP_PASSWORD='REPLACE_WITH_A_LONG_RANDOM_PASSWORD' mitchell-lab-web
```

Open http://localhost:8501 on the same computer.

On macOS, Docker Desktop can run the Linux container. Docker must be
installed and running; the image will be larger than a Python-only web app
because it includes R.

## Launch locally without Docker

Requires Python 3.12, an accessible `Rscript` executable, and R packages
readxl, dplyr, tidyr, VennDiagram, openxlsx, tibble, ggplot2, stringr,
ggrepel, gridExtra, lubridate, purrr.

```bash
cd web_app
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements-web.txt
streamlit run streamlit_app.py
```

**This does not reuse your Seagate `.venv-external`, and it does not modify the
desktop application.** If R is missing, Venn/Manhattan/QBA will fail with a
specific missing-R error; Convert Txt uses Python only.

## Host it on a web service (e.g., Render with Docker)

1. Create a **private** Git repository. Upload **only** the contents of
   `web_app/` to that repository (or set the service's root directory to
   `web_app/`). **Do not publish the entire master ZIP or desktop snapshot.**
2. Create a Docker-based Web Service from that repository. In Render use
   **Runtime = Docker**, choose the repository and root directory if needed.
3. Set `APP_REQUIRE_PASSWORD=1` and a strong `APP_PASSWORD` as **secret**
   environment variables. Do not put credentials in Git/source files.
4. The included Dockerfile launches Streamlit on `${PORT:-8501}` and binds
   `0.0.0.0`, as Docker-based web services require. Configure health/path if
   your service requires it.
5. Deploy, then test small **de-identified** workbooks. Compare output with
   the validated desktop analyses before giving colleagues access.
6. When COSINOR is ready, add its validated adapter to the web app behind an
   independent feature switch and redeploy; other programs can stay live.

No managed domain or hosted service is created by this ZIP; deployment requires
your cloud account, repository, and explicit hosting approval.

## Security and reliability caveats

- This starter uses an optional single shared lab password. **That is not
  enterprise SSO, user roles, an audit trail, or a substitute for institution-
  approved handling of regulated research data.** For public/clinical data,
  obtain IT/privacy approval first and add institutional authentication.
- File sizes are limited to 30 MB each and R execution to 10 minutes; hosts
  may impose stricter total request or memory limits.
- Each analysis is processed under a separate temporary job folder; the
  folder is deleted after the ZIP is assembled. Results stored in Streamlit
  session state remain in the browser session until it ends. No permanent
  database or object storage is configured.
- Don't send identifiers or patient/participant datasets to an unapproved host.
- Complex graphical R jobs and uploads may require more CPU/memory than a free
  or small hosted instance provides. Heavy multi-user operation needs job
  queues, concurrency limits, logs/metrics, and a worker process.
- Dependencies are installed at **Docker image build time**, not on request.
  A real deployment should pin exact package versions after validation.
- No `.xlsx` user datasets are bundled in the master archive.

## Files

- `streamlit_app.py` — multi-program web interface and password gate
- `analysis_core.py` — IOX conversion, R runner, temporary jobs, ZIP outputs
- `assets/venn_functions.R`, `assets/manhattan_functions.R` — original R routines
- `assets/analysis_runner.R` — original Venn/Manhattan R driver
- `assets/QB_Breathing_Tag_Updated_Fin.Rmd` — original QBA source
- `Dockerfile`, `requirements-web.txt`, `.streamlit/config.toml` — deployment
- `tests/test_core.py` — non-R regression checks

## Known limitations

- The Venn web interface supports one selected 2–4 set comparison or all
  pairwise comparisons among selected sets; the desktop tool's queued custom
  comparison list is not yet replicated.
- The Manhattan web interface covers primary column mapping, threshold,
  autosome count and labeling; the desktop's full plot styling options are
  retained in R defaults but not yet exposed.
- Convert Txt reproduces the original IOX table assumptions; if cleaning
  rejects a file, the raw converted workbook and conversion report are still
  downloadable for review.
- QBA uses the original Rmd's default protocol-block renaming and the original
  sigh/post-sigh overlap priority. It provides a three-column Baseline / Challenge / Excluded selector across uploaded workbooks. The web adaptation has not been compared to
  the latest locally patched QBA code on the user's external drive.
- The separately uploaded `app.py` snapshot may predate recent Seagate patches.
