"""Server-side adapters for the supplied Mitchell Lab analysis programs.

Each execution gets an isolated TemporaryDirectory; user uploads are not retained by
this app after a result is generated. R code loaded from project assets only.
"""
from __future__ import annotations

import hashlib
import io
import itertools
import json
import math
import os
from pathlib import Path
import re
import subprocess
import tempfile
import zipfile

import pandas as pd
from openpyxl import Workbook

APP_ROOT = Path(__file__).resolve().parent
ASSETS = APP_ROOT / 'assets'
ALLOWED_EXTS = {'.xlsx', '.xls'}
MAX_UPLOAD_BYTES = 30 * 1024 * 1024
MAX_R_SECONDS = 600


class AnalysisError(RuntimeError):
    pass


def validate_files(files, *, kinds=ALLOWED_EXTS, max_count=12):
    if not files or len(files) > max_count:
        raise AnalysisError(f'Upload between 1 and {max_count} files.')
    for name, data in files:
        suffix = Path(name).suffix.lower()
        if suffix not in kinds:
            raise AnalysisError(f'Unsupported extension: {name}')
        if not data or len(data) > MAX_UPLOAD_BYTES:
            raise AnalysisError(f'File empty or exceeds 30 MB: {name}')


def safe_file_name(name):
    name = Path(name.replace('\\', '/')).name
    name = re.sub(r'[^A-Za-z0-9_. -]+', '_', name).strip(' .')
    if not name or name.startswith('.') or len(name) > 120:
        raise AnalysisError('File name is invalid or too long.')
    return name


def staged_inputs(folder: Path, files):
    dest = folder / 'inputs'
    dest.mkdir(parents=True, exist_ok=True)
    paths = []
    for i, (filename, data) in enumerate(files, 1):
        p = dest / f'{i:02d}_{safe_file_name(filename)}'
        p.write_bytes(data)
        paths.append(p)
    return paths


def r_quote(value):
    if value is None:
        return 'NULL'
    if isinstance(value, bool):
        return 'TRUE' if value else 'FALSE'
    if isinstance(value, (int, float)):
        if not math.isfinite(value):
            raise AnalysisError('Invalid numeric configuration.')
        return str(value)
    if isinstance(value, str):
        # JSON strings are valid R string literals for these path/label values.
        return json.dumps(value, ensure_ascii=False)
    if isinstance(value, (list, tuple)):
        return 'list(' + ','.join(r_quote(v) for v in value) + ')'
    if isinstance(value, dict):
        if any(not re.fullmatch('[A-Za-z][A-Za-z_0-9]*', key) for key in value):
            raise AnalysisError('Invalid R configuration key.')
        return 'list(' + ','.join(f'{key}={r_quote(val)}' for key, val in value.items()) + ')'
    raise AnalysisError('Unsupported configuration type.')


def run_r(job: Path, config: dict, *, stage: str):
    rscript = os.getenv('RSCRIPT_PATH', 'Rscript')
    config_file = job / f'config_{stage}.R'
    config_file.write_text('cfg <- ' + r_quote(config) + '\n', encoding='utf-8')
    runner = ASSETS / 'analysis_runner.R'
    try:
        completed = subprocess.run(
            [rscript, '--vanilla', str(runner), str(config_file)],
            cwd=job, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, timeout=MAX_R_SECONDS, check=False,
            env={**os.environ, 'R_DEFAULT_INTERNET_TIMEOUT': '60'},
        )
    except FileNotFoundError as e:
        raise AnalysisError('Rscript is not installed. Use the provided Docker image.') from e
    except subprocess.TimeoutExpired as e:
        raise AnalysisError('R analysis exceeded the 10-minute limit.') from e
    output = completed.stdout[-30_000:]
    (job / f'run_{stage}.log').write_text(output, encoding='utf-8')
    if completed.returncode:
        raise AnalysisError(f'R {stage} failed:\n{output[-5000:]}')
    return output


def zipped_results(job: Path, *, exclude_inputs=True):
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as z:
        for p in sorted(job.rglob('*')):
            if not p.is_file():
                continue
            relative = p.relative_to(job)
            if exclude_inputs and relative.parts[0] == 'inputs':
                continue
            z.write(p, relative.as_posix())
    buffer.seek(0)
    return buffer.getvalue()


BASE = {
    'sheets': [], 'prefix': 'Mitchell', 'dpi': 220,
    'min_width': 8.0, 'max_width': 16.0,
    'min_height': 8.0, 'max_height': 14.0,
    'label_size': 1.0,
}


def validate_stub(s):
    if not isinstance(s, str) or not re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9_-]{0,45}', s):
        raise AnalysisError('Output prefix must use letters, numbers, hyphens or underscores (up to 46).')
    return s


def venn_config(job, files, prefix, sheets=None, *, id_column='', gene_column='', dpi=220):
    cfg = dict(BASE)
    cfg.update(kind='venn', stage='prepare', files=[str(p) for p in files],
               job=str(job), assets=str(ASSETS), prefix=validate_stub(prefix),
               sheets=sheets or [], id_column=id_column.strip(), gene_column=gene_column.strip(),
               excel_folder='Sets and Intersections', plot_folder='Venn Diagrams',
               colors=['#999999', '#E69F00', '#56B4E9', '#009E73'],
               count_size=3.0, dpi=int(dpi))
    return cfg


def prepare_venn(files, prefix='Venn', sheets=None, id_column='', gene_column=''):
    validate_files(files)
    with tempfile.TemporaryDirectory(prefix='mitchell_venn_preview_') as folder:
        job = Path(folder)
        paths = staged_inputs(job, files)
        cfg = venn_config(job, paths, prefix, sheets, id_column=id_column, gene_column=gene_column)
        run_r(job, cfg, stage='prepare')
        manifest = pd.read_csv(job / 'sets.tsv', sep='\t')
        return manifest.to_dict('records')


def venn_run(files, selected, *, prefix='Venn', sheets=None, id_column='', gene_column='', all_pairs=False):
    validate_files(files)
    if not 2 <= len(selected) <= 4:
        raise AnalysisError('Select 2–4 groups for a Venn diagram.')
    with tempfile.TemporaryDirectory(prefix='mitchell_venn_') as folder:
        job = Path(folder)
        paths = staged_inputs(job, files)
        cfg = venn_config(job, paths, prefix, sheets, id_column=id_column, gene_column=gene_column)
        run_r(job, cfg, stage='prepare')
        manifest = pd.read_csv(job / 'sets.tsv', sep='\t')
        labels = manifest['name'].astype(str).tolist()
        if len(set(labels)) != len(labels):
            raise AnalysisError('Set names are not unique.')
        if any(name not in labels for name in selected):
            raise AnalysisError('Selected groups changed since preview; preview again.')
        idx = [labels.index(s) + 1 for s in selected]
        combinations = list(itertools.combinations(idx, 2)) if all_pairs else [tuple(idx)]
        if len(combinations) > 20:
            raise AnalysisError('Too many pairwise comparisons; limit input sets.')
        cfg.update(stage='run', aliases=labels,
                   comparisons=[{'indices': list(pair), 'name': f'comparison_{n:02d}'}
                                for n, pair in enumerate(combinations, 1)])
        run_r(job, cfg, stage='run')
        return zipped_results(job)


def manhattan_config(job, files, *, prefix, columns, autosomes, sheets=None, cutoff=0.05, top_labels=10):
    if len(columns) != 4 or len(set(columns)) != 4 or any(not x for x in columns):
        raise AnalysisError('Specify four distinct column names.')
    if not (1 <= int(autosomes) <= 100 and 0 < cutoff <= 1 and 0 <= int(top_labels) <= 250):
        raise AnalysisError('Invalid Manhattan analysis parameters.')
    cfg = dict(BASE)
    cfg.update(kind='manhattan', stage='run', files=[str(p) for p in files],
               job=str(job), assets=str(ASSETS), prefix=validate_stub(prefix),
               sheets=sheets or [], columns=columns, autosomes=int(autosomes),
               cutoff=float(cutoff), top_labels=int(top_labels),
               colors=['#4D4D4D','#87CEEB'], threshold_color='#FF0000',
               alpha=.75, point_size=.8, label_size=3.0, font_size=14,
               chr_width=1., chr_gap=.12, suffix='DNA_methylation_manhattan',
               title='DNA Methylation Manhattan Plot',
               x_label='Chromosome', y_label='-log10(Adjusted P-value)',
               min_width=10., max_width=24., min_height=5., max_height=12.,
               dpi=220)
    return cfg


def manhattan_run(files, *, prefix='Manhattan', columns=('chr','pos','adj.P.Val','CpG'),
                  autosomes=22, sheets=None, cutoff=.05, top_labels=10):
    validate_files(files)
    with tempfile.TemporaryDirectory(prefix='mitchell_manhattan_') as folder:
        job = Path(folder)
        paths = staged_inputs(job, files)
        cfg = manhattan_config(job, paths, prefix=prefix, columns=list(columns),
                               autosomes=autosomes, sheets=sheets,
                               cutoff=cutoff, top_labels=top_labels)
        run_r(job, cfg, stage='run')
        return zipped_results(job)


def parse_iox_text(text):
    lines = [line.rstrip('\r\n') for line in text.splitlines() if line.strip()]
    if not lines:
        raise AnalysisError('Text file is empty.')
    sample = '\n'.join(lines[:10])
    delimiter = next((d for d in ('\t', ',', ';') if d in sample), None)
    return [line.split(delimiter) if delimiter else line.split() for line in lines]


def write_raw_io_text(rows, path):
    wb = Workbook(); ws = wb.active
    for ri, row in enumerate(rows, start=1):
        for ci, raw in enumerate(row, start=1):
            if ri == 1:
                val = str(raw)
            else:
                value = str(raw).strip()
                if value and re.fullmatch(r'[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?', value):
                    try: val = float(value)
                    except ValueError: val = raw
                else: val = raw
            ws.cell(row=ri, column=ci, value=val)
    wb.save(path)


def clean_io_excel(path):
    """Faithful adapter of notebook's cleanit() logic, with explicit shape errors."""
    df = pd.read_excel(path)
    if df.shape[1] < 21:
        raise AnalysisError('IOX workbook does not have the 21+ columns expected by the cleaning notebook.')
    df.columns.values[6] = 'c7'
    matches = df[df['c7'].astype(str).str.contains('parameter', case=False, na=False)].index
    if len(matches) == 0:
        raise AnalysisError('Could not find the IOX parameter row in column 7.')
    df = pd.read_excel(path, skiprows=int(matches[0]) + 1)
    if df.shape[1] < 21:
        raise AnalysisError('Workbook after the parameter row has insufficient columns.')
    df.drop(df.tail(6).index, inplace=True)
    df.rename(columns={df.columns[5]: 'c5'}, inplace=True)
    df = df[~df['c5'].astype(str).str.contains('analyzer|period-stop', regex=True, na=False)]
    df.rename(columns={df.columns[6]: 'c6', df.columns[8]: 'c8'}, inplace=True)
    df['c8'] = df.groupby(df['c6'].notnull().cumsum())['c6'].transform(lambda x: x.ffill())
    df = df.drop(columns=df.columns[0]); df = df.drop(columns=df.columns[4:7])
    if len(df) < 4: raise AnalysisError('Too few rows remain after IOX cleaning.')
    df.index = list(range(4)) + list(df.index[4:])
    df = df.drop(labels=range(0,4), axis=0)
    if df.shape[1] < 21: raise AnalysisError('IOX columns do not match the expected export structure.')
    df.columns.values[5] = 'c5'
    df = df.dropna(subset=['c5'])
    df.insert(0, 'Sample', df['c5'])
    names = {0:'Sample',1:'CPU Date',2:'CPU Time',3:'Site Time',4:'Period Time',
             5:'Protocol Type',6:'Storage ID',7:'First Beat ID',8:'Last Beat ID',
             9:'Ti (msec)',10:'Te (msec)',15:'RT (msec)',17:'P (msec)',
             18:'f (bpm)',19:'EIP (msec)',20:'EEP (msec)'}
    for index, label in names.items(): df.columns.values[index] = label
    return df.reset_index(drop=True)


def convert_io(files):
    validate_files(files, kinds={'.txt','.iox'}, max_count=10)
    with tempfile.TemporaryDirectory(prefix='mitchell_io_') as folder:
        job = Path(folder); outputs = job / 'converted'; outputs.mkdir()
        report = []
        for i, (name, data) in enumerate(files, 1):
            decoded = data.decode('utf-8-sig', errors='replace')
            rows = parse_iox_text(decoded)
            stem = safe_file_name(name).removesuffix('.txt').removesuffix('.iox')
            raw = outputs / f'{i:02d}_{stem}.xlsx'
            clean = outputs / f'{i:02d}_{stem}_cleaned.xlsx'
            write_raw_io_text(rows, raw)
            try:
                clean_io_excel(raw).to_excel(clean, index=False)
                report.append(f'{name}: cleaned successfully')
            except Exception as exc:
                report.append(f'{name}: raw conversion only; cleaning failed: {exc}')
        (job / 'conversion_report.txt').write_text('\n'.join(report)+'\n', encoding='utf-8')
        return zipped_results(job), report


def make_qba_script(template, path, prefix, baseline, excluded):
    """Adapt the original QBA Rmd for a non-interactive R process.

    No changes to thresholds, tagging, or class-priority rules.
    """
    chunks = re.findall(r'```\{r[^}]*\}\s*(.*?)```',template,re.I|re.S)
    if not chunks: raise AnalysisError('QBA Rmd has no runnable code chunks.')
    script = '\n\n'.join(chunks[1:])  # Skip knitr setup chunk.
    needle = 'file_path <- file.choose()'
    if script.count(needle) != 1: raise AnalysisError('QBA file input anchor changed.')
    script = script.replace(needle, 'file_path <- '+r_quote(str(path)))
    needle = 'rat_light_type <- readline(prompt = "Enter Rat Light Type: ")'
    if script.count(needle) != 1: raise AnalysisError('QBA prefix prompt anchor changed.')
    script = script.replace(needle, 'rat_light_type <- '+r_quote(prefix))
    needle = 'baseline_protocol <- c("Quiet Breathing", "Normoxia Baseline")'
    if script.count(needle) != 1: raise AnalysisError('QBA baseline protocol anchor changed.')
    script = script.replace(needle, 'baseline_protocol <- unlist('+r_quote(baseline)+')')
    needle = 'protocols_to_exclude <- c(baseline_protocol, "Acclimation")'
    if script.count(needle) != 1: raise AnalysisError('QBA exclusion protocol anchor changed.')
    script = script.replace(needle, 'protocols_to_exclude <- c(baseline_protocol, unlist('+r_quote(excluded)+'))')
    # Suppress only the OS-specific manual Excel-preview prompt block.
    open_pos = script.find('if (.Platform$OS.type == "windows") {')
    end_pos = script.find('# Optionally clean up the temporary file', open_pos)
    if open_pos == -1 or end_pos == -1:
        raise AnalysisError('QBA interactive preview anchor changed.')
    script = script[:open_pos] + 'cat("Web QBA: manual preview skipped.\\n")\n' + script[end_pos:]
    # Skip CRAN installs during an analysis; packages are installed in Docker image.
    script = script.replace('install_if_missing(required_packages)',
                            'invisible(lapply(required_packages, function(p) { if (!requireNamespace(p, quietly=TRUE)) stop(paste("Missing R package:",p)) }))')
    return script


def qba_protocol_labels(files):
    """Inspect the first worksheet of every uploaded QB workbook, just as R does."""
    validate_files(files, kinds={'.xlsx'}, max_count=5)
    labels = []
    for name, payload in files:
        try:
            protocol = pd.read_excel(io.BytesIO(payload), sheet_name=0, usecols=['Protocol Type'])
        except Exception as exc:
            raise AnalysisError(
                f'{name}: cannot read a Protocol Type column from the first worksheet: {exc}'
            ) from exc
        for label in protocol['Protocol Type'].dropna().unique():
            value = str(label)
            if value and value not in labels:
                labels.append(value)
    if not labels:
        raise AnalysisError('No nonempty Protocol Type values found in the input workbooks.')
    return labels


def qba_run(files, *, prefix, baseline, excluded, challenge=None):
    """Run the original QBA Rmd with explicit, validated protocol categories.

    Challenge blocks are intentionally calculated by the original R code as
    the complement of baseline and excluded labels. The thresholds, QB tagging
    and breath-classification priority are not modified.
    """
    validate_files(files, kinds={'.xlsx'}, max_count=5)
    prefix = validate_stub(prefix)
    available = set(qba_protocol_labels(files))
    baseline, excluded = list(baseline), list(excluded)
    if challenge is None:
        challenge = sorted(available - set(baseline) - set(excluded))
    else:
        challenge = list(challenge)
    classifications = baseline + challenge + excluded
    if not baseline:
        raise AnalysisError('Classify at least one protocol as Baseline.')
    if len(classifications) != len(set(classifications)):
        raise AnalysisError('Each protocol label must belong to exactly one category.')
    missing, unknown = available - set(classifications), set(classifications) - available
    if missing or unknown:
        raise AnalysisError(
            'Protocol classification mismatch. Unassigned: '
            + repr(sorted(missing)) + '; unknown: ' + repr(sorted(unknown))
        )
    with tempfile.TemporaryDirectory(prefix='mitchell_qba_') as folder:
        job=Path(folder); inputs=staged_inputs(job,files)
        template=(ASSETS/'QB_Breathing_Tag_Updated_Fin.Rmd').read_text(encoding='utf-8')
        for i, path in enumerate(inputs, 1):
            output=job / f'qba_{i:02d}';output.mkdir()
            script=make_qba_script(template, path, prefix, baseline, excluded)
            (output/'generated_qba.R').write_text(script, encoding='utf-8')
            try:
                completed=subprocess.run([os.getenv('RSCRIPT_PATH','Rscript'),'--vanilla','generated_qba.R'],
                                         cwd=output,stdout=subprocess.PIPE,stderr=subprocess.STDOUT,
                                         timeout=MAX_R_SECONDS,check=False,text=True)
            except FileNotFoundError as exc:
                raise AnalysisError('Rscript is not available.') from exc
            except subprocess.TimeoutExpired as exc:
                raise AnalysisError('QBA exceeded the time limit.') from exc
            (output/'run.log').write_text(completed.stdout[-30000:], encoding='utf-8')
            if completed.returncode:
                raise AnalysisError('QBA failed: '+completed.stdout[-5000:])
        return zipped_results(job)
