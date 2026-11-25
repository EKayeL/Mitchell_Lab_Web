import sys
import os
import json
import threading
import subprocess
import shutil
import re
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, filedialog
import tkinter.font as tkfont

# Optional drag+drop support
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES  # pip install tkinterdnd2
    BaseTk = TkinterDnD.Tk
except Exception:
    DND_FILES = None
    BaseTk = tk.Tk

import openpyxl  # pip install openpyxl
import pandas as pd  # pip install pandas
import numpy as np  # pip install numpy
# notebook imports scipy in cell 0
from scipy.stats import zscore, percentileofscore  # pip install scipy
from scipy.spatial.distance import mahalanobis


# -----------------------------
# Styling constants
# -----------------------------
DEFAULT_BG = "#2B4C7E"
BTN_BG = "#E7EEF9"
BTN_FG = "#0B1F3B"
BTN_ACTIVE_BG = "#D7E3F7"
BTN_BORDER = "#3C3C3C"

DROPDOWN_BORDER = "#203656"
DROPDOWN_ARROW_BG = "#D0D6E2"
DROPDOWN_ARROW_FG = "#0B1F3B"
PLACEHOLDER_FG = "#6B7280"

OUTPUT_FOLDER_NAME = "App QB XLSX files"
QBA_OUTPUT_FOLDER_NAME = "App QBA Outputs"

NOTEBOOK_FILENAME = "edu_iox_dataclean_notebook.ipynb"
QBA_RMD_FILENAME = "QB_Breathing_Tag_Updated_Fin.Rmd"
QBA_RMD_DOCX_FILENAME = "QB_Breathing_Tag_Updated_Fin.Rmd.docx"


COSINOR_OUTPUT_FOLDER_NAME = "App COSINOR Outputs"
COSINOR_RMD_FILENAME = "Cosinor_Sample_Markdown.Rmd"
COSINOR_RMD_NO_OUTLIER_FILENAME = "Cosinor_Sample_Markdown_NoOutlier.Rmd"
COSINOR_RMD_CLOCK_FILENAME = "Cosinor_Sample_Markdown_Clock.Rmd"
COSINOR_RMD_CLOCK_NO_OUTLIER_FILENAME = "Cosinor_Sample_Markdown_Clock_NoOutlier.Rmd"


# -----------------------------
# i18n (kept from prior version; minimal additions for new screens)
# -----------------------------
I18N = {
    "en": {
        "title": "My Launcher",
        "run_analysis": "Run Analysis",
        "settings": "Settings",
        "exit": "Exit",
        "font_size": "Font Size",
        "background": "Background",
        "language": "Language",
        "back": "Back",
        "accept": "Accept",
        "text": "Text",

        "select_protocol": "Select protocol...",
        "protocol_convert": "Convert Txt (Cleancode)",
        "protocol_qb": "Quiet Breathing Analysis",
        "protocol_cosinor": "COSINOR Modeling",

        "upload_files": "Upload file(s)",
        "run_files": "Run File(s)",
        "clear": "Clear",
        "no_files": "No files selected.",

        "file_prefix": "File Prefix",
        "convert_to_qba": "Convert to QBA",
        "err_no_prefix": "ERROR: No File Prefix Selected",
        "err_no_eligible": "ERROR: No Files Eligible for Conversion",
    }
}

# For now, keep other languages as previously implemented in your version.
# If you want, we can add translated protocol labels too.
LANG_OPTIONS = [("English", "en")]

BG_OPTIONS = [
    ("#2B4C7E", "#2B4C7E"),
    ("#1B2A49", "#1B2A49"),
    ("#3F5E8C", "#3F5E8C"),
    ("#2A7F7F", "#2A7F7F"),
    ("#2F6B3F", "#2F6B3F"),
    ("#6B2F3F", "#6B2F3F"),
    ("#4B2F6B", "#4B2F6B"),
    ("#2E2E2E", "#2E2E2E"),
    ("#2C3E50", "#2C3E50"),
    ("#34495E", "#34495E"),
]


# -----------------------------
# Helpers
# -----------------------------
def resource_path(rel_name: str) -> Path:
    """Get path to bundled resource (works in dev and PyInstaller)."""
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / rel_name
    return Path(__file__).resolve().parent / rel_name


def strip_iox_extensions(name: str) -> str:
    lower = name.lower()
    # handle .iox.txt (or .IOX.TXT)
    if lower.endswith(".iox.txt"):
        return name[:-len(".iox.txt")]
    if lower.endswith(".txt"):
        name = name[:-len(".txt")]
        lower = name.lower()
    if lower.endswith(".iox"):
        return name[:-len(".iox")]
    return Path(name).stem


def guess_delimiter(sample: str) -> str:
    if "\t" in sample:
        return "\t"
    if "," in sample:
        return ","
    if ";" in sample:
        return ";"
    return None  # whitespace-ish


def parse_text_to_rows(text: str):
    lines = [ln.rstrip("\n\r") for ln in text.splitlines() if ln.strip() != ""]
    if not lines:
        return []

    sample = "\n".join(lines[:10])
    delim = guess_delimiter(sample)

    rows = []
    for ln in lines:
        if delim:
            rows.append(ln.split(delim))
        else:
            # fallback: split on runs of whitespace
            rows.append(ln.split())
    return rows


def write_rows_to_xlsx(rows, out_path: Path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    num_re_int = re.compile(r"^-?\d+$")
    num_re_float = re.compile(r"^-?(?:\d+\.\d*|\d*\.\d+)(?:[eE]-?\d+)?$|^-?\d+(?:[eE]-?\d+)$")

    def coerce(v):
        # Keep None/blank as None
        if v is None:
            return None
        if isinstance(v, (int, float)):
            return v
        s = str(v).strip()
        if s == "":
            return None
        s2 = s.replace(",", "")
        # Only coerce if it's purely numeric (avoid things like '12 ms')
        if num_re_int.match(s2):
            try:
                return int(s2)
            except Exception:
                return s
        if num_re_float.match(s2):
            try:
                return float(s2)
            except Exception:
                return s
        return s

    for r_idx, row in enumerate(rows, start=1):
        for c_idx, val in enumerate(row, start=1):
            # Treat first row as header by default: keep strings as-is
            if r_idx == 1:
                ws.cell(row=r_idx, column=c_idx, value=str(val) if val is not None else "")
            else:
                ws.cell(row=r_idx, column=c_idx, value=coerce(val))

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)



def inject_paths_into_notebook(nb_json: dict, iox_xlsx: str, cleaned_xlsx: str) -> dict:
    """Replace ioxpath/file_path lines in the notebook JSON."""
    def py_string_literal(p: str) -> str:
        # Use raw string where safe; otherwise normal escaped string
        # Raw string can't end with a single backslash, but filepaths won't.
        if "\\" in p:
            esc = p.replace("\\", "\\\\")
            return f'"{esc}"'
        return f'r"{p}"'

    iox_lit = py_string_literal(iox_xlsx)
    clean_lit = py_string_literal(cleaned_xlsx)

    for cell in nb_json.get("cells", []):
        if cell.get("cell_type") != "code":
            continue
        src = "".join(cell.get("source", []))

        if "ioxpath" in src:
            src = _replace_assignment(src, "ioxpath", iox_lit)

        if "file_path" in src:
            src = _replace_assignment(src, "file_path", clean_lit)

        cell["source"] = src.splitlines(keepends=True)

    return nb_json


def _replace_assignment(src: str, varname: str, rhs_literal: str) -> str:
    # Replace lines like: var = r""
    # Keep indentation
    out_lines = []
    for line in src.splitlines():
        if line.strip().startswith(f"{varname}"):
            # only replace direct assignment lines
            if "=" in line:
                indent = line[:len(line) - len(line.lstrip())]
                out_lines.append(f"{indent}{varname} = {rhs_literal}")
                continue
        out_lines.append(line)
    return "\n".join(out_lines) + ("\n" if src.endswith("\n") else "")


def run_notebook_in_process(nb_path: Path, workdir: Path) -> None:
    """
    Execute notebook code cells in-process (no Jupyter launch).
    Raises exceptions on failure.
    """
    with open(nb_path, "r", encoding="utf-8") as f:
        nb = json.load(f)

    env = {}
    old_cwd = os.getcwd()
    try:
        os.chdir(workdir)
        for idx, cell in enumerate(nb.get("cells", [])):
            if cell.get("cell_type") != "code":
                continue
            code = "".join(cell.get("source", []))
            if not code.strip():
                continue
            try:
                exec(code, env, env)
            except Exception as e:
                raise RuntimeError(f"Notebook execution failed in cell {idx}: {e}") from e
    finally:
        os.chdir(old_cwd)


# -----------------------------
# R / QBA helpers
# -----------------------------
def _find_rscript() -> str:
    exe = shutil.which("Rscript") or shutil.which("Rscript.exe")
    if exe:
        return exe

    candidates = [
        "/opt/homebrew/bin/Rscript",
        "/usr/local/bin/Rscript",
        "/Library/Frameworks/R.framework/Resources/bin/Rscript",
    ]
    for c in candidates:
        if Path(c).exists():
            return c

    raise FileNotFoundError("Rscript was not found. Please install R and ensure Rscript is on your PATH.")


def _load_qba_rmd_text() -> str:
    """Load the QBA Rmd text (prefers .Rmd, falls back to the provided .docx)."""
    p = resource_path(QBA_RMD_FILENAME)
    if p.exists():
        return p.read_text(encoding="utf-8", errors="replace")

    p2 = resource_path(QBA_RMD_DOCX_FILENAME)
    if p2.exists():
        try:
            from docx import Document  # pip install python-docx
        except Exception as e:
            raise RuntimeError(
                f"Found {QBA_RMD_DOCX_FILENAME} but python-docx is not installed. "
                f"Either install python-docx or bundle {QBA_RMD_FILENAME} instead."
            ) from e

        doc = Document(str(p2))
        return "\n".join([para.text for para in doc.paragraphs])

    raise FileNotFoundError(
        f"Missing {QBA_RMD_FILENAME} (or {QBA_RMD_DOCX_FILENAME}). Place it next to app.py (or bundle it)."
    )




def _load_cosinor_rmd_text(template_key: str) -> str:
    """Load a COSINOR Rmd template by key."""
    key_map = {
        "base": COSINOR_RMD_FILENAME,
        "no_outlier": COSINOR_RMD_NO_OUTLIER_FILENAME,
        "clock": COSINOR_RMD_CLOCK_FILENAME,
        "clock_no_outlier": COSINOR_RMD_CLOCK_NO_OUTLIER_FILENAME,
    }
    if template_key not in key_map:
        raise ValueError(f"Unknown COSINOR template key: {template_key}")

    p = resource_path(key_map[template_key])
    if not p.exists():
        raise FileNotFoundError(
            f"Missing COSINOR template: {p.name}. Place it next to app.py (or bundle it)."
        )
    return p.read_text(encoding="utf-8", errors="replace")
def _extract_r_chunks(rmd_text: str) -> str:
    pattern = re.compile(r"```{r[^}]*}\s*(.*?)```", re.IGNORECASE | re.DOTALL)
    chunks = pattern.findall(rmd_text)
    if not chunks:
        raise RuntimeError("Could not find any ```{r} code chunks in the QBA RMD.")
    return "\n\n".join(chunks)


def _r_str_path(p: str) -> str:
    return p.replace("\\", "/")


def _build_qba_script(rmd_text: str, input_xlsx: str, file_prefix: str) -> str:
    script = _extract_r_chunks(rmd_text)

    # --- Robust numeric coercion for Te/Ti columns (prevents 'non-numeric argument' errors)
    # If the RMD contains: Ttotal = `Te (msec)` + `Ti (msec)`
    # we rewrite it to coerce those columns to numeric safely.
    script = re.sub(
        r"`Te\s*\(msec\)`\s*\+\s*`Ti\s*\(msec\)`",
        "as.numeric(gsub('[^0-9.-]', '', as.character(`Te (msec)`))) + as.numeric(gsub('[^0-9.-]', '', as.character(`Ti (msec)`)))",
        script
    )
    script = re.sub(
        r"`Ti\s*\(msec\)`\s*\+\s*`Te\s*\(msec\)`",
        "as.numeric(gsub('[^0-9.-]', '', as.character(`Ti (msec)`))) + as.numeric(gsub('[^0-9.-]', '', as.character(`Te (msec)`)))",
        script
    )

    input_xlsx_r = _r_str_path(input_xlsx)
    prefix_r = file_prefix.replace('"', '\"')

    header = (
        'options(repos=c(CRAN="https://cloud.r-project.org"))\n'
        "options(warn=1)\n"
        f'file_path <- "{input_xlsx_r}"\n'
        f'rat_light_type <- "{prefix_r}"\n'
    )

    script = re.sub(r"file_path\s*<-\s*file\.choose\(\)", f'file_path <- "{input_xlsx_r}"', script)

    script = re.sub(
        r"rat_light_type\s*<-\s*readline\([^\)]*Enter\s+Rat\s+Light\s+Type[^\)]*\)",
        f'rat_light_type <- "{prefix_r}"',
        script,
        flags=re.IGNORECASE,
    )

    script = re.sub(
        r"temp_file\s*<-\s*file\.path\(tempdir\(\),\s*\"QBfilter_temp\.xlsx\"\).*?readline\(\s*prompt\s*=\s*\"Press\s+Enter\s+to\s+continue\.\.\.\"\s*\)\s*",
        "",
        script,
        flags=re.IGNORECASE | re.DOTALL,
    )

    return header + "\n\n" + script + "\n"





def _sanitize_r_identifier_fragment(s: str) -> str:
    """Make a string safe to embed inside R object names."""
    s = (s or "").strip()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^A-Za-z0-9_]", "", s)
    s = re.sub(r"_+", "_", s).strip("_")
    return s or "X"


def _build_cosinor_script(
    rmd_text: str,
    input_path: str,
    measured_compound: str,
    abbrev_compound_name: str,
    abbrev_tag: str,
    sheet_name: str,
    cell_range: str,
) -> str:
    """Build an R script from a COSINOR Rmd template by extracting code chunks and substituting placeholders."""
    script = _extract_r_chunks(rmd_text)

    measured = (measured_compound or "").strip()
    measured_r = measured.replace("\\", "\\\\").replace('"', '\\"')

    # sanitize fragments that are embedded into object names
    abbrev_name = _sanitize_r_identifier_fragment(abbrev_compound_name)
    tag = _sanitize_r_identifier_fragment(abbrev_tag)

    # Replace placeholders wrapped in single quotes in the template
    script = script.replace("'Your Measured Compound'", measured)
    script = script.replace("'Your Abbrev Compound Name'", abbrev_name)
    script = script.replace("'Your Abbrev Tag'", tag)

    # Replace data file placeholder (template contains ../Your Data File.xlsx)
    in_r = _r_str_path(input_path).replace('"', '\\"')
    script = script.replace("../Your Data File.xlsx", in_r)
    script = script.replace("Your Data File.xlsx", in_r)

    # Replace sheet/range placeholders from user inputs
    sheet_in = (sheet_name or "").strip()
    range_in = (cell_range or "").strip()

    if sheet_in.isdigit():
        sheet_repl = sheet_in
    else:
        # Use double quotes for safety (handles apostrophes in names)
        sheet_repl = '"' + sheet_in.replace('\\', '\\\\').replace('"', '\\"') + '"'

    range_esc = range_in.replace('\\', '\\\\').replace('"', '\\"')

    # Simple placeholder replacement (keeps behavior consistent with other tags)
    script = script.replace("Your Range", range_esc)

    script = re.sub(r"sheet\s*=\s*'Your Sheet Name'", f"sheet = {sheet_repl}", script, flags=re.IGNORECASE)
    script = re.sub(r'sheet\s*=\s*"Your Sheet Name"', f"sheet = {sheet_repl}", script, flags=re.IGNORECASE)
    script = re.sub(r'range\s*=\s*"(?:Your Range|our Range)"', f'range = "{range_esc}"', script, flags=re.IGNORECASE)
    script = re.sub(r"range\s*=\s*'(?:Your Range|our Range)'", f'range = "{range_esc}"', script, flags=re.IGNORECASE)

    header = (
        'options(repos=c(CRAN="https://cloud.r-project.org"))\n'
        "options(warn=1)\n"
        f'file_path <- "{in_r}"\n'
        f'measured_compound <- "{measured_r}"\n'
        f'abbrev_compound_name <- "{abbrev_name}"\n'
        f'abbrev_tag <- "{tag}"\n'
    )

    return header + "\n" + script + "\n"

def _run_r_script(script_path: Path, workdir: Path) -> subprocess.CompletedProcess:
    rscript = _find_rscript()
    return subprocess.run(
        [rscript, str(script_path)],
        cwd=str(workdir),
        capture_output=True,
        text=True,
        check=True,
    )


# -----------------------------
# Custom button (kept from your optical-centering version)
# -----------------------------
class AppButton(tk.Frame):
    def __init__(self, parent, text: str, font: tkfont.Font, command=None, padx=24, pady=14, cursor="hand2"):
        super().__init__(parent, bg=BTN_BG, highlightthickness=3, highlightbackground=BTN_BORDER)
        self._text = text
        self._font = font
        self._command = command
        self._padx = padx
        self._pady = pady
        self._pressed = False

        self.label = tk.Label(self, text=self._text, font=self._font, bg=BTN_BG, fg=BTN_FG)

        for w in (self, self.label):
            w.bind("<Enter>", self._on_enter)
            w.bind("<Leave>", self._on_leave)
            w.bind("<ButtonPress-1>", self._on_press)
            w.bind("<ButtonRelease-1>", self._on_release)
            w.configure(cursor=cursor if command else "arrow")

        self.label.place(relx=0.5, rely=0.5, anchor="center")
        self.refresh()

    def _has_descenders(self, s: str) -> bool:
        return any(ch in "gjpqy" for ch in s.lower())

    def _optical_y_offset(self) -> int:
        descent = int(self._font.metrics("descent"))
        if descent <= 0:
            return 0
        factor = 0.50 if not self._has_descenders(self._text) else 0.25
        return int(round(descent * factor))

    def refresh(self):
        ascent = int(self._font.metrics("ascent"))
        descent = int(self._font.metrics("descent"))
        text_w = int(self._font.measure(self._text))
        text_h = ascent + descent

        w = text_w + 2 * self._padx
        h = text_h + 2 * self._pady
        self.configure(width=w, height=h)
        self.pack_propagate(False)
        # Hard-center the label each refresh (prevents text drifting when font size changes)
        self.label.place_configure(relx=0.5, rely=0.5, anchor="center", x=0, y=0)

    def set_text(self, text: str):
        self._text = text
        self.label.configure(text=text)
        self.refresh()

    def _set_bg(self, bg):
        self.configure(bg=bg)
        self.label.configure(bg=bg)

    def _on_enter(self, _e=None):
        if self._command and not self._pressed:
            self._set_bg(BTN_ACTIVE_BG)

    def _on_leave(self, _e=None):
        self._pressed = False
        self._set_bg(BTN_BG)

    def _on_press(self, _e=None):
        if self._command:
            self._pressed = True
            self._set_bg(BTN_ACTIVE_BG)

    def _on_release(self, _e=None):
        if not self._command:
            return
        was_pressed = self._pressed
        self._pressed = False
        self._set_bg(BTN_ACTIVE_BG)

        x, y = self.winfo_pointerx(), self.winfo_pointery()
        if was_pressed and self.winfo_containing(x, y) in (self, self.label):
            self._command()


# -----------------------------
# Dropdown (filled bar + arrow box)
# -----------------------------
class FilledDropdown(tk.Frame):
    def __init__(self, parent, app, options, initial_display="", initial_value=None,
                 fill_color=BTN_BG, on_select=None, width_px=520, height_px=52, placeholder=""):
        super().__init__(parent, bg=app.bg_color)
        self.app = app
        self.options = options[:]
        self.on_select = on_select
        self.value = initial_value
        self.placeholder = placeholder

        self.width_px = width_px
        self.height_px = height_px
        self.arrow_w = 44

        self.bar = tk.Frame(
            self,
            bg=fill_color,
            width=self.width_px,
            height=self.height_px,
            highlightthickness=3,
            highlightbackground=DROPDOWN_BORDER
        )
        self.bar.pack()
        self.bar.pack_propagate(False)

        self.text_region = tk.Frame(self.bar, bg=self.bar["bg"])
        self.text_region.place(x=0, y=0, width=self.width_px - self.arrow_w, height=self.height_px)

        self.value_label = tk.Label(
            self.text_region,
            text="",
            bg=self.bar["bg"],
            fg=BTN_FG,
            font=self.app.btn_font
        )
        self.value_label.place(relx=0.03, rely=0.5, anchor="w")

        self.arrow_region = tk.Frame(self.bar, bg=DROPDOWN_ARROW_BG)
        self.arrow_region.place(x=self.width_px - self.arrow_w, y=0, width=self.arrow_w, height=self.height_px)

        self.arrow_label = tk.Label(self.arrow_region, text="▼", bg=DROPDOWN_ARROW_BG, fg=DROPDOWN_ARROW_FG,
                                    font=self.app.btn_font)
        self.arrow_label.place(relx=0.5, rely=0.5, anchor="center")

        for w in (self.bar, self.text_region, self.value_label, self.arrow_region, self.arrow_label):
            w.bind("<Button-1>", self._open_menu)

        self.menu = tk.Menu(self, tearoff=False)
        self._rebuild_menu()

        self.set_fill(fill_color)
        self.set_display(initial_display)

        self.app.register_bg_widget(self)

    def _rebuild_menu(self):
        self.menu.delete(0, "end")
        for disp, val in self.options:
            self.menu.add_command(label=disp, command=lambda d=disp, v=val: self._select(d, v))

    def set_fill(self, color_hex):
        self.bar.configure(bg=color_hex)
        self.text_region.configure(bg=color_hex)
        self.value_label.configure(bg=color_hex)

    def set_display(self, text):
        if text is None:
            text = ""
        if text == "" and self.placeholder:
            self.value_label.configure(text=self.placeholder, fg=PLACEHOLDER_FG)
        else:
            self.value_label.configure(text=text, fg=BTN_FG)

    def _open_menu(self, _e=None):
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.height_px + 6
        self.menu.tk_popup(x, y)

    def _select(self, disp, val):
        self.value = val
        self.set_display(disp)
        if self.on_select:
            self.on_select(val)


# -----------------------------
# App + UI structure
# -----------------------------
class App(BaseTk):
    def __init__(self):
        super().__init__()

        self.bg_color = DEFAULT_BG
        self.lang = "en"  # defaults to English each launch

        self.base_font_size = 16
        self.button_font_size = self.base_font_size
        self.btn_font = tkfont.Font(family="Arial", size=self.button_font_size, weight="bold")
        self.small_btn_font = tkfont.Font(family="Arial", size=max(10, self.button_font_size - 3), weight="bold")

        self._button_entries = []
        self._bg_widgets = []
        self._i18n_buttons = []

        self.title(self.t("title"))
        self.geometry("900x600")
        self.minsize(700, 450)
        self.configure(bg=self.bg_color)

        self.container = tk.Frame(self, bg=self.bg_color)
        self.container.pack(fill="both", expand=True)
        self.register_bg_widget(self.container)

        self.frames = {}
        for FrameClass, name in [
            (MainMenuFrame, "main"),
            (SettingsFrame, "settings"),
            (RunAnalysisSelectFrame, "run_select"),
            (ConvertTxtFrame, "convert_txt"),
            (QBAnalysisFrame, "qba"),
            (CosinorModelingFrame, "cosinor"),
        ]:
            frame = FrameClass(parent=self.container, app=self)
            self.frames[name] = frame
            frame.place(relx=0, rely=0, relwidth=1, relheight=1)
            self.register_bg_widget(frame)

        self.after_idle(self._capture_button_bases)
        self.show_frame("main")

    def t(self, key: str) -> str:
        return I18N.get(self.lang, I18N["en"]).get(key, key)

    def register_i18n_button(self, btn: AppButton, key: str):
        self._i18n_buttons.append((btn, key))

    def refresh_language(self):
        self.title(self.t("title"))
        for btn, key in self._i18n_buttons:
            btn.set_text(self.t(key))
        self.after_idle(self._refresh_button_wrapper_sizes)

    def register_bg_widget(self, widget):
        self._bg_widgets.append(widget)

    def apply_background(self, new_bg: str):
        self.bg_color = new_bg
        self.configure(bg=new_bg)
        for w in self._bg_widgets:
            try:
                w.configure(bg=new_bg)
            except tk.TclError:
                pass

    def show_frame(self, name: str):
        frame = self.frames[name]
        if hasattr(frame, "on_show"):
            frame.on_show()
        frame.tkraise()

    # wrapper autosize
    def register_button_wrapper(self, wrapper: tk.Frame, button: AppButton):
        self._button_entries.append({"wrapper": wrapper, "button": button, "base_w": None, "base_h": None})
        self.register_bg_widget(wrapper)

    def _capture_button_bases(self):
        self.update_idletasks()
        for entry in self._button_entries:
            entry["button"].refresh()
            w = entry["button"].winfo_reqwidth()
            h = entry["button"].winfo_reqheight()
            entry["base_w"] = w
            entry["base_h"] = h
            entry["wrapper"].configure(width=w, height=h)
            entry["wrapper"].pack_propagate(False)

    def _refresh_button_wrapper_sizes(self):
        self.update_idletasks()
        font_changed = (self.button_font_size != self.base_font_size)
        for entry in self._button_entries:
            base_w, base_h = entry["base_w"], entry["base_h"]
            if base_w is None or base_h is None:
                continue
            entry["button"].refresh()
            req_w = entry["button"].winfo_reqwidth()
            req_h = entry["button"].winfo_reqheight()
            if font_changed and (req_w > base_w or req_h > base_h):
                entry["wrapper"].configure(width=max(base_w, req_w), height=max(base_h, req_h))
            else:
                entry["wrapper"].configure(width=base_w, height=base_h)

    def apply_font_size(self, new_size: int):
        self.button_font_size = new_size
        self.btn_font.configure(size=self.button_font_size)
        self.small_btn_font.configure(size=max(10, self.button_font_size - 3))
        self.after_idle(self._refresh_button_wrapper_sizes)


class BaseCenteredMenu(tk.Frame):
    def __init__(self, parent, app: App):
        super().__init__(parent, bg=app.bg_color)
        self.app = app
        self.center = tk.Frame(self, bg=app.bg_color)
        self.center.pack(expand=True)
        self.app.register_bg_widget(self.center)

    def add_button(self, key, command, small=False, pady=10):
        wrapper = tk.Frame(self.center, bg=self.app.bg_color)
        wrapper.pack(pady=pady)

        font = self.app.small_btn_font if small else self.app.btn_font
        padx = 18 if small else 24
        pady_in = 10 if small else 14

        btn = AppButton(wrapper, text=self.app.t(key), font=font, command=command, padx=padx, pady=pady_in)
        btn.place(relx=0.5, rely=0.5, anchor="center")

        self.app.register_button_wrapper(wrapper, btn)
        self.app.register_i18n_button(btn, key)
        return btn


class MainMenuFrame(BaseCenteredMenu):
    def __init__(self, parent, app: App):
        super().__init__(parent, app)
        self.add_button("run_analysis", lambda: self.app.show_frame("run_select"))
        self.add_button("settings", lambda: self.app.show_frame("settings"))
        self.add_button("exit", self.app.destroy)


class SettingsFrame(BaseCenteredMenu):
    def __init__(self, parent, app: App):
        super().__init__(parent, app)
        # Keeping this minimal here; you can merge back your full settings screens
        self.add_button("back", lambda: self.app.show_frame("main"), small=True, pady=(0, 0))


class RunAnalysisSelectFrame(tk.Frame):
    def __init__(self, parent, app: App):
        super().__init__(parent, bg=app.bg_color)
        self.app = app

        self.center = tk.Frame(self, bg=app.bg_color)
        self.center.pack(expand=True)
        self.app.register_bg_widget(self.center)

        self.protocols = [
            (self.app.t("protocol_convert"), "convert"),
            (self.app.t("protocol_qb"), "qb"),
            (self.app.t("protocol_cosinor"), "cosinor"),
        ]

        self.dropdown = FilledDropdown(
            self.center,
            app=self.app,
            options=self.protocols,
            initial_display="",
            initial_value=None,
            fill_color=BTN_BG,
            placeholder=self.app.t("select_protocol"),
            on_select=self.on_pick
        )
        self.dropdown.pack(pady=(0, 50))

        row = tk.Frame(self.center, bg=self.app.bg_color)
        row.pack()
        self.app.register_bg_widget(row)

        # Back to main (no changes to keep)
        self.back_wrapper = tk.Frame(row, bg=self.app.bg_color)
        self.back_wrapper.pack(side="left", padx=10)
        self.app.register_bg_widget(self.back_wrapper)

        self.back_btn = AppButton(self.back_wrapper, text=self.app.t("back"),
                                  font=self.app.small_btn_font, command=lambda: self.app.show_frame("main"),
                                  padx=18, pady=10)
        self.back_btn.place(relx=0.5, rely=0.5, anchor="center")
        self.app.register_button_wrapper(self.back_wrapper, self.back_btn)

        # Accept -> route to protocol screen
        self.accept_wrapper = tk.Frame(row, bg=self.app.bg_color)
        self.accept_wrapper.pack(side="left", padx=10)
        self.app.register_bg_widget(self.accept_wrapper)

        self.accept_btn = AppButton(self.accept_wrapper, text=self.app.t("accept"),
                                    font=self.app.small_btn_font, command=self.accept_protocol,
                                    padx=18, pady=10)
        self.accept_btn.place(relx=0.5, rely=0.5, anchor="center")
        self.app.register_button_wrapper(self.accept_wrapper, self.accept_btn)

        self.pending_protocol = None

    def on_show(self):
        # reset blank/unfilled
        self.pending_protocol = None
        self.dropdown.value = None
        self.dropdown.set_display("")
        self.dropdown.placeholder = self.app.t("select_protocol")
        self.dropdown.set_display("")

    def on_pick(self, protocol_key: str):
        self.pending_protocol = protocol_key

    def accept_protocol(self):
        if not self.pending_protocol:
            messagebox.showinfo("Select", "Please select a protocol first.")
            return
        if self.pending_protocol == "convert":
            self.app.show_frame("convert_txt")
        elif self.pending_protocol == "qb":
            self.app.show_frame("qba")
        elif self.pending_protocol == "cosinor":
            self.app.show_frame("cosinor")
        else:
            messagebox.showinfo("Coming soon", "That protocol screen is next.")
            # remain here


class ConvertTxtFrame(tk.Frame):
    def __init__(self, parent, app: App):
        super().__init__(parent, bg=app.bg_color)
        self.app = app

        self.center = tk.Frame(self, bg=app.bg_color)
        self.center.pack(expand=True, fill="both")
        self.app.register_bg_widget(self.center)

        # Upload box
        self.upload_box = tk.Frame(
            self.center,
            bg=BTN_BG,
            highlightthickness=3,
            highlightbackground=DROPDOWN_BORDER,
            width=700,
            height=180
        )
        self.upload_box.pack(pady=(0, 20))
        self.upload_box.pack_propagate(False)

        self.files_var = tk.StringVar(value=self.app.t("no_files"))
        self.files_label = tk.Label(self.upload_box, textvariable=self.files_var, bg=BTN_BG, fg=BTN_FG,
                                    font=self.app.btn_font, justify="center", wraplength=650)
        self.files_label.place(relx=0.5, rely=0.5, anchor="center")

        # Optional drag/drop
        if DND_FILES is not None and hasattr(self.upload_box, "drop_target_register"):
            self.upload_box.drop_target_register(DND_FILES)
            self.upload_box.dnd_bind("<<Drop>>", self._on_drop)

        # Controls row
        controls = tk.Frame(self.center, bg=app.bg_color)
        controls.pack(pady=(0, 25))
        self.app.register_bg_widget(controls)

        # Upload button
        self.upload_wrapper = tk.Frame(controls, bg=app.bg_color)
        self.upload_wrapper.pack(side="left", padx=10)
        self.app.register_bg_widget(self.upload_wrapper)

        self.upload_btn = AppButton(self.upload_wrapper, text=self.app.t("upload_files"),
                                    font=self.app.small_btn_font, command=self.browse_files,
                                    padx=18, pady=10)
        self.upload_btn.place(relx=0.5, rely=0.5, anchor="center")
        self.app.register_button_wrapper(self.upload_wrapper, self.upload_btn)

        # Clear button
        self.clear_wrapper = tk.Frame(controls, bg=app.bg_color)
        self.clear_wrapper.pack(side="left", padx=10)
        self.app.register_bg_widget(self.clear_wrapper)

        self.clear_btn = AppButton(self.clear_wrapper, text=self.app.t("clear"),
                                   font=self.app.small_btn_font, command=self.clear_files,
                                   padx=18, pady=10)
        self.clear_btn.place(relx=0.5, rely=0.5, anchor="center")
        self.app.register_button_wrapper(self.clear_wrapper, self.clear_btn)

        # Run Files button
        self.run_wrapper = tk.Frame(self.center, bg=app.bg_color)
        self.run_wrapper.pack(pady=(0, 30))
        self.app.register_bg_widget(self.run_wrapper)

        self.run_btn = AppButton(self.run_wrapper, text=self.app.t("run_files"),
                                 font=self.app.btn_font, command=self.run_files,
                                 padx=24, pady=14)
        self.run_btn.place(relx=0.5, rely=0.5, anchor="center")
        self.app.register_button_wrapper(self.run_wrapper, self.run_btn)

        # Back/Accept row (Back returns to Run Select; Accept not needed here, so Back only)
        row = tk.Frame(self.center, bg=app.bg_color)
        row.pack()
        self.app.register_bg_widget(row)

        self.back_wrapper = tk.Frame(row, bg=app.bg_color)
        self.back_wrapper.pack(side="left", padx=10)
        self.app.register_bg_widget(self.back_wrapper)

        self.back_btn = AppButton(self.back_wrapper, text=self.app.t("back"),
                                  font=self.app.small_btn_font,
                                  command=lambda: self.app.show_frame("run_select"),
                                  padx=18, pady=10)
        self.back_btn.place(relx=0.5, rely=0.5, anchor="center")
        self.app.register_button_wrapper(self.back_wrapper, self.back_btn)

        # Status label
        self.status_var = tk.StringVar(value="")
        self.status = tk.Label(self.center, textvariable=self.status_var,
                               bg=app.bg_color, fg="white", font=("Arial", 12, "bold"))
        self.status.pack(pady=(15, 0))

        # Convert to QBA button + error label
        self.convert_error_var = tk.StringVar(value="")
        self.convert_error = tk.Label(
            self.center,
            textvariable=self.convert_error_var,
            bg=app.bg_color,
            fg="#FFB4B4",
            font=("Arial", 12, "bold")
        )
        self.convert_error.pack(pady=(10, 0))
        self.app.register_bg_widget(self.convert_error)

        self.convert_qba_wrapper = tk.Frame(self.center, bg=app.bg_color)
        self.convert_qba_wrapper.pack(pady=(12, 20))
        self.app.register_bg_widget(self.convert_qba_wrapper)

        self.convert_qba_btn = AppButton(
            self.convert_qba_wrapper,
            text=self.app.t("convert_to_qba"),
            font=self.app.small_btn_font,
            command=self.convert_to_qba,
            padx=18,
            pady=10
        )
        self.convert_qba_btn.place(relx=0.5, rely=0.5, anchor="center")
        self.app.register_button_wrapper(self.convert_qba_wrapper, self.convert_qba_btn)
        self.app.register_bg_widget(self.status)

        self.selected_files = []

        self._qba_ready_cleaned_files = []
        self._qba_ready = False

    def on_show(self):
        self.clear_files()
        self.status_var.set("")
        self.convert_error_var.set("")
        self._qba_ready_cleaned_files = []
        self._qba_ready = False

    def _on_drop(self, event):
        # event.data can contain multiple paths
        raw = event.data
        paths = self._split_dnd_files(raw)
        self._add_files(paths)

    @staticmethod
    def _split_dnd_files(data: str):
        # Handles formats like: "{/path/one} {/path/two}" or plain paths
        parts = []
        buf = ""
        in_brace = False
        for ch in data:
            if ch == "{":
                in_brace = True
                buf = ""
            elif ch == "}":
                in_brace = False
                parts.append(buf)
                buf = ""
            elif ch == " " and not in_brace:
                if buf.strip():
                    parts.append(buf.strip())
                    buf = ""
            else:
                buf += ch
        if buf.strip():
            parts.append(buf.strip())
        return parts

    def browse_files(self):
        paths = filedialog.askopenfilenames(
            title="Select .iox / .iox.txt files",
            filetypes=[
                ("IOX files", "*.iox"),
                ("IOX TXT files", "*.iox.txt"),
                ("Text files", "*.txt"),
                ("All files", "*.*"),
            ]
        )
        if paths:
            self._add_files(paths)

    def _add_files(self, paths):
        # Filter to .iox or .iox.txt (or .txt that ends with .iox.txt)
        valid = []
        for p in paths:
            lp = p.lower()
            if lp.endswith(".iox") or lp.endswith(".iox.txt"):
                valid.append(p)
        if not valid:
            messagebox.showinfo("Files", "Please select .iox or .iox.txt files.")
            return

        # De-dupe
        existing = set(self.selected_files)
        for p in valid:
            if p not in existing:
                self.selected_files.append(p)

        self.files_var.set("\n".join([Path(p).name for p in self.selected_files]))

    def clear_files(self):
        self.selected_files = []
        self.files_var.set(self.app.t("no_files"))
        self.convert_error_var.set("")
        self._qba_ready_cleaned_files = []
        self._qba_ready = False

    def run_files(self):
        if not self.selected_files:
            messagebox.showinfo("Run", "No files selected.")
            return

        # Run in background thread so UI doesn't freeze
        thread = threading.Thread(target=self._run_files_worker, daemon=True)
        thread.start()

    def _run_files_worker(self):
        try:
            nb_src = resource_path(NOTEBOOK_FILENAME)
            if not nb_src.exists():
                self._set_status(f"Missing notebook: {nb_src}")
                return

            cleaned_outputs = []

            for i, fpath in enumerate(self.selected_files, start=1):
                self._set_status(f"Processing {i}/{len(self.selected_files)}: {Path(fpath).name}")

                src_path = Path(fpath)
                out_dir = src_path.parent / OUTPUT_FOLDER_NAME
                out_dir.mkdir(parents=True, exist_ok=True)

                base = strip_iox_extensions(src_path.name)
                raw_xlsx = out_dir / f"{base}.xlsx"
                cleaned_xlsx = out_dir / f"{base}_cleaned.xlsx"

                # 1) Read IOX text
                text = src_path.read_text(encoding="utf-8", errors="replace")

                # 2) Convert to xlsx (internal, no clipboard/excel needed)
                rows = parse_text_to_rows(text)
                if not rows:
                    raise RuntimeError(f"File appears empty: {src_path.name}")
                write_rows_to_xlsx(rows, raw_xlsx)

                # 3) Inject paths into notebook and execute it internally
                with open(nb_src, "r", encoding="utf-8") as f:
                    nb_json = json.load(f)

                nb_json = inject_paths_into_notebook(nb_json, str(raw_xlsx), str(cleaned_xlsx))

                tmp_nb = out_dir / f"_tmp_run_{base}.ipynb"
                with open(tmp_nb, "w", encoding="utf-8") as f:
                    json.dump(nb_json, f, ensure_ascii=False, indent=2)

                run_notebook_in_process(tmp_nb, workdir=out_dir)

                # 4) Verify cleaned file exists
                if not cleaned_xlsx.exists() or cleaned_xlsx.stat().st_size == 0:
                    raise RuntimeError(f"Cleaned file not created: {cleaned_xlsx}")

                cleaned_outputs.append(str(cleaned_xlsx))

                # cleanup temp notebook
                try:
                    tmp_nb.unlink()
                except Exception:
                    pass

            self._qba_ready_cleaned_files = cleaned_outputs
            self._qba_ready = True

            self._set_status("Done. Files saved in: " + str((Path(self.selected_files[0]).parent / OUTPUT_FOLDER_NAME)))
            self._show_info("Success", "All selected file(s) converted and cleaned successfully.")
        except Exception as e:
            self._qba_ready_cleaned_files = []
            self._qba_ready = False
            self._set_status("Error: " + str(e))
            self._show_info("Error", str(e))
    def convert_to_qba(self):
        # Only works if Convert Txt just ran successfully in this page instance
        if not self._qba_ready or not self._qba_ready_cleaned_files:
            self.convert_error_var.set(self.app.t("err_no_eligible"))
            return

        self.convert_error_var.set("")
        qba_frame = self.app.frames.get("qba")
        if qba_frame:
            qba_frame.prefill_files(self._qba_ready_cleaned_files)
        self.app.show_frame("qba")


    def _set_status(self, msg: str):
        self.after(0, lambda: self.status_var.set(msg))

    def _show_info(self, title: str, msg: str):
        self.after(0, lambda: messagebox.showinfo(title, msg))



class QBAnalysisFrame(tk.Frame):
    def __init__(self, parent, app: App):
        super().__init__(parent, bg=app.bg_color)
        self.app = app

        self.center = tk.Frame(self, bg=app.bg_color)
        self.center.pack(expand=True, fill="both")
        self.app.register_bg_widget(self.center)

        self.upload_box = tk.Frame(
            self.center,
            bg=BTN_BG,
            highlightthickness=3,
            highlightbackground=DROPDOWN_BORDER,
            width=700,
            height=180
        )
        self.upload_box.pack(pady=(0, 20))
        self.upload_box.pack_propagate(False)

        self.files_var = tk.StringVar(value=self.app.t("no_files"))
        self.files_label = tk.Label(
            self.upload_box,
            textvariable=self.files_var,
            bg=BTN_BG,
            fg=BTN_FG,
            font=self.app.btn_font,
            justify="center",
            wraplength=650
        )
        self.files_label.place(relx=0.5, rely=0.5, anchor="center")

        if DND_FILES is not None and hasattr(self.upload_box, "drop_target_register"):
            self.upload_box.drop_target_register(DND_FILES)
            self.upload_box.dnd_bind("<<Drop>>", self._on_drop)

        controls = tk.Frame(self.center, bg=app.bg_color)
        controls.pack(pady=(0, 20))
        self.app.register_bg_widget(controls)

        self.upload_wrapper = tk.Frame(controls, bg=app.bg_color)
        self.upload_wrapper.pack(side="left", padx=10)
        self.app.register_bg_widget(self.upload_wrapper)

        self.upload_btn = AppButton(
            self.upload_wrapper,
            text=self.app.t("upload_files"),
            font=self.app.small_btn_font,
            command=self.browse_files,
            padx=18,
            pady=10
        )
        self.upload_btn.place(relx=0.5, rely=0.5, anchor="center")
        self.app.register_button_wrapper(self.upload_wrapper, self.upload_btn)

        self.clear_wrapper = tk.Frame(controls, bg=app.bg_color)
        self.clear_wrapper.pack(side="left", padx=10)
        self.app.register_bg_widget(self.clear_wrapper)

        self.clear_btn = AppButton(
            self.clear_wrapper,
            text=self.app.t("clear"),
            font=self.app.small_btn_font,
            command=self.clear_files,
            padx=18,
            pady=10
        )
        self.clear_btn.place(relx=0.5, rely=0.5, anchor="center")
        self.app.register_button_wrapper(self.clear_wrapper, self.clear_btn)

        prefix_block = tk.Frame(self.center, bg=app.bg_color)
        prefix_block.pack(pady=(0, 10))
        self.app.register_bg_widget(prefix_block)

        self.prefix_label = tk.Label(
            prefix_block,
            text=self.app.t("file_prefix"),
            bg=app.bg_color,
            fg="white",
            font=("Arial", 14, "bold")
        )
        self.prefix_label.pack(pady=(0, 8))
        self.app.register_bg_widget(self.prefix_label)

        self.prefix_entry = tk.Entry(prefix_block, font=self.app.btn_font, width=28, justify="center")
        self.prefix_entry.pack()

        self.run_wrapper = tk.Frame(self.center, bg=app.bg_color)
        self.run_wrapper.pack(pady=(16, 10))
        self.app.register_bg_widget(self.run_wrapper)

        self.run_btn = AppButton(
            self.run_wrapper,
            text=self.app.t("run_files"),
            font=self.app.btn_font,
            command=self.run_files,
            padx=24,
            pady=14
        )
        self.run_btn.place(relx=0.5, rely=0.5, anchor="center")
        self.app.register_button_wrapper(self.run_wrapper, self.run_btn)

        self.err_var = tk.StringVar(value="")
        self.err_label = tk.Label(self.center, textvariable=self.err_var, bg=app.bg_color, fg="#FFB4B4",
                                  font=("Arial", 12, "bold"))
        self.err_label.pack(pady=(6, 0))
        self.app.register_bg_widget(self.err_label)

        self.status_var = tk.StringVar(value="")
        self.status = tk.Label(self.center, textvariable=self.status_var, bg=app.bg_color, fg="white",
                               font=("Arial", 12, "bold"))
        self.status.pack(pady=(10, 0))
        self.app.register_bg_widget(self.status)

        row = tk.Frame(self.center, bg=app.bg_color)
        row.pack(pady=(16, 0))
        self.app.register_bg_widget(row)

        self.back_wrapper = tk.Frame(row, bg=app.bg_color)
        self.back_wrapper.pack(side="left", padx=10)
        self.app.register_bg_widget(self.back_wrapper)

        self.back_btn = AppButton(
            self.back_wrapper,
            text=self.app.t("back"),
            font=self.app.small_btn_font,
            command=lambda: self.app.show_frame("run_select"),
            padx=18,
            pady=10
        )
        self.back_btn.place(relx=0.5, rely=0.5, anchor="center")
        self.app.register_button_wrapper(self.back_wrapper, self.back_btn)

        self.selected_files = []

    def on_show(self):
        self.err_var.set("")
        self.status_var.set("")

    def prefill_files(self, paths):
        self.selected_files = list(paths)
        self.files_var.set("\n".join([Path(p).name for p in self.selected_files]))

    def _on_drop(self, event):
        raw = event.data
        paths = ConvertTxtFrame._split_dnd_files(raw)
        self._add_files(paths)

    def browse_files(self):
        paths = filedialog.askopenfilenames(
            title="Select .xlsx file(s)",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if paths:
            self._add_files(paths)

    def _add_files(self, paths):
        valid = [p for p in paths if p.lower().endswith(".xlsx")]
        if not valid:
            messagebox.showinfo("Files", "Please select .xlsx files.")
            return

        existing = set(self.selected_files)
        for p in valid:
            if p not in existing:
                self.selected_files.append(p)

        self.files_var.set("\n".join([Path(p).name for p in self.selected_files]))
        self.err_var.set("")

    def clear_files(self):
        self.selected_files = []
        self.files_var.set(self.app.t("no_files"))
        self.err_var.set("")
        self.status_var.set("")

    def run_files(self):
        prefix = self.prefix_entry.get().strip()
        if not prefix:
            self.err_var.set(self.app.t("err_no_prefix"))
            return

        if not self.selected_files:
            messagebox.showinfo("Run", "No files selected.")
            return

        self.err_var.set("")
        threading.Thread(target=self._run_worker, args=(prefix,), daemon=True).start()

    def _run_worker(self, prefix: str):
        try:
            rmd_text = _load_qba_rmd_text()
        except Exception as e:
            self._set_status(f"Error: {e}")
            self._show_info("Error", str(e))
            return

        try:
            for i, fpath in enumerate(self.selected_files, start=1):
                src_path = Path(fpath)
                out_dir = src_path.parent / QBA_OUTPUT_FOLDER_NAME
                out_dir.mkdir(parents=True, exist_ok=True)

                self._set_status(f"Running QBA {i}/{len(self.selected_files)}: {src_path.name}")

                script_text = _build_qba_script(rmd_text, str(src_path), prefix)
                script_path = out_dir / f"_tmp_qba_run_{src_path.stem}.R"
                script_path.write_text(script_text, encoding="utf-8")

                try:
                    _run_r_script(script_path, workdir=out_dir)
                finally:
                    try:
                        script_path.unlink()
                    except Exception:
                        pass

                if not (out_dir / "Circadian QB Filter").exists():
                    raise RuntimeError("QBA finished but expected output folder was not found:\n" + str(out_dir / "Circadian QB Filter"))

            self._set_status("Done. Outputs saved under: " + str(Path(self.selected_files[0]).parent / QBA_OUTPUT_FOLDER_NAME))
            self._show_info("Success", "QBA completed successfully for all selected file(s).")
        except subprocess.CalledProcessError as e:
            msg = (e.stderr or e.stdout or str(e)).strip()
            self._set_status("Error running R: " + msg[:3000])
            self._show_info("Error (R)", msg[:3000])
        except Exception as e:
            self._set_status("Error: " + str(e))
            self._show_info("Error", str(e))

    def _set_status(self, msg: str):
        self.after(0, lambda: self.status_var.set(msg))

    def _show_info(self, title: str, msg: str):
        self.after(0, lambda: messagebox.showinfo(title, msg))




class CosinorModelingFrame(tk.Frame):
    """COSINOR Modeling screen under Run Analysis."""

    TEMPLATE_OPTIONS = [
        ("COSINOR Markdown", "base"),
        ("COSINOR Markdown No Outlier", "no_outlier"),
        ("COSINOR Markdown (Clock Genes)", "clock"),
        ("COSINOR Markdown (Clock Genes) No Outlier", "clock_no_outlier"),
    ]

    TEMPLATE_HELP = {
        None: "Choose Sample Template",
        "base": "COSINOR Analysis with Automatic Outlier filtering (larger datasets)",
        "no_outlier": "COSINOR Analysis without Automatic Outlier filtering",
        "clock": "COSINOR Clock Gene Analysis with Automatic Outlier filtering (larger datasets)",
        "clock_no_outlier": "COSINOR Clock Gene Analysis without Automatic Outlier filtering",
    }

    def __init__(self, parent, app: App):
        super().__init__(parent, bg=app.bg_color)
        self.app = app

        self.center = tk.Frame(self, bg=app.bg_color)
        self.center.pack(expand=True, fill="both")
        self.app.register_bg_widget(self.center)

        # --- Upload box (styled same as ConvertTxtFrame / QBAnalysisFrame) ---
        self.upload_box = tk.Frame(
            self.center,
            bg=BTN_BG,
            highlightthickness=3,
            highlightbackground=DROPDOWN_BORDER,
            width=700,
            height=180
        )
        self.upload_box.pack(pady=(0, 20))
        self.upload_box.pack_propagate(False)

        self.files_var = tk.StringVar(value=self.app.t("no_files"))
        self.files_label = tk.Label(
            self.upload_box,
            textvariable=self.files_var,
            bg=BTN_BG,
            fg=BTN_FG,
            font=self.app.btn_font,
            justify="center",
            wraplength=650
        )
        self.files_label.place(relx=0.5, rely=0.5, anchor="center")

        if DND_FILES is not None and hasattr(self.upload_box, "drop_target_register"):
            self.upload_box.drop_target_register(DND_FILES)
            self.upload_box.dnd_bind("<<Drop>>", self._on_drop)

        controls = tk.Frame(self.center, bg=app.bg_color)
        controls.pack(pady=(0, 25))
        self.app.register_bg_widget(controls)

        self.upload_wrapper = tk.Frame(controls, bg=app.bg_color)
        self.upload_wrapper.pack(side="left", padx=10)
        self.app.register_bg_widget(self.upload_wrapper)

        self.upload_btn = AppButton(
            self.upload_wrapper,
            text=self.app.t("upload_files"),
            font=self.app.small_btn_font,
            command=self.browse_files,
            padx=18,
            pady=10
        )
        self.upload_btn.place(relx=0.5, rely=0.5, anchor="center")
        self.app.register_button_wrapper(self.upload_wrapper, self.upload_btn)

        self.clear_wrapper = tk.Frame(controls, bg=app.bg_color)
        self.clear_wrapper.pack(side="left", padx=10)
        self.app.register_bg_widget(self.clear_wrapper)

        self.clear_btn = AppButton(
            self.clear_wrapper,
            text=self.app.t("clear"),
            font=self.app.small_btn_font,
            command=self.clear_files,
            padx=18,
            pady=10
        )
        self.clear_btn.place(relx=0.5, rely=0.5, anchor="center")
        self.app.register_button_wrapper(self.clear_wrapper, self.clear_btn)

        # --- 3-column block under upload ---
        columns = tk.Frame(self.center, bg=app.bg_color)
        columns.pack(pady=(5, 0), padx=30, fill="x")
        self.app.register_bg_widget(columns)

        columns.grid_columnconfigure(0, weight=1, uniform="col")
        columns.grid_columnconfigure(1, weight=1, uniform="col")
        columns.grid_columnconfigure(2, weight=1, uniform="col")

        # Left column: descriptor entries
        left = tk.Frame(columns, bg=app.bg_color)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 15))
        self.app.register_bg_widget(left)

        def make_labeled_entry(parent_frame, label_text):
            block = tk.Frame(parent_frame, bg=app.bg_color)
            block.pack(fill="x", pady=(0, 12))
            self.app.register_bg_widget(block)

            lbl = tk.Label(block, text=label_text, bg=app.bg_color, fg="white", font=("Arial", 12, "bold"))
            lbl.pack(anchor="w", pady=(0, 6))
            self.app.register_bg_widget(lbl)

            ent = tk.Entry(block, font=("Arial", 12), width=28)
            ent.pack(fill="x")
            return ent

        self.measured_entry = make_labeled_entry(left, "'Your Measured Compound'")
        self.abbrev_name_entry = make_labeled_entry(left, "'Your Abbrev Compound Name'")
        self.abbrev_tag_entry = make_labeled_entry(left, "'Your Abbrev Tag'")
        self.sheet_entry = make_labeled_entry(left, "'Your Sheet Name'")
        self.range_entry = make_labeled_entry(left, "'Your Range'")

        # Middle column: template dropdown
        mid = tk.Frame(columns, bg=app.bg_color)
        mid.grid(row=0, column=1, sticky="nsew", padx=15)
        self.app.register_bg_widget(mid)

        self.selected_template = None

        self.template_dropdown = FilledDropdown(
            mid,
            app=self.app,
            options=self.TEMPLATE_OPTIONS,
            initial_display="",
            initial_value=None,
            fill_color=BTN_BG,
            placeholder="",
            on_select=self._on_template_pick
        )
        self.template_dropdown.pack(pady=(0, 10))

        # Buttons under centered middle column
        btn_row = tk.Frame(mid, bg=app.bg_color)
        btn_row.pack(pady=(10, 0))
        self.app.register_bg_widget(btn_row)

        self.back_wrapper = tk.Frame(btn_row, bg=app.bg_color)
        self.back_wrapper.pack(side="left", padx=10)
        self.app.register_bg_widget(self.back_wrapper)

        self.back_btn = AppButton(
            self.back_wrapper,
            text=self.app.t("back"),
            font=self.app.small_btn_font,
            command=lambda: self.app.show_frame("run_select"),
            padx=18,
            pady=10
        )
        self.back_btn.place(relx=0.5, rely=0.5, anchor="center")
        self.app.register_button_wrapper(self.back_wrapper, self.back_btn)

        self.run_wrapper = tk.Frame(btn_row, bg=app.bg_color)
        self.run_wrapper.pack(side="left", padx=10)
        self.app.register_bg_widget(self.run_wrapper)

        self.run_btn = AppButton(
            self.run_wrapper,
            text=self.app.t("run_files"),
            font=self.app.small_btn_font,
            command=self.run_files,
            padx=18,
            pady=10
        )
        self.run_btn.place(relx=0.5, rely=0.5, anchor="center")
        self.app.register_button_wrapper(self.run_wrapper, self.run_btn)

        # Right column: non-interactable text window showing template help
        right = tk.Frame(columns, bg=app.bg_color)
        right.grid(row=0, column=2, sticky="nsew", padx=(15, 0))
        self.app.register_bg_widget(right)

        self.help_text = tk.Text(right, height=8, width=30, wrap="word", font=("Arial", 11))
        self.help_text.pack(fill="both", expand=True)
        self.help_text.configure(state="disabled")
        self._set_help_text(self.TEMPLATE_HELP[None])

        # Status / error (kept consistent with other screens)
        self.err_var = tk.StringVar(value="")
        self.err_label = tk.Label(self.center, textvariable=self.err_var, bg=app.bg_color, fg="yellow",
                                  font=("Arial", 12, "bold"))
        self.err_label.pack(pady=(12, 0))
        self.app.register_bg_widget(self.err_label)

        self.status_var = tk.StringVar(value="")
        self.status = tk.Label(self.center, textvariable=self.status_var, bg=app.bg_color, fg="white",
                               font=("Arial", 12, "bold"), wraplength=760, justify="center")
        self.status.pack(pady=(10, 0))
        self.app.register_bg_widget(self.status)

        self.selected_files = []

    def on_show(self):
        self.err_var.set("")
        self.status_var.set("")

    def _set_help_text(self, msg: str):
        self.help_text.configure(state="normal")
        self.help_text.delete("1.0", "end")
        self.help_text.insert("1.0", msg)
        self.help_text.configure(state="disabled")

    def _on_template_pick(self, template_key: str):
        self.selected_template = template_key
        msg = self.TEMPLATE_HELP.get(template_key, self.TEMPLATE_HELP[None])
        self._set_help_text(msg)

    def _set_status(self, msg: str):
        self.status_var.set(msg)

    def _show_info(self, title: str, msg: str):
        self.after(0, lambda: messagebox.showinfo(title, msg))

    def _show_error(self, title: str, msg: str):
        self.after(0, lambda: messagebox.showerror(title, msg))

    def _on_drop(self, event):
        raw = event.data
        paths = ConvertTxtFrame._split_dnd_files(raw)
        self._add_files(paths)

    def browse_files(self):
        paths = filedialog.askopenfilenames(
            title="Select .xlsx file(s)",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if paths:
            self._add_files(paths)

    def _add_files(self, paths):
        valid = [p for p in paths if p.lower().endswith(".xlsx")]
        if not valid:
            messagebox.showinfo("Files", "Please select .xlsx files.")
            return

        existing = set(self.selected_files)
        for p in valid:
            if p not in existing:
                self.selected_files.append(p)

        if self.selected_files:
            show = "\n".join([Path(p).name for p in self.selected_files])
            self.files_var.set(show)
        else:
            self.files_var.set(self.app.t("no_files"))

    def clear_files(self):
        self.selected_files = []
        self.files_var.set(self.app.t("no_files"))

    def _descriptor_error_message(
        self,
        measured: str,
        abbrev_name: str,
        abbrev_tag: str,
        sheet_name: str,
        cell_range: str,
    ) -> str:
        missing = []
        if not measured:
            missing.append("Measured Compound")
        if not abbrev_name:
            missing.append("Abbrev Compound Name")
        if not abbrev_tag:
            missing.append("Abbrev Tag")
        if not sheet_name:
            missing.append("Sheet Name")
        if not cell_range:
            missing.append("Range")

        if len(missing) == 5:
            return "Please fill in all Compound Descriptor Tags"

        if len(missing) == 1:
            return f"Please fill in {missing[0]} Bar"

        if len(missing) == 2:
            joined = f"{missing[0]} and {missing[1]}"
        else:
            joined = ", ".join(missing[:-1]) + f", and {missing[-1]}"

        return f"Please fill in {joined} Bars"


    def run_files(self):
        measured = self.measured_entry.get().strip()
        abbrev_name = self.abbrev_name_entry.get().strip()
        abbrev_tag = self.abbrev_tag_entry.get().strip()
        sheet_name = self.sheet_entry.get().strip()
        cell_range = self.range_entry.get().strip()

        if not measured or not abbrev_name or not abbrev_tag or not sheet_name or not cell_range:
            self.err_var.set(
                self._descriptor_error_message(measured, abbrev_name, abbrev_tag, sheet_name, cell_range)
            )
            return

        if not self.selected_files:
            messagebox.showinfo("Run", "No files selected.")
            return

        if not self.selected_template:
            messagebox.showinfo("Template", "Please select a COSINOR template.")
            return

        self.err_var.set("")
        self._set_status("")
        threading.Thread(
            target=self._run_worker,
            args=(measured, abbrev_name, abbrev_tag, sheet_name, cell_range, self.selected_template),
            daemon=True
        ).start()

    def _run_worker(self, measured: str, abbrev_name: str, abbrev_tag: str, sheet_name: str, cell_range: str, template_key: str):
        try:
            rmd_text = _load_cosinor_rmd_text(template_key)
        except Exception as e:
            self._set_status(f"Error: {e}")
            self._show_error("Error", str(e))
            return

        try:
            for i, fpath in enumerate(self.selected_files, start=1):
                src_path = Path(fpath)
                out_dir = src_path.parent / COSINOR_OUTPUT_FOLDER_NAME / src_path.stem
                out_dir.mkdir(parents=True, exist_ok=True)

                # ensure here::here() resolves under the output folder
                try:
                    (out_dir / ".here").write_text("", encoding="utf-8")
                except Exception:
                    pass

                self._set_status(f"Running COSINOR {i}/{len(self.selected_files)}: {src_path.name}")

                script_text = _build_cosinor_script(
                    rmd_text=rmd_text,
                    input_path=str(src_path),
                    measured_compound=measured,
                    abbrev_compound_name=abbrev_name,
                    abbrev_tag=abbrev_tag,
                    sheet_name=sheet_name,
                    cell_range=cell_range,
                )

                script_path = out_dir / f"run_cosinor_{template_key}.R"
                script_path.write_text(script_text, encoding="utf-8")

                _run_r_script(script_path, out_dir)

            self._set_status("Done. Outputs saved under: " + str(Path(self.selected_files[0]).parent / COSINOR_OUTPUT_FOLDER_NAME))
            self._show_info("Success", "COSINOR completed successfully for all selected file(s).")
        except subprocess.CalledProcessError as e:
            msg = (e.stderr or e.stdout or str(e)).strip()
            self._set_status("Error running R: " + msg[:3000])
            self._show_error("Error (R)", msg[:3000])
        except Exception as e:
            self._set_status("Error: " + str(e))
            self._show_error("Error", str(e))

if __name__ == "__main__":
    App().mainloop()
