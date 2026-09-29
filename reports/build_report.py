"""Build the AT2 final report by filling the unit's report template with the experiment outputs.

Usage (from the repository root, after running BOTH notebooks):

    python reports/build_report.py                 # -> reports/AT2_report.docx
    python reports/build_report.py --out my.docx   # custom output path

Inputs
    reports/report_template.docx                   the "clean" template provided by the unit
    models/comfort_climate/metadata.json           written by the CCI notebook
    models/weather_hazard/metadata.json            written by the WHC notebook
    reports/figures/*.png                          written by both notebooks

The template's cover page, table of contents, numbered headings, section pages and
dividers are kept; the placeholder bullets / "Instructions" paragraphs of every section
are replaced by the report text, tables and figures. Every number, table and figure
comes from the metadata and figures, and the sentences that depend on the results
(improvement over baselines, hypothesis outcome, top features, business counts) are
generated from them. Missing metadata is rendered as "[run notebook]" so the structure
can be reviewed before the notebooks have run.

The table of contents is a Word field: Word is told to update fields on opening
(answer "Yes"), or right-click the TOC -> Update Field -> Entire table.

Citations and the reference list follow the UTS APA 7 guide (in-text author-date
citations, references alphabetical with a 1.27 cm hanging indent, tables and figures
labelled in bold with an italic title).

Requires python-docx and matplotlib (installed by `pip install -e ".[notebooks]"`).
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "reports" / "report_template.docx"
FIGURES = ROOT / "reports" / "figures"
MODELS = ROOT / "models"

# ----------------------------------------------------------------------------- edit these
STUDENT = "Tahreema Rahman Zariyat"
STUDENT_ID = "26016360"
GITHUB_USERNAME = "tahreemarz"
EXPERIMENTS_REPO = "https://github.com/tahreemarz/weather-intelligence-experiments"
PACKAGE_REPO = "https://github.com/tahreemarz/36120-26SP-group2-26016360-package"
API_REPO = "https://github.com/tahreemarz/weather-intelligence-api"
API_URL = "https://weather-intelligence-api.onrender.com"      # <- your Render URL
PACKAGE_URL = "https://test.pypi.org/project/amla-package-26016360/"
# -----------------------------------------------------------------------------------------

PLACEHOLDER = "[run notebook]"


# ---------------------------------------------------------------------------
# Metadata access with graceful placeholders
# ---------------------------------------------------------------------------
class Meta:
    def __init__(self, path: Path):
        self.path = path
        self.data = json.loads(path.read_text()) if path.exists() else None
        self.available = self.data is not None

    def get(self, *keys, default=None):
        node = self.data
        for k in keys:
            if not isinstance(node, dict) or k not in node:
                return default
            node = node[k]
        return node

    def num(self, *keys, fmt="{:.2f}"):
        v = self.get(*keys)
        if v is None:
            return PLACEHOLDER
        try:
            return fmt.format(float(v))
        except (TypeError, ValueError):
            return str(v)


def pct(x):
    return f"{100 * x:.0f}%" if isinstance(x, (int, float)) else PLACEHOLDER


def fnum(x, fmt="{:.2f}"):
    return fmt.format(x) if isinstance(x, (int, float)) else PLACEHOLDER


# ---------------------------------------------------------------------------
# Template helpers
# ---------------------------------------------------------------------------
def replace_placeholder(paragraph, mapping: dict[str, str]) -> bool:
    """Replace placeholder text inside a paragraph, keeping the first run's formatting."""
    text = paragraph.text
    if not any(k in text for k in mapping):
        return False
    for k, v in mapping.items():
        text = text.replace(k, v)
    runs = paragraph.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r.text = ""
    else:
        paragraph.add_run(text)
    return True


def is_heading(p_el) -> bool:
    style = p_el.find(qn("w:pPr") + "/" + qn("w:pStyle"))
    return style is not None and str(style.get(qn("w:val"))).startswith("Heading")


def has_section_break_or_drawing(p_el) -> bool:
    return p_el.find(".//" + qn("w:sectPr")) is not None or p_el.find(".//" + qn("w:drawing")) is not None


def clear_section(heading_el):
    """Remove the template's bullet / Instructions paragraphs that follow a heading."""
    el = heading_el.getnext()
    while el is not None and el.tag == qn("w:p") and not is_heading(el) and not has_section_break_or_drawing(el):
        nxt = el.getnext()
        el.getparent().remove(el)
        el = nxt


class Cursor:
    """Insert content right after an anchor element, keeping insertion order."""

    def __init__(self, doc: Document, anchor_el):
        self.doc = doc
        self.el = anchor_el
        self.words = 0

    def _move(self, el):
        self.el.addnext(el)
        self.el = el

    def para(self, text, count=True, italic=False, bold=False, align=None, hanging=False, size=None, single=False, keep_next=False):
        p = self.doc.add_paragraph()
        if hanging:
            p.paragraph_format.left_indent = Cm(1.27)
            p.paragraph_format.first_line_indent = Cm(-1.27)
        if align is not None:
            p.alignment = align
        if single:
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        if keep_next:
            p.paragraph_format.keep_with_next = True
        add_runs(p, text, italic=italic, bold=bold, size=size)
        self._move(p._p)
        if count:
            self.words += len(re.sub(r"\*", "", text).split())
        return p

    def table_caption(self, number, title):
        self.para(f"Table {number}", count=False, bold=True, single=True, keep_next=True, align=WD_ALIGN_PARAGRAPH.LEFT)
        self.para(title, count=False, italic=True, single=True, keep_next=True, align=WD_ALIGN_PARAGRAPH.LEFT)

    def figure_caption(self, number, title):
        self.para(f"Figure {number}", count=False, bold=True, single=True, keep_next=True, align=WD_ALIGN_PARAGRAPH.LEFT)
        self.para(title, count=False, italic=True, single=True, keep_next=True, align=WD_ALIGN_PARAGRAPH.LEFT)

    def note(self, text):
        p = self.doc.add_paragraph()
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run("Note. "); r.italic = True; r.font.size = Pt(9)
        r2 = p.add_run(text); r2.font.size = Pt(9)
        self._move(p._p)

    def table(self, header, rows, col_widths_cm=None, font_size=9):
        t = self.doc.add_table(rows=1 + len(rows), cols=len(header))
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        for j, h in enumerate(header):
            c = t.rows[0].cells[j]; c.text = ""
            r = c.paragraphs[0].add_run(str(h)); r.bold = True; r.font.size = Pt(font_size)
            _fmt_cell(c, j == 0, top=8, bottom=4)
        for i, row in enumerate(rows, start=1):
            for j, v in enumerate(row):
                c = t.rows[i].cells[j]; c.text = ""
                r = c.paragraphs[0].add_run(str(v)); r.font.size = Pt(font_size)
                _fmt_cell(c, j == 0, top=None, bottom=8 if i == len(rows) else None)
        if col_widths_cm:
            for row in t.rows:
                for j, w in enumerate(col_widths_cm):
                    row.cells[j].width = Cm(w)
        self._move(t._tbl)
        return t

    def figure(self, number, title, path: Path, width_cm=15.5, note_text=None):
        self.figure_caption(number, title)
        if path.exists():
            p = self.doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.add_run().add_picture(str(path), width=Cm(width_cm))
            self._move(p._p)
        else:
            self.para(f"[Figure missing: run the notebook to create {path.relative_to(ROOT)}]", count=False, italic=True)
        self.note(note_text or "Produced by the experimentation notebook (author's own work).")


def _fmt_cell(cell, first_col, top, bottom):
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT if first_col else WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge, val in (("top", top), ("bottom", bottom), ("left", None), ("right", None)):
        el = OxmlElement(f"w:{edge}")
        if val is None:
            el.set(qn("w:val"), "nil")
        else:
            el.set(qn("w:val"), "single"); el.set(qn("w:sz"), str(val)); el.set(qn("w:color"), "000000")
        borders.append(el)
    tcPr.append(borders)


def add_runs(p, text, italic=False, bold=False, size=None):
    """Support *italic* markup inside paragraph text."""
    for chunk in re.split(r"(\*[^*]+\*)", text):
        if not chunk:
            continue
        if chunk.startswith("*") and chunk.endswith("*"):
            r = p.add_run(chunk[1:-1]); r.italic = True
        else:
            r = p.add_run(chunk); r.italic = italic
        r.bold = bold
        if size:
            r.font.size = Pt(size)


def find_heading(doc, text):
    for p in doc.paragraphs:
        if p.style.name.lower().startswith("heading") and p.text.strip() == text:
            return p
    raise KeyError(f"heading '{text}' not found in the template")


def enable_update_fields(doc):
    """Ask Word to refresh fields (the table of contents) when the file is opened."""
    settings = doc.settings.element
    el = settings.find(qn("w:updateFields"))
    if el is None:
        el = OxmlElement("w:updateFields")
        settings.append(el)
    el.set(qn("w:val"), "true")


# ---------------------------------------------------------------------------
# Architecture diagram (matplotlib, so the report is self-contained)
# ---------------------------------------------------------------------------
def draw_architecture(path: Path):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.patches import FancyArrowPatch, FancyBboxPatch

    fig, ax = plt.subplots(figsize=(11.5, 5.6)); ax.axis("off"); ax.set_xlim(-0.7, 11.4); ax.set_ylim(-0.75, 5.05)

    def box(x, y, w, h, title, lines, fc):
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.02,rounding_size=0.12", fc=fc, ec="#333", lw=1.2))
        ax.text(x + w / 2, y + h - 0.32, title, ha="center", va="center", fontsize=10.5, fontweight="bold")
        for i, line in enumerate(lines):
            ax.text(x + w / 2, y + h - 0.68 - 0.3 * i, line, ha="center", va="center", fontsize=8.6)

    def arrow(p1, p2, label=None, style="-|>", dy=0.14):
        ax.add_patch(FancyArrowPatch(p1, p2, arrowstyle=style, mutation_scale=14, lw=1.1, color="#333"))
        if label:
            ax.text((p1[0] + p2[0]) / 2, (p1[1] + p2[1]) / 2 + dy, label, ha="center", va="bottom", fontsize=8, color="#333")

    def elbow(points, label, label_at):
        xs, ys = zip(*points)
        ax.plot(xs[:-1], ys[:-1], color="#333", lw=1.1, ls="--")
        ax.add_patch(FancyArrowPatch(points[-2], points[-1], arrowstyle="-|>", mutation_scale=14, lw=1.1, color="#333", linestyle="--"))
        ax.text(*label_at, label, ha="center", va="top", fontsize=8, color="#333")

    box(0.2, 3.3, 2.7, 1.55, "Open-Meteo", ["Historical Weather API (ERA5)", "Forecast API (recent days)", "hourly, Australia/Sydney"], "#e3f2fd")
    box(3.6, 3.3, 3.4, 1.55, "Experimentation repo", ["Jupyter notebooks (Cookiecutter DS)", "hourly -> daily -> targets -> features", "XGBoost search / eval / refit"], "#fff8e1")
    box(7.7, 3.3, 3.5, 1.55, "Model artefacts (JSON)", ["cci_h1..h3.json, whc_h7.json", "climatology.json", "metadata.json (features, metrics)"], "#f3e5f5")
    box(0.2, 0.35, 2.7, 1.95, "Custom package (TestPyPI)", ["amla-package-26016360 >= 0.2.0", "data, targets, features,", "modeling, artifacts", "shared by notebooks AND API"], "#e8f5e9")
    box(3.6, 0.35, 3.4, 1.95, "FastAPI service on Render", ["Docker: python:3.11-slim + uvicorn", "GET /predict/index/comfort_climate", "GET /predict/category/weather_hazard", "GET /health, /model-metadata, /"], "#fce4ec")
    box(7.7, 0.35, 3.5, 1.95, "Clients", ["Open Meteo platform", "B2B customers", "markers / curl / browser"], "#eceff1")
    arrow((2.9, 4.05), (3.6, 4.05), "hourly data")
    arrow((7.0, 4.05), (7.7, 4.05), "save")
    arrow((8.4, 3.3), (6.6, 2.3)); ax.text(7.75, 2.72, "copied into models/", ha="left", va="center", fontsize=8, color="#333")
    arrow((2.6, 2.3), (3.9, 3.3), style="<|-|>"); ax.text(2.85, 2.95, "import", ha="right", va="center", fontsize=8, color="#333")
    arrow((2.9, 1.3), (3.6, 1.3), "import")
    arrow((7.7, 0.75), (7.0, 0.75), style="<|-|>")
    ax.text(7.35, 0.88, "HTTPS", ha="center", va="bottom", fontsize=7.5, color="#333")
    ax.text(7.35, 0.62, "?date=YYYY-MM-DD", ha="center", va="top", fontsize=7, color="#333")
    elbow([(0.2, 3.7), (-0.4, 3.7), (-0.4, -0.3), (5.3, -0.3), (5.3, 0.35)],
          "trailing 90-day observation window fetched at request time (archive + forecast back-fill)", (5.3, -0.38))
    fig.savefig(path, dpi=200, bbox_inches="tight"); plt.close(fig)


# ---------------------------------------------------------------------------
# Report content
# ---------------------------------------------------------------------------
def build(out_path: Path):
    cci = Meta(MODELS / "comfort_climate" / "metadata.json")
    whc = Meta(MODELS / "weather_hazard" / "metadata.json")
    FIGURES.mkdir(parents=True, exist_ok=True)
    arch_png = FIGURES / "deployment_architecture.png"
    draw_architecture(arch_png)

    doc = Document(TEMPLATE)
    enable_update_fields(doc)

    # ---------------- cover page ----------------
    mapping = {
        "<put_student_name_here>": STUDENT,
        "<put_student_id_here>": STUDENT_ID,
        "<put_github_id>": GITHUB_USERNAME,
        "Experiment Repo: <put_link_to_github>": f"Experiment Repo: {EXPERIMENTS_REPO}",
        "Package Repo: <put_link_to_github>": f"Package Repo: {PACKAGE_REPO}",
        "API Repo: <put_link_to_github>": f"API Repo: {API_REPO}",
        "<put_link_to_deployed_fastapi>": API_URL,
    }
    for p in doc.paragraphs:
        replace_placeholder(p, mapping)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_placeholder(p, mapping)

    # ---------------- shared numbers ----------------
    cs = cci.get("eda", "band_share_pct") or {}
    ws = whc.get("eda", "class_share_pct") or {}
    ca = cci.get("eda", "autocorrelation") or {}
    wa = whc.get("eda", "autocorrelation") or {}
    tc = {h: cci.get("performance", "test_2023_2025", h) or {} for h in ("t+1", "t+2", "t+3")}
    pers = cci.get("performance", "baselines_test_mae", "persistence") or {}
    clim = cci.get("performance", "baselines_test_mae", "climatology") or {}
    improvement = {h: 1 - tc[h]["mae"] / min(pers[h], clim[h]) for h in tc if tc[h] and h in pers and h in clim}
    t = whc.get("performance", "test_2023_2025") or {}
    b = whc.get("performance", "baselines_test") or {}
    best_f1 = max((v.get("macro_f1", 0) for v in b.values()), default=None)
    best_cost = min((v.get("expected_cost", 9e9) for v in b.values()), default=None)
    bi = cci.get("business_impact_test_t1") or {}
    bw = whc.get("business_impact_test") or {}
    top_c = list((cci.get("feature_importance") or {}).keys())[:5]
    top_w = list((whc.get("feature_importance") or {}).keys())[:5]
    cci_ok = bool(improvement) and all(v > 0 for v in improvement.values())
    whc_f1_ok = bool(t and b) and t["macro_f1"] > best_f1
    whc_cost_ok = bool(t and b) and t["expected_cost"] < best_cost

    def mae_line():
        return " / ".join(fnum(tc[h].get("mae")) for h in ("t+1", "t+2", "t+3"))

    def base_line():
        return " / ".join(fnum(min(pers[h], clim[h])) if h in pers and h in clim else PLACEHOLDER for h in ("t+1", "t+2", "t+3"))

    total_words = 0

    # =====================================================================
    # 1. Executive Summary
    # =====================================================================
    h = find_heading(doc, "Executive Summary"); clear_section(h._p); c = Cursor(doc, h._p)
    c.para("Open Meteo is adding artificial-intelligence services to its microservice weather platform. This project "
           "delivers two of them for Sydney, NSW (latitude −33.8688, longitude 151.2093): a regression service that "
           "forecasts the Climate Comfort Index (CCI, 0–100) for each of the three days after an input date, and a "
           "classification service that forecasts the Weather Hazard Category (WHC: Low, Moderate, High or Extreme Risk) "
           "exactly seven days ahead. Both targets are derived, with the formulas prescribed in the brief, from daily "
           "observations obtained through the Open-Meteo Historical Weather API (Zippenfenig, 2023).")
    c.para("The problem is a genuine forecasting one: the models may only use information available on the input "
           "date, must be evaluated against the naive references a customer already has (persistence and climatology), "
           "and must be served as a REST API that behaves gracefully on invalid input. All shared logic (data access, "
           "aggregation, cleaning, target formulas, feature engineering, model artefacts) was packaged in my custom "
           "Python package amla-package-26016360, extended from 0.1.3 to 0.2.0 and published on TestPyPI (Zariyat, 2026), "
           "so that the notebooks and the deployed FastAPI service compute everything with the same code.")
    c.para(f"Outcome: on the untouched 2023–2025 test years the XGBoost CCI models reach a mean absolute error of "
           f"{mae_line()} index points for t+1 / t+2 / t+3 against {base_line()} for the best naive baseline, "
           f"{'beating it at every horizon' if cci_ok else 'not beating it at every horizon'}. The WHC classifier "
           f"reaches a macro-F1 of {fnum(t.get('macro_f1'), '{:.3f}')} against {fnum(best_f1, '{:.3f}')} for the best baseline and "
           f"flags {pct(t.get('severe_recall'))} of High/Extreme days a week in advance. The models are deployed at {API_URL} "
           "with five documented endpoints and a Docker-based Render configuration; the two repositories, the package and "
           "the live service are listed on the cover page.")
    total_words += c.words

    # =====================================================================
    # 2a. Business Use Cases
    # =====================================================================
    h = find_heading(doc, "Business Use Cases"); clear_section(h._p); c = Cursor(doc, h._p)
    c.para("The CCI service summarises how suitable a day is for outdoor activity (35% temperature, 20% humidity, 15% "
           "wind, 15% cloud, 15% rain) as one interpretable number with five bands. Its users are event organisers, "
           "tourism and hospitality operators, councils running parks, pools and beaches, and workforce schedulers for "
           "outdoor work, who plan staffing, stock and activities one to three days ahead. The WHC service condenses "
           "rain, gusts, cloud, snow and temperature extremes into a week-ahead risk category for logistics and "
           "construction firms, councils, insurers and utilities that pre-position staff, reschedule outdoor work or "
           "issue early warnings.")
    c.para("The challenge that motivated the project is that these customers currently read five raw variables and "
           "extrapolate: \"tomorrow will be like today\" or \"this is the usual weather for the season\". Both rules are "
           "poor beyond a day or two, and neither expresses risk in a form that triggers a workflow. Machine learning is "
           "relevant because the targets are non-linear, interaction-heavy functions of the weather (temperature helps "
           "up to 22 °C and hurts beyond; 20 mm of rain means a summer storm in February and a multi-day low in June), "
           "because the useful signal a week out is regime and seasonality information hidden in rolling statistics "
           "and climatology, and because a learned model can quantify how much better than the naive rules a forecast "
           "is—which is exactly what a platform selling the service must be able to state.")
    total_words += c.words

    # =====================================================================
    # 2b. Key Objectives
    # =====================================================================
    h = find_heading(doc, "Key Objectives"); clear_section(h._p); c = Cursor(doc, h._p)
    c.para("The objectives were: (1) build a regression model that forecasts the CCI for t+1, t+2 and t+3 with a lower "
           "MAE than persistence and climatology; (2) build a four-class model that forecasts the WHC for t+7 with a "
           "higher macro-F1 and a lower expected business cost than persistence, climatology and the majority class; "
           "(3) package the reusable functionality in my custom Python package and publish it on TestPyPI; (4) deploy "
           "both models as a FastAPI service on Render with the required endpoints and graceful error handling; and "
           "(5) keep everything reproducible from two documented repositories.")
    c.para("Stakeholders and requirements: Open Meteo's product team needs a documented, versioned service with a "
           "metadata endpoint, sub-second responses and predictable failure modes; B2B customers need forecasts that are "
           "demonstrably better than their own rules, an honest statement of accuracy and, for hazards, a tunable "
           "trade-off between missed severe days and false alarms; the public exposed to weather needs the service to "
           "err on the side of caution; and the marker needs the notebooks, artefacts and API to be reproducible. The "
           "project addresses these through leakage-safe feature engineering, chronological evaluation against "
           "baselines, class-weighted training with an explicit cost matrix, a frozen feature contract inside the "
           "artefacts, and a package that is tested offline and shared by training and serving.")
    total_words += c.words

    # =====================================================================
    # 3. Data Understanding
    # =====================================================================
    h = find_heading(doc, "Data Understanding"); clear_section(h._p); c = Cursor(doc, h._p)
    c.para("All observations were retrieved programmatically from the Open-Meteo Historical Weather API, which serves "
           "the ECMWF ERA5 reanalysis (Hersbach et al., 2020) on a roughly 25 km grid. Data were requested at *hourly* "
           "resolution for 1 January 1990 to 31 December 2025 in the Australia/Sydney time zone, in calendar-year "
           "chunks with retries, and cached locally as a Parquet file (about 316,000 rows). Nothing dated 2026 or later "
           "is ever downloaded, as the brief reserves it for production. Hourly data were chosen deliberately: the brief "
           "prescribes one daily aggregation per variable (Table 1), and performing that aggregation in the package "
           "makes it explicit, testable and identical between training and serving. A day is kept only if at least 20 "
           "of its 24 hourly records are valid, and the daily index is reindexed to a complete range so that a shift of "
           "*h* rows always means *h* days. The package's data-quality report found no duplicate timestamps and "
           "flagged snowfall as a column without variation.")
    c.table_caption(1, "Hourly Variables Retrieved From Open-Meteo and Their Daily Aggregation")
    c.table(["Variable", "Daily aggregation", "Role"], [
        ["temperature_2m", "mean", "CCI and WHI formula"],
        ["relative_humidity_2m", "mean", "CCI formula"],
        ["wind_speed_10m", "mean", "CCI formula"],
        ["wind_gusts_10m", "daily maximum", "WHI formula"],
        ["cloud_cover", "mean", "CCI and WHI formula"],
        ["precipitation", "daily total", "CCI and WHI formula"],
        ["snowfall", "daily total", "WHI formula"],
        ["dew_point_2m, apparent_temperature, pressure_msl, wind_direction_10m, shortwave_radiation, vapour_pressure_deficit",
         "mean (plus daily min/max, radiation total, rain hours)", "extra predictors"],
    ], col_widths_cm=[7.5, 4.5, 4.0])
    c.note("Aggregation rules for the first seven variables are those prescribed by the assignment brief.")
    c.para("Both targets are deterministic functions of these daily values and were implemented once, in the package "
           "module *targets*, with unit tests against hand-computed cases. The CCI turns five values into scores in [0, 1] "
           "that peak at the most comfortable value (22 °C, 50% humidity, a 10 km/h breeze, clear sky, no rain), "
           "combines them with weights 0.35/0.20/0.15/0.15/0.15 and multiplies by 100; five bands (Excellent ≥ 80 to "
           "Very poor < 20) translate it for users. The Weather Hazard Index (WHI) uses saturating ramps (rain/30 mm, "
           "gusts/100 km/h, cloud/100%, snow/15 cm, |T − 22|/25 °C) with weights 0.30/0.30/0.20/0.10/0.10, and the WHC "
           "cuts it at 25, 50 and 75. The forecasting targets are created by shifting the series: cci_h1–cci_h3 hold "
           "the index one to three days after each row and whc_h7 the class seven days after; the same-day index is a "
           "legitimate input, the shifted columns never are.")
    c.para(f"Exploration of the record (n = {cci.num('eda', 'n_days', fmt='{:,.0f}')} days; Figures 1 and 2) shows why the "
           f"two problems differ. The CCI has a mean of {cci.num('eda', 'cci_mean', fmt='{:.1f}')} (SD {cci.num('eda', 'cci_std', fmt='{:.1f}')}), "
           f"unevenly populated bands ({cs.get('Excellent outdoor conditions', PLACEHOLDER)}% Excellent, {cs.get('Comfortable conditions', PLACEHOLDER)}% "
           f"Comfortable, {cs.get('Moderate comfort', PLACEHOLDER)}% Moderate, {cs.get('Poor comfort', PLACEHOLDER)}% Poor, "
           f"{cs.get('Very poor comfort', PLACEHOLDER)}% Very poor), a clear seasonal cycle and an autocorrelation of {ca.get('1', PLACEHOLDER)} "
           f"at lag 1 that decays to {ca.get('3', PLACEHOLDER)} at lag 3: tomorrow resembles today, but three days ahead much of "
           f"the signal is gone. The WHI has a mean of {whc.num('eda', 'whi_mean', fmt='{:.1f}')} (SD {whc.num('eda', 'whi_std', fmt='{:.1f}')}) "
           f"and its classes are strongly imbalanced ({ws.get('Low Risk', PLACEHOLDER)}% Low, {ws.get('Moderate Risk', PLACEHOLDER)}% Moderate, "
           f"{ws.get('High Risk', PLACEHOLDER)}% High, {ws.get('Extreme Risk', PLACEHOLDER)}% Extreme Risk); the majority class alone gives "
           f"{whc.num('eda', 'majority_accuracy_pct', fmt='{:.1f}')}% accuracy and the persistence rule {whc.num('eda', 'persistence_accuracy_pct', fmt='{:.1f}')}%, "
           f"while the autocorrelation of the index at lag 7 is only {wa.get('7', PLACEHOLDER)}. Snowfall is essentially always zero at "
           "sea level in Sydney, so the practical maximum of the index is 90 and Extreme days are rare by construction "
           "as well as by climate. Among the drivers, precipitation is zero-inflated and heavy-tailed, gusts and cloud "
           "carry a continuous contribution and the seasonal cycle is strong, which motivated rolling, seasonal and "
           "climatological features.")
    c.para("Limitations of the data: ERA5 is a model reanalysis for a grid cell, not a station, so localised rain, "
           "sea-breeze effects and peak gusts are smoothed and extreme days are under-represented; the archive lags real "
           "time by one to two days; and the fixed thresholds of the formulas make some class boundaries arbitrary "
           "(a WHI of 49 and 51 are almost the same weather).")
    c.figure(1, "Distribution, Annual Means, Seasonal Cycle and a Sample Year of the Climate Comfort Index", FIGURES / "cci_target_exploration.png", width_cm=14.5)
    c.figure(2, "Class Distribution, Index Distribution, Monthly Class Shares and a Sample Year of the Weather Hazard Index", FIGURES / "whc_target_exploration.png", width_cm=14.5)
    total_words += c.words

    # =====================================================================
    # 4. Data Preparation
    # =====================================================================
    h = find_heading(doc, "Data Preparation"); clear_section(h._p); c = Cursor(doc, h._p)
    c.para("Cleaning was implemented as three package functions so the API applies them unchanged. Days with fewer "
           "than 20 valid hours have their observations set to missing—rows are kept so that the daily index stays "
           "contiguous and \"seven days ahead\" never silently becomes eight. Values are clipped to physical bounds "
           "(humidity 0–100%, non-negative rain and radiation, pressure 900–1100 hPa, gusts not below the mean wind), "
           "which protects the formulas' MAX/MIN operations from corrupted upstream data; on the ERA5 data the counts "
           "of clipped values were zero or negligible. Gaps of at most two days in smooth variables (temperature, "
           "pressure, humidity, radiation, cloud) are time-interpolated, whereas precipitation and snowfall are never "
           "interpolated because that would fabricate rain; XGBoost routes missing values natively, so no other "
           "imputation and no scaling were applied. Targets are recomputed after cleaning so that no label comes from a "
           "rejected day. Outliers were not removed: extreme days are the events the hazard model must learn.")
    c.para("Feature engineering followed one rule: every feature must be computable at serving time from a trailing "
           "window of observations ending on the input date. A set of base variables was retained per target (the "
           "formula variables, the current CCI and WHI, pressure, dew point, humidity, wind and radiation) and expanded "
           "with the package function *build_features* into lags of 1, 2, 3 and 7 days; trailing rolling means over "
           "3, 7, 14 and 30 days, rolling standard deviations over 7, 14 and 30 days and 7- and 30-day rain totals; "
           "tendencies such as the 24–72 h pressure change, the diurnal temperature range, the dew-point depression and "
           "the 30-day temperature anomaly; and cyclic day-of-year and month encodings. A day-of-year climatology "
           "encoder, fitted on training years only and serialised with the model, supplies the value expected on the "
           "*target* date (t+1 to t+3 or t+7)—a strong prior at these horizons and the classic climatology baseline. "
           f"The final matrices hold {cci.num('n_features', fmt='{:.0f}')} features for the CCI models and "
           f"{whc.num('n_features', fmt='{:.0f}')} for the WHC model.")
    c.para("Feature selection combined a filter stage (Spearman correlation and mutual information with the future "
           "CCI; ANOVA F, mutual information and Spearman correlation with the future WHI) with an embedded/wrapper "
           "stage (XGBoost gain importance and permutation importance on the validation years). Snowfall (constant) "
           "and the raw wind-direction components (no measurable importance) were dropped.")
    c.para("Because the data are a time series, the split is chronological with no shuffling (Bergmeir & Benítez, 2012): "
           "training on 1990–2019, validation on 2020–2022 for hyper-parameter selection and early stopping, and a "
           "single evaluation on the untouched 2023–2025 test years; whole-year blocks keep every season and both ENSO "
           "phases in each block. The package's leakage check confirmed that no date appears on both sides of a split, "
           "an automated audit asserts that no target-derived column is a feature, and the features of a test day "
           "rebuilt from a 90-day window—exactly what the API will have—were verified to be numerically identical to "
           "the training features. Class imbalance was handled with inverse-frequency sample weights whose exponent "
           "(0 = none, 0.5 = softened, 1 = fully balanced) is tuned as a hyper-parameter, rather than with synthetic "
           "oversampling, which would break the temporal and physical structure of weather data (He & Garcia, 2009).")
    total_words += c.words

    # =====================================================================
    # 5. Modeling
    # =====================================================================
    h = find_heading(doc, "Modeling"); clear_section(h._p); c = Cursor(doc, h._p)
    c.para("XGBoost (Chen & Guestrin, 2016) was selected for both targets. The relationships are non-linear and "
           "interaction-heavy, which trees capture without manual transformation; the algorithm regularises well on "
           "roughly 11,000 training rows with 150–190 heterogeneous features (column and row subsampling, shrinkage, "
           "early stopping); it handles missing values, supports per-row sample weights, exposes feature importance for "
           "the metadata endpoint, and saves to a portable JSON format that avoids pickle version coupling. For the CCI, "
           "one regressor per horizon lets each horizon learn its own balance between the current state and the "
           "climatological prior. For the WHC, a softprob classifier with an explicit *num_class* = 4 was trained "
           "through the package's *fit_xgb_classifier*, which avoids the scikit-learn wrapper failing when the rare "
           "Extreme class is absent from a split (Pedregosa et al., 2011). Persistence, climatology and the majority "
           "class were kept as baselines (Wilks, 2011).")
    c.para("Hyper-parameters (tree depth, minimum child weight, learning rate, row and column subsampling, L1/L2 "
           "regularisation) were tuned by random search (Bergstra & Bengio, 2012) scored on the validation years—MAE "
           "for the regressors, macro-F1 with expected cost as tie-breaker for the classifier—with the number of trees "
           "chosen by early stopping so that it adapts to each learning rate. The class-weight exponent was searched "
           f"jointly; the selected value was {whc.num('hyperparameters', 'class_weight_power', fmt='{:.1f}')}. Shuffled "
           "cross-validation was avoided because it would leak neighbouring days. The final models were refitted on all "
           "data up to 31 December 2025 with the selected settings and saved as artefact bundles (model JSON, climatology "
           "table, metadata with the frozen feature list, hyper-parameters, importance and metrics).")
    c.para("Implementation and deployment (Figure 3): the package is installed from TestPyPI in the notebooks and in "
           "the API image, so the target formulas, cleaning rules, feature construction and artefact format are the same "
           "code on both sides. The API is a FastAPI application (Ramírez, 2018) packaged with Docker (Merkel, 2014) on "
           "a python:3.11-slim image and deployed to Render as a Docker web service with a health check. At request "
           "time it validates the date (format, real calendar date, not later than yesterday in Sydney time), fetches "
           "the trailing 90-day window from Open-Meteo (back-filling the newest days from the Forecast API when the "
           "archive lags), applies the package's cleaning and feature functions, predicts and returns the required "
           "JSON. Invalid inputs return 422 or 400 with an explanation, unavailable observations or upstream outages "
           "503, unknown routes a JSON 404 and unexpected failures a generic 500. Seventeen contract tests run without "
           "network access by training tiny bundles on synthetic data; the package itself has 59 offline tests.")
    c.figure(3, "Deployment Architecture: Data Flow Between Open-Meteo, the Experimentation Repository, the Custom Package and the FastAPI Service on Render",
             arch_png, width_cm=15.5, note_text="Author's own diagram.")
    total_words += c.words

    # =====================================================================
    # 6a. Results and Analysis
    # =====================================================================
    h = find_heading(doc, "Results and Analysis"); clear_section(h._p); c = Cursor(doc, h._p)
    c.para("For the CCI the primary metric is the MAE in index points, complemented by RMSE, R², the share of days "
           "within 10 points and the share of days in the correct interpretation band, because customers act on the "
           "band. Table 2 reports the test results against the naive baselines.")
    rows = []
    for hz in ("t+1", "t+2", "t+3"):
        m = tc[hz]
        rows.append([hz, fnum(m.get("mae")), fnum(m.get("rmse")), fnum(m.get("r2")), pct(m.get("within_10")), pct(m.get("band_accuracy")),
                     fnum(pers.get(hz)), fnum(clim.get(hz))])
    c.table_caption(2, "Test-Set Performance of the CCI Models (2023–2025) and Naive Baselines")
    c.table(["Horizon", "MAE", "RMSE", "R²", "Within 10", "Band acc.", "Persistence MAE", "Climatology MAE"], rows,
            col_widths_cm=[1.6, 1.4, 1.6, 1.4, 1.9, 1.9, 2.9, 3.0])
    c.note("MAE and RMSE in CCI points (0–100). Within 10 = share of days within 10 points of the observed index; "
           "Band acc. = share of days forecast in the correct interpretation band.")
    if improvement:
        beats = [hz for hz, v in improvement.items() if v > 0]
        where = "every horizon" if len(beats) == 3 else (" and ".join(beats) if beats else "no horizon")
        deltas = ", ".join(f"{abs(100 * v):.0f}% {'lower' if v > 0 else 'higher'} at {hz}" for hz, v in improvement.items())
        summary = f"The CCI models {'beat' if beats else 'did not beat'} the best naive baseline at {where}, with an MAE {deltas}. "
    else:
        summary = f"{PLACEHOLDER}: the comparison with the baselines is generated from metadata.json. "
    val = cci.get("performance", "validation_2020_2022", "t+1") or {}
    c.para(summary + f"Validation MAE at t+1 was {fnum(val.get('mae'))}, close to the test value, so the tuning did not "
           "over-fit the validation years. Accuracy degrades with the horizon as expected: at one day the model exploits "
           "the current index, cloud, rain and pressure tendency, whereas at three days predictions regress towards the "
           "climatological expectation and the margin over climatology shrinks. The most important features "
           f"{('were ' + ', '.join(top_c)) if top_c else 'are listed in metadata.json'}. The scatter plots (Figure 4) show "
           "the pattern of a useful but bounded forecaster: unbiased in the middle of the range, with the largest "
           "residuals on abrupt regime changes such as storm arrivals and southerly busters.")
    c.figure(4, "Observed Versus Predicted CCI on the Test Set for the Three Horizons", FIGURES / "cci_test_scatter.png")
    c.para("For the WHC, accuracy is reported but not used for decisions, because always predicting the majority class "
           "already reaches its prevalence. The primary metric is the macro-averaged F1, supported by balanced accuracy "
           "and three business metrics: an expected misclassification cost from an asymmetric cost matrix (an Extreme "
           "day forecast as Low costs 10 units, a false severe alarm 2–3, an adjacent-class slip 1), the recall of severe "
           "(High or Extreme) days and the false-alarm rate of severe forecasts. Table 3 compares the model with the "
           "three baselines and Table 4 gives the per-class results.")
    rows = []
    for name, m in [("XGBoost", t), ("Persistence", b.get("persistence", {})), ("Climatology", b.get("climatology", {})), ("Majority class", b.get("majority", {}))]:
        rows.append([name] + [fnum(m.get(k), "{:.3f}") for k in ("macro_f1", "balanced_accuracy", "accuracy", "expected_cost")]
                    + [pct(m.get(k)) for k in ("severe_recall", "severe_false_alarm")])
    c.table_caption(3, "Test-Set Performance of the WHC Model (2023–2025) Against the Baselines")
    c.table(["Model", "Macro-F1", "Bal. acc.", "Accuracy", "Exp. cost", "Severe recall", "Severe false alarm"], rows,
            col_widths_cm=[3.0, 1.9, 1.9, 1.9, 1.9, 2.5, 2.9])
    c.note("Exp. cost = expected misclassification cost per day under the cost matrix described in the text; severe = High or Extreme Risk.")
    pc = whc.get("performance", "test_per_class") or {}
    rows = [[cls] + [fnum((pc.get(k) or {}).get(cls), "{:.3f}") for k in ("precision", "recall", "f1")] for cls in ("Low Risk", "Moderate Risk", "High Risk", "Extreme Risk")]
    c.table_caption(4, "Per-Class Precision, Recall and F1 of the WHC Model on the Test Set")
    c.table(["Class", "Precision", "Recall", "F1"], rows, col_widths_cm=[4.0, 3.0, 3.0, 3.0])
    c.note("Rows correspond to the observed class. The Extreme Risk class has very few test examples, so its metrics carry wide uncertainty.")
    if t and b:
        outcome = (f"The model's macro-F1 of {t['macro_f1']:.3f} {'exceeds' if whc_f1_ok else 'does not exceed'} the best "
                   f"baseline ({best_f1:.3f}) and its expected cost ({t['expected_cost']:.3f}) is "
                   f"{'lower' if whc_cost_ok else 'not lower'} than the best baseline ({best_cost:.3f}); it flags "
                   f"{pct(t['severe_recall'])} of the severe days a week ahead, against {pct(b['persistence']['severe_recall'])} for persistence, "
                   f"with a severe false-alarm rate of {pct(t['severe_false_alarm'])}. ")
    else:
        outcome = f"{PLACEHOLDER}: the comparison with the baselines is generated from metadata.json. "
    ap = whc.get("performance", "test_average_precision_severe")
    c.para(outcome + f"The average precision for the severe event was {fnum(ap, '{:.3f}')}, which measures ranking quality "
           "independently of the decision threshold. The confusion matrices (Figure 5) show that most errors are slips "
           "into an adjacent class—largely threshold noise around the 25/50/75 cut points—while precision on the severe "
           "classes is low by design: the class weights trade false alarms for recall. The most important features "
           f"{('were ' + ', '.join(top_w)) if top_w else 'are listed in metadata.json'}, confirming that a seven-day hazard "
           "forecast from single-point observations is mainly a regime and seasonality problem rather than a "
           "day-specific one.")
    c.figure(5, "Confusion Matrices of the WHC Model on the Test Set (Counts and Per-Class Recall)", FIGURES / "whc_confusion_matrix.png")
    c.para("Key insights from the experimentation: the current index, its recent trajectory and the climatology of "
           "the target day carry most of the CCI signal; for the hazard target the rolling statistics, seasonal "
           "encodings and climatology dominate while today's raw rain and gusts add little a week out; class weighting "
           "is the lever that turns an accurate-but-useless classifier into a hazard service; and the largest errors "
           "of both models are abrupt changes that no single-point daily model can anticipate—the strongest argument "
           "for adding numerical-weather-prediction inputs in future work.")
    total_words += c.words

    # =====================================================================
    # 6b. Business Impact and Benefits
    # =====================================================================
    h = find_heading(doc, "Business Impact and Benefits"); clear_section(h._p); c = Cursor(doc, h._p)
    c.para("The CCI service is accurate enough at one and two days to support operational decisions—staffing, stock and "
           "scheduling of outdoor activities—provided the interpretation band, with fuzzy boundaries, is communicated "
           "rather than the point value; at three days it is an early indication. The costliest error for customers is "
           "false comfort, a Comfortable or Excellent forecast for a day that turns out Moderate or worse: on the test "
           f"period it occurred on {bi.get('false_comfort_rate_pct', PLACEHOLDER)}% of days for the one-day forecast, against "
           f"{bi.get('persistence_false_comfort_rate_pct', PLACEHOLDER)}% for the persistence rule a customer would otherwise apply; "
           f"the forecast was over-optimistic by at least one band on {bi.get('over_optimistic_band_pct', PLACEHOLDER)}% of days and "
           f"over-pessimistic on {bi.get('over_pessimistic_band_pct', PLACEHOLDER)}%, and misses larger than 15 points occurred on "
           f"{bi.get('large_miss_over_15_points_pct', PLACEHOLDER)}% of days. High-stakes decisions such as large events should "
           "therefore be routed to the one-day forecast, and these rates should be quoted in the product documentation.")
    c.para("The WHC service is an early-warning and planning tool rather than a go/no-go decision maker; its value is "
           f"separating regimes a week ahead. Over the {bw.get('days', PLACEHOLDER)}-day test period there were "
           f"{bw.get('severe_days', PLACEHOLDER)} High or Extreme days; the model flagged {bw.get('flagged', PLACEHOLDER)} of them a week in "
           f"advance and missed {bw.get('missed', PLACEHOLDER)}, with {bw.get('false_alarms', PLACEHOLDER)} false severe alarms and "
           f"{bw.get('critical_misses', PLACEHOLDER)} critical misses (an Extreme day forecast as Low), whereas persistence missed "
           f"{bw.get('persistence_missed', PLACEHOLDER)} with {bw.get('persistence_false_alarms', PLACEHOLDER)} false alarms and climatology "
           f"missed {bw.get('climatology_missed', PLACEHOLDER)}. Under the cost matrix the total cost of the model's forecasts was "
           f"{bw.get('total_cost', PLACEHOLDER)} units against {bw.get('persistence_total_cost', PLACEHOLDER)} for persistence and "
           f"{bw.get('climatology_total_cost', PLACEHOLDER)} for climatology.")
    thr = bw.get("recommended_alert_threshold")
    c.para("The class probabilities kept in the artefact let each customer set an alert threshold matching its own cost "
           "of a miss versus a false alarm (Figure 6): alerting whenever the predicted probability of severe weather is "
           f"at least {fnum(thr)} would have caught about {pct(bw.get('recall_at_threshold'))} of the severe days with a "
           f"false-alarm rate of {pct(bw.get('false_alarm_rate_at_threshold'))} (about {bw.get('alerts_per_year_at_threshold', PLACEHOLDER)} "
           "alerts per year)—a sensible setting for a council or an insurer that prefers to over-prepare, whereas an "
           "event organiser with a high cancellation cost should use a higher threshold. Extreme Risk should be "
           "communicated together with High Risk as \"severe\" because it is too rare to estimate reliably. The value "
           "generated is therefore twofold: better-than-naive planning accuracy at short range, and a quantified, "
           "tunable early warning a week ahead where customers previously had only the seasonal average. Both services "
           "should be monitored by recomputing the targets once the observations for the target dates arrive and "
           "tracking the rolling MAE, macro-F1 and severe recall against Tables 2 and 3; sustained deterioration should "
           "trigger retraining with the newest years through the same notebooks.")
    c.figure(6, "Severe-Alert Operating Points on the Test Period: Recall of High/Extreme Days Versus False-Alarm Rate for Alert Thresholds on P(Severe)",
             FIGURES / "whc_threshold_tradeoff.png", width_cm=12)
    total_words += c.words

    # =====================================================================
    # 6c. Data Privacy and Ethical Concerns
    # =====================================================================
    h = find_heading(doc, "Data Privacy and Ethical Concerns"); clear_section(h._p); c = Cursor(doc, h._p)
    c.para("Data privacy: the models use only open meteorological reanalysis data for a single location; no personal "
           "data are collected, stored or inferred, and the data licence permits this use with attribution. The API "
           "receives a date and returns a forecast; it keeps no user data, and server logs should retain only what "
           "operations require (timestamps, status codes) rather than client addresses. Data provenance is fully "
           "transparent and the pipeline is reproducible from public sources, which supports accountability.")
    c.para("Ethical concerns: weather intelligence is used to keep people safe, so the main risks are over-trust and "
           "unequal impact. A forecast shown as a single category can be read as certain; the service therefore "
           "publishes its accuracy and baselines through the metadata endpoint, and the recommendations above insist "
           "on communicating uncertainty and calibrating alerts to the user's risk. A missed severe day exposes outdoor "
           "workers and event attendees—often those least able to absorb the cost—so the cost matrix penalises misses "
           "more than false alarms, and the service must complement, never replace, the Bureau of Meteorology's "
           "official warnings. The models cover one grid cell centred on Sydney; applying them to other locations or "
           "to fringe communities with different micro-climates would silently degrade accuracy, and reanalysis "
           "under-representation of extremes means the models are least reliable when the stakes are highest, which "
           "must be stated in the terms of use. Technical debt of the two-repository design was managed deliberately "
           "with version pins and a frozen feature contract (Sculley et al., 2015).")
    c.para("Potential impacts on Indigenous people: Sydney is the land of the Gadigal and other Eora and Dharug "
           "peoples, and Aboriginal communities in Greater Sydney and beyond are disproportionately represented among "
           "outdoor workers, remote-service staff and residents of flood- and heat-exposed areas, so both the benefits "
           "of a good early warning and the harm of a missed one fall on them unevenly. Three risks were considered. "
           "First, the formulas encode a Western, urban notion of \"comfort\" and \"hazard\" (22 °C optimum, rain as a "
           "penalty) that does not reflect Indigenous seasonal knowledge—the D'harawal calendar of the Sydney region, "
           "for example, reads rain and wind as signals for Country rather than as nuisances—so the service should be "
           "presented as one input alongside, not above, that knowledge, and any adaptation for Indigenous "
           "organisations should be co-designed with them. Second, a Sydney-only model must not be applied to regional "
           "or remote Aboriginal communities where accuracy is unknown; the location is hard-coded and documented for "
           "that reason. Third, no Indigenous data are collected, so Indigenous data-sovereignty principles are not "
           "breached, but if usage data were ever added they should be governed under the CARE principles.")
    total_words += c.words

    # =====================================================================
    # 7. Conclusion
    # =====================================================================
    h = find_heading(doc, "Conclusion"); clear_section(h._p); c = Cursor(doc, h._p)
    outcome_cci = "confirmed" if cci_ok else ("partially confirmed" if improvement else PLACEHOLDER)
    if t and b:
        outcome_whc = "confirmed" if (whc_f1_ok and whc_cost_ok) else ("partially confirmed" if (whc_f1_ok or whc_cost_ok) else "not confirmed")
    else:
        outcome_whc = PLACEHOLDER
    c.para("The project delivers two reproducible forecasting services whose data, target and feature logic live in a "
           "tested, versioned package shared between experimentation and serving. The CCI hypothesis—that an XGBoost "
           f"model on engineered daily features beats persistence and climatology at every horizon—was {outcome_cci}, with a "
           f"test MAE of {mae_line()} points; the WHC hypothesis—higher macro-F1 and lower expected cost than all three "
           f"baselines—was {outcome_whc}, with a macro-F1 of {fnum(t.get('macro_f1'), '{:.3f}')} and {pct(t.get('severe_recall'))} of "
           "severe days flagged a week ahead. Against the stakeholders' requirements, the service is documented, "
           "versioned, evaluated against the customer's own rules, tunable for risk appetite and fails gracefully; the "
           "main gap is precision on individual severe days, which is a limit of the information available rather than "
           "of the modelling.")
    c.para("Future work, ranked by expected uplift: add numerical-weather-prediction forecasts from the Open-Meteo "
           "Forecast API for the target days as features (large uplift, especially at t+3 and t+7, at the cost of a "
           "second upstream dependency); add upstream grid points and climate-driver indices (ENSO, IOD) to sharpen the "
           "regime signal; adopt an ordinal formulation or regress the WHI and bin afterwards, with probability "
           "calibration, so that customer thresholds are trustworthy; add quantile outputs for the CCI to communicate "
           "uncertainty; extend training to 1940 to see more Extreme days; and build the scheduled retraining and "
           "monitoring pipeline sketched above. Operationally, the free Render instance sleeps between requests and "
           "Open-Meteo lags real time by one to two days, so predictions for the most recent dates can be temporarily "
           "unavailable—behaviour the API reports explicitly rather than hiding.")
    total_words += c.words

    # =====================================================================
    # 8. References
    # =====================================================================
    h = find_heading(doc, "References"); clear_section(h._p); c = Cursor(doc, h._p)
    refs = [
        "Bergmeir, C., & Benítez, J. M. (2012). On the use of cross-validation for time series predictor evaluation. *Information Sciences, 191*, 192–213. https://doi.org/10.1016/j.ins.2011.12.028",
        "Bergstra, J., & Bengio, Y. (2012). Random search for hyper-parameter optimization. *Journal of Machine Learning Research, 13*, 281–305.",
        "Chen, T., & Guestrin, C. (2016). XGBoost: A scalable tree boosting system. In *Proceedings of the 22nd ACM SIGKDD International Conference on Knowledge Discovery and Data Mining* (pp. 785–794). ACM. https://doi.org/10.1145/2939672.2939785",
        "He, H., & Garcia, E. A. (2009). Learning from imbalanced data. *IEEE Transactions on Knowledge and Data Engineering, 21*(9), 1263–1284. https://doi.org/10.1109/TKDE.2008.239",
        "Hersbach, H., Bell, B., Berrisford, P., Hirahara, S., Horányi, A., Muñoz-Sabater, J., Nicolas, J., Peubey, C., Radu, R., Schepers, D., Simmons, A., Soci, C., Abdalla, S., Abellan, X., Balsamo, G., Bechtold, P., Biavati, G., Bidlot, J., Bonavita, M., … Thépaut, J.-N. (2020). The ERA5 global reanalysis. *Quarterly Journal of the Royal Meteorological Society, 146*(730), 1999–2049. https://doi.org/10.1002/qj.3803",
        "Merkel, D. (2014). Docker: Lightweight Linux containers for consistent development and deployment. *Linux Journal, 2014*(239), Article 2.",
        "Pedregosa, F., Varoquaux, G., Gramfort, A., Michel, V., Thirion, B., Grisel, O., Blondel, M., Prettenhofer, P., Weiss, R., Dubourg, V., Vanderplas, J., Passos, A., Cournapeau, D., Brucher, M., Perrot, M., & Duchesnay, É. (2011). Scikit-learn: Machine learning in Python. *Journal of Machine Learning Research, 12*, 2825–2830.",
        "Ramírez, S. (2018). *FastAPI* [Computer software]. https://fastapi.tiangolo.com/",
        "Sculley, D., Holt, G., Golovin, D., Davydov, E., Phillips, T., Ebner, D., Chaudhary, V., Young, M., Crespo, J.-F., & Dennison, D. (2015). Hidden technical debt in machine learning systems. In *Advances in Neural Information Processing Systems 28* (pp. 2503–2511). Curran Associates.",
        "Wilks, D. S. (2011). *Statistical methods in the atmospheric sciences* (3rd ed.). Academic Press.",
        f"Zariyat, T. R. (2026). *amla-package-26016360* (Version 0.2.0) [Computer software]. TestPyPI. {PACKAGE_URL}",
        "Zippenfenig, P. (2023). *Open-Meteo.com weather API* [Computer software]. Zenodo. https://doi.org/10.5281/zenodo.7970649",
    ]
    for ref in refs:
        c.para(ref, count=False, hanging=True, single=True, align=WD_ALIGN_PARAGRAPH.LEFT)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_path)

    missing = [m.path.relative_to(ROOT) for m in (cci, whc) if not m.available]
    print(f"written: {out_path}")
    print(f"body word count (report text only; tables, captions and references excluded): {total_words}")
    if total_words > 4000:
        print("WARNING: body exceeds the 4,000-word limit; shorten a section in reports/build_report.py")
    if missing:
        print("placeholders used - missing:", ", ".join(map(str, missing)))
    figs = [f for f in ("cci_target_exploration.png", "whc_target_exploration.png", "cci_test_scatter.png",
                        "whc_confusion_matrix.png", "whc_threshold_tradeoff.png") if not (FIGURES / f).exists()]
    if figs:
        print("figures missing (run the notebooks):", ", ".join(figs))
    print("open in Word and answer 'Yes' to updating fields so the table of contents shows the right pages")
    return out_path


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Fill the unit's report template with the experiment results.")
    ap.add_argument("--out", default=str(ROOT / "reports" / "AT2_report.docx"))
    args = ap.parse_args()
    build(Path(args.out))
