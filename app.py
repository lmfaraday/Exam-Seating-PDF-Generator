"""
Exam Seating Generator
======================
Streamlit app for teaching assistants to generate exam seating plans
and student signature sheets from an Excel student list.

Workflow:
  1. Upload an Excel file with student data
  2. Configure columns, classrooms, and seating mode
  3. Generate and download a ZIP with two PDFs:
       - exam_seating.pdf    : seat assignments per classroom
       - signature_sheet.pdf : same layout with a blank Signature column
"""

import math
import os
import random
from io import BytesIO
from zipfile import ZipFile

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)


# ─────────────────────────────────────────────────────────────────────────────
# Fonts
# ─────────────────────────────────────────────────────────────────────────────

# Helvetica (ReportLab built-in) does not support Turkish characters (İ, Ş, Ğ, …).
# DejaVu Sans is bundled in the fonts/ directory alongside this script for full Unicode support.
# System fonts are tried first; the bundled fonts are the guaranteed fallback.

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

_FONT_CANDIDATES = {
    "regular": [
        os.path.join(_SCRIPT_DIR, "fonts", "DejaVuSans.ttf"),     # bundled (always present)
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",         # Linux / Streamlit Cloud
        "/usr/share/fonts/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/Library/Fonts/Arial.ttf",                                # macOS
        "/System/Library/Fonts/Supplemental/Arial.ttf",
    ],
    "bold": [
        os.path.join(_SCRIPT_DIR, "fonts", "DejaVuSans-Bold.ttf"), # bundled
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        "/Library/Fonts/Arial Bold.ttf",
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
    ],
}


def _register_unicode_fonts() -> tuple[str, str]:
    """Register a Unicode TTF font that covers Turkish characters.

    Tries candidates in order; the bundled DejaVu Sans is listed first so it
    is always found even when no system fonts are available.
    """
    reg, bold = "Helvetica", "Helvetica-Bold"

    def _try(path: str, name: str) -> bool:
        try:
            pdfmetrics.registerFont(TTFont(name, path))
            return True
        except Exception:
            return False

    for path in _FONT_CANDIDATES["regular"]:
        if os.path.exists(path) and _try(path, "UniFont"):
            reg = "UniFont"
            break

    for path in _FONT_CANDIDATES["bold"]:
        if os.path.exists(path) and _try(path, "UniFont-Bold"):
            bold = "UniFont-Bold"
            break

    return reg, bold


FONT_REG, FONT_BOLD = _register_unicode_fonts()



# ─────────────────────────────────────────────────────────────────────────────
# Data utilities
# ─────────────────────────────────────────────────────────────────────────────

def to_id_str(value) -> str:
    """Convert a cell value to a clean string student ID.

    Pandas reads integer IDs as floats (e.g. 2022401234.0).
    This strips the decimal point. Returns "" for NaN / None.
    """
    if pd.isna(value):
        return ""
    return str(int(value)) if isinstance(value, float) else str(value)


def first_match_index(options: list, candidates: list) -> int:
    """Return the index of the first candidate found in options, or 0."""
    for c in candidates:
        if c in options:
            return options.index(c)
    return 0


def find_duplicates(values: list) -> set:
    """Return the set of values that appear more than once in the list."""
    seen, dupes = set(), set()
    for v in values:
        (dupes if v in seen else seen).add(v)
    return dupes


# ─────────────────────────────────────────────────────────────────────────────
# Validation
# ─────────────────────────────────────────────────────────────────────────────

def get_hard_errors(
    classes: dict,
    included_students: list,
    seating_cols: list,
    sig_cols: list,
) -> list[str]:
    """Return error messages that must block PDF generation."""
    errors = []

    if not classes:
        errors.append("Enter at least one classroom name.")
    if not included_students:
        errors.append("No students are included.")
    if not seating_cols:
        errors.append("Select at least one column for the Seating Plan.")
    if not sig_cols:
        errors.append("Select at least one column for the Signature Sheet.")

    if classes and included_students:
        total_cap = sum(classes.values())
        n = len(included_students)
        if total_cap < n:
            errors.append(
                f"Total classroom capacity ({total_cap}) is less than the number of "
                f"included students ({n}). Increase capacity or exclude some students."
            )

        dupes = find_duplicates(included_students)
        if dupes:
            errors.append(
                f"Duplicate student IDs found: {', '.join(sorted(dupes))}. "
                "Each student must appear only once."
            )

    return errors


def get_soft_warnings(assignments: dict, classes: dict) -> list[str]:
    """Return warning messages about unusual but non-fatal assignment results."""
    warnings = []
    for cls, students in assignments.items():
        n, cap = len(students), classes[cls]
        if n == 0:
            warnings.append(
                f"**{cls}** received 0 students. "
                "Consider removing this classroom or reducing the classroom count."
            )
        elif n > cap:
            warnings.append(
                f"**{cls}** has {n} students but a capacity of {cap}. "
                "Some students may not have a seat."
            )
    return warnings


# ─────────────────────────────────────────────────────────────────────────────
# Seating assignment
# ─────────────────────────────────────────────────────────────────────────────

def _split_proportionally(students: list, classes: dict) -> dict[str, list]:
    """Distribute students across classrooms proportional to each room's capacity.

    The last classroom absorbs any rounding remainder so no student is lost.
    """
    total_cap = sum(classes.values())
    total_stu = len(students)
    assignments = {}
    cursor = 0

    for i, (cls, cap) in enumerate(classes.items()):
        if i == len(classes) - 1:               # last room gets the remainder
            assignments[cls] = students[cursor:]
        else:
            count = round(total_stu * cap / total_cap)
            assignments[cls] = students[cursor:cursor + count]
            cursor += count

    return assignments


def assign_randomly(students: list, classes: dict, seed: int | None = None) -> dict[str, list]:
    """(Mode A) Shuffle all students, then distribute proportionally by room capacity.

    Args:
        seed: Optional integer seed for reproducible assignments.
    """
    pool = students[:]
    rng = random.Random(seed)
    rng.shuffle(pool)
    return _split_proportionally(pool, classes)


def assign_alphabetically(
    students: list,
    classes: dict,
    df_lookup: pd.DataFrame,
    sort_col: str,
    ascending: bool = True,
) -> dict[str, list]:
    """(Mode B) Sort students by a chosen column, split proportionally, shuffle within rooms.

    Alphabetical grouping gives each room a distinct name range (A–F, G–N, …)
    while seat positions within each room remain randomised.

    Args:
        ascending: Sort A→Z when True, Z→A when False.
    """
    id_to_sort_key = dict(zip(df_lookup["_id_str"], df_lookup[sort_col].astype(str)))
    sorted_students = sorted(
        students,
        key=lambda sid: id_to_sort_key.get(sid, ""),
        reverse=not ascending,
    )
    assignments = _split_proportionally(sorted_students, classes)
    for cls in assignments:
        random.shuffle(assignments[cls])
    return assignments



# ─────────────────────────────────────────────────────────────────────────────
# PDF generation
# ─────────────────────────────────────────────────────────────────────────────

_PAGE_W  = A4[0] - 80   # usable width in points (~515 pt with 40 pt side margins)
_SIG_W   = 65            # fixed width of the "Signature" column
_NAME_W  = 90            # fixed width of the blank "Name" column on signature sheets
_CELL_PAD = 10           # horizontal padding added to every measured cell width


_HEADER_BG   = colors.HexColor("#1B3A5C")   # dark navy – header background
_HEADER_FG   = colors.white                 # header text colour
_ROW_ALT     = colors.HexColor("#EBF4FB")   # light-blue tint for odd data rows
_GRID_COLOR  = colors.HexColor("#B0BEC5")   # soft grey grid lines
_DIVIDER_CLR = colors.HexColor("#1B3A5C")   # colour of the centre divider


def _make_table_style(n_left_cols: int, row_height: int) -> TableStyle:
    """Return a styled table with a coloured header and alternating row tints.

    n_left_cols determines where the thick vertical divider is drawn between
    the left and right student halves on each page.
    """
    return TableStyle([
        # ── Header row ──────────────────────────────────────────────────────
        ("BACKGROUND", (0,  0), (-1,  0), _HEADER_BG),
        ("TEXTCOLOR",  (0,  0), (-1,  0), _HEADER_FG),
        ("FONTNAME",   (0,  0), (-1,  0), FONT_BOLD),
        ("FONTSIZE",   (0,  0), (-1,  0), 9),
        # ── Data rows ───────────────────────────────────────────────────────
        ("FONTNAME",   (0,  1), (-1, -1), FONT_REG),
        ("FONTSIZE",   (0,  1), (-1, -1), 8),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, _ROW_ALT]),
        # ── Layout ──────────────────────────────────────────────────────────
        ("ALIGN",    (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",   (0, 0), (-1, -1), "MIDDLE"),
        ("ROWHEIGHT", (0, 0), (-1, -1), row_height),
        # ── Grid & divider ──────────────────────────────────────────────────
        ("GRID",      (0, 0), (-1, -1), 0.5, _GRID_COLOR),
        ("LINEAFTER", (n_left_cols - 1, 0), (n_left_cols - 1, -1), 2.5, _DIVIDER_CLR),
    ])


def _make_title_style() -> ParagraphStyle:
    return ParagraphStyle("ExamTitle", parent=getSampleStyleSheet()["Title"], fontName=FONT_BOLD)


def _compute_col_widths(
    table_data: list[list],
    n_data_cols: int,
    include_signature: bool,
    n_blank_cols: int = 0,
) -> list[float]:
    """Calculate column widths based on actual cell content.

    For each column position in one half, we measure the widest text that
    appears in any row (header or data) using ReportLab's stringWidth.
    The Seat column is capped narrow; blank and Signature columns get fixed
    widths. All remaining space is distributed to the data columns
    proportionally to their measured natural widths.
    """
    from reportlab.pdfbase.pdfmetrics import stringWidth

    n_fixed_right = n_blank_cols + (1 if include_signature else 0)
    half_ncols = 1 + n_data_cols + n_fixed_right

    # Measure the natural (content-driven) width of each column
    natural = [0.0] * half_ncols
    for row_idx, row in enumerate(table_data):
        font = FONT_BOLD if row_idx == 0 else FONT_REG
        fsize = 9 if row_idx == 0 else 8
        for side in (0, 1):
            for c in range(half_ncols):
                src = side * half_ncols + c
                if src < len(row):
                    w = stringWidth(str(row[src]), font, fsize) + _CELL_PAD
                    natural[c] = max(natural[c], w)

    # Seat column: just needs to fit "Seat" + up to 3-digit numbers
    natural[0] = max(natural[0], 26)

    # Fixed-width trailing columns: blank cols then Signature (students write here)
    if include_signature:
        natural[-1] = _SIG_W
    for k in range(n_blank_cols):
        natural[-(1 + (1 if include_signature else 0) + k)] = _NAME_W

    # Distribute the available half-width:
    #   fixed portion  = Seat + blank cols + Signature (if present)
    #   flexible portion = data cols, scaled proportionally from measured widths
    available_half = _PAGE_W / 2
    n_data_end = -n_fixed_right if n_fixed_right > 0 else len(natural)
    data_natural = natural[1:n_data_end]

    fixed_w = natural[0] + (sum(natural[-n_fixed_right:]) if n_fixed_right > 0 else 0)
    data_available = available_half - fixed_w
    data_total_natural = sum(data_natural)

    if data_total_natural > 0:
        data_widths = [w * data_available / data_total_natural for w in data_natural]
    else:
        data_widths = [data_available / max(n_data_cols, 1)] * n_data_cols

    blank_widths = [_NAME_W] * n_blank_cols
    sig_width = [_SIG_W] if include_signature else []
    one_half = [natural[0]] + data_widths + blank_widths + sig_width
    return one_half * 2


def _cell_to_str(value) -> str:
    """Convert a cell value to a display string.

    Whole-number floats (e.g. 2022401234.0) are rendered without the decimal
    point.  Actual decimals (e.g. 85.5) are kept as-is.
    """
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value)


def _lookup_student_values(df_lookup: pd.DataFrame, student_id: str, columns: list) -> list:
    """Return column values for one student from the lookup dataframe."""
    row = df_lookup[df_lookup["_id_str"] == student_id]
    if row.empty:
        return [""] * len(columns)
    return [_cell_to_str(row.iloc[0][col]) for col in columns]


def _build_table_data(
    students: list,
    data_cols: list,
    df_lookup: pd.DataFrame,
    include_signature: bool,
    blank_cols: list | None = None,
) -> list[list]:
    """Build all rows (header + data) for a two-halves-per-page layout.

    The page is split into left and right halves. Each half shows:
        Seat | data_cols... | [blank_cols...] | [Signature]
    """
    blank_cols = blank_cols or []
    sig = ["Signature"] if include_signature else []
    trailing = blank_cols + sig          # user-defined blank cols, then Signature
    half = math.ceil(len(students) / 2)
    left_half, right_half = students[:half], students[half:]
    n_data = len(data_cols)

    header = ["Seat"] + data_cols + trailing + ["Seat"] + data_cols + trailing
    rows = [header]

    for i in range(half):
        left  = [str(i + 1)]        + _lookup_student_values(df_lookup, left_half[i],  data_cols) + [""] * len(trailing)
        right = ([str(i + 1 + half)] + _lookup_student_values(df_lookup, right_half[i], data_cols) + [""] * len(trailing)
                 if i < len(right_half) else [""] * (1 + n_data + len(trailing)))
        rows.append(left + right)

    return rows


def _build_pdf(
    assignments: dict,
    columns: list,
    df_lookup: pd.DataFrame,
    include_signature: bool,
    blank_cols: list | None = None,
) -> BytesIO:
    """Core PDF builder used by both the seating plan and signature sheet."""
    blank_cols = blank_cols or []
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40)
    title_style = _make_title_style()

    n_left_cols = 1 + len(columns) + len(blank_cols) + (1 if include_signature else 0)
    row_height  = 22 if include_signature else 16

    elements = []
    for i, (classroom, students) in enumerate(assignments.items()):
        if i > 0:
            elements.append(PageBreak())

        heading = f"<b>{classroom}{' – Signature Sheet' if include_signature else ''}</b>"
        elements.append(Paragraph(heading, title_style))
        elements.append(Spacer(1, 10))

        table_data = _build_table_data(students, columns, df_lookup, include_signature, blank_cols)
        # Compute widths from actual content so every column fits its data
        col_widths = _compute_col_widths(table_data, len(columns), include_signature, len(blank_cols))
        table = Table(table_data, colWidths=col_widths)
        table.setStyle(_make_table_style(n_left_cols, row_height))
        elements.append(table)

    doc.build(elements)
    buf.seek(0)
    return buf


def build_seating_pdf(assignments: dict, columns: list, df_lookup: pd.DataFrame) -> BytesIO:
    """Generate the seating plan PDF (seat numbers + selected columns)."""
    return _build_pdf(assignments, columns, df_lookup, include_signature=False)


def build_signature_pdf(assignments: dict, columns: list, df_lookup: pd.DataFrame, blank_cols: list | None = None) -> BytesIO:
    """Generate the signature sheet PDF (data columns + user-defined blank columns + Signature)."""
    return _build_pdf(assignments, columns, df_lookup, include_signature=True, blank_cols=blank_cols)


def build_zip(seating_buf: BytesIO, signature_buf: BytesIO) -> BytesIO:
    """Pack both PDFs into a single ZIP archive for download."""
    zip_buf = BytesIO()
    with ZipFile(zip_buf, "w") as zf:
        zf.writestr("exam_seating.pdf",    seating_buf.getvalue())
        zf.writestr("signature_sheet.pdf", signature_buf.getvalue())
    zip_buf.seek(0)
    return zip_buf


# ─────────────────────────────────────────────────────────────────────────────
# Streamlit UI
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Exam Seating Generator", layout="wide", page_icon="🪑")

st.markdown("""
<style>
/* ── Page background ─────────────────────────────────────────────────────── */
[data-testid="stAppViewContainer"] {
    background: linear-gradient(160deg, #0f1117 0%, #1a1f2e 60%, #0f1117 100%);
}
[data-testid="stHeader"] { background: transparent; }

/* ── Title ───────────────────────────────────────────────────────────────── */
h1 {
    background: linear-gradient(90deg, #63b3ed, #90cdf4, #4299e1);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    font-size: 2.2rem !important; font-weight: 800 !important;
    letter-spacing: -0.5px; margin-bottom: 0.1rem !important;
}
.subtitle {
    color: #718096; font-size: 0.95rem; margin-bottom: 1.5rem;
}

/* ── Step expanders ──────────────────────────────────────────────────────── */
[data-testid="stExpander"] > details {
    border: 1px solid rgba(99,179,237,0.18);
    border-radius: 12px;
    margin-bottom: 0.8rem;
    background: rgba(26,32,44,0.55);
    backdrop-filter: blur(6px);
    box-shadow: 0 2px 12px rgba(0,0,0,0.3);
    transition: border-color 0.2s;
}
[data-testid="stExpander"] > details:hover {
    border-color: rgba(99,179,237,0.35);
}
[data-testid="stExpander"] > details > summary {
    font-size: 1rem; font-weight: 700;
    color: #90cdf4; padding: 0.75rem 1.2rem;
    letter-spacing: 0.01em;
}
[data-testid="stExpander"] > details > summary:hover { color: #bee3f8; }
[data-testid="stExpander"] > details > summary::marker { color: #4299e1; }

/* ── Section labels inside expanders ─────────────────────────────────────── */
.section-label {
    font-size: 0.72rem; font-weight: 700; letter-spacing: 0.1em;
    text-transform: uppercase; color: #4299e1; margin-bottom: 0.3rem;
    padding-left: 2px;
}

/* ── Generate button ─────────────────────────────────────────────────────── */
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #2b6cb0 0%, #3182ce 50%, #4299e1 100%);
    color: white; border: none; border-radius: 10px;
    font-size: 1.1rem; font-weight: 700;
    padding: 0.75rem 1rem; transition: all 0.25s ease;
    box-shadow: 0 2px 10px rgba(49,130,206,0.3);
}
.stButton > button[kind="primary"]:hover {
    background: linear-gradient(135deg, #3182ce 0%, #4299e1 100%);
    box-shadow: 0 6px 24px rgba(49,130,206,0.55);
    transform: translateY(-2px);
}
.stButton > button[kind="primary"]:active { transform: translateY(0); }

/* ── Download button ─────────────────────────────────────────────────────── */
.stDownloadButton > button {
    background: linear-gradient(135deg, #1c4532, #276749, #38a169) !important;
    color: white !important; border: none !important; border-radius: 10px !important;
    font-size: 1.05rem !important; font-weight: 700 !important;
    transition: all 0.25s ease !important;
    box-shadow: 0 2px 10px rgba(56,161,105,0.25) !important;
}
.stDownloadButton > button:hover {
    background: linear-gradient(135deg, #276749, #48bb78) !important;
    box-shadow: 0 6px 24px rgba(56,161,105,0.45) !important;
    transform: translateY(-2px) !important;
}

/* ── Metric cards ────────────────────────────────────────────────────────── */
[data-testid="stMetric"] {
    background: linear-gradient(135deg, rgba(49,130,206,0.1), rgba(66,153,225,0.06));
    border: 1px solid rgba(99,179,237,0.25);
    border-radius: 10px; padding: 0.75rem 1.2rem;
    box-shadow: 0 1px 6px rgba(0,0,0,0.2);
}
[data-testid="stMetricLabel"] { color: #718096 !important; font-size: 0.8rem !important; }
[data-testid="stMetricValue"] { color: #90cdf4 !important; font-weight: 800 !important; font-size: 1.6rem !important; }

/* ── Alerts ──────────────────────────────────────────────────────────────── */
[data-testid="stAlert"] { border-radius: 10px; }

/* ── Divider ─────────────────────────────────────────────────────────────── */
hr { border-color: rgba(99,179,237,0.15) !important; }

/* ── Number inputs / text inputs ─────────────────────────────────────────── */
[data-testid="stNumberInput"] input,
[data-testid="stTextInput"] input {
    border-radius: 8px !important;
}

/* ── Multiselect tags ────────────────────────────────────────────────────── */
[data-testid="stMultiSelect"] span[data-baseweb="tag"] {
    background-color: rgba(49,130,206,0.25) !important;
    border: 1px solid rgba(99,179,237,0.4) !important;
    border-radius: 6px !important;
}
</style>
""", unsafe_allow_html=True)

st.title("Exam Seating Generator")
st.markdown('<p class="subtitle">Generate randomised seating plans and signature sheets for exams.</p>', unsafe_allow_html=True)

for key in ["pdf_buffer", "signature_buffer"]:
    if key not in st.session_state:
        st.session_state[key] = None

# ── Step 1: Upload ────────────────────────────────────────────────────────────
with st.expander("Step 1 – Upload Student List", expanded=True):
    uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

if not uploaded_file:
    st.info("Upload a student list to get started.")
    st.stop()

df_raw = pd.read_excel(uploaded_file)
if "Include?" not in df_raw.columns:
    df_raw.insert(0, "Include?", True)

raw_cols = [col for col in df_raw.columns if col != "Include?"]

# ── Step 2: Students ──────────────────────────────────────────────────────────
with st.expander("Step 2 – Student Preview & Settings", expanded=True):

    # Always create a "Name - Surname" column; let the user choose the source columns.
    c1, c2 = st.columns(2)
    name_col    = c1.selectbox("First name column", raw_cols, index=first_match_index(raw_cols, ["First name", "first name", "Ad", "Name"]))
    surname_col = c2.selectbox("Last name column",  raw_cols, index=first_match_index(raw_cols, ["Last name", "last name", "Soyad", "Surname"]))

    df_work = df_raw.copy()
    df_work.insert(1, "Name - Surname", df_work[name_col].astype(str) + " " + df_work[surname_col].astype(str))
    all_cols = [col for col in df_work.columns if col != "Include?"]

    c1, c2, c3 = st.columns(3)
    id_col    = c1.selectbox("Student ID column", all_cols, index=first_match_index(all_cols, ["ID number", "id", "ID"]))
    sort_col  = c2.selectbox("Sort preview by",   all_cols)
    ascending = c3.checkbox("Ascending order", value=True)

    df_sorted = df_work.sort_values(by=sort_col, ascending=ascending).reset_index(drop=True)

    # Columns starting with "group" are not relevant for seating; hide them.
    hidden_cols = {c for c in all_cols if c.lower().startswith("group")}

    st.markdown('<p class="section-label">Select students to include</p>', unsafe_allow_html=True)
    edited_df = st.data_editor(
        df_sorted[["Include?"] + all_cols],
        column_config={
            "Include?": st.column_config.CheckboxColumn("Include?", default=True, width="small"),
            **{c: None for c in hidden_cols},   # hide group columns
        },
        disabled=all_cols,          # only the Include? checkbox is editable
        use_container_width=True,
        height=300,
        key=f"editor_{uploaded_file.name}",
    )

    n_included = int(edited_df["Include?"].sum())
    n_excluded = len(edited_df) - n_included
    ca, cb = st.columns(2)
    ca.caption(f"✅ {n_included} students included")
    if n_excluded:
        cb.caption(f"⛔ {n_excluded} students excluded")

# Flat list of included student IDs (strings) used by assignment and validation
included_students = (
    edited_df[edited_df["Include?"] == True][id_col]
    .dropna()
    .apply(to_id_str)
    .tolist()
)

# Lookup dataframe: one row per student, indexed by "_id_str" for fast PDF row population
df_lookup = edited_df.copy()
df_lookup["_id_str"] = df_lookup[id_col].apply(to_id_str)

# ── Step 3: PDF Columns ───────────────────────────────────────────────────────
if "sig_blank_cols" not in st.session_state:
    st.session_state.sig_blank_cols = []

with st.expander("Step 3 – PDF Column Selection", expanded=True):
    st.caption("Select columns in the order you want them to appear in the PDF.")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<p class="section-label">Seating Plan columns</p>', unsafe_allow_html=True)
        seating_cols = st.multiselect("Columns in Seating Plan", all_cols, default=[id_col], key="seating_cols")
    with c2:
        st.markdown('<p class="section-label">Signature Sheet columns</p>', unsafe_allow_html=True)
        sig_cols = st.multiselect("Columns in Signature Sheet", all_cols, default=[id_col], key="sig_cols")

    st.markdown('<p class="section-label">Blank columns on Signature Sheet</p>', unsafe_allow_html=True)
    st.caption("These columns appear empty in the signature sheet so students can fill them in by hand.")
    ba, bb = st.columns([4, 1])
    new_blank = ba.text_input("Blank column name", key="new_blank_col_input", label_visibility="collapsed", placeholder="e.g. Name, Department…")
    if bb.button("Add", key="add_blank_col", use_container_width=True):
        name = new_blank.strip()
        if name:
            st.session_state.sig_blank_cols.append(name)
            st.rerun()
    for idx, col_name in enumerate(st.session_state.sig_blank_cols):
        r1, r2 = st.columns([5, 1])
        r1.text(f"  {idx + 1}. {col_name}")
        if r2.button("Remove", key=f"remove_blank_{idx}", use_container_width=True):
            st.session_state.sig_blank_cols.pop(idx)
            st.rerun()

# ── Step 4: Classrooms ────────────────────────────────────────────────────────
with st.expander("Step 4 – Classroom Configuration", expanded=True):
    n_classes = int(st.number_input("Number of classrooms", min_value=1, max_value=20, value=2, step=1))
    classes = {}
    grid = st.columns(min(n_classes, 4))
    for i in range(n_classes):
        with grid[i % len(grid)]:
            cls_name = st.text_input(f"Class {i + 1} name", key=f"cls_name_{i}")
            cls_cap  = st.number_input("Capacity", min_value=1, max_value=500, value=30, key=f"cls_cap_{i}")
            if cls_name:
                classes[cls_name] = cls_cap

# ── Step 5: Seating Mode ──────────────────────────────────────────────────────
_SEATING_MODES = {
    "random":       "Completely Random",
    "alphabetical": "Alphabetical Split → Random within rooms",
}

with st.expander("Step 5 – Seating Mode", expanded=True):
    mode_key = st.radio(
        "How should students be assigned to classrooms?",
        list(_SEATING_MODES.keys()),
        format_func=lambda k: _SEATING_MODES[k],
        horizontal=True,
    )

    # ── Random options ────────────────────────────────────────────────────────
    random_seed: int | None = None
    if mode_key == "random":
        use_seed = st.checkbox("Use fixed seed for reproducible results", value=False)
        if use_seed:
            random_seed = int(st.number_input("Random seed", min_value=0, max_value=999999, value=42, step=1))

    # ── Alphabetical options ──────────────────────────────────────────────────
    alpha_sort_col: str | None = None
    alpha_ascending: bool = True
    if mode_key == "alphabetical":
        bc1, bc2 = st.columns(2)
        alpha_sort_col = bc1.selectbox("Sort students by", all_cols, index=first_match_index(all_cols, ["Last name", "Name - Surname", "First name"]))
        alpha_ascending = bc2.checkbox("Ascending (A → Z)", value=True)

# ── Generate ──────────────────────────────────────────────────────────────────
st.divider()
if st.button("Generate Seating", type="primary", use_container_width=True):

    errors = get_hard_errors(classes, included_students, seating_cols, sig_cols)
    for msg in errors:
        st.error(msg)
    if errors:
        st.stop()

    try:
        if mode_key == "random":
            assignments = assign_randomly(included_students, classes, seed=random_seed)
        else:  # alphabetical
            assignments = assign_alphabetically(
                included_students, classes, df_lookup, alpha_sort_col, ascending=alpha_ascending
            )

        for msg in get_soft_warnings(assignments, classes):
            st.warning(msg)

        st.session_state.pdf_buffer       = build_seating_pdf(assignments, seating_cols, df_lookup)
        st.session_state.signature_buffer = build_signature_pdf(assignments, sig_cols, df_lookup, blank_cols=st.session_state.sig_blank_cols)

        st.success("Seating plan and signature sheet generated successfully!")
        metrics = st.columns(len(assignments))
        for (cls, students), col in zip(assignments.items(), metrics):
            col.metric(cls, f"{len(students)} students")

    except Exception as e:
        st.error(f"Failed to generate PDFs: {e}")

# ── Download ──────────────────────────────────────────────────────────────────
if st.session_state.pdf_buffer and st.session_state.signature_buffer:
    st.divider()
    st.download_button(
        label="⬇  Download Seating Plan & Signature Sheet (ZIP)",
        data=build_zip(st.session_state.pdf_buffer, st.session_state.signature_buffer),
        file_name="exam_documents.zip",
        mime="application/zip",
        use_container_width=True,
    )
