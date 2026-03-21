import streamlit as st
import pandas as pd
import random
import math
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from zipfile import ZipFile

st.set_page_config(page_title="Exam Seating Generator", layout="wide")

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* Page title */
h1 { color: #63b3ed; font-size: 2rem !important; margin-bottom: 0.25rem !important; }

/* Expander border + rounded corners */
[data-testid="stExpander"] > details {
    border: 1px solid rgba(99,179,237,0.25);
    border-radius: 10px;
    margin-bottom: 0.75rem;
    background: rgba(26,32,44,0.3);
}
[data-testid="stExpander"] > details > summary {
    font-size: 1rem;
    font-weight: 600;
    color: #90cdf4;
    padding: 0.6rem 1rem;
}
[data-testid="stExpander"] > details > summary:hover { color: #bee3f8; }

/* Section headers inside expanders */
.section-label {
    font-size: 0.78rem;
    font-weight: 600;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: #63b3ed;
    margin-bottom: 0.25rem;
}

/* Primary generate button */
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #2b6cb0, #3182ce);
    color: white;
    border: none;
    border-radius: 8px;
    font-size: 1.1rem;
    font-weight: 700;
    padding: 0.7rem 1rem;
    transition: all 0.2s ease;
}
.stButton > button[kind="primary"]:hover {
    background: linear-gradient(135deg, #3182ce, #4299e1);
    box-shadow: 0 4px 18px rgba(49,130,206,0.45);
    transform: translateY(-1px);
}

/* Download button */
.stDownloadButton > button {
    background: linear-gradient(135deg, #276749, #38a169) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-size: 1.05rem !important;
    font-weight: 600 !important;
    transition: all 0.2s ease !important;
}
.stDownloadButton > button:hover {
    background: linear-gradient(135deg, #38a169, #48bb78) !important;
    box-shadow: 0 4px 18px rgba(56,161,105,0.4) !important;
    transform: translateY(-1px) !important;
}

/* Metric cards */
[data-testid="stMetric"] {
    background: rgba(49,130,206,0.12);
    border: 1px solid rgba(99,179,237,0.3);
    border-radius: 8px;
    padding: 0.6rem 1rem;
}
[data-testid="stMetricValue"] { color: #90cdf4 !important; font-weight: 700; }

/* Success / info alerts */
[data-testid="stAlert"] { border-radius: 8px; }

/* Divider */
hr { border-color: rgba(99,179,237,0.2) !important; }
</style>
""", unsafe_allow_html=True)

st.title("Exam Seating Generator")

for key in ['pdf_buffer', 'signature_buffer']:
    if key not in st.session_state:
        st.session_state[key] = None

# ── Step 1: Upload ────────────────────────────────────────────────────────────
with st.expander("Step 1 – Upload Student List", expanded=True):
    uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

if not uploaded_file:
    st.info("Upload a student list to get started.")
    st.stop()

df_raw = pd.read_excel(uploaded_file)
if 'Include?' not in df_raw.columns:
    df_raw.insert(0, 'Include?', True)

raw_data_cols = [col for col in df_raw.columns if col != "Include?"]

# ── Step 2: Student Preview & Settings ───────────────────────────────────────
with st.expander("Step 2 – Student Preview & Settings", expanded=True):

    # --- Combined column ---
    combine_on = st.checkbox("Create a combined column (e.g. Full Name)")
    combined_col_name = None
    if combine_on:
        cc1, cc2, cc3, cc4 = st.columns([2, 2, 2, 1])
        col_a = cc1.selectbox("First column", raw_data_cols, key="ca")
        col_b = cc2.selectbox("Second column", raw_data_cols,
                               index=min(1, len(raw_data_cols)-1), key="cb")
        combined_col_name = cc3.text_input("New column name", "Full Name")
        separator = cc4.text_input("Separator", " ")

    # Apply combined column to dataframe
    df_work = df_raw.copy()
    if combine_on and combined_col_name:
        df_work.insert(
            1, combined_col_name,
            df_work[col_a].astype(str) + separator + df_work[col_b].astype(str)
        )

    data_cols = [col for col in df_work.columns if col != "Include?"]

    # --- Sort & ID column ---
    sc1, sc2, sc3 = st.columns(3)
    with sc1:
        id_column = st.selectbox("Student ID column", data_cols,
                                  index=data_cols.index("ID number") if "ID number" in data_cols else 0)
    with sc2:
        sort_column = st.selectbox("Sort preview by", data_cols)
    with sc3:
        ascending = st.checkbox("Ascending order", value=True)

    df_sorted = df_work.sort_values(by=sort_column, ascending=ascending).reset_index(drop=True)

    # --- Editable student table with Include? checkbox ---
    st.markdown('<p class="section-label">Select students to include</p>', unsafe_allow_html=True)
    display_order = ['Include?'] + data_cols
    edited_df = st.data_editor(
        df_sorted[display_order],
        column_config={
            "Include?": st.column_config.CheckboxColumn("Include?", default=True, width="small")
        },
        disabled=data_cols,
        use_container_width=True,
        height=300,
        key=f"editor_{uploaded_file.name}",
    )

    included_count = int(edited_df['Include?'].sum())
    excluded_count = len(edited_df) - included_count
    ic1, ic2 = st.columns(2)
    ic1.caption(f"✅ {included_count} students included")
    if excluded_count:
        ic2.caption(f"⛔ {excluded_count} students excluded")

included_students = (
    edited_df[edited_df['Include?'] == True][id_column]
    .dropna()
    .apply(lambda x: str(int(x)) if isinstance(x, float) and not pd.isna(x) else ("" if pd.isna(x) else str(x)))
    .tolist()
)

# Build a lookup df aligned with edited_df
df_lookup = edited_df.copy()
df_lookup["_id_str"] = (
    df_lookup[id_column]
    .apply(lambda x: str(int(x)) if isinstance(x, float) and not pd.isna(x) else ("" if pd.isna(x) else str(x)))
)

# ── Step 3: PDF Column Selection ──────────────────────────────────────────────
with st.expander("Step 3 – PDF Column Selection", expanded=True):
    pc1, pc2 = st.columns(2)
    with pc1:
        st.markdown('<p class="section-label">Seating Plan columns</p>', unsafe_allow_html=True)
        seating_cols = st.multiselect(
            "Columns to include in Seating Plan",
            data_cols,
            default=[id_column],
            key="seating_cols",
        )
    with pc2:
        st.markdown('<p class="section-label">Signature Sheet columns</p>', unsafe_allow_html=True)
        sig_cols = st.multiselect(
            "Columns to include in Signature Sheet",
            data_cols,
            default=[id_column],
            key="sig_cols",
        )

# ── Step 4: Classroom Configuration ──────────────────────────────────────────
with st.expander("Step 4 – Classroom Configuration", expanded=True):
    class_count = st.number_input("Number of classrooms", min_value=1, max_value=20, value=2, step=1)
    classes = {}
    grid_cols = st.columns(min(int(class_count), 4))
    for i in range(int(class_count)):
        with grid_cols[i % len(grid_cols)]:
            cls_name = st.text_input(f"Class {i+1} name", key=f"class_name_{i}")
            capacity = st.number_input("Capacity", min_value=1, max_value=500, value=30, key=f"capacity_{i}")
            if cls_name:
                classes[cls_name] = capacity

# ── Step 5: Seating Mode ──────────────────────────────────────────────────────
with st.expander("Step 5 – Seating Mode", expanded=True):
    seating_mode = st.radio(
        "Seating mode",
        ["Completely Random", "Alphabetically Split, then Random"],
        horizontal=True,
    )
    if seating_mode == "Alphabetically Split, then Random":
        name_column = st.selectbox("Sort students alphabetically by", data_cols)

# ── Generate ──────────────────────────────────────────────────────────────────
st.divider()
generate = st.button("Generate Seating", type="primary", use_container_width=True)

if generate:
    errors = []
    if not classes:
        errors.append("Enter at least one classroom name.")
    if not seating_cols:
        errors.append("Select at least one column for the Seating Plan.")
    if not sig_cols:
        errors.append("Select at least one column for the Signature Sheet.")
    if not included_students:
        errors.append("No students are included.")
    for err in errors:
        st.error(err)
    if errors:
        st.stop()

    try:
        total_capacity = sum(classes.values())
        assignments = {}
        index = 0

        if seating_mode == "Completely Random":
            students = included_students[:]
            random.shuffle(students)
            total_students = len(students)
            for i, (cls, cap) in enumerate(classes.items()):
                if i == len(classes) - 1:
                    assignments[cls] = students[index:]
                else:
                    num = round(total_students * cap / total_capacity)
                    assignments[cls] = students[index:index + num]
                    index += num
        else:
            included_df = edited_df[edited_df['Include?'] == True].copy()
            included_df = included_df.sort_values(by=name_column)
            students_sorted = (
                included_df[id_column]
                .dropna()
                .apply(lambda x: str(int(x)) if isinstance(x, float) and not pd.isna(x) else ("" if pd.isna(x) else str(x)))
                .tolist()
            )
            total_students = len(students_sorted)
            for i, (cls, cap) in enumerate(classes.items()):
                if i == len(classes) - 1:
                    group = students_sorted[index:]
                else:
                    num = round(total_students * cap / total_capacity)
                    group = students_sorted[index:index + num]
                    index += num
                random.shuffle(group)
                assignments[cls] = group

        def get_values(student_id, columns):
            row = df_lookup[df_lookup["_id_str"] == student_id]
            if row.empty:
                return [""] * len(columns)
            return [str(row.iloc[0][c]) for c in columns]

        PAGE_W = A4[0] - 80  # ~515 pt usable width
        HDR_COLOR = colors.HexColor('#1a365d')
        ROW_ALT   = colors.HexColor('#ebf4ff')
        DIVIDER   = colors.HexColor('#2b6cb0')

        def base_style(n_left_cols, n_total_cols, n_rows):
            cmds = [
                ('BACKGROUND', (0, 0), (-1, 0), HDR_COLOR),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#d0dde8')),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 9),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                # Bold divider between the two halves
                ('LINEAFTER', (n_left_cols - 1, 0), (n_left_cols - 1, -1), 2.5, DIVIDER),
            ]
            for row_i in range(1, n_rows):
                if row_i % 2 == 0:
                    cmds.append(('BACKGROUND', (0, row_i), (-1, row_i), ROW_ALT))
            return cmds

        # ── Seating PDF ───────────────────────────────────────────────────────
        n_sc = len(seating_cols)
        seat_w = 32
        data_w = (PAGE_W / 2 - seat_w) / n_sc
        seating_cw = ([seat_w] + [data_w] * n_sc) * 2

        seating_buffer = BytesIO()
        doc = SimpleDocTemplate(seating_buffer, pagesize=A4,
                                leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40)
        elements = []
        styles = getSampleStyleSheet()

        for idx_cls, (cls, student_list) in enumerate(assignments.items()):
            if idx_cls > 0:
                elements.append(PageBreak())
            elements.append(Paragraph(f"<b>{cls}</b>", styles['Title']))
            elements.append(Spacer(1, 10))

            half = math.ceil(len(student_list) / 2)
            col1, col2 = student_list[:half], student_list[half:]
            header = ["Seat"] + list(seating_cols) + ["Seat"] + list(seating_cols)
            table_data = [header]
            for j in range(half):
                left = [str(j + 1)] + get_values(col1[j], seating_cols)
                right = (
                    [str(j + 1 + half)] + get_values(col2[j], seating_cols)
                    if j < len(col2) else [""] * (1 + n_sc)
                )
                table_data.append(left + right)

            cmds = base_style(1 + n_sc, len(header), len(table_data))
            cmds.append(('ROWHEIGHT', (0, 0), (-1, -1), 16))
            t = Table(table_data, colWidths=seating_cw)
            t.setStyle(TableStyle(cmds))
            elements.append(t)

        doc.build(elements)
        seating_buffer.seek(0)
        st.session_state.pdf_buffer = seating_buffer

        # ── Signature PDF ─────────────────────────────────────────────────────
        n_sg = len(sig_cols)
        sig_w = 62
        data_w2 = max((PAGE_W / 2 - seat_w - sig_w) / n_sg, 38)
        sig_cw = ([seat_w] + [data_w2] * n_sg + [sig_w]) * 2

        sig_buffer = BytesIO()
        doc = SimpleDocTemplate(sig_buffer, pagesize=A4,
                                leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40)
        elements = []

        for idx_cls, (cls, student_list) in enumerate(assignments.items()):
            if idx_cls > 0:
                elements.append(PageBreak())
            elements.append(Paragraph(f"<b>{cls} – Signature Sheet</b>", styles['Title']))
            elements.append(Spacer(1, 10))

            half = math.ceil(len(student_list) / 2)
            col1, col2 = student_list[:half], student_list[half:]
            header = ["Seat"] + list(sig_cols) + ["Signature"] + ["Seat"] + list(sig_cols) + ["Signature"]
            table_data = [header]
            for j in range(half):
                left = [str(j + 1)] + get_values(col1[j], sig_cols) + [""]
                right = (
                    [str(j + 1 + half)] + get_values(col2[j], sig_cols) + [""]
                    if j < len(col2) else [""] * (2 + n_sg)
                )
                table_data.append(left + right)

            cmds = base_style(1 + n_sg + 1, len(header), len(table_data))
            cmds.append(('ROWHEIGHT', (0, 0), (-1, -1), 22))
            t = Table(table_data, colWidths=sig_cw)
            t.setStyle(TableStyle(cmds))
            elements.append(t)

        doc.build(elements)
        sig_buffer.seek(0)
        st.session_state.signature_buffer = sig_buffer

        st.success("Seating plan and signature sheet generated successfully!")
        metric_cols = st.columns(len(assignments))
        for (cls, lst), col in zip(assignments.items(), metric_cols):
            col.metric(cls, f"{len(lst)} students")

    except Exception as e:
        st.error(f"Failed to generate PDFs: {e}")

# ── Download ──────────────────────────────────────────────────────────────────
if st.session_state.pdf_buffer and st.session_state.signature_buffer:
    st.divider()
    zip_buffer = BytesIO()
    with ZipFile(zip_buffer, 'w') as zf:
        zf.writestr("exam_seating.pdf", st.session_state.pdf_buffer.getvalue())
        zf.writestr("signature_sheet.pdf", st.session_state.signature_buffer.getvalue())
    zip_buffer.seek(0)
    st.download_button(
        label="⬇ Download Seating Plan & Signature Sheet (ZIP)",
        data=zip_buffer,
        file_name="exam_documents.zip",
        mime="application/zip",
        use_container_width=True,
    )
