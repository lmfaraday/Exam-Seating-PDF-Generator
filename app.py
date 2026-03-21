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
st.title("Exam Seating Generator")

for key in ['assignments', 'pdf_buffer', 'signature_buffer']:
    if key not in st.session_state:
        st.session_state[key] = None

# ── Step 1: Upload ────────────────────────────────────────────────────────────
with st.expander("Step 1 – Upload Student List", expanded=True):
    uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

if not uploaded_file:
    st.info("Upload a student list to get started.")
    st.stop()

df_students = pd.read_excel(uploaded_file)
if 'Include?' not in df_students.columns:
    df_students['Include?'] = True

data_cols = [col for col in df_students.columns if col != "Include?"]

# ── Step 2: Student Preview ───────────────────────────────────────────────────
with st.expander("Step 2 – Student Preview & Settings", expanded=True):
    c1, c2, c3 = st.columns(3)
    with c1:
        id_column = st.selectbox("Student ID column", data_cols)
    with c2:
        sort_column = st.selectbox("Preview sort by", data_cols)
    with c3:
        ascending = st.checkbox("Ascending order", value=True)

    df_sorted = df_students.sort_values(by=sort_column, ascending=ascending)
    st.dataframe(df_sorted.drop(columns=['Include?']), use_container_width=True, height=260)
    st.caption(f"{len(df_sorted)} students loaded")

included_students = (
    df_sorted[df_sorted['Include?'] == True][id_column]
    .dropna()
    .apply(lambda x: str(int(x)) if isinstance(x, float) else str(x))
    .tolist()
)

# ── Step 3: PDF Column Selection ──────────────────────────────────────────────
with st.expander("Step 3 – PDF Column Selection", expanded=True):
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Seating Plan columns**")
        seating_cols = st.multiselect(
            "Columns to include in Seating Plan",
            data_cols,
            default=[id_column],
            key="seating_cols",
        )
    with c2:
        st.markdown("**Signature Sheet columns**")
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
    if errors:
        for e in errors:
            st.error(e)
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
            included_df = df_sorted[df_sorted['Include?'] == True].copy()
            included_df = included_df.sort_values(by=name_column)
            students_sorted = (
                included_df[id_column]
                .dropna()
                .apply(lambda x: str(int(x)) if isinstance(x, float) else str(x))
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

        st.session_state.assignments = assignments

        # Lookup helper: student_id → list of column values
        df_lookup = df_sorted.copy()
        df_lookup["_id_str"] = (
            df_lookup[id_column]
            .apply(lambda x: str(int(x)) if isinstance(x, float) else str(x))
        )

        def get_values(student_id, columns):
            row = df_lookup[df_lookup["_id_str"] == student_id]
            if row.empty:
                return [""] * len(columns)
            return [str(row.iloc[0][c]) for c in columns]

        # ── Seating PDF ───────────────────────────────────────────────────────
        PAGE_W = A4[0] - 80  # ~515 pt usable width
        n_sc = len(seating_cols)
        seat_w = 35
        data_w = (PAGE_W / 2 - seat_w) / n_sc
        seating_col_widths = ([seat_w] + [data_w] * n_sc) * 2

        seating_buffer = BytesIO()
        doc = SimpleDocTemplate(
            seating_buffer, pagesize=A4,
            leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40,
        )
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
                    if j < len(col2)
                    else [""] * (1 + n_sc)
                )
                table_data.append(left + right)

            style_cmds = [
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2c3e50')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cccccc')),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 9),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('LINEAFTER', (n_sc, 0), (n_sc, -1), 2, colors.HexColor('#2c3e50')),
                ('ROWHEIGHT', (0, 0), (-1, -1), 16),
            ]
            # Alternating row shading
            for row_i in range(1, len(table_data)):
                bg = colors.white if row_i % 2 == 1 else colors.HexColor('#f0f4f8')
                style_cmds.append(('BACKGROUND', (0, row_i), (-1, row_i), bg))

            t = Table(table_data, colWidths=seating_col_widths)
            t.setStyle(TableStyle(style_cmds))
            elements.append(t)

        doc.build(elements)
        seating_buffer.seek(0)
        st.session_state.pdf_buffer = seating_buffer

        # ── Signature PDF ─────────────────────────────────────────────────────
        n_sg = len(sig_cols)
        sig_w = 60
        data_w2 = max((PAGE_W / 2 - seat_w - sig_w) / n_sg, 40)
        sig_col_widths = ([seat_w] + [data_w2] * n_sg + [sig_w]) * 2

        sig_buffer = BytesIO()
        doc = SimpleDocTemplate(
            sig_buffer, pagesize=A4,
            leftMargin=40, rightMargin=40, topMargin=40, bottomMargin=40,
        )
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
                    if j < len(col2)
                    else [""] * (2 + n_sg)
                )
                table_data.append(left + right)

            divider_col = n_sg + 1  # after Signature column of left half
            style_cmds = [
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2c3e50')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cccccc')),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 9),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('LINEAFTER', (divider_col, 0), (divider_col, -1), 2, colors.HexColor('#2c3e50')),
                ('ROWHEIGHT', (0, 0), (-1, -1), 22),
            ]
            for row_i in range(1, len(table_data)):
                bg = colors.white if row_i % 2 == 1 else colors.HexColor('#f0f4f8')
                style_cmds.append(('BACKGROUND', (0, row_i), (-1, row_i), bg))

            t = Table(table_data, colWidths=sig_col_widths)
            t.setStyle(TableStyle(style_cmds))
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
        label="Download Seating Plan & Signature Sheet (ZIP)",
        data=zip_buffer,
        file_name="exam_documents.zip",
        mime="application/zip",
        use_container_width=True,
    )
