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

st.title("Exam Seating PDF Generator")

# Initialize session state for assignments
if 'assignments' not in st.session_state:
    st.session_state.assignments = None
if 'pdf_buffer' not in st.session_state:
    st.session_state.pdf_buffer = None
if 'signature_buffer' not in st.session_state:
    st.session_state.signature_buffer = None

uploaded_file = st.file_uploader("Upload the student list as an Excel file", type=["xlsx"])

if uploaded_file:
    df_students = pd.read_excel(uploaded_file)

    if 'Include?' not in df_students.columns:
        df_students['Include?'] = True

    # Sorting options
    sort_columns = [col for col in df_students.columns if col != "Include?"]
    sort_column = st.selectbox("Which column would you like to sort the table by?", sort_columns, index=0)
    ascending = st.checkbox("Ascending order", value=True)

    df_sorted = df_students.sort_values(by=sort_column, ascending=ascending)

    st.subheader("Student Preview and Selection")
    display_df = df_sorted.drop(columns=['Include?'])
    edited_df = st.data_editor(display_df, num_rows="dynamic")

    id_column = st.text_input("Name of the student ID column", "ID number")
    included_students = df_sorted[df_sorted['Include?'] == True][id_column].dropna().astype(int).astype(str).tolist()

    st.subheader("Enter Classroom Information")
    class_count = st.number_input("Number of classrooms", min_value=1, max_value=20, value=2, step=1)
    classes = {}
    for i in range(class_count):
        cls_name = st.text_input(f"Class {i+1} name", key=f"class_name_{i}")
        capacity = st.number_input(f"{cls_name} capacity", min_value=1, max_value=500, key=f"capacity_{i}")
        if cls_name:
            classes[cls_name] = capacity

    # Generate Seating - creates both PDFs at once
    if st.button("Generate Seating"):
        try:
            random.shuffle(included_students)
            
            # Calculate proportional distribution
            total_capacity = sum(classes.values())
            total_students = len(included_students)
            
            assignments = {}
            index = 0
            
            # Distribute students proportionally based on classroom capacity
            for i, (cls, capacity) in enumerate(classes.items()):
                if i == len(classes) - 1:  # Last classroom gets remaining students
                    assignments[cls] = included_students[index:]
                else:
                    # Calculate proportional share
                    proportion = capacity / total_capacity
                    num_students = round(total_students * proportion)
                    assignments[cls] = included_students[index:index+num_students]
                    index += num_students

            # Store assignments in session state
            st.session_state.assignments = assignments

            # Generate Seating PDF
            seating_buffer = BytesIO()
            doc = SimpleDocTemplate(seating_buffer, pagesize=A4)
            elements = []
            styles = getSampleStyleSheet()
            first_class = True

            for cls, student_list in assignments.items():
                if not first_class:
                    elements.append(PageBreak())
                first_class = False

                elements.append(Paragraph(f"<b>{cls}</b>", styles['Title']))
                elements.append(Spacer(1, 12))

                half = math.ceil(len(student_list)/2)
                col1 = student_list[:half]
                col2 = student_list[half:]

                table_data = [["Seat", "ID", "Seat", "ID"]]
                for i in range(half):
                    row = [str(i+1), col1[i]]
                    if i < len(col2):
                        row += [str(i+1+half), col2[i]]
                    else:
                        row += ["", ""]
                    table_data.append(row)

                table = Table(table_data, colWidths=[40, 80, 40, 80])
                table_style = TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.grey),
                    ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
                    ('ALIGN',(0,0),(-1,-1),'CENTER'),
                    ('GRID', (0,0), (-1,-1), 1, colors.black),
                    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('FONTNAME', (0,1), (-1,-1), 'Helvetica')
                ])
                table.setStyle(table_style)
                elements.append(table)
                elements.append(Spacer(1, 24))

            doc.build(elements)
            seating_buffer.seek(0)
            st.session_state.pdf_buffer = seating_buffer

            # Generate Signature Sheet PDF
            signature_buffer = BytesIO()
            doc = SimpleDocTemplate(signature_buffer, pagesize=A4)
            elements = []
            first_class = True

            for cls, student_list in assignments.items():
                if not first_class:
                    elements.append(PageBreak())
                first_class = False

                elements.append(Paragraph(f"<b>{cls} - Signature Sheet</b>", styles['Title']))
                elements.append(Spacer(1, 12))

                half = math.ceil(len(student_list)/2)
                col1 = student_list[:half]
                col2 = student_list[half:]

                table_data = [["Seat", "ID", "Signature", "Seat", "ID", "Signature"]]
                for i in range(half):
                    row = [str(i+1), col1[i], ""]
                    if i < len(col2):
                        row += [str(i+1+half), col2[i], ""]
                    else:
                        row += ["", "", ""]
                    table_data.append(row)

                table = Table(table_data, colWidths=[40, 80, 80, 40, 80, 80])
                table_style = TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.grey),
                    ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
                    ('ALIGN',(0,0),(-1,-1),'CENTER'),
                    ('GRID', (0,0), (-1,-1), 1, colors.black),
                    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('FONTNAME', (0,1), (-1,-1), 'Helvetica')
                ])
                table.setStyle(table_style)
                elements.append(table)
                elements.append(Spacer(1, 24))

            doc.build(elements)
            signature_buffer.seek(0)
            st.session_state.signature_buffer = signature_buffer

            st.success("Seating plan and signature sheet successfully generated!")
            
            # Show distribution info
            st.info("Distribution:")
            for cls, student_list in assignments.items():
                st.write(f"{cls}: {len(student_list)} students")

        except Exception as e:
            st.error(f"Failed to generate PDFs: {e}")

    # Single Download button for both PDFs in a ZIP file
    if st.session_state.pdf_buffer and st.session_state.signature_buffer:
        # Create ZIP file containing both PDFs
        zip_buffer = BytesIO()
        with ZipFile(zip_buffer, 'w') as zip_file:
            zip_file.writestr("exam_seating.pdf", st.session_state.pdf_buffer.getvalue())
            zip_file.writestr("signature_sheet.pdf", st.session_state.signature_buffer.getvalue())
        zip_buffer.seek(0)
        
        st.download_button(
            label="Download Seating Plan & Signature Sheet",
            data=zip_buffer,
            file_name="exam_documents.zip",
            mime="application/zip"
        )
