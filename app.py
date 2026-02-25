import streamlit as st
import pandas as pd
import random
import math
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

st.title("Exam Seating PDF Generator")

uploaded_file = st.file_uploader("Öğrenci listesini Excel olarak yükleyin", type=["xlsx"])

pdf_buffer = None  # PDF buffer global

if uploaded_file:
    df_students = pd.read_excel(uploaded_file)

    if 'Include?' not in df_students.columns:
        df_students['Include?'] = True

    # Sıralama seçenekleri
    sort_columns = [col for col in df_students.columns if col != "Include?"]
    sort_column = st.selectbox("Tabloyu hangi sütuna göre sıralamak istersiniz?", sort_columns, index=0)
    ascending = st.checkbox("Artan sıra", value=True)

    df_sorted = df_students.sort_values(by=sort_column, ascending=ascending)

    st.subheader("Öğrenci Önizleme ve Seçim")
    display_df = df_sorted.drop(columns=['Include?'])
    edited_df = st.data_editor(display_df, num_rows="dynamic")

    id_column = st.text_input("Okul numarası sütun adı", "ID number")
    included_students = df_sorted[df_sorted['Include?'] == True][id_column].dropna().astype(int).astype(str).tolist()

    st.subheader("Sınıf Bilgilerini Girin")
    class_count = st.number_input("Kaç sınıf var?", min_value=1, max_value=20, value=2, step=1)
    classes = {}
    for i in range(class_count):
        cls_name = st.text_input(f"Sınıf {i+1} adı", key=f"class_name_{i}")
        capacity = st.number_input(f"{cls_name} kapasitesi", min_value=1, max_value=500, key=f"capacity_{i}")
        if cls_name:
            classes[cls_name] = capacity

    # Planla butonu
    if st.button("Planla"):
        try:
            random.shuffle(included_students)
            assignments = {}
            index = 0
            for cls, capacity in classes.items():
                assignments[cls] = included_students[index:index+capacity]
                index += capacity

            buffer = BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4)
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
            buffer.seek(0)
            pdf_buffer = buffer  # PDF buffer global olarak sakla
            st.success("PDF başarıyla oluşturuldu!")
        except Exception as e:
            st.error(f"PDF oluşturulamadı: {e}")

    # Eğer PDF başarılı oluşturulduysa indirme butonunu göster
    if pdf_buffer:
        st.download_button(
            label="PDF İndir",
            data=pdf_buffer,
            file_name="exam_seating.pdf",
            mime="application/pdf"
        )