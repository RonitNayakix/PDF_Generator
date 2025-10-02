import os
import pandas as pd
import sqlite3
import tempfile
import streamlit as st
from docxtpl import DocxTemplate
import mammoth
from weasyprint import HTML

# ==============================
# Save to SQLite
# ==============================
def save_to_sqlite(df, db_file="uploaded_data.db"):
    conn = sqlite3.connect(db_file)
    df.to_sql("uploaded_data", conn, if_exists="replace", index=False)
    conn.close()

def load_from_sqlite(db_file="uploaded_data.db"):
    conn = sqlite3.connect(db_file)
    df = pd.read_sql("SELECT * FROM uploaded_data", conn)
    conn.close()
    return df

# ==============================
# PDF generation (preserve DOCX formatting)
# ==============================
def generate_pdfs_weasy(template_file, df):
    pdf_paths = []

    for idx, row in df.iterrows():
        context = row.to_dict()

        # Render DOCX template with docxtpl
        doc = DocxTemplate(template_file)
        doc.render(context)
        tmp_docx = os.path.join(tempfile.gettempdir(), f"filled_{idx}.docx")
        doc.save(tmp_docx)

        # Convert DOCX → HTML with Mammoth
        with open(tmp_docx, "rb") as f:
            result = mammoth.convert_to_html(f)
            html_content = result.value  # The generated HTML

        # Convert HTML → PDF with WeasyPrint
        out_pdf = os.path.join(
            tempfile.gettempdir(),
            f"{context.get('OrderID', idx)}_{context.get('Name','Record')}.pdf"
        )
        HTML(string=html_content).write_pdf(out_pdf)

        pdf_paths.append(out_pdf)

    return pdf_paths

# ==============================
# Streamlit UI
# ==============================
st.set_page_config(page_title="Custom PDF Generator", page_icon="📄", layout="centered")
st.title("📄 Styled Invoice PDF Generator")
st.markdown(
    "Upload a DOCX template with placeholders (`{{ColumnName}}`) and a CSV/Excel file. "
    "The app fills in the template for each row and generates PDFs with formatting preserved."
)

template_file = st.file_uploader("Upload DOCX Template", type=["docx"])
data_file = st.file_uploader("Upload CSV/Excel Data", type=["csv", "xlsx"])

if template_file and data_file:
    # Save template to temp file
    tmp_tpl_path = os.path.join(tempfile.gettempdir(), template_file.name)
    with open(tmp_tpl_path, "wb") as f:
        f.write(template_file.read())

    # Load data
    if data_file.name.endswith(".csv"):
        df = pd.read_csv(data_file)
    else:
        df = pd.read_excel(data_file)

    # Save to SQLite
    save_to_sqlite(df)

    st.write("### 📊 Preview of Uploaded Data")
    st.dataframe(df.head())

    if st.button("Generate Styled PDFs"):
        pdf_paths = generate_pdfs_weasy(tmp_tpl_path, df)
        st.success("✅ PDFs generated successfully!")

        for pdf in pdf_paths:
            with open(pdf, "rb") as f:
                st.download_button(
                    label=f"⬇ Download {os.path.basename(pdf)}",
                    data=f,
                    file_name=os.path.basename(pdf),
                    mime="application/pdf"
                )
