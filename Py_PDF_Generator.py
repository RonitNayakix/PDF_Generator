import os
import pandas as pd
import sqlite3
import tempfile
import shutil
import subprocess
from docxtpl import DocxTemplate
from docx2pdf import convert
import streamlit as st

# ==============================
# LibreOffice fallback for PDFs
# ==============================
def convert_docx_to_pdf(docx_path, pdf_path):
    soffice = shutil.which("soffice")
    if not soffice:
        possible_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
        ]
        for path in possible_paths:
            if os.path.exists(path):
                soffice = path
                break
    if not soffice:
        raise RuntimeError("LibreOffice 'soffice.exe' not found. Please install LibreOffice or update path.")
    
    outdir = os.path.dirname(pdf_path)
    cmd = [soffice, "--headless", "--convert-to", "pdf", "--outdir", outdir, docx_path]
    subprocess.run(cmd, check=True)

    if not os.path.exists(pdf_path):
        raise RuntimeError(f"PDF not created: {pdf_path}")

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
# Generate PDFs
# ==============================
def generate_pdfs(template_file, df):
    pdf_paths = []
    for idx, row in df.iterrows():
        doc = DocxTemplate(template_file)
        context = row.to_dict()

        tmp_docx = os.path.join(tempfile.gettempdir(), f"temp_{idx}.docx")
        out_pdf = os.path.join(tempfile.gettempdir(), f"{context.get('OrderID', idx)}_{context.get('Name', 'Record')}.pdf")

        doc.render(context)
        doc.save(tmp_docx)

        try:
            convert(tmp_docx, out_pdf)  # First try docx2pdf
        except Exception as e:
            print("docx2pdf failed, falling back to LibreOffice:", e)
            convert_docx_to_pdf(tmp_docx, out_pdf)

        pdf_paths.append(out_pdf)
        os.remove(tmp_docx)

    return pdf_paths

# ==============================
# Streamlit UI
# ==============================
st.set_page_config(page_title="Custom PDF Generator", page_icon="📄", layout="centered")

st.title("📄 Custom PDF Generator")
st.markdown("Upload a **DOCX template** (with placeholders) and a **CSV/Excel file**. The app will fill in the template with each row of data and generate PDFs.")

template_file = st.file_uploader("Upload DOCX Template", type=["docx"])
data_file = st.file_uploader("Upload CSV/Excel Data", type=["csv", "xlsx"])

if template_file and data_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_tpl:
        tmp_tpl.write(template_file.read())
        template_path = tmp_tpl.name

    if data_file.name.endswith(".csv"):
        df = pd.read_csv(data_file)
    else:
        df = pd.read_excel(data_file)

    # Save to SQLite
    save_to_sqlite(df)

    st.write("### Preview of Uploaded Data")
    st.dataframe(df.head())

    if st.button("Generate PDFs"):
        pdf_paths = generate_pdfs(template_path, df)

        st.success("✅ PDFs generated successfully!")
        for pdf in pdf_paths:
            with open(pdf, "rb") as f:
                st.download_button(
                    label=f"⬇ Download {os.path.basename(pdf)}",
                    data=f,
                    file_name=os.path.basename(pdf),
                    mime="application/pdf"
                )
