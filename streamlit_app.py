import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import zipfile
from pypdf import PdfReader
from reportlab.pdfgen import canvas

# --- PAGE CONFIG ---
st.set_page_config(page_title="DocuGen Pro", layout="wide")
st.title("📄 Document Generator & Processor")

# --- SIDEBAR NAVIGATION ---
mode = st.sidebar.radio("Choose an Operation:", 
                       ["Bulk Generate (Word + Excel)", "Read/Extract PDF", "Generate Basic PDF"])

# ==========================================
# MODE 1: BULK GENERATE WORD DOCS FROM EXCEL
# ==========================================
if mode == "Bulk Generate (Word + Excel)":
    st.subheader("Bulk Generate Word Documents")
    st.write("1. Create a Word document. Use `{{ column_name }}` for variables.")
    st.write("2. Upload the template and an Excel file with matching column names.")

    col1, col2 = st.columns(2)
    with col1:
        template_file = st.file_uploader("1. Upload Word Template (.docx)", type=["docx"])
    with col2:
        data_file = st.file_uploader("2. Upload Data (.xlsx)", type=["xlsx"])

    if template_file and data_file:
        # Read Data
        df = pd.read_excel(data_file)
        st.success("Data loaded successfully!")
        st.dataframe(df.head()) # Preview data

        naming_col = st.selectbox("Which column should be used to name the files?", df.columns)

        if st.button("Generate Documents"):
            with st.spinner("Generating documents..."):
                # Create a ZIP file in memory
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    
                    for index, row in df.iterrows():
                        # Load template
                        doc = DocxTemplate(template_file)
                        
                        # Convert row to dictionary
                        context = row.to_dict()
                        
                        # Render template with data
                        doc.render(context)
                        
                        # Save to memory stream
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        doc_io.seek(0)
                        
                        # Add to ZIP
                        file_name = f"{row[naming_col]}_document.docx"
                        zip_file.writestr(file_name, doc_io.getvalue())
                
                st.success("Documents generated successfully!")
                
                # Download Button
                st.download_button(
                    label="⬇️ Download All as ZIP",
                    data=zip_buffer.getvalue(),
                    file_name="Generated_Documents.zip",
                    mime="application/zip"
                )

# ==========================================
# MODE 2: READ AND EXTRACT PDF
# ==========================================
elif mode == "Read/Extract PDF":
    st.subheader("Extract Text from PDF")
    pdf_file = st.file_uploader("Upload a PDF Document", type=["pdf"])

    if pdf_file:
        reader = PdfReader(pdf_file)
        num_pages = len(reader.pages)
        st.write(f"**Total Pages:** {num_pages}")
        
        page_num = st.number_input("Select Page to Extract", min_value=1, max_value=num_pages, value=1)
        
        if st.button("Extract Text"):
            page = reader.pages[page_num - 1]
            text = page.extract_text()
            st.text_area("Extracted Text", text, height=300)

# ==========================================
# MODE 3: GENERATE BASIC PDF
# ==========================================
elif mode == "Generate Basic PDF":
    st.subheader("Generate a Simple PDF")
    
    pdf_title = st.text_input("Document Title", "My Generated PDF")
    pdf_body = st.text_area("Document Content", "Write your text here...")
    
    if st.button("Generate PDF"):
        # Create PDF in memory
        pdf_buffer = io.BytesIO()
        c = canvas.Canvas(pdf_buffer)
        
        # Add Title
        c.setFont("Helvetica-Bold", 20)
        c.drawString(100, 750, pdf_title)
        
        # Add Body Text
        c.setFont("Helvetica", 12)
        text_object = c.beginText(100, 700)
        for line in pdf_body.split('\n'):
            text_object.textLine(line)
        c.drawText(text_object)
        
        c.save()
        pdf_buffer.seek(0)
        
        st.success("PDF created!")
        st.download_button(
            label="⬇️ Download PDF",
            data=pdf_buffer,
            file_name=f"{pdf_title.replace(' ', '_')}.pdf",
            mime="application/pdf"
        ) streamlit as st

st.title("🎈 My new app")
st.write(
    "Let's start building! For help and inspiration, head over to [docs.streamlit.io](https://docs.streamlit.io/)."
)
