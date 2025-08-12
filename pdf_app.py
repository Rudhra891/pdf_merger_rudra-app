import streamlit as st
import pandas as pd
import tempfile, os
from io import BytesIO
from PyPDF2 import PdfMerger
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, LongTable, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.lib import colors
import fitz  # PyMuPDF
from PIL import Image

# Helpers for Excel → PDF conversion
def sanitize_columns(df):
    return df.rename(columns=lambda c: "" if "Unnamed" in str(c) else c)

def compute_stretched_widths(df, font, size, avail, pad=6, minw=50, maxw=400):
    cols = df.columns.tolist()
    if not cols:
        return []
    raw = []
    for col in cols:
        try:
            max_w = stringWidth(str(col), font, size)
            sample = df[col].dropna().astype(str).sample(min(200, len(df)))
            for v in sample:
                max_w = max(max_w, stringWidth(v, font, size))
            raw.append(max(minw, min(max_w + pad, maxw)))
        except Exception:
            raw.append(minw)
    total = sum(raw)
    return [avail / len(raw)] * len(raw) if total == 0 else [w * avail / total for w in raw]

def df_to_flowables(df, sheet, font="Helvetica", size=8):
    df = df.where(pd.notnull(df), "")
    df = sanitize_columns(df)
    hdr = ParagraphStyle("hdr", fontName=f"{font}-Bold", fontSize=size + 1)
    cell = ParagraphStyle("cell", fontName=font, fontSize=size)
    if df.empty or df.shape[1] == 0:
        return [
            Paragraph(f"<b>Sheet: {sheet}</b>", hdr),
            Spacer(1, 6),
            Paragraph("(Empty sheet)", cell),
            Spacer(1, 6),
        ]
    page_w, _ = landscape(A4)
    avail = page_w - 30 * mm
    widths = compute_stretched_widths(df, font, size, avail)
    flow = [Paragraph(f"<b>Sheet: {sheet}</b>", hdr), Spacer(1, 6)]
    data = [[Paragraph(str(c), hdr) for c in df.columns]] + [
        [Paragraph(str(v), cell) for v in row] for _, row in df.iterrows()
    ]
    tbl = LongTable(data, colWidths=widths or None)
    tbl.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.25, colors.black),
        ("FONTSIZE", (0, 0), (-1, -1), size),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    flow += [tbl, Spacer(1, 6)]
    return flow

def make_pdf(data_bytes, out_path, fsize=8):
    try:
        sheets = pd.read_excel(BytesIO(data_bytes), sheet_name=None, engine="openpyxl")
    except Exception as e:
        raise ValueError(f"Error reading Excel file: {e}")
    doc = SimpleDocTemplate(out_path, pagesize=landscape(A4),
                            leftMargin=15*mm, rightMargin=15*mm,
                            topMargin=15*mm, bottomMargin=15*mm)
    elems = []
    for name, df in sheets.items():
        elems += df_to_flowables(df, name, size=fsize)
    doc.build(elems)

# UI setup with styling
st.set_page_config(page_title="PDF Tool_Rudra", layout="wide")

st.markdown("""
<div style="text-align: center; margin-bottom: 10px;">
    <img src="https://bebpl.com/wp-content/uploads/2023/07/BLUE-ENERGY-lFINAL-LOGO.png" width="250"
         width="250" alt="BEBPL Logo">
</div>
""", unsafe_allow_html=True)

st.markdown("""
<style>
  .stApp { background: linear-gradient(to right, #ffff1c, #00c3ff); }
  .css-1d391kg { background: #ff00cc; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
  .css-1v3fvcr { background-color: #4CAF50; color: white; font-family: 'Arial', sans-serif; font-size: 24px;
                 padding: 10px; border-radius: 5px; }
  .stButton button { background-color: #4CAF50; color: white; border-radius: 5px;
                     padding: 10px 20px; font-size: 16px; border: none; }
  .stButton button:hover { background-color: #45a049; }
  .stTextInput input, .stTextArea textarea { border-radius: 5px; border: 1px solid #ccc;
                                            padding: 10px; font-size: 16px; }
  .stFileUploader input[type="file"] { border-radius: 5px; border: 1px solid #ccc;
                                       padding: 10px; font-size: 16px; }
</style>
""", unsafe_allow_html=True)

# Sidebar menu
st.sidebar.title("Select Task")
option = st.sidebar.radio("Choose operation:", ["Merge PDFs", "Image to PDF", "Excel to PDF","Compress PDF","Split PDF"])
st.title("📄 PDF WORLD-BEBPL")

# Tool-specific sections
if option == "Compress PDF":
    st.header("Compress PDF")
    uploaded_pdf = st.file_uploader("Upload a PDF file", type="pdf")
    dpi = st.slider("Select DPI (Lower = Smaller file)", 50, 200, 100, 10)
    if uploaded_pdf and st.button("Compress PDF"):
        with st.spinner("Compressing PDF..."):
            doc = fitz.open(stream=uploaded_pdf.read(), filetype="pdf")
            comp = fitz.open()
            for page in doc:
                pix = page.get_pixmap(dpi=dpi)
                rect = fitz.Rect(0, 0, pix.width, pix.height)
                newp = comp.new_page(width=rect.width, height=rect.height)
                newp.insert_image(rect, pixmap=pix)
            out = tempfile.mktemp(suffix=".pdf")
            comp.save(out, garbage=4, deflate=True)
            comp.close()
            doc.close()
        st.success("Compression complete!")
        with open(out, "rb") as f:
            st.download_button("Download Compressed PDF", f, file_name="compressed.pdf")



elif option == "Merge PDFs":
    st.header("Merge PDF Files")
    pdfs = st.file_uploader("Upload PDFs", type="pdf", accept_multiple_files=True)
    if pdfs:
        names = [f.name for f in pdfs]
        order = st.multiselect("Select merge order", options=names, default=names)
        if len(order) == len(names) and st.button("Merge"):
            merger = PdfMerger()
            for name in order:
                file = next(f for f in pdfs if f.name == name)
                merger.append(file)
            out = tempfile.mktemp(suffix=".pdf")
            merger.write(out)
            merger.close()
            st.success("Merged successfully!")
            with open(out, "rb") as f:
                st.download_button("Download Merged PDF", f, file_name="merged.pdf")

elif option == "Image to PDF":
    st.header("Convert Images to PDF & Merge")
    imgs = st.file_uploader("Upload Images (JPG/PNG)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
    if imgs:
        pdfs = []
        for img in imgs:
            tmp = tempfile.mktemp(suffix=os.path.splitext(img.name)[1])
            pdf_out = os.path.splitext(tmp)[0] + ".pdf"
            Image.open(img).convert("RGB").save(pdf_out, "PDF")
            pdfs.append(pdf_out)
            st.success(f"{img.name} → PDF")
            with open(pdf_out, "rb") as f:
                st.download_button(f"Download {os.path.basename(pdf_out)}", f, file_name=os.path.basename(pdf_out))
        if pdfs and st.button("Merge Images to PDF"):
            merger = PdfMerger()
            for p in pdfs:
                merger.append(p)
            out = tempfile.mktemp(suffix=".pdf")
            merger.write(out)
            merger.close()
            st.success("Images merged!")
            with open(out, "rb") as f:
                st.download_button("Download Merged Images PDF", f, file_name="images_merged.pdf")
                
                
elif option == "Excel to PDF":
    st.header("Convert Excel to Merged PDF")
    uploaded_excels = st.file_uploader("Upload Excel files", type=["xls", "xlsx"], accept_multiple_files=True)
    if uploaded_excels and st.button("Create Merged PDF"):
        tempdir = tempfile.mkdtemp()
        pdf_paths = []
        for f in uploaded_excels:
            st.info(f"Processing {f.name} …")
            out_pdf = os.path.join(tempdir, os.path.splitext(f.name)[0] + ".pdf")
            try:
                make_pdf(f.getvalue(), out_pdf, fsize=8)
                pdf_paths.append(out_pdf)
                st.success(f"{f.name} converted")
            except Exception as e:
                st.error(f"Error converting {f.name}: {e}")
        if pdf_paths:
            merged = os.path.join(tempdir, "merged_excel.pdf")
            merger = PdfMerger()
            for p in pdf_paths:
                merger.append(p)
            merger.write(merged)
            merger.close()
            with open(merged, "rb") as m:
                st.download_button("Download Merged PDF", m, file_name="merged_excel.pdf")                
                
                
                
elif option == "Split PDF":
    st.header("Extract Specific Page Range from PDF")
    uploaded_pdf = st.file_uploader("Upload a PDF file", type="pdf")
    
    from_page = st.number_input("From page (starting from 1)", min_value=1, step=1)
    to_page = st.number_input("To page (starting from 1)", min_value=1, step=1)

    if uploaded_pdf:
        doc = fitz.open(stream=uploaded_pdf.read(), filetype="pdf")
        total_pages = len(doc)
        st.info(f"Total pages: {total_pages}")

        if from_page < 1 or to_page > total_pages or from_page > to_page:
            st.error("Invalid page range. Please enter valid numbers.")
        else:
            if st.button("Extract Pages"):
                with st.spinner("Extracting selected pages..."):
                    extracted = fitz.open()
                    extracted.insert_pdf(doc, from_page=from_page - 1, to_page=to_page - 1)
                    out_path = tempfile.mktemp(suffix=".pdf")
                    extracted.save(out_path)
                    extracted.close()
                    doc.close()

                st.success(f"Pages {from_page} to {to_page} extracted successfully!")
                with open(out_path, "rb") as f:
                    st.download_button("Download Extracted PDF", f, file_name="extracted_pages.pdf")



st.link_button("About Us","https://bebpl.com/")
