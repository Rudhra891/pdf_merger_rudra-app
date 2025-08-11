import streamlit as st
from PyPDF2 import PdfMerger
import fitz  # PyMuPDF
from PIL import Image
import xlwings as xw
import tempfile
import os




# Set page configuration
st.set_page_config(page_title="PDF Tool_Rudra", layout="wide")

# background image




# Embed BEBPL logo at the top
st.markdown(
    """
    <div style="text-align: center; margin-bottom: 10px;">
        <img src="https://bebpl.com/wp-content/uploads/2023/07/BLUE-ENERGY-lFINAL-LOGO.png" width="250" alt="BEBPL Logo">
    </div>
    """,
    unsafe_allow_html=True
)


# Inject custom CSS for styling
st.markdown("""
    <style>
        /* Set background color */
        .stApp {
            background-color:background: background: #00c3ff;  /* fallback for old browsers */
background: -webkit-linear-gradient(to right, #ffff1c, #00c3ff);  /* Chrome 10-25, Safari 5.1-6 */
background: linear-gradient(to right, #ffff1c, #00c3ff); /* W3C, IE 10+/ Edge, Firefox 16+, Chrome 26+, Opera 12+, Safari 7+ */


  ;
        }

        /* Style the sidebar */
        .css-1d391kg {
            background-color: #0000FF;
            border-radius: 10px;
            background: #ff00cc;
	    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            
        }

        /* Style the header */
        .css-1v3fvcr {
            background-color: #4CAF50;
            color: white;
            font-family: 'Arial', sans-serif;
            font-size: 24px;
            padding: 10px;
            border-radius: 5px;
        }

        /* Style buttons */
        .stButton button {
            background-color: #4CAF50;
            color: white;
            border-radius: 5px;
            padding: 10px 20px;
            font-size: 16px;
            border: none;
        }

        .stButton button:hover {
            background-color: #45a049;
        }

        /* Style input fields */
        .stTextInput input, .stTextArea textarea {
            border-radius: 5px;
            border: 1px solid #ccc;
            padding: 10px;
            font-size: 16px;
        }

        /* Style file uploader */
        .stFileUploader input[type="file"] {
            border-radius: 5px;
            border: 1px solid #ccc;
            padding: 10px;
            font-size: 16px;
        }
    </style>
""", unsafe_allow_html=True)

# Sidebar for task selection
st.sidebar.title("Select Task")
option = st.sidebar.radio("Choose operation:", ("Merge PDFs", "Excel to PDF", "Image to PDF", "Compress PDF"))

st.title("📄 PDF Tool — BEBPL")


# 1. Merge PDFs with Custom Order
if option == "Merge PDFs":
    st.header("Merge PDF Files")
    uploaded = st.file_uploader("Upload PDF files", type="pdf", accept_multiple_files=True)
    if uploaded:
        st.markdown("**Select PDFs in your desired merge order:**")
        names = [f.name for f in uploaded]
        order = st.multiselect("Merge Order", options=names, default=names)
        if len(order) == len(names):
            if st.button("Merge in Selected Order"):
                merger = PdfMerger()
                for name in order:
                    f = next(file for file in uploaded if file.name == name)
                    merger.append(f)
                merged = tempfile.mktemp(suffix=".pdf")
                merger.write(merged)
                merger.close()
                st.success("PDFs merged in specified order.")
                with open(merged, "rb") as f:
                    st.download_button("Download Merged PDF", f, file_name="merged.pdf")
        else:
            st.warning("Please select all files in the order you want them merged.")

# 2. Excel to PDF (and merge)
elif option == "Excel to PDF":
    st.header("Convert Excel to PDF and Merge")
    ex_uploads = st.file_uploader("Upload Excel files", type=["xls", "xlsx"], accept_multiple_files=True)
    if ex_uploads and st.button("Convert & Merge Excel files"):
        pdf_paths = []
        tmp_dir = tempfile.mkdtemp()
        app = xw.App(visible=False)
        for u in ex_uploads:
            tmp = os.path.join(tmp_dir, u.name)
            with open(tmp, "wb") as of:
                of.write(u.getbuffer())
            wb = app.books.open(tmp)
            out_pdf = os.path.splitext(tmp)[0] + ".pdf"
            for sheet in wb.sheets:
                # Adjust page setup to fit content on one page
                sheet.page_setup.fit_to_pages_width = 1
                sheet.page_setup.fit_to_pages_height = 1
                # Optional: Set margins
                sheet.page_setup.left_margin = app.api.InchesToPoints(0.5)
                sheet.page_setup.right_margin = app.api.InchesToPoints(0.5)
                sheet.page_setup.top_margin = app.api.InchesToPoints(0.5)
                sheet.page_setup.bottom_margin = app.api.InchesToPoints(0.5)
            wb.to_pdf(out_pdf)
            wb.close()
            pdf_paths.append(out_pdf)
            st.success(f"{u.name} → PDF")
        app.quit()
        if pdf_paths:
            merger = PdfMerger()
            for p in pdf_paths:
                merger.append(p)
            merged = tempfile.mktemp(suffix=".pdf")
            merger.write(merged)
            merger.close()
            st.success("Excel PDFs merged successfully.")
            with open(merged, "rb") as f:
                st.download_button("Download Merged Excel PDF", f, file_name="excel_merged.pdf")

# 3. Image to PDF (and merge)
elif option == "Image to PDF":
    st.header("Convert Images to PDF and Merge")
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
        if pdfs and st.button("Merge Image PDFs"):
            merger = PdfMerger()
            for p in pdfs:
                merger.append(p)
            merged = tempfile.mktemp(suffix=".pdf")
            merger.write(merged)
            merger.close()
            st.success("Images merged into PDF.")
            with open(merged, "rb") as f:
                st.download_button("Download Merged Images PDF", f, file_name="images_merged.pdf")

# 4. Compress PDF
elif option == "Compress PDF":
    st.header("Compress PDF")
    uploaded_pdf = st.file_uploader("Upload a PDF", type="pdf", accept_multiple_files=False)
    dpi = st.slider("Select DPI (lower = smaller size)", 50, 200, 100, 10)
    if uploaded_pdf and st.button("Compress PDF"):
        with st.spinner("Compressing..."):
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
            st.success("PDF compressed.")
            with open(out, "rb") as f:
                st.download_button("Download Compressed PDF", f, file_name="compressed.pdf")

st.link_button("About Us","https://bebpl.com/")
