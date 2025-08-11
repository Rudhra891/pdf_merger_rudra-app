import streamlit as st
from PyPDF2 import PdfMerger
import fitz  # PyMuPDF
import tempfile
import os
import base64

# Configure page layout
st.set_page_config(page_title="PDF Merger & Compressor_Rudra", layout="wide")

# Embed BEBPL logo at the top
st.markdown(
    """
    <div style="text-align: center; margin-bottom: 10px;">
        <img src="https://github.com/Rudhra891/pdf_merger_rudra-app/blob/main/bebpl.png" width="250" alt="BEBPL Logo">
    </div>
    """,
    unsafe_allow_html=True
)

# Styling and banner section
st.markdown(
    """
    <style>
      [data-testid="stAppViewContainer"] {
         background: #FC5C7D;  /* fallback for old browsers */
background: -webkit-linear-gradient(to right, #6A82FB, #FC5C7D);  /* Chrome 10-25, Safari 5.1-6 */
background: linear-gradient(to right, #6A82FB, #FC5C7D); /* W3C, IE 10+/ Edge, Firefox 16+, Chrome 26+, Opera 12+, Safari 7+ */
;
      }
      [data-testid="stHeader"] {
          visibility: hidden;
      }
      .banner img {
          width: 100%;
          height: auto;
          object-fit: cover;
          border-radius: 8px;
          margin-bottom: 20px;
      }
      button {
          background-color: #4A90E2;
          color: white;
          border-radius: 8px;
          padding: 0.5em 1em;
      }
      button:hover {
          background-color: #357ABD;
      }
      .card {
          background-color: white;
          padding: 15px;
          border-radius: 8px;
          box-shadow: 0 4px 6px rgba(0,0,0,0.1);
          margin-bottom: 20px;
      }
    </style>
    """,
    unsafe_allow_html=True
)

st.title("📄 PDF Merger and Compressor")

# Main UI container
st.markdown('<div class="card">', unsafe_allow_html=True)

# File uploader
uploaded_files = st.file_uploader(
    "Upload PDF files", type=["pdf"], accept_multiple_files=True
)

# Initialize session state if not already set
if "merged_pdf_path" not in st.session_state:
    st.session_state.merged_pdf_path = None
if "compressed_pdf_path" not in st.session_state:
    st.session_state.compressed_pdf_path = None

# Merge functionality
if uploaded_files:
    st.subheader("🔢 Set Merge Order")
    order_info = []
    for i, file in enumerate(uploaded_files):
        order = st.number_input(
            f"Order for **{file.name}**",
            min_value=1,
            max_value=len(uploaded_files),
            value=i + 1,
            key=f"order_{file.name}"
        )
        order_info.append((order, file))

    if st.button("🔗 Merge PDFs"):
        with st.spinner("Merging..."):
            try:
                ordered_files = [file for _, file in sorted(order_info, key=lambda x: x[0])]
                merger = PdfMerger()
                for pdf in ordered_files:
                    merger.append(pdf)
                merged_path = tempfile.mktemp(suffix=".pdf")
                merger.write(merged_path)
                merger.close()
                st.session_state.merged_pdf_path = merged_path
                merged_size = os.path.getsize(merged_path) / 1024
                st.success(f"✅ Merged PDF created! Size: {merged_size:.2f} KB")
                with open(merged_path, "rb") as f:
                    st.download_button("📥 Download Merged PDF", f, file_name="merged.pdf")
            except Exception as e:
                st.error(f"❌ Merge failed: {e}")

# Compression functionality
if st.session_state.merged_pdf_path:
    st.subheader("🗜️ Compress Merged PDF")
    dpi = st.slider(
        "Choose DPI (Lower = Smaller size, Lower quality)",
        50,
        200,
        100,
        10
    )
    if st.button("🔧 Compress Now"):
        with st.spinner("Compressing PDF..."):
            try:
                src_doc = fitz.open(st.session_state.merged_pdf_path)
                compressed_doc = fitz.open()
                for page in src_doc:
                    pix = page.get_pixmap(dpi=dpi)
                    rect = fitz.Rect(0, 0, pix.width, pix.height)
                    new_page = compressed_doc.new_page(width=rect.width, height=rect.height)
                    new_page.insert_image(rect, pixmap=pix)
                compressed_path = tempfile.mktemp(suffix=".pdf")
                compressed_doc.save(compressed_path, deflate=True)
                compressed_doc.close()
                src_doc.close()
                st.session_state.compressed_pdf_path = compressed_path
                compressed_size = os.path.getsize(compressed_path) / 1024
                st.success(f"✅ Compressed PDF created! Size: {compressed_size:.2f} KB")
            except Exception as e:
                st.error(f"❌ Compression failed: {e}")

# Download compressed PDF
if st.session_state.compressed_pdf_path:
    with open(st.session_state.compressed_pdf_path, "rb") as f:
        st.download_button("📥 Download Compressed PDF", f, file_name="compressed_merged.pdf")

st.markdown('</div>', unsafe_allow_html=True)
# By Rudra - Geophysicist


st.link_button("About Us","https://bebpl.com/")
