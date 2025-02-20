import streamlit as st
import pandas as pd
from fpdf import FPDF
import zipfile
import io
import re
import time

# ============ 🧹 Utility Functions ============

def clean_text(text):
    """Remove unsupported Unicode characters."""
    return re.sub(r'[^\x00-\x7F]+', '', str(text))

def sanitize_filename(text):
    """Remove special characters and limit filename length."""
    sanitized = re.sub(r'[\\/*?:"<>|]', "", str(text))
    return sanitized[:20]  # Limit filename length

# ============ 🧾 PDF Invoice Generator ============

def generate_invoice(row, logo_bytes):
    pdf = FPDF()
    pdf.add_page()

    # Using Helvetica (native font)
    pdf.set_font("Helvetica", '', 14)

    # Colors
    navy_color = (0, 0, 128)
    white_color = (255, 255, 255)

    # Save logo temporarily in memory
    logo_path = "temp_logo.png"
    with open(logo_path, "wb") as logo_file:
        logo_file.write(logo_bytes)

    # Add logo
    pdf.image(logo_path, x=10, y=8, w=100)

    # Invoice title
    pdf.set_font("Helvetica", 'B', 20)
    pdf.set_xy(150, 15)
    pdf.cell(50, 10, 'INVOICE', ln=True, align='R')
    pdf.ln(25)

    # Invoice number (cleaned)
    reference_no = sanitize_filename(row['Reference No'])
    pdf.set_font("Helvetica", '', 12)
    pdf.cell(0, 10, f"Invoice Number: {clean_text(reference_no)}", ln=True, align='R')
    pdf.ln(5)

    # Client Information
    pdf.set_fill_color(*navy_color)
    pdf.set_text_color(*white_color)
    pdf.set_font("Helvetica", 'B', 14)
    pdf.cell(0, 10, 'Client Information:', ln=True, fill=True)

    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Helvetica", '', 12)
    pdf.cell(0, 8, f"Client Name: {clean_text(row['Client Name'])}", ln=True)
    pdf.cell(0, 8, f"Client ID: {clean_text(row['Client ID'])}", ln=True)
    pdf.cell(0, 8, f"Address: {clean_text(row['Client Address'])}", ln=True)
    pdf.cell(0, 8, f"Region: {clean_text(row['Client Region'])}", ln=True)
    pdf.ln(5)

    # Service Details
    pdf.set_fill_color(*navy_color)
    pdf.set_text_color(*white_color)
    pdf.set_font("Helvetica", 'B', 14)
    pdf.cell(0, 10, 'Service Details:', ln=True, fill=True)

    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Helvetica", '', 12)
    pdf.cell(0, 8, f"Service Start: {clean_text(row['Starting At'])}", ln=True)
    pdf.cell(0, 8, f"Service End: {clean_text(row['Ending At'])}", ln=True)
    pdf.multi_cell(0, 8, f"Appointment Attributes:\n{clean_text(row['Appointment Attributes'])}")
    pdf.set_font("Helvetica", 'B', 12)
    pdf.cell(0, 8, f"Assigned Crew: {clean_text(row['Assigned Crew Member'])}", ln=True)
    pdf.ln(5)

    # Payment Information
    pdf.set_fill_color(*navy_color)
    pdf.set_text_color(*white_color)
    pdf.set_font("Helvetica", 'B', 14)
    pdf.cell(0, 10, 'Payment Information:', ln=True, fill=True)

    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Helvetica", '', 12)
    pdf.cell(0, 8, f"Booking Amount: AED {clean_text(row['Booking Amount'])}", ln=True)
    pdf.cell(0, 8, f"Payment Method: {clean_text(row['Payment Method'])}", ln=True)
    pdf.ln(10)

    # Footer
    pdf.set_fill_color(*navy_color)
    pdf.set_text_color(*white_color)
    pdf.set_font("Helvetica", 'I', 10)
    pdf.cell(0, 10, "Phone: +971 568780406 | Email: Happysweep.cleaning@gmail.com | Dubai - United Arab Emirates", ln=True, align='C', fill=True)

    # Return PDF as bytes
    pdf_output = pdf.output(dest='S').encode('latin1')
    filename = f"Invoice_{reference_no}.pdf"
    return filename, pdf_output

# ============ 🚀 Streamlit App ============

st.title("📋 Invoice Generator - Happy Sweep")
st.write("Upload your Excel file and generate invoices. Invoices will be downloaded directly through the browser.")

# Upload logo
logo = st.file_uploader("Upload Company Logo", type=["png", "jpg", "jpeg"])

# Upload Excel file
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file and logo:
    df = pd.read_excel(uploaded_file)

    if st.button("Generate Invoices"):
        # Read logo into bytes
        logo_bytes = logo.read()

        # Progress Bar
        progress_bar = st.progress(0)
        status_text = st.empty()
        total_rows = len(df)

        # In-memory list to store invoices
        invoices = []

        # Generate invoices with progress updates
        for idx, (_, row) in enumerate(df.iterrows()):
            filename, pdf_bytes = generate_invoice(row, logo_bytes)
            invoices.append((filename, pdf_bytes))

            # Update progress bar
            progress = (idx + 1) / total_rows
            progress_bar.progress(progress)
            status_text.text(f"Processing Invoice {idx + 1} of {total_rows}")

            time.sleep(0.01)  # Small delay for UI responsiveness

        # Complete progress bar
        progress_bar.progress(1.0)
        status_text.text("✅ Invoices Generated Successfully!")

        # ✅ Generate ZIP for All Invoices (In-Memory)
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as zipf:
            for filename, pdf_bytes in invoices:
                zipf.writestr(filename, pdf_bytes)

        zip_buffer.seek(0)  # Reset buffer pointer

        # ✅ Download Button for ZIP (FIRST OPTION)
        st.subheader("📁 Download All Invoices")
        st.download_button(
            label="📥 Download All Invoices (ZIP)",
            data=zip_buffer,
            file_name="Invoices.zip",
            mime="application/zip"
        )

        st.subheader("📄 Download Individual Invoices")
        # ✅ Individual Download Buttons (BELOW ZIP OPTION)
        for filename, pdf_bytes in invoices:
            st.download_button(
                label=f"📄 Download {filename}",
                data=pdf_bytes,
                file_name=filename,
                mime='application/pdf'
            )

        st.success("📁 Invoices are ready for download. Check your Downloads folder.")
