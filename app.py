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

def extract_crew_names(crew_string):
    """Extract only crew names without extra text like 'Excl Happy Sweep HC3'."""
    names = re.sub(r'Excl Happy Sweep HC\d*', '', crew_string)
    names = ', '.join([name.strip() for name in names.split(',') if name.strip()])
    return names

# ============ 🧾 PDF Invoice Generator ============

def generate_invoice(row, logo_bytes):
    pdf = FPDF()
    pdf.add_page()

    # Colors
    navy_color = (54, 79, 107)
    light_gray = (240, 240, 240)
    dark_gray = (90, 90, 90)
    accent_color = (93, 173, 226)  # Soft blue for highlights

    # Fonts
    pdf.set_font("Helvetica", '', 12)

    # Save logo temporarily
    logo_path = "temp_logo.png"
    with open(logo_path, "wb") as logo_file:
        logo_file.write(logo_bytes)

    # ✅ Add logo
    pdf.image(logo_path, x=10, y=5, w=70)

    # Invoice Title
    pdf.set_font("Helvetica", 'B', 24)
    pdf.set_text_color(*navy_color)
    pdf.set_xy(140, 12)
    pdf.cell(50, 10, 'INVOICE', ln=True, align='R')

    # Company Info
    pdf.set_font("Helvetica", '', 10)
    pdf.set_text_color(*dark_gray)
    pdf.set_xy(140, 25)
    pdf.multi_cell(60, 5, "Happy Sweep Cleaning Company\nPhone: +971 568780406\nEmail: happysweep.cleaning@gmail.com\nDubai - United Arab Emirates", align='R')
    pdf.ln(3)

    # Divider Line
    pdf.set_draw_color(220, 220, 220)
    pdf.set_line_width(0.5)
    pdf.line(10, 52, 200, 52)
    pdf.ln(5)

    # Invoice Number & Date
    pdf.set_font("Helvetica", '', 12)
    reference_no = sanitize_filename(row['Reference No'])
    pdf.cell(0, 8, f"Invoice Number: {clean_text(reference_no)}", ln=True)
    pdf.cell(0, 8, f"Date: {clean_text(row.get('Starting At', 'N/A'))}", ln=True)
    pdf.ln(5)

    # Client Information
    pdf.set_fill_color(*light_gray)
    pdf.set_font("Helvetica", 'B', 14)
    pdf.cell(0, 10, 'Client Information', ln=True, fill=True)

    pdf.set_font("Helvetica", '', 12)
    pdf.cell(0, 8, f"Client Name: {clean_text(row['Client Name'])}", ln=True)
    pdf.cell(0, 8, f"Client ID: {clean_text(row['Client ID'])}", ln=True)
    pdf.cell(0, 8, f"Address: {clean_text(row['Client Address'])}", ln=True)
    pdf.cell(0, 8, f"Region: {clean_text(row['Client Region'])}", ln=True)
    pdf.ln(5)

    # Service Details
    pdf.set_fill_color(*light_gray)
    pdf.set_font("Helvetica", 'B', 14)
    pdf.cell(0, 10, 'Service Details', ln=True, fill=True)

    pdf.set_font("Helvetica", '', 12)
    pdf.cell(0, 8, f"Service Start: {clean_text(row['Starting At'])}", ln=True)
    pdf.cell(0, 8, f"Service End: {clean_text(row['Ending At'])}", ln=True)

    # Extract clean crew names
    appointment_attributes = clean_text(row.get('Appointment Attributes', 'N/A'))
    assigned_crew_raw = clean_text(row.get('Assigned Crew Member', 'N/A'))
    assigned_crew = extract_crew_names(assigned_crew_raw)

    pdf.multi_cell(0, 8, f"Appointment Attributes:\n{appointment_attributes}")
    pdf.set_font("Helvetica", 'B', 12)
    pdf.cell(0, 8, f"Assigned Crew: {assigned_crew}", ln=True)
    pdf.ln(5)

    # Payment Information with VAT
    pdf.set_fill_color(*light_gray)
    pdf.set_font("Helvetica", 'B', 14)
    pdf.cell(0, 10, 'Payment Information', ln=True, fill=True)

    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Helvetica", '', 12)

    try:
        booking_amount = float(row['Booking Amount'])
    except ValueError:
        booking_amount = 0.0

    vat = booking_amount * 0.05
    total = booking_amount + vat

    pdf.cell(0, 8, f"Booking Amount: AED {booking_amount:.2f}", ln=True)
    pdf.cell(0, 8, f"VAT (5%): AED {vat:.2f}", ln=True)
    pdf.set_font("Helvetica", 'B', 14)
    pdf.cell(0, 10, f"Total Amount: AED {total:.2f}", ln=True)
    pdf.set_font("Helvetica", '', 12)
    pdf.cell(0, 8, f"Payment Method: {clean_text(row['Payment Method'])}", ln=True)

    # ✅ Divider after Payment Method
    pdf.set_draw_color(220, 220, 220)
    pdf.set_line_width(0.5)
    pdf.line(10, pdf.get_y() + 3, 200, pdf.get_y() + 3)
    pdf.ln(8)

    # ✅ Final Footer Line (Smaller, Italic, Black for Thank You)
    pdf.set_text_color(0, 0, 0)  # Black color
    pdf.set_font("Helvetica", 'I', 10)  # Italic and smaller
    pdf.cell(0, 10, "Thank you for choosing Happy Sweep Cleaning Company!", ln=True, align='C')

    pdf_output = pdf.output(dest='S').encode('latin1')
    filename = f"Invoice_{reference_no}.pdf"
    return filename, pdf_output

# ============ 🚀 Streamlit App ============

st.title("📋 Invoice Generator - Happy Sweep (Final Version)")
st.write("Upload your Excel file and generate invoices. Invoices will be downloaded directly through the browser.")

# Initialize session state
if 'invoices' not in st.session_state:
    st.session_state.invoices = []
if 'uploaded_file' not in st.session_state:
    st.session_state.uploaded_file = None
if 'logo' not in st.session_state:
    st.session_state.logo = None

# Upload logo
logo = st.file_uploader("Upload Company Logo", type=["png", "jpg", "jpeg"])
if logo:
    st.session_state.logo = logo.read()

# Upload Excel file
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
if uploaded_file:
    st.session_state.uploaded_file = uploaded_file

# Check if both files are uploaded
if st.session_state.uploaded_file and st.session_state.logo:
    df = pd.read_excel(st.session_state.uploaded_file)

    if st.button("Generate Invoices"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        total_rows = len(df)

        invoices = []

        for idx, (_, row) in enumerate(df.iterrows()):
            filename, pdf_bytes = generate_invoice(row, st.session_state.logo)
            invoices.append((filename, pdf_bytes))

            progress = (idx + 1) / total_rows
            progress_bar.progress(progress)
            status_text.text(f"Processing Invoice {idx + 1} of {total_rows}")

            time.sleep(0.01)

        st.session_state.invoices = invoices

        progress_bar.progress(1.0)
        status_text.text("✅ Invoices Generated Successfully!")

# If invoices exist, allow download
if st.session_state.invoices:
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zipf:
        for filename, pdf_bytes in st.session_state.invoices:
            zipf.writestr(filename, pdf_bytes)

    zip_buffer.seek(0)

    st.subheader("📁 Download All Invoices")
    st.download_button(
        label="📥 Download All Invoices (ZIP)",
        data=zip_buffer,
        file_name="Invoices.zip",
        mime="application/zip"
    )

    st.subheader("📄 Download Individual Invoices")
    for filename, pdf_bytes in st.session_state.invoices:
        st.download_button(
            label=f"📄 Download {filename}",
            data=pdf_bytes,
            file_name=filename,
            mime='application/pdf'
        )

    st.success("📁 Invoices are ready for download. Check your Downloads folder.")
