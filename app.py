import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
import zipfile
import math

# --- FUNGSI TERBILANG ---
def terbilang(n):
    angka = ["", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas"]
    if n < 12:
        return angka[n]
    elif n < 20:
        return terbilang(n - 10) + " Belas"
    elif n < 100:
        return terbilang(n // 10) + " Puluh " + terbilang(n % 10)
    elif n < 200:
        return "Seratus " + terbilang(n - 100)
    elif n < 1000:
        return terbilang(n // 100) + " Ratus " + terbilang(n % 100)
    elif n < 2000:
        return "Seribu " + terbilang(n - 1000)
    elif n < 1000000:
        return terbilang(n // 1000) + " Ribu " + terbilang(n % 1000)
    elif n < 1000000000:
        return terbilang(n // 1000000) + " Juta " + terbilang(n % 1000000)
    elif n < 1000000000000:
        return terbilang(n // 1000000000) + " Milyar " + terbilang(n % 1000000000)
    else:
        return "Angka terlalu besar"

# --- FUNGSI FORMAT TANGGAL ---
def format_tanggal(date_val):
    bulan_indo = ["", "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"]
    date_str = str(int(date_val)).strip() 
    if len(date_str) == 8:
        y = date_str[:4]
        m = int(date_str[4:6])
        d = date_str[6:]
        return f"{d} {bulan_indo[m]} {y}"
    return str(date_val)

# --- FUNGSI GENERATE PDF ---
def generate_kwitansi(data):
    # A5 Landscape: Lebar 210mm, Tinggi 148mm
    pdf = FPDF(orientation='L', unit='mm', format='A5') 
    pdf.add_page()
    
    # Judul Kwitansi (Tebal & Tengah)
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "KWITANSI", ln=1, align='C')
    pdf.ln(5)
    
    # Fungsi pembantu untuk baris dengan teks panjang (Bold di Kiri, Regular di Kanan)
    def add_wrapped_row(label, value):
        # Set Font Label ke Bold
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(50, 6, label, 0, 0)
        pdf.cell(5, 6, ":", 0, 0)
        
        # Set Font Value ke Regular (Biasa)
        pdf.set_font("Arial", '', 10)
        pdf.multi_cell(0, 6, str(value), 0, 'L')

    # Mengisi Data Header
    add_wrapped_row("NO", data['Agreement No.'])
    add_wrapped_row("Sudah terima dari", "PT. YAMAHA INDONESIA MOTOR MANUFACTURING")
    add_wrapped_row("Terbilang", data['Terbilang'])
    add_wrapped_row("Untuk pembayaran", data['Activity Theme'])
    
    pdf.ln(5) # Spasi sebelum tabel angka
    
    # --- Tabel Rincian Angka ---
    
    # Subtotal
    pdf.set_font("Arial", 'B', 10) # Label Bold
    pdf.cell(50, 6, "Subtotal", 0, 0)
    pdf.cell(5, 6, "", 0, 0)
    
    pdf.set_font("Arial", '', 10) # Angka Regular
    pdf.cell(10, 6, "Rp", 0, 0)
    pdf.cell(35, 6, f"{data['Claim Amount']:,.2f}", 0, 1, 'R')
    
    # Flag Faktur
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(50, 6, "Faktur Pajak Required Flag", 0, 0)
    pdf.cell(5, 6, ":", 0, 0)
    
    pdf.set_font("Arial", '', 10)
    pdf.cell(0, 6, "Input Faktur Pajak 11%", 0, 1)
    
    # PPN
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(50, 6, "PPN", 0, 0)
    pdf.cell(5, 6, "", 0, 0)
    
    pdf.set_font("Arial", '', 10)
    pdf.cell(10, 6, "Rp", 0, 0)
    pdf.cell(35, 6, f"{data['PPN']:,.2f}", 0, 1, 'R')
    
    # Pajak Reklame
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(50, 6, "Pajak Reklame", 0, 0)
    pdf.cell(5, 6, "", 0, 0)
    
    pdf.set_font("Arial", '', 10)
    pdf.cell(10, 6, "Rp", 0, 0)
    pdf.cell(35, 6, f"{data['Pajak Reklame']:,.2f}", 0, 1, 'R')
    
    pdf.ln(2) 
    
    # Lokasi dan Tanggal (Regular)
    pdf.set_font("Arial", '', 10)
    pdf.cell(105, 6, "", 0, 0)
    pdf.cell(85, 6, f"{data['Lokasi']} , {data['Tanggal_Format']}", 0, 1, 'C')
    
    pdf.ln(1)
    
    # Jumlah (Label Bold, Angka Regular)
    pdf.set_font("Arial", 'B', 11)
    pdf.cell(20, 6, "", 0, 0) 
    pdf.cell(30, 6, "Jumlah", 0, 0, 'C')
    
    pdf.set_font("Arial", '', 11)
    pdf.cell(15, 6, "Rp", 0, 0, 'C')
    pdf.cell(35, 6, f"{data['Jumlah']:,.2f}", 0, 1, 'R')
    
    # Ruang kosong untuk Tanda Tangan
    pdf.ln(15) 
    
    # Nama PIC di Kanan Bawah (Regular)
    pdf.set_font("Arial", '', 10)
    pdf.cell(105, 6, "", 0, 0)
    pdf.cell(85, 6, data['PIC'], 0, 1, 'C')
    
    return pdf.output(dest='S').encode('latin-1')


# --- ANTARMUKA STREAMLIT ---
st.set_page_config(page_title="Kwitansi Generator", layout="wide")
st.title("🧾 Generator Kwitansi Yamaha")
st.write("Upload file Excel. Teks kiri tebal (bold), teks kanan biasa, dan dirapikan otomatis.")

uploaded_file = st.file_uploader("Upload Excel Template", type=['xlsx'])

if uploaded_file is not None:
    # Baca Excel
    df = pd.read_excel(uploaded_file)
    
    # Kalkulasi Otomatis
    df['Claim Amount'] = pd.to_numeric(df['Claim Amount'], errors='coerce').fillna(0)
    df['Pajak Reklame'] = pd.to_numeric(df['Pajak Reklame'], errors='coerce').fillna(0)
    
    df['PPN'] = df['Claim Amount'] * 0.11
    df['Jumlah'] = df['Claim Amount'] + df['PPN'] + df['Pajak Reklame']
    
    # Terapkan Terbilang & Tanggal
    df['Terbilang'] = df['Jumlah'].apply(lambda x: (" ".join(terbilang(math.floor(x)).split()) + " RUPIAH").upper())
    df['Tanggal_Format'] = df['Tax Invoice Date'].apply(format_tanggal)
    
    st.write("### Preview Data Setelah Kalkulasi:")
    st.dataframe(df[['Agreement No.', 'Activity Theme', 'Jumlah', 'Terbilang']])
    
    if st.button("Proses & Buat Kwitansi"):
        with st.spinner('Menyusun tata letak dan membuat kwitansi...'):
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for index, row in df.iterrows():
                    data_dict = row.to_dict()
                    pdf_bytes = generate_kwitansi(data_dict)
                    
                    safe_agreement_no = str(data_dict['Agreement No.']).replace('/', '-')
                    nama_file = f"Kwitansi_{data_dict['PIC']}_{safe_agreement_no}.pdf"
                    
                    zip_file.writestr(nama_file, pdf_bytes)
            
            st.success(f"Selesai! {len(df)} Kwitansi berhasil dibuat dengan rapi.")
            st.download_button(
                label="📥 Download Semua Kwitansi (.zip)",
                data=zip_buffer.getvalue(),
                file_name="Kwitansi_Batch_Final.zip",
                mime="application/zip"
            )

