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

# --- FUNGSI GENERATE PDF (DIPERBARUI) ---
def generate_kwitansi(data):
    # A5 Landscape: Lebar 210mm, Tinggi 148mm, Margin default 10mm (Area kerja: 190mm)
    pdf = FPDF(orientation='L', unit='mm', format='A5') 
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    
    # Judul
    pdf.cell(0, 10, "KWITANSI", ln=1, align='C')
    pdf.ln(5)
    
    # Fungsi pembantu untuk baris dengan teks panjang (Auto Text-Wrap)
    def add_wrapped_row(label, value, is_red_value=False):
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Arial", 'B', 10)
        
        # Kolom Label (Lebar 50mm)
        pdf.cell(50, 6, label, 0, 0)
        # Kolom Titik Dua (Lebar 5mm)
        pdf.cell(5, 6, ":", 0, 0)
        
        # Kolom Value dengan Multi-cell (Bisa otomatis turun ke baris baru)
        if is_red_value:
            pdf.set_text_color(255, 0, 0)
        
        # 0 artinya lebar teks akan mengisi sisa kertas sampai batas margin kanan
        pdf.multi_cell(0, 6, str(value), 0, 'L')
        pdf.set_text_color(0, 0, 0) # Kembalikan ke warna hitam

    # Mengisi Data Header (Teks panjang akan otomatis dirapikan)
    add_wrapped_row("NO", data['Agreement No.'])
    add_wrapped_row("Sudah terima dari", "PT. YAMAHA INDONESIA MOTOR MANUFACTURING", is_red_value=True)
    add_wrapped_row("Terbilang", data['Terbilang'])
    add_wrapped_row("Untuk pembayaran", data['Activity Theme'])
    
    pdf.ln(5) # Spasi antara deskripsi dan angka
    
    # --- Tabel Rincian Angka ---
    pdf.set_font("Arial", 'B', 10)
    
    # Subtotal
    pdf.cell(50, 6, "Subtotal", 0, 0)
    pdf.cell(5, 6, "", 0, 0) # Penyeimbang posisi titik dua
    pdf.cell(10, 6, "Rp", 0, 0)
    pdf.cell(35, 6, f"{data['Claim Amount']:,.2f}", 0, 1, 'R')
    
    # Flag Faktur
    pdf.cell(50, 6, "Faktur Pajak Required Flag", 0, 0)
    pdf.cell(5, 6, ":", 0, 0)
    pdf.cell(0, 6, "Input Faktur Pajak 11%", 0, 1)
    
    # PPN
    pdf.cell(50, 6, "PPN", 0, 0)
    pdf.cell(5, 6, "", 0, 0)
    pdf.cell(10, 6, "Rp", 0, 0)
    pdf.cell(35, 6, f"{data['PPN']:,.2f}", 0, 1, 'R')
    
    # Pajak Reklame
    pdf.cell(50, 6, "Pajak Reklame", 0, 0)
    pdf.cell(5, 6, "", 0, 0)
    pdf.cell(10, 6, "Rp", 0, 0)
    pdf.cell(35, 6, f"{data['Pajak Reklame']:,.2f}", 0, 1, 'R')
    
    pdf.ln(2) # Spasi sebelum blok lokasi/tanggal
    
    # Lokasi dan Tanggal di Kanan
    pdf.cell(105, 6, "", 0, 0) # Spasi kosong untuk mendorong ke kanan
    pdf.cell(85, 6, f"{data['Lokasi']} , {data['Tanggal_Format']}", 0, 1, 'C')
    
    pdf.ln(1)
    
    # Jumlah
    pdf.set_font("Arial", 'B', 11)
    pdf.cell(20, 6, "", 0, 0) # Geser tulisan Jumlah ke posisi pas
    pdf.cell(30, 6, "Jumlah", 0, 0, 'C')
    pdf.cell(15, 6, "Rp", 0, 0, 'C')
    pdf.cell(35, 6, f"{data['Jumlah']:,.2f}", 0, 1, 'R')
    
    # Ruang kosong agar ada tempat untuk Tanda Tangan
    pdf.ln(15) 
    
    # Nama PIC di Kanan Bawah
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(105, 6, "", 0, 0)
    pdf.cell(85, 6, data['PIC'], 0, 1, 'C')
    
    return pdf.output(dest='S').encode('latin-1')


# --- ANTARMUKA STREAMLIT ---
st.set_page_config(page_title="Kwitansi Generator", layout="wide")
st.title("🧾 Generator Kwitansi Batch")
st.write("Upload file Excel. Teks otomatis disesuaikan (auto-wrap) dan dirapikan.")

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
                file_name="Kwitansi_Batch_Rapi.zip",
                mime="application/zip"
            )