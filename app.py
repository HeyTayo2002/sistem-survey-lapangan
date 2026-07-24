import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Alignment
import io
import os

# Konfigurasi Halaman
st.set_page_config(page_title="Form Berita Acara Survey", layout="centered")

# 1. Bagian Header & Input Nama Anggota
st.title("📋 Form Berita Acara Survey")
st.markdown("### 👤 Informasi Anggota")
nama_anggota = st.text_input("Nama dan Nomor Anggota", placeholder="Ketik nama dan nomor anggota di sini...")
st.markdown("---")

st.write("Silakan isi data hasil survey lapangan pada form di bawah ini.")

# Pastikan nama file sesuai dengan file Anda
file_template = "Template_Survey.xlsx" 

if not os.path.exists(file_template):
    st.error(f"⚠️ File '{file_template}' tidak ditemukan! Mohon masukkan file Excel tersebut ke dalam folder yang sama.")
else:
    wb = openpyxl.load_workbook(file_template)
    sheet = wb.active
    
    jawaban_user = {}
    sel_warna_input = [] # Menyimpan daftar kotak input (kuning & hijau)
    
    # Membaca Form mulai baris ke-4
    for row in range(4, sheet.max_row + 1):
        kategori = sheet.cell(row=row, column=3).value
        pertanyaan = sheet.cell(row=row, column=4).value
        sel_jawaban = sheet.cell(row=row, column=5)
        
        if type(sel_jawaban).__name__ == 'MergedCell':
            continue
            
        nilai_bawaan = sel_jawaban.value
        warna_sel = sel_jawaban.fill.start_color if sel_jawaban.fill else None
        
        if kategori and isinstance(kategori, str) and kategori.strip() != "":
            st.markdown(f"<br>### 📌 {kategori.upper()}", unsafe_allow_html=True)
            
        if pertanyaan and isinstance(pertanyaan, str) and pertanyaan.strip() != "":
            is_yellow = False
            is_green = False
            
            if warna_sel and warna_sel.type == 'theme':
                if warna_sel.theme == 7:
                    is_yellow = True
                    sel_warna_input.append(sel_jawaban)
                elif warna_sel.theme == 9:
                    is_green = True
                    sel_warna_input.append(sel_jawaban)
                    
            if is_green and isinstance(nilai_bawaan, str):
                pilihan_mentah = nilai_bawaan.split('/')
                pilihan_bersih = [p.strip() for p in pilihan_mentah if p.strip()] 
                pilihan_bersih.insert(0, "- Pilih Salah Satu -")
                
                jawaban = st.selectbox(label=pertanyaan, options=pilihan_bersih, key=f"baris_{row}")
                if jawaban != "- Pilih Salah Satu -":
                    jawaban_user[row] = jawaban
                    
            elif is_yellow:
                teks_petunjuk = str(nilai_bawaan).strip() if nilai_bawaan else ""
                jawaban = st.text_input(label=pertanyaan, placeholder=teks_petunjuk, key=f"baris_{row}")
                if jawaban:
                    jawaban_user[row] = jawaban
            else:
                st.markdown(f"**{pertanyaan}**")
                
    st.markdown("---")
    
    # Tombol Eksekusi
    if st.button("💾 Simpan Data & Buat Excel"):
        no_fill = PatternFill(fill_type=None)
        
        # 1. Masukkan jawaban & Hapus warna (Hanya pada kotak input kuning & hijau)
        for row, isi_jawaban in jawaban_user.items():
            sel = sheet.cell(row=row, column=5)
            sel.value = isi_jawaban
            
        for sel in sel_warna_input:
            sel.fill = no_fill
            
        # 2. Menggabungkan teks "Berita Acara Survey : " + Nama Anggota ke Kolom B2
        cell_b2 = sheet.cell(row=2, column=2)
        # Jika kolom nama kosong, beri teks aman agar tidak error
        nama_bersih = nama_anggota.strip() if nama_anggota else "Tanpa_Nama"
        cell_b2.value = f"Berita Acara Survey : {nama_bersih}"
        cell_b2.fill = no_fill  
        
        # Merge & Center dari Kolom B2 sampai Kolom F2
        sheet.merge_cells(start_row=2, start_column=2, end_row=2, end_column=6)
        cell_b2.alignment = Alignment(horizontal='center', vertical='center')
                    
        # 3. Mengatur Tinggi Baris (Row Height)
        for r in range(1, sheet.max_row + 1):
            sheet.row_dimensions[r].height = 30.05
            
        if sheet.max_row >= 1:
            sheet.row_dimensions[1].height = 13.05
        if sheet.max_row >= 86:
            sheet.row_dimensions[86].height = 13.05
                
        # 4. Menyimpan output dengan nama file dinamis sesuai permintaan
        output = io.BytesIO()
        wb.save(output)
        excel_data = output.getvalue()
        
        # Membuat nama file otomatis ("Pertanyaan Survey [Nama].xlsx")
        nama_file_aman = "".join(c for c in nama_bersih if c.isalnum() or c in (' ', '_', '-')).strip()
        nama_file_final = f"Pertanyaan Survey {nama_file_aman}.xlsx"
        
        st.success(f"✅ Berhasil! File siap diunduh dengan nama: **{nama_file_final}**")
        
        st.download_button(
            label="📥 Unduh File Excel Final",
            data=excel_data,
            file_name=nama_file_final,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )