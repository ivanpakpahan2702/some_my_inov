import pandas as pd
import requests
import os
from urllib.parse import urlparse, parse_qs
import time

def download_google_drive_file(url, filename, save_folder='downloaded_files'):
    """
    Download file dari Google Drive
    """
    try:
        # Membuat folder jika belum ada
        os.makedirs(save_folder, exist_ok=True)
        
        # Extract file ID dari URL Google Drive
        if 'open?id=' in url:
            file_id = url.split('open?id=')[1]
        elif '/d/' in url:
            file_id = url.split('/d/')[1].split('/')[0]
        elif 'id=' in url:
            file_id = url.split('id=')[1].split('&')[0]
        else:
            file_id = url
        
        # URL untuk download
        download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
        
        # Session untuk menangani cookies
        session = requests.Session()
        
        print(f"Mendownload: {filename}")
        
        # Request ke Google Drive
        response = session.get(download_url, stream=True)
        
        # Cek jika file besar (Google Drive akan memberikan konfirmasi)
        if 'confirm=' not in response.url:
            # Tidak perlu konfirmasi, langsung download
            pass
        else:
            # Jika perlu konfirmasi, tambahkan parameter confirm
            params = parse_qs(urlparse(response.url).query)
            if 'confirm' in params:
                download_url = f"https://drive.google.com/uc?export=download&id={file_id}&confirm={params['confirm'][0]}"
                response = session.get(download_url, stream=True)
        
        # Simpan file
        filepath = os.path.join(save_folder, filename)
        with open(filepath, 'wb') as f:
            for chunk in response.iter_content(chunk_size=32768):
                if chunk:
                    f.write(chunk)
        
        print(f"✓ Berhasil disimpan: {filepath}")
        return True
        
    except Exception as e:
        print(f"✗ Gagal mendownload {filename}: {str(e)}")
        return False

def clean_filename(name):
    """
    Membersihkan nama file dari karakter yang tidak valid
    """
    # Ganti karakter yang tidak valid untuk nama file
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        name = name.replace(char, '_')
    return name.strip()

def main():
    # Baca file Excel utama
    input_file = 'data_pengadilan.xlsx'  # Ganti dengan nama file Anda
    
    try:
        # Baca data dari Excel
        df = pd.read_excel(input_file)
        
        print(f"Menemukan {len(df)} data pengadilan")
        print("-" * 50)
        
        success_count = 0
        fail_count = 0
        
        # Iterasi setiap baris
        for index, row in df.iterrows():
            # Ambil data dari setiap kolom
            timestamp = str(row['Timestamp'])
            pengadilan = str(row['Pengadilan Tinggi / Pengadilan Negeri'])
            nama_contact = str(row['Nama Contact Person'])
            url = str(row['Data Pengadilan (format excel)'])
            
            # Buat nama file yang unik
            # Gunakan timestamp dan nama pengadilan untuk menghindari duplikat
            clean_pengadilan = clean_filename(pengadilan)
            clean_nama = clean_filename(nama_contact)
            filename = f"{index+1:03d}_{clean_pengadilan}_{clean_nama}.xlsx"
            
            # Download file
            if url.startswith('http'):
                success = download_google_drive_file(url, filename)
                if success:
                    success_count += 1
                else:
                    fail_count += 1
                
                # Delay kecil untuk menghindari blokir
                time.sleep(1)
            else:
                print(f"✗ URL tidak valid pada baris {index+1}")
                fail_count += 1
        
        print("\n" + "="*50)
        print(f"DOWNLOAD SELESAI!")
        print(f"Berhasil: {success_count} file")
        print(f"Gagal: {fail_count} file")
        print(f"File tersimpan di folder: downloaded_files/")
        
    except FileNotFoundError:
        print(f"File '{input_file}' tidak ditemukan!")
        print("Pastikan file Excel ada di folder yang sama dengan script ini.")
    except Exception as e:
        print(f"Terjadi error: {str(e)}")

if __name__ == "__main__":
    print("="*50)
    print("GOOGLE DRIVE DOWNLOADER - DATA PENGADILAN")
    print("="*50)
    main()