from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import pandas as pd
import shutil  # Untuk mengganti nama file

# Atur opsi untuk WebDriver
options = webdriver.ChromeOptions()
download_dir = r'C:\Users\hendr\Downloads\cek data'  # Ganti dengan direktori download yang diinginkan
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
options.add_experimental_option("prefs", prefs)

# Membuat instance WebDriver
driver = webdriver.Chrome(options=options)

# List untuk menyimpan hasil jumlah data dan id
result_list = []

try:
    # Buka halaman login
    driver.get('https://url')  # Ganti dengan URL login yang tepat
    time.sleep(2)

    # Masukkan admin_username dan admin_password
    admin_username_input = driver.find_element(By.NAME, 'admin_username')
    admin_password_input = driver.find_element(By.NAME, 'admin_password')
    login_button = driver.find_element(By.XPATH, '//button[@type="submit"]')

    admin_username_input.send_keys('username')
    admin_password_input.send_keys('password')
    login_button.click()

    # Tunggu proses login selesai
    time.sleep(5)

    # Looping untuk mengakses halaman dengan id dari 1 hingga 48
    for entry_id in range(123,127):  # Looping dari ID 1 sampai 48
        # Buka URL untuk mengunduh file berdasarkan ID
        driver.get(f'https://url/view_entry.php?form_id=88530&entry_id={entry_id}')
        time.sleep(5)  # Tunggu beberapa detik agar halaman sepenuhnya dimuat

        # Explicit wait untuk memastikan elemen muncul
        wait = WebDriverWait(driver, 20)
        download_button = wait.until(
            EC.presence_of_element_located((By.CLASS_NAME, 'entry_link'))
        )

        # Klik tombol unduh
        download_button.click()
        time.sleep(10)  # Tunggu beberapa detik agar unduhan selesai

        # Cari file terbaru di direktori unduhan
        files = os.listdir(download_dir)
        paths = [os.path.join(download_dir, file) for file in files if file.endswith('.xlsx')]
        if paths:
            latest_file = max(paths, key=os.path.getctime)  # Pilih file yang paling baru
            print(f"File yang diunduh: {latest_file}")

            # Ganti nama file menjadi '204.xlsx', '205.xlsx', dan seterusnya
            new_file_name = f"{203 + entry_id}.xlsx"  # Mengganti dengan nama baru
            new_file_path = os.path.join(download_dir, new_file_name)
            
            # Mengganti nama file
            shutil.move(latest_file, new_file_path)  # Mengganti nama file
            print(f"File telah diganti nama menjadi: {new_file_path}")

            # Membaca file Excel menggunakan pandas
            df = pd.read_excel(new_file_path, header=None)  # Baca menggunakan nama baru

            # Menemukan kolom "No" secara dinamis
            no_column_index = None
            for i in range(df.shape[1]):
                if df[i].astype(str).str.contains('No', case=False).any():
                    no_column_index = i
                    break

            if no_column_index is not None:
                # Mengonversi kolom "No" ke tipe numeric
                df[no_column_index] = pd.to_numeric(df[no_column_index], errors='coerce')

                # Menghitung jumlah entri di kolom "No"
                max_no = df[no_column_index].max()  # Mendapatkan nilai maksimum di kolom "No"
                print(f"Jumlah data untuk ID {entry_id}: {max_no}")
                
                # Menyimpan hasil ke list
                result_list.append({
                    "ID": entry_id,
                    "Jumlah Data": max_no
                })
            else:
                print(f"Kolom 'No' tidak ditemukan untuk ID {entry_id}.")
        else:
            print(f"Tidak ada file Excel yang ditemukan untuk ID {entry_id}.")
finally:
    driver.quit()  # Pastikan untuk menutup browser

# Setelah looping selesai, simpan hasil ke dalam file Excel
df_result = pd.DataFrame(result_list)  # Membuat DataFrame dari hasil
df_result.to_excel(os.path.join(download_dir, 'hasil_perhitungan.xlsx'), index=False)  # Simpan hasil sebagai file Excel
print("Hasil perhitungan disimpan di 'hasil_perhitungan.xlsx'.")
