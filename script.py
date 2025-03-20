import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time

# Setup ChromeDriver menggunakan WebDriver Manager
service = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
options.add_argument("--user-data-dir=selenium_session")  # Simpan sesi login
driver = webdriver.Chrome(service=service, options=options)

# Buka WhatsApp Web dan tunggu pengguna login
driver.get("https://web.whatsapp.com")
input("🔄 Scan QR Code di WhatsApp Web lalu tekan ENTER...")

# Membaca nomor dari file CSV
file_csv = "nomor.csv"  # Pastikan file ini ada di direktori yang sama dengan script
df = pd.read_csv(file_csv, dtype={"nomor": str})  # Pastikan nomor terbaca sebagai string
df.columns = df.columns.str.strip()  # Bersihkan spasi tersembunyi
df = df.dropna(subset=['nomor'])  # Hapus baris kosong

# Konversi data ke list
numbers = df['nomor'].astype(str).str.strip().tolist()

# Memasukkan pesan
message = input("Masukkan pesan yang ingin dikirim: ")

# Kirim pesan ke setiap nomor dengan jeda 10 detik
for number in numbers:
    try:
        url = f"https://web.whatsapp.com/send?phone={number}&text={message}"
        print(f"🔍 Membuka: {url}")
        driver.get(url)

        # Tunggu hingga input chat tersedia
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true']"))
        )
        time.sleep(2)  # Tambahan jeda agar lebih stabil

        # Klik tombol kirim
        send_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Kirim']"))
        )
        send_button.click()
        time.sleep(5)  # Tunggu agar pesan terkirim

        print(f"✅ Pesan berhasil dikirim ke {number}")

        # Jeda 10 detik sebelum mengirim ke nomor berikutnya
        print("⏳ Menunggu 10 detik sebelum mengirim ke nomor berikutnya...")
        time.sleep(10)

    except Exception as e:
        print(f"❌ Gagal mengirim pesan ke {number}: {e}")

# Tutup browser setelah selesai
driver.quit()
