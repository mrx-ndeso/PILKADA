import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

# Setup Google Sheets API credentials
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

# Minta inputan dari pengguna
spreadsheet_id = input("Masukkan ID Google Sheet: ")  # Input ID Google Sheet
sheet_name = input("Masukkan nama sheet: ")  # Input nama sheet

# Buka Google Sheet dengan ID
sheet = client.open_by_key(spreadsheet_id).worksheet(sheet_name)  # Buka sheet dengan nama sesuai inputan

# Ambil data dari rentang A:S
data = sheet.get_all_values()  # Ambil semua nilai

# Pilih kolom A hingga S (index 0 hingga 18)
data_filtered = [row[:19] for row in data]

# Ambil header
headers = data_filtered[0]  # Header baris pertama
data_filtered = data_filtered[1:]  # Data tanpa header

# Cari indeks kolom untuk NIK dan NKK
index_nik = headers.index("NIK")
index_nkk = headers.index("NKK")

# Konversi nilai di kolom NIK dan NKK menjadi string
for row in data_filtered:
    row[index_nik] = str(row[index_nik])  # Kolom NIK
    row[index_nkk] = str(row[index_nkk])  # Kolom NKK

# Convert data to DataFrame
df = pd.DataFrame(data_filtered, columns=headers)

# Simpan DataFrame ke file Excel
df.to_excel("manual_output.xlsx", index=False, engine='openpyxl')

print("Data berhasil disimpan ke manual_output.xlsx")
