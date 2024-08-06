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

# Konversi nilai di kolom D dan E menjadi string (index 3 dan 4)
for row in data_filtered:
    row[3] = str(row[3])  # Kolom D
    row[4] = str(row[4])  # Kolom E

# Convert data to DataFrame
df = pd.DataFrame(data_filtered, columns=headers)

# Simpan DataFrame ke file Excel
df.to_excel("manual_output.xlsx", index=False, engine='openpyxl')

print("Data berhasil disimpan ke output.xlsx")
