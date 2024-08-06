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
excel_filename = input("Masukkan nama file Excel (misalnya sidalih.xlsx): ")  # Input nama file Excel

try:
    # Akses sheet berdasarkan nama
    sheet = client.open_by_key(spreadsheet_id).worksheet(sheet_name)
except gspread.exceptions.WorksheetNotFound:
    print(f"Sheet dengan nama '{sheet_name}' tidak ditemukan.")
    exit()
except gspread.exceptions.SpreadsheetNotFound:
    print(f"Spreadsheet dengan ID '{spreadsheet_id}' tidak ditemukan.")
    exit()

# Baca data dari file Excel
try:
    df = pd.read_excel(excel_filename, engine='openpyxl')
except FileNotFoundError:
    print(f"File '{excel_filename}' tidak ditemukan.")
    exit()
except Exception as e:
    print(f"Terjadi kesalahan saat membaca file Excel: {e}")
    exit()

# Konversi kolom NIK dan NKK menjadi string jika kolom tersebut ada
if 'NIK' in df.columns:
    df['NIK'] = df['NIK'].astype(str)
else:
    print("Kolom 'NIK' tidak ditemukan di file Excel.")
    
if 'NKK' in df.columns:
    df['NKK'] = df['NKK'].astype(str)
else:
    print("Kolom 'NKK' tidak ditemukan di file Excel.")

# Gantikan nilai NaN dengan string kosong
df = df.fillna("")

# Mengirim data ke Google Sheets
try:
    sheet.clear()  # Opsional: Hapus data yang ada sebelum mengirim data baru
    sheet.update([df.columns.values.tolist()] + df.values.tolist())
    print("Data berhasil dikirim ke Google Sheets")
except Exception as e:
    print(f"Terjadi kesalahan saat mengirim data ke Google Sheets: {e}")
