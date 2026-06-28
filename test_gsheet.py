import os
import gspread
from google.oauth2.service_account import Credentials

# Cargar variables de entorno (si usas .env)
from dotenv import load_dotenv
load_dotenv()

SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "1ADAkwPy_4Gyaf4ydM_fub3k2x7NuXPcTJhLDjDFVjV8")
CREDS_FILE = os.getenv("CREDS_FILE", "creds_nuevo.json")

print(f"🔍 Usando SHEET_ID: {SHEET_ID}")
print(f"🔍 Usando CREDS_FILE: {CREDS_FILE}")

try:
    creds = Credentials.from_service_account_file(CREDS_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SHEET_ID)
    print("✅ ¡ÉXITO! Hoja encontrada. Pestañas:", [ws.title for ws in sh.worksheets()])
except Exception as e:
    print("❌ ERROR:", e)