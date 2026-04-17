import sys
import io
import pandas as pd
import gspread
import time
from google.oauth2.service_account import Credentials
from gspread.exceptions import WorksheetNotFound

# ======================
# FIX ENCODING (SAFE)
# ======================
try:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding='utf-8')
except:
    pass

# ======================
# CONFIG
# ======================
EXCEL_FILE = "output.xlsx"
SPREADSHEET_ID = "1FVu_-BWCk_c7rjtC5ovq4wSish8U7bx3ay-KhNiYqXY"
TARGET_SHEET = "upload"
SERVICE_FILE = "responsive-task-492802-h3-0f08af796138.json"

# ======================
# READ EXCEL
# ======================
df = pd.read_excel(EXCEL_FILE)
df = df.astype(str)

print(f"Upload {len(df)} dong...")

# ======================
# AUTH GOOGLE
# ======================
scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

creds = Credentials.from_service_account_file(
    SERVICE_FILE,
    scopes=scope
)

client = gspread.authorize(creds)
spreadsheet = client.open_by_key(SPREADSHEET_ID)

# ======================
# GET SHEET
# ======================
try:
    worksheet = spreadsheet.worksheet(TARGET_SHEET)
except WorksheetNotFound:
    worksheet = spreadsheet.add_worksheet(
        title=TARGET_SHEET,
        rows="1000",
        cols="20"
    )

worksheet.clear()

# ======================
# PREP DATA
# ======================
data = [df.columns.tolist()] + df.values.tolist()

# ======================
# UPLOAD SAFE
# ======================
for i in range(3):
    try:
        worksheet.update(
            range_name="A1",
            values=data,
            value_input_option="USER_ENTERED"
        )
        print("Upload thanh cong!")
        break

    except Exception as e:
        print(f"Retry {i+1}/3: {e}")
        time.sleep(2)

else:
    print("Upload that bai")
