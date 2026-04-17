import sys
import io
import re
import pandas as pd

# =========================
# FIX WINDOWS ENCODING
# =========================
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# =========================
# FILE CONFIG
# =========================
input_file = "datasauget.xlsx"
output_file = "output.xlsx"

# =========================
# READ EXCEL
# =========================
df = pd.read_excel(input_file)

# Normalize columns
df.columns = df.columns.str.strip().str.lower()
print("Cac cot tim thay:", df.columns.tolist())

# =========================
# COLUMN MAP
# =========================
col_ma_kh = "ma_kh"
col_time = "thoi_gian_tra_cuu"
col_noidung = "ket_qua"

rows = []

# =========================
# PROCESS ROWS
# =========================
for _, row in df.iterrows():

    ma_kh = row.get(col_ma_kh, "")
    thoi_gian_tra_cuu = row.get(col_time, "")
    text = str(row.get(col_noidung, ""))

    # =========================
    # EXTRACT INFO
    # =========================
    kh_match = re.search(r"KHÁCH HÀNG:\s*(.+)", text, re.IGNORECASE)
    dc_match = re.search(r"ĐỊA CHỈ:\s*(.+)", text, re.IGNORECASE)

    khach_hang = kh_match.group(1).strip() if kh_match else ""
    dia_chi = dc_match.group(1).strip() if dc_match else ""

    # =========================
    # SPLIT BLOCKS (SAFE)
    # =========================
    lich_blocks = re.split(r"(?=MÃ.*LỊCH)", text)

    for block in lich_blocks:

        ma_lich_match = re.search(r"MÃ.*LỊCH:\s*(\d+)", block, re.IGNORECASE)

        tg_match = re.search(
            r"THỜI GIAN:\s*từ (.+?) ngày (.+?) đến (.+?) ngày (.+)",
            block,
            re.IGNORECASE
        )

        ly_do_match = re.search(
            r"LÝ DO NGỪNG CUNG CẤP ĐIỆN:\s*(.+)",
            block,
            re.IGNORECASE
        )

        if ma_lich_match and tg_match:

            ma_lich = ma_lich_match.group(1)

            start_time = tg_match.group(1).strip()
            start_date = tg_match.group(2).strip()
            end_time = tg_match.group(3).strip()
            end_date = tg_match.group(4).strip()

            ly_do = ly_do_match.group(1).strip() if ly_do_match else ""

            rows.append([
                ma_kh,
                thoi_gian_tra_cuu,
                khach_hang,
                dia_chi,
                ma_lich,
                start_date,
                start_time,
                end_date,
                end_time,
                ly_do
            ])

# =========================
# EXPORT EXCEL
# =========================
result = pd.DataFrame(rows, columns=[
    "MA_KH",
    "Thoi gian tra cuu",
    "Khach hang",
    "Dia chi",
    "Ma lich",
    "Ngay bat dau",
    "Gio bat dau",
    "Ngay ket thuc",
    "Gio ket thuc",
    "Ly do"
])

result.to_excel(output_file, index=False)

print("HOAN TAT:", output_file)
print("Tong dong:", len(result))