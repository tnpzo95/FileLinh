import pandas as pd
import re
import os

# ==================================================
# ĐƯỜNG DẪN FILE
# ==================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

INPUT_FILE = os.path.join(BASE_DIR, "Total.xlsx")
OUTPUT_FILE = os.path.join(BASE_DIR, "output_sorted.xlsx")

print("INPUT :", INPUT_FILE)
print("OUTPUT:", OUTPUT_FILE)

# ==================================================
# ĐỌC EXCEL
# ==================================================
df = pd.read_excel(INPUT_FILE, dtype=str)

# fill empty
df = df.fillna("")

# ==================================================
# HÀM SẮP XẾP NGÀY
# ==================================================
def parse_sort_key(value):
    text = str(value).strip()

    # dd/MM/yyyy
    full_date_pattern = r"^(\d{1,2})/(\d{1,2})/(\d{4})$"

    # /MM/yyyy
    month_only_pattern = r"^/(\d{1,2})/(\d{4})$"

    # Có ngày đầy đủ
    m1 = re.match(full_date_pattern, text)
    if m1:
        day = int(m1.group(1))
        month = int(m1.group(2))
        year = int(m1.group(3))

        return (
            year,
            month,
            0,
            day
        )

    # Chỉ có tháng
    m2 = re.match(month_only_pattern, text)
    if m2:
        month = int(m2.group(1))
        year = int(m2.group(2))

        return (
            year,
            month,
            1,
            999
        )

    # lỗi -> xuống cuối
    return (
        9999,
        9999,
        9999,
        9999
    )

# ==================================================
# CỘT E = INDEX 4
# ==================================================
df["__sort_key__"] = df.iloc[:, 4].apply(parse_sort_key)

# ==================================================
# SORT
# ==================================================
df = df.sort_values(by="__sort_key__")

# xóa cột tạm
df = df.drop(columns=["__sort_key__"])

# ==================================================
# EXPORT
# ==================================================
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    df.to_excel(writer, index=False)

print("DONE")
print("FILE:", OUTPUT_FILE)