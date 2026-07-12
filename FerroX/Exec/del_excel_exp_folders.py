from pathlib import Path
import re

from openpyxl import load_workbook


# ==========================
# 使用者設定
# ==========================
INPUT_PATH = Path("MFIS_dataset.xlsx")
OUTPUT_PATH = Path("MFIS_dataset.xlsx")

SHEET_NAME = "experiments"
HEADER_ROW = 1

# 填入要刪除的 MFIS_t_7_nomi_(正整數)
# 例如會刪除：
# MFIS_t_7_nomi_1
# MFIS_t_7_nomi_2
# MFIS_t_7_nomi_6
DELETE_NOMI_NUMBERS = {8,9,10,11,12,13,14,15,16,17}


# ==========================
# 判斷 folder 是否需要刪除
# ==========================
def should_delete(folder_value) -> bool:
    if folder_value is None:
        return False

    folder = str(folder_value).strip()

    # 條件 1：完全符合 MFIS_t_8_nomi_31
    if folder == "MFIS_t_8_nomi_31":
        return True

    # 條件 2：MFIS_t_7_nomi_(指定正整數)
    match = re.fullmatch(r"MFIS_t_7_nomi_(\d+)", folder)

    if match:
        number = int(match.group(1))

        if number in DELETE_NOMI_NUMBERS:
            return True

    # 條件 3：MFIS_t_7_nomi_4_1e-6 或 MFIS_t_7_nomi_5_1e-6
    if re.fullmatch(r"MFIS_t_7_nomi_(4|5)_1e-6", folder):
        return True

    # 條件 4：MFIS_t_7_nomi_4_bg=50 或 MFIS_t_7_nomi_4_bg=200
    if re.fullmatch(r"MFIS_t_7_nomi_4_bg=(50|200)", folder):
        return True

    return False


# ==========================
# 讀取 Excel
# ==========================
workbook = load_workbook(INPUT_PATH)

if SHEET_NAME not in workbook.sheetnames:
    raise ValueError(
        f"找不到工作表 {SHEET_NAME!r}，"
        f"目前工作表為：{workbook.sheetnames}"
    )

worksheet = workbook[SHEET_NAME]


# ==========================
# 找出 folder 欄位
# ==========================
header_to_column = {}

for column_index in range(1, worksheet.max_column + 1):
    header = worksheet.cell(
        row=HEADER_ROW,
        column=column_index
    ).value

    if header is not None:
        header_to_column[str(header).strip()] = column_index

if "folder" not in header_to_column:
    raise ValueError(
        f"第 {HEADER_ROW} 列找不到 'folder' 欄位。\n"
        f"目前欄位：{list(header_to_column.keys())}"
    )

folder_column = header_to_column["folder"]


# ==========================
# 找出需要刪除的 row
# ==========================
rows_to_delete = []
deleted_folders = []

for row_index in range(HEADER_ROW + 1, worksheet.max_row + 1):
    folder = worksheet.cell(
        row=row_index,
        column=folder_column
    ).value

    if should_delete(folder):
        rows_to_delete.append(row_index)
        deleted_folders.append(str(folder).strip())


# ==========================
# 從最後一列往上刪除
# 避免 row index 因刪除而改變
# ==========================
for row_index in reversed(rows_to_delete):
    worksheet.delete_rows(row_index, 1)


# ==========================
# 儲存結果
# ==========================
workbook.save(OUTPUT_PATH)

print(f"輸入檔案：{INPUT_PATH.resolve()}")
print(f"輸出檔案：{OUTPUT_PATH.resolve()}")
print(f"共刪除 {len(rows_to_delete)} 列")

if deleted_folders:
    print("\n刪除的 folder：")

    for folder in deleted_folders:
        print(f"  - {folder}")
else:
    print("沒有找到符合條件的資料列。")