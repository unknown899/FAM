from pathlib import Path
import shutil
import subprocess
import shlex


# ==========================
# 使用者設定
# ==========================

# 指定要處理的 Exec 資料夾
EXEC_DIRS = [
    Path("/home/bowei/FAM/FerroX/Exec/"),

    # 有其他 Exec 資料夾就繼續加入
    # Path("/home/gkaiwi58/FAM/FerroX_other/Exec"),
    # Path("/home/gkaiwi58/project2/Exec"),
]

# notebook 來源資料夾
SOURCE_FOLDER = "MFIS_t_8_nomi_33"

# notebook 名稱
NOTEBOOK_NAME = "analyze.ipynb"

# MFIS_t_7_nomi_(某些正整數)
DELETE_NOMI_NUMBERS = {
    8, 9, 10, 11, 12, 13, 14, 15, 16, 17
}

# 某個 notebook 執行失敗後，是否繼續執行其他資料夾
CONTINUE_ON_ERROR = True


# ==========================
# 建立目標資料夾名稱
# ==========================
TARGET_FOLDER_NAMES = [
    "MFIS_t_8_nomi_31",

    *[
        f"MFIS_t_7_nomi_{number}"
        for number in sorted(DELETE_NOMI_NUMBERS)
    ],

    "MFIS_t_7_nomi_4_1e-6",
    "MFIS_t_7_nomi_5_1e-6",

    "MFIS_t_7_nomi_4_bg=50",
    "MFIS_t_7_nomi_4_bg=200",
]


# ==========================
# 統計資訊
# ==========================
copied_count = 0
success_count = 0
failed_count = 0
missing_count = 0

failed_notebooks = []
missing_folders = []


# ==========================
# 逐一處理 Exec 資料夾
# ==========================
for exec_dir in EXEC_DIRS:
    exec_dir = exec_dir.expanduser().resolve()

    print("\n" + "=" * 80)
    print(f"處理 Exec：{exec_dir}")

    if not exec_dir.is_dir():
        print(f"[錯誤] Exec 資料夾不存在：{exec_dir}")
        missing_count += len(TARGET_FOLDER_NAMES)
        continue

    # 每個 Exec 使用自己 MFIS_t_8_nomi_33 內的 analyze.ipynb
    source_notebook = (
        exec_dir
        / SOURCE_FOLDER
        / NOTEBOOK_NAME
    )

    print(f"來源 notebook：{source_notebook}")

    if not source_notebook.is_file():
        print(f"[錯誤] 找不到來源 notebook：{source_notebook}")
        missing_count += len(TARGET_FOLDER_NAMES)
        continue

    # ==========================
    # 處理每個目標資料夾
    # ==========================
    for folder_name in TARGET_FOLDER_NAMES:
        target_dir = exec_dir / folder_name
        target_notebook = target_dir / NOTEBOOK_NAME

        print("\n" + "-" * 80)
        print(f"目標資料夾：{target_dir}")

        if not target_dir.is_dir():
            print(f"[跳過] 資料夾不存在：{target_dir}")

            missing_count += 1
            missing_folders.append(target_dir)
            continue

        # ==========================
        # 複製並覆蓋 analyze.ipynb
        # ==========================
        try:
            shutil.copy2(
                source_notebook,
                target_notebook,
            )

            copied_count += 1
            print(f"[完成] notebook 已覆蓋：{target_notebook}")

        except Exception as error:
            print(f"[失敗] 無法複製 notebook：{error}")

            failed_count += 1
            failed_notebooks.append(target_notebook)

            if not CONTINUE_ON_ERROR:
                raise

            continue

        # ==========================
        # 執行 analyze.ipynb
        # ==========================
        command = [
            "jupyter",
            "nbconvert",
            "--to",
            "notebook",
            "--execute",
            "--inplace",
            NOTEBOOK_NAME,
        ]

        print(f"[執行] {shlex.join(command)}")
        print(f"[工作目錄] {target_dir}")

        try:
            result = subprocess.run(
                command,

                # 重要：
                # 必須在目標資料夾內執行，
                # analyze.ipynb 裡的相對路徑才會指向正確資料
                cwd=target_dir,

                # 不直接丟出 CalledProcessError，
                # 由 returncode 判斷是否成功
                check=False,
            )

        except FileNotFoundError:
            print(
                "[失敗] 找不到 jupyter 指令，"
                "請確認目前 Python/Conda 環境已安裝 Jupyter。"
            )

            failed_count += 1
            failed_notebooks.append(target_notebook)

            if not CONTINUE_ON_ERROR:
                raise

            continue

        except Exception as error:
            print(f"[失敗] 執行 notebook 時發生錯誤：{error}")

            failed_count += 1
            failed_notebooks.append(target_notebook)

            if not CONTINUE_ON_ERROR:
                raise

            continue

        if result.returncode == 0:
            success_count += 1
            print(f"[成功] notebook 執行完成：{target_notebook}")

        else:
            failed_count += 1
            failed_notebooks.append(target_notebook)

            print(
                f"[失敗] notebook 執行失敗，"
                f"return code = {result.returncode}"
            )

            if not CONTINUE_ON_ERROR:
                raise RuntimeError(
                    f"Notebook 執行失敗：{target_notebook}"
                )


# ==========================
# 顯示執行結果
# ==========================
print("\n" + "=" * 80)
print("全部處理完成")
print("=" * 80)

print(f"成功複製 notebook：{copied_count}")
print(f"成功執行 notebook：{success_count}")
print(f"執行或複製失敗：{failed_count}")
print(f"不存在的資料夾：{missing_count}")

if missing_folders:
    print("\n不存在的目標資料夾：")

    for folder in missing_folders:
        print(f"  - {folder}")

if failed_notebooks:
    print("\n執行或複製失敗的 notebook：")

    for notebook in failed_notebooks:
        print(f"  - {notebook}")