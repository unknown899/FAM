
#!/usr/bin/env python3

from pathlib import Path
import argparse
import html
import re

import pandas as pd


def count_folder_rows(html_text: str, folder: str) -> int:
    """
    只計算 HTML 表格中：
        <tr><td>folder</td>
    的出現次數，不計算 href、src 等位置。
    """
    escaped_folder = re.escape(html.escape(folder))

    pattern = re.compile(
        rf"<tr\s*>\s*<td\s*>\s*{escaped_folder}\s*</td\s*>",
        flags=re.IGNORECASE,
    )

    return len(pattern.findall(html_text))


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "檢查 Excel 的 folder 是否在 index.html "
            "表格中恰好出現一次。"
        )
    )
    parser.add_argument(
        "html_path",
        type=Path,
        help="index.html 路徑",
    )
    parser.add_argument(
        "--excel",
        type=Path,
        default=Path("MFIS_dataset.xlsx"),
        help="Excel 路徑，預設為 MFIS_dataset.xlsx",
    )
    parser.add_argument(
        "--sheet",
        default="experiments",
        help='Worksheet 名稱，預設為 "experiments"',
    )
    args = parser.parse_args()

    experiments = pd.read_excel(
        args.excel,
        sheet_name=args.sheet,
    )

    if "folder" not in experiments.columns:
        raise KeyError(
            f'Worksheet "{args.sheet}" 中找不到 folder 欄位'
        )

    html_text = args.html_path.read_text(
        encoding="utf-8",
        errors="replace",
    )

    problem_count = 0

    for excel_row, value in enumerate(
        experiments["folder"],
        start=2,
    ):
        if pd.isna(value):
            continue

        folder = str(value).strip()

        if not folder:
            continue

        count = count_folder_rows(
            html_text=html_text,
            folder=folder,
        )

        if count == 0:
            problem_count += 1
            print(
                f"[MISSING] Excel row {excel_row}: {folder}"
            )

        elif count >= 2:
            problem_count += 1
            print(
                f"[DUPLICATE x{count}] "
                f"Excel row {excel_row}: {folder}"
            )

    if problem_count == 0:
        print(
            "[OK] 所有 folder 都在 index.html "
            "表格中恰好出現一次。"
        )
    else:
        print(f"\n共發現 {problem_count} 筆異常。")


if __name__ == "__main__":
    main()

