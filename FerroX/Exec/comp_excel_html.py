#!/usr/bin/env python3

from argparse import ArgumentParser
from collections import Counter
from html.parser import HTMLParser
from pathlib import Path

import pandas as pd


class FolderTableParser(HTMLParser):
    """
    從指定 table 的每一個 <tr> 讀取第一個 <td>，
    視為該列的 folder。
    """

    def __init__(self, table_id: str = "tbl") -> None:
        super().__init__(convert_charrefs=True)

        self.table_id = table_id
        self.in_target_table = False
        self.target_table_depth = 0

        self.in_row = False
        self.td_index = 0
        self.capture_first_td = False
        self.first_td_parts: list[str] = []

        self.folders: list[str] = []

    def handle_starttag(
        self,
        tag: str,
        attrs: list[tuple[str, str | None]],
    ) -> None:
        attrs_dict = dict(attrs)

        if tag == "table":
            if self.in_target_table:
                self.target_table_depth += 1

            elif attrs_dict.get("id") == self.table_id:
                self.in_target_table = True
                self.target_table_depth = 1

            return

        if not self.in_target_table:
            return

        if tag == "tr":
            self.in_row = True
            self.td_index = 0
            self.capture_first_td = False
            self.first_td_parts = []

        elif tag == "td" and self.in_row:
            self.td_index += 1

            if self.td_index == 1:
                self.capture_first_td = True

    def handle_data(self, data: str) -> None:
        if self.capture_first_td:
            self.first_td_parts.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag == "table" and self.in_target_table:
            self.target_table_depth -= 1

            if self.target_table_depth == 0:
                self.in_target_table = False

            return

        if not self.in_target_table:
            return

        if tag == "td" and self.capture_first_td:
            self.capture_first_td = False

        elif tag == "tr" and self.in_row:
            folder = "".join(self.first_td_parts).strip()

            if folder:
                self.folders.append(folder)

            self.in_row = False
            self.capture_first_td = False
            self.first_td_parts = []


def read_excel_folders(
    excel_path: Path,
    sheet_name: str,
) -> list[str]:
    experiments = pd.read_excel(
        excel_path,
        sheet_name=sheet_name,
    )

    if "folder" not in experiments.columns:
        raise KeyError(
            f'Worksheet "{sheet_name}" 中找不到 folder 欄位'
        )

    folders = []

    for value in experiments["folder"]:
        if pd.isna(value):
            continue

        folder = str(value).strip()

        if folder:
            folders.append(folder)

    return folders


def read_html_folders(
    html_path: Path,
    table_id: str,
) -> list[str]:
    html_text = html_path.read_text(
        encoding="utf-8",
        errors="replace",
    )

    parser = FolderTableParser(table_id=table_id)
    parser.feed(html_text)

    return parser.folders


def report_counter(
    title: str,
    values: dict[str, int],
) -> None:
    if not values:
        return

    print(f"\n{title}")

    for folder, count in sorted(values.items()):
        print(f"  {folder}: x{count}")


def main() -> None:
    parser = ArgumentParser(
        description=(
            "比較 MFIS_dataset.xlsx 與 index.html 中的 folder，"
            "檢查缺失、重複及額外 case。"
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
        help="Excel 路徑",
    )

    parser.add_argument(
        "--sheet",
        default="experiments",
        help='Worksheet 名稱，預設為 "experiments"',
    )

    parser.add_argument(
        "--table-id",
        default="tbl",
        help='HTML table id，預設為 "tbl"',
    )

    args = parser.parse_args()

    excel_folders = read_excel_folders(
        excel_path=args.excel,
        sheet_name=args.sheet,
    )

    html_folders = read_html_folders(
        html_path=args.html_path,
        table_id=args.table_id,
    )

    excel_counts = Counter(excel_folders)
    html_counts = Counter(html_folders)

    excel_duplicates = {
        folder: count
        for folder, count in excel_counts.items()
        if count >= 2
    }

    html_duplicates = {
        folder: count
        for folder, count in html_counts.items()
        if count >= 2
    }

    missing_in_html = {
        folder: excel_counts[folder]
        for folder in excel_counts
        if folder not in html_counts
    }

    extra_in_html = {
        folder: html_counts[folder]
        for folder in html_counts
        if folder not in excel_counts
    }

    report_counter(
        "[DUPLICATE IN EXCEL]",
        excel_duplicates,
    )

    report_counter(
        "[DUPLICATE IN HTML]",
        html_duplicates,
    )

    report_counter(
        "[MISSING IN HTML]",
        missing_in_html,
    )

    report_counter(
        "[EXTRA IN HTML]",
        extra_in_html,
    )

    abnormal_folders = (
        set(excel_duplicates)
        | set(html_duplicates)
        | set(missing_in_html)
        | set(extra_in_html)
    )

    print("\n========== SUMMARY ==========")
    print(
        f"Excel：{len(excel_folders)} rows, "
        f"{len(excel_counts)} unique folders"
    )
    print(
        f"HTML ：{len(html_folders)} rows, "
        f"{len(html_counts)} unique folders"
    )

    if abnormal_folders:
        print(
            f"[FAILED] 共 {len(abnormal_folders)} 個 "
            "unique folder 有異常。"
        )
        raise SystemExit(1)

    print("[OK] Excel 與 HTML 的 folder 完全一致，且都沒有重複。")


if __name__ == "__main__":
    main()