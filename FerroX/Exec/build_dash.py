#!/usr/bin/env python3
"""Build the MFIS HTML dashboard from an Excel workbook.

The ``experiments`` worksheet supplies parameters, times, and status values,
including ``valid_loop`` and ``has_multi``.
PV/Pz images are located under each case folder in ``--data-root``. Use
``python build_dash.py --help`` for copyable command and import examples.
"""

from __future__ import annotations

import argparse
import html
import os
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook


DEFAULT_EXCEL_NAME = "MFIS_dataset.xlsx"
DEFAULT_SHEET_NAME = "experiments"
DEFAULT_OUTPUT_NAME = "index.html"

WANTED = [
    "alpha",
    "beta",
    "gamma",
    "BigGamma",
    "g11",
    "g44",
    "FE_lo",
    "FE_hi",
    "Ec",
    "Pr",
    "Pc0",
    "rp",
    "landau_consistency_pass",
    "valid_loop",
    "has_multi",
]

IMAGE_FILES = {
    "PV Curve": "MFIS_PV_curve.png",
    "Pz Stack": "Pz_FE_layer_stack.png",
}

HELP_EPILOG = r"""
資料來源:
  Excel 的 experiments worksheet：folder、參數、時間、valid_loop、has_multi 等欄位
  --data-root/folder/figs/：MFIS_PV_curve.png、Pz_FE_layer_stack.png

範例 1：使用預設檔案
  python build_dash.py

  預設讀取 ./MFIS_dataset.xlsx，從目前目錄尋找 case 圖片，產生 ./index.html。

範例 2：指定 Excel、圖片根目錄與輸出 HTML
  python build_dash.py \
      --excel temporary_dataset.xlsx \
      --data-root /home/bowei/FAM/FerroX/Exec \
      --output valid_preview.html

產生 HTML 後可在搜尋欄篩選狀態:
  has_multi=1 valid_loop=1

在其他 Python 程式中呼叫:
  import build_dash

  build_dash.build_dashboard(
      "temporary_dataset.xlsx",
      data_root="/home/bowei/FAM/FerroX/Exec",
      output_path="valid_preview.html",
      sheet_name="experiments",
  )
"""


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build index.html from the experiments worksheet.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=HELP_EPILOG,
    )
    parser.add_argument(
        "--excel",
        type=Path,
        default=Path(DEFAULT_EXCEL_NAME),
        help=f"Workbook to read (default: {DEFAULT_EXCEL_NAME}).",
    )
    parser.add_argument(
        "--sheet",
        default=DEFAULT_SHEET_NAME,
        help=f"Worksheet to read (default: {DEFAULT_SHEET_NAME}).",
    )
    parser.add_argument(
        "--data-root",
        type=Path,
        default=Path.cwd(),
        help="Directory containing the MFIS case folders (default: cwd).",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path(DEFAULT_OUTPUT_NAME),
        help=f"Output HTML path (default: {DEFAULT_OUTPUT_NAME}).",
    )
    return parser.parse_args()


def normalized_name(value: object) -> str:
    return "" if value is None else str(value).strip().casefold()


def resolve_under_root(path: Path, root: Path) -> Path:
    return path if path.is_absolute() else root / path


def sci(value: object) -> str:
    if value is None:
        return ""
    return f"{float(value):.3e}"


def format_parameter(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (int, float)):
        return sci(value)
    return str(value)


def format_plain(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    return str(value)


def find_images(folder: Path) -> dict[str, Path | None]:
    images: dict[str, Path | None] = {
        "PV Curve": None,
        "Pz Stack": None,
    }

    if not folder.is_dir():
        return images

    for title, filename in IMAGE_FILES.items():
        matches = list(folder.rglob(filename))
        if matches:
            images[title] = matches[0]

    return images


def read_dashboard_rows(
    excel_path: Path,
    data_root: Path,
    sheet_name: str,
) -> list[dict[str, object]]:
    workbook = load_workbook(excel_path, data_only=True, read_only=True)
    try:
        if sheet_name not in workbook.sheetnames:
            raise KeyError(
                f"Worksheet {sheet_name!r} was not found. "
                f"Available worksheets: {workbook.sheetnames}"
            )

        worksheet = workbook[sheet_name]
        header_values = next(
            worksheet.iter_rows(min_row=1, max_row=1, values_only=True),
            (),
        )
        header_lookup: dict[str, int] = {}
        for index, value in enumerate(header_values):
            name = normalized_name(value)
            if not name:
                continue
            if name in header_lookup:
                raise ValueError(f"Duplicate worksheet header: {value!r}")
            header_lookup[name] = index

        folder_column = next(
            (
                header_lookup[name]
                for name in ("folder", "file_name", "run_id")
                if name in header_lookup
            ),
            None,
        )
        if folder_column is None:
            raise KeyError(
                "No folder column was found. Expected one of: "
                "folder, file_name, run_id."
            )

        def row_value(values: tuple[object, ...], column_name: str) -> object:
            index = header_lookup.get(normalized_name(column_name))
            if index is None or index >= len(values):
                return None
            return values[index]

        rows: list[dict[str, object]] = []
        for values in worksheet.iter_rows(min_row=2, values_only=True):
            raw_folder = values[folder_column] if folder_column < len(values) else None
            if raw_folder is None or not str(raw_folder).strip():
                continue

            folder_name = str(raw_folder).strip()
            images = find_images(data_root / folder_name)
            rows.append(
                {
                    "folder": folder_name,
                    "params": {
                        parameter: row_value(values, parameter)
                        for parameter in WANTED
                    },
                    "tfe": row_value(values, "T_FE"),
                    "Start Time": row_value(values, "Start Time"),
                    "End Time": row_value(values, "End Time"),
                    "Elapsed": row_value(values, "Elapsed"),
                    **images,
                }
            )

        return sorted(rows, key=lambda row: str(row["folder"]))
    finally:
        workbook.close()


def image_href(image_path: Path, output_path: Path) -> str:
    relative_path = os.path.relpath(
        image_path.resolve(),
        start=output_path.parent.resolve(),
    )
    return Path(relative_path).as_posix()


def render_dashboard(rows: list[dict[str, object]], output_path: Path) -> None:
    last_update = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    html_text = f"""
<!DOCTYPE html>

<html>

<head>

<meta charset="utf-8">

<title>MFIS Dashboard</title>

<style>

body{{
    font-family:Arial;
    margin:25px;
}}

h1{{
    margin-bottom:10px;
}}

.subtitle{{
    color:#000;
    font-size:20px;
    font-weight:bold;
    margin-top:0px;
    margin-bottom:10px;
}}

.note{{
    color:#444;
    font-size:15px;
    margin-bottom:20px;
}}

input{{
    width:350px;
    padding:8px;
    font-size:15px;
    margin-bottom:15px;
}}

table{{
    border-collapse:collapse;
    width:100%;
}}

th{{
    position:sticky;
    top:0;
    background:#f2f2f2;
}}

th,td{{
    border:1px solid #bbb;
    padding:8px;
    text-align:center;
}}

tr:nth-child(even){{
    background:#fafafa;
}}

tr:hover{{
    background:#ffffdd;
}}

img{{
    max-width:180px;
    max-height:120px;
    transition:.2s;
}}

img:hover{{
    transform:scale(1.8);
    z-index:100;
}}

</style>

</head>

<body>

<h1>MFIS Dashboard</h1>

<p class="subtitle">
<div>
Last Update :
<span id="last_update">{last_update}</span>
</div>

<div>
Total Experiments :
<span id="total_exp">{len(rows)}</span>
</div>
</p>

<p class="note">
* All parameters are in SI units.
</p>

<p class="note">
* Searching format: parameter1.operator1(=, >=, <=, >, <).value1 (space) parameter2.operator2.value2 ...
</p>

<p class="note">
* Searching example: gamma<=1.5e11 T_FE>=8e-9.
</p>

<input id="search" placeholder="Search...">

<table id="tbl">

<thead>

<tr>

<th>Folder</th>
"""

    displayed_parameters = [
        parameter
        for parameter in WANTED
        if parameter not in {"FE_lo", "FE_hi"}
    ]
    for parameter in displayed_parameters:
        html_text += f"<th>{html.escape(parameter)}</th>\n"

    html_text += "<th>T_FE</th>"
    html_text += "<th>Start Time</th>"
    html_text += "<th>End Time</th>"
    html_text += "<th>Elapsed</th>"
    for title in IMAGE_FILES:
        html_text += f"<th>{html.escape(title)}</th>"

    html_text += """
</tr>

</thead>

<tbody id="exp_table">
"""

    for row in rows:
        html_text += "<tr>"
        html_text += f"<td>{html.escape(str(row['folder']))}</td>"

        params = row["params"]
        assert isinstance(params, dict)
        for parameter in displayed_parameters:
            text_value = format_parameter(params.get(parameter))
            html_text += f"<td>{html.escape(text_value)}</td>"

        html_text += f"<td>{html.escape(format_parameter(row['tfe']))}</td>"
        html_text += f"<td>{html.escape(format_plain(row['Start Time']))}</td>"
        html_text += f"<td>{html.escape(format_plain(row['End Time']))}</td>"
        html_text += f"<td>{html.escape(format_plain(row['Elapsed']))}</td>"

        for title in IMAGE_FILES:
            image = row[title]
            if isinstance(image, Path):
                href = html.escape(image_href(image, output_path), quote=True)
                html_text += f"""
<td>
<a href="{href}" target="_blank">
<img src="{href}">
</a>
</td>
"""
            else:
                html_text += "<td>No Image</td>"

        html_text += "</tr>"

    column_indexes = {
        parameter: index
        for index, parameter in enumerate(displayed_parameters, start=1)
    }
    tfe_index = len(displayed_parameters) + 1
    js_row_fields = []
    for parameter, index in column_indexes.items():
        if parameter == "landau_consistency_pass":
            expression = (
                f'cells[{index}].textContent.trim() === "True" ? 1 : 0'
            )
        else:
            expression = f"Number(cells[{index}].textContent)"
        js_row_fields.append(f"            {parameter}: {expression}")
    js_row_fields.append(f"            T_FE: Number(cells[{tfe_index}].textContent)")
    js_row_mapping = ",\n".join(js_row_fields)

    script = """

</tbody>

</table>

<script>

const search = document.getElementById("search");

function updateTable() {

    const filter = search.value.trim();
    const rows = document.querySelectorAll("#tbl tbody tr");

    let count = 0;

    // 將搜尋字串拆成多個條件
    // 例如 "alpha=-8e9 T_FE=8e-9"
    const conditions = filter === ""
        ? []
        : filter.split(/\\s+/).map(str => {
            const m = str.match(/^(\\w+)(<=|>=|=|<|>)(.+)$/);
            if (!m) return null;

            return {
                key: m[1],
                op: m[2],
                value: Number(m[3])
            };
        }).filter(c => c);

    rows.forEach(r => {

        const cells = r.querySelectorAll("td");

        if (cells.length < 8) {
            r.style.display = "none";
            return;
        }

        const row = {
__JS_ROW_MAPPING__
        };

        let show = true;

        for (const c of conditions) {

            const v = row[c.key];

            if (v === undefined) {
                show = false;
                break;
            }

            switch (c.op) {

                case "=":
                    if (v !== c.value) show = false;
                    break;

                case ">":
                    if (!(v > c.value)) show = false;
                    break;

                case "<":
                    if (!(v < c.value)) show = false;
                    break;

                case ">=":
                    if (!(v >= c.value)) show = false;
                    break;

                case "<=":
                    if (!(v <= c.value)) show = false;
                    break;
            }

            if (!show) break;
        }

        r.style.display = show ? "" : "none";

        if (show) count++;

    });

    document.getElementById("total_exp").textContent = count;
}

search.onkeyup = updateTable;
updateTable();

</script>

</body>

</html>
"""
    html_text += script.replace("__JS_ROW_MAPPING__", js_row_mapping)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(html_text, encoding="utf-8")


def build_dashboard(
    excel_path: str | Path = DEFAULT_EXCEL_NAME,
    *,
    data_root: str | Path = Path.cwd(),
    output_path: str | Path = DEFAULT_OUTPUT_NAME,
    sheet_name: str = DEFAULT_SHEET_NAME,
) -> int:
    """Read one Excel workbook and create the corresponding dashboard."""

    excel = Path(excel_path).expanduser().resolve()
    root = Path(data_root).expanduser().resolve()
    output = Path(output_path).expanduser().resolve()

    if not excel.is_file():
        raise FileNotFoundError(f"Workbook does not exist: {excel}")
    if not root.is_dir():
        raise NotADirectoryError(f"Data root does not exist: {root}")

    rows = read_dashboard_rows(excel, root, sheet_name)
    render_dashboard(rows, output)
    print(f"Generated {output} ({len(rows)} experiments)")
    return len(rows)


def main() -> int:
    args = parse_args()
    data_root = args.data_root.expanduser().resolve()
    excel_path = resolve_under_root(args.excel.expanduser(), data_root).resolve()
    output_path = resolve_under_root(args.output.expanduser(), data_root).resolve()

    try:
        build_dashboard(
            excel_path,
            data_root=data_root,
            output_path=output_path,
            sheet_name=args.sheet,
        )
    except (FileNotFoundError, NotADirectoryError, KeyError, ValueError) as exc:
        raise SystemExit(f"ERROR: {exc}") from exc
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
