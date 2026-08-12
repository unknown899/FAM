#!/usr/bin/env python3

from __future__ import annotations

import argparse
import ast
import re
import sys
from pathlib import Path
from typing import Any

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


# ============================================================
# Constants
# ============================================================

DEFAULT_CSV_RELATIVE_PATH = Path("figs/MFIS_PV_curve.csv")
DEFAULT_INPUTS_FILENAME = "inputs"

WANTED_INPUT_PARAMETERS = {
    "alpha",
    "beta",
    "gamma",
    "BigGamma",
    "g11",
    "g44",
    "FE_lo",
    "FE_hi",
}

from functools import lru_cache

SCRIPT_DIR = Path(__file__).resolve().parent

# 相對路徑一律以這支 script 所在資料夾為基準
DEFAULT_DATASET_PATH = Path("MFIS_dataset.xlsx")

EXPERIMENT_SHEET = "experiments"
EXPERIMENT_FOLDER_COLUMN = "folder"

# ============================================================
# 在這裡手動填入要讀取的 parameter 欄位名稱
# 名稱必須和 experiments worksheet 的欄名完全相同
# ============================================================
WANTED_EXPERIMENT_PARAMETERS: tuple[str, ...] = (
     "alpha",
     "beta",
     "gamma",
     "Ec",
     "Pr",
     "rp",
)


DEFAULT_LEGEND_PARAMETERS = [
    #"alpha",
    #"beta",
    #"gamma",
    "Ec",
    "Pr",
    "rp",
]

PARAMETER_DISPLAY_NAMES = {
    "alpha": r"$\alpha$",
    "beta": r"$\beta$",
    "gamma": r"$\gamma$",
    "BigGamma": r"$\Gamma$",
    "g11": r"$g_{11}$",
    "g44": r"$g_{44}$",
    "Ec": r"$E_c$",
    "Pr": r"$P_r$",
    "rp": r"$r_p$",
    "FE_lo": "FE_lo",
    "FE_hi": "FE_hi",
    "t_FE": r"$t_{\mathrm{FE}}$",
}

def resolve_dataset_path(
    dataset_path: str | Path | None = None,
) -> Path:
    """
    未指定時，使用 script 同資料夾下的 MFIS_dataset.xlsx。

    若傳入相對路徑，例如 data/MFIS_dataset.xlsx，
    則以 script 所在資料夾為基準。
    """
    if dataset_path is None:
        path = DEFAULT_DATASET_PATH
    else:
        path = Path(dataset_path).expanduser()

    if not path.is_absolute():
        path = SCRIPT_DIR / path

    return path.resolve()


@lru_cache(maxsize=None)
def load_experiments(excel_path: Path) -> pd.DataFrame:
    """
    Excel 只讀取一次，避免每處理一個 folder 就重新讀取整份檔案。
    """
    return pd.read_excel(
        excel_path,
        sheet_name=EXPERIMENT_SHEET,
    )


def read_experiment_parameters(
    folder_name: str,
    dataset_path: str | Path | None = None,
) -> dict[str, Any]:
    """
    從 MFIS_dataset.xlsx 的 experiments worksheet，
    讀取指定 folder 對應的參數。
    """
    parameters: dict[str, Any] = {}

    excel_path = resolve_dataset_path(dataset_path)

    if not excel_path.is_file():
        print(
            f"[WARNING] Dataset does not exist: {excel_path}",
            file=sys.stderr,
        )
        return parameters

    experiments = load_experiments(excel_path)

    required_columns = {
        EXPERIMENT_FOLDER_COLUMN,
        *WANTED_EXPERIMENT_PARAMETERS,
    }
    missing_columns = sorted(
        required_columns - set(experiments.columns)
    )

    if missing_columns:
        raise KeyError(
            "The following columns do not exist in "
            f"{excel_path.name}/{EXPERIMENT_SHEET}: "
            f"{missing_columns}"
        )

    folder_mask = (
        experiments[EXPERIMENT_FOLDER_COLUMN]
        .astype("string")
        .str.strip()
        .eq(str(folder_name).strip())
        .fillna(False)
    )

    matched_rows = experiments.loc[folder_mask]

    if matched_rows.empty:
        print(
            f"[WARNING] Folder not found in Excel: {folder_name}",
            file=sys.stderr,
        )
        return parameters

    if len(matched_rows) > 1:
        raise ValueError(
            f"Folder {folder_name!r} appears "
            f"{len(matched_rows)} times in the experiments worksheet."
        )

    row = matched_rows.iloc[0]

    for parameter_name in WANTED_EXPERIMENT_PARAMETERS:
        value = row[parameter_name]
        #print(f"[DEBUG] {folder_name}: {parameter_name} = {value}")

        # 空白 Excel cell 不放入 parameters
        if pd.isna(value):
            continue

        # 將 numpy scalar 轉成一般 Python scalar
        if hasattr(value, "item"):
            value = value.item()

        parameters[parameter_name] = value

    return parameters

# ============================================================
# Voltage sequence
# ============================================================

def make_expected_voltage_sequence() -> np.ndarray:
    """
    建立 37 個電壓點：

    -4.5, -4.0, -3.5, ..., 4.5,
     4.0,  3.5,  3.0, ..., -4.5
    """

    # 使用整數再除以 2，避免 np.arange 的浮點誤差。
    forward = np.arange(-9, 10, dtype=float) / 2.0
    backward = np.arange(8, -10, -1, dtype=float) / 2.0

    voltage = np.concatenate([forward, backward])

    if len(voltage) != 37:
        raise RuntimeError(
            f"Unexpected voltage sequence length: {len(voltage)}"
        )

    return voltage


# ============================================================
# inputs parsing
# ============================================================

def parse_value(values: str) -> float | list[float]:
    """
    將 FerroX inputs 中的數值字串轉成 float 或 list[float]。

    例如：
        "-8.0e9"       -> -8.0e9
        "0 0 2.0e-9"   -> [0.0, 0.0, 2.0e-9]
    """

    values = values.split("#", 1)[0].strip()

    if not values:
        raise ValueError("Empty parameter value")

    parsed = [float(value) for value in values.split()]

    if len(parsed) == 1:
        return parsed[0]

    return parsed


def read_inputs_parameters(inputs_path: Path) -> dict[str, Any]:
    """
    從 inputs 檔案讀取指定參數。
    """

    parameters: dict[str, Any] = {}

    if not inputs_path.is_file():
        return parameters

    with inputs_path.open(
        "r",
        encoding="utf-8",
        errors="replace",
    ) as file:
        for raw_line in file:
            line = raw_line.strip()

            if not line or line.startswith("#") or "=" not in line:
                continue

            key, values = line.split("=", 1)
            key = key.strip()

            if key not in WANTED_INPUT_PARAMETERS:
                continue

            try:
                parameters[key] = parse_value(values.strip())
            except ValueError as error:
                print(
                    f"[WARNING] Cannot parse {key} in "
                    f"{inputs_path}: {error}",
                    file=sys.stderr,
                )

    return parameters


def calculate_t_fe_nm(
    parameters: dict[str, Any],
) -> float | None:
    """
    使用 FE_hi[2] - FE_lo[2] 計算鐵電層厚度，單位轉成 nm。
    """

    fe_lo = parameters.get("FE_lo")
    fe_hi = parameters.get("FE_hi")

    if not isinstance(fe_lo, list):
        return None

    if not isinstance(fe_hi, list):
        return None

    if len(fe_lo) < 3 or len(fe_hi) < 3:
        return None

    return (fe_hi[2] - fe_lo[2]) * 1.0e9


# ============================================================
# CSV reading
# ============================================================

def find_column(
    dataframe: pd.DataFrame,
    requested_name: str,
) -> str | None:
    """
    忽略欄位名稱前後空白及大小寫來尋找欄位。
    """

    normalized_requested = requested_name.strip().lower()

    for column in dataframe.columns:
        if str(column).strip().lower() == normalized_requested:
            return str(column)

    return None


def parse_list_cell(value: Any) -> list[float]:
    """
    處理 CSV 中可能以字串形式儲存的陣列。

    例如：
        "[-0.2, -0.1, 0.0]"
        "-0.2 -0.1 0.0"
        "-0.2,-0.1,0.0"
    """

    if pd.isna(value):
        return []

    if isinstance(value, (list, tuple, np.ndarray)):
        return [float(item) for item in value]

    text = str(value).strip()

    if not text:
        return []

    try:
        parsed = ast.literal_eval(text)

        if isinstance(parsed, (list, tuple, np.ndarray)):
            return [float(item) for item in parsed]

        return [float(parsed)]

    except (ValueError, SyntaxError):
        pass

    text = text.strip("[]()")

    parts = re.split(r"[\s,;]+", text)

    result = []

    for part in parts:
        if not part:
            continue
        result.append(float(part))

    return result


def extract_numeric_column(
    dataframe: pd.DataFrame,
    column_name: str,
) -> np.ndarray:
    """
    讀取一般逐列數值欄位，也支援單一儲存格內存放整個 list。
    """

    actual_column = find_column(dataframe, column_name)

    if actual_column is None:
        raise KeyError(
            f"Column '{column_name}' was not found. "
            f"Available columns: {list(dataframe.columns)}"
        )

    series = dataframe[actual_column]

    numeric_series = pd.to_numeric(series, errors="coerce")
    numeric_values = numeric_series.dropna().to_numpy(dtype=float)

    # 一般情況：每一列是一個數值。
    if len(numeric_values) > 1:
        return numeric_values

    # 單列純數字也直接回傳。
    if len(numeric_values) == 1 and len(series) == 1:
        return numeric_values

    # 嘗試解析字串形式的 list。
    parsed_values: list[float] = []

    for value in series:
        try:
            parsed_values.extend(parse_list_cell(value))
        except ValueError:
            continue

    if not parsed_values:
        raise ValueError(
            f"No numeric values could be read from column "
            f"'{actual_column}'."
        )

    return np.asarray(parsed_values, dtype=float)


def read_pv_curve(
    csv_path: Path,
    use_csv_voltage: bool,
) -> tuple[np.ndarray, np.ndarray]:
    """
    回傳：
        voltage, polarization

    預設使用固定的 37 點 voltage sequence。

    若 --use-csv-voltage：
        改讀取 CSV 中的 Vg_mean。
    """

    dataframe = pd.read_csv(csv_path)
    dataframe.columns = [
        str(column).strip()
        for column in dataframe.columns
    ]

    polarization = extract_numeric_column(
        dataframe,
        "P_mean",
    )

    expected_voltage = make_expected_voltage_sequence()

    if use_csv_voltage:
        voltage = extract_numeric_column(
            dataframe,
            "Vg_mean",
        )

        if len(voltage) != len(polarization):
            raise ValueError(
                f"Vg_mean length ({len(voltage)}) does not match "
                f"P_mean length ({len(polarization)})."
            )

        return voltage, polarization

    if len(polarization) == len(expected_voltage):
        return expected_voltage, polarization

    # 若不是 37 點，嘗試使用 CSV 內的 Vg_mean。
    vg_column = find_column(dataframe, "Vg_mean")

    if vg_column is not None:
        voltage = extract_numeric_column(
            dataframe,
            "Vg_mean",
        )

        if len(voltage) == len(polarization):
            print(
                f"[WARNING] {csv_path}: P_mean has "
                f"{len(polarization)} points instead of 37. "
                "Using Vg_mean from the CSV.",
                file=sys.stderr,
            )
            return voltage, polarization

    raise ValueError(
        f"P_mean contains {len(polarization)} points, "
        f"but the expected voltage sequence contains "
        f"{len(expected_voltage)} points."
    )


# ============================================================
# Folder discovery
# ============================================================

def natural_sort_key(path: Path) -> list[Any]:
    """
    讓 var_2 排在 var_10 前面。
    """

    text = path.as_posix()

    parts = re.split(r"(\d+(?:\.\d+)?)", text)

    key: list[Any] = []

    for part in parts:
        if re.fullmatch(r"\d+(?:\.\d+)?", part):
            key.append((0, float(part)))
        else:
            key.append((1, part.lower()))

    return key


def contains_target_csv(
    case_dir: Path,
    csv_relative_path: Path,
) -> bool:
    return (case_dir / csv_relative_path).is_file()


def discover_case_directories(
    root: Path,
    explicit_folders: list[str],
    prefixes: list[str],
    contains_strings: list[str],
    csv_relative_path: Path,
    recursive: bool,
) -> list[Path]:
    """
    所有篩選條件採聯集：

    1. --folders 明確指定的資料夾
    2. --prefix 開頭相符
    3. --contains 名稱或相對路徑包含字串

    若三者都未指定，會選取 root 下所有具有目標 CSV 的資料夾。
    """

    found: dict[Path, Path] = {}

    # --------------------------------------------------------
    # 1. Explicit folders
    # --------------------------------------------------------

    for folder_text in explicit_folders:
        folder = Path(folder_text).expanduser()

        if not folder.is_absolute():
            folder = root / folder

        folder = folder.resolve()

        if not folder.is_dir():
            print(
                f"[WARNING] Folder does not exist: {folder}",
                file=sys.stderr,
            )
            continue

        if not contains_target_csv(folder, csv_relative_path):
            print(
                f"[WARNING] CSV not found: "
                f"{folder / csv_relative_path}",
                file=sys.stderr,
            )
            continue

        found[folder] = folder

    # --------------------------------------------------------
    # 2. Search pool
    # --------------------------------------------------------

    if recursive:
        search_pool = (
            path
            for path in root.rglob("*")
            if path.is_dir()
        )
    else:
        search_pool = (
            path
            for path in root.iterdir()
            if path.is_dir()
        )

    has_name_filters = bool(prefixes or contains_strings)

    for candidate in search_pool:
        if not contains_target_csv(
            candidate,
            csv_relative_path,
        ):
            continue

        try:
            relative_text = candidate.relative_to(root).as_posix()
        except ValueError:
            relative_text = candidate.as_posix()

        folder_name = candidate.name

        prefix_match = any(
            folder_name.startswith(prefix)
            for prefix in prefixes
        )

        contains_match = any(
            text in folder_name or text in relative_text
            for text in contains_strings
        )

        # 沒有提供任何篩選條件時，選取所有 case。
        select_all = (
            not explicit_folders
            and not has_name_filters
        )

        if prefix_match or contains_match or select_all:
            resolved = candidate.resolve()
            found[resolved] = resolved

    return sorted(
        found.values(),
        key=natural_sort_key,
    )


# ============================================================
# Legend formatting
# ============================================================

def format_scalar(value: float) -> str:
    """
    適合 Landau 係數的簡短科學記號。
    """

    if value == 0:
        return "0"

    abs_value = abs(value)

    if abs_value >= 1.0e4 or abs_value < 1.0e-3:
        return f"{value:.3e}"

    return f"{value:.5g}"


def format_parameter_value(
    parameter_name: str,
    value: Any,
) -> str:
    if parameter_name == "t_FE":
        return f"{float(value):.4g} nm"

    if isinstance(value, list):
        formatted = ", ".join(
            format_scalar(float(item))
            for item in value
        )
        return f"[{formatted}]"

    return format_scalar(float(value))


def make_legend_label(
    case_dir: Path,
    root: Path,
    parameters: dict[str, Any],
    legend_parameters: list[str],
    use_relative_name: bool,
) -> str:
    """
    圖例第一行為 case 名稱，第二行為參數。
    """

    if use_relative_name:
        try:
            case_name = case_dir.relative_to(root).as_posix()
        except ValueError:
            case_name = case_dir.name
    else:
        case_name = case_dir.name

    displayed_parameters = dict(parameters)

    t_fe_nm = calculate_t_fe_nm(parameters)

    if t_fe_nm is not None:
        displayed_parameters["t_FE"] = t_fe_nm

    pieces = []

    for parameter_name in legend_parameters:
        if parameter_name not in displayed_parameters:
            continue

        display_name = PARAMETER_DISPLAY_NAMES.get(
            parameter_name,
            parameter_name,
        )

        value_text = format_parameter_value(
            parameter_name,
            displayed_parameters[parameter_name],
        )

        pieces.append(
            f"{display_name}={value_text}"
        )

    if not pieces:
        return case_name

    return case_name + "\n" + ", ".join(pieces)


# ============================================================
# Plotting
# ============================================================

def plot_overlay(
    case_directories: list[Path],
    root: Path,
    csv_relative_path: Path,
    inputs_filename: str,
    legend_parameters: list[str],
    use_csv_voltage: bool,
    use_relative_name: bool,
    output_path: Path,
    title: str,
    show_markers: bool,
    show_plot: bool,
    dpi: int,
) -> None:
    figure, axis = plt.subplots(
        figsize=(10, 6.5),
        constrained_layout=True,
    )

    number_plotted = 0

    for case_dir in case_directories:
        csv_path = case_dir / csv_relative_path
        inputs_path = case_dir / inputs_filename

        try:
            voltage, polarization = read_pv_curve(
                csv_path=csv_path,
                use_csv_voltage=use_csv_voltage,
            )
            '''
            parameters = read_inputs_parameters(
                inputs_path,
            )
            '''
            parameters = read_experiment_parameters(
                folder_name=case_dir.name,
                dataset_path="./MFIS_dataset.xlsx",
            )
            #print(f"[INFO] {case_dir.name}: Read parameters: {parameters}")
            label = make_legend_label(
                case_dir=case_dir,
                root=root,
                parameters=parameters,
                legend_parameters=legend_parameters,
                use_relative_name=use_relative_name,
            )

            plot_kwargs = {
                "linewidth": 1.7,
                "alpha": 0.82,
                "label": label,
            }

            if show_markers:
                plot_kwargs.update(
                    {
                        "marker": "o",
                        "markersize": 3.0,
                    }
                )

            axis.plot(
                voltage,
                polarization,
                **plot_kwargs,
            )

            number_plotted += 1

            print(
                f"[OK] {case_dir.name}: "
                f"{len(polarization)} points"
            )

        except Exception as error:
            print(
                f"[SKIP] {case_dir}: {error}",
                file=sys.stderr,
            )

    if number_plotted == 0:
        plt.close(figure)
        raise RuntimeError(
            "No valid P–V curves were plotted."
        )

    axis.set_xlabel("Gate voltage, $V_g$ (V)")
    axis.set_ylabel(
        "Mean ferroelectric polarization, "
        "$P_{\\mathrm{mean}}$"
    )
    axis.set_title(title)

    axis.axhline(
        0.0,
        linewidth=0.8,
        alpha=0.4,
    )
    axis.axvline(
        0.0,
        linewidth=0.8,
        alpha=0.4,
    )

    axis.grid(
        True,
        alpha=0.25,
    )

    # 將圖例放在圖外右側，避免遮住 P–V loop。
    axis.legend(
        loc="upper left",
        bbox_to_anchor=(1.02, 1.0),
        borderaxespad=0.0,
        fontsize=8,
        frameon=True,
    )

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    figure.savefig(
        output_path,
        dpi=dpi,
        bbox_inches="tight",
    )

    print()
    print(f"Plotted curves : {number_plotted}")
    print(f"Saved figure   : {output_path.resolve()}")

    if show_plot:
        plt.show()
    else:
        plt.close(figure)


# ============================================================
# Command line interface
# ============================================================

def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Overlay P_mean curves from multiple "
            "figs/MFIS_PV_curve.csv files."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )

    parser.add_argument(
        "--root",
        type=Path,
        default=Path("."),
        help="搜尋 case 資料夾的根目錄。",
    )

    parser.add_argument(
        "--folders",
        nargs="*",
        default=[],
        help=(
            "明確指定 case 資料夾。相對路徑以 --root 為基準。"
        ),
    )

    parser.add_argument(
        "--prefix",
        action="append",
        default=[],
        help=(
            "選取資料夾名稱以此字串開頭的 case。"
            "需要多個 prefix 時可重複使用此參數。"
        ),
    )

    parser.add_argument(
        "--contains",
        action="append",
        default=[],
        help=(
            "選取資料夾名稱或相對路徑包含此字串的 case。"
            "需要多個字串時可重複使用此參數。"
        ),
    )

    parser.add_argument(
        "--recursive",
        action="store_true",
        help="遞迴搜尋 --root 下的所有子資料夾。",
    )

    parser.add_argument(
        "--csv-relative-path",
        type=Path,
        default=DEFAULT_CSV_RELATIVE_PATH,
        help="case 資料夾內 P–V CSV 的相對路徑。",
    )

    parser.add_argument(
        "--inputs-filename",
        default=DEFAULT_INPUTS_FILENAME,
        help="case 資料夾內的 inputs 檔名。",
    )

    parser.add_argument(
        "--legend-params",
        nargs="+",
        default=DEFAULT_LEGEND_PARAMETERS,
        choices=[
            "alpha",
            "beta",
            "gamma",
            "BigGamma",
            "g11",
            "g44",
            "FE_lo",
            "FE_hi",
            "t_FE",
        ],
        help="顯示在圖例中的參數。",
    )

    parser.add_argument(
        "--use-csv-voltage",
        action="store_true",
        help=(
            "使用 CSV 的 Vg_mean，而不是預設的 37 點電壓序列。"
        ),
    )

    parser.add_argument(
        "--relative-name",
        action="store_true",
        help=(
            "圖例顯示相對於 --root 的路徑，"
            "而不只顯示資料夾名稱。"
        ),
    )

    parser.add_argument(
        "--markers",
        action="store_true",
        help="在每個電壓點加上 marker。",
    )

    parser.add_argument(
        "--output",
        type=Path,
        default=Path("pv_overlay.png"),
        help="輸出圖片路徑。",
    )

    parser.add_argument(
        "--title",
        default="Overlay of MFIS P–V curves",
        help="圖片標題。",
    )

    parser.add_argument(
        "--dpi",
        type=int,
        default=200,
        help="輸出圖片 DPI。",
    )

    parser.add_argument(
        "--no-show",
        action="store_true",
        help="只存檔，不顯示 Matplotlib 視窗。",
    )

    return parser.parse_args()


def main() -> None:
    args = parse_arguments()

    root = args.root.expanduser().resolve()

    if not root.is_dir():
        raise NotADirectoryError(
            f"Root directory does not exist: {root}"
        )

    case_directories = discover_case_directories(
        root=root,
        explicit_folders=args.folders,
        prefixes=args.prefix,
        contains_strings=args.contains,
        csv_relative_path=args.csv_relative_path,
        recursive=args.recursive,
    )

    if not case_directories:
        raise FileNotFoundError(
            "No matching case folders containing "
            f"'{args.csv_relative_path}' were found."
        )

    print(f"Root directory : {root}")
    print(f"Matched cases  : {len(case_directories)}")

    for case_dir in case_directories:
        try:
            display_path = case_dir.relative_to(root)
        except ValueError:
            display_path = case_dir

        print(f"  - {display_path}")

    print()

    plot_overlay(
        case_directories=case_directories,
        root=root,
        csv_relative_path=args.csv_relative_path,
        inputs_filename=args.inputs_filename,
        legend_parameters=args.legend_params,
        use_csv_voltage=args.use_csv_voltage,
        use_relative_name=args.relative_name,
        output_path=args.output,
        title=args.title,
        show_markers=args.markers,
        show_plot=not args.no_show,
        dpi=args.dpi,
    )


if __name__ == "__main__":
    main()
