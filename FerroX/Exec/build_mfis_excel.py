#!/usr/bin/env python3
"""
Build an Excel dataset index for FerroX / MFIS simulations.

Output design:
  1. summary          : generation metadata
  2. experiments      : one row per MFIS simulation folder
  3. pv_curve         : long table, one row per voltage point
  4. pz_stack_index   : one row per steady-state Pz slice; actual 2D arrays are saved in compressed .npz files
  5. warnings         : missing files / mismatch messages

Default folder layout expected:
  ./MFIS*/inputs
  ./MFIS*/MFIS_PV_curve.csv or somewhere below MFIS*/
  ./MFIS*/plts/plt########
  ./MFIS*/run.log or ./MFIS*/plts/run.log

Dependencies:
  pip install numpy pandas openpyxl yt
"""

from __future__ import annotations

from pathlib import Path
from datetime import datetime, timedelta
import argparse
import re
import sys

import numpy as np
import pandas as pd


# ==========================
# User settings / defaults
# ==========================

WANTED_PARAMS = [
    "alpha",
    "beta",
    "gamma",
    "BigGamma",
    "g11",
    "g44",
    "FE_lo",
    "FE_hi",
]

P_FIELD = ("boxlib", "Pz")
PHI_FIELD = ("boxlib", "Phi")

# Keep the same convention as your plotting code: P *= -1
P_SIGN = -1.0

# y slice: None means y center
Y_INDEX = None

# If True, only keep FE x range based on FE_lo[0], FE_hi[0].
# If False, keep full x range.
USE_FE_X_RANGE = False

# Plotfile filtering, copied from your Pz stack code
SKIP_INITIAL = True
SKIP_FIRST_N = 9

# Applied gate sweep used if V_applied is not already in MFIS_PV_curve.csv
DEFAULT_VMIN = -4.5
DEFAULT_VMAX = 4.5
DEFAULT_DV = 0.5


# ==========================
# Basic helpers
# ==========================

def parse_value(values: str):
    """Parse FerroX input values like '-8e9' or '0 0 8e-9'."""
    values = values.split("#", 1)[0].strip()
    vals = [float(v) for v in values.split()]
    return vals[0] if len(vals) == 1 else vals


def sci_or_blank(v):
    if v is None or (isinstance(v, float) and not np.isfinite(v)):
        return ""
    return f"{float(v):.6e}"


def get_component(v, idx: int):
    if isinstance(v, (list, tuple)) and len(v) > idx:
        return v[idx]
    if idx == 0 and isinstance(v, (int, float)):
        return v
    return None


def get_step_from_name(name) -> int | None:
    m = re.search(r"plt(\d+)$", Path(name).name)
    return int(m.group(1)) if m else None


def find_plot_names(plot_dir: Path) -> list[str]:
    if not plot_dir.exists():
        return []
    names = [p.name for p in plot_dir.iterdir() if p.is_dir() and re.match(r"plt\d+$", p.name)]
    return sorted(names, key=lambda n: get_step_from_name(n) if get_step_from_name(n) is not None else 10**99)


def find_first_file(folder: Path, filename: str) -> Path | None:
    direct = folder / filename
    if direct.exists():
        return direct
    matches = sorted(folder.rglob(filename))
    return matches[0] if matches else None


def build_default_sweep(vmin=DEFAULT_VMIN, vmax=DEFAULT_VMAX, dv=DEFAULT_DV):
    # Example: -4.5, -4.0, ..., 4.5, 4.0, ..., -4.5 => 37 points
    up = np.arange(vmin, vmax + 0.5 * dv, dv)
    down = np.arange(vmax - dv, vmin - 0.5 * dv, -dv)
    values = np.round(np.concatenate([up, down]), 12)
    branches = ["positive_sweep"] * len(up) + ["negative_sweep"] * len(down)
    return values, branches


def normalize_col(c) -> str:
    return re.sub(r"[^a-z0-9]", "", str(c).strip().lower())


def find_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Robustly find a column by normalized names."""
    norm_map = {normalize_col(c): c for c in df.columns}
    norm_candidates = [normalize_col(c) for c in candidates]

    # exact match first
    for cand in norm_candidates:
        if cand in norm_map:
            return norm_map[cand]

    # then partial match, but avoid too-short ambiguous candidates first
    for cand in norm_candidates:
        if len(cand) < 2:
            continue
        for norm, raw in norm_map.items():
            if cand in norm:
                return raw

    return None


# ==========================
# Read original metadata
# ==========================

def read_inputs(path: Path) -> dict:
    params = {}
    with open(path, errors="ignore") as f:
        for line in f:
            line = line.strip()
            if not line or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            if key not in WANTED_PARAMS:
                continue
            try:
                params[key] = parse_value(value)
            except Exception:
                params[key] = value.split("#", 1)[0].strip()
    return params


def parse_date_line(line: str) -> datetime | None:
    """Parse date output like 'Fri Jun 19 00:36:36 CST 2026'.

    Python's %Z parsing for CST is system-dependent, so remove the timezone token.
    """
    line = line.strip()
    m = re.match(
        r"^(?P<dow>\w{3})\s+(?P<mon>\w{3})\s+(?P<day>\d{1,2})\s+"
        r"(?P<hms>\d{2}:\d{2}:\d{2})\s+(?P<tz>\S+)\s+(?P<year>\d{4})$",
        line,
    )
    if not m:
        return None
    s = f"{m.group('dow')} {m.group('mon')} {m.group('day')} {m.group('hms')} {m.group('year')}"
    try:
        return datetime.strptime(s, "%a %b %d %H:%M:%S %Y")
    except Exception:
        return None


def read_run_log(folder: Path):
    logfile = folder / "run.log"
    if not logfile.exists():
        logfile = folder / "plts" / "run.log"
    if not logfile.exists():
        return None, None, None, None

    with open(logfile, errors="ignore") as f:
        lines = [line.strip() for line in f if line.strip()]

    # Keep behavior close to your original script
    lines_tail = lines[-200:]

    end_time = None
    start_time = None
    elapsed_hms = None
    elapsed_sec = None

    if lines_tail:
        end_time = parse_date_line(lines_tail[-1])

    for line in lines_tail:
        if line.startswith("Total run time"):
            m = re.search(r"([\d.]+)", line)
            if m:
                elapsed_sec = float(m.group(1))
            break

    if elapsed_sec is not None:
        if end_time is not None:
            start_time = end_time - timedelta(seconds=elapsed_sec)

        h = int(elapsed_sec // 3600)
        m = int((elapsed_sec % 3600) // 60)
        s = int(elapsed_sec % 60)
        elapsed_hms = f"{h:02d}:{m:02d}:{s:02d}"

    return start_time, end_time, elapsed_hms, elapsed_sec


# ==========================
# Read P-V CSV
# ==========================

def read_pv_curve(folder: Path, run_id: str, sweep_values: np.ndarray, sweep_branches: list[str]):
    csv_path = find_first_file(folder, "MFIS_PV_curve.csv")
    if csv_path is None:
        return None, None, f"{run_id}: MFIS_PV_curve.csv not found"

    try:
        df_raw = pd.read_csv(csv_path)
    except Exception as e:
        return None, str(csv_path), f"{run_id}: failed to read {csv_path}: {e}"

    p_col = find_column(df_raw, ["P_mean", "Pavg", "P_average", "P"])
    vg_col = find_column(df_raw, ["Vg_mean", "Vtop_mean", "V_mean", "Vg", "V"])
    vapplied_col = find_column(df_raw, ["V_applied", "Vapp", "V_app", "gate_voltage", "Vg_applied"])
    step_col = find_column(df_raw, ["step_id", "step", "index"])

    if p_col is None or vg_col is None:
        return None, str(csv_path), (
            f"{run_id}: cannot identify P_mean/Vg_mean columns in {csv_path}. "
            f"Columns = {list(df_raw.columns)}"
        )

    n = len(df_raw)
    out = pd.DataFrame(index=np.arange(n))

    #out["run_id"] = run_id
    out["folder"] = folder.name
    out["point_id"] = np.arange(n, dtype=int)
    #out["step_id"] = df_raw[step_col].values if step_col else np.arange(n, dtype=int)

    if vapplied_col is not None:
        out["V_applied"] = pd.to_numeric(df_raw[vapplied_col], errors="coerce")
        out["branch"] = infer_branch_from_voltage(out["V_applied"].to_numpy())
    elif n == len(sweep_values):
        out["V_applied"] = sweep_values
        out["branch"] = sweep_branches
    else:
        out["V_applied"] = np.nan
        out["branch"] = "unknown"

    #out["Vg_mean"] = pd.to_numeric(df_raw[vg_col], errors="coerce")
    out["P_mean"] = pd.to_numeric(df_raw[p_col], errors="coerce")

    # Keep optional statistics if your CSV already contains them
    optional = {
        "P_min": ["P_min", "Pmin"],
        "P_max": ["P_max", "Pmax"],
        "P_std": ["P_std", "Pstd"],
        "Vg_min": ["Vg_min", "Vmin"],
        "Vg_max": ["Vg_max", "Vmax"],
        "Vg_std": ["Vg_std", "Vstd"],
    }
    
    '''
    
    for out_col, candidates in optional.items():
        c = find_column(df_raw, candidates)
        if c is not None:
            out[out_col] = pd.to_numeric(df_raw[c], errors="coerce")
            
    '''

    #out["pv_csv_path"] = str(csv_path)
    return out, str(csv_path), None


def infer_branch_from_voltage(v: np.ndarray):
    labels = []
    for i, x in enumerate(v):
        if i == 0 or not np.isfinite(x) or not np.isfinite(v[i - 1]):
            labels.append("start")
            continue
        dv = x - v[i - 1]
        if dv > 1e-10:
            labels.append("positive_sweep")
        elif dv < -1e-10:
            labels.append("negative_sweep")
        else:
            labels.append("same_V")
    return labels


# ==========================
# Read Pz steady-state slices from plotfiles
# ==========================

def to_numpy(field_data):
    try:
        return field_data.to_ndarray()
    except AttributeError:
        try:
            return field_data.v
        except AttributeError:
            return np.asarray(field_data)


def read_full_field_from_grids(ds, field):
    """Read a full-domain field.

    First try yt.covering_grid. If that fails, merge individual grids using their
    global start indices. This is intended for level-0 uniform AMReX/BoxLib data.
    """
    dims = tuple(int(x) for x in ds.domain_dimensions)

    try:
        cg = ds.covering_grid(level=0, left_edge=ds.domain_left_edge, dims=dims)
        return to_numpy(cg[field])
    except Exception:
        arr = np.full(dims, np.nan, dtype=np.float64)
        for g in ds.index.grids:
            data = to_numpy(g[field])
            start = np.array(g.get_global_startindex(), dtype=int)
            end = start + np.array(data.shape, dtype=int)
            arr[start[0]:end[0], start[1]:end[1], start[2]:end[2]] = data
        return arr


def read_one_plot(plotfile: Path, params: dict, yt_module, y_index=None):
    ds = yt_module.load(str(plotfile))

    if P_FIELD not in ds.field_list:
        raise RuntimeError(f"{plotfile}: field {P_FIELD} not found. Available fields: {ds.field_list}")

    P = read_full_field_from_grids(ds, P_FIELD).astype(np.float32)
    P *= P_SIGN

    if PHI_FIELD in ds.field_list:
        Phi = read_full_field_from_grids(ds, PHI_FIELD).astype(np.float32)
    else:
        Phi = None

    Nx, Ny, Nz = P.shape

    lo = np.asarray(ds.domain_left_edge.to_value(), dtype=float)
    hi = np.asarray(ds.domain_right_edge.to_value(), dtype=float)
    dx = (hi - lo) / np.array([Nx, Ny, Nz], dtype=float)

    x_nm = (lo[0] + (np.arange(Nx) + 0.5) * dx[0]) * 1e9
    y_nm = (lo[1] + (np.arange(Ny) + 0.5) * dx[1]) * 1e9
    z_nm = (lo[2] + (np.arange(Nz) + 0.5) * dx[2]) * 1e9

    if y_index is None:
        y_index_use = Ny // 2
    else:
        y_index_use = int(y_index)

    fe_lo = params.get("FE_lo")
    fe_hi = params.get("FE_hi")

    fe_z_lo = get_component(fe_lo, 2)
    fe_z_hi = get_component(fe_hi, 2)
    fe_x_lo = get_component(fe_lo, 0)
    fe_x_hi = get_component(fe_hi, 0)

    if fe_z_lo is not None and fe_z_hi is not None:
        z_lo_nm = min(fe_z_lo, fe_z_hi) * 1e9
        z_hi_nm = max(fe_z_lo, fe_z_hi) * 1e9
        z_sel = (z_nm >= z_lo_nm) & (z_nm <= z_hi_nm)
    else:
        z_sel = np.ones_like(z_nm, dtype=bool)

    if USE_FE_X_RANGE and fe_x_lo is not None and fe_x_hi is not None:
        x_lo_nm = min(fe_x_lo, fe_x_hi) * 1e9
        x_hi_nm = max(fe_x_lo, fe_x_hi) * 1e9
        x_sel = (x_nm >= x_lo_nm) & (x_nm <= x_hi_nm)
    else:
        x_sel = np.ones_like(x_nm, dtype=bool)

    P_slice = P[np.ix_(x_sel, [y_index_use], z_sel)][:, 0, :]

    if Phi is not None:
        top_phi = Phi[:, :, -1]
        Vtop_mean = float(np.nanmean(top_phi))
        Vtop_min = float(np.nanmin(top_phi))
        Vtop_max = float(np.nanmax(top_phi))
        Vtop_std = float(np.nanstd(top_phi))
    else:
        Vtop_mean = np.nan
        Vtop_min = np.nan
        Vtop_max = np.nan
        Vtop_std = np.nan

    return {
        "P_slice": P_slice,
        "x_nm": x_nm[x_sel],
        "z_nm": z_nm[z_sel],
        "y_index": y_index_use,
        "y_nm": float(y_nm[y_index_use]),
        "Vtop_mean_from_plot": Vtop_mean,
        "Vtop_min_from_plot": Vtop_min,
        "Vtop_max_from_plot": Vtop_max,
        "Vtop_std_from_plot": Vtop_std,
        "full_shape": tuple(int(x) for x in P.shape),
        "slice_shape": tuple(int(x) for x in P_slice.shape),
    }


def extract_pz_stacks(
    folder: Path,
    run_id: str,
    params: dict,
    pv_df: pd.DataFrame | None,
    pz_dir: Path,
    y_index=None,
    skip_initial=True,
    skip_first_n=9,
):
    try:
        import yt  # imported only when extraction is enabled
    except Exception as e:
        raise RuntimeError("yt is required for Pz extraction. Install with: pip install yt") from e

    plot_dir = folder / "plts"
    if not plot_dir.exists():
        plot_dir = folder

    plot_names = find_plot_names(plot_dir)

    if skip_initial:
        plot_names = [n for n in plot_names if get_step_from_name(n) != 0]

    if skip_first_n > 0:
        plot_names = plot_names[skip_first_n:]

    if not plot_names:
        return [], None, f"{run_id}: no steady plotfiles found after filtering"

    out_folder = pz_dir / run_id
    out_folder.mkdir(parents=True, exist_ok=True)
    out_npz = out_folder / "Pz_FE_all_voltage.npz"

    npz_payload = {}
    index_rows = []
    x_nm_ref = None
    z_nm_ref = None

    for i, name in enumerate(plot_names):
        plotfile = plot_dir / name
        info = read_one_plot(plotfile, params=params, yt_module=yt, y_index=y_index)

        key = f"Pz_{i:03d}"
        P_slice = info["P_slice"].astype(np.float32)
        npz_payload[key] = P_slice

        if x_nm_ref is None:
            x_nm_ref = info["x_nm"]
            z_nm_ref = info["z_nm"]

        if pv_df is not None and i < len(pv_df):
            pv_row = pv_df.iloc[i]
            V_applied = pv_row.get("V_applied", np.nan)
            Vg_mean_csv = pv_row.get("Vg_mean", np.nan)
            P_mean_csv = pv_row.get("P_mean", np.nan)
            branch = pv_row.get("branch", "")
            point_id = int(pv_row.get("point_id", i))
        else:
            V_applied = np.nan
            Vg_mean_csv = np.nan
            P_mean_csv = np.nan
            branch = "unknown"
            point_id = i

        index_rows.append({
            "run_id": run_id,
            "folder": folder.name,
            "point_id": point_id,
            "plot_index": i,
            "plotfile": str(plotfile),
            "plot_step": get_step_from_name(name),
            "branch": branch,
            "V_applied": V_applied,
            "Vg_mean_from_csv": Vg_mean_csv,
            "P_mean_from_csv": P_mean_csv,
            "Vtop_mean_from_plot": info["Vtop_mean_from_plot"],
            "Vtop_min_from_plot": info["Vtop_min_from_plot"],
            "Vtop_max_from_plot": info["Vtop_max_from_plot"],
            "Vtop_std_from_plot": info["Vtop_std_from_plot"],
            "Pz_mean_from_slice": float(np.nanmean(P_slice)),
            "Pz_min_from_slice": float(np.nanmin(P_slice)),
            "Pz_max_from_slice": float(np.nanmax(P_slice)),
            "Pz_std_from_slice": float(np.nanstd(P_slice)),
            "y_index": info["y_index"],
            "y_nm": info["y_nm"],
            "full_shape": str(info["full_shape"]),
            "slice_shape_Nx_Nz": str(info["slice_shape"]),
            "npz_path": str(out_npz),
            "array_key": key,
            "x_nm_key": "x_nm",
            "z_nm_key": "z_nm",
        })

    # Save coordinate arrays and simple metadata arrays in the same file
    npz_payload["x_nm"] = np.asarray(x_nm_ref, dtype=np.float64)
    npz_payload["z_nm"] = np.asarray(z_nm_ref, dtype=np.float64)
    npz_payload["V_applied"] = np.asarray([r["V_applied"] for r in index_rows], dtype=np.float64)
    npz_payload["Vg_mean_from_csv"] = np.asarray([r["Vg_mean_from_csv"] for r in index_rows], dtype=np.float64)
    npz_payload["P_mean_from_csv"] = np.asarray([r["P_mean_from_csv"] for r in index_rows], dtype=np.float64)
    npz_payload["Vtop_mean_from_plot"] = np.asarray([r["Vtop_mean_from_plot"] for r in index_rows], dtype=np.float64)
    npz_payload["plot_step"] = np.asarray([r["plot_step"] if r["plot_step"] is not None else -1 for r in index_rows], dtype=np.int64)

    np.savez_compressed(out_npz, **npz_payload)

    warning = None
    if pv_df is not None and len(pv_df) != len(plot_names):
        warning = f"{run_id}: PV rows ({len(pv_df)}) != selected plotfiles ({len(plot_names)})"

    return index_rows, str(out_npz), warning


# ==========================
# Excel writing
# ==========================

def format_excel(out_xlsx: Path):
    """Light formatting using openpyxl after pandas writes the workbook."""
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import Font, PatternFill, Alignment
    except Exception:
        return

    wb = load_workbook(out_xlsx)

    header_fill = PatternFill("solid", fgColor="D9EAF7")
    header_font = Font(bold=True)

    for ws in wb.worksheets:
        if ws.max_row >= 1:
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

        # reasonable width cap
        for col_cells in ws.columns:
            col_letter = col_cells[0].column_letter
            max_len = 0
            for cell in col_cells[:200]:
                if cell.value is None:
                    continue
                max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 42)

        # number formats
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, float):
                    cell.number_format = "0.000E+00"
                elif isinstance(cell.value, datetime):
                    cell.number_format = "yyyy-mm-dd hh:mm:ss"

    wb.save(out_xlsx)


def write_workbook(
    out_xlsx: Path,
    summary_rows: list[dict],
    exp_rows: list[dict],
    pv_rows_all: list[pd.DataFrame],
    pz_index_rows: list[dict],
    warnings: list[str],
):
    summary_df = pd.DataFrame(summary_rows)
    exp_df = pd.DataFrame(exp_rows)
    pv_df = pd.concat(pv_rows_all, ignore_index=True) if pv_rows_all else pd.DataFrame()
    pz_df = pd.DataFrame(pz_index_rows)
    warn_df = pd.DataFrame({"warning": warnings}) if warnings else pd.DataFrame({"warning": []})

    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="summary", index=False)
        exp_df.to_excel(writer, sheet_name="experiments", index=False)
        pv_df.to_excel(writer, sheet_name="pv_curve", index=False)
        pz_df.to_excel(writer, sheet_name="pz_stack_index", index=False)
        warn_df.to_excel(writer, sheet_name="warnings", index=False)

    format_excel(out_xlsx)


# ==========================
# Main build
# ==========================

def build_dataset(args):
    root = Path(args.root).resolve()
    out_xlsx = Path(args.out).resolve()
    pz_dir = Path(args.pz_dir).resolve()
    pz_dir.mkdir(parents=True, exist_ok=True)

    sweep_values, sweep_branches = build_default_sweep(args.vmin, args.vmax, args.dv)

    exp_rows = []
    pv_rows_all = []
    pz_index_rows = []
    warnings = []

    folders = sorted([p for p in root.glob("MFIS*") if p.is_dir()])

    for folder in folders:
        run_id = folder.name
        inp = folder / "inputs"
        if not inp.exists():
            warnings.append(f"{run_id}: inputs not found; skipped")
            continue

        print(f"Processing {run_id}")
        params = read_inputs(inp)

        fe_lo_z = get_component(params.get("FE_lo"), 2)
        fe_hi_z = get_component(params.get("FE_hi"), 2)
        t_fe = fe_hi_z - fe_lo_z if fe_lo_z is not None and fe_hi_z is not None else np.nan

        start_time, end_time, elapsed_hms, elapsed_sec = read_run_log(folder)

        pv_df, pv_csv_path, pv_warning = read_pv_curve(folder, run_id, sweep_values, sweep_branches)
        if pv_warning:
            warnings.append(pv_warning)
        if pv_df is not None:
            pv_rows_all.append(pv_df)

        pz_npz_path = None
        n_pz_stacks = 0
        if args.extract_pz:
            try:
                rows, pz_npz_path, pz_warning = extract_pz_stacks(
                    folder=folder,
                    run_id=run_id,
                    params=params,
                    pv_df=pv_df,
                    pz_dir=pz_dir,
                    y_index=args.y_index,
                    skip_initial=not args.keep_initial,
                    skip_first_n=args.skip_first_n,
                )
                pz_index_rows.extend(rows)
                n_pz_stacks = len(rows)
                if pz_warning:
                    warnings.append(pz_warning)
            except Exception as e:
                warnings.append(f"{run_id}: Pz extraction failed: {e}")

        exp_row = {
            "run_id": run_id,
            "folder": folder.name,
            "inputs_path": str(inp),
            "pv_csv_path": pv_csv_path,
            "pz_npz_path": pz_npz_path,
            "n_pv_points": 0 if pv_df is None else len(pv_df),
            "n_pz_stacks": n_pz_stacks,
            "T_FE": t_fe,
            "T_FE_nm": t_fe * 1e9 if np.isfinite(t_fe) else np.nan,
            "Start Time": start_time,
            "End Time": end_time,
            "Elapsed": elapsed_hms,
            "Elapsed_sec": elapsed_sec,
        }

        for p in WANTED_PARAMS:
            if p in ["FE_lo", "FE_hi"]:
                v = params.get(p)
                exp_row[p] = " ".join(sci_or_blank(x) for x in v) if isinstance(v, list) else sci_or_blank(v)
            else:
                exp_row[p] = params.get(p, np.nan)

        exp_rows.append(exp_row)

    summary_rows = [
        {"item": "Last Update", "value": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
        {"item": "Root", "value": str(root)},
        {"item": "Total MFIS folders", "value": len(folders)},
        {"item": "Experiments written", "value": len(exp_rows)},
        {"item": "P-V points written", "value": sum(len(df) for df in pv_rows_all)},
        {"item": "Pz stacks indexed", "value": len(pz_index_rows)},
        {"item": "Default sweep", "value": f"{args.vmin} -> {args.vmax} -> {args.vmin}, dV={args.dv}"},
        {"item": "Pz storage", "value": "2D Pz arrays are saved in compressed .npz files; Excel stores npz_path + array_key."},
        {"item": "P sign convention", "value": f"Pz multiplied by {P_SIGN}, same as plotting code."},
    ]

    write_workbook(
        out_xlsx=out_xlsx,
        summary_rows=summary_rows,
        exp_rows=exp_rows,
        pv_rows_all=pv_rows_all,
        pz_index_rows=pz_index_rows,
        warnings=warnings,
    )

    print("Done")
    print(f"Excel saved: {out_xlsx}")
    print(f"Pz npz folder: {pz_dir}")
    if warnings:
        print(f"Warnings: {len(warnings)}; see warnings sheet")


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description="Build FerroX MFIS Excel dataset index")
    parser.add_argument("--root", default=".", help="Root folder containing MFIS* simulation folders")
    parser.add_argument("--out", default="MFIS_dataset.xlsx", help="Output Excel filename")
    parser.add_argument("--pz-dir", default="extracted_pz", help="Folder for compressed Pz .npz outputs")
    parser.add_argument("--no-pz", dest="extract_pz", action="store_false", help="Do not read plotfiles / do not extract Pz stacks")
    parser.set_defaults(extract_pz=True)
    parser.add_argument("--vmin", type=float, default=DEFAULT_VMIN)
    parser.add_argument("--vmax", type=float, default=DEFAULT_VMAX)
    parser.add_argument("--dv", type=float, default=DEFAULT_DV)
    parser.add_argument("--skip-first-n", type=int, default=SKIP_FIRST_N, help="Skip first N steady-state plotfiles after initial plt00000000 removal")
    parser.add_argument("--keep-initial", action="store_true", help="Do not remove plt00000000")
    parser.add_argument("--y-index", type=int, default=Y_INDEX, help="y index for Pz slice; default is center")
    return parser.parse_args(argv)


if __name__ == "__main__":
    build_dataset(parse_args())
