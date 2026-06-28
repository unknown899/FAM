#!/usr/bin/env python3

from pathlib import Path
import html
import re
from datetime import datetime, timedelta

def read_run_log(folder):
    logfile = folder / "run.log"
    if not logfile.exists():
        logfile = folder / "plts/run.log"
    if not logfile.exists():
        return None, None, None
    
    print(f"Reading run log: {logfile}")
    with open(logfile, errors="ignore") as f:
        lines = [line.strip() for line in f if line.strip()]

    # 只取最後 200 行
    lines = lines[-200:]

    end_time = None
    start_time = None
    elapsed_hms = None
    elapsed_sec = None

    # 最後一行 = End Time
    if lines:
        try:
            end_time_str = lines[-1]
            end_time = datetime.strptime(
                end_time_str,
                "%a %b %d %H:%M:%S %Z %Y"
            )
        except:
            pass

    # 找 Total run time
    for line in lines:
        if line.startswith("Total run time"):
            m = re.search(r'([\d.]+)', line)
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
    print(f"Start Time: {start_time}, End Time: {end_time}, Elapsed: {elapsed_hms}")
    return start_time, end_time, elapsed_hms


wanted = [
    "alpha",
    "beta",
    "gamma",
    "BigGamma",
    "g11",
    "g44",
    "FE_lo",
    "FE_hi",
]

IMAGE_PATTERNS = [
    "*.png",
    "*.jpg",
    "*.jpeg",
    "*.svg",
]

IMAGE_FILES = {
    "PV Curve": "MFIS_PV_curve.png",
    "Pz Stack": "Pz_FE_layer_stack.png",
}


def parse_value(values):
    vals = [float(v) for v in values.split()]
    return vals[0] if len(vals) == 1 else vals


def sci(v):
    if v is None:
        return ""
    return f"{float(v):.3e}"


def read_inputs(path):
    params = {}

    with open(path) as f:
        for line in f:

            line = line.strip()

            if not line:
                continue

            if "=" not in line:
                continue

            key, value = line.split("=", 1)
            key = key.strip()

            if key not in wanted:
                continue

            params[key] = parse_value(value)

    return params


def find_images(folder):
    images = {
        "PV": None,
        "Pz": None,
    }

    pv = list(folder.rglob("MFIS_PV_curve.png"))
    if pv:
        images["PV"] = pv[0]

    pz = list(folder.rglob("Pz_FE_layer_stack.png"))
    if pz:
        images["Pz"] = pz[0]

    return images


rows = []

for folder in sorted(Path(".").glob("MFIS*")):

    if not folder.is_dir():
        continue

    inp = folder / "inputs"

    if not inp.exists():
        continue

    params = read_inputs(inp)

    fe_lo = params.get("FE_lo")
    fe_lo = fe_lo[2]

    fe_hi = params.get("FE_hi")
    fe_hi = fe_hi[2]

    t_fe = None
    if fe_lo is not None and fe_hi is not None:
        t_fe = fe_hi - fe_lo

    imgs = find_images(folder)

    # 找 MFIS*/run.log
    start_time, end_time, elapsed_hms = read_run_log(
        folder
    )

    rows.append({
        "folder": folder.name,
        "params": params,
        "tfe": t_fe,

        "Start Time": start_time,
        "End Time": end_time,
        "Elapsed": elapsed_hms,
        
        "PV Curve": imgs["PV"],
        "Pz Stack": imgs["Pz"],
    })

from datetime import datetime

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
<span id="total_exp">27</span>
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

for p in wanted:
    if p != "FE_lo" and p != "FE_hi":
        html_text += f"<th>{p}</th>\n"
    


html_text += "<th>T_FE</th>"
html_text += "<th>Start Time</th>"
html_text += "<th>End Time</th>"
html_text += "<th>Elapsed</th>"
for title in IMAGE_FILES:
    html_text += f"<th>{title}</th>"

html_text += """
</tr>

</thead>

<tbody id="exp_table">
"""

for r in rows:

    html_text += "<tr>"

    html_text += f"<td>{html.escape(r['folder'])}</td>"

    for p in wanted:
        if p == "FE_lo" or p == "FE_hi":
            continue
        value = r["params"].get(p)

        if isinstance(value, list):
            txt = " ".join(sci(v) for v in value)
        elif value is None:
            txt = ""
        else:
            txt = sci(value)

        html_text += f"<td>{txt}</td>"

    html_text += f"<td>{sci(r['tfe'])}</td>"
    html_text += f"<td>{r['Start Time']}</td>"
    html_text += f"<td>{r['End Time']}</td>"
    html_text += f"<td>{r['Elapsed']}</td>"

    for key in IMAGE_FILES:
        img = r[key]

        if img:
            rel = img.as_posix()

            html_text += f"""
<td>
<a href="{rel}" target="_blank">
<img src="{rel}">
</a>
</td>
"""
        else:
            html_text += "<td>No Image</td>"

    html_text += "</tr>"

html_text += """

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
        : filter.split(/\s+/).map(str => {
            const m = str.match(/^(\w+)(<=|>=|=|<|>)(.+)$/);
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

        // 每列資料
        const row = {
            folder : cells[0].textContent.trim(),
            alpha  : Number(cells[1].textContent),
            beta   : Number(cells[2].textContent),
            gamma  : Number(cells[3].textContent),
            g11    : Number(cells[4].textContent),
            t_DE   : Number(cells[5].textContent),
            t_OX   : Number(cells[6].textContent),
            T_FE   : Number(cells[7].textContent),
        };

        // 沒輸入任何條件 → 全部顯示
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

with open("index.html", "w", encoding="utf8") as f:
    f.write(html_text)

print(f"Generated index.html ({len(rows)} experiments)")