#!/usr/bin/env python3

from pathlib import Path
import html

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
    fe_lo=fe_lo[2]
    fe_hi = params.get("FE_hi")
    fe_hi=fe_hi[2]

    t_fe = None
    if fe_lo is not None and fe_hi is not None:
        t_fe = (fe_hi - fe_lo)

    imgs = find_images(folder)

    rows.append({
    "folder": folder.name,
    "params": params,
    "tfe": t_fe,
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
<span id="total_exp">{len(rows)}</span>
</div>
</p>

<p class="note">
* All parameters are in SI units.
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

const search=document.getElementById("search");

search.onkeyup=function(){

let filter=this.value.toLowerCase();

let rows=document.querySelectorAll("#tbl tbody tr");

rows.forEach(function(r){

let txt=r.innerText.toLowerCase();

r.style.display=txt.includes(filter)?"":"none";

});

};

</script>

</body>

</html>

"""

with open("index.html", "w", encoding="utf8") as f:
    f.write(html_text)

print(f"Generated index.html ({len(rows)} experiments)")