from flask import Flask, request, render_template_string, jsonify, send_file
import io
import math
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

app = Flask(__name__)

# =====================================================
# ⚙️ CONFIG (INDUSTRY STANDARD)
# =====================================================
SCALE = 10   # 0.1 mm precision

# =====================================================
# 🔥 KNAPSACK OPTIMIZER
# =====================================================
def best_combination(sizes, max_width, max_use):
    dp = {0: [0]*len(sizes)}

    for i, size in enumerate(sizes):
        new_dp = dp.copy()

        for width, counts in dp.items():
            for k in range(1, max_use[size] + 1):
                new_width = width + size*k
                if new_width > max_width:
                    break

                new_counts = counts.copy()
                new_counts[i] += k

                if new_width not in new_dp:
                    new_dp[new_width] = new_counts

        dp = new_dp

    best_width = max(dp.keys())
    return best_width, dp[best_width]

# =====================================================
# 🚀 MAIN PLANNER
# =====================================================
def plan_multi_coil(order, master_width, coil_weight, tolerance, min_util):

    # 🔹 SCALE INPUTS
    order = [(int(round(s*SCALE)), w) for s, w in order]
    master_width_scaled = int(round(master_width*SCALE))

    # 🔹 Weight per mm (REAL mm)
    weight_per_mm = (coil_weight*1000) / master_width

    # 🔹 Convert demand → slits
    demand = {}
    weight_per_slit_map = {}

    for size, wt in order:
        real_size = size / SCALE
        wps = (real_size * weight_per_mm) / 1000
        slits = max(1, math.ceil((wt - tolerance) / wps))

        demand[size] = slits
        weight_per_slit_map[size] = wps

    sizes = sorted(demand.keys())
    plans = []

    while any(v > 0 for v in demand.values()):

        base = [0]*len(sizes)
        used_width = 0

        # 🔹 Mandatory 1 slit each
        for i, size in enumerate(sizes):
            if demand[size] > 0:
                base[i] += 1
                demand[size] -= 1
                used_width += size

        remaining = master_width_scaled - used_width

        max_use = {s: demand[s]+2 for s in sizes}

        best_w, combo = best_combination(sizes, remaining, max_use)

        final = [base[i] + combo[i] for i in range(len(sizes))]

        # 🔹 Reduce demand (ONLY combo)
        for i, size in enumerate(sizes):
            demand[size] -= combo[i]
            if demand[size] < 0:
                demand[size] = 0

        # 🔹 Build coil
        coil = []
        total_weight = 0
        used_scaled = 0

        for i, size in enumerate(sizes):
            slits = final[i]
            if slits == 0:
                continue

            width_scaled = slits * size
            real_size = size / SCALE
            real_width = width_scaled / SCALE

            wps = weight_per_slit_map[size]
            total = slits * wps

            total_weight += total
            used_scaled += width_scaled

            coil.append({
                "size": round(real_size,2),
                "slits": slits,
                "width": round(real_width,2),
                "weight_per_mm": round(weight_per_mm,2),
                "weight_per_slit": round(wps,3),
                "total_weight": round(total,3)
            })

        used = used_scaled / SCALE
        remaining = (master_width_scaled - used_scaled) / SCALE
        util = used_scaled / master_width_scaled

        plans.append({
            "coil": sorted(coil, key=lambda x: x["size"]),
            "used_width": round(used,2),
            "remaining_width": round(remaining,2),
            "utilization": round(util,4),
            "total_weight": round(total_weight,3)
        })

    return plans

# =====================================================
# 🎨 FRONTEND
# =====================================================
HTML = """
<!DOCTYPE html>
<html>
<head>
<title>Coil Optimizer</title>
<style>
body{font-family:Segoe UI;background:#f8fafc;padding:20px;}
.card{background:white;padding:20px;border-radius:10px;margin-bottom:20px;}
input{padding:10px;margin:5px;}
button{padding:10px;margin:5px;}
table{width:100%;margin-top:10px;border-collapse:collapse;}
th,td{padding:8px;border-bottom:1px solid #ddd;text-align:center;}
.total{font-weight:bold;color:green;}
</style>
</head>

<body>
<h2>Coil Optimization</h2>

<div class="card">
<input id="mw" placeholder="Master Width (mm)">
<input id="cw" placeholder="Coil Weight (MT)">
<input id="tol" placeholder="Tolerance (Kg)">
<input id="util" placeholder="Min Util %">
</div>

<div class="card">
<table id="tbl">
<tr><th>Size (mm)</th><th>Weight (MT)</th><th></th></tr>
</table>
<button onclick="add()">+ Add</button>
</div>

<button onclick="run()">Generate</button>

<div id="out"></div>

<script>
function add(){
 let r=document.getElementById("tbl").insertRow();
 r.innerHTML=`<td><input></td><td><input></td><td><button onclick="this.parentElement.parentElement.remove()">X</button></td>`;
}
add(); add();

async function run(){
 let rows=document.querySelectorAll("#tbl tr");
 let order=[];
 rows.forEach((r,i)=>{
  if(i==0)return;
  let i1=r.children[0].querySelector("input").value;
  let i2=r.children[1].querySelector("input").value;
  if(i1&&i2)order.push([parseFloat(i1),parseFloat(i2)]);
 });

 let res=await fetch("/plan",{method:"POST",headers:{"Content-Type":"application/json"},
 body:JSON.stringify({
 master_width:mw.value,
 coil_weight:cw.value,
 tolerance:tol.value,
 min_utilization:util.value,
 order:order
 })});

 let data=await res.json();
 window.data=data;

 let out=document.getElementById("out");
 out.innerHTML="";

 data.forEach((c,i)=>{
 let html=`<div class="card">
 <h3>Coil ${i+1}</h3>
 Used: ${c.used_width} mm | Remaining: ${c.remaining_width} mm
 <div class="total">Weight: ${c.total_weight} MT</div>
 <table>
 <tr><th>Size</th><th>Slits</th><th>Width</th><th>Total</th></tr>`;

 c.coil.forEach(x=>{
 html+=`<tr>
 <td>${x.size}</td>
 <td>${x.slits}</td>
 <td>${x.width}</td>
 <td>${x.total_weight}</td>
 </tr>`;
 });

 html+="</table></div>";
 out.innerHTML+=html;
 });
}

async function download(){
 let res=await fetch("/export",{method:"POST",
 headers:{"Content-Type":"application/json"},
 body:JSON.stringify(window.data)});
 let blob=await res.blob();
 let a=document.createElement("a");
 a.href=URL.createObjectURL(blob);
 a.download="plan.xlsx";
 a.click();
}
</script>

<button onclick="download()">Download Excel</button>

</body>
</html>
"""

# =====================================================
# ROUTES
# =====================================================
@app.route("/")
def home():
    return render_template_string(HTML)

@app.route("/plan", methods=["POST"])
def plan_api():
    d = request.json
    return jsonify(plan_multi_coil(
        d["order"],
        float(d["master_width"]),
        float(d["coil_weight"]),
        float(d["tolerance"])/1000,
        float(d["min_utilization"])/100
    ))

@app.route("/export", methods=["POST"])
def export():
    data = request.json

    wb = Workbook()
    ws = wb.active

    headers = ["Coil","Size(mm)","Slits","Width(mm)","Weight(MT)"]
    ws.append(headers)

    for c in ws[1]:
        c.font = Font(bold=True)

    tw, twt = 0,0

    for i,c in enumerate(data):
        for it in c["coil"]:
            tw += it["width"]
            twt += it["total_weight"]
            ws.append([i+1,it["size"],it["slits"],it["width"],it["total_weight"]])

    ws.append(["TOTAL","","",round(tw,2),round(twt,2)])

    for c in ws[ws.max_row]:
        c.font = Font(bold=True)

    for row in ws.iter_rows():
        for c in row:
            c.alignment = Alignment(horizontal="center")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    return send_file(buf, download_name="coil_plan.xlsx", as_attachment=True)

# =====================================================
# RUN
# =====================================================
if __name__ == "__main__":
    app.run()
