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
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

<style>
body {
  font-family:'Segoe UI';
  background:#f8fafc;
  color:#1e293b;
  padding:20px;
}

.container {
  max-width:900px;
  margin:auto;
}

.card {
  background:white;
  padding:20px;
  border-radius:12px;
  margin-bottom:20px;
  box-shadow: 0 4px 12px rgba(0,0,0,0.05);
}

input {
  padding:10px;
  margin:6px;
  border-radius:6px;
  border:1px solid #e2e8f0;
  width:140px;
}

button {
  padding:10px 15px;
  border-radius:6px;
  border:none;
  cursor:pointer;
}

.btn-primary {background:#22c55e;color:white;}
.btn-secondary {background:#3b82f6;color:white;}
.btn-danger {background:#ef4444;color:white;}

table {
  width:100%;
  border-collapse:collapse;
  margin-top:10px;
}

th {
  font-weight:700;
  background:#f1f5f9;
}

th,td {
  padding:10px;
  border-bottom:1px solid #e2e8f0;
  text-align:center;
}

.coil-card {
  background:white;
  padding:15px;
  margin-top:15px;
  border-radius:10px;
  box-shadow: 0 2px 10px rgba(0,0,0,0.04);
}

.total-text {
  font-weight:700;
  color:#16a34a;
}

.input-wrapper {
  position: relative;
  display: inline-block;
}

.input-wrapper input {
  padding: 10px 40px 10px 10px;
  border-radius: 8px;
  border: 1px solid #e2e8f0;
}

.input-wrapper span {
  position: absolute;
  right: 10px;
  top: 50%;
  transform: translateY(-50%);
  color: #64748b;
  font-size: 13px;
}
</style>
</head>
<body>
<div class="container">
<h1>Coil Optimization System</h1>

<div class="card">
<div class="input-wrapper">
  <input id="master_width" placeholder="Master Width">
  <span>mm</span>
</div>

<div class="input-wrapper">
  <input id="coil_weight" placeholder="Coil Weight (MT)">
  <span>MT</span>
</div>

<div class="input-wrapper">
  <input id="tolerance" placeholder="Tolerance (Kg)">
  <span>Kg</span>
</div>

<div class="input-wrapper">
  <input id="min_utilization" placeholder="Min Utilization %">
  <span>%</span>
</div>
</div>

<div class="card">
<table id="orderTable">
<tr><th>Size (mm)</th><th>Weight (MT)</th><th>Action</th></tr>
</table>
<button class="btn-secondary" onclick="addRow()">+ Add Row</button>
</div>

<button class="btn-primary" onclick="generate()">Generate Plan</button>

<div id="output"></div>

<button class="btn-secondary" onclick="downloadExcel()">Download Excel</button>
</div>

<script>
function addRow(size="", weight="") {
  const table = document.getElementById("orderTable");
  const row = table.insertRow();

  row.innerHTML = `
    <td><input value="${size}"> mm</td>
    <td><input value="${weight}"> MT</td>
    <td><button onclick="this.parentElement.parentElement.remove()">X</button></td>
  `;
}
addRow(); addRow();

async function generate(){
  const rows=document.querySelectorAll("#orderTable tr");
  let order=[];

  rows.forEach((row,i)=>{
    if(i===0)return;
    const inputs=row.querySelectorAll("input");
    const s=parseFloat(inputs[0].value);
    const w=parseFloat(inputs[1].value);
    if(!isNaN(s)&&!isNaN(w)) order.push([s,w]);
  });

  const payload={
    master_width:document.getElementById("master_width").value,
    coil_weight:document.getElementById("coil_weight").value,
    tolerance:document.getElementById("tolerance").value,
    min_utilization:document.getElementById("min_utilization").value,
    order:order
  };

  const res=await fetch("/plan",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(payload)});
  const data=await res.json();
  window.latestPlans=data;

  render(data);
}

function render(data){
  const out=document.getElementById("output");
  out.innerHTML="<h2>Results</h2>";

  data.forEach((c,i)=>{
    let total = c.coil.reduce((a,b)=>a+b.total_weight,0);

    let html=`<div class="coil-card">
    <h3>Coil ${i+1}</h3>
    <p>Used Width: <b>${c.used_width} mm</b></p>
    <p>Remaining: <b>${c.remaining_width} mm</b></p>
    <p class="total-text">Total Coil Weight: ${total.toFixed(2)} MT</p>

    <table>
    <tr>
    <th>Size</th><th>Slits</th><th>Width</th><th>Wt/mm</th><th>Wt/Slit</th><th>Total</th>
    </tr>`;

    c.coil.forEach(it=>{
      html+=`<tr>
<td>${it.size} mm</td>
<td>${it.slits}</td>
<td>${it.width} mm</td>
<td>${it.weight_per_mm} kg</td>
<td>${it.weight_per_slit} MT</td>
<td>${it.total_weight} MT</td>
</tr>`;
    });

    html+="</table></div>";
    out.innerHTML+=html;
  });
}

async function downloadExcel(){
  const res=await fetch("/export",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(window.latestPlans)});
  const blob=await res.blob();
  const a=document.createElement("a");
  a.href=URL.createObjectURL(blob);
  a.download="coil_plan.xlsx";
  a.click();
}
</script>
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
