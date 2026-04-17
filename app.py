from flask import Flask, request, render_template_string, jsonify, send_file
import pandas as pd
import io
import math
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

app = Flask(__name__)

# =========================
# 🔥 CORE LOGIC (UPDATED)
# =========================
import math

import math

# -------------------------------------------------
# 🔥 STEP 1: BEST COMBINATION (Knapsack DP)
# -------------------------------------------------

import math

# -------------------------------------------------
# 🔥 STEP 1: BEST COMBINATION (Knapsack DP)
# -------------------------------------------------
def best_coil_combination(sizes, max_width, max_use):
    dp = {0: [0] * len(sizes)}

    for i, size in enumerate(sizes):
        new_dp = dp.copy()

        for width, counts in dp.items():
            k = 1
            while True:
                new_width = width + size * k

                if new_width > max_width:
                    break

                if k > max_use[size]:
                    break

                new_counts = counts.copy()
                new_counts[i] += k

                if new_width not in new_dp:
                    new_dp[new_width] = new_counts

                k += 1

        dp = new_dp

    best_width = max(dp.keys())
    return best_width, dp[best_width]


# -------------------------------------------------
# 🚀 MAIN MULTI-COIL PLANNER
# -------------------------------------------------
def plan_multi_coil(order, master_width, COIL_WEIGHT, TOLERANCE, MIN_UTILIZATION):

    print("\n🚀 Production Planning Started...\n")

    # Normalize
    # order = [(int(round(size)), weight) f
    SCALE = 10
    order = [(int(round(size * SCALE)), weight) for size, weight in order]
    master_width = int(round(master_width * SCALE))

    weight_per_mm_kg = (COIL_WEIGHT * 1000) / master_width / SCALE

    demand_slits = {}
    weight_per_slit_map = {}

    # Convert MT → slits
    for size, demand_mt in order:
        weight_per_slit = (size * weight_per_mm_kg) / 1000

        min_slits = max(1, math.ceil((demand_mt - TOLERANCE) / weight_per_slit))

        demand_slits[size] = min_slits
        weight_per_slit_map[size] = weight_per_slit

        print(f"📦 Size {size} → Required Slits: {min_slits}")

    sizes = sorted(demand_slits.keys())
    plans = []

    # 🔒 Safety
    max_iterations = 50
    iteration = 0

    # -------------------------------------------------
    # 🔁 MULTI-COIL LOOP
    # -------------------------------------------------
    while any(v > 0 for v in demand_slits.values()):

        iteration += 1
        if iteration > max_iterations:
            print("❌ Safety break (infinite loop protection)")
            break

        print("\n🌀 New Coil")

        base_counts = [0] * len(sizes)
        used_width = 0

        # 🔥 max_use (CORRECT PLACE)
        max_use = {}
        for size in sizes:
            remaining = demand_slits[size]
            if remaining <= 0:
                max_use[size] = 0
            else:
                max_use[size] = remaining + 2  # small buffer

        # 🔹 STEP 1: Mandatory allocation
        for i, size in enumerate(sizes):
            if demand_slits[size] > 0:
                base_counts[i] += 1
                demand_slits[size] -= 1
                used_width += size

        remaining_width = master_width - used_width

        # 🔹 STEP 2: Optimize remaining width
        best_width, combo = best_coil_combination(sizes, remaining_width, max_use)

        # Merge
        final_counts = []
        for i in range(len(sizes)):
            final_counts.append(base_counts[i] + combo[i])

        # 🔥 CRITICAL FIX: Reduce demand AFTER DP
        for i, size in enumerate(sizes):
            demand_slits[size] -= final_counts[i]
            if demand_slits[size] < 0:
                demand_slits[size] = 0

        # 🔹 STEP 3: Build output
        coil_plan = []
        total_weight = 0

        for i, size in enumerate(sizes):
            slits = final_counts[i]
            if slits == 0:
                continue

            width = slits * size
            weight_per_slit = weight_per_slit_map[size]
            total = slits * weight_per_slit

            total_weight += total

            coil_plan.append({
                # "size": size,
                "size": round(size / SCALE,2),
                "slits": slits,
                # "width": width,
                "width": round(width/SCALE,2,
                "weight_per_mm": round(weight_per_mm_kg, 2),
                "weight_per_slit": round(weight_per_slit, 3),
                "total_weight": round(total, 3)
            })

        # Sort for clean UI (like screenshot)
        coil_plan = sorted(coil_plan, key=lambda x: x["size"])

        used_width = sum(x["width"] for x in coil_plan)
        remaining_width = master_width - used_width
        utilization = used_width / master_width

        print(f"➡️ Used Width: {used_width}")
        print(f"➡️ Utilization: {utilization:.4f}")

        plans.append({
            "coil": coil_plan,
            "used_width": round(used_width, 2),
            "remaining_width": round(remaining_width, 2),
            "utilization": round(utilization, 4),
            "total_weight": round(total_weight, 3)
        })

    print("\n✅ Production Planning Complete\n")
    return plans

# =========================
# 🎨 FRONTEND (LIGHT UI)
# =========================
HTML = """<!DOCTYPE html>
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

# =========================
# ROUTES
# =========================
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
        float(d["tolerance"]) / 1000,
        float(d["min_utilization"]) / 100
    ))

@app.route("/export", methods=["POST"])
def export():
    data = request.json

    wb = Workbook()
    ws = wb.active

    headers = ["Coil", "Size (mm)", "Slits", "Width (mm)", "Weight (MT)"]
    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True)

    total_width = 0
    total_weight = 0

    for i, c in enumerate(data):
        for it in c["coil"]:
            total_width += it["width"]
            total_weight += it["total_weight"]

            ws.append([
                i+1,
                f"{it['size']}mm",
                it["slits"],
                f"{it['width']}mm",
                f"{it['total_weight']}MT"
            ])

    ws.append(["TOTAL","","",f"{total_width}mm",f"{round(total_weight,2)}MT"])

    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal="center")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    return send_file(buf, download_name="coil_plan.xlsx", as_attachment=True)

# =========================
# RUN
# =========================
if __name__ == "__main__":
    app.run(debug=True)
