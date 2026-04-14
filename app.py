from flask import Flask, request, render_template_string, jsonify, send_file
import pandas as pd
import io

app = Flask(__name__)

# ---------------- BACKEND LOGIC ----------------
def plan_multi_coil(order, master_width, COIL_WEIGHT, TOLERANCE, MIN_UTILIZATION):
    produced = {size: 0.0 for size, _ in order}
    demand = {size: weight for size, weight in order}
    sizes = sorted(demand.keys(), reverse=True)

    plans = []

    while True:
        best_plan = None
        min_remaining = master_width

        def backtrack(i, used_width, plan):
            nonlocal best_plan, min_remaining

            if i == len(sizes):
                remaining = master_width - used_width
                if plan and remaining < min_remaining:
                    best_plan = plan.copy()
                    min_remaining = remaining
                return

            size = sizes[i]
            weight_per_slit = (size / master_width) * COIL_WEIGHT
            max_width = (master_width - used_width) // size

            remaining_weight = demand[size] + TOLERANCE - produced[size]
            max_weight = int(max(0, remaining_weight / weight_per_slit))
            max_slits = int(min(max_width, max_weight))

            for s in range(max_slits, -1, -1):
                if s > 0:
                    plan.append((size, s))

                backtrack(i + 1, used_width + s * size, plan)

                if s > 0:
                    plan.pop()

        backtrack(0, 0, [])

        if not best_plan:
            break

        used_width = sum(size * slits for size, slits in best_plan)

        if used_width / master_width < MIN_UTILIZATION:
            if sum(demand.values()) >= COIL_WEIGHT:
                break

        coil_plan = []

        for size, slits in best_plan:
            weight_per_mm_kg = (COIL_WEIGHT *1000) / master_width
            weight_per_slit = size * (weight_per_mm_kg) / 1000
            total_weight = slits * weight_per_slit

            produced[size] += total_weight

            coil_plan.append({
                "size": size,
                "slits": slits,
                "width": slits * size,
                "weight_per_mm": round(weight_per_mm_kg, 2),
                "weight_per_slit": round(weight_per_slit, 2),
                "total_weight": round(total_weight, 2)
            })

        plans.append({
            "coil": coil_plan,
            "used_width": used_width,
            "remaining_width": master_width - used_width,
            "utilization": round(used_width / master_width, 3)
        })

        if all(produced[s] >= demand[s] for s in sizes):
            break

    return plans


# ---------------- FRONTEND ----------------
HTML = """
<!DOCTYPE html>
<html>
<head>
<title>Coil Optimizer</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

<style>
body {font-family:'Segoe UI';background:#0b1220;color:#e5e7eb;padding:20px;}
.container{max-width:900px;margin:auto;}
.card{background:#1e293b;padding:20px;border-radius:12px;margin-bottom:20px;}
input{padding:10px;margin:6px;border-radius:6px;border:none;width:140px;background:#0f172a;color:#e5e7eb;}
button{padding:10px 15px;border-radius:6px;border:none;cursor:pointer;}
.btn-primary{background:#22c55e;color:white;}
.btn-secondary{background:#3b82f6;color:white;}
.btn-danger{background:#ef4444;color:white;}
table{width:100%;border-collapse:collapse;}
th,td{padding:10px;border-bottom:1px solid #334155;text-align:center;}
.coil-card{background:#020617;padding:15px;margin-top:15px;border-radius:10px;}
</style>
</head>

<body>
<div class="container">
<h1>Coil Optimization System</h1>

<div class="card">
<input id="master_width" placeholder="Master Width (mm)">
<input id="coil_weight" placeholder="Coil Weight (MT)">
<input id="tolerance" placeholder="Tolerance (Kg)">
<input id="min_utilization" placeholder="Min Utilization %">
</div>

<div class="card">
<table id="orderTable">
<tr><th>Size (mm)</th><th>Weight (MT)</th><th>Action</th></tr>
</table>
<button class="btn-secondary" onclick="addRow()">+ Add Row</button>
</div>

<button class="btn-primary" onclick="generate()">Generate Plan</button>

<div id="output"></div>
<canvas id="chart"></canvas>

<button class="btn-secondary" onclick="downloadExcel()">Download Excel</button>
</div>

<script>
function addRow(size="", weight="") {
  const table = document.getElementById("orderTable");
  const rowCount = table.rows.length;

  const row = table.insertRow();

  if (rowCount === 1) {
    row.innerHTML = `
      <td><input placeholder="mm" value="${size}"></td>
      <td><input placeholder="MT" value="${weight}"></td>
      <td>—</td>
    `;
  } else {
    row.innerHTML = `
      <td><input placeholder="mm" value="${size}"></td>
      <td><input placeholder="MT" value="${weight}"></td>
      <td><button class="btn-danger" onclick="deleteRow(this)">X</button></td>
    `;
  }
}

function deleteRow(btn) {
  const table = document.getElementById("orderTable");
  if (table.rows.length <= 2) {
    alert("At least one row is required!");
    return;
  }
  btn.parentElement.parentElement.remove();
}

addRow();
addRow();

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

  if(order.length === 0){
    alert("Please add at least one valid row!");
    return;
  }

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
  chart(data);
}

function render(data){
  const out=document.getElementById("output");
  out.innerHTML="<h2>Results</h2>";

  data.forEach((c,i)=>{
    let total = c.coil.reduce((a,b)=>a+b.total_weight,0);

    let html=`<div class="coil-card">
    <h3>Coil ${i+1}</h3>
    <p>Used Width: ${c.used_width} mm</p>
    <p>Remaining: ${c.remaining_width} mm</p>
    <p><b>Total Coil Weight:</b> ${total.toFixed(0)} MT</p>

    <table>
    <tr>
    <th>Size</th>
    <th>Slits</th>
    <th>Width (mm)</th>
    <th>Wt/mm (Kg)</th>
    <th>Wt/Slit (MT)</th>
    <th>Total Wt (MT)</th>
    </tr>`;

    c.coil.forEach(it=>{
      html+=`<tr>
<td>${it.size}</td>
<td>${it.slits}</td>
<td>${it.width}</td>
<td>${it.weight_per_mm}</td>
<td>${it.weight_per_slit}</td>
<td>${it.total_weight}</td>
</tr>`;
    });

    html+="</table></div>";
    out.innerHTML+=html;
  });
}

function chart(data){
  const ctx=document.getElementById("chart");
  new Chart(ctx,{
    type:'bar',
    data:{
      labels:data.map((_,i)=>"Coil "+(i+1)),
      datasets:[{label:'Utilization',data:data.map(d=>d.utilization)}]
    }
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

# ---------------- ROUTES ----------------
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
    rows = []

    for i,c in enumerate(data):
        for it in c["coil"]:
            rows.append({
                "Coil": i+1,
                "Size (mm)": it["size"],
                "Slits": it["slits"],
                "Width (mm)": it["width"],
                "Weight (MT)": it["total_weight"]
            })

    df = pd.DataFrame(rows)
    df.loc[len(df)] = ["", "TOTAL", "", df["Width (mm)"].sum(), df["Weight (MT)"].sum()]

    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)

    return send_file(buf, download_name="coil_plan.xlsx", as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
