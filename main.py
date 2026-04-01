from flask import Flask, request, jsonify
from openpyxl import load_workbook
import io, re, csv

app = Flask(__name__)

ROW_INCOME = 8
ROW_MATERIALS = 22
ROW_LABOR = 33
ROW_SUBCON = 26
ROW_COGS = 34
ROW_GROSS = 35

def clean(value):
    if value is None: return 0.0
    if isinstance(value, (int, float)): return round(float(value), 2)
    return 0.0

def has_data(all_rows, col_idx):
    for r in [ROW_INCOME, ROW_MATERIALS, ROW_LABOR, ROW_SUBCON, ROW_COGS]:
        v = all_rows[r][col_idx]
        if v and isinstance(v, (int, float)) and v != 0: return True
    return False

def parse_job_number(name):
    m = re.search(r'[-#]\s*(\d{3,5})\b', str(name))
    return m.group(1) if m else ""

@app.route('/process', methods=['POST'])
def process():
    file = request.files.get('file')
    if not file: return jsonify({"error": "No file"}), 400

    wb = load_workbook(io.BytesIO(file.read()), data_only=True)
    ws = wb.active
    all_rows = list(ws.iter_rows(values_only=True))
    period = str(all_rows[2][0]).strip()
    headers = all_rows[4]

    col_to_customer = {}
    pending = []
    for i, h in enumerate(headers):
        if not h: continue
        h = str(h).strip()
        if h in ("Total", "Not specified", ""): continue
        if h.startswith("Total for "):
            for pi in pending: col_to_customer[pi] = h.replace("Total for ", "")
            pending = []
        else:
            pending.append(i)

    records = []
    for col_i, customer in col_to_customer.items():
        if not has_data(all_rows, col_i): continue
        project = str(headers[col_i]).strip()
        records.append({
            "period": period,
            "customer": customer,
            "project": "" if project == customer else project,
            "job_number": parse_job_number(project),
            "revenue": clean(all_rows[ROW_INCOME][col_i]),
            "materials": clean(all_rows[ROW_MATERIALS][col_i]),
            "labor": clean(all_rows[ROW_LABOR][col_i]),
            "subcontractors": clean(all_rows[ROW_SUBCON][col_i]),
            "total_cogs": clean(all_rows[ROW_COGS][col_i]),
            "gross_profit": clean(all_rows[ROW_GROSS][col_i]),
        })

    return jsonify(records)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
