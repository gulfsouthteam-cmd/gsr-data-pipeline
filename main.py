from flask import Flask, request, jsonify
from openpyxl import load_workbook
import io, re, base64

app = Flask(__name__)

ROWS = {
    "bond":                     18,
    "cogs_30_total":            20,
    "cogs_base":                16,
    "commercial_package":       39,
    "commissions_cogs":         15,
    "equipment_hauling":        23,
    "equipment_rental":         22,
    "equipment_total":          24,
    "field_401k":               46,
    "field_health_insurance":   47,
    "field_payroll_fees":       48,
    "field_payroll_taxes":      43,
    "field_pto":                44,
    "field_wages":              45,
    "fortified_inspections":    21,
    "fuel_gas":                 25,
    "gross_profit":             51,
    "hotels_travel":            26,
    "income_equipter":           8,
    "income_returns":            9,
    "income_scrap_metal":        7,
    "income_services":           6,
    "income_total":             12,
    "job_income":               10,
    "job_plans":                27,
    "labor_service_fees":       35,
    "labor_total":              49,
    "licenses_permits":         14,
    "materials":                28,
    "mgmt_fees_cogs":           29,
    "misc_service_cost":        30,
    "permits":                  31,
    "purchases":                32,
    "shipping":                 33,
    "subcontractors":           34,
    "subcontractors_total":     36,
    "tool_inventory":           37,
    "total_cogs":               50,
    "waste_removal":            38,
    "wcp_builders_mutual":      17,
}

def clean(value):
    if value is None: return 0.0
    if isinstance(value, (int, float)): return round(float(value), 2)
    return 0.0

def has_data(all_rows, col_idx):
    for r in [ROWS["income_total"], ROWS["materials"], ROWS["labor_total"], ROWS["total_cogs"]]:
        v = all_rows[r][col_idx]
        if v and isinstance(v, (int, float)) and v != 0: return True
    return False

def parse_job_number(name):
    m = re.search(r'[-#]\s*(\d{3,5})\b', str(name))
    return m.group(1) if m else ""

def process_workbook(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
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
        record = {
            "customer":   customer,
            "job_number": parse_job_number(project),
            "period":     period,
            "project":    "" if project == customer else project,
        }
        for key in sorted(ROWS.keys()):
            record[key] = clean(all_rows[ROWS[key]][col_i])
        records.append(record)

    return records

@app.route('/process', methods=['POST'])
def process():
    file = request.files.get('file')
    if file:
        return jsonify(process_workbook(file.read()))

    if request.data:
        try:
            return jsonify(process_workbook(request.data))
        except: pass

    if request.json:
        data = request.json.get('data') or request.json.get('file')
        if data:
            try:
                return jsonify(process_workbook(base64.b64decode(data)))
            except Exception as e:
                return jsonify({"error": str(e)}), 400

    return jsonify({"error": "No file received"}), 400

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
