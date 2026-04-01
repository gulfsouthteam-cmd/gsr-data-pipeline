from flask import Flask, request, jsonify
from openpyxl import load_workbook
import io, re

app = Flask(__name__)

# Row indexes (0-based) for every COGS line item
ROWS = {
    "income_services":          6,   # 4000 Services
    "income_scrap_metal":       7,   # 4050 Scrap Metal Sales
    "income_total":             8,   # Total Income
    "licenses_permits":        10,   # 5010 Business licenses & permits
    "cogs_base":               11,   # 5030 Cost of goods sold
    "wcp_builders_mutual":     12,   # 6115 WCP Builders Mutual Policy
    "bond":                    13,   # Bond
    "cogs_total":              14,   # Total 5030
    "fortified_inspections":   15,   # 5035 Fortified Inspections
    "equipment_rental":        16,   # 5040 Equipment Rental
    "equipment_hauling":       17,   # Equipment Hauling
    "equipment_total":         18,   # Total 5040
    "fuel_gas":                19,   # 5200 Fuel & Gas
    "hotels_travel":           20,   # 5210 Hotels/Travel
    "job_plans":               21,   # 5220 Job Plans
    "materials":               22,   # 5230 Job Supplies & Materials
    "permits":                 23,   # 5260 Permits
    "subcontractors":          24,   # 5290 Subcontractor expenses
    "labor_service_fees":      25,   # 5295 Labor Service Fees
    "subcontractors_total":    26,   # Total 5290
    "tool_inventory":          27,   # 5300 Tool Inventory
    "waste_removal":           28,   # 5310 Waste Removal
    "commercial_package":      29,   # 5320 Commercial Package Policy
    "field_payroll_taxes":     31,   # 5100 Field Payroll Taxes
    "field_wages":             32,   # 5110 Field Salaries & Wages
    "labor_total":             33,   # Total Field Staff Payroll
    "total_cogs":              34,   # Total Cost of Goods Sold
    "gross_profit":            35,   # Gross Profit
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
            "period": period,
            "customer": customer,
            "project": "" if project == customer else project,
            "job_number": parse_job_number(project),
        }
        for key, row_idx in ROWS.items():
            record[key] = clean(all_rows[row_idx][col_i])
        records.append(record)
    return records

@app.route('/process', methods=['POST'])
def process():
    # Try multipart file upload
    file = request.files.get('file')
    if file:
        return jsonify(process_workbook(file.read()))

    # Try raw binary body
    if request.data:
        try:
            return jsonify(process_workbook(request.data))
        except: pass

    # Try base64 JSON body
    if request.json:
        import base64
        data = request.json.get('data') or request.json.get('file')
        if data:
            try:
                return jsonify(process_workbook(base64.b64decode(data)))
            except Exception as e:
                return jsonify({"error": str(e)}), 400

    return jsonify({"error": "No file received"}), 400

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
