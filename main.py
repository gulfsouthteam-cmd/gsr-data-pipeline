
from flask import Flask, request, jsonify
from openpyxl import load_workbook
import io, re, base64

app = Flask(__name__)

ROWS = {
    "income_services":          6,   # 4000 Services
    "income_scrap_metal":       7,   # 4050 Scrap Metal Sales
    "income_equipter":          8,   # 4080 Equipter Lease Income
    "income_returns":           9,   # 4090 Returns
    "job_income":              10,   # Job Income
    "income_total":            12,   # Total for Income
    "licenses_permits":        14,   # 5010 Business licenses & permits
    "commissions_cogs":        15,   # 5020 Commissions COGS
    "cogs_base":               16,   # 5030 Cost of goods sold
    "wcp_builders_mutual":     17,   # 6115 WCP Builders Mutual Policy
    "bond":                    18,   # Bond
    "cogs_30_total":           20,   # Total for 5030
    "fortified_inspections":   21,   # 5035 Fortified Inspections
    "equipment_rental":        22,   # 5040 Equipment Rental for Jobs
    "equipment_hauling":       23,   # Equipment Hauling
    "equipment_total":         24,   # Total for 5040
    "fuel_gas":                25,   # 5200 Fuel & Gas
    "hotels_travel":           26,   # 5210 Hotels/Travel COGS
    "job_plans":               27,   # 5220 Job Plans
    "materials":               28,   # 5230 Job Supplies & Materials
    "mgmt_fees_cogs":          29,   # 5240 Management Fees COGS
    "misc_service_cost":       30,   # 5250 Other Miscellaneous Service Cost
    "permits":                 31,   # 5260 Permits
    "purchases":               32,   # 5270 Purchases
    "shipping":                33,   # 5280 Shipping
    "subcontractors":          34,   # 5290 Subcontractor expenses
    "labor_service_fees":      35,   # 5295 Labor Service Fees
    "subcontractors_total":    36,   # Total for 5290
    "tool_inventory":          37,   # 5300 Tool Inventory
    "waste_removal":           38,   # 5310 Waste Removal
    "commercial_package":      39,   # 5320 Commercial Package Policy
    "field_payroll_taxes":     43,   # 5100 Field Payroll Taxes
    "field_pto":               44,   # 5105 Field PTO
    "field_wages":             45,   # 5110 Field Salaries & Wages
    "field_401k":              46,   # 5115 Field Staff 401K Match
    "field_health_insurance":  47,   # 5120 Field Staff Health Insurance
    "field_payroll_fees":      48,   # 5125 Field Staff Payroll Fees
    "labor_total":             49,   # Total for Field Staff Payroll
    "total_cogs":              50,   # Total for Cost of Goods Sold
    "gross_profit":            51,   # Gross Profit
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
            "period":     period,
            "customer":   customer,
            "project":    "" if project == customer else project,
            "job_number": parse_job_number(project),
        }
        for key, row_idx in ROWS.items():
            record[key] = clean(all_rows[row_idx][col_i])
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
