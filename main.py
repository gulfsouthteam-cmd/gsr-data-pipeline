from flask import Flask, request, jsonify
from openpyxl import load_workbook
import io, re, base64

app = Flask(__name__)

# Account name → output field name
ACCOUNT_MAP = {
    "4000 Services":                    "income_services",
    "4050 Scrap Metal Sales":           "income_scrap_metal",
    "4080 Equipter Lease Income":       "income_equipter",
    "4090 Returns":                     "income_returns",
    "Unapplied Cash Payment Income":    "unapplied_cash",
    "Total for Income":                 "income_total",
    "5010 Business licenses & permits": "licenses_permits",
    "5020 Commissions COGS":            "commissions_cogs",
    "5030 Cost of goods sold":          "cogs_base",
    "6115 WCP -Builders Mutual Policy": "wcp_builders_mutual",
    "Bond":                             "bond",
    "Total for 5030 Cost of goods sold":"cogs_30_total",
    "5035 Fortified Inspections":       "fortified_inspections",
    "5040 Equipment Rental for Jobs":   "equipment_rental",
    "Equipment Hauling":                "equipment_hauling",
    "Total for 5040 Equipment Rental for Jobs": "equipment_total",
    "5200 Fuel & Gas":                  "fuel_gas",
    "5210 Hotels/Travel COGS":          "hotels_travel",
    "5220 Job Plans":                   "job_plans",
    "5230 Job Supplies & Materials":    "materials",
    "5240 Management Fees COGS":        "mgmt_fees_cogs",
    "5250 Other Miscellaneous Service Cost": "misc_service_cost",
    "5260 Permits":                     "permits",
    "5270 Purchases":                   "purchases",
    "5280 Shipping":                    "shipping",
    "5290 Subcontractor expenses":      "subcontractors",
    "5295 Labor Service Fees":          "labor_service_fees",
    "Total for 5290 Subcontractor expenses": "subcontractors_total",
    "5300 Tool Inventory":              "tool_inventory",
    "5310 Waste Removal":               "waste_removal",
    "5320 Commercial Package Policy w/BM": "commercial_package",
    "5100 Field Payroll Taxes":         "field_payroll_taxes",
    "5105 Field PTO":                   "field_pto",
    "5110 Field Salaries & Wages":      "field_wages",
    "5115 Field Staff 401K Match":      "field_401k",
    "5120 Field Staff Health Insurance":"field_health_insurance",
    "5125 Field Staff Payroll Fees":    "field_payroll_fees",
    "Total for Field Staff Payroll":    "labor_total",
    "Total for Cost of Goods Sold":     "total_cogs",
    "Gross Profit":                     "gross_profit",
}

ALL_FIELDS = sorted(ACCOUNT_MAP.values())

def clean(value):
    if value is None: return 0.0
    if isinstance(value, (int, float)): return round(float(value), 2)
    return 0.0

def parse_job_number(name):
    m = re.search(r'[-#]\s*(\d{3,5})\b', str(name))
    return m.group(1) if m else ""

def process_workbook(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    all_rows = list(ws.iter_rows(values_only=True))
    period = str(all_rows[2][0]).strip()
    headers = all_rows[4]

    # Build row index by account name
    row_index = {}
    for i, row in enumerate(all_rows):
        label = row[0]
        if label and str(label).strip() in ACCOUNT_MAP:
            key = ACCOUNT_MAP[str(label).strip()]
            row_index[key] = i

    # Build customer/project column map
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

    # Build records
    records = []
    for col_i, customer in col_to_customer.items():
        # Check if column has any data
        has_data = False
        for field, row_i in row_index.items():
            v = all_rows[row_i][col_i]
            if v and isinstance(v, (int, float)) and v != 0:
                has_data = True
                break
        if not has_data:
            continue

        project = str(headers[col_i]).strip()
        record = {
            "customer":   customer,
            "job_number": parse_job_number(project),
            "period":     period,
            "project":    "" if project == customer else project,
        }
        for field in ALL_FIELDS:
            row_i = row_index.get(field)
            record[field] = clean(all_rows[row_i][col_i]) if row_i is not None else 0.0

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
