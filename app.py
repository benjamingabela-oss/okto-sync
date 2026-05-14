import os
import json
import time
import tempfile
from flask import Flask, request, jsonify, send_from_directory
import requests
import openpyxl
import anthropic

app = Flask(__name__, static_folder="static")

# ── Anthropic key lives on the server — never exposed to users
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

ASANA_FIELDS = {
    "market_value":  "1209691181990947",
    "debt":          "1209690936207910",
    "noi":           "1209690933762869",
    "dscr":          "1209691184219184",
    "ltv":           "1209691185448001",
    "cap_rate":      "1209691329150592",
    "cost":          "1209693994446597",
    "units":         "1209720926990044",
    "sf_building":   "1209720821050620",
    "sf_land":       "1209759849844611",
    "gross_income":  "1209762019720933",
    "cashflow":      "1210006419611186",
    "expense":       "1210006388018561",
    "cashdown":      "1210333099357386",
    "price":         "1209636012311297",
    "cash_call":     "1210006382272158",
}

PROJECT_GIDS = {
    "Multi-Res IPP - Calgary - Riverview":           "1210853284875107",
    "Land - NCC Mangin Land":                        "1209331149696129",
    "Ind IPP - 700 Avenue Beaumont":                 "1209331149696124",
    "Multi-Res IPP - Edmonton - Garneau":            "1209331149696094",
    "Multi-Res IPP - Edmonton - Imperial":           "1209556797269329",
    "Multi-Res IPP - Edmonton - Insignia":           "1209624097039783",
    "Multi-Res IPP - Edmonton - Capilano Tower":     "1209331149696099",
    "Multi-Res - Allumetieres":                      "1209331149696158",
    "Multi-Res - Place Laval Land":                  "1209331149696104",
    "Multi-Res IPP - Edmonton - Lansdowne":          "1209771721895696",
    "Multi-Res IPP - Edmonton - Galbraith":          "1209771721895687",
    "Multi-Res IPP - Edmonton - Axcess":             "1209198438088491",
    "Multi-Res IPP - Edmonton - Hamptons":           "1209198438088484",
    "Multi-Res IPP - Gatineau - Eleonore":           "1203839657006504",
    "Multi-Res - Alexandra":                         "1203878731582881",
    "Multi-Res IPP - Dorval - Walt":                 "1203878731582875",
    "Multi-Res - Moqueurs":                          "1203878731582869",
    "Multi-Res - 720 MTL-TO (Romeo)":               "1209120065416144",
    "Multi-Res - 720 MTL-TO (Roméo)":               "1209120065416144",
    "720 MTL-TO":                                    "1209120065416144",
    "720 MTL-TO Limited Partnership":                "1209120065416144",
    "Multi-Res - 605 Wilfrid-Hamel (Charles)":      "1203878731582893",
    "Multi-Res - 4717 Wellington":                   "1209671658434874",
    "Multi-Res - 311 de l'Eglise":                  "1209671805092934",
    "Commercial - Loblaws Cowansville":              "1210023255872586",
    "Storage - 12455 Sherbrooke Est (BB)":           "1206322186668459",
    "Land - Laurier":                                "1209120065416220",
    "Commercial - Loblaws Acton Vale":               "1203831185421665",
    "Commercial - 1400 Taschereau":                  "1209569869960748",
    "Land - Land CineParc":                          "1203878731582887",
    "Land - Land Walkley":                           "1203831300537692",
    "Mokto Walkley LP":                              "1203831300537692",
    "Ind IPP - Couture":                             "1209198438088461",
    "Ind IPP - 7725 Cordner":                        "1206343135145159",
    "Ind IPP - Chateauguay 315 BLVD Industrial":     "1209331149696153",
    "Ind IPP - Hamel":                               "1207216084898282",
    "Ind IPP - Newton":                              "1209785685167226",
    "Storage - Charest (BB)":                        "1203877426888063",
    "Storage - Seigneuriale (BB)":                   "1203877426888066",
    "Storage - Leman, Laval (BB)":                   "1203877426888084",
}


def read_excel_to_text(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    output = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        output.append(f"=== SHEET: {sheet_name} ===")
        for row in ws.iter_rows(values_only=True):
            if any(v is not None for v in row):
                clean = [str(v) if v is not None else "" for v in row]
                output.append("\t".join(clean))
    return "\n".join(output)


def ask_claude(excel_text):
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    system_prompt = """You are a real estate financial data extraction assistant for Okto, a real estate investment firm.

You will receive raw Excel file content. Your job is to:
1. Identify which Okto project(s) the file relates to
2. Extract any financial figures that map to these Asana fields:
   - market_value: Market/appraised value of the property
   - debt: Total debt / loan balance
   - noi: Net Operating Income
   - dscr: Debt Service Coverage Ratio
   - ltv: Loan to Value ratio (as decimal e.g. 0.75)
   - cap_rate: Capitalization rate (as decimal e.g. 0.065)
   - cost: Total cost / acquisition cost
   - units: Number of units
   - sf_building: Building square footage
   - sf_land: Land square footage
   - gross_income: Gross rental income
   - cashflow: Net cash flow
   - expense: Total expenses
   - cashdown: Cash down / equity contributed
   - price: Purchase price
   - cash_call: Capital call amount

Return ONLY a valid JSON object, nothing else:
{
  "projects": [
    {
      "name": "exact project name",
      "fields": { "field_name": numeric_value },
      "confidence": "high/medium/low",
      "notes": "brief explanation"
    }
  ]
}

Known aliases:
- 720 MTL-TO / 720 MTL-TO Limited Partnership → "Multi-Res - 720 MTL-TO (Romeo)"
- Mokto Walkley / Mokto Walkley LP → "Land - Land Walkley"

Only include fields with actual non-zero values. Only numeric values."""

    message = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=1000,
        messages=[{"role": "user", "content": f"Excel content:\n\n{excel_text[:15000]}"}],
        system=system_prompt
    )

    raw = message.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw.strip())


def update_asana_project(gid, custom_fields, asana_token):
    r = requests.put(
        f"https://app.asana.com/api/1.0/projects/{gid}",
        headers={"Authorization": f"Bearer {asana_token}", "Content-Type": "application/json"},
        json={"data": {"custom_fields": custom_fields}},
        timeout=15
    )
    if r.status_code == 200:
        return True, None
    try:
        err = r.json().get("errors", [{}])[0].get("message", r.text)
    except Exception:
        err = r.text
    return False, err


# ── Routes ────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.route("/analyze", methods=["POST"])
def analyze():
    """Step 1: Read file + ask Claude. Returns what it found before pushing."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    if not file.filename.endswith((".xlsx", ".xls")):
        return jsonify({"error": "Please upload an Excel file (.xlsx or .xls)"}), 400

    if not ANTHROPIC_API_KEY:
        return jsonify({"error": "Anthropic API key not configured on server"}), 500

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        file.save(tmp.name)
        try:
            excel_text = read_excel_to_text(tmp.name)
            result = ask_claude(excel_text)
            return jsonify(result)
        except json.JSONDecodeError:
            return jsonify({"error": "Claude couldn't parse the file. Try a different file."}), 400
        except Exception as e:
            return jsonify({"error": str(e)}), 500
        finally:
            os.unlink(tmp.name)


@app.route("/sync", methods=["POST"])
def sync():
    """Step 2: Push confirmed updates to Asana."""
    data = request.json
    asana_token = data.get("asana_token", "").strip()
    projects = data.get("projects", [])

    if not asana_token:
        return jsonify({"error": "Asana token required"}), 400

    results = []
    for p in projects:
        name = p.get("name")
        fields = p.get("fields", {})

        if not name or not fields:
            continue

        gid = PROJECT_GIDS.get(name)
        if not gid:
            results.append({"name": name, "status": "skipped", "reason": "Project not found in GID map"})
            continue

        custom_fields = {}
        for field_name, value in fields.items():
            field_gid = ASANA_FIELDS.get(field_name)
            if field_gid:
                custom_fields[field_gid] = float(value)

        if not custom_fields:
            results.append({"name": name, "status": "skipped", "reason": "No recognized fields"})
            continue

        ok, err = update_asana_project(gid, custom_fields, asana_token)
        time.sleep(0.2)

        if ok:
            results.append({"name": name, "status": "success"})
        else:
            results.append({"name": name, "status": "error", "reason": err})

    return jsonify({"results": results})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)


@app.route("/chat", methods=["POST"])
def chat():
    data = request.json
    message = data.get("message", "").strip()
    asana_token = data.get("asana_token", "").strip()

    if not message:
        return jsonify({"error": "No message provided"}), 400
    if not asana_token:
        return jsonify({"error": "Asana token required"}), 400
    if not ANTHROPIC_API_KEY:
        return jsonify({"error": "Anthropic API key not configured on server"}), 500

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    project_list = "\n".join(f"- {n}" for n in PROJECT_GIDS.keys())
    field_list = "\n".join(f"- {n}" for n in ASANA_FIELDS.keys())

    system_prompt = f"""You are an Asana update assistant for Okto, a real estate firm.
Parse the user's natural language instruction and return JSON of what to update.

Available projects:
{project_list}

Available fields:
{field_list}

Return ONLY valid JSON:
{{"updates": [{{"name": "exact project name", "fields": {{"field_name": numeric_value}}, "summary": "what changed"}}], "message": "optional fallback"}}

Match project names loosely. Only numeric values. If unparseable return {{"updates": [], "message": "Could not parse. Try: update [project] [field] to [value]"}}"""

    try:
        response = client.messages.create(
            model="claude-sonnet-4-5",
            max_tokens=800,
            system=system_prompt,
            messages=[{"role": "user", "content": message}]
        )
        raw = response.content[0].text.strip()
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
        parsed = json.loads(raw.strip())
    except Exception as e:
        return jsonify({"error": f"AI error: {str(e)}"}), 500

    updates = parsed.get("updates", [])
    if not updates:
        return jsonify({"updates": [], "message": parsed.get("message", "Nothing to update.")})

    results = []
    for u in updates:
        name = u.get("name")
        fields = u.get("fields", {})
        summary = u.get("summary", "")
        gid = PROJECT_GIDS.get(name)
        if not gid:
            results.append({"name": name, "status": "skipped", "reason": "Project not found", "summary": summary})
            continue
        custom_fields = {}
        for field_name, value in fields.items():
            field_gid = ASANA_FIELDS.get(field_name)
            if field_gid:
                custom_fields[field_gid] = float(value)
        if not custom_fields:
            results.append({"name": name, "status": "skipped", "reason": "No valid fields", "summary": summary})
            continue
        ok, err = update_asana_project(gid, custom_fields, asana_token)
        time.sleep(0.2)
        if ok:
            results.append({"name": name, "status": "success", "summary": summary})
        else:
            results.append({"name": name, "status": "error", "reason": err, "summary": summary})

    return jsonify({"updates": results})
