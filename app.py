import os, json, time, tempfile, requests
from flask import Flask, request, jsonify, send_from_directory, redirect, session
from urllib.parse import urlencode
import openpyxl, anthropic

app = Flask(__name__, static_folder="static")
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "okto-sync-secret-2024-xK9p")

ANTHROPIC_API_KEY   = os.environ.get("ANTHROPIC_API_KEY", "")
AZURE_CLIENT_ID     = os.environ.get("AZURE_CLIENT_ID", "c4875a3b-811a-428b-8a34-a492257db518")
AZURE_CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "")
AZURE_TENANT_ID     = os.environ.get("AZURE_TENANT_ID", "7e4f0df0-f07f-4934-97a9-872d0c5292f9")
REDIRECT_URI        = os.environ.get("REDIRECT_URI", "https://okto-sync-production.up.railway.app/auth/callback")
SCOPES              = "openid profile email Files.Read.All Sites.Read.All User.Read offline_access"

ASANA_FIELDS = {
    "market_value":"1209691181990947","debt":"1209690936207910","noi":"1209690933762869",
    "dscr":"1209691184219184","ltv":"1209691185448001","cap_rate":"1209691329150592",
    "cost":"1209693994446597","units":"1209720926990044","sf_building":"1209720821050620",
    "sf_land":"1209759849844611","gross_income":"1209762019720933","cashflow":"1210006419611186",
    "expense":"1210006388018561","cashdown":"1210333099357386","price":"1209636012311297",
    "cash_call":"1210006382272158",
}
PROJECT_GIDS = {
    "Multi-Res IPP - Calgary - Riverview":"1210853284875107",
    "Land - NCC Mangin Land":"1209331149696129",
    "Ind IPP - 700 Avenue Beaumont":"1209331149696124",
    "Multi-Res IPP - Edmonton - Garneau":"1209331149696094",
    "Multi-Res IPP - Edmonton - Imperial":"1209556797269329",
    "Multi-Res IPP - Edmonton - Insignia":"1209624097039783",
    "Multi-Res IPP - Edmonton - Capilano Tower":"1209331149696099",
    "Multi-Res - Allumetieres":"1209331149696158",
    "Multi-Res - Place Laval Land":"1209331149696104",
    "Multi-Res IPP - Edmonton - Lansdowne":"1209771721895696",
    "Multi-Res IPP - Edmonton - Galbraith":"1209771721895687",
    "Multi-Res IPP - Edmonton - Axcess":"1209198438088491",
    "Multi-Res IPP - Edmonton - Hamptons":"1209198438088484",
    "Multi-Res IPP - Gatineau - Eleonore":"1203839657006504",
    "Multi-Res - Alexandra":"1203878731582881",
    "Multi-Res IPP - Dorval - Walt":"1203878731582875",
    "Multi-Res - Moqueurs":"1203878731582869",
    "Multi-Res - 720 MTL-TO (Romeo)":"1209120065416144",
    "720 MTL-TO":"1209120065416144",
    "Multi-Res - 605 Wilfrid-Hamel (Charles)":"1203878731582893",
    "Multi-Res - 4717 Wellington":"1209671658434874",
    "Multi-Res - 311 de l'Eglise":"1209671805092934",
    "Commercial - Loblaws Cowansville":"1210023255872586",
    "Storage - 12455 Sherbrooke Est (BB)":"1206322186668459",
    "Land - Laurier":"1209120065416220",
    "Commercial - Loblaws Acton Vale":"1203831185421665",
    "Commercial - 1400 Taschereau":"1209569869960748",
    "Land - Land CineParc":"1203878731582887",
    "Land - Land Walkley":"1203831300537692",
    "Ind IPP - Couture":"1209198438088461",
    "Ind IPP - 7725 Cordner":"1206343135145159",
    "Ind IPP - Chateauguay 315 BLVD Industrial":"1209331149696153",
    "Ind IPP - Hamel":"1207216084898282",
    "Ind IPP - Newton":"1209785685167226",
    "Storage - Charest (BB)":"1203877426888063",
    "Storage - Seigneuriale (BB)":"1203877426888066",
    "Storage - Leman, Laval (BB)":"1203877426888084",
}

# ── Auth ──────────────────────────────────────────────────────────────
@app.route("/auth/login")
def auth_login():
    params = {"client_id":AZURE_CLIENT_ID,"response_type":"code","redirect_uri":REDIRECT_URI,"scope":SCOPES,"response_mode":"query"}
    return redirect(f"https://login.microsoftonline.com/{AZURE_TENANT_ID}/oauth2/v2.0/authorize?{urlencode(params)}")

@app.route("/auth/callback")
def auth_callback():
    code = request.args.get("code")
    if not code: return redirect("/?error=no_code")
    resp = requests.post(f"https://login.microsoftonline.com/{AZURE_TENANT_ID}/oauth2/v2.0/token", data={
        "client_id":AZURE_CLIENT_ID,"client_secret":AZURE_CLIENT_SECRET,
        "code":code,"redirect_uri":REDIRECT_URI,"grant_type":"authorization_code","scope":SCOPES,
    })
    tokens = resp.json()
    if "error" in tokens: return redirect(f"/?error=auth_failed")
    session["access_token"] = tokens.get("access_token")
    me = requests.get("https://graph.microsoft.com/v1.0/me", headers={"Authorization":f"Bearer {session['access_token']}"}).json()
    session["user_name"]  = me.get("displayName","User")
    session["user_email"] = me.get("mail") or me.get("userPrincipalName","")
    return redirect("/")

@app.route("/auth/logout")
def auth_logout():
    session.clear()
    return redirect("/")

@app.route("/auth/me")
def auth_me():
    if "access_token" not in session: return jsonify({"logged_in":False})
    return jsonify({"logged_in":True,"name":session.get("user_name"),"email":session.get("user_email")})

# ── SharePoint helpers ────────────────────────────────────────────────
def sp_search(query, access_token):
    r = requests.post("https://graph.microsoft.com/v1.0/search/query",
        headers={"Authorization":f"Bearer {access_token}","Content-Type":"application/json"},
        json={"requests":[{"entityTypes":["driveItem"],"query":{"queryString":f"{query} filetype:xlsx"},
            "fields":["name","webUrl","lastModifiedDateTime","parentReference"],
            "from":0,"size":5,"sortProperties":[{"name":"lastModifiedDateTime","isDescending":True}]}]})
    results = []
    try:
        for hit in r.json()["value"][0]["hitsContainers"][0]["hits"]:
            res = hit.get("resource",{})
            pr  = res.get("parentReference",{})
            results.append({"name":res.get("name",""),"webUrl":res.get("webUrl",""),
                "driveId":pr.get("driveId",""),"itemId":res.get("id",""),
                "modified":res.get("lastModifiedDateTime","")})
    except: pass
    return results

def read_excel_from_sp(drive_id, item_id, access_token):
    r = requests.get(f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content",
        headers={"Authorization":f"Bearer {access_token}"},allow_redirects=True)
    if r.status_code != 200: return None
    with tempfile.NamedTemporaryFile(suffix=".xlsx",delete=False) as tmp:
        tmp.write(r.content); tmp_path=tmp.name
    try:
        wb = openpyxl.load_workbook(tmp_path,data_only=True)
        out = []
        for sn in wb.sheetnames:
            ws=wb[sn]; out.append(f"=== {sn} ===")
            for row in ws.iter_rows(values_only=True):
                if any(v is not None for v in row):
                    out.append("\t".join([str(v) if v is not None else "" for v in row]))
        return "\n".join(out)
    except: return None
    finally: os.unlink(tmp_path)

def claude_extract(message, file_contents, file_names):
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    pl = "\n".join(f"- {n}" for n in PROJECT_GIDS); fl = "\n".join(f"- {n}" for n in ASANA_FIELDS)
    files_text = "".join(f"\n\n=== FILE: {n} ===\n{c[:6000]}" for n,c in zip(file_names,file_contents))
    resp = client.messages.create(model="claude-sonnet-4-5",max_tokens=1000,
        system=f'Extract financial data for Okto projects from SharePoint files.\nProjects:\n{pl}\nFields:\n{fl}\nReturn ONLY JSON: {{"updates":[{{"name":"project name","fields":{{"field":value}},"summary":"what changed + filename"}}],"message":"explanation"}}',
        messages=[{"role":"user","content":f"User asked: {message}\n\nFiles:{files_text}"}])
    raw=resp.content[0].text.strip()
    if raw.startswith("```"): raw=raw.split("```")[1]; raw=raw[4:] if raw.startswith("json") else raw
    return json.loads(raw.strip())

def claude_parse(message):
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    pl = "\n".join(f"- {n}" for n in PROJECT_GIDS); fl = "\n".join(f"- {n}" for n in ASANA_FIELDS)
    resp = client.messages.create(model="claude-sonnet-4-5",max_tokens=600,
        system=f'Parse Asana update instruction.\nProjects:\n{pl}\nFields:\n{fl}\nReturn ONLY JSON: {{"updates":[{{"name":"project","fields":{{"field":value}},"summary":"desc"}}],"message":"fallback","needs_sharepoint":true/false,"sharepoint_query":"search term"}}',
        messages=[{"role":"user","content":message}])
    raw=resp.content[0].text.strip()
    if raw.startswith("```"): raw=raw.split("```")[1]; raw=raw[4:] if raw.startswith("json") else raw
    return json.loads(raw.strip())

def push_asana(updates, asana_token):
    results=[]
    for u in updates:
        name=u.get("name"); fields=u.get("fields",{}); summary=u.get("summary","")
        gid=PROJECT_GIDS.get(name)
        if not gid: results.append({"name":name,"status":"skipped","reason":"Project not found","summary":summary}); continue
        cf={ASANA_FIELDS[fn]:float(v) for fn,v in fields.items() if fn in ASANA_FIELDS}
        if not cf: results.append({"name":name,"status":"skipped","reason":"No valid fields","summary":summary}); continue
        r=requests.put(f"https://app.asana.com/api/1.0/projects/{gid}",
            headers={"Authorization":f"Bearer {asana_token}","Content-Type":"application/json"},
            json={"data":{"custom_fields":cf}},timeout=15)
        time.sleep(0.2)
        if r.status_code==200: results.append({"name":name,"status":"success","summary":summary})
        else:
            try: err=r.json().get("errors",[{}])[0].get("message",r.text)
            except: err=r.text
            results.append({"name":name,"status":"error","reason":err,"summary":summary})
    return results

# ── Routes ────────────────────────────────────────────────────────────
@app.route("/")
def index(): return send_from_directory("static","index.html")

@app.route("/chat",methods=["POST"])
def chat():
    data=request.json; message=data.get("message","").strip(); asana_token=data.get("asana_token","").strip()
    if not message: return jsonify({"error":"No message"}),400
    if not asana_token: return jsonify({"error":"Asana token required"}),400
    if not ANTHROPIC_API_KEY: return jsonify({"error":"Anthropic key not configured"}),500
    try: parsed=claude_parse(message)
    except Exception as e: return jsonify({"error":f"AI error: {e}"}),500

    source_files=[]
    if parsed.get("needs_sharepoint") and "access_token" in session:
        query=parsed.get("sharepoint_query",message)
        files=sp_search(query,session["access_token"])
        if files:
            contents,names=[],[]
            for f in files[:3]:
                c=read_excel_from_sp(f["driveId"],f["itemId"],session["access_token"])
                if c: contents.append(c); names.append(f["name"])
            if contents:
                try: parsed=claude_extract(message,contents,names); source_files=names
                except Exception as e: return jsonify({"error":f"SharePoint error: {e}"}),500
    elif parsed.get("needs_sharepoint") and "access_token" not in session:
        return jsonify({"error":"sign_in_required","message":"Sign in with Microsoft to search SharePoint automatically"})

    updates=parsed.get("updates",[])
    if not updates: return jsonify({"updates":[],"message":parsed.get("message","Nothing to update.")})
    results=push_asana(updates,asana_token)
    return jsonify({"updates":results,"source_files":source_files})

@app.route("/analyze",methods=["POST"])
def analyze():
    if "file" not in request.files: return jsonify({"error":"No file"}),400
    file=request.files["file"]
    if not ANTHROPIC_API_KEY: return jsonify({"error":"Anthropic key not configured"}),500
    with tempfile.NamedTemporaryFile(suffix=".xlsx",delete=False) as tmp:
        file.save(tmp.name)
        try:
            wb=openpyxl.load_workbook(tmp.name,data_only=True)
            out=[]
            for sn in wb.sheetnames:
                ws=wb[sn]; out.append(f"=== {sn} ===")
                for row in ws.iter_rows(values_only=True):
                    if any(v is not None for v in row):
                        out.append("\t".join([str(v) if v is not None else "" for v in row]))
            excel_text="\n".join(out)
            client=anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
            pl="\n".join(f"- {n}" for n in PROJECT_GIDS)
            resp=client.messages.create(model="claude-sonnet-4-5",max_tokens=1000,
                system=f'Extract Okto project financial data.\nProjects:\n{pl}\nFields: market_value,debt,noi,dscr,ltv,cap_rate,cost,units,sf_building,sf_land,gross_income,cashflow,expense,cashdown,price,cash_call\nReturn ONLY JSON: {{"projects":[{{"name":"project","fields":{{"field":value}},"confidence":"high/medium/low","notes":"desc"}}]}}',
                messages=[{"role":"user","content":excel_text[:15000]}])
            raw=resp.content[0].text.strip()
            if raw.startswith("```"): raw=raw.split("```")[1]; raw=raw[4:] if raw.startswith("json") else raw
            return jsonify(json.loads(raw.strip()))
        except Exception as e: return jsonify({"error":str(e)}),500
        finally: os.unlink(tmp.name)

@app.route("/sync",methods=["POST"])
def sync():
    data=request.json; asana_token=data.get("asana_token","").strip(); projects=data.get("projects",[])
    if not asana_token: return jsonify({"error":"Asana token required"}),400
    results=[]
    for p in projects:
        name=p.get("name"); fields=p.get("fields",{})
        if not name or not fields: continue
        gid=PROJECT_GIDS.get(name)
        if not gid: results.append({"name":name,"status":"skipped","reason":"Not found"}); continue
        cf={ASANA_FIELDS[fn]:float(v) for fn,v in fields.items() if fn in ASANA_FIELDS}
        if not cf: results.append({"name":name,"status":"skipped","reason":"No valid fields"}); continue
        r=requests.put(f"https://app.asana.com/api/1.0/projects/{gid}",
            headers={"Authorization":f"Bearer {asana_token}","Content-Type":"application/json"},
            json={"data":{"custom_fields":cf}},timeout=15)
        time.sleep(0.2)
        if r.status_code==200: results.append({"name":name,"status":"success"})
        else:
            try: err=r.json().get("errors",[{}])[0].get("message",r.text)
            except: err=r.text
            results.append({"name":name,"status":"error","reason":err})
    return jsonify({"results":results})

if __name__=="__main__":
    port=int(os.environ.get("PORT",5000))
    app.run(host="0.0.0.0",port=port)
