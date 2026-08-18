# =============================================================================
#  APPLY — Sales Management Portal  |  streamlit run app.py
# =============================================================================

import streamlit as st
import pandas as pd
import json, os, hashlib, io, re
from datetime import datetime, date
from collections import Counter

# ══════════════════════════════════════════════════════════════════════
#  SESSION STATE FOR JSON UPLOAD
# ══════════════════════════════════════════════════════════════════════
if "uploaded_json_data" not in st.session_state:
    st.session_state.uploaded_json_data = None
if "use_uploaded_json" not in st.session_state:
    st.session_state.use_uploaded_json = False

# ══════════════════════════════════════════════════════════════════════
#  PATHS
# ══════════════════════════════════════════════════════════════════════
DATA_DIR   = "apply_data"
DATA_FILE  = os.path.join(DATA_DIR, "data.json")
USERS_FILE = os.path.join(DATA_DIR, "users.json")
LAST_DF    = os.path.join(DATA_DIR, "last_df.xlsx")
EXCEL_SEED = "Final_Apply_Feedback.xlsx"
os.makedirs(DATA_DIR, exist_ok=True)

# ══════════════════════════════════════════════════════════════════════
#  CONSTANTS
# ══════════════════════════════════════════════════════════════════════
ADMIN_EMAIL = "admin@apply.com"
ADMIN_PASS  = "Apply@Admin2026"

FEEDBACK_OPTIONS = [
    "done", "recall", "N.A", "closed", "travel",
    "out of area", "member", "not interested",
    "Member With other Phone",
]

FB_COLORS = {
    "done": "#15803d", "recall": "#b45309", "N.A": "#64748b",
    "closed": "#b91c1c", "travel": "#1d4ed8", "out of area": "#6d28d9",
    "member": "#0e7490", "not interested": "#475569", "Member With other Phone": "#5b21b6",
}
FB_BG = {
    "done": "#dcfce7", "recall": "#fef3c7", "N.A": "#f1f5f9",
    "closed": "#fee2e2", "travel": "#dbeafe", "out of area": "#ede9fe",
    "member": "#cffafe", "not interested": "#f1f5f9", "Member With other Phone": "#ede9fe",
}

DS_OPTIONS = ["old member", "recommendation", "Facebook", "Instagram", "other"]
DS_COLORS = {"old member": "#0e7490", "recommendation": "#15803d", "Facebook": "#1d4ed8", "Instagram": "#be185d", "other": "#475569"}
DS_BG = {"old member": "#cffafe", "recommendation": "#dcfce7", "Facebook": "#dbeafe", "Instagram": "#fce7f3", "other": "#f1f5f9"}

MONTHS_AR = {1:"يناير", 2:"فبراير", 3:"مارس", 4:"أبريل", 5:"مايو", 6:"يونيو", 7:"يوليو", 8:"أغسطس", 9:"سبتمبر", 10:"أكتوبر", 11:"نوفمبر", 12:"ديسمبر"}

# ══════════════════════════════════════════════════════════════════════
#  NORMALIZATION FUNCTIONS
# ══════════════════════════════════════════════════════════════════════

def normalize_feedback(fb_value) -> str:
    if not fb_value or str(fb_value).strip() in ("", "None", "nan"):
        return ""
    fb = str(fb_value).strip().lower()
    fb_map = {
        "done": "done", "dona": "done", "don": "done", "dn": "done", "dne": "done", "do": "done",
        "recall": "recall", "recal": "recall", "reca": "recall", "rcl": "recall", "rcall": "recall", "rec": "recall",
        "closed": "closed", "close": "closed", "clos": "closed", "cls": "closed", "clse": "closed", "clsd": "closed",
        "not interested": "not interested", "not intersted": "not interested", "not": "not interested",
        "ni": "not interested", "not int": "not interested", "notinterest": "not interested", "uninterested": "not interested",
        "n.a": "N.A", "na": "N.A", "n/a": "N.A", "n_a": "N.A",
        "travel": "travel", "trvl": "travel", "travl": "travel", "trv": "travel",
        "out of area": "out of area", "out area": "out of area", "outside": "out of area", "out of": "out of area", "out": "out of area",
        "member": "member", "memb": "member", "mbr": "member",
        "member with other phone": "Member With other Phone", "member other phone": "Member With other Phone",
        "other phone": "Member With other Phone", "mem with other": "Member With other Phone", "member other": "Member With other Phone",
    }
    return fb_map.get(fb, fb)

def normalize_datasource(ds_value) -> str:
    if not ds_value or str(ds_value).strip() in ("", "None", "nan"):
        return ""
    ds = str(ds_value).strip().lower()
    ds_map = {
        "old member": "old member", "oldmember": "old member", "old": "old member", "old mem": "old member",
        "recommendation": "recommendation", "recommend": "recommendation", "recom": "recommendation", "rcm": "recommendation", "reco": "recommendation",
        "facebook": "Facebook", "fb": "Facebook", "face": "Facebook", "facebok": "Facebook", "fbook": "Facebook",
        "instagram": "Instagram", "insta": "Instagram", "ig": "Instagram", "instgram": "Instagram",
        "other": "other", "oth": "other", "others": "other",
    }
    return ds_map.get(ds, ds)

def extract_records_from_json(json_data):
    """تستخرج الـ records من أي شكل JSON - تدعم {data: []} أو [] مباشرة"""
    if isinstance(json_data, list):
        return json_data
    elif isinstance(json_data, dict):
        if "data" in json_data and isinstance(json_data["data"], list):
            return json_data["data"]
        if "records" in json_data and isinstance(json_data["records"], list):
            return json_data["records"]
        if "results" in json_data and isinstance(json_data["results"], list):
            return json_data["results"]
        first_key = next(iter(json_data.keys())) if json_data else None
        if first_key and str(first_key).lstrip('-').isdigit():
            return list(json_data.values())
    return None

# ══════════════════════════════════════════════════════════════════════
#  LOAD & SAVE RECORDS
# ══════════════════════════════════════════════════════════════════════
def _load_records() -> list:
    if st.session_state.use_uploaded_json and st.session_state.uploaded_json_data is not None:
        return st.session_state.uploaded_json_data

    if os.path.exists(DATA_FILE) and os.path.getsize(DATA_FILE) > 0:
        try:
            with open(DATA_FILE, encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError:
            return []
    return []

def _save_records(recs: list):
    if st.session_state.use_uploaded_json:
        st.session_state.uploaded_json_data = recs
    else:
        with open(DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(recs, f, ensure_ascii=False, indent=2)

# ══════════════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════════════
def hash_pw(pw: str) -> str:
    return hashlib.sha256(pw.encode()).hexdigest()

def pct(a, b):
    return round(a / b * 100) if b else 0

def bar_html(val, total, color):
    p = pct(val, total)
    return f'<div class="kpi-bar-wrap"><div class="kpi-bar" style="width:{p}%;background:{color}"></div></div>'

def fix_mobile(x) -> str:
    if x is None or str(x).strip() in ("", "None", "nan"):
        return ""
    try:
        clean = re.sub(r"[^\d]", "", str(x).split(".")[0])
        return clean.zfill(11) if clean else ""
    except Exception:
        return str(x).strip()

# ══════════════════════════════════════════════════════════════════════
#  USERS PERSISTENCE
# ══════════════════════════════════════════════════════════════════════
def _load_users() -> dict:
    if os.path.exists(USERS_FILE) and os.path.getsize(USERS_FILE) > 0:
        try:
            with open(USERS_FILE, encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError:
            pass
            
    default = {ADMIN_EMAIL: {"password": hash_pw(ADMIN_PASS), "role": "admin", "agent_code": None, "name": "Admin"}}
    _save_users(default)
    return default

def _save_users(u: dict):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(u, f, ensure_ascii=False, indent=2)

def _migrate_feedback(records: list) -> int:
    n = 0
    for r in records:
        old_fb = r.get("Feedback (Sales)", "")
        new_fb = normalize_feedback(old_fb)
        if old_fb != new_fb:
            r["Feedback (Sales)"] = new_fb
            n += 1
        old_ds = r.get("Data Source Feedback", "")
        new_ds = normalize_datasource(old_ds)
        if old_ds != new_ds:
            r["Data Source Feedback"] = new_ds
            n += 1
    return n

def _normalise_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.where(pd.notna(df), None)
    for col, fn in [
        ("Assign Data", lambda x: str(x)[:10] if x else ""),
        ("Agent Code", lambda x: str(int(float(x))) if x is not None and str(x).strip() not in ("", "None") else ""),
        ("Mobile", fix_mobile),
    ]:
        if col in df.columns:
            df[col] = df[col].apply(fn)
    if "Feedback (Sales)" in df.columns:
        df["Feedback (Sales)"] = df["Feedback (Sales)"].apply(normalize_feedback)
    if "Data Source Feedback" in df.columns:
        df["Data Source Feedback"] = df["Data Source Feedback"].apply(normalize_datasource)
    return df

def _to_excel_bytes(records: list) -> bytes:
    df = pd.DataFrame(records)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Apply")
        ws = w.sheets["Apply"]
        headers = [ws.cell(1, c).value for c in range(1, ws.max_column+1)]
        if "Mobile" in headers:
            col_idx = headers.index("Mobile") + 1
            for row in range(2, ws.max_row+1):
                cell = ws.cell(row, col_idx)
                val = str(cell.value) if cell.value is not None else ""
                if val and val not in ("None","nan",""):
                    cell.value = fix_mobile(val)
                    cell.number_format = "@"
    return buf.getvalue()

def _save_df(df: pd.DataFrame):
    try:
        df = df.copy()
        if "Mobile" in df.columns:
            df["Mobile"] = df["Mobile"].apply(fix_mobile)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df.to_excel(w, index=False, sheet_name="Apply")
            ws = w.sheets["Apply"]
            headers = [ws.cell(1,c).value for c in range(1, ws.max_column+1)]
            if "Mobile" in headers:
                ci = headers.index("Mobile") + 1
                for row in range(2, ws.max_row+1):
                    ws.cell(row, ci).number_format = "@"
        with open(LAST_DF, "wb") as f:
            f.write(buf.getvalue())
    except Exception:
        pass

def _auto_create_agents(records: list) -> list:
    users = _load_users()
    created = []
    for r in records:
        ac = str(r.get("Agent Code","")).strip()
        if not ac: continue
        email = f"agent{ac}@apply.com"
        if email not in users:
            users[email] = {"password": hash_pw(f"Apply{ac}"), "role": "sales", "agent_code": ac, "name": f"Agent {ac}"}
            created.append(email)
    if created: _save_users(users)
    return created

def _get_stats(records: list) -> dict:
    total = len(records)
    fb_cnt = Counter()
    ds_cnt = Counter()
    for r in records:
        fb = normalize_feedback(r.get("Feedback (Sales)", ""))
        ds = normalize_datasource(r.get("Data Source Feedback", ""))
        if fb: fb_cnt[fb] += 1
        if ds: ds_cnt[ds] += 1
    return {
        "total": total,
        "fb": dict(fb_cnt),
        "ds": dict(ds_cnt),
        "done": fb_cnt.get("done", 0),
        "recall": fb_cnt.get("recall", 0),
        "closed": fb_cnt.get("closed", 0),
        "ni": fb_cnt.get("not interested", 0),
        "na": fb_cnt.get("N.A", 0),
    }

def _build_html_report(records: list) -> str:
    df = pd.DataFrame(records)
    html_table = df.to_html(classes="table table-striped", index=False)
    return f"<html><head><title>Apply Report</title><link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css'></head><body class='p-4'><h2>Sales Report</h2>{html_table}</body></html>"

def _build_excel_dashboard(records: list) -> bytes:
    return _to_excel_bytes(records)

# ══════════════════════════════════════════════════════════════════════
#  PAGE CONFIG + CSS
# ══════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="Apply", page_icon="📋", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800;900&family=DM+Sans:wght@300;400;500;600&display=swap');
*, *::before, *::after { font-family: 'DM Sans', sans-serif !important; box-sizing: border-box; }
html, body, .stApp { background: #f6f8fc !important; color: #1e2d45 !important; font-size: 15px !important; }
.stApp > header { background: transparent !important; }
section[data-testid="stSidebar"] { background: #ffffff !important; border-right: 1px solid #e4eaf3 !important; box-shadow: 2px 0 10px rgba(0,0,0,.05); }
section[data-testid="stSidebar"] * { color: #334155 !important; font-size: 14px !important; }
h1 { color: #1e40af !important; font-family: 'Syne', sans-serif !important; font-weight: 900 !important; font-size: 1.9rem !important; }
h2 { color: #1e40af !important; font-family: 'Syne', sans-serif !important; font-weight: 800 !important; }
h3 { color: #2563eb !important; font-weight: 700 !important; }
.stTextInput input, .stTextArea textarea, .stNumberInput input { background: #ffffff !important; border: 1.5px solid #d0dae8 !important; border-radius: 10px !important; color: #1e2d45 !important; font-size: 14px !important; }
.stButton > button, [data-testid="stFormSubmitButton"] > button { background: linear-gradient(135deg, #1d4ed8, #2563eb) !important; color: #fff !important; border: none !important; border-radius: 10px !important; font-weight: 700 !important; font-size: 14px !important; }
.stDownloadButton > button { background: linear-gradient(135deg, #065f46, #059669) !important; color: #fff !important; border: none !important; border-radius: 10px !important; font-weight: 700 !important; font-size: 14px !important; }
.stTabs [data-baseweb="tab-list"] { background: #eaf0f8 !important; border-radius: 12px; padding: 4px; gap: 4px; border: none !important; }
.stTabs [data-baseweb="tab"] { color: #64748b !important; font-weight: 600; border-radius: 9px; font-size: 14px !important; }
.stTabs [aria-selected="true"] { background: #ffffff !important; color: #1d4ed8 !important; font-weight: 700 !important; }
.ap-hero { background: linear-gradient(135deg, #ffffff 0%, #f0f5ff 100%); border: 1px solid #dbeafe; border-radius: 18px; padding: 26px 30px; margin-bottom: 22px; }
.ap-hero-title { font-family: 'Syne', sans-serif; font-size: 1.6rem; font-weight: 900; color: #1e2d45; }
.ap-hero-sub { color: #64748b; font-size: 14px; margin-top: 5px; }
.kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 14px; margin: 14px 0; }
.kpi-card { background: #ffffff; border: 1px solid #e4eaf3; border-radius: 14px; padding: 20px 22px; }
.kpi-ttl { font-size: 11px; color: #94a3b8; font-weight: 700; margin-bottom: 8px; text-transform: uppercase; }
.kpi-val { font-family: 'Syne', sans-serif; font-size: 2.4rem; font-weight: 900; }
.kpi-sub { font-size: 14px; color: #64748b; margin-top: 6px; }
.kpi-bar-wrap { background: #f1f5f9; border-radius: 5px; height: 7px; margin-top: 10px; overflow: hidden; }
.kpi-bar { height: 7px; border-radius: 5px; }
.sec-title { font-family: 'Syne', sans-serif; font-size: 1.05rem; font-weight: 800; color: #1e40af; border-bottom: 2px solid #dbeafe; padding-bottom: 9px; margin: 24px 0 16px; }
.stat-cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(135px, 1fr)); gap: 12px; margin: 14px 0; }
.stat-card { background: #ffffff; border: 1px solid #e4eaf3; border-radius: 12px; padding: 18px 12px; text-align: center; }
.stat-num { font-family: 'Syne', sans-serif; font-size: 2rem; font-weight: 900; }
.stat-lbl { font-size: 11px; color: #64748b; margin-top: 6px; text-transform: uppercase; }
.login-wrap { max-width: 420px; margin: 50px auto; }
.login-box { background: #ffffff; border: 1px solid #e4eaf3; border-radius: 20px; padding: 38px 34px; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════
#  SESSION INIT
# ══════════════════════════════════════════════════════════════════════
for k, v in [("logged_in", False), ("role", None), ("agent_code", None), ("user_name", ""), ("user_email", "")]:
    if k not in st.session_state:
        st.session_state[k] = v

# ══════════════════════════════════════════════════════════════════════
#  LOGIN
# ══════════════════════════════════════════════════════════════════════
if not st.session_state.logged_in:
    st.markdown('<div class="login-wrap">', unsafe_allow_html=True)
    st.markdown("""
    <div class="login-box">
      <div style="text-align:center;margin-bottom:24px">
        <div style="font-family:'Syne',sans-serif;font-size:2.6rem;font-weight:900;
                    background:linear-gradient(135deg,#3b82f6,#06b6d4);
                    -webkit-background-clip:text;-webkit-text-fill-color:transparent;
                    letter-spacing:5px">APPLY</div>
        <div style="color:#94a3b8;font-size:.9rem;letter-spacing:1px;margin-top:6px">
          Sales Management Portal
        </div>
      </div>
    </div>""", unsafe_allow_html=True)
    email = st.text_input("📧 Email", placeholder="agent250712@apply.com")
    password = st.text_input("🔒 Password", type="password", placeholder="••••••••")
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("Sign In →", use_container_width=True):
        e = email.strip().lower()
        users = _load_users()
        if e == ADMIN_EMAIL and password == ADMIN_PASS:
            st.session_state.update(logged_in=True, role="admin", user_email=e, user_name="Admin")
            st.rerun()
        elif e in users and users[e]["password"] == hash_pw(password):
            u = users[e]
            st.session_state.update(logged_in=True, role=u["role"], user_email=e, agent_code=u.get("agent_code"), user_name=u["name"])
            st.rerun()
        else:
            st.error("❌ Incorrect email or password.")
    st.markdown("</div>", unsafe_allow_html=True)
    st.stop()

# Auto-migrate legacy feedback values
if os.path.exists(DATA_FILE):
    _mr = _load_records()
    _mn = _migrate_feedback(_mr)
    if _mn > 0:
        _save_records(_mr)

# Seed on first run
if not os.path.exists(DATA_FILE) and os.path.exists(EXCEL_SEED):
    _seed = _normalise_df(pd.read_excel(EXCEL_SEED))
    _recs = _seed.to_dict(orient="records")
    _save_records(_recs)
    _auto_create_agents(_recs)
    _save_df(_seed)

# ══════════════════════════════════════════════════════════════════════
#  SIDEBAR
# ══════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(f"""
    <div style="padding:18px 4px 12px;text-align:center">
      <div style="font-family:'Syne',sans-serif;font-size:1.65rem;font-weight:900;
                  background:linear-gradient(135deg,#3b82f6,#06b6d4);
                  -webkit-background-clip:text;-webkit-text-fill-color:transparent;
                  letter-spacing:4px">APPLY</div>
    </div>
    <div style="background:#f0f5ff;border:1px solid #dbeafe;border-radius:12px;
                padding:12px 14px;margin-bottom:14px">
      <div style="font-weight:700;color:#1e2d45;font-size:.95rem">{st.session_state.user_name}</div>
      <div style="color:#64748b;font-size:.78rem;margin-top:2px">
        {'🛡️ Administrator' if st.session_state.role == 'admin' else f'👤 Agent · {st.session_state.agent_code}'}
      </div>
    </div>""", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("**📁 Load JSON Data**")

    if st.session_state.use_uploaded_json:
        st.warning("⚠️ وضع العرض المؤقت - التعديلات مش بتتحفظ تلقائياً")
        st.success(f"✅ جاري استخدام JSON المرفوع ({len(st.session_state.uploaded_json_data)} سجل)")
        col1, col2 = st.columns(2)
        with col1:
            json_str = json.dumps(st.session_state.uploaded_json_data, ensure_ascii=False, indent=2)
            st.download_button("💾 Export JSON", data=json_str, file_name=f"data_export_{datetime.now():%Y%m%d_%H%M}.json", mime="application/json", key="json_download_temp")
        with col2:
            if st.button("🔗 Merge Local", use_container_width=True):
                local_data = []
                if os.path.exists(DATA_FILE):
                    with open(DATA_FILE, encoding="utf-8") as f:
                        local_data = json.load(f)
                existing_mobiles = {str(r.get("Mobile", "")).strip().lstrip("0") for r in local_data}
                new_records = []
                for r in st.session_state.uploaded_json_data:
                    mob_key = str(r.get("Mobile", "")).strip().lstrip("0")
                    if mob_key and mob_key not in existing_mobiles:
                        new_records.append(r)
                        existing_mobiles.add(mob_key)
                merged = local_data + new_records
                with open(DATA_FILE, "w", encoding="utf-8") as f:
                    json.dump(merged, f, ensure_ascii=False, indent=2)
                _auto_create_agents(merged)
                st.session_state.use_uploaded_json = False
                st.session_state.uploaded_json_data = None
                st.success(f"✅ تم دمج {len(new_records)} سجل جديد!")
                st.rerun()
        if st.button("🔄 Local Data Only", use_container_width=True):
            st.session_state.use_uploaded_json = False
            st.session_state.uploaded_json_data = None
            st.rerun()
    else:
        st.info("📄 Using local data from apply_data/")
        local_count = len(_load_records())
        st.caption(f"Local records: {local_count}")
        uploaded_json = st.file_uploader("Upload JSON file", type=["json"], key="json_uploader", label_visibility="collapsed")
        if uploaded_json is not None:
            try:
                data = json.load(uploaded_json)
                records = extract_records_from_json(data)
                if records is not None and isinstance(records, list):
                    st.session_state.uploaded_json_data = records
                    st.session_state.use_uploaded_json = True
                    st.success(f"✅ Loaded {len(records)} records!")
                    st.rerun()
                else:
                    st.error("Invalid JSON format.")
            except Exception as e:
                st.error(f"Error reading JSON: {e}")

    st.markdown("---")
    pages = (["Dashboard", "All Data", "Upload Data", "Users", "Monthly Report"] if st.session_state.role == "admin" else ["My Dashboard", "My Clients"])
    page = st.radio("nav", pages, label_visibility="collapsed")

    st.markdown("---")
    if st.button("Sign Out", use_container_width=True):
        for k in ["logged_in","role","agent_code","user_name","user_email"]:
            st.session_state[k] = False if k == "logged_in" else None if k in ["role","agent_code"] else ""
        st.rerun()

# ══════════════════════════════════════════════════════════════════════
#  SHARED STATS RENDERERS
# ══════════════════════════════════════════════════════════════════════
def _render_fb_cards(records: list):
    s = _get_stats(records)
    total = s["total"]
    if not total: return
    cards = ""
    for fb in FEEDBACK_OPTIONS:
        cnt = s["fb"].get(fb, 0)
        color = FB_COLORS.get(fb, "#475569")
        bg = FB_BG.get(fb, "#f8fafc")
        cards += f'<div class="stat-card" style="border-left:3px solid {color};background:{bg}"><div class="stat-num" style="color:{color}">{cnt}</div><div class="stat-lbl" style="color:{color}">{fb}</div><div style="font-size:12px;color:#94a3b8;margin-top:4px">{pct(cnt,total)}%</div></div>'
    st.markdown(f'<div class="stat-cards">{cards}</div>', unsafe_allow_html=True)

def _render_ds_cards(records: list):
    s = _get_stats(records)
    total = s["total"]
    if not total: return
    all_ds = set(str(r.get("Data Source Feedback","")).strip() for r in records) | set(DS_OPTIONS)
    all_ds.discard(""); all_ds.discard("None"); all_ds.discard("nan")
    cards = ""
    for ds in sorted(all_ds):
        cnt = s["ds"].get(ds, 0)
        if cnt == 0: continue
        color = DS_COLORS.get(ds, "#475569")
        bg = DS_BG.get(ds, "#f8fafc")
        cards += f'<div class="stat-card" style="border-left:3px solid {color};background:{bg}"><div class="stat-num" style="color:{color}">{cnt}</div><div class="stat-lbl" style="color:{color}">{ds}</div><div style="font-size:12px;color:#94a3b8;margin-top:4px">{pct(cnt,total)}%</div></div>'
    if cards:
        st.markdown(f'<div class="stat-cards">{cards}</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════
#  ADMIN PAGES
# ══════════════════════════════════════════════════════════════════════
def page_dashboard():
    records = _load_records()
    s = _get_stats(records)
    if st.session_state.use_uploaded_json:
        st.info("ℹ️ أنت في وضع العرض المؤقت - البيانات من JSON المرفوع")
    st.markdown(f'<div class="ap-hero"><div class="ap-hero-title">🛡️ Admin Dashboard</div><div class="ap-hero-sub">Total records: {s["total"]}</div></div>', unsafe_allow_html=True)
    k1, k2 = st.columns(2)
    k1.metric("Total", s["total"])
    k2.metric("Done ✅", s["done"], f"{pct(s['done'], s['total'])}%")
    if not records:
        st.info("No data yet — upload a client sheet to get started.")
        return
    st.markdown('<div class="sec-title">Feedback Breakdown</div>', unsafe_allow_html=True)
    _render_fb_cards(records)
    st.markdown('<div class="sec-title">Data Source Breakdown</div>', unsafe_allow_html=True)
    _render_ds_cards(records)
    st.markdown('<div class="sec-title">KPI Bars</div>', unsafe_allow_html=True)
    st.markdown(f"""
    <div class="kpi-grid">
      <div class="kpi-card"><div class="kpi-ttl">Done Rate</div><div class="kpi-val" style="color:#15803d">{pct(s['done'], s['total'])}%</div><div class="kpi-sub">{s['done']} of {s['total']}</div>{bar_html(s['done'], s['total'], '#15803d')}</div>
      <div class="kpi-card"><div class="kpi-ttl">Recall Rate</div><div class="kpi-val" style="color:#b45309">{pct(s['recall'], s['total'])}%</div><div class="kpi-sub">{s['recall']} clients</div>{bar_html(s['recall'], s['total'], '#b45309')}</div>
      <div class="kpi-card"><div class="kpi-ttl">Closed Rate</div><div class="kpi-val" style="color:#b91c1c">{pct(s['closed'], s['total'])}%</div><div class="kpi-sub">{s['closed']} clients</div>{bar_html(s['closed'], s['total'], '#b91c1c')}</div>
      <div class="kpi-card"><div class="kpi-ttl">Not Interested</div><div class="kpi-val" style="color:#475569">{pct(s['ni'], s['total'])}%</div><div class="kpi-sub">{s['ni']} clients</div>{bar_html(s['ni'], s['total'], '#475569')}</div>
    </div>""", unsafe_allow_html=True)
    
    st.markdown('<div class="sec-title">Export Options</div>', unsafe_allow_html=True)
    c1, c2, _ = st.columns([1, 1, 2])
    with c1:
        st.download_button("⬇️ Download HTML Report", data=_build_html_report(records), file_name=f"apply_report_{datetime.now():%Y%m%d_%H%M}.html", mime="text/html", use_container_width=True)
    with c2:
        st.download_button("📊 Download Excel Dashboard", data=_build_excel_dashboard(records), file_name=f"apply_dashboard_{datetime.now():%Y%m%d_%H%M}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

def page_all_data():
    st.markdown("## 📊 All Data")
    records = _load_records()
    if not records:
        st.info("No records available.")
        return
    df = pd.DataFrame(records)
    st.dataframe(df, use_container_width=True)

def page_upload_data():
    st.markdown("## 📤 Upload Data")
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        df_clean = _normalise_df(df)
        recs = df_clean.to_dict(orient="records")
        if st.button("Save Uploaded Data"):
            _save_records(recs)
            _auto_create_agents(recs)
            _save_df(df_clean)
            st.success("Data successfully saved!")

def page_users():
    st.markdown("## 👥 User Management")
    users = _load_users()
    st.json(users)

def page_monthly_report():
    st.markdown("## 📅 Monthly Report")
    st.info("Monthly performance stats summary.")

# ══════════════════════════════════════════════════════════════════════
#  AGENT PAGES
# ══════════════════════════════════════════════════════════════════════
def page_agent_dashboard():
    records = [r for r in _load_records() if str(r.get("Agent Code", "")).strip() == str(st.session_state.agent_code)]
    st.markdown(f"## 👤 My Dashboard (Agent {st.session_state.agent_code})")
    _render_fb_cards(records)

def page_agent_clients():
    st.markdown("## 📋 My Clients")
    records = [r for r in _load_records() if str(r.get("Agent Code", "")).strip() == str(st.session_state.agent_code)]
    if records:
        st.dataframe(pd.DataFrame(records), use_container_width=True)
    else:
        st.info("No clients assigned.")

# ══════════════════════════════════════════════════════════════════════
#  ROUTER
# ══════════════════════════════════════════════════════════════════════
if st.session_state.role == "admin":
    if page == "Dashboard": page_dashboard()
    elif page == "All Data": page_all_data()
    elif page == "Upload Data": page_upload_data()
    elif page == "Users": page_users()
    elif page == "Monthly Report": page_monthly_report()
else:
    if page == "My Dashboard": page_agent_dashboard()
    elif page == "My Clients": page_agent_clients()
