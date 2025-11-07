import io, os, re, csv, json, yaml, hashlib, time, shutil
from datetime import datetime
import pandas as pd
import streamlit as st
from bs4 import BeautifulSoup
from mailparser import MailParser
from pdfminer.high_level import extract_text as pdf_extract_text
import mammoth

# Αν τρέχει σε Windows host, χρησιμοποιούμε python-magic-bin από requirements
try:
    import magic
    HAS_MAGIC = True
except Exception:
    HAS_MAGIC = False

# ------------- Ρυθμίσεις μόνιμης αποθήκευσης -------------
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
STORAGE_DIR = os.path.join(BASE_DIR, "storage")
BANKS_DIR   = os.path.join(STORAGE_DIR, "banks")
REPORTS_DIR = os.path.join(STORAGE_DIR, "reports")
os.makedirs(BANKS_DIR, exist_ok=True)
os.makedirs(REPORTS_DIR, exist_ok=True)

def now_tag():
    return datetime.utcnow().strftime("%Y%m%d-%H%M%S")

def file_sha(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()[:16]

# ------------- Βοηθητικά -------------
def normalize(s: str) -> str:
    return (s or "").replace("\xa0", " ").replace("\r", " ")

def to_number(x):
    s = str(x)
    s = re.sub(r"[^0-9,.\-]", "", s)
    s = re.sub(r"\.(?=.*\.)", "", s)  # κράτα το τελευταίο dot
    s = s.replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0

def money(n): return f"{n:,.2f}"

def guess_mime(name: str, data: bytes) -> str:
    if HAS_MAGIC:
        try:
            return magic.Magic(mime=True).from_buffer(data[:2048] if data else b"")
        except Exception:
            pass
    # fallback
    n = name.lower()
    if n.endswith(".pdf"): return "application/pdf"
    if n.endswith(".docx"): return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    if n.endswith(".eml"): return "message/rfc822"
    if n.endswith(".csv"): return "text/csv"
    if n.endswith(".xlsx"): return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    if n.endswith(".xls"): return "application/vnd.ms-excel"
    if n.endswith(".txt"): return "text/plain"
    return "application/octet-stream"

# ------------- Extractors -------------
def extract_text_from_file(uploaded) -> str:
    data = uploaded.read()
    uploaded.seek(0)
    mime = guess_mime(uploaded.name, data)
    name = uploaded.name.lower()

    # PDF (όχι OCR για σκαναρισμένα)
    if name.endswith(".pdf") or mime == "application/pdf":
        return pdf_extract_text(io.BytesIO(data))

    # DOCX
    if name.endswith(".docx") or "officedocument.wordprocessingml.document" in mime:
        res = mammoth.convert_to_html(io.BytesIO(data))
        html = res.value
        soup = BeautifulSoup(html, "html.parser")
        return soup.get_text(" ")

    # EML (email)
    if name.endswith(".eml") or mime in ("message/rfc822",):
        mail = MailParser()
        mail.parse_from_bytes(data)
        parts = [mail.subject or "", mail.body or ""]
        return "\n".join(parts)

    # TXT / CSV (ως κείμενο)
    if name.endswith(".txt") or name.endswith(".csv") or mime.startswith("text/"):
        try:
            return data.decode("utf-8", errors="ignore")
        except:
            return data.decode("latin-1", errors="ignore")

    # Excel → μετατροπή σε text (όλες οι στήλες σαν κείμενο)
    if name.endswith(".xlsx"):
        df = pd.read_excel(io.BytesIO(data), engine="openpyxl")
        return df.astype(str).to_csv(index=False)
    if name.endswith(".xls"):
        df = pd.read_excel(io.BytesIO(data), engine="xlrd")
        return df.astype(str).to_csv(index=False)

    return ""

def parse_rules_yaml_bytes(file_bytes: bytes) -> dict:
    return yaml.safe_load(file_bytes.decode("utf-8", errors="ignore"))

def parse_volumes_csv_bytes(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_csv(io.BytesIO(file_bytes))
    cols = {c.lower(): c for c in df.columns}
    type_col = cols.get("type") or list(df.columns)[0]
    count_col = cols.get("count") or (list(df.columns)[1] if len(df.columns)>1 else None)
    amt_col = cols.get("amount_total") or cols.get("total_amount") or (list(df.columns)[2] if len(df.columns)>2 else None)
    out = pd.DataFrame({
        "type": df[type_col].astype(str).str.strip(),
        "count": pd.to_numeric(df[count_col], errors="coerce").fillna(0) if count_col else 0,
        "amount_total": pd.to_numeric(df[amt_col], errors="coerce").fillna(0.0) if amt_col else 0.0
    })
    return out

def parse_invoice_lines(text: str) -> pd.DataFrame:
    norm = normalize(text).lower()
    items = []

    pat1 = re.compile(r"(.*?)(€|eur)\s*([0-9]+[.,][0-9]{1,4}|[0-9]+)(?![\d.,])", re.IGNORECASE)
    pat2 = re.compile(r"(.*?)([0-9]+[.,][0-9]{1,4}|[0-9]+)\s*(€|eur)", re.IGNORECASE)

    for m in pat1.finditer(norm):
        desc = (m.group(1) or "").strip().replace("•","").replace("–","-")
        amt = to_number(m.group(3))
        items.append((desc[-140:], amt, "EUR"))

    for m in pat2.finditer(norm):
        desc = (m.group(1) or "").strip().replace("•","").replace("–","-")
        amt = to_number(m.group(2))
        items.append((desc[-140:], amt, "EUR"))

    df = pd.DataFrame(items, columns=["description","amount","currency"])
    if df.empty:
        return df
    grouped = df.groupby("description", as_index=False).agg({"amount":"sum","currency":"first"})
    return grouped

# ------------- Υπολογισμοί -------------
def guess_count(vols: pd.DataFrame) -> float:
    if vols is None or vols.empty: return 0
    inc = vols[vols["type"].str.upper().str.startswith("INCOMING")]
    if not inc.empty: return float(inc["count"].sum())
    swift = vols[vols["type"].str.upper().str.contains("SWIFT")]
    if not swift.empty: return float(swift["count"].sum())
    return float(vols["count"].sum())

def guess_amount(vols: pd.DataFrame) -> float:
    if vols is None or vols.empty: return 0.0
    inc = vols[vols["type"].str.upper().str.startswith("INCOMING")]
    if not inc.empty: return float(inc["amount_total"].sum())
    return float(vols["amount_total"].sum())

def expected_amount(rule: dict, vols: pd.DataFrame) -> float:
    mode = rule.get("mode")
    val = float(rule.get("value", 0) or 0)
    if mode == "per_item":
        return round((guess_count(vols) * val), 2)
    if mode == "percent":
        base = guess_amount(vols)
        res = base * val
        if rule.get("min") is not None: res = max(res, float(rule["min"]))
        if rule.get("max") is not None: res = min(res, float(rule["max"]))
        return round(res, 2)
    return round(val, 2)

def audit(rules_cfg: dict, invoice_df: pd.DataFrame, vols_df: pd.DataFrame):
    currency = rules_cfg.get("currency", "EUR")
    tol_pct = float(rules_cfg.get("tolerance_pct", 1.0)) / 100.0
    rules = rules_cfg.get("rules", [])

    matches, missing, extras = [], [], invoice_df.copy()

    for r in rules:
        keywords = [k.lower() for k in (r.get("keywords") or []) if k]
        hit = None
        if not invoice_df.empty:
            for _, row in invoice_df.iterrows():
                desc = str(row["description"]).lower()
                if any(k in desc for k in keywords) or r.get("label","").lower() in desc:
                    hit = row; break

        if hit is not None:
            exp = expected_amount(r, vols_df)
            diff = round(float(hit["amount"]) - exp, 2)
            if r.get("mode") == "per_item":
                note = f"count={int(guess_count(vols_df))} × unit={r.get('value')}"
            elif r.get("mode") == "percent":
                note = f"{float(r.get('value'))*100:.4f}% επί ποσού {money(guess_amount(vols_df))} (min={r.get('min','—')}, max={r.get('max','—')})"
            else:
                note = f"πάγιο {r.get('value')}"

            matches.append({
                "rule_code": r.get("code"),
                "rule_label": r.get("label"),
                "invoice_desc": hit["description"],
                "expected": exp,
                "invoiced": float(hit["amount"]),
                "diff": diff,
                "note": note
            })
            extras = extras[extras["description"] != hit["description"]]
        else:
            missing.append({"rule_code": r.get("code"), "rule_label": r.get("label")})

    totals = {
        "expected": round(sum(expected_amount(r, vols_df) for r in rules), 2),
        "invoiced": round(float(invoice_df["amount"].sum()) if not invoice_df.empty else 0.0, 2),
    }

    return {
        "currency": currency,
        "tolerance_pct": tol_pct,
        "matches": matches,
        "missing": missing,
        "extras": extras.to_dict(orient="records"),
        "totals": totals
    }

def build_report_csv(res: dict) -> bytes:
    rows = [["type","rule_code","rule_label","invoice_desc","expected","invoiced","diff","note"]]
    for m in res["matches"]:
        rows.append(["match", m["rule_code"], m["rule_label"], m["invoice_desc"], m["expected"], m["invoiced"], m["diff"], m["note"]])
    for m in res["missing"]:
        rows.append(["missing", m["rule_code"], m["rule_label"], "", "", "", "", "no invoice line"])
    for x in res["extras"]:
        rows.append(["extra", "", "", x["description"], "", x["amount"], "", "no matching rule"])
    rows.append(["totals","","","", res["totals"]["expected"], res["totals"]["invoiced"], round(res["totals"]["invoiced"]-res["totals"]["expected"],2), ""])
    buf = io.StringIO(); w = csv.writer(buf); w.writerows(rows)
    return buf.getvalue().encode("utf-8")

# ------------- Αποθήκευση/ανάγνωση “κανόνων ανά Τράπεζα” -------------
def bank_path(bank_name: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", bank_name.strip())
    p = os.path.join(BANKS_DIR, safe)
    os.makedirs(p, exist_ok=True)
    return p

def load_bank_index(bank_name: str) -> dict:
    p = bank_path(bank_name)
    idx_file = os.path.join(p, "index.json")
    if os.path.exists(idx_file):
        return json.load(open(idx_file, "r", encoding="utf-8"))
    return {"active": None, "versions": []}

def save_bank_index(bank_name: str, data: dict):
    p = bank_path(bank_name)
    idx_file = os.path.join(p, "index.json")
    with open(idx_file, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def save_rules_file(bank_name: str, uploaded) -> dict:
    data = uploaded.read(); uploaded.seek(0)
    p = bank_path(bank_name)
    ver = now_tag() + "_" + file_sha(data) + os.path.splitext(uploaded.name)[1].lower()
    full = os.path.join(p, ver)
    with open(full, "wb") as f: f.write(data)
    idx = load_bank_index(bank_name)
    idx["versions"].insert(0, {"file": ver, "original": uploaded.name, "size": len(data), "ts": time.time()})
    # δεν αλλάζουμε αυτόματα το active — ο χρήστης πατά “Ορισμός ως ενεργό”
    save_bank_index(bank_name, idx)
    return {"file": ver, "path": full}

def set_active_rules(bank_name: str, filename: str):
    idx = load_bank_index(bank_name)
    idx["active"] = filename
    save_bank_index(bank_name, idx)

def get_active_rules_text(bank_name: str) -> str:
    idx = load_bank_index(bank_name)
    if not idx.get("active"): return ""
    full = os.path.join(bank_path(bank_name), idx["active"])
    if not os.path.exists(full): return ""
    # Επιστρέφει ως κείμενο (για εμφάνιση). Για YAML κάνει parse ο χρήστης ξεχωριστά.
    with open(full, "rb") as f:
        fake_upload = type("U", (), {"read": lambda self=f: self.read(), "seek": lambda *args, **kwargs: None, "name": full})
        return extract_text_from_file(fake_upload)

def get_active_rules_yaml(bank_name: str) -> dict | None:
    """Αν το ενεργό αρχείο είναι YAML, το επιστρέφει ως dict, αλλιώς None."""
    idx = load_bank_index(bank_name)
    if not idx.get("active"): return None
    full = os.path.join(bank_path(bank_name), idx["active"])
    ext = os.path.splitext(full)[1].lower()
    if ext in (".yaml", ".yml"):
        return yaml.safe_load(open(full, "r", encoding="utf-8").read())
    return None

# ------------- UI (Streamlit) -------------
st.set_page_config(page_title="Έλεγχος Τιμολόγησης Τραπεζών", layout="wide")
st.title("Έλεγχος Τιμολόγησης Τραπεζών — Αποθήκευση Κανόνων & Audit (MVP)")

with st.sidebar:
    st.header("🧭 Πλοήγηση")
    page = st.radio("Μενού", ["1) Κανόνες ανά Τράπεζα", "2) Έλεγχος Τιμολόγησης"])

# --- 1) Διαχείριση Κανόνων ---
if page.startswith("1"):
    st.subheader("1) Κανόνες Τιμολόγησης ανά Τράπεζα")

    # Επιλογή/Εισαγωγή ονόματος Τράπεζας
    banks = sorted([d for d in os.listdir(BANKS_DIR) if os.path.isdir(os.path.join(BANKS_DIR, d))])
    colA, colB = st.columns([1,1])
    with colA:
        bank = st.selectbox("Διάλεξε τράπεζα (υπάρχουσα)", banks) if banks else None
    with colB:
        new_bank = st.text_input("Ή γράψε νέα τράπεζα", placeholder="π.χ. CRBAGRAAXXX ή ALPHA_BANK")
        if new_bank:
            bank = new_bank

    if not bank:
        st.info("Διάλεξε/γράψε όνομα τράπεζας για να συνεχίσεις.")
        st.stop()

    st.markdown(f"**Τρέχουσα τράπεζα:** `{bank}`")
    idx = load_bank_index(bank)

    st.markdown("### ➕ Ανέβασμα/Αποθήκευση νέου αρχείου κανόνων")
    rules_up = st.file_uploader("Αρχείο Κανόνων (PDF / DOCX / EML / TXT / YAML / XLSX / XLS / CSV)", type=["pdf","docx","eml","txt","yaml","yml","xlsx","xls","csv"])
    if rules_up is not None and st.button("Αποθήκευση αρχείου κανόνων"):
        saved = save_rules_file(bank, rules_up)
        st.success(f"Αποθηκεύτηκε: {saved['file']}")
        st.experimental_rerun()

    st.markdown("### 📚 Εκδόσεις κανόνων που έχουν αποθηκευτεί")
    if not idx["versions"]:
        st.info("Δεν υπάρχουν αποθηκευμένοι κανόνες για αυτή την τράπεζα.")
    else:
        for v in idx["versions"]:
            c1, c2, c3, c4 = st.columns([3,2,2,2])
            with c1:
                st.write(f"**{v['original']}**  \n_({v['file']})_")
            with c2:
                st.write(f"Μέγεθος: {v['size']} bytes")
            with c3:
                st.write(f"Ημ/νία: {datetime.fromtimestamp(v['ts']).strftime('%Y-%m-%d %H:%M:%S')}")
            with c4:
                if st.button("Ορισμός ως ενεργό", key=f"set_{v['file']}"):
                    set_active_rules(bank, v["file"])
                    st.success("Ορίστηκε ως ενεργό.")
                    st.experimental_rerun()

    st.markdown("### ✅ Ενεργό αρχείο κανόνων")
    if idx.get("active"):
        st.success(f"Ενεργό: {idx['active']}")
        # Προεπισκόπηση ως κείμενο
        with st.expander("Προεπισκόπηση (ως κείμενο)"):
            st.text(get_active_rules_text(bank)[:4000] or "—")
    else:
        st.warning("Δεν έχει οριστεί ενεργό αρχείο κανόνων. Επέλεξε μία έκδοση και πάτα «Ορισμός ως ενεργό».")

# --- 2) Έλεγχος Τιμολόγησης ---
if page.startswith("2"):
    st.subheader("2) Έλεγχος Τιμολογίου με βάση το ενεργό αρχείο κανόνων")

    # Επιλογή τράπεζας
    banks = sorted([d for d in os.listdir(BANKS_DIR) if os.path.isdir(os.path.join(BANKS_DIR, d))])
    if not banks:
        st.info("Δεν έχεις τράπεζες. Πήγαινε στο βήμα «1) Κανόνες ανά Τράπεζα» και πρόσθεσε.")
        st.stop()

    bank = st.selectbox("Διάλεξε τράπεζα", banks)
    idx = load_bank_index(bank)
    if not idx.get("active"):
        st.warning("Για αυτή την τράπεζα δεν υπάρχει ενεργό αρχείο κανόνων. Όρισε πρώτα στο μενού 1).")
        st.stop()

    active_yaml = get_active_rules_yaml(bank)
    st.caption("Αν το ενεργό αρχείο είναι YAML, χρησιμοποιούνται οι κανόνες από εκεί. Αλλιώς επιχειρείται απλό text-based parsing.")

    # Upload invoice
    inv = st.file_uploader("Αρχείο Τιμολόγησης (PDF / DOCX / TXT / CSV / XLSX / XLS)", type=["pdf","docx","txt","csv","xlsx","xls"])
    invoice_df = pd.DataFrame()
    if inv is not None:
        if inv.name.lower().endswith(".csv"):
            tmp = pd.read_csv(inv)
            cols = {c.lower(): c for c in tmp.columns}
            invoice_df = pd.DataFrame({
                "description": tmp[cols.get("description", tmp.columns[0])].astype(str),
                "amount": pd.to_numeric(tmp[cols.get("amount", tmp.columns[1])], errors="coerce").fillna(0.0),
                "currency": "EUR"
            })
        elif inv.name.lower().endswith((".xlsx",".xls")):
            xdf = pd.read_excel(inv, engine=("openpyxl" if inv.name.lower().endswith(".xlsx") else "xlrd"))
            # Προσπαθούμε να βρούμε περιγραφή/ποσό
            cols = {c.lower(): c for c in xdf.columns}
            desc_col = cols.get("description") or cols.get("περιγραφή") or xdf.columns[0]
            amount_col = cols.get("amount") or cols.get("ποσό") or (xdf.columns[1] if len(xdf.columns)>1 else xdf.columns[0])
            invoice_df = pd.DataFrame({
                "description": xdf[desc_col].astype(str),
                "amount": pd.to_numeric(xdf[amount_col], errors="coerce").fillna(0.0),
                "currency": "EUR"
            })
        else:
            text = extract_text_from_file(inv)
            invoice_df = parse_invoice_lines(text)

    # Volumes (optional)
    vol = st.file_uploader("Volumes (CSV προαιρετικό: type,count,amount_total)", type=["csv"])
    volumes_df = pd.DataFrame()
    if vol is not None:
        volumes_df = parse_volumes_csv_bytes(vol.read())

    # Πλήκτρο ελέγχου
    can_run = (invoice_df is not None) and (not invoice_df.empty)
    run = st.button("▶️ Τρέξε Έλεγχο", disabled=not can_run)

    if not can_run:
        st.info("Φόρτωσε ένα αρχείο τιμολογίου για να ενεργοποιηθεί ο έλεγχος.")
    elif run:
        # κανόνες: αν ενεργό YAML → πάρ’ το, αλλιώς ένα ελάχιστο default για demo
        if active_yaml:
            rules_cfg = active_yaml
        else:
            st.warning("Το ενεργό αρχείο κανόνων δεν είναι YAML. Θα χρησιμοποιηθεί απλό default demo config.")
            rules_cfg = {
                "currency": "EUR",
                "tolerance_pct": 1,
                "rules": [
                    {"code": "R1","label":"Χρέωση ανά εντολή","mode":"per_item","value":0.30,"keywords":["ανά εντολή","incoming","swift"]},
                    {"code": "R2","label":"Ποσοστιαία","mode":"percent","value":0.001,"min":2,"max":50,"keywords":["%","ad valorem"]},
                    {"code": "R3","label":"Πάγιο","mode":"flat","value":100,"keywords":["πάγιο","monthly"]}
                ]
            }

        result = audit(rules_cfg, invoice_df, volumes_df)

        # Σύνολα
        c1, c2 = st.columns(2)
        with c1:
            st.metric("Σύνολο τιμολογίου", f"{money(result['totals']['invoiced'])} {rules_cfg.get('currency','EUR')}")
        with c2:
            st.metric("Αναμενόμενο σύνολο", f"{money(result['totals']['expected'])} {rules_cfg.get('currency','EUR')}")

        st.markdown("### Αντιστοιχίσεις")
        if result["matches"]:
            st.dataframe(pd.DataFrame(result["matches"]))
        else:
            st.info("—")

        a, b = st.columns(2)
        with a:
            st.markdown("### Κανόνες χωρίς αντίστοιχη χρέωση")
            st.dataframe(pd.DataFrame(result["missing"]) if result["missing"] else pd.DataFrame())
        with b:
            st.markdown("### Χρεώσεις εκτός κανόνων")
            st.dataframe(pd.DataFrame(result["extras"]) if result["extras"] else pd.DataFrame())

        # Αποθήκευση αναφοράς
        rep_bytes = build_report_csv(result)
        rep_name = f"report_{bank}_{now_tag()}.csv"
        # γράψε στο storage/reports
        with open(os.path.join(REPORTS_DIR, rep_name), "wb") as f:
            f.write(rep_bytes)
        st.download_button("⬇️ Λήψη Αναφοράς CSV", data=rep_bytes, file_name=rep_name, mime="text/csv")

        st.success("Ο έλεγχος ολοκληρώθηκε.")
