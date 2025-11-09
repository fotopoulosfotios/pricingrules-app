import io, os, re, json, yaml
import pandas as pd
import streamlit as st
import chardet
from pdfminer.high_level import extract_text as pdf_extract_text
from docx import Document

# ---------------------------
#  Ρυθμίσεις σελίδας
# ---------------------------
st.set_page_config(page_title="Έλεγχος Τιμολογήσεων", page_icon="💶", layout="wide")
st.title("💶 Σύστημα Ελέγχου Τιμολογήσεων Εισερχόμενων Εντολών")

# ---------------------------
#  Βοηθητικά: Κανόνες
# ---------------------------
def load_rules_from_path(path="rules_example.yaml"):
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
        return data.get("rules", [])
    except Exception as e:
        st.warning(f"Δεν βρέθηκαν προεπιλεγμένοι κανόνες ({path}): {e}")
        return []

def validate_rules(rules):
    required = {"id", "description", "condition", "action"}
    for i, r in enumerate(rules, start=1):
        if not isinstance(r, dict):
            return f"Κανόνας #{i} δεν είναι αντικείμενο."
        missing = required - set(r.keys())
        if missing:
            return f"Κανόνας #{i} λείπει πεδία: {', '.join(missing)}"
        if "field" not in (r.get("condition") or {}):
            return f"Κανόνας #{i} λείπει condition.field"
    return None

# ---------------------------
#  Βοηθητικά: Ονοματολογία πεδίων
# ---------------------------
REQUIRED_COLS = ["BIC", "Currency", "Amount", "ChargeBearer"]

SYNONYMS = {
    "bic": ["bic", "sender_bic", "bank_bic"],
    "currency": ["currency", "ccy", "curr"],
    "amount": ["amount", "amt", "value", "sum"],
    "chargebearer": ["chargebearer", "charge_bearer", "charge-bearer", "charges", "expenses_type"],
}

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols_map = {}
    lower_cols = {c.lower(): c for c in df.columns}
    for target, alts in SYNONYMS.items():
        for cand in alts:
            if cand in lower_cols:
                cols_map[lower_cols[cand]] = target.capitalize() if target != "chargebearer" else "ChargeBearer"
                break
    # διατήρησε ό,τι ήδη ταιριάζει κανονικά
    for c in df.columns:
        if c in REQUIRED_COLS:
            cols_map[c] = c
    df = df.rename(columns=cols_map)
    return df

def require_columns(df: pd.DataFrame):
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    return missing

# ---------------------------
#  Αναγνώστες αρχείων
# ---------------------------
def read_csv_file(file) -> pd.DataFrame:
    raw = file.read()
    enc = chardet.detect(raw).get("encoding") or "utf-8"
    return pd.read_csv(io.BytesIO(raw), encoding=enc)

def read_xlsx_file(file) -> pd.DataFrame:
    return pd.read_excel(file)

def read_json_file(file) -> pd.DataFrame:
    try:
        obj = json.load(file)
    except Exception:
        # αν είναι bytes
        obj = json.loads(file.read().decode("utf-8"))
    if isinstance(obj, list):
        return pd.DataFrame(obj)
    elif isinstance(obj, dict):
        # προσπαθεί να βρει πίνακα
        for v in obj.values():
            if isinstance(v, list) and v and isinstance(v[0], dict):
                return pd.DataFrame(v)
        return pd.DataFrame([obj])
    return pd.DataFrame()

def read_txt_file(file) -> pd.DataFrame:
    # προσπαθεί διαχωρισμό με tab/;/
    raw = file.read()
    enc = chardet.detect(raw).get("encoding") or "utf-8"
    text = raw.decode(enc, errors="ignore")
    # δοκίμασε πρώτα tab
    if "\t" in text:
        return pd.read_csv(io.StringIO(text), sep="\t")
    # δοκίμασε ;
    if ";" in text:
        return pd.read_csv(io.StringIO(text), sep=";")
    # fallback: space-split -> σε λίστα
    rows = [re.split(r"\s{2,}|\s*,\s*|\s+", ln.strip()) for ln in text.splitlines() if ln.strip()]
    maxlen = max((len(r) for r in rows), default=0)
    rows = [r + [""]*(maxlen-len(r)) for r in rows]
    return pd.DataFrame(rows)

def parse_table_from_text(text: str) -> pd.DataFrame:
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    # απλή εικασία: διάστημα/κόμμα/; ως διαχωριστές
    rows = [re.split(r"\s{2,}|;|,|\t", ln) for ln in lines]
    # αν η πρώτη γραμμή μοιάζει με header, χρησιμοποίησέ την
    if rows:
        header = [h.strip() for h in rows[0]]
        body = rows[1:]
        # αν τα header είναι πολύ «χαοτικά», φτιάξε generic
        if any(len(h) == 0 for h in header) or len(set(header)) != len(header):
            width = max(len(r) for r in rows)
            cols = [f"Col{i+1}" for i in range(width)]
            body = [r + [""]*(len(cols)-len(r)) for r in rows]
            return pd.DataFrame(body, columns=cols)
        else:
            width = len(header)
            body = [r + [""]*(width-len(r)) for r in body]
            return pd.DataFrame(body, columns=header)
    return pd.DataFrame()

def read_pdf_file(file) -> pd.DataFrame:
    # Προσπάθεια εξαγωγής text από PDF (όχι OCR)
    try:
        text = pdf_extract_text(file)
    except Exception as e:
        st.error(f"Αποτυχία ανάγνωσης PDF: {e}")
        return pd.DataFrame()
    df = parse_table_from_text(text)
    return df

def read_docx_file(file) -> pd.DataFrame:
    # Προσπάθεια ανάγνωσης πινάκων DOCX
    byts = file.read()
    doc = Document(io.BytesIO(byts))
    tables = doc.tables
    if tables:
        tbl = tables[0]
        data = []
        for r in tbl.rows:
            data.append([c.text.strip() for c in r.cells])
        # θεώρησε 1η γραμμή ως header αν φαίνεται κατάλληλη
        if data:
            header = data[0]
            if len(set(header)) == len(header) and all(h.strip() for h in header):
                rows = data[1:]
                rows = [r + [""]*(len(header)-len(r)) for r in rows]
                return pd.DataFrame(rows, columns=header)
        # αλλιώς generic
        width = max((len(r) for r in data), default=0)
        data = [r + [""]*(width-len(r)) for r in data]
        return pd.DataFrame(data)
    # fallback: παράγραφοι
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return parse_table_from_text("\n".join(paras))

def read_eml_file(file) -> pd.DataFrame:
    # mail-parser για .eml
    from mailparser import MailParser
    raw = file.read()
    parser = MailParser()
    parser.parse_from_bytes(raw)
    # προτιμά text/plain
    body = parser.text_plain[0] if parser.text_plain else (parser.body or "")
    return parse_table_from_text(body)

def read_msg_file(file) -> pd.DataFrame:
    # extract-msg για .msg (Outlook)
    import extract_msg
    byts = file.read()
    with open("_tmp_msg.msg", "wb") as f:
        f.write(byts)
    try:
        msg = extract_msg.Message("_tmp_msg.msg")
        body = (msg.body or "").strip()
        return parse_table_from_text(body)
    finally:
        try:
            os.remove("_tmp_msg.msg")
        except Exception:
            pass

# Προαιρετικό OCR μέσω OCR.space (αν υπάρχει API key σε secrets)
def read_with_ocr_if_possible(file_bytes: bytes, filename: str) -> pd.DataFrame:
    api_key = st.secrets.get("OCRSPACE_API_KEY", None)
    if not api_key:
        st.info("🔍 Το αρχείο φαίνεται σκαναρισμένο/εικόνα. "
                "Για αυτόματο OCR πρόσθεσε OCRSPACE_API_KEY στα Streamlit secrets.")
        return pd.DataFrame()
    try:
        import requests
        files = {"file": (filename, io.BytesIO(file_bytes))}
        data = {"apikey": api_key, "language": "eng"}
        r = requests.post("https://api.ocr.space/parse/image", files=files, data=data, timeout=60)
        js = r.json()
        text = "\n".join([p.get("ParsedText", "") for p in js.get("ParsedResults", [])])
        return parse_table_from_text(text)
    except Exception as e:
        st.error(f"OCR αποτυχία: {e}")
        return pd.DataFrame()

def detect_scanned_pdf_need_ocr(df: pd.DataFrame) -> bool:
    # χονδρική ένδειξη: PDF έβγαλε πολύ λίγο/άχρηστο κείμενο
    return df.empty or (df.shape[1] <= 2 and df.shape[0] <= 3)

# ---------------------------
#  Εφαρμογή κανόνων
# ---------------------------
def apply_rules(df: pd.DataFrame, rules):
    out = []
    for i, row in df.iterrows():
        matched = []
        for rule in rules:
            cond = rule["condition"]
            field = cond.get("field")
            val = row.get(field)
            try:
                if "equals" in cond and str(val) == str(cond["equals"]):
                    matched.append(rule["id"])
                elif "in" in cond and str(val) in [str(x) for x in cond["in"]]:
                    matched.append(rule["id"])
                elif "greater_than" in cond and float(val) > float(cond["greater_than"]):
                    matched.append(rule["id"])
            except Exception:
                pass
        out.append({
            "Index": i+1,
            "BIC": row.get("BIC", ""),
            "Currency": row.get("Currency", ""),
            "Amount": row.get("Amount", ""),
            "ChargeBearer": row.get("ChargeBearer", ""),
            "MatchedRules": ", ".join(matched) if matched else "—"
        })
    return pd.DataFrame(out)

# ---------------------------
#  Sidebar: Κανόνες
# ---------------------------
st.sidebar.header("⚙️ Κανόνες")
rules_file = st.sidebar.file_uploader("Ανέβασε YAML κανόνων", type=["yaml", "yml"])
if rules_file:
    try:
        rules = (yaml.safe_load(rules_file) or {}).get("rules", [])
        msg = validate_rules(rules)
        if msg:
            st.sidebar.error(f"❌ Σφάλμα κανόνων: {msg}")
            rules = []
        else:
            st.sidebar.success("✅ Κανόνες φορτώθηκαν.")
    except Exception as e:
        st.sidebar.error(f"Μη έγκυρο YAML: {e}")
        rules = []
else:
    rules = load_rules_from_path()
    if rules:
        st.sidebar.info("Χρήση προεπιλεγμένων κανόνων (rules_example.yaml)")
    else:
        st.sidebar.warning("Δεν βρέθηκαν κανόνες.")

# ---------------------------
#  Αρχείο εντολών
# ---------------------------
st.subheader("📂 Ανέβασε τιμολογημένες εντολές")
uploaded = st.file_uploader(
    "Υποστηριζόμενα: CSV, XLSX, PDF (text), DOCX, EML, MSG, TXT, JSON. Για σκαναρισμένα: δείτε OCR σημείωση.",
    type=["csv", "xlsx", "pdf", "docx", "eml", "msg", "txt", "json"]
)

df = pd.DataFrame()
if uploaded:
    name = uploaded.name.lower()
    try:
        if name.endswith(".csv"):
            df = read_csv_file(uploaded)
        elif name.endswith(".xlsx"):
            df = read_xlsx_file(uploaded)
        elif name.endswith(".json"):
            df = read_json_file(uploaded)
        elif name.endswith(".txt"):
            df = read_txt_file(uploaded)
        elif name.endswith(".pdf"):
            df = read_pdf_file(uploaded)
            # αν μοιάζει σκαναρισμένο, προσπάθησε OCR αν υπάρχει API key
            if detect_scanned_pdf_need_ocr(df):
                uploaded.seek(0)
                byts = uploaded.read()
                df = read_with_ocr_if_possible(byts, uploaded.name)
        elif name.endswith(".docx"):
            df = read_docx_file(uploaded)
        elif name.endswith(".eml"):
            df = read_eml_file(uploaded)
        elif name.endswith(".msg"):
            df = read_msg_file(uploaded)
        else:
            st.error("Μη υποστηριζόμενος τύπος αρχείου.")
            df = pd.DataFrame()
    except Exception as e:
        st.error(f"❌ Σφάλμα ανάγνωσης αρχείου: {e}")
        df = pd.DataFrame()

    if not df.empty:
        df = normalize_columns(df)
        st.success(f"✅ Φορτώθηκαν {len(df)} γραμμές.")
        st.dataframe(df.head(50), use_container_width=True)

        missing = require_columns(df)
        if missing:
            st.warning(f"⚠️ Λείπουν απαιτούμενες στήλες: {', '.join(missing)}")
            st.info("Μπορείς να μετονομάσεις στήλες στο αρχείο σου σε: "
                    "'BIC', 'Currency', 'Amount', 'ChargeBearer' ή αντίστοιχα συνώνυμα.")
        else:
            st.subheader("📊 Αποτελέσματα Ελέγχου")
            if rules:
                result = apply_rules(df, rules)
                st.dataframe(result, use_container_width=True)
                # Στατιστικά
                st.subheader("📈 Στατιστικά")
                vc = result["MatchedRules"].value_counts().reset_index()
                vc.columns = ["Κανόνας/Κατάσταση", "Πλήθος"]
                st.table(vc)

                # Λήψη Excel
                out = io.BytesIO()
                result.to_excel(out, index=False)
                st.download_button(
                    "📥 Λήψη Αποτελεσμάτων (Excel)",
                    out.getvalue(),
                    "audit_results.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.error("Δεν υπάρχουν ενεργοί κανόνες για εφαρμογή.")
    else:
        st.warning("Δεν προέκυψαν δεδομένα από το αρχείο.")

else:
    st.info("➡️ Ανέβασε ένα αρχείο για έλεγχο.")

st.markdown("---")
st.caption("© 2025 Σύστημα Ελέγχου Τιμολογήσεων | CRBAGRAAXXX — Υποστήριξη αρχείων: CSV, XLSX, PDF(text), DOCX, EML, MSG, TXT, JSON. Προαιρετικό OCR με API.")
