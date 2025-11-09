import io
import os
import yaml
import pandas as pd
import streamlit as st

# ===============================
#  💶 Εφαρμογή Ελέγχου Τιμολογήσεων
# ===============================
st.set_page_config(page_title="Σύστημα Ελέγχου Τιμολογήσεων", page_icon="💶", layout="wide")
st.title("💶 Σύστημα Ελέγχου Τιμολογήσεων Εισερχόμενων Εντολών")

# ===============================
#  🧠 Βοηθητικές Συναρτήσεις
# ===============================
def load_rules(file_path="rules_example.yaml"):
    """Φόρτωση κανόνων από YAML αρχείο"""
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)
        if not data or "rules" not in data:
            raise ValueError("Το αρχείο δεν περιέχει ενότητα 'rules'.")
        return data["rules"]
    except Exception as e:
        st.error(f"⚠️ Δεν ήταν δυνατή η φόρτωση κανόνων: {e}")
        return []


def validate_rules(rules):
    """Έλεγχος εγκυρότητας κανόνων YAML"""
    required_keys = {"id", "description", "condition", "action"}
    for i, rule in enumerate(rules, start=1):
        missing = required_keys - rule.keys()
        if missing:
            return f"Ο κανόνας #{i} λείπει τα πεδία: {', '.join(missing)}"
        if "field" not in rule["condition"]:
            return f"Ο κανόνας #{i} λείπει πεδίο 'field' στο condition"
    return None


def validate_file_structure(df, required_columns):
    """Έλεγχος αν υπάρχουν τα απαραίτητα πεδία στο αρχείο εντολών"""
    missing = [c for c in required_columns if c not in df.columns]
    if missing:
        return f"Λείπουν τα πεδία: {', '.join(missing)}"
    return None


def apply_rules_to_data(df, rules):
    """Εφαρμόζει τους κανόνες στις εντολές"""
    results = []

    for i, row in df.iterrows():
        matched = []
        for rule in rules:
            cond = rule["condition"]
            field = cond.get("field")
            value = row.get(field)

            try:
                if "equals" in cond and str(value) == str(cond["equals"]):
                    matched.append(rule["id"])
                elif "in" in cond and str(value) in [str(x) for x in cond["in"]]:
                    matched.append(rule["id"])
                elif "greater_than" in cond and float(value) > float(cond["greater_than"]):
                    matched.append(rule["id"])
            except Exception:
                continue

        results.append({
            "Index": i + 1,
            "BIC": row.get("BIC", ""),
            "Currency": row.get("Currency", ""),
            "Amount": row.get("Amount", ""),
            "ChargeBearer": row.get("ChargeBearer", ""),
            "MatchedRules": ", ".join(matched) if matched else "—",
        })

    return pd.DataFrame(results)


# ===============================
#  📘 Sidebar - Φόρτωση Κανόνων
# ===============================
st.sidebar.header("⚙️ Ρυθμίσεις Κανόνων")

uploaded_rules = st.sidebar.file_uploader("Ανέβασε αρχείο κανόνων (YAML):", type=["yaml", "yml"])

if uploaded_rules:
    try:
        rules = yaml.safe_load(uploaded_rules).get("rules", [])
        msg = validate_rules(rules)
        if msg:
            st.sidebar.error(f"❌ Σφάλμα δομής αρχείου κανόνων:\n{msg}")
            rules = []
        else:
            st.sidebar.success("✅ Φορτώθηκαν επιτυχώς οι κανόνες από το αρχείο!")
    except Exception as e:
        st.sidebar.error(f"❌ Μη έγκυρο αρχείο YAML: {e}")
        rules = []
else:
    rules = load_rules()
    if rules:
        st.sidebar.info("Χρησιμοποιούνται οι προεπιλεγμένοι κανόνες (`rules_example.yaml`).")

# ===============================
#  📋 Προβολή Κανόνων
# ===============================
st.subheader("🔍 Ενεργοί Κανόνες Τιμολόγησης")

if rules:
    for r in rules:
        st.markdown(f"- **{r['id']}**: {r['description']}")
else:
    st.warning("Δεν βρέθηκαν έγκυροι κανόνες. Ανέβασε σωστό YAML αρχείο.")

# ===============================
#  📥 Ανέβασμα Εντολών
# ===============================
st.subheader("📂 Ανέβασε λίστα τιμολογημένων εντολών (CSV ή Excel)")

uploaded_file = st.file_uploader("Επίλεξε αρχείο:", type=["csv", "xlsx"])
required_columns = ["BIC", "Currency", "Amount", "ChargeBearer"]

if uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        # Έλεγχος στηλών
        msg = validate_file_structure(df, required_columns)
        if msg:
            st.error(f"❌ Μη έγκυρο αρχείο: {msg}")
        else:
            st.success(f"✅ Φορτώθηκαν {len(df)} εγγραφές.")
            st.dataframe(df.head())

            # Εφαρμογή κανόνων
            st.subheader("📊 Αποτελέσματα Ελέγχου")
            result_df = apply_rules_to_data(df, rules)

            # Εμφάνιση αναφοράς
            st.dataframe(result_df)

            # Αναλυτικά στατιστικά
            st.subheader("📈 Αναφορά Ελέγχου")
            summary = result_df["MatchedRules"].value_counts().reset_index()
            summary.columns = ["Κανόνας / Κατάσταση", "Πλήθος"]
            st.table(summary)

            # Λήψη αρχείου Excel
            output = io.BytesIO()
            result_df.to_excel(output, index=False)
            st.download_button(
                label="📥 Λήψη Αποτελεσμάτων (Excel)",
                data=output.getvalue(),
                file_name="audit_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"⚠️ Σφάλμα κατά την επεξεργασία του αρχείου: {e}")
else:
    st.info("➡️ Ανέβασε αρχείο εντολών για έλεγχο (CSV ή XLSX).")

# ===============================
#  🏁 Τέλος
# ===============================
st.markdown("---")
st.caption("© 2025 Σύστημα Ελέγχου Τιμολογήσεων | CRBAGRAAXXX")
