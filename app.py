import io
import os
import yaml
import pandas as pd
import streamlit as st

# ===============================
#  🔧 Ρυθμίσεις Εφαρμογής
# ===============================
st.set_page_config(page_title="Σύστημα Ελέγχου Τιμολογήσεων", page_icon="💶", layout="wide")
st.title("💶 Σύστημα Ελέγχου Τιμολογήσεων Εισερχόμενων Εντολών")

# ===============================
#  🧠 Φόρτωση Κανόνων YAML
# ===============================
def load_rules(file_path="rules_example.yaml"):
    """Φόρτωση κανόνων από YAML αρχείο"""
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)
        return data.get("rules", [])
    except Exception as e:
        st.warning(f"⚠️ Δεν βρέθηκε το αρχείο κανόνων ({file_path}) ή δεν είναι έγκυρο.\n{e}")
        return []

# ===============================
#  🧾 Εφαρμογή Κανόνων σε Εντολές
# ===============================
def apply_rules_to_data(df, rules):
    """Εφαρμόζει τους κανόνες στις εντολές και επιστρέφει αποτελέσματα"""
    results = []

    for i, row in df.iterrows():
        matched = []
        for rule in rules:
            field = rule["condition"].get("field")
            cond = rule["condition"]
            value = row.get(field)

            # Έλεγχοι τύπων συνθηκών
            if "equals" in cond and str(value) == str(cond["equals"]):
                matched.append(rule["id"])
            elif "in" in cond and str(value) in cond["in"]:
                matched.append(rule["id"])
            elif "greater_than" in cond:
                try:
                    if float(value) > float(cond["greater_than"]):
                        matched.append(rule["id"])
                except Exception:
                    pass

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
#  📤 Ανέβασμα Κανόνων
# ===============================
st.sidebar.header("📘 Ρυθμίσεις")
uploaded_rules = st.sidebar.file_uploader("Ανέβασε αρχείο κανόνων (YAML):", type=["yaml", "yml"])

if uploaded_rules:
    rules = yaml.safe_load(uploaded_rules).get("rules", [])
    st.sidebar.success("✅ Φορτώθηκαν επιτυχώς οι κανόνες από το αρχείο!")
else:
    rules = load_rules()
    st.sidebar.info("Χρησιμοποιούνται οι προεπιλεγμένοι κανόνες (`rules_example.yaml`).")

# Εμφάνιση κανόνων
st.subheader("🔍 Ενεργοί Κανόνες Τιμολόγησης")
if rules:
    for r in rules:
        st.markdown(f"- **{r['id']}**: {r['description']}")
else:
    st.warning("Δεν βρέθηκαν κανόνες. Ανέβασε αρχείο YAML με κανόνες.")

# ===============================
#  📥 Ανέβασμα Λίστας Εντολών
# ===============================
st.subheader("📂 Ανέβασε λίστα τιμολογημένων εντολών")
uploaded_file = st.file_uploader("Επίλεξε αρχείο CSV ή Excel", type=["csv", "xlsx"])

if uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        st.success(f"✅ Φορτώθηκαν {len(df)} εγγραφές επιτυχώς.")
        st.dataframe(df.head())

        # Εφαρμογή κανόνων
        st.subheader("📊 Αποτελέσματα Ελέγχου")
        result_df = apply_rules_to_data(df, rules)
        st.dataframe(result_df)

        # Export αποτελεσμάτων
        output = io.BytesIO()
        result_df.to_excel(output, index=False)
        st.download_button(
            label="📥 Λήψη Αποτελεσμάτων (Excel)",
            data=output.getvalue(),
            file_name="audit_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Σφάλμα κατά την επεξεργασία αρχείου: {e}")
else:
    st.info("➡️ Ανέβασε αρχείο με εντολές (π.χ. export από το σύστημα).")

# ===============================
#  🏁 Τέλος Εφαρμογής
# ===============================
st.markdown("---")
st.caption("© 2025 Σύστημα Ελέγχου Τιμολογήσεων | CRBAGRAAXXX")
