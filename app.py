import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="RE-Excel Sheet Converter", layout="wide")

EXPECTED_GOLDEN = 159

# ---------- UI ----------
st.title("RE-Excel Sheet Converter")
st.subheader("Convert Type 1 Excel into structured statement-score data")

# ---------- HELPERS ----------
def is_numeric(x):
    return re.match(r"^-?\d+(\.\d+)?$", str(x).strip()) is not None

def norm_statement(s):
    return (
        str(s or "")
        .lower()
        .replace("“", "'")
        .replace("”", "'")
        .replace('"', "'")
        .replace("’", "'")
        .strip()
    )

def clean_cell(x):
    x = str(x or "")
    x = re.sub(r"202\d\s*\|\s*Page\s*\d+\s*of:\s*\d+", "", x, flags=re.I)
    x = re.sub(r"Page\s*\d+\s*of\s*:\s*\d+", "", x, flags=re.I)
    return x.strip()

# ---------- EXACT PARSER (PORTED) ----------
def parse_type1(file):

    if file.name.lower().endswith(".csv"):
        raw = pd.read_csv(file, header=None, dtype=str).fillna("")
        rows = raw.values.tolist()
        source_type = "csv"
    else:
        raw = pd.read_excel(file, header=None, dtype=str).fillna("")
        rows = raw.values.tolist()
        source_type = "excel"

    total_rows = len(rows)

    # 🔴 EXACT skip (DO NOT CHANGE)
    body = rows[45:]

    # 🔴 EXACT row repair
    target_excel_rows = [137, 138, 144]
    repairs = []

    for excel_row in target_excel_rows:
        idx = excel_row - 46
        if 0 <= idx < len(body):
            body[idx] = [clean_cell(v) for v in body[idx]]
            repairs.append(excel_row)

    out = []
    warnings = []

    def push_pair(s, sc):
        try:
            sc = float(sc)
            if (
                isinstance(s, str)
                and s.strip()
                and isinstance(sc, float)
                and sc >= 0
                and sc <= 10
            ):
                out.append({"statement": s.strip(), "score": sc})
                return True
        except:
            return False
        return False

    # 🔴 EXACT parsing loop
    for i, r in enumerate(body):

        A = str(r[0]) if len(r) > 0 else ""
        B = str(r[1]) if len(r) > 1 else ""
        F = str(r[5]) if len(r) > 5 else ""

        # Case 1
        if A and not is_numeric(A) and is_numeric(F):
            if push_pair(A, F):
                continue

        # Case 2
        if A and re.match(r"^\s*\d+(\.\d+)?\s+", A):
            m = re.match(r"^\s*(\d+(?:\.\d+)?)\s+(.+)$", A)
            if m:
                if push_pair(m.group(2), m.group(1)):
                    continue

        # Case 3
        if is_numeric(A) and B:
            if push_pair(B, A):
                continue

        # Case 4
        if A and is_numeric(B):
            if push_pair(A, B):
                continue

        # 🔴 EXACT warning logic
        joined = " ".join([str(v) for v in r]).strip()
        if joined and len(warnings) < 8:
            warnings.append(f"Unparsed @ line {45 + i + 1}: {joined[:80]}")

    df = pd.DataFrame(out)

    return df, source_type, total_rows, repairs, warnings


# ---------- UI FLOW ----------
file = st.file_uploader("Upload Type 1 Excel or CSV", type=["xlsx", "xls", "csv"])

if file:

    df, source, total_rows, repairs, warnings = parse_type1(file)

    st.success("File processed")

    # Metrics
    col1, col2, col3 = st.columns(3)
    col1.metric("Source", source.upper())
    col2.metric("Raw rows", total_rows)
    col3.metric("Usable pairs", len(df))

    # Golden check
    if len(df) == EXPECTED_GOLDEN:
        st.success(f"✅ GOLDEN: {EXPECTED_GOLDEN} rows")
    else:
        st.warning(f"{len(df)} rows (Expected {EXPECTED_GOLDEN})")

    # Repairs info
    if repairs:
        st.info(f"Auto-repaired rows: {', '.join(map(str, repairs))}")

    # Warnings
    if warnings:
        with st.expander("Warnings"):
            for w in warnings:
                st.write(w)

    # Output
    st.subheader("Structured Output")
    st.dataframe(df, use_container_width=True)

    # Download
    csv = df.to_csv(index=False).encode("utf-8")

    st.download_button(
        "Download structured CSV",
        data=csv,
        file_name="type1_structured.csv",
        mime="text/csv",
    )

else:
    st.info("Upload a Type 1 Excel file to begin.")
