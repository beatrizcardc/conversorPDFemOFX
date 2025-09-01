import io
import uuid
from datetime import datetime
from dateutil import parser as dateparser

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Excel/CSV → OFX Converter",
    page_icon="💳",
    layout="centered",
)
st.title("💳 Excel/CSV → OFX Converter")
st.caption("Converta planilhas (CSV/XLSX/XLSB) para OFX 1.03/SGML.")

def load_dataframe(uploaded):
    if uploaded is None:
        return None

    name = uploaded.name.lower()
    raw = uploaded.read()
    bio = io.BytesIO(raw)

    if name.endswith(".csv"):
        try:
            bio.seek(0)
            return pd.read_csv(bio)
        except UnicodeDecodeError:
            bio.seek(0)
            return pd.read_csv(bio, encoding="latin-1", engine="python", sep=None)
    elif name.endswith(".xlsx"):
        bio.seek(0)
        return pd.read_excel(bio, engine="openpyxl")
    elif name.endswith(".xlsb"):
        bio.seek(0)
        return pd.read_excel(bio, engine="pyxlsb")
    else:
        raise ValueError("Formato não suportado. Envie CSV, XLSX ou XLSB.")

def norm_date(x, fmt_hint: str):
    if pd.isna(x):
        return None
    try:
        if fmt_hint.strip():
            return datetime.strptime(str(x), fmt_hint)
        return dateparser.parse(str(x), dayfirst=True)
    except Exception:
        return None

def norm_amount(x):
    if pd.isna(x):
        return None
    s = str(x).strip().replace("R$", "").replace(" ", "")
    if s.count(",") == 1 and s.count(".") == 0:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s and "." in s and s.rfind(",") > s.rfind("."):
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except:
        return None

def fmt_dtposted(dt: datetime):
    return dt.strftime("%Y%m%d")

def ofx_header():
    return (
        "OFXHEADER:100\n"
        "DATA:OFXSGML\n"
        "VERSION:103\n"
        "SECURITY:NONE\n"
        "ENCODING:USASCII\n"
        "CHARSET:1252\n"
        "COMPRESSION:NONE\n"
        "OLDFILEUID:NONE\n"
        "NEWFILEUID:NONE\n\n"
    )

def ofx_open(cur, bank, acct, accttype):
    now = datetime.now().strftime("%Y%m%d%H%M%S")
    return (
        "<OFX>\n"
        "  <SIGNONMSGSRSV1>\n"
        "    <SONRS>\n"
        "      <STATUS><CODE>0</CODE><SEVERITY>INFO</SEVERITY></STATUS>\n"
        f"      <DTSERVER>{now}</DTSERVER>\n"
        "      <LANGUAGE>POR</LANGUAGE>\n"
        "    </SONRS>\n"
        "  </SIGNONMSGSRSV1>\n"
        "  <BANKMSGSRSV1>\n"
        "    <STMTTRNRS>\n"
        "      <TRNUID>1</TRNUID>\n"
        "      <STATUS><CODE>0</CODE><SEVERITY>INFO</SEVERITY></STATUS>\n"
        "      <STMTRS>\n"
        f"        <CURDEF>{cur}\n"
        "        <BANKACCTFROM>\n"
        f"          <BANKID>{bank}\n"
        f"          <ACCTID>{acct}\n"
        f"          <ACCTTYPE>{accttype}\n"
        "        </BANKACCTFROM>\n"
        "        <BANKTRANLIST>\n"
    )

def ofx_close():
    return (
        "        </BANKTRANLIST>\n"
        "      </STMTRS>\n"
        "    </STMTTRNRS>\n"
        "  </BANKMSGSRSV1>\n"
        "</OFX>\n"
    )

# Interface
file = st.file_uploader("Excel (.xlsx, .xlsb) ou CSV (.csv)", type=["xlsx", "xlsb", "csv"])
df = load_dataframe(file) if file else None
if df is not None:
    st.dataframe(df)
