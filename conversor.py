import io
import uuid
from datetime import datetime
from dateutil import parser as dateparser

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Excel/CSV → OFX Converter", page_icon="💳", layout="centered")

st.title("💳 Excel/CSV → OFX Converter")
st.caption("Converta extratos de planilha em arquivo OFX (1.03/SGML).")

st.markdown("### 1) Envie seu arquivo")
file = st.file_uploader("Excel (.xlsx, .xls) ou CSV (.csv)", type=["xlsx", "xls", "csv"])

# Configurações básicas do OFX
st.markdown("### 2) Configurações do OFX")
col_a, col_b = st.columns(2)
with col_a:
    bank_id = st.text_input("Bank ID (agência/ident.)", value="00000000")
    acct_id = st.text_input("Account ID (conta)", value="0000000000")
    acct_type = st.selectbox("Account Type", ["CHECKING", "SAVINGS", "CREDITLINE"], index=0)
with col_b:
    currency = st.text_input("Currency", value="BRL")
    fitid_col_generate = st.checkbox("Gerar FITID automaticamente (UUID)", value=True)

st.divider()

def load_dataframe(uploaded):
    if uploaded is None:
        return None
    if uploaded.name.lower().endswith(".csv"):
        return pd.read_csv(uploaded)
    else:
        return pd.read_excel(uploaded)

df = load_dataframe(file)

if df is not None:
    st.markdown("### 3) Mapeie as colunas")
    st.write("Prévia da planilha:")
    st.dataframe(df.head(10), use_container_width=True)

    cols = ["<nenhuma>"] + list(df.columns)

    # Campos obrigatórios
    date_col = st.selectbox("Coluna de DATA", cols)
    amount_col = st.selectbox("Coluna de VALOR (positivo/crédito, negativo/débito)", cols)
    memo_col = st.selectbox("Coluna de DESCRIÇÃO/MEMO", cols)

    # Campos opcionais
    fitid_col = st.selectbox("Coluna de ID (FITID) [opcional]", cols, index=0)
    trntype_col = st.selectbox("Coluna de TIPO (CREDIT/DEBIT) [opcional]", cols, index=0)

    date_parse_hint = st.text_input("Formato da data (opcional, ex: %d/%m/%Y). Se vazio, detecção automática.", value="")

    def norm_date(x):
        if pd.isna(x):
            return None
        try:
            if date_parse_hint.strip():
                return datetime.strptime(str(x), date_parse_hint)
            return dateparser.parse(str(x), dayfirst=True)
        except Exception:
            return None

    def norm_amount(x):
        if pd.isna(x):
            return None
        s = str(x).replace("R$", "").replace(".", "").replace(",", ".").strip()
        try:
            return float(s)
        except Exception:
            try:
                return float(x)
            except Exception:
                return None

    def infer_trntype(v):
        if v is None:
            return "DEBIT"
        return "CREDIT" if v >= 0 else "DEBIT"

    def fmt_dtposted(dt: datetime):
        # OFX 1.03 costuma aceitar YYYYMMDD ou YYYYMMDDHHMMSS
        return dt.strftime("%Y%m%d")

    def ofx_header():
        # Cabeçalho OFX 1.03/SGML (amplamente aceito)
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

    def ofx_open():
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
            "        <CURDEF>{cur}\n"
            "        <BANKACCTFROM>\n"
            "          <BANKID>{bank}\n"
            "          <ACCTID>{acct}\n"
            "          <ACCTTYPE>{accttype}\n"
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

    # Validação mínima
    ready = True
    missing = []
    if date_col == "<nenhuma>":
        ready = False; missing.append("DATA")
    if amount_col == "<nenhuma>":
        ready = False; missing.append("VALOR")
    if memo_col == "<nenhuma>":
        ready = False; missing.append("DESCRIÇÃO/MEMO")

    if not ready:
        st.warning("Selecione as colunas obrigatórias: " + ", ".join(missing))
    else:
        # Monta transações
        trns = []
        for _, row in df.iterrows():
            dt = norm_date(row[date_col]) if date_col in df.columns else None
            amt = norm_amount(row[amount_col]) if amount_col in df.columns else None
            memo = str(row[memo_col]) if memo_col in df.columns else ""
            if dt is None or amt is None:
                continue

            if trntype_col != "<nenhuma>":
                ttype_raw = str(row[trntype_col]).strip().upper()
                ttype = "CREDIT" if "CREDIT" in ttype_raw or "CR" == ttype_raw else ("DEBIT" if "DEBIT" in ttype_raw or "DR" == ttype_raw else infer_trntype(amt))
            else:
                ttype = infer_trntype(amt)

            if fitid_col != "<nenhuma>":
                fitid = str(row[fitid_col])
            else:
                fitid = uuid.uuid4().hex if fitid_col_generate else f"{int(dt.timestamp())}-{abs(hash((amt, memo)))%10_000_000}"

            trns.append({
                "DTPOSTED": fmt_dtposted(dt),
                "TRNAMT": f"{amt:.2f}",
                "TRNTYPE": ttype,
                "FITID": fitid,
                "MEMO": memo,
            })

        if len(trns) == 0:
            st.error("Nenhuma transação válida encontrada com o mapeamento atual.")
        else:
            # Datas de início/fim
            dts = [t["DTPOSTED"] for t in trns]
            dtstart, dtend = min(dts), max(dts)

            # Construir OFX
            body = []
            body.append(ofx_open().format(cur=currency, bank=bank_id, acct=acct_id, accttype=acct_type))
            body.append(f"          <DTSTART>{dtstart}\n")
            body.append(f"          <DTEND>{dtend}\n")
            for t in trns:
                body.append("          <STMTTRN>\n")
                body.append(f"            <TRNTYPE>{t['TRNTYPE']}\n")
                body.append(f"            <DTPOSTED>{t['DTPOSTED']}\n")
                body.append(f"            <TRNAMT>{t['TRNAMT']}\n")
                body.append(f"            <FITID>{t['FITID']}\n")
                body.append(f"            <MEMO>{t['MEMO']}\n")
                body.append("          </STMTTRN>\n")
            body.append(ofx_close())
            ofx_text = ofx_header() + "".join(body)

            st.success(f"Gerado {len(trns)} lançamento(s).")
            fname = f"export_{acct_id}_{dtstart}_{dtend}.ofx"
            st.download_button("⬇️ Baixar OFX", data=ofx_text.encode("utf-8"), file_name=fname, mime="application/x-ofx")

            with st.expander("📄 Visualizar OFX"):
                st.code(ofx_text, language="xml")
else:
    st.info("Envie um arquivo Excel (.xlsx/.xls) ou CSV para começar.")
