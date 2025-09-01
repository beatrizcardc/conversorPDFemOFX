import io
import uuid
from datetime import datetime
from dateutil import parser as dateparser

import chardet
import pandas as pd
import streamlit as st

# ---------------------- Setup da página ----------------------
st.set_page_config(
    page_title="Excel/CSV → OFX Converter",
    page_icon="💳",
    layout="centered",
)
st.title("💳 Excel/CSV → OFX Converter")
st.caption("Converta planilhas (CSV/XLSX/XLS/XLSB) para OFX 1.03/SGML.")

# ---------------------- Helpers de leitura ----------------------
def sniff_encoding(file_bytes: bytes, default="utf-8"):
    try:
        res = chardet.detect(file_bytes)
        enc = res.get("encoding") or default
        return enc
    except Exception:
        return default

def load_dataframe(uploaded):
    """Carrega CSV/XLSX/XLS/XLSB de maneira robusta."""
    if uploaded is None:
        return None

    name = uploaded.name.lower()

    # Precisamos do conteúdo em bytes para reabrir várias vezes
    raw = uploaded.read()
    bio = io.BytesIO(raw)

    if name.endswith(".csv"):
        # 1) tenta utf-8 padrão
        try:
            bio.seek(0)
            return pd.read_csv(bio)
        except UnicodeDecodeError:
            # 2) tenta detectar encoding
            enc = sniff_encoding(raw, default="latin-1")
            # tenta separadores comuns
            for sep in [",", ";", "\t", "|"]:
                try:
                    bio.seek(0)
                    return pd.read_csv(bio, encoding=enc, sep=sep)
                except Exception:
                    continue
            # último recurso
            bio.seek(0)
            return pd.read_csv(bio, encoding=enc, engine="python", sep=None)
    elif name.endswith(".xlsx"):
        bio.seek(0)
        return pd.read_excel(bio, engine="openpyxl")
    elif name.endswith(".xls"):
        bio.seek(0)
        return pd.read_excel(bio, engine="xlrd")
    elif name.endswith(".xlsb"):
        bio.seek(0)
        return pd.read_excel(bio, engine="pyxlsb")
    else:
        raise ValueError("Formato não suportado. Envie CSV, XLSX, XLS ou XLSB.")

# ---------------------- Normalizações ----------------------
def norm_date(x, fmt_hint: str):
    if pd.isna(x):
        return None
    try:
        if fmt_hint.strip():
            return datetime.strptime(str(x), fmt_hint)
        # dayfirst=True ajuda em dd/mm/yyyy
        return dateparser.parse(str(x), dayfirst=True)
    except Exception:
        return None

def norm_amount(x):
    """Normaliza valores como: 'R$ 1.234,56' → 1234.56  | '1,234.56' → 1234.56"""
    if pd.isna(x):
        return None
    s = str(x).strip()
    # remove R$ e espaços
    s = s.replace("R$", "").replace(" ", "")
    # heurística: se tiver vírgula e NÃO tiver ponto como decimal, trocar vírgula por ponto
    # também remove separadores de milhar
    if s.count(",") == 1 and s.count(".") == 0:
        s = s.replace(".", "")      # remove eventuais milhares antigos
        s = s.replace(",", ".")     # vírgula decimal → ponto
    else:
        # tentar remover milhares (pontos) quando vier em estilo brasileiro
        # '1.234,56' -> '1234,56' -> depois vira 1234.56 no try abaixo
        if "," in s and "." in s and s.rfind(",") > s.rfind("."):
            s = s.replace(".", "")
            s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        try:
            return float(str(x))
        except Exception:
            return None

def infer_trntype(v):
    if v is None:
        return "DEBIT"
    return "CREDIT" if v >= 0 else "DEBIT"

def fmt_dtposted(dt: datetime):
    # OFX 1.03 costuma aceitar YYYYMMDD ou YYYYMMDDHHMMSS (vamos num simples)
    return dt.strftime("%Y%m%d")

# ---------------------- OFX Builders ----------------------
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

# ---------------------- UI ----------------------
st.markdown("### 1) Envie seu arquivo")
file = st.file_uploader("Excel (.xlsx, .xls, .xlsb) ou CSV (.csv)", type=["xlsx", "xls", "xlsb", "csv"])

st.markdown("### 2) Configurações do OFX")
col_a, col_b = st.columns(2)
with col_a:
    bank_id = st.text_input("Bank ID (agência/ident.)", value="00000000")
    acct_id = st.text_input("Account ID (conta)", value="0000000000")
    acct_type = st.selectbox("Account Type", ["CHECKING", "SAVINGS", "CREDITLINE"], index=0)
with col_b:
    currency = st.text_input("Currency", value="BRL")
    fitid_auto = st.checkbox("Gerar FITID automaticamente (UUID)", value=True)

st.markdown("### 3) Mapeamento de colunas")
date_hint = st.text_input("Formato da data (opcional, ex: %d/%m/%Y). Se vazio, detecção automática.", value="")

df = None
error_box = st.empty()
try:
    if file is not None:
        # volta o ponteiro para o início, pois o st.file_uploader moveu
        file.seek(0)
        df = load_dataframe(file)
except Exception as e:
    error_box.error(f"Erro ao ler arquivo: {e}")

if df is not None:
    st.write("Prévia da planilha:")
    st.dataframe(df.head(10), use_container_width=True)

    cols = ["<nenhuma>"] + list(df.columns)
    date_col   = st.selectbox("Coluna de DATA", cols)
    amount_col = st.selectbox("Coluna de VALOR (positivo/crédito, negativo/débito)", cols)
    memo_col   = st.selectbox("Coluna de DESCRIÇÃO/MEMO", cols)
    fitid_col  = st.selectbox("Coluna de ID (FITID) [opcional]", cols, index=0)
    ttype_col  = st.selectbox("Coluna de TIPO (CREDIT/DEBIT) [opcional]", cols, index=0)

    missing = []
    ready = True
    if date_col == "<nenhuma>":
        ready = False; missing.append("DATA")
    if amount_col == "<nenhuma>":
        ready = False; missing.append("VALOR")
    if memo_col == "<nenhuma>":
        ready = False; missing.append("DESCRIÇÃO/MEMO")

    if not ready:
        st.warning("Selecione as colunas obrigatórias: " + ", ".join(missing))
    else:
        # ---------------------- Monta transações ----------------------
        transactions = []
        skipped = []  # para logar linhas ignoradas

        for idx, row in df.iterrows():
            dt = norm_date(row[date_col], date_hint) if date_col in df.columns else None
            amt = norm_amount(row[amount_col]) if amount_col in df.columns else None
            memo = str(row[memo_col]) if memo_col in df.columns else ""

            if dt is None or amt is None:
                reason = "data inválida" if dt is None else "valor inválido"
                skipped.append({"linha": int(idx), "motivo": reason, "memo": memo})
                continue

            if ttype_col != "<nenhuma>":
                raw_t = str(row[ttype_col]).strip().upper()
                if "CREDIT" in raw_t or raw_t == "CR":
                    ttype = "CREDIT"
                elif "DEBIT" in raw_t or raw_t == "DR":
                    ttype = "DEBIT"
                else:
                    ttype = "CREDIT" if amt >= 0 else "DEBIT"
            else:
                ttype = "CREDIT" if amt >= 0 else "DEBIT"

            if fitid_col != "<nenhuma>":
                fitid = str(row[fitid_col])
            else:
                fitid = uuid.uuid4().hex if fitid_auto else f"{int(dt.timestamp())}-{abs(hash((amt, memo)))%10_000_000}"

            transactions.append(
                {
                    "DTPOSTED": fmt_dtposted(dt),
                    "TRNAMT": f"{amt:.2f}",
                    "TRNTYPE": ttype,
                    "FITID": fitid,
                    "MEMO": memo,
                }
            )

        if len(transactions) == 0:
            st.error("Nenhuma transação válida encontrada com o mapeamento atual.")
        else:
            dts = [t["DTPOSTED"] for t in transactions]
            dtstart, dtend = min(dts), max(dts)

            body = []
            body.append(ofx_open(currency, bank_id, acct_id, acct_type))
            body.append(f"          <DTSTART>{dtstart}\n")
            body.append(f"          <DTEND>{dtend}\n")
            for t in transactions:
                body.append("          <STMTTRN>\n")
                body.append(f"            <TRNTYPE>{t['TRNTYPE']}\n")
                body.append(f"            <DTPOSTED>{t['DTPOSTED']}\n")
                body.append(f"            <TRNAMT>{t['TRNAMT']}\n")
                body.append(f"            <FITID>{t['FITID']}\n")
                body.append(f"            <MEMO>{t['MEMO']}\n")
                body.append("          </STMTTRN>\n")
            body.append(ofx_close())

            ofx_text = ofx_header() + "".join(body)

            st.success(f"Gerado {len(transactions)} lançamento(s).")
            fname = f"export_{acct_id}_{dtstart}_{dtend}.ofx"
            st.download_button("⬇️ Baixar OFX", data=ofx_text.encode("utf-8"), file_name=fname, mime="application/x-ofx")

            with st.expander("📄 Visualizar OFX"):
                st.code(ofx_text, language="xml")

            if skipped:
                st.warning(f"{len(skipped)} linha(s) foram ignoradas.")
                st.dataframe(pd.DataFrame(skipped), use_container_width=True)
else:
    st.info("Envie um arquivo Excel (.xlsx/.xls/.xlsb) ou CSV para começar.")

