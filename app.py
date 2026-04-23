import streamlit as st
import pandas as pd
from datetime import datetime
from decimal import Decimal

# ===================== CONFIG =====================

st.set_page_config(page_title="Excel → OFX", layout="centered")
st.markdown("""
    <style>
        /* remove menu superior */
        [data-testid="stToolbar"] {
            display: none;
        }

        /* remove decoração lateral */
        [data-testid="stDecoration"] {
            display: none;
        }

        /* remove status */
        [data-testid="stStatusWidget"] {
            display: none;
        }

        /* remove botão flutuante (esse vermelho aí) */
        .stActionButton {
            display: none;
        }

        /* remove footer */
        footer {
            visibility: hidden;
        }
    </style>
""", unsafe_allow_html=True)

# ===================== UTIL =====================

def ler_excel_inteligente(uploaded_file):
    for i in range(10):
        try:
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, header=i, dtype=str, engine="openpyxl")
            cols = [str(c).lower() for c in df.columns]

            if (
                any('data' in c for c in cols) and
                (any('valor' in c for c in cols) or any('r$' in c for c in cols))
            ):
                return df, i
        except Exception as e:
            st.error(str(e))
            raise

    raise ValueError("Não foi possível identificar automaticamente o cabeçalho.")


def extrair_info_bancaria(uploaded_file):
    uploaded_file.seek(0)
    df_topo = pd.read_excel(uploaded_file, header=None, nrows=1, engine="openpyxl")

    valores = df_topo.iloc[0].tolist()

    agencia = None
    conta = None

    for i, val in enumerate(valores):
        v = str(val).strip().lower()

        if v == "agencia" and i + 1 < len(valores):
            agencia = str(valores[i+1]).strip()

        if v == "conta" and i + 1 < len(valores):
            conta = str(valores[i+1]).strip()

    return agencia, conta


def parse_valor_br(valor):
    if pd.isna(valor) or str(valor).strip() in ['', '-']:
        return Decimal('0')

    s = str(valor).strip()

    if ',' in s:
        s = s.replace('.', '').replace(',', '.')

    if s.startswith('(') and s.endswith(')'):
        s = '-' + s[1:-1]

    try:
        return Decimal(s)
    except:
        return Decimal('0')


def clean_text(texto):
    return str(texto).encode('latin-1', 'ignore').decode('latin-1')


def detectar_colunas(df):
    rename = {}

    for col in df.columns:
        cl = col.lower().strip()

        if 'data' in cl:
            rename[col] = 'data'
        elif 'hist' in cl:
            rename[col] = 'historico'
        elif 'doc' in cl:
            rename[col] = 'documento'
        elif 'valor' in cl or 'r$' in cl:
            rename[col] = 'valor'
        elif 'saldo' in cl:
            rename[col] = 'saldo'

    return df.rename(columns=rename)


def validar_colunas(df):
    obrigatorias = ['data', 'valor']
    faltando = [c for c in obrigatorias if c not in df.columns]

    if faltando:
        raise ValueError(f"Colunas obrigatórias ausentes: {', '.join(faltando)}")


# ===================== CONVERSÃO =====================

def converter_para_ofx(df, agencia, conta, bank_id):
    
    # ===== LIMPEZA CRÍTICA (ADICIONAR AQUI) =====
    agencia = ''.join(filter(str.isdigit, str(agencia or ''))).zfill(4)
    conta = ''.join(filter(str.isdigit, str(conta or '')))
    bank_id = ''.join(filter(str.isdigit, str(bank_id or '')))
    
    conta = conta[:-1] if len(conta) > 1 else conta
    
    if len(bank_id) < 3:
        raise ValueError("Código do banco inválido (use código FEBRABAN, ex: 341, 237, 033)")
    
    if not conta:
        raise ValueError("Conta inválida (vazia ou formato errado)")

    if not bank_id:
        raise ValueError("Código do banco inválido")
    
    df = detectar_colunas(df)

    # 🔥 remove colunas duplicadas (ESSENCIAL)
    df = df.loc[:, ~df.columns.duplicated()]

    validar_colunas(df)

    # 🔥 garante que 'valor' não virou DataFrame
    if isinstance(df['valor'], pd.DataFrame):
        df['valor'] = df['valor'].iloc[:, 0]

    df['data'] = pd.to_datetime(df['data'], dayfirst=True, errors='coerce')
    df['valor'] = df['valor'].apply(parse_valor_br)

    if 'saldo' in df.columns:
        df['saldo'] = df['saldo'].apply(parse_valor_br)
    else:
        df['saldo'] = [Decimal('0')] * len(df)

    df = df.dropna(subset=['data'])

    # 🔥 evita bug do pandas
    df_mov = df[df['valor'].apply(lambda x: float(x) != 0)].copy()

    if df_mov.empty:
        raise ValueError("Nenhuma movimentação encontrada.")

    df_mov = df_mov.sort_values('data')

    start_dt = df_mov['data'].min().strftime("%Y%m%d")
    end_dt = df_mov['data'].max().strftime("%Y%m%d%H%M%S")

    saldo_series = df['saldo'].dropna()
    saldo_final = saldo_series.iloc[-1] if not saldo_series.empty else Decimal(a'0.00')

    # ===================== OFX =====================

    ofx = f"""OFXHEADER:100
DATA:OFXSGML
VERSION:103
SECURITY:NONE
ENCODING:USASCII
CHARSET:1252
COMPRESSION:NONE
OLDFILEUID:NONE
NEWFILEUID:NONE

<OFX>
<SIGNONMSGSRSV1>
<SONRS>
<STATUS><CODE>0</CODE><SEVERITY>INFO</SEVERITY></STATUS>
<DTSERVER>{datetime.now().strftime("%Y%m%d%H%M%S")}</DTSERVER>
<LANGUAGE>POR</LANGUAGE>
<FI><ORG>{bank_id}</ORG><FID>{bank_id}</FID></FI>
<INTU.BID>{bank_id}</INTU.BID>
</SONRS>
</SIGNONMSGSRSV1>
<BANKMSGSRSV1>
<STMTTRNRS>
<TRNUID>1</TRNUID>
<STATUS><CODE>0</CODE><SEVERITY>INFO</SEVERITY></STATUS>
<STMTRS>
<CURDEF>BRL</CURDEF>
<BANKACCTFROM>
<BANKID>{bank_id}</BANKID>
<BRANCHID>{agencia}</BRANCHID>
<ACCTID>{conta}</ACCTID>
<ACCTTYPE>CHECKING</ACCTTYPE>
</BANKACCTFROM>
<BANKTRANLIST>
<DTSTART>{start_dt}</DTSTART>
<DTEND>{end_dt}</DTEND>
"""

    for idx, row in df_mov.iterrows():
        valor = float(row['valor'])
        tipo = "CREDIT" if valor >= 0 else "DEBIT"
        data = row['data'].strftime("%Y%m%d%H%M%S")

        fitid = f"{data}{idx:06d}"

        nome = clean_text(row.get('historico', ''))[:32]
        memo = clean_text(f"{row.get('historico', '')} | Doc: {row.get('documento', '')}")[:255]

        ofx += f"""<STMTTRN>
<TRNTYPE>{tipo}</TRNTYPE>
<DTPOSTED>{data}</DTPOSTED>
<TRNAMT>{valor:.2f}</TRNAMT>
<FITID>{fitid}</FITID>
<NAME>{nome}</NAME>
<MEMO>{memo}</MEMO>
</STMTTRN>
"""

    ofx += f"""</BANKTRANLIST>
<LEDGERBAL>
<BALAMT>{saldo_final:.2f}</BALAMT>
<DTASOF>{end_dt}</DTASOF>
</LEDGERBAL>
</STMTRS>
</STMTTRNRS>
</BANKMSGSRSV1>
</OFX>"""

    return ofx.encode("utf-8")


# ===================== UI =====================

st.title("🧾 Conversor Excel para OFX")

uploaded_file = st.file_uploader("Envie seu extrato (.xls/.xlsx)", type=["xls", "xlsx"])

col1, col2 = st.columns(2)

agencia_auto, conta_auto = (None, None)

if uploaded_file:
    try:
        agencia_auto, conta_auto = extrair_info_bancaria(uploaded_file)
    except Exception as e:
        st.warning("Não foi possível extrair agência/conta automaticamente")

agencia = col1.text_input("Agência", value=agencia_auto or "")
conta = col2.text_input("Conta", value=conta_auto or "")

bank_id = st.text_input("Código do Banco", "")

if uploaded_file:

    try:
        df, header = ler_excel_inteligente(uploaded_file)

        st.success(f"Cabeçalho detectado automaticamente (linha {header})")

        st.subheader("Preview do arquivo")
        st.dataframe(df.head())

        if st.button("🚀 Converter para OFX", type="primary"):

            with st.spinner("Processando..."):
                ofx = converter_para_ofx(df, agencia, conta, bank_id)

                nome = uploaded_file.name.rsplit(".", 1)[0] + ".ofx"

                st.success("Arquivo gerado com sucesso")

                st.download_button(
                    "📥 Baixar OFX",
                    data=ofx,
                    file_name=nome,
                    mime="text/plain"
                )

    except Exception as e:
        st.error(f"Erro: {str(e)}")
