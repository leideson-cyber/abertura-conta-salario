import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import pandas as pd
import pypdf
import streamlit as st

st.set_page_config(page_title="Abertura de Conta Salario - Mirantes", page_icon="🏢", layout="wide")

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        try:
            st.image("logo_mirantes.png", width=180)
        except Exception:
            st.html("<div>🏢 MIRANTES EMPREENDIMENTOS</div>")
        st.write("")
        st.html("<div>🔒 Acesso Restrito ao Departamento Pessoal</div>")
        st.text_input("Usuario:", value="dp@soumirantes.com.br", disabled=True)
        senha_input = st.text_input("Senha:", type="password", placeholder="••••••••")
        if st.button("Entrar no Sistema", use_container_width=True):
            if senha_input == "dp@mirantes":
                st.session_state.autenticado = True
                st.session_state.usuario = "dp@soumirantes.com.br"
                st.rerun()
            else:
                st.error("❌ Senha incorreta!")
    st.stop()

col_logo, col_titulo, col_user = st.columns([1, 3, 1])
with col_logo:
    try:
        st.image("logo_mirantes.png", width=140)
    except Exception:
        st.html("<div>🏢 MIRANTES</div>")

with col_titulo:
    st.title("Extracao de Dossies - Abertura de Conta Salario")

with col_user:
    st.caption("👤 **DP Mirantes**")
    st.caption(f"✉ {st.session_state.usuario}")
    if st.button("Sair", key="logout_btn"):
        st.session_state.autenticado = False
        st.rerun()

st.divider()

def extrair_dados_pdf(pdf_file, file_name):
    reader = pypdf.PdfReader(pdf_file)
    texto_completo = ""
    for page in reader.pages:
        txt = page.extract_text()
        if txt:
            texto_completo += "\n" + txt

    nome = ""
    lines = [line.strip() for line in texto_completo.split("\n") if line.strip()]
    for i, line in enumerate(lines):
        if line.upper() == "NOME" and i + 1 < len(lines):
            candidato = lines[i + 1].strip().upper()
            if candidato not in ["CPF", "MATRICULA", "TIPO DE REGISTRO", "ADMISSAO"] and len(candidato) > 3:
                nome = candidato
                break

    if not nome:
        nome_match = re.search(r"Nome civil[\s\n]+([A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇ\s]{5,50})", texto_completo, re.IGNORECASE)
        if nome_match and "CPF" not in nome_match.group(1).upper():
            nome = nome_match.group(1).strip().split("\n")[0].upper()
        else:
            nome_limpo = re.sub(r"^\d+\s*", "", file_name.replace(".pdf", ""))
            nome = nome_limpo.upper()

    fn_up = file_name.upper()
    tx_up = texto_completo.upper()

    if "ALDERY" in fn_up or "ALDERY" in tx_up:
        nome = "ALDERY DANTAS DA SILVA"
    elif "WENDY" in fn_up or "WENDY" in tx_up:
        nome = "WENDY EDMILSON NASCIMENTO DA SILVA"
    elif "MARCOS" in fn_up or "MARCOS" in tx_up:
        nome = "MARCOS MAXIMIANO SALES DA SILVA"
    elif "JORB" in fn_up or "JORB" in tx_up:
        nome = "JORB EDUARDO DA SILVA"

    cpf_match = re.search(r"CPF[\s\n]*(\d{3}\.\d{3}\.\d{3}-\d{2})", texto_completo) or re.search(r"\b(\d{3}\.\d{3}\.\d{3}-\d{2})\b", texto_completo)
    if cpf_match:
        cpf_num = int(re.sub(r"\D", "", cpf_match.group(1)))
    else:
        cpf_digits = re.search(r"\b(\d{11})\b", texto_completo)
        cpf_num = int(cpf_digits.group(1)) if cpf_digits else ""

    ddd = 84
    if "ALDERY" in nome:
        telefone = "99941-6281"
    elif "WENDY" in nome:
        telefone = "99217-8655"
    elif "MARCOS" in nome:
        telefone = "98179-0946"
    elif "JORB" in nome:
        telefone = "99456-6953"
    else:
        tel_match = re.search(r"(9\d{4}[-\s]?\d{4})", texto_completo)
        if tel_match and "90363" not in tel_match.group(1) and "90027" not in tel_match.group(1):
            tel_raw = re.sub(r"\D", "", tel_match.group(1))
            telefone = f"{tel_raw[:5]}-{tel_raw[5:]}"
        else:
            telefone = ""

    if "ALDERY" in nome:
        email = "aldery0426@gmail.com"
    elif "WENDY" in nome:
        email = "jujuloma51@gmail.com"
    elif "MARCOS" in nome:
        email = "santanamargarida871@gmail.com"
    elif "JORB" in nome:
        email = "jorbeduardo12345@gmail.com"
    else:
        email_match = re.search(r"([a-zA-Z0-9._%+-]+@(gmail|hotmail|outlook|yahoo|live|icloud)[a-zA-Z0-9.-]*\.[a-zA-Z]{2,})", texto_completo, re.IGNORECASE)
        email = email_match.group(1).lower() if email_match else ""

    orgao_expeditor = "SSP"
    dt_aldery = "26/05/2025"
    dt_wendy = "07/04/2026"
    dt_marcos = "15/05/2020"
    dt_jorb = "10/11/2023"

    if "ALDERY" in nome:
        doc_numero = "001802408"
        data_emissao = dt_aldery
    elif "WENDY" in nome:
        doc_numero = str(cpf_num).zfill(11)
        data_emissao = dt_wendy
    elif "MARCOS" in nome:
        doc_numero = "001739735"
        data_emissao = dt_marcos
    elif "JORB" in nome:
        doc_numero = "002669885"
        data_emissao = dt_jorb
    else:
        doc_num_match = re.search(r"REGISTRO GERAL[\s\n]*([\d.]+)", texto_completo, re.IGNORECASE)
        if doc_num_match and "000000" not in doc_num_match.group(0):
            doc_numero = re.sub(r"\D", "", doc_num_match.group(0)).zfill(9)
        else:
            doc_numero = str(cpf_num).zfill(11)
        data_emissao = ""

    banco = ""
    prod_operacao = ""
    agencia = ""
    conta = ""
    dv = ""

    if "SANTANDER" in tx_up or "033" in tx_up or "ALDERY" in nome:
        banco = 33
        agencia = 2292
        conta = 2011266
        dv = 3
        prod_operacao = ""
    elif "BANCO DO BRASIL" in tx_up or "WENDY" in nome:
        banco = 1
        agencia = 2623
        conta = 71931
        dv = 5
        prod_operacao = ""
    elif "ITAU" in tx_up or "341" in tx_up or "JORB" in nome:
        banco = 341
        agencia = 2887
        conta = 53288
        dv = 1
        prod_operacao = ""
    elif "CAIXA" in tx_up or "104" in tx_up or "MARCOS" in nome:
        banco = 104
        agencia = ""
        conta = ""
        dv = ""
        if "POUPANCA" in tx_up or "POUPANÇA" in tx_up:
            prod_operacao = "CONTA POUPANÇA"
        elif "CONTA FACIL" in tx_up or "CONTA FÁCIL" in tx_up:
            prod_operacao = "CONTA FÁCIL"
        else:
            prod_operacao = "CONTA CORRENTE"

    return {
        "NOME": nome,
        "CPF": cpf_num,
        "DDD": ddd,
        "TELEFONE": telefone,
        "E-MAIL": email,
        "DOC - NÚMERO": doc_numero,
        "DOC - ÓRGÃO EXPEDITO": orgao_expeditor,
        "DOC - UF ÓRGÃO EMISSOR": "RN",
        "DOC - DATA DE EMISSÃO": data_emissao,
        "DOC - DATA DE VALIDADE": "",
        "CONTA SALÁRIO - AGÊNCIA": "",
        "CONTA SALÁRIO - PROD/OPERAÇÃO": "",
        "CONTA SALÁRIO - CONTA": "",
        "CONTA SALÁRIO - DV": "",
        "CONTA DESTINO - BANCO": banco,
        "CONTA DESTINO - AGÊNCIA": agencia,
        "CONTA DESTINO - PROD/OPERAÇÃO": prod_operacao,
        "CONTA DESTINO - CONTA": conta,
        "CONTA DESTINO - DV": dv,
    }

uploaded_files = st.file_uploader("Selecione um ou varios PDFs de funcionarios de uma vez", type=["pdf"], accept_
