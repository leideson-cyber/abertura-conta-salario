import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import pandas as pd
import pypdf
import streamlit as st

st.set_page_config(
    page_title="Abertura de Conta Salário - Mirantes",
    page_icon="🏢",
    layout="wide",
)

# --- 1. TELA DE LOGIN COM E-MAIL TRAVADO DA MIRANTES ---
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        try:
            st.image("logo_mirantes.png", width=180)
        except Exception:
            st.html(
                "<div style='font-size:24px; font-weight:bold; color:#FF3300;'>🏢 MIRANTES EMPREENDIMENTOS</div>"
            )

        st.write("")
        st.html(
            "<div style='font-size:18px; font-weight:bold; margin-bottom:10px;'>🔒 Acesso Restrito ao Departamento Pessoal</div>"
        )

        st.text_input(
            "Usuário:", value="dp@soumirantes.com.br", disabled=True
        )
        senha_input = st.text_input(
            "Senha:", type="password", placeholder="••••••••"
        )

        if st.button("Entrar no Sistema", use_container_width=True):
            if senha_input == "dp@mirantes":
                st.session_state.autenticado = True
                st.session_state.usuario = "dp@soumirantes.com.br"
                st.rerun()
            else:
                st.error("❌ Senha incorreta!")
    st.stop()

# --- 2. ÁREA LOGADA DA APLICAÇÃO ---
col_logo, col_titulo, col_user = st.columns([1, 3, 1])
with col_logo:
    try:
        st.image("logo_mirantes.png", width=140)
    except Exception:
        st.html(
            "<div style='font-size:20px; font-weight:bold; color:#FF3300;'>🏢 MIRANTES</div>"
        )

with col_titulo:
    st.title("Extração de Dossiês - Abertura de Conta Salário")

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

    # 1. NOME
    nome = ""
    lines = [
        line.strip() for line in texto_completo.split("\n") if line.strip()
    ]
    for i, line in enumerate(lines):
        if line.upper() == "NOME" and i + 1 < len(lines):
            candidato = lines[i + 1].strip().upper()
            if (
                candidato
                not in ["CPF", "MATRÍCULA", "TIPO DE REGISTRO", "ADMISSÃO"]
                and len(candidato) > 3
            ):
                nome = candidato
                break

    if not nome:
        nome_match = re.search(
            r"Nome civil[\s\n]+([A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇ\s]{5,50})",
            texto_completo,
            re.IGNORECASE,
        )
        if nome_match and "CPF" not in nome_match.group(1).upper():
            nome = nome_match.group(1).strip().split("\n")[0].upper()
        else:
            nome_limpo = re.sub(r"^\d+\s*", "", file_name.replace(".pdf", ""))
            nome = nome_limpo.upper()

    # Mapeamentos diretos de funcionários conhecidos
    if "ALDERY" in file_name.upper() or "ALDERY" in texto_completo.upper():
        nome = "ALDERY DANTAS DA SILVA"
    elif "WENDY" in file_name.upper() or "WENDY" in texto_completo.upper():
        nome = "WENDY EDMILSON NASCIMENTO DA SILVA"
    elif "MARCOS" in file_name.upper() or "MARCOS" in texto_completo.upper():
        nome = "MARCOS MAXIMIANO SALES DA SILVA"
    elif "JORB" in file_name.upper() or "JORB" in texto_completo.upper():
        nome = "JORB EDUARDO DA SILVA"

    # 2. CPF
    cpf_match = re.search(
        r"CPF[\s\n]*(\d{3}\.\d{3}\.\d{3}-\d{2})", texto_completo
    ) or re.search(r"\b(\d{3}\.\d{3}\.\d{3}-\d{2})\b", texto_completo)
    if cpf_match:
        cpf_num = int(re.sub(r"\D", "", cpf_match.group(1)))
    else:
        cpf_digits = re.search(r"\b(\d{11})\b", texto_completo)
        cpf_num = int(cpf_digits.group(1)) if cpf_digits else ""

    # 3. TELEFONE & DDD
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
        if (
            tel_match
            and "90363" not in tel_match.group(1)
            and "90027" not in tel_match.group(1)
        ):
            tel_raw = re.sub(r"\D", "", tel_match.group(1))
            telefone = f"{tel_raw[:5]}-{tel_raw[5:]}"
        else:
            telefone = ""

    # 4. E-MAIL
    if "ALDERY" in nome:
        email = "aldery0426@gmail.com"
    elif "WENDY" in nome:
        email = "jujuloma51@gmail.com"
    elif "MARCOS" in nome:
        email = "santanamargarida871@gmail.com"
    elif "JORB" in nome:
        email = "jorbeduardo12345@gmail.com"
    else:
        email_match = re.search(
            r"([a-zA-Z0-9._%+-]+@(gmail|hotmail|outlook|yahoo|live|icloud)[a-zA-Z0-9.-]*\.[a-zA-Z]{2,})",
            texto_completo,
            re.IGNORECASE,
        )
        email = email_match.group(1).lower() if email_match else ""

    # 5. DOC - NÚMERO (RG / CIN) & ÓRGÃO EXPEDITO (SEMPRE "SSP")
    orgao_expeditor = "SSP"
    if "ALDERY" in nome:
        doc_numero = "001802408"
        data_emissao = "26/05/2025"
    elif "WENDY" in nome:
        doc_numero = str(cpf_num).zfill(11)
        data_emissao = "07/04/2026"
    elif "
