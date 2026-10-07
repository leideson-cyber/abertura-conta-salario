import io
import re
import pandas as pd
import pypdf
import streamlit as st

st.set_page_config(
    page_title="Gerador de Planilha Abertura de Conta",
    page_icon="🏦",
    layout="wide",
)

st.title("🏦 Extração de Dossiês - Abertura de Conta Salário")
st.markdown(
    "Arraste ou selecione os arquivos **PDF dos funcionários** para gerar a planilha formatada."
)


def extrair_dados_pdf(pdf_file, file_name):
    reader = pypdf.PdfReader(pdf_file)
    texto_completo = ""
    for page in reader.pages:
        txt = page.extract_text()
        if txt:
            texto_completo += "\n" + txt

    # 1. NOME
    nome_match = re.search(
        r"Nome civil\s*([A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇ\s]{3,50})", texto_completo
    ) or re.search(r"Nome\s*([A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇ\s]{3,50})", texto_completo)
    if nome_match:
        nome = nome_match.group(1).strip().split("\n")[0].upper()
    else:
        nome = file_name.replace(".pdf", "").upper()

    # 2. CPF (11 dígitos numéricos puros)
    cpf_match = re.search(r"\b(\d{3}\.\d{3}\.\d{3}-\d{2})\b", texto_completo)
    if cpf_match:
        cpf_num = re.sub(r"\D", "", cpf_match.group(1))
    else:
        cpf_digits = re.search(r"\b(\d{11})\b", texto_completo)
        cpf_num = cpf_digits.group(1) if cpf_digits else ""

    # 3. TELEFONE & DDD
    tel_match = re.search(r"(?:84|084)?\s*(9\d{4}[-\s]?\d{4})", texto_completo)
    if tel_match:
        tel_raw = re.sub(r"\D", "", tel_match.group(1))
        telefone = f"{tel_raw[:5]}-{tel_raw[5:]}"
    else:
        telefone = ""
    ddd = "084"

    # 4. E-MAIL
    email_match = re.search(
        r"([a-zA-Z0-9._%+-]+@(gmail|hotmail|outlook|yahoo|live|icloud)[a-zA-Z0-9.-]*\.[a-zA-Z]{2,})",
        texto_completo,
        re.IGNORECASE,
    )
    email = email_match.group(1).lower() if email_match else ""

    # 5. DOC - NÚMERO (RG com zeros mantidos)
    doc_num_match = re.search(
        r"REGISTRO GERAL\s*([\d.]+)", texto_completo, re.IGNORECASE
    ) or re.search(r"00\d{7}", texto_completo)
    doc_numero = (
        re.sub(r"\D", "", doc_num_match.group(0)).zfill(9)
        if doc_num_match
        else ""
    )

    # 6. DADOS BANCÁRIOS
    banco = ""
    if "ITAU" in texto_completo.upper() or "341" in texto_completo:
        banco = "341"
    elif "CAIXA" in texto_completo.upper() or "104" in texto_completo:
        banco = "104"

    # Regra Banco 104 (Caixa)
    prod_operacao = "CONTA CORRENTE" if str(banco) == "104" else ""

    ag_match = re.search(r"Ag\s*(\d+)", texto_completo, re.IGNORECASE)
    agencia = ag_match.group(1) if ag_match else ""

    cc_match = re.search(
        r"(?:CC|Conta)\s*(\d+)[\s-]*(\d{1})", texto_completo, re.IGNORECASE
    )
    if cc_match:
        conta = cc_match.group(1)
        dv = cc_match.group(2)
    else:
        conta, dv = "", ""

    return {
        "NOME": nome,
        "CPF": str(cpf_num),
        "DDD": ddd,
        "TELEFONE": telefone,
        "E-MAIL": email,
        "DOC - NÚMERO": str(doc_numero),
        "DOC - ÓRGÃO EXPEDITO": "ITEP",
        "DOC - UF ÓRGÃO EMISSOR": "RN",
        "DOC - DATA DE EMISSÃO": "",
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


uploaded_files = st.file_uploader(
    "Selecione um ou vários PDFs de funcionários de uma vez",
    type=["pdf"],
    accept_multiple_files=True,
)

if uploaded_files:
    registros = []
    for file in uploaded_files:
        dados = extrair_dados_pdf(file, file.name)
        registros.append(dados)

    df = pd.DataFrame(registros)

    st.subheader("Pré-visualização dos Dados Extraídos")
    st.dataframe(df, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, startrow=1)
        worksheet = writer.sheets["Sheet1"]
        worksheet["A1"] = "CNPJ"
        worksheet["B1"] = "49.036.333/0001-60"

    excel_data = output.getvalue()

    st.download_button(
        label="📥 Baixar Planilha Excel Formatada (.xlsx)",
        data=excel_data,
        file_name="Planilha_Abertura_Conta_Salario.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
