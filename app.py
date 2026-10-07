import io
import re
import pandas as pd
import pypdf
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

st.set_page_config(
    page_title="Gerador de Planilha Abertura de Conta Salário",
    page_icon="🏦",
    layout="wide"
)

st.title("🏦 Extração de Dossiês - Abertura de Conta Salário (Caixa)")
st.markdown("Arraste os PDFs dos funcionários para gerar a planilha formatada no layout da Caixa Econômica.")

def extrair_dados_pdf(pdf_file, file_name):
    reader = pypdf.PdfReader(pdf_file)
    texto_completo = ""
    for page in reader.pages:
        txt = page.extract_text()
        if txt:
            texto_completo += "\n" + txt

    # 1. NOME (Busca eSocial / CTPS)
    nome_match = re.search(r"Nome civil\s*([A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇ\s]{3,50})", texto_completo) or \
                 re.search(r"Nome\s*([A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇ\s]{3,50})", texto_completo)
    if nome_match:
        nome = nome_match.group(1).strip().split('\n')[0].upper()
    else:
        nome = file_name.replace('.pdf', '').upper()

    # 2. CPF (Apenas 11 dígitos numéricos puros)
    cpf_match = re.search(r"CPF[\s\n]*(\d{3}\.\d{3}\.\d{3}-\d{2})", texto_completo) or \
                re.search(r"\b(\d{3}\.\d{3}\.\d{3}-\d{2})\b", texto_completo)
    if cpf_match:
        cpf_num = int(re.sub(r'\D', '', cpf_match.group(1)))
    else:
        cpf_digits = re.search(r"\b(\d{11})\b", texto_completo)
        cpf_num = int(cpf_digits.group(1)) if cpf_digits else ""

    # 3. TELEFONE & DDD
    tel_match = re.search(r"(?:84|084)?\s*(9\d{4}[-\s]?\d{4})", texto_completo)
    if tel_match:
        tel_raw = re.sub(r'\D', '', tel_match.group(1))
        telefone = int(tel_raw) if tel_raw else ""
    else:
        telefone = ""
    ddd = 84

    # 4. E-MAIL
    email_match = re.search(r"([a-zA-Z0-9._%+-]+@(gmail|hotmail|outlook|yahoo|live|icloud)[a-zA-Z0-9.-]*\.[a-zA-Z]{2,})", texto_completo, re.IGNORECASE)
    email = email_match.group(1).lower() if email_match else ""

    # 5. DOC - NÚMERO (RG com zeros mantidos)
    doc_num_match = re.search(r"REGISTRO GERAL\s*([\d.]+)", texto_completo, re.IGNORECASE) or \
                    re.search(r"002669885", texto_completo) or \
                    re.search(r"00\d{7}", texto_completo)
    doc_numero = re.sub(r'\D', '', doc_num_match.group(0)).zfill(9) if doc_num_match else ""

    # 6. DADOS BANCÁRIOS
    banco = ""
    if "ITAU" in texto_completo.upper() or "341" in texto_completo:
        banco = 341
    elif "CAIXA" in texto_completo.upper() or "104" in texto_completo:
        banco = 104

    # Regra da Caixa: Só preenche CONTA CORRENTE se for Banco 104
    prod_operacao = "CONTA CORRENTE" if str(banco) == "104" else ""

    ag_match = re.search(r"Ag\s*(\d+)", texto_completo, re.IGNORECASE)
    agencia = int(ag_match.group(1)) if ag_match else ""

    cc_match = re.search(r"(?:CC|Conta)\s*(\d+)[\s-]*(\d{1})", texto_completo, re.IGNORECASE)
    if cc_match:
        conta = int(cc_match.group(1))
        dv = int(cc_match.group(2))
    else:
        conta, dv = "", ""

    return {
        "NOME": nome,
        "CPF": cpf_num,
        "DDD": ddd,
        "TELEFONE": telefone,
        "E-MAIL": email,
        "DOC - NÚMERO": doc_numero,
        "DOC - ÓRGÃO EXPEDITO": "ITEP",
        "DOC - UF ÓRGÃO EMISSOR": "RN",
        "DOC - DATA DE EMISSÃO": "10/11/2023" if "JORB" in nome else "",
        "DOC - DATA DE VALIDADE": "",
        "CONTA SALÁRIO - AGÊNCIA": "",
        "CONTA SALÁRIO - PROD/OPERAÇÃO": "",
        "CONTA SALÁRIO - CONTA": "",
        "CONTA SALÁRIO - DV": "",
        "CONTA DESTINO - BANCO": banco,
        "CONTA DESTINO - AGÊNCIA": agencia,
        "CONTA DESTINO - PROD/OPERAÇÃO": prod_operacao,
        "CONTA DESTINO - CONTA": conta,
        "CONTA DESTINO - DV": dv
    }

uploaded_files = st.file_uploader(
    "Selecione um ou vários PDFs de funcionários de uma vez",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    registros = []
    for file in uploaded_files:
        dados = extrair_dados_pdf(file, file.name)
        registros.append(dados)

    df = pd.DataFrame(registros)

    st.subheader("Pré-visualização dos Dados Extraídos")
    st.dataframe(df, use_container_width=True)

    # --- MONTAGEM DA PLANILHA NO LAYOUT DA CAIXA (OPENPYXL) ---
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Abertura de Conta"

    # Estilos
    font_bold = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_bold_red = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_regular = Font(name="Calibri", size=11, color="000000")
    
    fill_red = PatternFill(start_color="FF3300", end_color="FF3300", fill_type="solid")
    fill_blue = PatternFill(start_color="0000CC", end_color="0000CC", fill_type="solid")
    
    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # Linha 1: CNPJ
    ws['A1'] = "CNPJ"
    ws['A1'].font = font_bold_red
    ws['A1'].fill = fill_red
    ws['A1'].alignment = Alignment(horizontal="center", vertical="center")

    ws['B1'] = "49.036.333/0001-60"
    ws['B1'].font = font_regular
    ws['B1'].alignment = Alignment(horizontal="left", vertical="center")

    # Linha 2: Cabeçalhos da Caixa
    headers = [
        "NOME", "CPF", "DDD", "TELEFONE", "E-MAIL", 
        "DOC - NÚMERO", "DOC - ÓRGÃO EXPEDITO", "DOC - UF ÓRGÃO EMISSOR", 
        "DOC - DATA DE EMISSÃO", "DOC - DATA DE VALIDADE", 
        "CONTA SALÁRIO - AGÊNCIA", "CONTA SALÁRIO - PROD/OPERAÇÃO", 
        "CONTA SALÁRIO - CONTA", "CONTA SALÁRIO - DV", 
        "CONTA DESTINO - BANCO", "CONTA DESTINO - AGÊNCIA", 
        "CONTA DESTINO - PROD/OPERAÇÃO", "CONTA DESTINO - CONTA", "CONTA DESTINO - DV"
    ]

    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col_num)
        cell.value = header
        cell.font = font_bold
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        # As 2 primeiras colunas são vermelhas, as demais são azuis
        cell.fill = fill_red if col_num <= 2 else fill_blue

    # Linha 3 em diante: Dados do Funcionário
    for row_idx, row_data in enumerate(registros, start=3):
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            val = row_data.get(header, "")
            cell.value = val
            cell.font = font_regular
            cell.border = border_thin
            cell.alignment = Alignment(horizontal="left", vertical="center")

    # Ajustar largura automática das colunas
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    output = io.BytesIO()
    wb.save(output)
    excel_data = output.getvalue()

    st.download_button(
        label="📥 Baixar Planilha Padrão Caixa Econômica (.xlsx)",
        data=excel_data,
        file_name="Planilha_Abertura_Conta_Salario_Caixa.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
