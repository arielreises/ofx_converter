# ============================================================
# Aplicativo: Conversor de OFX para XLSX
# Criado por: Ariel Reises
# Empresa: Reises Co.
# CNPJ: 57.975.327/0001-04
# Contato: ariel@reises.com.br
# Website: www.reises.com.br

# Uso exclusivo para: Movendo LTDA
# CNPJ: 51.265.918/0001-01
# ============================================================

import streamlit as st
import pandas as pd
from ofxparse import OfxParser
import io
import re

# Função para processar o arquivo OFX e converter para DataFrame
def process_ofx(ofx_file):
    ofx = OfxParser.parse(ofx_file)
    transactions = []
    for account in ofx.accounts:
        for transaction in account.statement.transactions:
            transactions.append({
                "Data": transaction.date.strftime("%Y-%m-%d"),
                "Valor": transaction.amount,
                "Descrição": transaction.memo,
                "Tipo": transaction.type,
                "ID": transaction.id,
            })
    return pd.DataFrame(transactions)

# Função para exibir cabeçalho do aplicativo
def display_app_header():
    st.title("Conversor de OFX para XLSX")
    st.write("Este aplicativo converte arquivos OFX em arquivos XLSX.")
    st.write("**Movendo LTDA.**")
    st.markdown("---")

# Exibe o cabeçalho
display_app_header()

# Campo para inserir o nome do cliente
client_name = st.text_input("Digite o nome do cliente")

# Upload do arquivo OFX
uploaded_file = st.file_uploader("Faça o upload de um arquivo OFX", type=["ofx"])

if uploaded_file is not None:
    if client_name.strip() == "":
        st.error("Por favor, insira o nome do cliente antes de prosseguir.")
    else:
        try:
            # Processa o arquivo OFX
            data = process_ofx(uploaded_file)
            st.success("Arquivo OFX processado com sucesso!")

            # Exibe o resumo
            st.write(f"Total de transações processadas: {len(data)}")
            st.dataframe(data)

            # Botão para exportar os dados como XLSX
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                data.to_excel(writer, index=False, sheet_name="Extrato")
            buffer.seek(0)

            # Nome do arquivo com validação
            sanitized_name = re.sub(r'[^a-zA-Z0-9_]', '', client_name.replace(' ', '_'))
            file_name = f"extrato_{sanitized_name}.xlsx"

            st.download_button(
                label="Baixar como XLSX",
                data=buffer,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Ocorreu um erro ao processar o arquivo: {e}")
