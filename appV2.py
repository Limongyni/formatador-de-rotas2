import streamlit as st
import pandas as pd
import requests
import pdfplumber
import io
import time
import os
import socket

# --- Utilitário para checar se há internet ---
def internet_disponivel():
    try:
        socket.create_connection(("8.8.8.8", 53), timeout=2)
        return True
    except OSError:
        return False

# --- Funções auxiliares ---
def formatar_cep(cep):
    cep = ''.join(filter(str.isdigit, str(cep)))
    return f"{cep[:5]}-{cep[5:]}" if len(cep) == 8 else cep

def limpar_float_texto(valor):
    valor_str = str(valor)
    return valor_str.replace('.0', '') if valor_str.endswith('.0') else valor_str

@st.cache_data(show_spinner=False)
def buscar_endereco_por_cep(cep):
    cep = ''.join(filter(str.isdigit, str(cep)))
    if len(cep) != 8:
        return {'logradouro': '', 'bairro': '', 'localidade': ''}
    try:
        time.sleep(0.2)  # evita bloqueio da API
        response = requests.get(f"https://viacep.com.br/ws/{cep}/json/", timeout=10)
        response.raise_for_status()
        data = response.json()
        return {
            'logradouro': data.get('logradouro', ''),
            'bairro': data.get('bairro', ''),
            'localidade': data.get('localidade', '')
        }
    except requests.exceptions.RequestException as e:
        st.warning(f"⚠️ Erro ao consultar CEP {cep}: {e}")
        return {'logradouro': '', 'bairro': '', 'localidade': ''}

def extrair_tabela_pdf(arquivo_pdf):
    dados = []
    with pdfplumber.open(arquivo_pdf) as pdf:
        for pagina in pdf.pages:
            tabela = pagina.extract_table()
            if tabela:
                for linha in tabela[1:]:
                    if any(linha):
                        dados.append(linha)
    colunas = ['Parada', 'ID do Pacote', 'Cliente', 'Endereco', 'Numero', 'Complemento', 'Bairro', 'Cidade', 'CEP', 'Tipo', 'Assinatura']
    df = pd.DataFrame(dados, columns=colunas[:len(dados[0])])
    return df

def processar_dataframe(df):
    for col in ['Numero', 'CEP', 'ID do Pacote', 'Parada']:
        if col in df.columns:
            df[col] = df[col].apply(limpar_float_texto)

    df['CEP'] = df['CEP'].apply(formatar_cep)

    ceps_unicos = df['CEP'].dropna().unique()
    total = len(ceps_unicos)
    st.info(f"🔎 Consultando endereços para {total} CEP(s) único(s)...")

    enderecos_dict = {}
    progress = st.progress(0)

    if internet_disponivel():
        for i, cep in enumerate(ceps_unicos):
            enderecos_dict[cep] = buscar_endereco_por_cep(cep)
            progress.progress((i + 1) / total)
    else:
        st.warning("🔌 Sem conexão com a internet. Pulando consulta de CEPs.")
        for cep in ceps_unicos:
            enderecos_dict[cep] = {'logradouro': '', 'bairro': '', 'localidade': ''}

    df['Endereco'] = df['CEP'].apply(lambda c: enderecos_dict.get(c, {}).get('logradouro', ''))
    df['Bairro'] = df['CEP'].apply(lambda c: enderecos_dict.get(c, {}).get('bairro', ''))
    df['Cidade'] = df['CEP'].apply(lambda c: enderecos_dict.get(c, {}).get('localidade', ''))

    df['Address Line'] = df['Endereco'].astype(str).str.strip() + ', ' + df['Numero'].astype(str).str.strip()

    df_formatado = pd.DataFrame({
        'Parada': df['Parada'],
        'ID do Pacote': df['ID do Pacote'],
        'Address Line': df['Address Line'],
        'Complemento': df.get('Complemento', ''),
        'Secondary Address Line': df['Bairro'],
        'City': df['Cidade'],
        'State': 'São Paulo',
        'Zip Code': df['CEP']
    })

    df_formatado['Parada_Base'] = df_formatado['Parada'].astype(str).str.extract(r'(\d+)')
    df_formatado.dropna(subset=['Parada_Base'], inplace=True)
    df_formatado['Parada_Base'] = df_formatado['Parada_Base'].astype(int)

    df_agrupado = df_formatado.groupby('Parada_Base').agg({
        'ID do Pacote': lambda x: ', '.join(x.astype(str)),
        'Address Line': 'first',
        'Complemento': 'first',
        'Secondary Address Line': 'first',
        'City': 'first',
        'State': 'first',
        'Zip Code': 'first'
    }).reset_index()

    df_agrupado['Total de Pacotes'] = df_agrupado['ID do Pacote'].str.split(',').str.len()
    df_agrupado['Total de Pacotes'] = df_agrupado['Total de Pacotes'].apply(
        lambda x: f"{x} pacote" if x == 1 else f"{x} pacotes"
    )
    df_agrupado['Parada'] = df_agrupado['Parada_Base'].apply(lambda x: f"Parada {x}")

    colunas_finais = [
        'Parada',
        'ID do Pacote',
        'Total de Pacotes',
        'Address Line',
        'Complemento',
        'Secondary Address Line',
        'City',
        'State',
        'Zip Code'
    ]

    return df_agrupado[colunas_finais]

# --- App Streamlit ---
st.set_page_config(page_title="Formatador de Rota", layout="centered")

st.title("📦 Formatador de Rota com Agrupamento")
st.write("Envie um arquivo `.pdf` ou `.xlsx` com a rota, e baixe o arquivo formatado com agrupamento por parada.")

arquivo = st.file_uploader("Selecione o arquivo de rota", type=["pdf", "xlsx"])

if arquivo:
    nome_base = os.path.splitext(arquivo.name)[0]
    st.success(f"📁 Arquivo carregado: {arquivo.name}")

    if arquivo.name.endswith(".pdf"):
        st.info("🔍 Extraindo dados do PDF...")
        df_raw = extrair_tabela_pdf(arquivo)
    elif arquivo.name.endswith(".xlsx"):
        st.info("📖 Lendo planilha Excel...")
        df_raw = pd.read_excel(arquivo)
        df_raw = df_raw.iloc[1:]
        df_raw.columns = df_raw.columns.str.strip()
    else:
        st.error("❌ Formato de arquivo não suportado.")
        st.stop()

    st.info("⚙️ Processando dados...")
    df_final = processar_dataframe(df_raw)

    if df_final.empty:
        st.warning("⚠️ Nenhum dado processado.")
    else:
        st.success("✅ Dados processados com sucesso!")
        st.dataframe(df_final)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='RotaFormatada')
        buffer.seek(0)

        nome_saida = f"{nome_base}_rota_formatada.xlsx"
        st.download_button(
            label="📥 Baixar Rota Formatada",
            data=buffer,
            file_name=nome_saida,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


