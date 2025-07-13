import streamlit as st
import pandas as pd
import io

# --- Configuração da Página do App ---
st.set_page_config(
    page_title="Formatador de Rotas",
    page_icon="🚚",
    layout="centered"
)

st.title("🚚 Formatador de Rotas para Circuit")
st.write("Faça o upload da sua planilha Excel para agrupar e formatar os pacotes por parada.")

# --- Funções auxiliares ---
def formatar_cep(cep):
    cep = ''.join(filter(str.isdigit, str(cep)))
    return f"{cep[:5]}-{cep[5:]}" if len(cep) == 8 else cep

def limpar_float_texto(valor):
    valor_str = str(valor)
    return valor_str.replace('.0', '') if valor_str.endswith('.0') else valor_str

# --- Função principal ---
def processar_arquivo_excel(arquivo_bytes):
    try:
        df = pd.read_excel(arquivo_bytes)
    except Exception as e:
        st.error(f"❌ Erro ao ler o arquivo Excel: {e}")
        return None

    mapa_colunas = {
        'N.º': 'Parada', 'ID do pacote': 'ID do Pacote', 'Endereço': 'Endereco',
        'N.º.1': 'Numero', 'Bairro': 'Bairro', 'Cidade': 'Cidade', 'CEP': 'CEP'
    }
    df.rename(columns=mapa_colunas, inplace=True)

    colunas_essenciais = ['Parada', 'ID do Pacote', 'Endereco', 'Numero', 'Bairro', 'CEP']
    if not all(col in df.columns for col in colunas_essenciais):
        st.error("⚠️ Colunas esperadas não encontradas. Verifique o nome das colunas no seu arquivo original.")
        return None

    # Limpeza e formatação
    for col in ['Numero', 'CEP', 'ID do Pacote', 'Parada']:
        if col in df.columns:
            df[col] = df[col].apply(limpar_float_texto)

    df['Address Line'] = df['Endereco'].astype(str).str.strip() + ', ' + df['Numero'].astype(str).str.strip()
    df['CEP'] = df['CEP'].apply(formatar_cep)

    df_formatado = pd.DataFrame({
        'Parada': df['Parada'], 'ID do Pacote': df['ID do Pacote'], 'Address Line': df['Address Line'],
        'Secondary Address Line': df['Bairro'], 'City': df.get('Cidade', 'São José dos Campos'),
        'State': 'São Paulo', 'Zip Code': df['CEP']
    })

    # Agrupamento por número base
    df_formatado['Parada_Base'] = df_formatado['Parada'].astype(str).str.extract(r'(\d+)')
    
    linhas_invalidas = df_formatado['Parada_Base'].isnull()
    if linhas_invalidas.any():
        st.warning(f"Atenção: {linhas_invalidas.sum()} linha(s) foram ignoradas por não terem um número de parada válido.")
        df_formatado.dropna(subset=['Parada_Base'], inplace=True)

    if df_formatado.empty:
        st.error("Nenhuma linha válida encontrada para processar após a limpeza.")
        return None
        
    df_formatado['Parada_Base'] = df_formatado['Parada_Base'].astype(int)

    df_agrupado = df_formatado.groupby('Parada_Base').agg({
        'ID do Pacote': lambda x: ', '.join(x.astype(str)),
        'Address Line': 'first',
        'Secondary Address Line': 'first',
        'City': 'first',
        'State': 'first',
        'Zip Code': 'first'
    }).reset_index()

    # Formatando colunas conforme solicitado
    df_agrupado['Total de Pacotes'] = df_agrupado['ID do Pacote'].str.split(',').str.len().astype(int)
    df_agrupado['Total de Pacotes'] = df_agrupado['Total de Pacotes'].apply(
        lambda x: f"{x} pacote" if x == 1 else f"{x} pacotes"
    )
    df_agrupado.rename(columns={'Parada_Base': 'Parada_Num'}, inplace=True)
    df_agrupado['Parada'] = df_agrupado['Parada_Num'].apply(lambda x: f"Parada {x}")

    colunas_finais = ['Parada', 'ID do Pacote', 'Total de Pacotes', 'Address Line', 'Secondary Address Line', 'City', 'State', 'Zip Code']
    return df_agrupado[colunas_finais]

# --- Interface do App ---
uploaded_file = st.file_uploader(
    "Escolha o arquivo Excel da rota (.xlsx)",
    type="xlsx"
)

if uploaded_file is not None:
    st.info("🔄 Processando o arquivo... Aguarde.")
    
    df_final = processar_arquivo_excel(uploaded_file)
    
    if df_final is not None:
        st.success("✅ Arquivo processado com sucesso!")
        st.dataframe(df_final)

        # Criar arquivo Excel em memória
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='RotaFormatada')
        excel_bytes = output.getvalue()

        # Botão de download
        st.download_button(
            label="⬇️ Baixar Arquivo Formatado (.xlsx)",
            data=excel_bytes,
            file_name="rota_formatada_agrupada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
