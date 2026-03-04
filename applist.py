import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import re
import unicodedata

# Configuração
st.set_page_config(page_title="PISF - Gerador de Termos", page_icon="💧")

URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQk-RTbvVDPlwxIJFaEKeR1WPRaNSFGioF8DIYD1_mQ-M6a7O20-7TXmx8fBAlDg/pub?gid=502195603&single=true&output=csv"

def normalizar_coluna(txt):
    """Remove acentos, símbolos e espaços para criar tags limpas"""
    if not isinstance(txt, str): return str(txt)
    # Remove acentos
    txt = "".join(c for c in unicodedata.normalize('NFD', txt) if unicodedata.category(c) != 'Mn')
    # Remove símbolos como ⌀, ( ), \n e espaços extras
    txt = txt.replace('\n', ' ').replace('⌀', ' ').strip()
    txt = re.sub(r'[^\w\s]', '', txt)
    txt = txt.replace(' ', '_').upper()
    
    # Mapeamento de nomes específicos para o Template
    correcoes = {
        'PROPRIETARIO': 'PROPRIETARIO',
        'TUBO_MM': 'TUBO_MM',
        'ESTRUTURA_WBS': 'ESTRUTURA_WBS',
        'VAZAO_ESTIMADA_M3MES': 'VAZAO_ESTIMADA'
    }
    return correcoes.get(txt, txt)

@st.cache_data(ttl=300)
def carregar_planilha():
    # Carrega pulando as 5 linhas de metadados
    df = pd.read_csv(URL, skiprows=5)
    # Aplica a normalização em todas as colunas
    df.columns = [normalizar_coluna(c) for c in df.columns]
    # Garante que o ID seja tratado como texto para busca exata
    df['ID'] = df['ID'].astype(str).str.replace('.0', '', regex=False)
    return df

st.title("📄 Emissor de Termos PISF")

try:
    df = carregar_planilha()
    
    id_busca = st.text_input("Insira o ID da Captação:", placeholder="Ex: 562026")

    if id_busca:
        registro = df[df['ID'] == id_busca]

        if not registro.empty:
            dados = registro.iloc[0].to_dict()
            
            # Preview dos dados
            st.success(f"Usuário encontrado: {dados['PROPRIETARIO']}")
            col1, col2 = st.columns(2)
            col1.metric("Município", dados['MUNICIPIO'])
            col2.metric("Situação", dados['SITUACAO'])

            if st.button("Gerar Termo de Entrega"):
                doc = DocxTemplate("template_pisf.docx")
                doc.render(dados)
                
                output = io.BytesIO()
                doc.save(output)
                output.seek(0)
                
                st.download_button(
                    label="📥 Baixar Documento Preenchido",
                    data=output,
                    file_name=f"Termo_ID_{id_busca}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        else:
            st.error("ID não localizado na base de dados do Eixo Leste.")

except Exception as e:
    st.error(f"Erro na conexão ou processamento: {e}")