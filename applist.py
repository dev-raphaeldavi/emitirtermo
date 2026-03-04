import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import re
import unicodedata

# Configuração da página
st.set_page_config(page_title="PISF - Gerador de Termos", page_icon="💧")

URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQk-RTbvVDPlwxIJFaEKeR1WPRaNSFGioF8DIYD1_mQ-M6a7O20-7TXmx8fBAlDg/pub?gid=502195603&single=true&output=csv"

def normalizar_coluna(txt):
    """Remove acentos, símbolos e espaços para criar tags limpas APENAS no cabeçalho"""
    if not isinstance(txt, str): return str(txt)
    txt = "".join(c for c in unicodedata.normalize('NFD', txt) if unicodedata.category(c) != 'Mn')
    txt = txt.replace('\n', ' ').replace('⌀', ' ').strip()
    txt = re.sub(r'[^\w\s]', '', txt)
    txt = txt.replace(' ', '_').upper()
    
    correcoes = {
        'PROPRIETARIO': 'PROPRIETARIO',
        'TUBO_MM': 'TUBO_MM',
        'ESTRUTURA_WBS': 'ESTRUTURA_WBS',
        'VAZAO_ESTIMADA_M3MES': 'VAZAO_ESTIMADA',
        'HIDROMETRO': 'HIDROMETRO',
        'LACRE': 'LACRE',
        'SEQUENCIAL': 'SEQUENCIAL'
    }
    return correcoes.get(txt, txt)

@st.cache_data(ttl=300)
def carregar_planilha():
    # keep_default_na=False e dtype=str PROÍBEM o sistema de adivinhar ou formatar dados. Tudo vira texto bruto.
    df = pd.read_csv(URL, skiprows=5, dtype=str, keep_default_na=False)
    
    # Reforço de segurança para garantir que nenhum dado na linha seja alterado
    df = df.astype(str)
    
    # Aplica a normalização exclusivamente nos títulos das colunas
    df.columns = [normalizar_coluna(c) for c in df.columns]
    
    # Limpa a coluna ID para a busca funcionar perfeitamente
    if 'ID' in df.columns:
        df['ID'] = df['ID'].str.replace('.0', '', regex=False).str.strip()
        
    return df

st.title("📄 Emissor de Termos PISF")

try:
    df = carregar_planilha()
    
    id_busca = st.text_input("Insira o ID da Captação:", placeholder="Ex: 562026")

    if id_busca:
        registro = df[df['ID'] == id_busca]

        if not registro.empty:
            # Extrai os dados EXATOS da planilha como texto
            dados = registro.iloc[0].to_dict()
            
            # Preview dos dados
            st.success(f"Usuário encontrado: {dados.get('PROPRIETARIO', 'Não informado')}")
            col1, col2 = st.columns(2)
            col1.metric("Município", dados.get('MUNICIPIO', '-'))
            col2.metric("Situação", dados.get('SITUACAO', '-'))

            if st.button("Gerar Termo de Entrega"):
                try:
                    # Carrega o template Word
                    doc = DocxTemplate("template_pisf.docx")
                    
                    # Preenche as tags {{ }} com os dados exatos da planilha
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
                except Exception as erro_word:
                    st.error(f"Erro ao preencher o Word. Verifique o template: {erro_word}")
        else:
            st.error("ID não localizado na base de dados do Eixo Leste.")

except Exception as e:
    st.error(f"Erro na conexão ou processamento: {e}")
