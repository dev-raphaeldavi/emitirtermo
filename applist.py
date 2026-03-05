import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import re
import unicodedata
import os

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="Central PISF - Eixo Leste", page_icon="💧", layout="wide")

# CSS Customizado
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        color: #1E3A8A;
        text-align: center;
        margin-top: -20px;
        margin-bottom: 0px;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #6B7280;
        text-align: center;
        margin-bottom: 2rem;
    }
    .info-box {
        background-color: #EFF6FF;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 5px solid #3B82F6;
        margin-bottom: 1rem;
        color: #1E3A8A;
    }
    </style>
""", unsafe_allow_html=True)

# Cabeçalho da Aplicação
st.markdown('<div class="main-header">Central de Emissão de Documentos PISF</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Regularização de Captações de Pequenos Usuários - Eixo Leste</div>', unsafe_allow_html=True)
st.divider()

URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSQk-RTbvVDPlwxIJFaEKeR1WPRaNSFGioF8DIYD1_mQ-M6a7O20-7TXmx8fBAlDg/pub?gid=502195603&single=true&output=csv"

# 2. FUNÇÕES DE TRATAMENTO
def normalizar_coluna(txt):
    if not isinstance(txt, str): return str(txt)
    txt = "".join(c for c in unicodedata.normalize('NFD', txt) if unicodedata.category(c) != 'Mn')
    txt = txt.replace('\n', ' ').replace('⌀', ' ').strip()
    txt = re.sub(r'[^\w\s]', '', txt)
    txt = txt.replace(' ', '_').upper()
    return txt

@st.cache_data(ttl=300)
def carregar_planilha():
    df = pd.read_csv(URL, skiprows=5, dtype=str, keep_default_na=False)
    df = df.astype(str)
    
    df.columns = [normalizar_coluna(c) for c in df.columns]
    
    if 'ID' in df.columns:
        df['ID'] = df['ID'].str.replace('.0', '', regex=False).str.strip()
    return df

# 3. INTERFACE E LÓGICA DE ABAS
try:
    df = carregar_planilha()
    
    # Criação das 3 Abas
    aba_termo, aba_materiais, aba_projetos = st.tabs([
        "📄 Termos de Responsabilidade", 
        "🛠️ Ficha de Materiais (Kits)", 
        "📐 Projetos de Captação"
    ])

    # ==========================================================
    # ABA 1: TERMO DE RESPONSABILIDADE
    # ==========================================================
    with aba_termo:
        st.markdown("#### Emissão do Termo de Entrega")
        st.write("Gere o documento oficial de responsabilidade e posse dos equipamentos padronizados.")
        
        c_t1, c_t2, c_t3 = st.columns([1, 2, 1])
        with c_t2:
            id_termo = st.text_input("🔍 Digite o ID da Captação para o Termo:", placeholder="Ex: 562026", key="input_termo")

        if id_termo:
            registro_termo = df[df['ID'] == id_termo]

            if not registro_termo.empty:
                dados_termo = registro_termo.iloc[0].to_dict()
                st.success(f"✅ Usuário: **{dados_termo.get('PROPRIETARIO', 'Não informado')}**")
                
                col1, col2, col3 = st.columns(3)
                col1.metric("Município", dados_termo.get('MUNICIPIO', '-'))
                col2.metric("Sistema", dados_termo.get('SISTEMA', '-'))
                col3.metric("Situação", dados_termo.get('SITUACAO', '-'))

                if st.button("🚀 Gerar Termo de Responsabilidade (Word)", type="primary", key="btn_termo"):
                    try:
                        doc = DocxTemplate("template_pisf.docx")
                        doc.render(dados_termo)
                        output = io.BytesIO()
                        doc.save(output)
                        output.seek(0)
                        
                        st.download_button(
                            label="📥 Baixar Termo",
                            data=output,
                            file_name=f"Termo_ID_{id_termo}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    except Exception as e:
                        st.error(f"Erro ao gerar o Termo: {e}. Verifique o arquivo 'template_pisf.docx'.")
            else:
                st.error("ID não localizado.")

    # ==========================================================
    # ABA 2: FICHA DE MATERIAIS
    # ==========================================================
    with aba_materiais:
        st.markdown("#### Emissão da Ficha de Materiais")
        st.write("Gere a lista técnica de compra parametrizada pelo tipo de sistema do usuário.")
        
        c_m1, c_m2, c_m3 = st.columns([1, 2, 1])
        with c_m2:
            id_mat = st.text_input("🔍 Digite o ID da Captação para a Ficha:", placeholder="Ex: 562026", key="input_mat")

        if id_mat:
            registro_mat = df[df['ID'] == id_mat]

            if not registro_mat.empty:
                dados_mat = registro_mat.iloc[0].to_dict()
                
                sistema = str(dados_mat.get('SISTEMA', '')).strip().upper()
                estaca = str(dados_mat.get('ESTACA', '')).strip().upper()
                
                template_ficha = None
                tipo_perfil = ""
                
                if sistema == 'GRAVIDADE':
                    template_ficha = "template_pisf-ATERRO.docx"
                    tipo_perfil = "Aterro (Sistema por Gravidade)"
                elif sistema == 'BOMBEAMENTO':
                    if 'RESERVAT' in estaca:
                        template_ficha = "template_pisf-RES.docx"
                        tipo_perfil = "Reservatório (Bombeamento em Reservatório)"
                    else:
                        template_ficha = "template_pisf-BOMBEAMENTO.docx"
                        tipo_perfil = "Bombeamento (Captação Direta no Canal)"
                
                st.success(f"✅ Usuário: **{dados_mat.get('PROPRIETARIO', 'Não informado')}**")
                
                cm1, cm2, cm3 = st.columns(3)
                cm1.metric("Estaca / Localização", dados_mat.get('ESTACA', '-'))
                cm2.metric("Sistema Lançado", dados_mat.get('SISTEMA', '-'))
                cm3.metric("Ficha Selecionada", tipo_perfil.split(' ')[0] if tipo_perfil else "Erro")

                if template_ficha:
                    st.markdown(f'<div class="info-box"><strong>⚙️ Análise do Sistema:</strong> Detectamos o perfil <strong>{tipo_perfil}</strong>. O documento será gerado usando a base <code>{template_ficha}</code>.</div>', unsafe_allow_html=True)
                    
                    if st.button("🚀 Gerar Ficha de Materiais (Word)", type="primary", key="btn_mat"):
                        try:
                            doc_mat = DocxTemplate(template_ficha)
                            doc_mat.render(dados_mat)
                            output_mat = io.BytesIO()
                            doc_mat.save(output_mat)
                            output_mat.seek(0)
                            
                            st.download_button(
                                label="📥 Baixar Ficha de Materiais",
                                data=output_mat,
                                file_name=f"Ficha_Materiais_ID_{id_mat}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                        except Exception as e:
                            st.error(f"Erro ao carregar o template. Certifique-se de que o arquivo '{template_ficha}' está na mesma pasta. Detalhe: {e}")
                else:
                    st.warning("⚠️ Atenção: Não foi possível determinar o template automaticamente. Verifique se a coluna 'SISTEMA' na planilha está preenchida corretamente.")
            else:
                st.error("ID não localizado.")

    # ==========================================================
    # ABA 3: PROJETOS DE CAPTAÇÃO (NOVO)
    # ==========================================================
    with aba_projetos:
        st.markdown("#### Projetos de Captação Padronizada")
        st.write("Baixe o arquivo PDF contendo os projetos técnicos e desenhos padronizados para as captações.")
        
        c_p1, c_p2, c_p3 = st.columns([1, 2, 1])
        with c_p2:
            st.info("📂 **Arquivo:** `projeto.pdf`\n\nEste documento contém os detalhamentos técnicos para instalação das estruturas.")
            
            # Verifica se o arquivo existe na pasta
            caminho_projeto = "projeto.pdf"
            
            if os.path.exists(caminho_projeto):
                with open(caminho_projeto, "rb") as pdf_file:
                    pdf_bytes = pdf_file.read()
                
                st.download_button(
                    label="📥 Baixar Projeto (PDF)",
                    data=pdf_bytes,
                    file_name="projeto.pdf",
                    mime="application/pdf",
                    type="primary",
                    use_container_width=True # Deixa o botão largo e bonito na coluna
                )
            else:
                st.warning("⚠️ O arquivo `projeto.pdf` ainda não foi adicionado ao sistema. Faça o upload dele para a mesma pasta do código fonte para habilitar o download.")

except Exception as e:
    st.error(f"Erro fatal na aplicação: {e}")

# 4. RODAPÉ FIXO
st.markdown("---")
st.markdown("<p style='text-align: center; color: gray;'>App por Raphael Davi</p>", unsafe_allow_html=True)
