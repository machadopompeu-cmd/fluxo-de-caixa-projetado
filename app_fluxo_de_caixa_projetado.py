import streamlit as st
import openpyxl
import pandas as pd
import io

# 1. Configuração Inicial da Página Web (Título e Layout Largo)
st.set_page_config(
    page_title="Fluxo de Caixa Projetado - Açougue & Empório",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📊 Fluxo de Caixa Projetado")
st.markdown("""
Este aplicativo web inteligente ajuda-te a projetar receitas, despesas e a calcular a lucratividade líquida 
do teu comércio de carnes ou empório com base no teu histórico financeiro, considerando datas sazonais e feriados.
""")

# 2. BARRA LATERAL (Configurações e Ajustes de Parâmetros)
st.sidebar.header("⚙️ Parâmetros de Projeção")

# Campo para o usuário fazer o upload da planilha
arquivo_carregado = st.sidebar.file_uploader(
    "Carrega a tua Planilha Excel (.xlsx)", 
    type=["xlsx"]
)

# Sliders interativos para substituir caixas de texto chatas
perc_receita = st.sidebar.slider(
    "Aumento/Diminuição da Receita (%)", 
    min_value=-50.0, 
    max_value=100.0, 
    value=3.0, 
    step=0.5
)

perc_lucro = st.sidebar.slider(
    "Lucratividade Desejada (%)", 
    min_value=0.0, 
    max_value=50.0, 
    value=8.0, 
    step=0.5
)

# 3. PROCESSAMENTO DOS DADOS (Quando o arquivo é carregado)
if arquivo_carregado is not None:
    try:
        # Lê os bytes do arquivo enviado pelo usuário
        conteudo_bytes = arquivo_carregado.read()
        
        # Abre o arquivo com o OpenPyXL para manter as fórmulas do Excel intactas
        wb = openpyxl.load_workbook(io.BytesIO(conteudo_bytes))
        
        # Verifica se a aba correta existe no arquivo
        if 'FC_PROJETADO' in wb.sheetnames:
            sheet = wb['FC_PROJETADO']
            
            # Insere os dados dos Sliders diretamente nas células B2 e B3
            # Dividimos por 100 pois o Excel trabalha com formato decimal para percentagens (ex: 0.03 = 3%)
            sheet['B2'] = perc_receita / 100.0
            sheet['B3'] = perc_lucro / 100.0
            
            # Salva o arquivo atualizado temporariamente na memória RAM do servidor
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            
            # Exibe confirmação visual e os indicadores principais na tela
            st.success("✅ Planilha processada e parâmetros atualizados com sucesso!")
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric(label="Aumento de Receita Aplicado", value=f"{perc_receita}%")
            with col2:
                st.metric(label="Meta de Lucratividade Líquida", value=f"{perc_lucro}%")
            
            st.markdown("### 📥 Obter o arquivo projetado")
            st.info("O teu arquivo está pronto! Clica no botão abaixo para descarregar a planilha com todos os cálculos, fórmulas e sazonalidades recalculados.")
            
            # Botão para o usuário descarregar o Excel pronto
            st.download_button(
                label="📥 Descarregar Planilha Projetada (Excel)",
                data=output,
                file_name="FLUXO_DE_CAIXA_PROJETADO_ATUALIZADO.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Erro: A aba 'FC_PROJETADO' não foi encontrada neste arquivo Excel.")
            
    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o arquivo: {e}")
else:
    # Mensagem informativa enquanto o usuário não faz o upload
    st.info("ℹ️ Aguardando o carregamento da tua planilha Excel (.xlsx) na barra lateral para iniciar a projeção.")