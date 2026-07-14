import streamlit as st
import pandas as pd
import openpyxl
import io
from PIL import Image

# 1. Configuração Inicial da Página
st.set_page_config(
    page_title="Renato Frigotudo & Associados - Fluxo de Caixa",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CABEÇALHO PERSONALIZADO ---
# Criamos duas colunas: uma pequena para a Logo no canto esquerdo e outra para o Título
col_logo, col_titulo = st.columns([1, 4])

with col_logo:
    try:
        # Carrega a imagem da logo (utilizando a imagem disponibilizada no sistema)
        logo = Image.open("input_file_1.png")
        st.image(logo, width=150)
    except Exception:
        # Caso a imagem não seja encontrada localmente, exibe um espaço reservado
        st.subheader("🥩 LOGO")

with col_titulo:
    st.title("Renato Frigotudo & Associados")
    st.subheader("Sistema Inteligente de Fluxo de Caixa Projetado")

st.markdown("---")

# 2. BARRA LATERAL (Configurações e Ajustes)
st.sidebar.header("⚙️ Identificação e Parâmetros")

# Novo campo para inserção do nome do cliente
nome_cliente = st.sidebar.text_input(
    "Nome do Cliente",
    value="RENATO"
)

arquivo_carregado = st.sidebar.file_uploader(
    "Carrega a Planilha Histórica (.xlsx)", 
    type=["xlsx"]
)

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

# 3. PROCESSAMENTO E EXIBIÇÃO
if arquivo_carregado is not None:
    try:
        # Mensagem de boas-vindas personalizada ao cliente ativo
        st.info(f"👤 **Cliente Ativo:** {nome_cliente.upper()}")
        
        # Leitura da planilha base
        xls = pd.ExcelFile(arquivo_carregado)
        aba_nome = xls.sheet_names[0]
        
        # Carregamos com openpyxl na memória para reescrever os dados sem alterar o original
        arquivo_carregado.seek(0)
        wb = openpyxl.load_workbook(io.BytesIO(arquivo_carregado.read()))
        sheet_origem = wb[aba_nome]
        
        # Cria a nova aba FC_PROJETADO limpa
        if 'FC_PROJETADO' in wb.sheetnames:
            del wb['FC_PROJETADO']
        sheet_proj = wb.create_sheet(title="FC_PROJETADO")
        
        # Clona o conteúdo da aba base mantendo as cores e estilos
        for row in sheet_origem.iter_rows():
            for cell in row:
                new_cell = sheet_proj.cell(row=cell.row, column=cell.column, value=cell.value)
                if cell.has_style:
                    new_cell.font = openpyxl.styles.Font(
                        name=cell.font.name, size=cell.font.size, bold=cell.font.bold, 
                        italic=cell.font.italic, color=cell.font.color
                    )
                    new_cell.fill = openpyxl.styles.PatternFill(
                        fill_type=cell.fill.fill_type, start_color=cell.fill.start_color, end_color=cell.fill.end_color
                    )
                    new_cell.alignment = openpyxl.styles.Alignment(
                        horizontal=cell.alignment.horizontal, vertical=cell.alignment.vertical
                    )
                    new_cell.border = openpyxl.styles.Border(
                        left=cell.border.left, right=cell.border.right, top=cell.border.top, bottom=cell.border.bottom
                    )
                    new_cell.number_format = cell.number_format

        # Salva o nome do cliente diretamente no cabeçalho interno da planilha gerada
        sheet_proj['A1'] = f"CLIENTE: {nome_cliente.upper()}"
        sheet_proj['A1'].font = openpyxl.styles.Font(name="Calibri", size=12, bold=True)
        
        # Define os parâmetros de simulação nas células do Excel
        r_adj = perc_receita / 100.0
        l_adj = perc_lucro / 100.0
        sheet_proj['B2'] = r_adj
        sheet_proj['B3'] = l_adj

        # Fórmula matemática do Fator de Ajuste de Despesas
        formula_fator_ajuste = (
            f"=IF(SUM(D14,G14,J14,M14,P14,S14,V14,Y14,AB14,AE14,AH14,AK14,D21,G21,J21,M21,P21,S21,V21,Y21,AB21,AE21,AH21,AK21)=0,1,"
            f"MAX(0,MIN(1,(B4*(1-B3)-(L4-SUM(D14,G14,J14,M14,P14,S14,V14,Y14,AB14,AE14,AH14,AK14,D21,G21,J21,M21,P21,S21,V21,Y21,AB21,AE21,AH21,AK21)))"
            f"/SUM(D14,G14,J14,M14,P14,S14,V14,Y14,AB14,AE14,AH14,AK14,D21,G21,J21,M21,P21,S21,V21,Y21,AB21,AE21,AH21,AK21))))"
        )
        sheet_proj['J4'] = formula_fator_ajuste

        # Processa as linhas do fluxo alterando as receitas e ajustando compras/insumos
        for r in range(1, sheet_proj.max_row + 1):
            cell_conta = sheet_proj.cell(row=r, column=3).value
            if cell_conta == "1.1 Venda Consumidor Final":
                for col_idx in [4, 7, 10, 13, 16, 19, 22, 25, 28, 31, 34, 37]:
                    val_hist = sheet_origem.cell(row=r, column=col_idx).value
                    if isinstance(val_hist, (int, float)):
                        sheet_proj.cell(row=r, column=col_idx, value=val_hist * (1 + r_adj))

            elif cell_conta in ["3.1 Compra de Mercadoria", "3.8 Insumos"]:
                for col_idx in [4, 7, 10, 13, 16, 19, 22, 25, 28, 31, 34, 37]:
                    val_hist = sheet_origem.cell(row=r, column=col_idx).value
                    if isinstance(val_hist, (int, float)):
                        sheet_proj.cell(row=r, column=col_idx, value=f"={sheet_origem.title}!{openpyxl.utils.get_column_letter(col_idx)}{r}*$J$4")

        # Conversão final para download em memória
        output_web = io.BytesIO()
        wb.save(output_web)
        output_web.seek(0)

        # Dashboard de Visualização das Métricas
        st.markdown(f"### 📈 Projeção Atual para {nome_cliente.upper()}")
        
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.metric("Meta de Aumento de Receitas", f"{perc_receita}%")
        with col_m2:
            st.metric("Alvo de Lucratividade Desejada", f"{perc_lucro}%")

        st.markdown("---")
        st.markdown("### 📥 Exportação de Dados")
        st.download_button(
            label="📥 Descarregar Planilha com FC_PROJETADO",
            data=output_web,
            file_name=f"FC_PROJETADO_{nome_cliente.upper()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Ocorreu um erro no processamento: {e}")
else:
    st.info("ℹ️ Insira o nome do cliente e carregue a sua planilha base para iniciar.")
