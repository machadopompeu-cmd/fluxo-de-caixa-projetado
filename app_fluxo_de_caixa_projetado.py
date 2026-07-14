import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
import io
from PIL import Image

# 1. Configuração da Página Web
st.set_page_config(
    page_title="Renato Frigotudo & Associados",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CABEÇALHO PERSONALIZADO ---
col_logo, col_titulo = st.columns([1, 4])

with col_logo:
    try:
        # Carrega a imagem enviada (logo.JPG)
        logo = Image.open("logo.JPG")
        st.image(logo, width=150)
    except Exception:
        st.subheader("🥩 LOGO")

with col_titulo:
    st.title("Renato Frigotudo & Associados")
    st.subheader("Simulador Dinâmico de Fluxo de Caixa Projetado")

st.markdown("---")

# 2. BARRA LATERAL (Entradas do Usuário)
st.sidebar.header("⚙️ Configurações da Simulação")

nome_cliente = st.sidebar.text_input("Nome do Cliente", value="RENATO")

arquivo_carregado = st.sidebar.file_uploader(
    "Carrega a Planilha Base (.xlsx)", 
    type=["xlsx"]
)

# Ajuste de Receita e Lucratividade
perc_receita = st.sidebar.number_input(
    "Ajuste de Receita (%)", 
    min_value=-100.0, 
    max_value=300.0, 
    value=3.0, 
    step=0.5
)

perc_lucro = st.sidebar.number_input(
    "Lucratividade Desejada (%)", 
    min_value=0.0, 
    max_value=100.0, 
    value=8.0, 
    step=0.5
)

# 3. COMPUTAÇÃO E SIMULAÇÃO DINÂMICA
if arquivo_carregado is not None:
    try:
        r_adj = perc_receita / 100.0
        l_adj = perc_lucro / 100.0

        # Carrega a planilha base
        xls = pd.ExcelFile(arquivo_carregado)
        aba_nome = xls.sheet_names[0]
        
        # Carregamento com OpenPyXL para podermos calcular fórmulas e salvar
        arquivo_carregado.seek(0)
        wb = openpyxl.load_workbook(io.BytesIO(arquivo_carregado.read()))
        sheet_origem = wb[aba_nome]
        
        # Cria ou limpa a aba FC_PROJETADO
        if 'FC_PROJETADO' in wb.sheetnames:
            del wb['FC_PROJETADO']
        sheet_proj = wb.create_sheet(title="FC_PROJETADO")
        
        # Copia todos os dados e estilos da aba original
        for r in range(1, sheet_origem.max_row + 1):
            for c in range(1, sheet_origem.max_column + 1):
                cell = sheet_origem.cell(row=r, column=c)
                new_cell = sheet_proj.cell(row=r, column=c, value=cell.value)
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

        # Injeta os valores dos controles nas células de controle
        sheet_proj['B2'] = r_adj
        sheet_proj['B3'] = l_adj
        sheet_proj['A1'] = f"CLIENTE: {nome_cliente.upper()}"
        
        # Fator de Ajuste dinâmico (Fórmula baseada na planilha original)
        formula_fator = (
            f"=IF(SUM(D14,G14,J14,M14,P14,S14,V14,Y14,AB14,AE14,AH14,AK14,D21,G21,J21,M21,P21,S21,V21,Y21,AB21,AE21,AH21,AK21)=0,1,"
            f"MAX(0,MIN(1,(B4*(1-B3)-(L4-SUM(D14,G14,J14,M14,P14,S14,V14,Y14,AB14,AE14,AH14,AK14,D21,G21,J21,M21,P21,S21,V21,Y21,AB21,AE21,AH21,AK21)))"
            f"/SUM(D14,G14,J14,M14,P14,S14,V14,Y14,AB14,AE14,AH14,AK14,D21,G21,J21,M21,P21,S21,V21,Y21,AB21,AE21,AH21,AK21))))"
        )
        sheet_proj['J4'] = formula_fator

        # Modifica as células de receita e custo dinamicamente
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

        # Salva o arquivo em memória para download
        output_web = io.BytesIO()
        wb.save(output_web)
        output_web.seek(0)

        # Mostra o painel com as novas projeções
        st.success(f"🎉 Simulação gerada com sucesso para o cliente **{nome_cliente.upper()}**!")
        
        # Métricas na Tela
        col1, col2 = st.columns(2)
        col1.metric("Projeção de Receita (Vendas)", f"+ {perc_receita}%")
        col2.metric("Margem de Lucro Alvo", f"{perc_lucro}%")

        st.markdown("### 📥 Salvar o cenário atual")
        st.download_button(
            label="📥 Exportar Excel com aba 'FC_PROJETADO'",
            data=output_web,
            file_name=f"FC_PROJETADO_{nome_cliente.upper()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Ocorreu um erro no cálculo dinâmico: {e}")
else:
    st.info("ℹ️ Aguardando carregamento do arquivo Excel base para iniciar a simulação dinâmica.")
