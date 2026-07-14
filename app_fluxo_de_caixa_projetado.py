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
        # Tenta carregar a imagem da logo enviada
        logo = Image.open("logo.JPG")
        st.image(logo, width=150)
    except Exception:
        try:
            logo = Image.open("input_file_1.png")
            st.image(logo, width=150)
        except Exception:
            st.subheader("🥩 LOGO")

with col_titulo:
    st.title("Renato Frigotudo & Associados")
    st.subheader("Simulador Dinâmico de Fluxo de Caixa Projetado")

st.markdown("---")

# 2. BARRA LATERAL (Entradas Numéricas Diretas do Usuário)
st.sidebar.header("⚙️ Configurações da Simulação")

nome_cliente = st.sidebar.text_input("Nome do Cliente", value="RENATO")

arquivo_carregado = st.sidebar.file_uploader(
    "Carrega a Planilha Base (.xlsx)", 
    type=["xlsx"]
)

st.sidebar.markdown("### 📊 Digite os percentuais de simulação:")

# Alterado de Sliders para campos de digitação direta (number_input) conforme solicitado
perc_receita = st.sidebar.number_input(
    "Aumento/Diminuição da Receita (%)", 
    min_value=-100.0, 
    max_value=300.0, 
    value=3.0, 
    step=0.1,
    format="%.2f"
)

perc_lucro = st.sidebar.number_input(
    "Lucratividade Líquida Desejada (%)", 
    min_value=0.0, 
    max_value=100.0, 
    value=8.0, 
    step=0.1,
    format="%.2f"
)

# 3. PROCESSAMENTO E EXIBIÇÃO EM TEMPO REAL
if arquivo_carregado is not None:
    try:
        r_adj = perc_receita / 100.0
        l_adj = perc_lucro / 100.0

        # Abrir o arquivo em memória
        arquivo_carregado.seek(0)
        conteudo_ficheiro = arquivo_carregado.read()
        wb = openpyxl.load_workbook(io.BytesIO(conteudo_ficheiro), data_only=True)
        
        # Seleciona a primeira aba disponível (BAL_25)
        aba_nome = wb.sheetnames[0]
        sheet_bal = wb[aba_nome]

        # Índices das colunas dos meses históricos (Colunas D, G, J, M, P, S, V, Y, AB, AE, AH, AK)
        colunas_meses_indices = [4, 7, 10, 13, 16, 19, 22, 25, 28, 31, 34, 37]
        nomes_meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

        # --- PROCESSAMENTO EXATO DO HISTÓRICO ---
        receita_total_historica = 0.0
        despesa_total_historica = 0.0
        compra_mercadoria_historica = 0.0
        insumos_historicos = 0.0

        for col_idx in colunas_meses_indices:
            # Receitas Operacionais (Linha 6 na planilha original)
            val_rec = sheet_bal.cell(row=6, column=col_idx).value
            receita_total_historica += float(val_rec) if isinstance(val_rec, (int, float)) else 0.0
            
            # Despesa Total (Linha 98 na planilha original)
            val_desp = sheet_bal.cell(row=98, column=col_idx).value
            despesa_total_historica += float(val_desp) if isinstance(val_desp, (int, float)) else 0.0

            # Compra de Mercadoria (Linha 14 na planilha original)
            val_compra = sheet_bal.cell(row=14, column=col_idx).value
            compra_mercadoria_historica += float(val_compra) if isinstance(val_compra, (int, float)) else 0.0

            # Insumos (Linha 21 na planilha original)
            val_insumos = sheet_bal.cell(row=21, column=col_idx).value
            insumos_historicos += float(val_insumos) if isinstance(val_insumos, (int, float)) else 0.0

        # --- CÁLCULO DOS INDICADORES REAIS (LINHA 4) ---
        # Receita Projetada (B4)
        receita_projetada = receita_total_historica * (1 + r_adj)
        
        # Despesa Base (L4)
        despesa_base = despesa_total_historica
        
        # Despesas Variáveis (Compras + Insumos)
        despesas_variaveis_base = compra_mercadoria_historica + insumos_historicos

        # Fator de Ajuste de Despesas (J4)
        if despesas_variaveis_base == 0:
            fator_ajuste = 1.0
        else:
            numerador = (receita_projetada * (1 - l_adj)) - (despesa_base - despesas_variaveis_base)
            fator_ajuste = numerador / despesas_variaveis_base
            fator_ajuste = max(0.0, min(1.0, fator_ajuste))

        # Despesa Projetada (D4)
        despesas_fixas = despesa_base - despesas_variaveis_base
        despesa_projetada = despesas_fixas + (despesas_variaveis_base * fator_ajuste)

        # Resultado Projetado (F4)
        resultado_projetado = receita_projetada - despesa_projetada

        # Margem Projetada (H4)
        margem_projetada = (resultado_projetado / receita_projetada) if receita_projetada != 0 else 0.0

        # --- APRESENTAÇÃO DOS INDICADORES NA TELA ---
        st.markdown("### 📊 Indicadores Financeiros do Cenário (Equivalente à Linha 4)")
        
        col_r1, col_r2, col_r3, col_r4, col_r5 = st.columns(5)
        with col_r1:
            st.metric("Receita Projetada (B4)", f"R$ {receita_projetada:,.2f}")
        with col_r2:
            st.metric("Despesa Projetada (D4)", f"R$ {despesa_projetada:,.2f}")
        with col_r3:
            st.metric("Resultado Projetado (F4)", f"R$ {resultado_projetado:,.2f}")
        with col_r4:
            st.metric("Margem Projetada (H4)", f"{margem_projetada * 100:.2f}%")
        with col_r5:
            st.metric("Fator de Ajuste (J4)", f"{fator_ajuste * 100:.2f}%")

        # --- CONSTRUÇÃO DE TABELA INTERATIVA EM TELA ---
        st.markdown("### 📈 Visualização Dinâmica da Planilha FC_PROJETADO")
        
        # Varre as linhas da planilha base para montar a visualização idêntica na tela
        linhas_visualizacao = []
        for r in range(4, sheet_bal.max_row + 1):
            conta = sheet_bal.cell(row=r, column=3).value
            tipo = sheet_bal.cell(row=r, column=2).value
            
            if conta is not None and str(conta).strip() != "":
                valores_originais = []
                for col_idx in colunas_meses_indices:
                    val = sheet_bal.cell(row=r, column=col_idx).value
                    valores_originais.append(float(val) if isinstance(val, (int, float)) else 0.0)
                
                # Aplicação dinâmica das regras de cálculo
                if conta == "RECEITAS OPERACIONAIS" or conta == "1.1 Venda Consumidor Final":
                    valores_projetados = [v * (1 + r_adj) for v in valores_originais]
                elif conta in ["3.1 Compra de Mercadoria", "3.8 Insumos"]:
                    valores_projetados = [v * fator_ajuste for v in valores_originais]
                else:
                    valores_projetados = valores_originais
                
                total_projetado = sum(valores_projetados)
                
                linhas_visualizacao.append({
                    "Conta / Hierarquia": str(conta).strip(),
                    "Tipo": str(tipo).strip() if tipo else "",
                    "Total Projetado": total_projetado,
                    **{nomes_meses[i]: valores_projetados[i] for i in range(12)}
                })
                
        df_visual = pd.DataFrame(linhas_visualizacao)
        
        # Formata o DataFrame para exibição limpa com moeda e duas casas decimais
        df_formatado = df_visual.copy()
        colunas_formatar = ["Total Projetado"] + nomes_meses
        for col in colunas_formatar:
            df_formatado[col] = df_formatado[col].apply(lambda x: f"R$ {x:,.2f}" if isinstance(x, (int, float)) else x)
            
        # Exibe a planilha interativa diretamente na tela com rolagem e busca
        st.dataframe(df_formatado, use_container_width=True, height=450)

        # --- PREPARAÇÃO DO EXCEL PARA EXPORTAÇÃO ---
        # Abre o arquivo com fórmulas ativas para o download final
        wb_formulas = openpyxl.load_workbook(io.BytesIO(conteudo_ficheiro), data_only=False)
        sheet_bal_form = wb_formulas[aba_nome]
        
        if 'FC_PROJETADO' in wb_formulas.sheetnames:
            del wb_formulas['FC_PROJETADO']
        sheet_proj = wb_formulas.create_sheet(title="FC_PROJETADO")
        
        # Duplica estrutura para manter design da exportação
        for r in range(1, sheet_bal_form.max_row + 1):
            for c in range(1, sheet_bal_form.max_column + 1):
                cell = sheet_bal_form.cell(row=r, column=c)
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

        # Injeta variáveis digitadas nas posições padrão
        sheet_proj['B2'] = r_adj
        sheet_proj['B3'] = l_adj
        sheet_proj['A1'] = f"CLIENTE: {nome_cliente.upper()}"
        sheet_proj['J4'] = (
            f"=IF(SUM(BAL_25!D14,BAL_25!G14,BAL_25!J14,BAL_25!M14,BAL_25!P14,BAL_25!S14,BAL_25!V14,BAL_25!Y14,BAL_25!AB14,BAL_25!AE14,BAL_25!AH14,BAL_25!AK14,BAL_25!D21,BAL_25!G21,BAL_25!J21,BAL_25!M21,BAL_25!P21,BAL_25!S21,BAL_25!V21,BAL_25!Y21,BAL_25!AB21,BAL_25!AE21,BAL_25!AH21,BAL_25!AK21)=0,1,"
            f"MAX(0,MIN(1,(B4*(1-B3)-(L4-SUM(BAL_25!D14,BAL_25!G14,BAL_25!J14,BAL_25!M14,BAL_25!P14,BAL_25!S14,BAL_25!V14,BAL_25!Y14,BAL_25!AB14,BAL_25!AE14,BAL_25!AH14,BAL_25!AK14,BAL_25!D21,BAL_25!G21,BAL_25!J21,BAL_25!M21,BAL_25!P21,BAL_25!S21,BAL_25!V21,BAL_25!Y21,BAL_25!AB21,BAL_25!AE21,BAL_25!AH21,BAL_25!AK21)))"
            f"/SUM(BAL_25!D14,BAL_25!G14,BAL_25!J14,BAL_25!M14,BAL_25!P14,BAL_25!S14,BAL_25!V14,BAL_25!Y14,BAL_25!AB14,BAL_25!AE14,BAL_25!AH14,BAL_25!AK14,BAL_25!D21,BAL_25!G21,BAL_25!J21,BAL_25!M21,BAL_25!P21,BAL_25!S21,BAL_25!V21,BAL_25!Y21,BAL_25!AB21,BAL_25!AE21,BAL_25!AH21,BAL_25!AK21))))"
        )

        # Modifica dinamicamente as receitas e despesas com as novas metas para exportação
        for r in range(1, sheet_proj.max_row + 1):
            cell_conta = sheet_bal_form.cell(row=r, column=3).value
            if cell_conta == "1.1 Venda Consumidor Final":
                for col_idx in colunas_meses_indices:
                    val_hist = sheet_bal_form.cell(row=r, column=col_idx).value
                    if isinstance(val_hist, (int, float)):
                        sheet_proj.cell(row=r, column=col_idx, value=val_hist * (1 + r_adj))
            elif cell_conta in ["3.1 Compra de Mercadoria", "3.8 Insumos"]:
                for col_idx in colunas_meses_indices:
                    val_hist = sheet_bal_form.cell(row=r, column=col_idx).value
                    if isinstance(val_hist, (int, float)):
                        sheet_proj.cell(row=r, column=col_idx, value=f"={sheet_bal_form.title}!{openpyxl.utils.get_column_letter(col_idx)}{r}*$J$4")

        # Salvar em formato binário para download
        output_web = io.BytesIO()
        wb_formulas.save(output_web)
        output_web.seek(0)

        st.markdown("---")
        st.markdown("### 📥 Exportar Resultados")
        st.download_button(
            label="📥 Exportar Planilha com FC_PROJETADO Oficial",
            data=output_web,
            file_name=f"FC_PROJETADO_{nome_cliente.upper()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro no processamento matemático: {e}")
else:
    st.info("ℹ️ Insira os dados de simulação e carregue a planilha base para ver os resultados correspondentes à linha 4.")
