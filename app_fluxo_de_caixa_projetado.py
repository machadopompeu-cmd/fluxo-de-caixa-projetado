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

# 3. PROCESSAMENTO E EXIBIÇÃO DINÂMICA
if arquivo_carregado is not None:
    try:
        r_adj = perc_receita / 100.0
        l_adj = perc_lucro / 100.0

        # Carregar planilha para o Pandas para calcular e exibir de forma dinâmica na tela
        xls = pd.ExcelFile(arquivo_carregado)
        aba_nome = xls.sheet_names[0]
        
        # Lemos a tabela ignorando as linhas vazias do topo para processar os números
        df_dados = pd.read_excel(arquivo_carregado, sheet_name=aba_nome, skiprows=2)
        
        # Limpar linhas e colunas completamente nulas
        df_dados = df_dados.dropna(subset=[df_dados.columns[2]])
        
        # --- MOTOR DE CÁLCULO DINÂMICO EM PYTHON ---
        col_conta = df_dados.columns[2]
        col_tipo = df_dados.columns[1] if len(df_dados.columns) > 1 else None
        
        # Identificar as colunas de meses (colunas que contêm os valores financeiros)
        # No formato do arquivo, as colunas de dados começam a partir da quarta coluna (índice 3) de 3 em 3
        colunas_meses_indices = [3, 6, 9, 12, 15, 18, 21, 24, 27, 30, 33, 36]
        nomes_meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
        
        # Criamos uma estrutura para guardar os dados calculados do FC_PROJETADO
        linhas_projetadas = []
        
        # 1. Passo: Calcular Receita Projetada Total
        receita_total_projetada = 0.0
        despesa_variavel_base = 0.0
        despesa_fixa_base = 0.0
        
        # Vamos primeiro varrer os dados para calcular os agregados e o fator de ajuste
        for idx, row in df_dados.iterrows():
            conta = str(row[col_conta]).strip()
            
            # Identifica valores mensais válidos
            valores_meses = []
            for col_idx in colunas_meses_indices:
                if col_idx < len(row):
                    val = row.iloc[col_idx]
                    valores_meses.append(float(val) if pd.notna(val) and isinstance(val, (int, float)) else 0.0)
                else:
                    valores_meses.append(0.0)
            
            total_linha = sum(valores_meses)
            
            if conta == "1.1 Venda Consumidor Final":
                receita_total_projetada = total_linha * (1 + r_adj)
            elif conta in ["3.1 Compra de Mercadoria", "3.8 Insumos"]:
                despesa_variavel_base += total_linha
            elif "Despesa" in str(row.get(col_tipo, "")) and conta not in ["3.1 Compra de Mercadoria", "3.8 Insumos"]:
                despesa_fixa_base += total_linha

        # 2. Passo: Calcular o Fator de Ajuste das Despesas Variáveis (J4)
        # Fórmula: Fator = (ReceitaProjetada * (1 - LucratividadeDesejada) - DespesaFixa) / DespesaVariavel
        if despesa_variavel_base > 0:
            numerador = (receita_total_projetada * (1 - l_adj)) - despesa_fixa_base
            fator_ajuste = numerador / despesa_variavel_base
            fator_ajuste = max(0.0, min(1.0, fator_ajuste)) # Limita entre 0 e 100%
        else:
            fator_ajuste = 1.0

        # 3. Passo: Montar a tabela visual para exibição na tela
        for idx, row in df_dados.iterrows():
            conta = str(row[col_conta]).strip()
            tipo = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
            
            valores_originais = []
            for col_idx in colunas_meses_indices:
                if col_idx < len(row):
                    val = row.iloc[col_idx]
                    valores_originais.append(float(val) if pd.notna(val) and isinstance(val, (int, float)) else 0.0)
                else:
                    valores_originais.append(0.0)
            
            # Aplicar as regras de projeção para cada linha
            valores_projetados = []
            if conta == "1.1 Venda Consumidor Final" or conta == "RECEITAS OPERACIONAIS":
                valores_projetados = [v * (1 + r_adj) for v in valores_originais]
            elif conta in ["3.1 Compra de Mercadoria", "3.8 Insumos"]:
                valores_projetados = [v * fator_ajuste for v in valores_originais]
            else:
                # Despesas fixas e saldos não sofrem alteração pelo fator de ajuste
                valores_projetados = valores_originais
                
            total_projetado = sum(valores_projetados)
            
            # Adiciona ao nosso relatório visual
            linhas_projetadas.append({
                "Conta / Hierarquia": conta,
                "Tipo": tipo,
                "Total Projetado": total_projetado,
                **{nomes_meses[i]: valores_projetados[i] for i in range(12)}
            })
            
        df_relatorio = pd.DataFrame(linhas_projetadas)

        # Interface do Usuário com Visualização Dinâmica
        st.success(f"🎉 Simulação gerada com sucesso para o cliente **{nome_cliente.upper()}**!")
        
        # Painel de Métricas Principais baseadas no cenário simulado
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("Receita Total Projetada", f"R$ {receita_total_projetada:,.2f}")
        col_m2.metric("Meta de Lucratividade", f"{perc_lucro}%")
        col_m3.metric("Fator de Ajuste de Compras", f"{fator_ajuste * 100:.1f}%")

        # --- EXIBIÇÃO DA PLANILHA EM TELA ---
        st.markdown("### 📋 Visualização da Planilha FC_PROJETADO")
        st.markdown("Esta tabela abaixo atualiza-se **imediatamente** sempre que moves os seletores na barra lateral!")
        
        # Formatando os valores para exibição como moeda brasileira
        df_visual = df_relatorio.copy()
        colunas_valores = ["Total Projetado"] + nomes_meses
        for col in colunas_valores:
            df_visual[col] = df_visual[col].apply(lambda x: f"R$ {x:,.2f}" if isinstance(x, (int, float)) else x)
            
        st.dataframe(df_visual, use_container_width=True, height=500)

        # --- EXPORTAÇÃO EXCEL COM FÓRMULAS ---
        # Recarrega para gerar o arquivo Excel físico contendo as fórmulas originais para o usuário
        arquivo_carregado.seek(0)
        wb = openpyxl.load_workbook(io.BytesIO(arquivo_carregado.read()))
        sheet_origem = wb[aba_nome]
        
        if 'FC_PROJETADO' in wb.sheetnames:
            del wb['FC_PROJETADO']
        sheet_proj = wb.create_sheet(title="FC_PROJETADO")
        
        # Copia dados e estilos
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

        # Injeta os parâmetros nas células do Excel para que o arquivo físico funcione
        sheet_proj['B2'] = r_adj
        sheet_proj['B3'] = l_adj
        sheet_proj['A1'] = f"CLIENTE: {nome_cliente.upper()}"
        sheet_proj['J4'] = (
            f"=IF(SUM(D14,G14,J14,M14,P14,S14,V14,Y14,AB14,AE14,AH14,AK14,D21,G21,J21,M21,P21,S21,V21,Y21,AB21,AE21,AH21,AK21)=0,1,"
            f"MAX(0,MIN(1,(B4*(1-B3)-(L4-SUM(D14,G14,J14,M14,P14,S14,V14,Y14,AB14,AE14,AH14,AK14,D21,G21,J21,M21,P21,S21,V21,Y21,AB21,AE21,AH21,AK21)))"
            f"/SUM(D14,G14,J14,M14,P14,S14,V14,Y14,AB14,AE14,AH14,AK14,D21,G21,J21,M21,P21,S21,V21,Y21,AB21,AE21,AH21,AK21))))"
        )

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

        output_web = io.BytesIO()
        wb.save(output_web)
        output_web.seek(0)

        st.markdown("---")
        st.markdown("### 📥 Exportar Resultados")
        st.download_button(
            label="📥 Exportar Planilha com FC_PROJETADO",
            data=output_web,
            file_name=f"FC_PROJETADO_{nome_cliente.upper()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Ocorreu um erro no cálculo dinâmico: {e}")
else:
    st.info("ℹ️ Aguardando carregamento do arquivo Excel base para iniciar a simulação dinâmica.")
