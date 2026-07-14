import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
import io
import os
from PIL import Image

# 1. Configuração da Página Web
st.set_page_config(
    page_title="Renato Frigotudo & Associados",
    layout="wide",
    initial_sidebar_state="expanded"
)

ARQUIVO_TEMPLATE = "GOIAS NOVO DFC PROJETADO-SEM CORTE.xlsx"

# --- CABEÇALHO PERSONALIZADO ---
col_logo, col_titulo = st.columns([1, 4])

with col_logo:
    try:
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

# 2. BARRA LATERAL (Campos Digitados de Precisão)
st.sidebar.header("⚙️ Configurações da Simulação")

nome_cliente = st.sidebar.text_input("Nome do Cliente", value="RENATO")

arquivo_carregado = st.sidebar.file_uploader(
    "Carrega a Planilha de Dados Históricos (.xlsx)", 
    type=["xlsx"]
)

st.sidebar.markdown("### 📊 Digite os percentuais de simulação:")

# Campos de digitação manual direta para precisão
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

# Função para localizar dinamicamente a linha pelo nome da conta
def localizar_linha(sheet, nome_conta):
    for r in range(1, sheet.max_row + 1):
        val = sheet.cell(row=r, column=3).value
        if val and str(val).strip().upper() == nome_conta.strip().upper():
            return r
    return None

# 3. COMPUTAÇÃO E SIMULAÇÃO DINÂMICA
if arquivo_carregado is not None:
    if not os.path.exists(ARQUIVO_TEMPLATE):
        st.error(f"Erro: O arquivo modelo '{ARQUIVO_TEMPLATE}' não foi encontrado no servidor.")
    else:
        try:
            r_adj = perc_receita / 100.0
            l_adj = perc_lucro / 100.0

            # 1. Carregar os dados históricos enviados pelo usuário (ex: TESTE.xlsx)
            arquivo_carregado.seek(0)
            wb_usuario = openpyxl.load_workbook(io.BytesIO(arquivo_carregado.read()), data_only=True)
            aba_usuario_nome = wb_usuario.sheetnames[0]
            sheet_usuario = wb_usuario[aba_usuario_nome]

            # 2. Carregar o arquivo de template (com fórmulas completas de FC_PROJETADO)
            wb_template = openpyxl.load_workbook(ARQUIVO_TEMPLATE, data_only=False)
            sheet_bal_template = wb_template['BAL_25']
            sheet_fc_template = wb_template['FC_PROJETADO']

            # Copiar os dados históricos do usuário para a aba BAL_25 do template
            for r in range(1, sheet_usuario.max_row + 1):
                for c in range(1, sheet_usuario.max_column + 1):
                    val = sheet_usuario.cell(row=r, column=c).value
                    if val is not None:
                        sheet_bal_template.cell(row=r, column=c, value=val)

            # Injetar os parâmetros configurados pelo utilizador no Excel físico
            sheet_fc_template['B2'] = r_adj
            sheet_fc_template['B3'] = l_adj
            sheet_fc_template['A1'] = f"CLIENTE: {nome_cliente.upper()}"

            # 3. LOCALIZAÇÃO DINÂMICA DAS LINHAS CHAVE
            linha_receitas_operacionais = localizar_linha(sheet_bal_template, "RECEITAS OPERACIONAIS")
            linha_venda_consumidor = localizar_linha(sheet_bal_template, "1.1 Venda Consumidor Final")
            linha_compra_mercadoria = localizar_linha(sheet_bal_template, "3.1 Compra de Mercadoria")
            linha_insumos = localizar_linha(sheet_bal_template, "3.8 Insumos")
            linha_despesa_total = localizar_linha(sheet_bal_template, "DESPESA TOTAL")

            # Verificação de segurança
            if not all([linha_receitas_operacionais, linha_compra_mercadoria, linha_insumos, linha_despesa_total]):
                st.error("Erro: Não foi possível mapear todas as contas-chave necessárias na sua planilha. Verifique os nomes das contas na Coluna C.")
            else:
                colunas_meses = [4, 7, 10, 13, 16, 19, 22, 25, 28, 31, 34, 37] # D, G, J, M, P, S, V, Y, AB, AE, AH, AK
                nomes_meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

                # --- PROCESSAMENTO DOS VALORES HISTÓRICOS ---
                receita_total_hist = sum(float(sheet_bal_template.cell(row=linha_receitas_operacionais, column=col).value or 0) for col in colunas_meses)
                despesa_total_hist = sum(float(sheet_bal_template.cell(row=linha_despesa_total, column=col).value or 0) for col in colunas_meses)
                compra_mercadoria_hist = sum(float(sheet_bal_template.cell(row=linha_compra_mercadoria, column=col).value or 0) for col in colunas_meses)
                insumos_hist = sum(float(sheet_bal_template.cell(row=linha_insumos, column=col).value or 0) for col in colunas_meses)

                # --- CÁLCULO DOS INDICADORES E FATOR DE AJUSTE (J4) ---
                receita_projetada = receita_total_hist * (1 + r_adj)
                despesas_variaveis_base = compra_mercadoria_hist + insumos_hist
                
                if despesas_variaveis_base == 0:
                    fator_ajuste = 1.0
                else:
                    numerador = (receita_projetada * (1 - l_adj)) - (despesa_total_hist - despesas_variaveis_base)
                    fator_ajuste = max(0.0, min(1.0, numerador / despesas_variaveis_base))

                despesas_fixas = despesa_total_hist - despesas_variaveis_base
                despesa_projetada = despesas_fixas + (despesas_variaveis_base * fator_ajuste)
                resultado_projetado = receita_projetada - despesa_projetada
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

                # --- CONSTRUÇÃO DA TABELA DINÂMICA INTERATIVA ---
                st.markdown("### 📈 Visualização Dinâmica da Planilha FC_PROJETADO")
                
                linhas_visualizacao = []
                for r in range(4, sheet_bal_template.max_row + 1):
                    conta = sheet_bal_template.cell(row=r, column=3).value
                    tipo = sheet_bal_template.cell(row=r, column=2).value
                    
                    if conta is not None and str(conta).strip() != "":
                        valores_originais = []
                        for col_idx in colunas_meses:
                            val = sheet_bal_template.cell(row=r, column=col_idx).value
                            valores_originais.append(float(val) if isinstance(val, (int, float)) else 0.0)
                        
                        # Aplicação dinâmica das regras baseada no mapeamento encontrado
                        if r == linha_receitas_operacionais or r == linha_venda_consumidor:
                            valores_projetados = [v * (1 + r_adj) for v in valores_originais]
                        elif r in [linha_compra_mercadoria, linha_insumos]:
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
                
                # Formatação de visualização de moeda em tela
                df_formatado = df_visual.copy()
                for col in ["Total Projetado"] + nomes_meses:
                    df_formatado[col] = df_formatado[col].apply(lambda x: f"R$ {x:,.2f}" if isinstance(x, (int, float)) else x)
                    
                st.dataframe(df_formatado, use_container_width=True, height=450)

                # --- EXCEL FÍSICO COM FÓRMULAS ---
                # Atualiza as receitas e despesas com as novas metas para exportação
                for r in range(1, sheet_fc_template.max_row + 1):
                    cell_conta = sheet_bal_template.cell(row=r, column=3).value
                    if cell_conta == "1.1 Venda Consumidor Final":
                        for col_idx in colunas_meses:
                            val_hist = sheet_bal_template.cell(row=r, column=col_idx).value
                            if isinstance(val_hist, (int, float)):
                                sheet_fc_template.cell(row=r, column=col_idx, value=val_hist * (1 + r_adj))
                    elif cell_conta in ["3.1 Compra de Mercadoria", "3.8 Insumos"]:
                        for col_idx in colunas_meses:
                            val_hist = sheet_bal_template.cell(row=r, column=col_idx).value
                            if isinstance(val_hist, (int, float)):
                                sheet_fc_template.cell(row=r, column=col_idx, value=f"={sheet_bal_template.title}!{openpyxl.utils.get_column_letter(col_idx)}{r}*$J$4")

                # Salvar em formato binário para download
                output_web = io.BytesIO()
                wb_template.save(output_web)
                output_web.seek(0)

                st.markdown("---")
                st.markdown("### 📥 Exportar Resultados")
                st.download_button(
                    label="📥 Exportar Planilha com FC_PROJETADO Oficial (Com Fórmulas)",
                    data=output_web,
                    file_name=f"FC_PROJETADO_{nome_cliente.upper()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Erro ao processar as fórmulas com o modelo: {e}")
else:
    st.info("ℹ️ Insira os dados de simulação e carregue a planilha base para ver os resultados.")
