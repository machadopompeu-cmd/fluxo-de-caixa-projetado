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

# Nome do arquivo de modelo que contém a estrutura e as fórmulas corretas
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

# Campos de introdução numérica direta via digitação
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

# 3. COMPUTAÇÃO E SIMULAÇÃO DINÂMICA
if arquivo_carregado is not None:
    # Verificar se o arquivo de modelo (template) existe no diretório
    if not os.path.exists(ARQUIVO_TEMPLATE):
        st.error(f"Erro: O arquivo modelo '{ARQUIVO_TEMPLATE}' não foi encontrado no servidor para servir de template.")
    else:
        try:
            r_adj = perc_receita / 100.0
            l_adj = perc_lucro / 100.0

            # 1. Carregar os dados históricos enviados pelo usuário
            arquivo_carregado.seek(0)
            wb_usuario = openpyxl.load_workbook(io.BytesIO(arquivo_carregado.read()), data_only=True)
            aba_usuario_nome = wb_usuario.sheetnames[0]
            sheet_usuario = wb_usuario[aba_usuario_nome]

            # 2. Carregar o arquivo de template (com fórmulas completas de FC_PROJETADO)
            wb_template = openpyxl.load_workbook(ARQUIVO_TEMPLATE, data_only=False)
            sheet_bal_template = wb_template['BAL_25']
            sheet_fc_template = wb_template['FC_PROJETADO']

            # 3. Copiar os dados históricos do usuário para a aba BAL_25 do nosso template
            for r in range(1, sheet_usuario.max_row + 1):
                for c in range(1, sheet_usuario.max_column + 1):
                    val = sheet_usuario.cell(row=r, column=c).value
                    if val is not None:
                        sheet_bal_template.cell(row=r, column=c, value=val)

            # 4. Injetar as variáveis nos campos de controle exatos do template
            sheet_fc_template['B2'] = r_adj
            sheet_fc_template['B3'] = l_adj
            sheet_fc_template['A1'] = f"CLIENTE: {nome_cliente.upper()}"

            # 5. Motor de cálculo em Python para exibir os números em tempo real no ecrã
            colunas_meses_indices = [4, 7, 10, 13, 16, 19, 22, 25, 28, 31, 34, 37] # D, G, J, M, P, S, V, Y, AB, AE, AH, AK
            nomes_meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

            # Realizar a leitura do BAL_25 copiado para calcular os indicadores em tempo real
            receita_total_historica = 0.0
            despesa_total_historica = 0.0
            compra_mercadoria_historica = 0.0
            insumos_historicos = 0.0

            for col_idx in colunas_meses_indices:
                # Receitas Operacionais (Linha 6)
                val_rec = sheet_bal_template.cell(row=6, column=col_idx).value
                receita_total_historica += float(val_rec) if isinstance(val_rec, (int, float)) else 0.0
                
                # Despesa Total (Linha 98)
                val_desp = sheet_bal_template.cell(row=98, column=col_idx).value
                despesa_total_historica += float(val_desp) if isinstance(val_desp, (int, float)) else 0.0

                # Compra de Mercadoria (Linha 14)
                val_compra = sheet_bal_template.cell(row=14, column=col_idx).value
                compra_mercadoria_historica += float(val_compra) if isinstance(val_compra, (int, float)) else 0.0

                # Insumos (Linha 21)
                val_insumos = sheet_bal_template.cell(row=21, column=col_idx).value
                insumos_historicos += float(val_insumos) if isinstance(val_insumos, (int, float)) else 0.0

            # Cálculos dos Indicadores (Idênticos ao Excel)
            receita_projetada = receita_total_historica * (1 + r_adj)
            despesa_base = despesa_total_historica
            despesas_variaveis_base = compra_mercadoria_historica + insumos_historicos

            if despesas_variaveis_base == 0:
                fator_ajuste = 1.0
            else:
                numerador = (receita_projetada * (1 - l_adj)) - (despesa_base - despesas_variaveis_base)
                fator_ajuste = numerador / despesas_variaveis_base
                fator_ajuste = max(0.0, min(1.0, fator_ajuste))

            despesas_fixas = despesa_base - despesas_variaveis_base
            despesa_projetada = despesas_fixas + (despesas_variaveis_base * fator_ajuste)
            resultado_projetado = receita_projetada - despesa_projetada
            margem_projetada = (resultado_projetado / receita_projetada) if receita_projetada != 0 else 0.0

            # Exibição dos indicadores em tela
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
                    for col_idx in colunas_meses_indices:
                        val = sheet_bal_template.cell(row=r, column=col_idx).value
                        valores_originais.append(float(val) if isinstance(val, (int, float)) else 0.0)
                    
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
            
            # Formatação de visualização de moeda em tela
            df_formatado = df_visual.copy()
            for col in ["Total Projetado"] + nomes_meses:
                df_formatado[col] = df_formatado[col].apply(lambda x: f"R$ {x:,.2f}" if isinstance(x, (int, float)) else x)
                
            st.dataframe(df_formatado, use_container_width=True, height=450)

            # 6. Salvar o arquivo de modelo modificado na memória para exportação
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
