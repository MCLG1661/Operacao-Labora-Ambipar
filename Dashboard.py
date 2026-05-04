import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import openpyxl

# Configuração da página
st.set_page_config(
    page_title="Dashboard LABORA/AMBIPAR",
    page_icon="📊",
    layout="wide"
)

# Título principal
st.title("📊 LABORA/AMBIPAR - ANÁLISE DE DADOS 📝")
st.markdown("---")

# Carrega a planilha
@st.cache_data
def carregar_dados():
    df = pd.read_excel(
        r'F:\Pessoal\Labora\Project Manager\Planilhas\PLANILHA DE RECOLHA DE NOTAS.xlsx',
        engine='openpyxl'
    )
    # Renomeia as colunas para facilitar
    df.columns = ['POSTO', 'B', 'C', 'CNPJ_AMBIPAR', 'E', 'F', 'NOTAS_POR_POSTO', 'H', 'I', 'VALOR_TOTAL_RECOLHIDO']
    
    # Remove linhas vazias (última linha com SUM)
    df = df.dropna(subset=['POSTO', 'NOTAS_POR_POSTO', 'VALOR_TOTAL_RECOLHIDO'])
    
    # Converte para numérico
    df['NOTAS_POR_POSTO'] = pd.to_numeric(df['NOTAS_POR_POSTO'], errors='coerce')
    df['VALOR_TOTAL_RECOLHIDO'] = pd.to_numeric(df['VALOR_TOTAL_RECOLHIDO'], errors='coerce')
    
    return df

df = carregar_dados()

# ==================== MÉTRICAS PRINCIPAIS ====================
st.subheader("📈 VISÃO GERAL")

col1, col2, col3, col4 = st.columns(4)

with col1:
    total_postos = df['POSTO'].nunique()
    st.metric("🏪 Total de Postos", total_postos)

with col2:
    total_cnpjs = df['CNPJ_AMBIPAR'].nunique()
    st.metric("🏢 Total de CNPJs", total_cnpjs)

with col3:
    total_notas = df['NOTAS_POR_POSTO'].sum()
    st.metric("📄 Total de Notas Fiscais", f"{total_notas:,.0f}")

with col4:
    total_recolhido = df['VALOR_TOTAL_RECOLHIDO'].sum()
    st.metric("💰 Valor Total Recolhido", f"R$ {total_recolhido:,.2f}")

st.markdown("---")

# ==================== GRÁFICOS ====================

# Layout com 2 colunas para os gráficos
col_esq, col_dir = st.columns(2)

# Gráfico 1: Top 10 postos com maior valor recolhido
with col_esq:
    st.subheader("🏆 Top 10 Postos - Maior Recolhimento")
    top_valor = df.groupby('POSTO')['VALOR_TOTAL_RECOLHIDO'].sum().reset_index()
    top_valor = top_valor.sort_values('VALOR_TOTAL_RECOLHIDO', ascending=False).head(10)
    
    fig1 = px.bar(
        top_valor, 
        x='VALOR_TOTAL_RECOLHIDO', 
        y='POSTO',
        orientation='h',
        title="Top 10 Postos por Valor Recolhido (R$)",
        color='VALOR_TOTAL_RECOLHIDO',
        color_continuous_scale='Blues',
        text_auto='.2s'
    )
    fig1.update_layout(height=500, yaxis={'categoryorder':'total ascending'})
    st.plotly_chart(fig1, use_container_width=True)

# Gráfico 2: Top 10 CNPJs com maior valor recolhido
with col_dir:
    st.subheader("🏆 Top 10 CNPJs - Maior Recolhimento")
    top_cnpj_valor = df.groupby('CNPJ_AMBIPAR').agg({
        'VALOR_TOTAL_RECOLHIDO': 'sum',
        'NOTAS_POR_POSTO': 'sum'
    }).reset_index()
    top_cnpj_valor = top_cnpj_valor.sort_values('VALOR_TOTAL_RECOLHIDO', ascending=False).head(10)
    
    # Formata o CNPJ para exibição (XX.XXX.XXX/XXXX-XX)
    top_cnpj_valor['CNPJ_FORMATADO'] = top_cnpj_valor['CNPJ_AMBIPAR'].astype(str).str.zfill(14)
    top_cnpj_valor['CNPJ_FORMATADO'] = top_cnpj_valor['CNPJ_FORMATADO'].str.replace(
        r'(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})', r'\1.\2.\3/\4-\5', regex=True
    )
    
    fig2 = px.bar(
        top_cnpj_valor,
        x='VALOR_TOTAL_RECOLHIDO',
        y='CNPJ_FORMATADO',
        orientation='h',
        title="Top 10 CNPJs por Valor Recolhido (R$)",
        color='VALOR_TOTAL_RECOLHIDO',
        color_continuous_scale='Purples',
        text_auto='.2s'
    )
    fig2.update_layout(height=500, yaxis={'categoryorder':'total ascending'})
    st.plotly_chart(fig2, use_container_width=True)

# ==================== SEGUNDA LINHA DE GRÁFICOS ====================

col_notas, col_cnpj_notas = st.columns(2)

# Gráfico 3: Top 10 postos com mais notas fiscais
with col_notas:
    st.subheader("📄 Top 10 Postos - Mais Notas Fiscais")
    top_notas = df.groupby('POSTO')['NOTAS_POR_POSTO'].sum().reset_index()
    top_notas = top_notas.sort_values('NOTAS_POR_POSTO', ascending=False).head(10)
    
    fig3 = px.bar(
        top_notas,
        x='NOTAS_POR_POSTO',
        y='POSTO',
        orientation='h',
        title="Top 10 Postos por Quantidade de Notas",
        color='NOTAS_POR_POSTO',
        color_continuous_scale='Greens',
        text_auto=True
    )
    fig3.update_layout(height=500, yaxis={'categoryorder':'total ascending'})
    st.plotly_chart(fig3, use_container_width=True)

# Gráfico 4: Top 10 CNPJs com mais notas fiscais
with col_cnpj_notas:
    st.subheader("📄 Top 10 CNPJs - Mais Notas Fiscais")
    top_cnpj_notas = df.groupby('CNPJ_AMBIPAR').agg({
        'NOTAS_POR_POSTO': 'sum',
        'VALOR_TOTAL_RECOLHIDO': 'sum'
    }).reset_index()
    top_cnpj_notas = top_cnpj_notas.sort_values('NOTAS_POR_POSTO', ascending=False).head(10)
    
    # Formata o CNPJ
    top_cnpj_notas['CNPJ_FORMATADO'] = top_cnpj_notas['CNPJ_AMBIPAR'].astype(str).str.zfill(14)
    top_cnpj_notas['CNPJ_FORMATADO'] = top_cnpj_notas['CNPJ_FORMATADO'].str.replace(
        r'(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})', r'\1.\2.\3/\4-\5', regex=True
    )
    
    fig4 = px.bar(
        top_cnpj_notas,
        x='NOTAS_POR_POSTO',
        y='CNPJ_FORMATADO',
        orientation='h',
        title="Top 10 CNPJs por Quantidade de Notas",
        color='NOTAS_POR_POSTO',
        color_continuous_scale='Oranges',
        text_auto=True
    )
    fig4.update_layout(height=500, yaxis={'categoryorder':'total ascending'})
    st.plotly_chart(fig4, use_container_width=True)

# ==================== TABELAS COMPLETAS ====================

st.markdown("---")
st.subheader("📋 ANÁLISE COMPLETA POR POSTO")

# Agrupa os dados por posto
df_agrupado = df.groupby('POSTO').agg({
    'CNPJ_AMBIPAR': 'first',
    'NOTAS_POR_POSTO': 'sum',
    'VALOR_TOTAL_RECOLHIDO': 'sum'
}).reset_index()

# Adiciona coluna de média por nota
df_agrupado['MÉDIA_POR_NOTA'] = df_agrupado['VALOR_TOTAL_RECOLHIDO'] / df_agrupado['NOTAS_POR_POSTO']
df_agrupado['MÉDIA_POR_NOTA'] = df_agrupado['MÉDIA_POR_NOTA'].fillna(0)

# Ordena por valor recolhido
df_agrupado = df_agrupado.sort_values('VALOR_TOTAL_RECOLHIDO', ascending=False)

# Exibe a tabela
st.dataframe(
    df_agrupado,
    column_config={
        "POSTO": "Nome do Posto",
        "CNPJ_AMBIPAR": "CNPJ",
        "NOTAS_POR_POSTO": st.column_config.NumberColumn("Total de Notas", format="%d"),
        "VALOR_TOTAL_RECOLHIDO": st.column_config.NumberColumn("Valor Recolhido (R$)", format="R$ %.2f"),
        "MÉDIA_POR_NOTA": st.column_config.NumberColumn("Média por Nota (R$)", format="R$ %.2f"),
    },
    use_container_width=True,
    hide_index=True
)

# ==================== TABELA POR CNPJ (SEM POSTO E SEM CNPJ CRU) ====================

st.markdown("---")
st.subheader("📋 ANÁLISE COMPLETA POR CNPJ")

# Agrupa os dados por CNPJ
df_cnpj = df.groupby('CNPJ_AMBIPAR').agg({
    'NOTAS_POR_POSTO': 'sum',
    'VALOR_TOTAL_RECOLHIDO': 'sum'
}).reset_index()

# Adiciona coluna de média por nota
df_cnpj['MÉDIA_POR_NOTA'] = df_cnpj['VALOR_TOTAL_RECOLHIDO'] / df_cnpj['NOTAS_POR_POSTO']
df_cnpj['MÉDIA_POR_NOTA'] = df_cnpj['MÉDIA_POR_NOTA'].fillna(0)

# Ordena por valor recolhido
df_cnpj = df_cnpj.sort_values('VALOR_TOTAL_RECOLHIDO', ascending=False)

# Formata o CNPJ para exibição
df_cnpj['CNPJ_FORMATADO'] = df_cnpj['CNPJ_AMBIPAR'].astype(str).str.zfill(14)
df_cnpj['CNPJ_FORMATADO'] = df_cnpj['CNPJ_FORMATADO'].str.replace(
    r'(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})', r'\1.\2.\3/\4-\5', regex=True
)

# Exibe a tabela (somente CNPJ formatado, Notas, Valor e Média)
st.dataframe(
    df_cnpj[['CNPJ_FORMATADO', 'NOTAS_POR_POSTO', 'VALOR_TOTAL_RECOLHIDO', 'MÉDIA_POR_NOTA']],
    column_config={
        "CNPJ_FORMATADO": "CNPJ",
        "NOTAS_POR_POSTO": st.column_config.NumberColumn("Total de Notas", format="%d"),
        "VALOR_TOTAL_RECOLHIDO": st.column_config.NumberColumn("Valor Recolhido (R$)", format="R$ %.2f"),
        "MÉDIA_POR_NOTA": st.column_config.NumberColumn("Média por Nota (R$)", format="R$ %.2f"),
    },
    use_container_width=True,
    hide_index=True
)

# ==================== TOP 10 CNPJs EM DESTAQUE ====================

st.markdown("---")
st.subheader("🏆 RESULTADOS COMPLETOS - TOP 10 CNPJs")

col_top_valor, col_top_notas = st.columns(2)

# Top 10 CNPJs por Valor (sem Posto Associado e sem CNPJ cru)
with col_top_valor:
    st.markdown("### 💰 Maior Valor Recolhido")
    top10_valor = df_cnpj[['CNPJ_FORMATADO', 'VALOR_TOTAL_RECOLHIDO']].head(10)
    st.dataframe(
        top10_valor,
        column_config={
            "CNPJ_FORMATADO": "CNPJ",
            "VALOR_TOTAL_RECOLHIDO": st.column_config.NumberColumn("Valor (R$)", format="R$ %.2f"),
        },
        hide_index=True,
        use_container_width=True
    )

# Top 10 CNPJs por Notas (sem Posto Associado e sem CNPJ cru)
with col_top_notas:
    st.markdown("### 📄 Maior Quantidade de Notas")
    top10_notas = df_cnpj[['CNPJ_FORMATADO', 'NOTAS_POR_POSTO']].head(10)
    st.dataframe(
        top10_notas,
        column_config={
            "CNPJ_FORMATADO": "CNPJ",
            "NOTAS_POR_POSTO": st.column_config.NumberColumn("Notas", format="%d"),
        },
        hide_index=True,
        use_container_width=True
    )

# ==================== DOWNLOAD DOS DADOS ====================
st.markdown("---")
st.subheader("📥 Exportar Dados")

col_download1, col_download2 = st.columns(2)

with col_download1:
    csv_postos = df_agrupado.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📎 Baixar dados por POSTO (CSV)",
        data=csv_postos,
        file_name="analise_postos_ambipar.csv",
        mime="text/csv",
    )

with col_download2:
    # Para o CSV de CNPJ, usar o DataFrame sem as colunas removidas
    df_cnpj_export = df_cnpj[['CNPJ_FORMATADO', 'NOTAS_POR_POSTO', 'VALOR_TOTAL_RECOLHIDO', 'MÉDIA_POR_NOTA']]
    csv_cnpjs = df_cnpj_export.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📎 Baixar dados por CNPJ (CSV)",
        data=csv_cnpjs,
        file_name="analise_cnpjs_ambipar.csv",
        mime="text/csv",
    )

# ==================== RODAPÉ ====================
st.markdown("---")
st.caption(f"📊 Dados atualizados | Total de registros: {len(df)} | Postos únicos: {total_postos} | CNPJs únicos: {total_cnpjs}")