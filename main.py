import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import streamlit as st
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go

# 🔹 Configuração da página
st.set_page_config(
    page_title="Dashboard de Chamados",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 🔹 Estado da Sessão para persistência de dados
if "df" not in st.session_state:
    st.session_state.df = None
if "filters" not in st.session_state:
    st.session_state.filters = {}
if "uploaded_file_name" not in st.session_state:
    st.session_state.uploaded_file_name = None

# 🔹 Upload do arquivo pelo usuário
uploaded_file = st.sidebar.file_uploader("Selecione um arquivo Excel", type=["xlsx"])

@st.cache_data(ttl=3600)  # Cache de 1 hora
def load_data(file):
    try:
        df = pd.read_excel(file)
        date_cols = ["Data da abertura", "Data de término do atendimento"]
        for col in date_cols:
            df[col] = pd.to_datetime(df[col], errors='coerce')
        return df
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {str(e)}")
        return pd.DataFrame()

# 🔹 Carregar dados apenas se o usuário fizer upload e manter entre recarregamentos
if uploaded_file is not None:
    if uploaded_file.name != st.session_state.uploaded_file_name or st.session_state.df is None:
        st.session_state.df = load_data(uploaded_file)
        st.session_state.uploaded_file_name = uploaded_file.name

df = st.session_state.df

if df is None or df.empty:
    st.warning("Por favor, faça o upload de um arquivo Excel para continuar.")
    st.stop()

# 🔹 Pré-processamento
df["Tempo de Atendimento (dias)"] = (df["Data de término do atendimento"] - df["Data da abertura"]).dt.days
df["Categoria 1"] = df.apply(lambda x: "INC-Sistemas Corporativos" if x["Categoria 2"] == "SAP" else x["Categoria 1"], axis=1)
df["Mês"] = df["Data da abertura"].dt.to_period("M").astype(str)

# 🔹 Sidebar - Filtros
st.sidebar.header("📊 Filtros")

# Filtro de data (não vai mudar)
min_date = df["Data da abertura"].min().date()
max_date = df["Data da abertura"].max().date()

date_range = st.sidebar.date_input(
    "Selecione o período:",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date
)

# Filtros Avançados
with st.sidebar.expander("🔍 Filtros Avançados", expanded=True):
    # Definição de valores padrão
    categorias_default = df["Categoria 2"].dropna().unique().tolist()
    situacoes_default = df["Situação"].dropna().unique().tolist()
    tipos_default = df["Categoria 1"].dropna().unique().tolist()
    etapas_default = df["Etapa"].dropna().unique().tolist()
    equipes_default = df["Equipe"].dropna().unique().tolist()
    responsaveis_default = df["Responsável"].dropna().unique().tolist()

    # Filtro "Situação"
    situacoes = st.multiselect("Situações:", options=df["Situação"].dropna().unique(), default=situacoes_default)

    # Filtrando o DataFrame de acordo com as situações selecionadas
    filtered_df = df[df["Situação"].isin(situacoes)]

    # 🔹 Filtrar categorias disponíveis dependendo das situações
    categorias_options = filtered_df["Categoria 2"].dropna().unique()
    categorias_default = [categoria for categoria in categorias_default if categoria in categorias_options]

    # Remover 'MTEBatch' caso não esteja nas opções disponíveis
    if 'MTEBatch' not in categorias_options:
        categorias_default = [categoria for categoria in categorias_default if categoria != 'MTEBatch']

    # Filtro "Categoria 2"
    categorias = st.multiselect(
        "Categorias:",
        options=categorias_options,
        default=categorias_default
    )
    
    # Filtro "Categoria 1", dependendo das "Categorias" selecionadas
    tipos_options = filtered_df[filtered_df["Categoria 2"].isin(categorias)]["Categoria 1"].dropna().unique()
    tipos_default = [tipo for tipo in tipos_default if tipo in tipos_options]

    tipos = st.multiselect(
        "Categoria 1:",
        options=tipos_options,
        default=tipos_default
    )

    # Filtro "Etapa", dependendo das "Situações"
    etapas_options = filtered_df["Etapa"].dropna().unique()
    etapas_default = [etapa for etapa in etapas_default if etapa in etapas_options]

    # Caso o filtro "Etapa" esteja vazio, defina um valor padrão (como uma lista vazia)
    if not etapas_default:
        etapas_default = []

    etapas = st.multiselect(
        "Etapa:",
        options=etapas_options,
        default=etapas_default
    )

    # Filtro "Equipe", dependendo das "Situações"
    equipes_options = filtered_df["Equipe"].dropna().unique()
    equipes_default = [equipe for equipe in equipes_default if equipe in equipes_options]

    equipes = st.multiselect(
        "Equipe:",
        options=equipes_options,
        default=equipes_default
    )

    # 🔹 Filtro "Responsável", dependendo das "Situações"
    responsaveis_options = filtered_df["Responsável"].dropna().unique()
    responsaveis_default = [responsavel for responsavel in responsaveis_default if responsavel in responsaveis_options]

    # Caso o filtro "Responsável" esteja vazio, defina um valor padrão (como uma lista vazia)
    if not responsaveis_default:
        responsaveis_default = []

    responsaveis = st.multiselect(
        "Responsável:",
        options=responsaveis_options,
        default=responsaveis_default
    )

# 🔹 Aplicar filtros
start_date = pd.to_datetime(date_range[0]) if len(date_range) > 0 else min_date
end_date = pd.to_datetime(date_range[1]) if len(date_range) > 1 else max_date

mask = (
    (df["Data da abertura"].between(start_date, end_date)) & 
    (df["Categoria 2"].isin(categorias)) & 
    (df["Situação"].isin(situacoes)) & 
    (df["Categoria 1"].isin(tipos)) & 
    (df["Equipe"].isin(equipes)) & 
    (df["Responsável"].isin(responsaveis)) & 
    (df["Etapa"].isin(etapas))
)

df_filtered = df[mask].copy()

# 🔹 Verificação de dados filtrados
if df_filtered.empty:
    st.warning("Nenhum dado encontrado com os filtros selecionados.")
    st.stop()

# 🔹 Métricas rápidas
st.subheader("📈 Métricas Chave")
col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    st.metric("Total Chamados", df_filtered.shape[0], help="Total de chamados no período selecionado")
with col2:
    avg_time = df_filtered["Tempo de Atendimento (dias)"].mean()
    st.metric("Tempo Médio", f"{avg_time:.1f} dias", delta=f"{avg_time - df['Tempo de Atendimento (dias)'].mean():.1f} vs histórico")
with col3:
    sla_servico = df_filtered["Atraso no serviço"].value_counts(normalize=True).get("Não", 0)*100
    st.metric("SLA Serviço", f"{sla_servico:.1f}%", help="Percentual de chamados dentro do SLA de serviço")
with col4:
    incidentes = df_filtered[df_filtered["Categoria 1"] == "INC-Sistemas Corporativos"].shape[0]
    st.metric("Incidentes", incidentes)
with col5:
    requisicoes = df_filtered[df_filtered["Categoria 1"] == "REQ-Sistemas Corporativos"].shape[0]
    st.metric("Requisições", requisicoes)

# 🔹 Abas para organização
tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Distribuição", "📅 Evolução Temporal", "⚙️ Análise de Etapas", "📋 Detalhes", "🏆 Top 10"])

with tab1:
    # Gráficos de distribuição
    col1, col2 = st.columns(2)
    
    with col1:
        if not df_filtered.empty and "Categoria 2" in df_filtered.columns and "Categoria 1" in df_filtered.columns:
            fig = px.bar(df_filtered, y="Categoria 2", color="Categoria 1", title="Distribuição por Categoria e Tipo", 
                         labels={"Categoria 2": "Categoria", "Categoria 1": "Tipo de Chamado"})
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Nenhum dado disponível para o gráfico de Distribuição por Categoria e Tipo.")
    
    with col2:
        if not df_filtered.empty and "Tempo de Atendimento (dias)" in df_filtered.columns and "Categoria 1" in df_filtered.columns:
            df_valid = df_filtered.dropna(subset=["Tempo de Atendimento (dias)"])
            if not df_valid.empty:
                fig = px.histogram(df_valid, x="Tempo de Atendimento (dias)", color="Categoria 1", 
                                    title="Distribuição do Tempo de Atendimento", 
                                    labels={"Tempo de Atendimento (dias)": "Dias para Resolução", "Categoria 1": "Tipo de Chamado"})
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("Nenhum dado válido para o gráfico de Tempo de Atendimento.")
        else:
            st.warning("Nenhum dado disponível para o gráfico de Distribuição do Tempo de Atendimento.")

with tab2:
    # Evolução Temporal
    st.subheader("Evolução Temporal de Chamados")
    if not df_filtered.empty and "Mês" in df_filtered.columns:
        fig = px.line(df_filtered.groupby("Mês").size().reset_index(name="Chamados"), x="Mês", y="Chamados", 
                      title="Evolução de Chamados ao Longo do Tempo", labels={"Mês": "Mês", "Chamados": "Quantidade de Chamados"})
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("Nenhum dado disponível para a análise de Evolução Temporal.")

with tab3:
    # Análise por Etapas
    st.subheader("Distribuição por Etapas")
    if not df_filtered.empty and "Etapa" in df_filtered.columns:
        # Conta as ocorrências de cada etapa
        etapas_count = df_filtered["Etapa"].value_counts().reset_index()
        
        # Renomeia as colunas corretamente
        etapas_count.columns = ["Etapa", "Chamados"]  # "index" vira "Etapa" e a contagem vira "Chamados"
        
        # Cria o gráfico de barras
        fig = px.bar(etapas_count, x="Etapa", y="Chamados", 
                     title="Distribuição de Chamados por Etapa", 
                     labels={"Etapa": "Etapa", "Chamados": "Quantidade de Chamados"})
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("Nenhum dado disponível para a análise de Etapas.")

with tab4:
    # Tabela detalhada
    with st.expander("📋 Dados Completos", expanded=True):
        columns_to_show = [
            "Atendimento",
            "Categoria 1",
            "Categoria 2",
            "Etapa",
            "Situação",
            "Prioridade",
            "Tempo de Atendimento (dias)",
            "Data da abertura"
        ]
        
        if all(col in df_filtered.columns for col in columns_to_show):
            st.dataframe(
                df_filtered[columns_to_show].sort_values("Tempo de Atendimento (dias)", ascending=False),
                height=500,
                use_container_width=True
            )
        else:
            st.warning("Algumas colunas não foram encontradas nos dados.")

with tab5:
    # Top 10
    st.subheader("🏆 Top 10")
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Top 10 Equipes")
        if not df_filtered.empty and "Equipe" in df_filtered.columns:
            # Conta as ocorrências de cada equipe e renomeia as colunas
            top_equipes = df_filtered["Equipe"].value_counts().nlargest(10).reset_index()
            top_equipes.columns = ["Equipe", "Chamados"]  # Renomeia "index" para "Equipe"
            
            # Cria o gráfico de barras
            fig = px.bar(top_equipes, x="Equipe", y="Chamados", title="Top 10 Equipes", 
                         labels={"Equipe": "Equipe", "Chamados": "Quantidade de Chamados"})
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Nenhum dado disponível para o gráfico de Top 10 Equipes.")
    
    with col2:
        st.subheader("Top 10 Responsáveis")
        if not df_filtered.empty and "Responsável" in df_filtered.columns:
            # Conta as ocorrências de cada responsável e renomeia as colunas
            top_responsaveis = df_filtered["Responsável"].value_counts().nlargest(10).reset_index()
            top_responsaveis.columns = ["Responsável", "Chamados"]  # Renomeia "index" para "Responsável"
            
            # Cria o gráfico de barras
            fig = px.bar(top_responsaveis, x="Responsável", y="Chamados", title="Top 10 Responsáveis", 
                         labels={"Responsável": "Responsável", "Chamados": "Quantidade de Chamados"})
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Nenhum dado disponível para o gráfico de Top 10 Responsáveis.")
    

# 🔹 Estilos CSS
st.markdown(""" 
    <style>
    .stMetric {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 15px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .stMetric h3 {
        color: #1e88e5 !important;
        font-size: 1.2rem !important;
    }
    .stMetric label {
        font-size: 0.9rem !important;
        color: #6c757d !important;
    }
    </style>
""", unsafe_allow_html=True)

# 🔹 Rodapé
st.markdown("---")
st.markdown("**Dashboard desenvolvido por Bionovis S/A - Atualizado em {}".format(datetime.now().strftime("%d/%m/%Y")))