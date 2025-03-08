from matplotlib import pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# 🔄 Configuração da página
st.set_page_config(
    page_title="Dashboard de Chamados - Bionovis",
    layout="wide",
    initial_sidebar_state="expanded",
    page_icon="📊"
)

# 📥 Função para carregar dados
@st.cache_data(ttl=3600)
def load_data(file):
    try:
        # Carregamento do arquivo
        df = pd.read_excel(file)
        
        # Tratamento de colunas de data
        date_cols = ["Data da abertura", "Data de término do atendimento"]
        for col in date_cols:
            df[col] = pd.to_datetime(df[col], errors='coerce')
        
        # NOVO TRATAMENTO DE DADOS
        # Criar máscaras para as condições
        mask_cancelado = df['Situação'] == 'Cancelado'
        mask_atendimento = df['Situação'].isin(['Em Atendimento', 'Aguardando Atendimento'])
        mask_etapa_reprovado = df['Etapa'].isin(['Reprovado', 'Atendimento Reprovado'])
        mask_encerrado = df['Situação'] == 'Aguardando confirmação de encerramento'
        mask_suspenso = df['Situação'] == 'Suspenso'
        mask_incidentes = df['Categoria 2'] == 'SAP'

        # Aplicar as regras de substituição usando .loc
        df.loc[mask_cancelado & mask_etapa_reprovado, ['Situação', 'Etapa']] = 'Reprovado'
        df.loc[mask_cancelado & ~mask_etapa_reprovado, ['Situação', 'Etapa']] = 'Cancelado'
        df.loc[mask_encerrado, 'Situação'] = 'Encerrado'
        df.loc[mask_atendimento, 'Situação'] = 'Em Atendimento'
        df.loc[mask_incidentes, 'Categoria 2'] = 'Incidente'
        df.loc[mask_suspenso, 'Situação'] = 'Aguardando Aprovação/Pausado'

        # Cálculo de métricas temporais
        df["Tempo de Atendimento (dias)"] = (df["Data de término do atendimento"] - df["Data da abertura"]).dt.days
        
        # Criação de novas features temporais
        df["Mês"] = df["Data da abertura"].dt.to_period("M").astype(str)
        
        # Mapeamento dos dias da semana para português
        dias_semana_ingles = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        dias_semana_portugues = ['Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado', 'Domingo']
        df["Dia da Semana"] = df["Data da abertura"].dt.day_name().map(dict(zip(dias_semana_ingles, dias_semana_portugues)))
        
        df["Hora"] = df["Data da abertura"].dt.hour

        # Ordenação dos dias da semana
        dias_semana = ['Segunda-feira', 'Terça-feira', 'Quarta-feira', 
                    'Quinta-feira', 'Sexta-feira', 'Sábado', 'Domingo']
        df['Dia da Semana'] = pd.Categorical(
            df['Dia da Semana'], 
            categories=dias_semana, 
            ordered=True
        )
        
        # Validação de dados críticos
        if df['Situação'].isnull().any():
            st.warning("⚠️ Existem registros com situação indefinida")
        
        return df
    
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {str(e)}")
        return pd.DataFrame()
    
def sidebar_filters(df):
    # Garantir que estamos trabalhando com uma cópia
    df = df.copy()
    st.sidebar.header("🎛️ Filtros")
    
    # Filtro de data com slider
    st.sidebar.markdown("### 📅 Selecione o Período")
    
    # Garantindo que as datas sejam extraídas corretamente
    min_date = df["Data da abertura"].min().date()
    max_date = df["Data da abertura"].max().date()

    # Slider para seleção de intervalo de datas
    date_range = st.sidebar.slider(
        "Selecione o período:",
        min_value=min_date,
        max_value=max_date,
        value=(min_date, max_date),
        format="DD/MM/YYYY"
    )

    # Filtros de categorias
    with st.sidebar.expander("📂 Filtros Principais", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            # Usando .loc para evitar SettingWithCopyWarning
            situacoes = st.multiselect(
                "Situação dos Chamados:",
                options=df["Situação"].unique(),
                default=df["Situação"].unique(),
                key="situacoes"
            )
            categorias = st.multiselect(
                "Tipos e Sistemas:",
                options=df["Categoria 2"].unique(),
                default=df["Categoria 2"].unique(),
                key="categorias"
            )
        with col2:
            tipos = st.multiselect(
                "Categorias:",
                options=df["Categoria 1"].unique(),
                default=df["Categoria 1"].unique(),
                key="tipos"
            )
            prioridades = st.multiselect(
                "Criticidade:",
                options=df["Prioridade"].unique(),
                default=df["Prioridade"].unique(),
                key="prioridades"
            )
    
    # Filtros de equipe
    with st.sidebar.expander("👥 Outros Filtros", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            equipes = st.multiselect(
                "Equipes:",
                options=df["Equipe"].unique(),
                default=df["Equipe"].unique(),
                key="equipes"
            )
        with col2:
            responsaveis = st.multiselect(
                "Responsáveis:",
                options=df["Responsável"].unique(),
                default=df["Responsável"].unique(),
                key="responsaveis"
            )
    
    return date_range, situacoes, categorias, tipos, equipes, responsaveis, prioridades

# 📊 Visualizações
def create_metrics(df):
    st.subheader("📈 Métricas Principais")
    cols = st.columns(6)
    
    sla_percent = df['Atraso no serviço'].value_counts(normalize=True).get('Não', 0)*100
    tempo_medio = df['Tempo de Atendimento (dias)'].mean()
    
    metrics = [
        ("Total Chamados", df.shape[0], "📋", "#000000"),
        ("Tempo Médio", f"{tempo_medio:.1f} dias", "⏳", "#ff9800"),
        ("Dentro do SLA", f"{sla_percent:.1f}%", "✅", "#4caf50"),
        ("Incidentes", df[df["Categoria 1"] == "INC-Sistemas Corporativos"].shape[0], "⚠️", "#f44336"),
        ("Requisições", df[df["Categoria 1"] == "REQ-Sistemas Corporativos"].shape[0], "📋", "#9c27b0"),
        ("Reprovados", df[df['Situação'] == 'Reprovado'].shape[0], "⚠️", "#f44336")
    ]
    
    for col, (title, value, icon, color) in zip(cols, metrics):
        with col:
            st.markdown(f"""
                <div class="metric-card">
                    <h3>{icon} {title}</h3>
                    <div style="font-size: 24px; font-weight: bold; color: {color};">
                        {value}
                    </div>
                </div>
            """, unsafe_allow_html=True)

def safe_plot(func, *args, **kwargs):
    """Helper function para evitar quebras por erros nos gráficos"""
    try:
        return func(*args, **kwargs)
    except Exception as e:
        st.warning(f"Não foi possível exibir este gráfico: {str(e)}")
        return None

def create_charts(df, template):
    tabs = st.tabs(["📈 Análise Geral", "📅 Evolução Temporal", "⚙️ Detalhamento", "📊 Comparativos", "🔍 Insights", "📋 Dados Completos", "📉 Backlog"])

    # Mapeamento de dias da semana
    dias_semana = {
        'Monday': 'Segunda',
        'Tuesday': 'Terça',
        'Wednesday': 'Quarta',
        'Thursday': 'Quinta',
        'Friday': 'Sexta',
        'Saturday': 'Sábado',
        'Sunday': 'Domingo'
    }
      
    with tabs[0]:
        col1, col2 = st.columns(2)
        
        with col1:
            # Gráfico de Pizza - Top 10 Categorias
            if 'Categoria 2' in df.columns:
                top_categorias = df['Categoria 2'].value_counts().nlargest(10)
                if not top_categorias.empty:
                    fig = px.pie(
                        top_categorias,
                        names=top_categorias.index,
                        values=top_categorias.values,
                        title="Top 10 Categorias (Volume de Chamados)",
                        template=template,
                        hole=0.4,
                        color_discrete_sequence=px.colors.qualitative.Pastel
                    )
                    fig.update_traces(
                        textposition='inside',
                        textinfo='percent+label',
                        hovertemplate="<b>%{label}</b><br>Chamados: %{value}<br>(%{percent})"
                    )
                    fig.update_layout(
                        uniformtext_minsize=12,
                        uniformtext_mode='hide',
                        legend=dict(
                            orientation="h",
                            yanchor="bottom",
                            y=-0.2,
                            xanchor="center",
                            x=0.5
                        )
                    )
                    st.plotly_chart(fig, use_container_width=True)
            
            # Heatmap de Chamados por Hora/Dia
            if 'Data da abertura' in df.columns:
                df_heatmap = df.copy()  # Cria uma cópia explícita
                df_heatmap.loc[:, 'Hora'] = df_heatmap['Data da abertura'].dt.hour
                df_heatmap['Dia da Semana'] = df_heatmap['Data da abertura'].dt.day_name().map(dias_semana)
                
                heatmap_data = df_heatmap.pivot_table(
                    index='Hora',
                    columns='Dia da Semana',
                    values='Atendimento',
                    aggfunc='count',
                    fill_value=0
                )
                
                # Ordenando os dias da semana em português
                dias_ordem = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo']
                heatmap_data = heatmap_data.reindex(columns=dias_ordem)
                
                fig = px.imshow(
                    heatmap_data,
                    labels=dict(x="Dia da Semana", y="Hora do Dia", color="Chamados"),
                    title="Distribuição de Chamados por Hora/Dia",
                    template=template,
                    color_continuous_scale='Viridis'
                )
                fig.update_layout(
                    xaxis_nticks=7,
                    yaxis_nticks=24,
                    coloraxis_colorbar=dict(
                        title="Chamados",
                        thickness=15
                    )
                )
                st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            # Histograma de Tempo de Atendimento
            if 'Tempo de Atendimento (dias)' in df.columns:
                fig = px.histogram(
                    df,
                    x="Tempo de Atendimento (dias)",
                    color="Categoria 1",
                    title="Distribuição do Tempo de Atendimento",
                    template=template,
                    nbins=30,
                    opacity=0.8,
                    barmode='overlay',
                    color_discrete_sequence=px.colors.qualitative.Pastel
                )
                fig.update_layout(
                    xaxis_title="Dias para Resolução",
                    yaxis_title="Quantidade de Chamados",
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=-0.3,
                        xanchor="center",
                        x=0.5
                    )
                )
                st.plotly_chart(fig, use_container_width=True)
            
            # Top Categorias com Maior Tempo Médio
            if 'Categoria 2' in df.columns and 'Tempo de Atendimento (dias)' in df.columns:
                top_categorias_tempo = df.groupby('Categoria 2')['Tempo de Atendimento (dias)'].mean().nlargest(10)
                
                fig = px.bar(
                    top_categorias_tempo,
                    orientation='h',
                    title="Top 10 Categorias com Maior Tempo Médio",
                    template=template,
                    text_auto='.1f',
                    color=top_categorias_tempo.values,
                    color_continuous_scale='Viridis'
                )
                fig.update_layout(
                    xaxis_title="Tempo Médio (dias)",
                    yaxis_title="Categoria",
                    coloraxis_showscale=False,
                    yaxis={'categoryorder':'total ascending'}
                )
                fig.update_traces(
                    textposition='outside',
                    hovertemplate="<b>%{y}</b><br>Tempo Médio: %{x:.1f} dias"
                )
                st.plotly_chart(fig, use_container_width=True)
    
    with tabs[1]:
        # Mapeamento de meses para português
        meses_pt = {
            1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
            5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
            9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
        }

        # Gráfico 1: Evolução Mensal e SLA
        try:
            # Criando subplots
            fig = make_subplots(
                rows=2, cols=1,
                subplot_titles=('Evolução Mensal de Chamados', 'Tendência de SLA'),
                vertical_spacing=0.15
            )
            
            # Processando dados para evolução mensal
            df['Mês_num'] = df['Data da abertura'].dt.month
            evolucao_mensal = df.groupby("Mês_num").size()
            evolucao_mensal.index = evolucao_mensal.index.map(meses_pt)
            
            # Adicionando gráfico de linha para evolução
            fig.add_trace(
                go.Scatter(
                    x=evolucao_mensal.index,
                    y=evolucao_mensal.values,
                    mode='lines+markers',
                    name='Chamados',
                    line=dict(color='#1f77b4', width=3),
                    marker=dict(size=8, color='#1f77b4'),
                    hovertemplate='<b>%{x}</b><br>Chamados: %{y}'
                ),
                row=1, col=1
            )
            
            # Processando dados para SLA
            evolucao_sla = df.groupby("Mês_num")['Atraso no serviço'].apply(
                lambda x: (x == 'Não').mean() * 100
            )
            evolucao_sla.index = evolucao_sla.index.map(meses_pt)
            
            # Adicionando gráfico de linha para SLA
            fig.add_trace(
                go.Scatter(
                    x=evolucao_sla.index,
                    y=evolucao_sla.values,
                    mode='lines+markers',
                    name='% Dentro SLA',
                    line=dict(color='#2ca02c', width=3),
                    marker=dict(size=8, color='#2ca02c'),
                    hovertemplate='<b>%{x}</b><br>SLA: %{y:.1f}%'
                ),
                row=2, col=1
            )
            
            # Atualizando layout
            fig.update_layout(
                height=700,
                template=template,
                showlegend=False,
                margin=dict(t=50, b=50),
                hovermode='x unified'
            )
            
            # Configurações específicas dos eixos
            fig.update_yaxes(title_text="Quantidade de Chamados", row=1, col=1)
            fig.update_yaxes(title_text="% Dentro do SLA", row=2, col=1)
            fig.update_xaxes(title_text="Mês", row=2, col=1)
            
            st.plotly_chart(fig, use_container_width=True)
            
        except Exception as e:
            st.warning(f"Erro ao criar gráfico temporal: {str(e)}")
        
        # Gráfico 2: Evolução por Categoria
        try:
            # Processando dados
            evolucao_categorias = df.groupby(['Mês_num', 'Categoria 1']).size().unstack().fillna(0)
            evolucao_categorias.index = evolucao_categorias.index.map(meses_pt)
            
            # Criando gráfico de área
            fig = px.area(
                evolucao_categorias,
                title="Evolução por Tipo de Chamado",
                template=template,
                labels={'value': 'Quantidade', 'Mês_num': 'Mês'},
                color_discrete_sequence=px.colors.qualitative.Pastel
            )
            
            # Atualizando layout
            fig.update_layout(
                height=500,
                xaxis_title='Mês',
                yaxis_title='Quantidade de Chamados',
                legend_title='Categoria',
                hovermode='x unified',
                margin=dict(t=50, b=50)
            )
            
            # Melhorando tooltips
            fig.update_traces(
                hovertemplate='<b>%{x}</b><br>Chamados: %{y}'
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
        except Exception as e:
            st.warning(f"Erro ao criar gráfico de evolução: {str(e)}")
    
    with tabs[2]:
        col1, col2 = st.columns(2)
        
        with col1:
            # Gráfico Sunburst - Distribuição por Equipe
            try:
                # Pré-processamento dos dados
                sunburst_df = df[['Equipe', 'Responsável', 'Situação']].copy()
                
                # Preenchendo valores ausentes de forma hierárquica
                sunburst_df['Equipe'] = sunburst_df['Equipe'].fillna('Equipe Não Informada')
                sunburst_df['Responsável'] = sunburst_df['Responsável'].fillna(
                    sunburst_df['Equipe'] + ' - Responsável Não Informado'
                )
                sunburst_df['Situação'] = sunburst_df['Situação'].fillna('Situação Não Informada')
                
                # Criando o gráfico
                fig = px.sunburst(
                    sunburst_df,
                    path=['Equipe', 'Responsável', 'Situação'],
                    title="Distribuição por Equipe",
                    template=template,
                    color_discrete_sequence=px.colors.qualitative.Pastel,
                    height=500
                )
                
                # Melhorando a aparência
                fig.update_traces(
                    textinfo='label+percent parent',
                    hovertemplate='<b>%{label}</b><br>Chamados: %{value}',
                    insidetextorientation='radial'
                )
                fig.update_layout(
                    margin=dict(t=40, b=20, l=20, r=20),
                    uniformtext=dict(minsize=12, mode='hide')
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
            except Exception as e:
                st.warning(f"Erro no gráfico sunburst: {str(e)}")
            
            # Gráfico Treemap - Hierarquia de Chamados
            try:
                # Pré-processamento dos dados
                treemap_df = df[['Categoria 1', 'Categoria 2', 'Etapa']].copy()
                treemap_df = treemap_df.fillna('Não Informado')
                
                # Criando o gráfico
                fig = px.treemap(
                    treemap_df,
                    path=['Categoria 1', 'Categoria 2', 'Etapa'],
                    title="Hierarquia de Chamados",
                    template=template,
                    color_discrete_sequence=px.colors.qualitative.Pastel,
                    height=500
                )
                
                # Melhorando a aparência
                fig.update_traces(
                    textinfo='label+percent parent',
                    hovertemplate='<b>%{label}</b><br>Chamados: %{value}',
                    textposition='middle center'
                )
                fig.update_layout(
                    margin=dict(t=40, b=20, l=20, r=20),
                    uniformtext=dict(minsize=12, mode='hide')
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
            except Exception as e:
                st.warning(f"Erro no gráfico treemap: {str(e)}")
        
        with col2:
            # Gráfico de Barras - Prioridade e Situação
            try:
                # Processando os dados
                pivot_data = df.pivot_table(
                    index='Prioridade',
                    columns='Situação',
                    values='Atendimento',
                    aggfunc='count',
                    fill_value=0
                )
                
                # Criando o gráfico
                fig = px.bar(
                    pivot_data,
                    barmode='stack',
                    title="Distribuição por Prioridade e Situação",
                    template=template,
                    labels={'value': 'Quantidade', 'Situação': 'Situação'},
                    color_discrete_sequence=px.colors.qualitative.Pastel,
                    height=500
                )
                
                # Melhorando a aparência
                fig.update_layout(
                    xaxis_title='Prioridade',
                    yaxis_title='Quantidade de Chamados',
                    legend_title='Situação',
                    hovermode='x unified',
                    margin=dict(t=40, b=20, l=20, r=20)
                )
                fig.update_traces(
                    hovertemplate='<b>%{x}</b><br>Chamados: %{y}'
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
            except Exception as e:
                st.warning(f"Erro no gráfico de prioridades: {str(e)}")
            
            # Gráfico Boxplot - Tempo de Atendimento
            try:
                # Filtrando outliers extremos
                q_low = df['Tempo de Atendimento (dias)'].quantile(0.01)
                q_hi  = df['Tempo de Atendimento (dias)'].quantile(0.99)
                boxplot_df = df[(df['Tempo de Atendimento (dias)'] >= q_low) & 
                            (df['Tempo de Atendimento (dias)'] <= q_hi)]
                
                # Criando o gráfico
                fig = px.box(
                    boxplot_df,
                    x='Categoria 1',
                    y='Tempo de Atendimento (dias)',
                    color='Prioridade',
                    title="Distribuição de Tempo por Categoria",
                    template=template,
                    labels={'Tempo de Atendimento (dias)': 'Dias', 'Categoria 1': 'Categoria'},
                    color_discrete_sequence=px.colors.qualitative.Pastel,
                    height=500
                )
                
                # Melhorando a aparência
                fig.update_layout(
                    xaxis_title='Categoria',
                    yaxis_title='Tempo de Atendimento (dias)',
                    legend_title='Prioridade',
                    hovermode='x unified',
                    margin=dict(t=40, b=20, l=20, r=20)
                )
                fig.update_traces(
                    hovertemplate='<b>%{x}</b><br>Dias: %{y}'
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
            except Exception as e:
                st.warning(f"Erro no gráfico boxplot: {str(e)}")
        
    with tabs[3]:
        col1, col2 = st.columns(2)
        
        with col1:
            # Gráfico de Barras - Chamados por Prioridade
            try:
                prioridade_data = df.groupby('Prioridade', observed=True).size().reset_index(name='Count')
                
                fig = px.bar(
                    prioridade_data,
                    x='Prioridade',
                    y='Count',
                    title='Distribuição de Chamados por Prioridade',
                    template=template,
                    text_auto=True,
                    color='Prioridade',
                    color_discrete_sequence=px.colors.qualitative.Pastel,
                    height=400
                )
                
                # Melhorando a aparência
                fig.update_layout(
                    xaxis_title='Prioridade',
                    yaxis_title='Quantidade de Chamados',
                    hovermode='x unified',
                    margin=dict(t=40, b=20, l=20, r=20)
                )
                fig.update_traces(
                    textposition='outside',
                    hovertemplate='<b>%{x}</b><br>Chamados: %{y}'
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
            except Exception as e:
                st.warning(f"Erro no gráfico de prioridades: {str(e)}")
            
            # Gráfico Comparativo SLA
            try:
                sla_comparativo = df.groupby(['Categoria 2', 'Atraso no serviço']).size().unstack().fillna(0)
                sla_comparativo = sla_comparativo.nlargest(10, columns=['Não', 'Sim'])
                
                fig = px.bar(
                    sla_comparativo,
                    barmode='group',
                    title='Comparativo SLA por Categoria (Top 10)',
                    template=template,
                    labels={'value': 'Quantidade', 'Atraso no serviço': 'Dentro do SLA'},
                    color_discrete_sequence=px.colors.qualitative.Pastel,
                    height=500
                )
                
                # Melhorando a aparência
                fig.update_layout(
                    xaxis_title='Categoria',
                    yaxis_title='Quantidade de Chamados',
                    legend_title='Dentro do SLA',
                    hovermode='x unified',
                    margin=dict(t=40, b=20, l=20, r=20)
                )
                fig.update_traces(
                    hovertemplate='<b>%{x}</b><br>Chamados: %{y}'
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
            except Exception as e:
                st.warning(f"Erro no gráfico comparativo: {str(e)}")
        
        with col2:
            # Gráfico de Dispersão - Tempo vs Prioridade
            try:
                # Filtrando outliers extremos
                q_low = df['Tempo de Atendimento (dias)'].quantile(0.01)
                q_hi  = df['Tempo de Atendimento (dias)'].quantile(0.99)
                scatter_df = df[(df['Tempo de Atendimento (dias)'] >= q_low) & 
                            (df['Tempo de Atendimento (dias)'] <= q_hi)]
                
                fig = px.scatter(
                    scatter_df,
                    x='Tempo de Atendimento (dias)',
                    y='Prioridade',
                    color='Categoria 1',
                    title='Relação Tempo vs Prioridade',
                    template=template,
                    labels={'Tempo de Atendimento (dias)': 'Dias', 'Prioridade': 'Prioridade'},
                    color_discrete_sequence=px.colors.qualitative.Pastel,
                    height=400
                )
                
                # Melhorando a aparência
                fig.update_layout(
                    xaxis_title='Tempo de Atendimento (dias)',
                    yaxis_title='Prioridade',
                    legend_title='Categoria',
                    hovermode='closest',
                    margin=dict(t=40, b=20, l=20, r=20)
                )
                fig.update_traces(
                    marker=dict(size=8, opacity=0.7),
                    hovertemplate='<b>%{y}</b><br>Dias: %{x}'
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
            except Exception as e:
                st.warning(f"Erro no gráfico de dispersão: {str(e)}")
    
    with tabs[4]:
        col1, col2 = st.columns(2)
        
        with col1:
            # Gráfico de Violino - Análise de Tempo de Atendimento
            try:
                # Filtrando outliers extremos
                q_low = df['Tempo de Atendimento (dias)'].quantile(0.01)
                q_hi  = df['Tempo de Atendimento (dias)'].quantile(0.99)
                violin_df = df[(df['Tempo de Atendimento (dias)'] >= q_low) & 
                            (df['Tempo de Atendimento (dias)'] <= q_hi)]
                
                fig = px.violin(
                    violin_df,
                    y='Tempo de Atendimento (dias)',
                    box=True,
                    points="all",
                    title="Distribuição de Tempo de Atendimento",
                    template=template,
                    color_discrete_sequence=['#1f77b4'],
                    height=500
                )
                
                # Melhorando a aparência
                fig.update_layout(
                    yaxis_title='Tempo de Atendimento (dias)',
                    xaxis_title='Distribuição',
                    hovermode='y unified',
                    margin=dict(t=40, b=20, l=20, r=20)
                )
                fig.update_traces(
                    hovertemplate='<b>Tempo:</b> %{y} dias',
                    meanline_visible=True
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
            except Exception as e:
                st.warning(f"Erro no gráfico de violino: {str(e)}")
        
        with col2:
            # Mapa de Calor Temporal
            try:
                # Processando os dados
                df_heatmap = df.copy()  # Cria uma cópia explícita
                df_heatmap['Hora'] = df_heatmap['Data da abertura'].dt.hour
                df_heatmap['Dia da Semana'] = df_heatmap['Data da abertura'].dt.day_name()
                
                # Mapeamento de dias da semana para português
                dias_semana = {
                    'Monday': 'Segunda',
                    'Tuesday': 'Terça',
                    'Wednesday': 'Quarta',
                    'Thursday': 'Quinta',
                    'Friday': 'Sexta',
                    'Saturday': 'Sábado',
                    'Sunday': 'Domingo'
                }
                df_heatmap['Dia da Semana'] = df_heatmap['Dia da Semana'].map(dias_semana)
                
                heatmap_data = df_heatmap.groupby(['Dia da Semana', 'Hora']).size().unstack().fillna(0)
                
                # Ordenando os dias da semana corretamente
                dias_ordem = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo']
                heatmap_data = heatmap_data.reindex(dias_ordem)
                
                fig = px.imshow(
                    heatmap_data,
                    labels=dict(x="Hora do Dia", y="Dia da Semana", color="Chamados"),
                    title="Distribuição de Chamados por Dia/Hora",
                    template=template,
                    color_continuous_scale='Viridis',
                    height=500
                )
                
                # Melhorando a aparência
                fig.update_layout(
                    xaxis_nticks=24,
                    yaxis_nticks=7,
                    coloraxis_colorbar=dict(
                        title="Chamados",
                        thickness=15
                    ),
                    margin=dict(t=40, b=20, l=20, r=20)
                )
                fig.update_traces(
                    hovertemplate='<b>%{y} às %{x}h</b><br>Chamados: %{z}'
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
            except Exception as e:
                st.warning(f"Erro no mapa de calor: {str(e)}")
    
    with tabs[5]:
        # Definindo colunas essenciais e opcionais
        essential_columns = ["Atendimento", "Título do atendimento", "Categoria 1", "Situação"]
        optional_columns = [
            "Categoria 2", "Etapa", "Equipe", "Responsável", "Prioridade",
            "Tempo de Atendimento (dias)", "Data da abertura",
            "Data da previsão de término", "Data de término do atendimento"
        ]
        
        # Verificando colunas disponíveis
        available_columns = [col for col in essential_columns if col in df.columns]
        available_columns += [col for col in optional_columns if col in df.columns]
        
        if not available_columns:
            st.error("Nenhuma coluna relevante encontrada nos dados.")
            st.stop()
        
        # Adicionando seleção de colunas pelo usuário
        with st.expander("⚙️ Configurações da Tabela", expanded=True):
            selected_columns = st.multiselect(
                "Selecione as colunas para exibir:",
                options=available_columns,
                default=available_columns,
                key="table_columns"
            )
            
            # Configurações adicionais
            col1, col2 = st.columns(2)
            with col1:
                sort_column = st.selectbox(
                    "Ordenar por:",
                    options=selected_columns,
                    index=selected_columns.index("Tempo de Atendimento (dias)") if "Tempo de Atendimento (dias)" in selected_columns else 0
                )
            with col2:
                sort_order = st.radio(
                    "Ordem:",
                    ["Descendente", "Ascendente"],
                    index=0,
                    horizontal=True
                )
        
        # Processando os dados
        try:
            # Criando DataFrame filtrado
            display_df = df[selected_columns].copy()
            
            # Ordenação especial para a coluna "Atendimento"
            if "Atendimento" in display_df.columns:
                display_df = display_df.copy()  # Cria uma cópia explícita
                # Extrai o número da string (e.g., "REQ-123" -> 123)
                display_df.loc[:, 'Atendimento_Num'] = (
                    display_df['Atendimento']
                    .str.extract(r'(\d+)')
                    .astype(float)
                )
                # Ordena pelo número extraído se a coluna "Atendimento" for selecionada
                if sort_column == "Atendimento":
                    display_df = display_df.sort_values(
                        'Atendimento_Num',
                        ascending=(sort_order == "Ascendente")
                    )
                else:
                    display_df = display_df.sort_values(
                        sort_column,
                        ascending=(sort_order == "Ascendente")
                    )
                # Remove a coluna auxiliar
                display_df = display_df.drop(columns=['Atendimento_Num'])
            else:
                # Ordenação padrão para outras colunas
                display_df = display_df.sort_values(
                    sort_column,
                    ascending=(sort_order == "Ascendente")
                )
                        
            # Adicionando formatação condicional
            if "Tempo de Atendimento (dias)" in display_df.columns:
                display_df = display_df.copy()  # Cria uma cópia explícita
                display_df['Tempo de Atendimento (dias)'] = pd.to_numeric(display_df['Tempo de Atendimento (dias)'], errors='coerce')
                conditions = [
                    (display_df['Tempo de Atendimento (dias)'] <= 1),
                    (display_df['Tempo de Atendimento (dias)'] <= 3),
                    (display_df['Tempo de Atendimento (dias)'] > 3)
                ]
                colors = ['#4CAF50', '#FFC107', '#F44336']  # Verde, Amarelo, Vermelho
                display_df['Tempo_Cor'] = np.select(conditions, colors, default='#FFFFFF')

            if "Situação" in display_df.columns:
                display_df = display_df.copy()  # Cria uma cópia explícita
                situacao_cores = {
                    'Resolvido': '#4CAF50',
                    'Em andamento': '#FFC107',
                    'Pendente': '#F44336',
                    'Cancelado': '#9E9E9E'
                }
                display_df['Situação_Cor'] = display_df['Situação'].map(situacao_cores).fillna('#FFFFFF')
            
            # Configurações de colunas
            column_config = {}
            date_columns = [col for col in display_df.columns if 'Data' in col]
            for col in date_columns:
                column_config[col] = st.column_config.DatetimeColumn(
                    format="DD/MM/YYYY"
                )
            
            # Exibindo a tabela
            st.dataframe(
                display_df,
                height=600,
                use_container_width=True,
                column_config=column_config,
                column_order=selected_columns,
                hide_index=True
            )
            
            # Adicionando métricas rápidas
            if "Tempo de Atendimento (dias)" in display_df.columns:
                st.markdown("### 📊 Métricas de Tempo de Atendimento")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Média", f"{display_df['Tempo de Atendimento (dias)'].mean():.1f} dias")
                with col2:
                    st.metric("Mediana", f"{display_df['Tempo de Atendimento (dias)'].median():.1f} dias")
                with col3:
                    st.metric("Máximo", f"{display_df['Tempo de Atendimento (dias)'].max():.1f} dias")
            
        except Exception as e:
            st.error(f"Erro ao processar os dados: {str(e)}")

    with tabs[6]:
        st.subheader("📉 Análise de Backlog de Chamados")
        
        # Cálculo do backlog
        df_backlog = df.copy()
        df_backlog['Mês Abertura'] = df_backlog['Data da abertura'].dt.to_period('M').astype(str)
        df_backlog['Mês Fechamento'] = df_backlog['Data de término do atendimento'].dt.to_period('M').astype(str)
        
        # Chamados abertos por mês
        abertos_por_mes = df_backlog.groupby('Mês Abertura').size().reset_index(name='Chamados Abertos')
        
        # Chamados fechados no mesmo mês
        fechados_mesmo_mes = df_backlog[df_backlog['Mês Abertura'] == df_backlog['Mês Fechamento']]
        fechados_mesmo_mes = fechados_mesmo_mes.groupby('Mês Abertura').size().reset_index(name='Chamados Fechados no Mesmo Mês')
        
        # Chamados transferidos para o próximo mês
        transferidos_proximo_mes = df_backlog[df_backlog['Mês Abertura'] != df_backlog['Mês Fechamento']]
        transferidos_proximo_mes = transferidos_proximo_mes.groupby('Mês Abertura').size().reset_index(name='Chamados Transferidos')
        
        # Chamados em aberto (não fechados)
        em_aberto = df_backlog[df_backlog['Data de término do atendimento'].isna()]
        em_aberto_por_mes = em_aberto.groupby('Mês Abertura').size().reset_index(name='Chamados em Aberto')
        
        # Combinando os dados
        backlog_df = abertos_por_mes.merge(fechados_mesmo_mes, on='Mês Abertura', how='left')
        backlog_df = backlog_df.merge(transferidos_proximo_mes, on='Mês Abertura', how='left')
        backlog_df = backlog_df.merge(em_aberto_por_mes, on='Mês Abertura', how='left')
        backlog_df = backlog_df.fillna(0)  # Preenche valores NaN com 0
        
        # Mapeamento dos meses para português
        meses_pt = {
            '1': 'Janeiro',
            '2': 'Fevereiro',
            '3': 'Março',
            '4': 'Abril',
            '5': 'Maio',
            '6': 'Junho',
            '7': 'Julho',
            '8': 'Agosto',
            '9': 'Setembro',
            '10': 'Outubro',
            '11': 'Novembro',
            '12': 'Dezembro'
        }
        
        # Função para converter o formato do mês e ano (ex: '2023-04' -> 'Abril-2023')
        def converter_mes_ano(periodo):
            # Extrai o ano e o mês do período (ex: '2023-04' -> '2023' e '04')
            ano, mes = periodo.split('-')
            # Remove o zero à esquerda do mês (ex: '04' -> '4')
            mes = mes.lstrip('0')
            # Obtém o nome do mês em português
            mes_pt = meses_pt.get(mes, 'Mês Desconhecido')
            # Retorna no formato 'Mês-Ano' (ex: 'Abril-2023')
            return f"{mes_pt}-{ano}"
        
        # Convertendo os números dos meses para nomes em português e adicionando o ano
        backlog_df['Mês Abertura'] = backlog_df['Mês Abertura'].apply(converter_mes_ano)
        
        # Calculando o backlog acumulado
        backlog_df['Backlog Acumulado'] = backlog_df['Chamados Abertos'] - backlog_df['Chamados Fechados no Mesmo Mês']
        backlog_df['Backlog Acumulado'] = backlog_df['Backlog Acumulado'].cumsum()
        
        # Ajustando o backlog acumulado para incluir os chamados em aberto
        backlog_df['Backlog Acumulado'] += backlog_df['Chamados em Aberto']
        
        # Exibindo métricas de backlog
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            st.metric("Total de Chamados Abertos", int(backlog_df['Chamados Abertos'].sum()))
        with col2:
            st.metric("Total de Chamados Fechados no Mesmo Mês", int(backlog_df['Chamados Fechados no Mesmo Mês'].sum()))
        with col3:
            st.metric("Total de Chamados Transferidos", int(backlog_df['Chamados Transferidos'].sum()))
        with col4:
            st.metric("Backlog Acumulado Atual", int(backlog_df['Backlog Acumulado'].iloc[-1]))
        with col5:
            st.metric("Chamados em Aberto Atuais", int(backlog_df['Chamados em Aberto'].sum()))
        
        # Gráfico de Backlog
        fig = go.Figure()
        
        # Adicionando barras para chamados abertos, fechados e transferidos
        fig.add_trace(go.Bar(
            x=backlog_df['Mês Abertura'],
            y=backlog_df['Chamados Abertos'],
            name='Chamados Abertos',
            marker_color='#1f77b4'
        ))
        fig.add_trace(go.Bar(
            x=backlog_df['Mês Abertura'],
            y=backlog_df['Chamados Fechados no Mesmo Mês'],
            name='Chamados Fechados no Mesmo Mês',
            marker_color='#2ca02c'
        ))
        fig.add_trace(go.Bar(
            x=backlog_df['Mês Abertura'],
            y=backlog_df['Chamados Transferidos'],
            name='Chamados Transferidos',
            marker_color='#ff7f0e'
        ))
        fig.add_trace(go.Bar(
            x=backlog_df['Mês Abertura'],
            y=backlog_df['Chamados em Aberto'],
            name='Chamados em Aberto',
            marker_color='#d62728'
        ))
        
        # Adicionando linha para backlog acumulado
        fig.add_trace(go.Scatter(
            x=backlog_df['Mês Abertura'],
            y=backlog_df['Backlog Acumulado'],
            name='Backlog Acumulado',
            mode='lines+markers',
            line=dict(color='#9467bd', width=3),
            marker=dict(size=8, color='#9467bd')
        ))
        
        # Layout do gráfico
        fig.update_layout(
            title="Evolução do Backlog de Chamados",
            xaxis_title="Mês-Ano",
            yaxis_title="Quantidade de Chamados",
            barmode='group',
            template=template,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            hovermode="x unified"
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        # Tabela de detalhes do backlog
        st.subheader("📋 Detalhes do Backlog por Mês-Ano")
        st.dataframe(backlog_df, use_container_width=True)
        
# 📌 Função principal
def main():

    # 📤 Upload de arquivo
    uploaded_file = st.sidebar.file_uploader("📤 Carregar arquivo Excel", type=["xlsx"])
    if not uploaded_file:
        st.info("ℹ️ Por favor, faça upload de um arquivo Excel para iniciar.")
        return
    
    # 📦 Carregar dados
    df = load_data(uploaded_file)
    if df.empty:
        st.error("Erro ao carregar dados. Verifique o formato do arquivo.")
        return
    
    # 🎛️ Aplicar filtros
    date_range, situacoes, categorias, tipos, equipes, responsaveis, prioridades = sidebar_filters(df)
    
    # ⚙️ Filtrar dados
    mask = (
        df["Data da abertura"].between(pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])) &
        df["Situação"].isin(situacoes) &
        df["Categoria 2"].isin(categorias) &
        df["Categoria 1"].isin(tipos) &
        df["Equipe"].isin(equipes) &
        df["Responsável"].isin(responsaveis) &
        df["Prioridade"].isin(prioridades)
    )
    df_filtered = df[mask].copy()
    
    if df_filtered.empty:
        st.warning("⚠️ Nenhum dado encontrado com os filtros selecionados.")
        return
    
    # 📊 Renderizar conteúdo
    create_metrics(df_filtered)
    create_charts(df_filtered, plotly_template)
    
    # 📥 Download dos dados filtrados
    st.sidebar.download_button(
        label="📥 Baixar Dados Filtrados",
        data=df_filtered.to_csv(index=False).encode('utf-8'),
        file_name='dados_filtrados.csv',
        mime='text/csv'
    )

if __name__ == "__main__":
    main()
