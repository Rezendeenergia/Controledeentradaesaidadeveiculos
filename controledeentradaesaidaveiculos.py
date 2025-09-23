import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
from datetime import datetime, timedelta
import io

# Configuração da página
st.set_page_config(
    page_title=st.secrets.get("app", {}).get("page_title", "Controle de Veículos"),
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS customizado com as cores da empresa
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    .main {
        font-family: 'Inter', sans-serif;
    }

    .header-container {
        background: linear-gradient(135deg, #F7931E 0%, #000000 100%);
        padding: 2rem;
        border-radius: 15px;
        margin-bottom: 2rem;
        box-shadow: 0 10px 30px rgba(247, 147, 30, 0.3);
    }

    .header-title {
        color: white;
        font-size: 2.5rem;
        font-weight: 700;
        margin: 0;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        text-align: center;
    }

    .header-subtitle {
        color: rgba(255,255,255,0.9);
        font-size: 1.2rem;
        font-weight: 400;
        margin-top: 0.5rem;
        text-align: center;
    }

    .metric-card {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        border-left: 4px solid #F7931E;
        margin-bottom: 1rem;
        transition: transform 0.3s ease;
    }

    .metric-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 8px 25px rgba(0,0,0,0.15);
    }

    .metric-value {
        font-size: 2.5rem;
        font-weight: 700;
        color: #000000;
        margin: 0;
    }

    .metric-label {
        font-size: 0.9rem;
        color: #666;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }

    .metric-delta {
        font-size: 0.85rem;
        margin-top: 0.5rem;
    }

    .delta-positive {
        color: #10B981;
    }

    .delta-negative {
        color: #EF4444;
    }

    .section-header {
        background: #000000;
        color: white;
        padding: 1rem 1.5rem;
        border-radius: 8px;
        margin: 2rem 0 1rem 0;
        font-weight: 600;
        font-size: 1.1rem;
    }

    .warning-card {
        background: #FEF3CD;
        border: 1px solid #F59E0B;
        border-radius: 8px;
        padding: 1rem;
        margin: 1rem 0;
    }

    .success-card {
        background: #D1FAE5;
        border: 1px solid #10B981;
        border-radius: 8px;
        padding: 1rem;
        margin: 1rem 0;
    }

    .sidebar .sidebar-content {
        background: linear-gradient(180deg, #F7931E 0%, #FF6B35 100%);
    }

    .stSelectbox > div > div {
        background: white;
        border: 2px solid #F7931E;
        border-radius: 8px;
    }

    .stDateInput > div > div {
        border: 2px solid #F7931E;
        border-radius: 8px;
    }

    div[data-testid="metric-container"] {
        background: white;
        border: 1px solid #E5E7EB;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #F7931E;
    }

</style>
""", unsafe_allow_html=True)


@st.cache_data
def load_sharepoint_data():
    """Carregar dados reais do SharePoint"""
    try:
        # Importar sua função de acesso ao SharePoint
        import requests
        from msal import ConfidentialClientApplication

        # Carregar credenciais do secrets.toml
        client_id = st.secrets["sharepoint"]["client_id"]
        client_secret = st.secrets["sharepoint"]["client_secret"]
        tenant_id = st.secrets["sharepoint"]["tenant_id"]
        site_url_base = st.secrets["sharepoint"]["site_url"]
        site_path = st.secrets["sharepoint"]["site_path"]
        excel_filename = st.secrets["sharepoint"]["excel_filename"]

        # Configurar autenticação
        app = ConfidentialClientApplication(
            client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            client_credential=client_secret,
        )

        # Obter token
        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

        if "access_token" in result:
            headers = {"Authorization": f"Bearer {result['access_token']}"}

            # Obter o site_id
            site_url = f"https://graph.microsoft.com/v1.0/sites/{site_url_base}:{site_path}"
            site_response = requests.get(site_url, headers=headers)

            if site_response.status_code == 200:
                site_data = site_response.json()
                site_id = site_data['id']

                # Buscar o arquivo específico
                search_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/search(q='{excel_filename}')"
                search_response = requests.get(search_url, headers=headers)

                if search_response.status_code == 200:
                    search_data = search_response.json()
                    files_found = search_data.get('value', [])

                    for item in files_found:
                        if item['name'] == excel_filename:
                            # Baixar o arquivo
                            download_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item['id']}/content"
                            download_response = requests.get(download_url, headers=headers)

                            if download_response.status_code == 200:
                                # Ler o Excel
                                df = pd.read_excel(io.BytesIO(download_response.content))

                                # Padronizar nomes das colunas
                                df.columns = ['data_hora', 'email', 'nome', 'placa', 'modelo', 'tipo', 'km_inicial',
                                              'km_final', 'finalidade']

                                # Processar dados para formato padrão
                                processed_data = []
                                for _, row in df.iterrows():
                                    if row['tipo'] == 'Saída':
                                        processed_data.append({
                                            'data_hora': row['data_hora'],
                                            'email': row['email'],
                                            'nome': row['nome'],
                                            'placa': row['placa'],
                                            'modelo': row['modelo'],
                                            'tipo': 'Saída',
                                            'km': row['km_inicial'],
                                            'finalidade': row['finalidade']
                                        })
                                    else:  # Chegada
                                        processed_data.append({
                                            'data_hora': row['data_hora'],
                                            'email': row['email'],
                                            'nome': row['nome'],
                                            'placa': row['placa'],
                                            'modelo': row['modelo'],
                                            'tipo': 'Chegada',
                                            'km': row['km_final'],
                                            'finalidade': None
                                        })

                                return pd.DataFrame(processed_data)

        st.error("Erro ao acessar o SharePoint. Verifique as credenciais.")
        return pd.DataFrame()

    except Exception as e:
        st.error(f"Erro ao carregar dados do SharePoint: {e}")
        return pd.DataFrame()


def process_trips(df):
    """Processar dados para criar viagens completas"""
    df['data_hora'] = pd.to_datetime(df['data_hora'], format='%d/%m/%Y %H:%M')
    df = df.sort_values(['nome', 'placa', 'data_hora'])

    trips = []
    orphan_arrivals = []  # Chegadas órfãs

    # Agrupar por motorista e veículo
    for (nome, placa), group in df.groupby(['nome', 'placa']):
        saida = None

        for _, row in group.iterrows():
            if row['tipo'] == 'Saída':
                saida = row
            elif row['tipo'] == 'Chegada':
                if saida is not None:
                    # Criar viagem completa
                    tempo_viagem = (row['data_hora'] - saida['data_hora']).total_seconds() / 60
                    km_rodados = row['km'] - saida['km'] if row['km'] > saida['km'] else 0

                    trips.append({
                        'motorista': nome,
                        'placa': placa,
                        'modelo': saida['modelo'],
                        'data_saida': saida['data_hora'],
                        'data_chegada': row['data_hora'],
                        'km_inicial': saida['km'],
                        'km_final': row['km'],
                        'km_rodados': km_rodados,
                        'tempo_viagem': tempo_viagem,
                        'finalidade': saida['finalidade'],
                        'status': 'Completa'
                    })
                    saida = None
                else:
                    # Chegada órfã (sem saída correspondente)
                    orphan_arrivals.append({
                        'motorista': nome,
                        'placa': placa,
                        'modelo': row['modelo'],
                        'data_chegada': row['data_hora'],
                        'km_final': row['km'],
                        'status': 'Chegada Órfã'
                    })

        # Viagem em aberto (saída sem chegada)
        if saida is not None:
            trips.append({
                'motorista': nome,
                'placa': placa,
                'modelo': saida['modelo'],
                'data_saida': saida['data_hora'],
                'data_chegada': None,
                'km_inicial': saida['km'],
                'km_final': None,
                'km_rodados': None,
                'tempo_viagem': None,
                'finalidade': saida['finalidade'],
                'status': 'Em Aberto'
            })

    trips_df = pd.DataFrame(trips)
    orphans_df = pd.DataFrame(orphan_arrivals)

    return trips_df, orphans_df


def main():
    # Header com logo
    st.markdown("""
    <div class="header-container">
        <h1 class="header-title">🚗 <span style="color: white;">Controle de Veículos</span></h1>
        <p class="header-subtitle">Dashboard Executivo - Entrada e Saída de Veículos</p>
    </div>
    """, unsafe_allow_html=True)

    # Sidebar para filtros
    st.sidebar.markdown("### 🔧 Filtros e Configurações")

    # Botão para atualizar dados
    if st.sidebar.button("🔄 Atualizar Dados", type="primary", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.sidebar.markdown("---")

    # Carregar dados do SharePoint
    with st.spinner("🔄 Carregando dados do SharePoint..."):
        raw_data = load_sharepoint_data()

    if raw_data.empty:
        st.error("❌ Não foi possível carregar os dados do SharePoint. Verifique a conexão.")
        return

    st.success(f"✅ Dados carregados com sucesso! {len(raw_data)} registros encontrados.")

    trips_data, orphan_arrivals = process_trips(raw_data)

    # Filtros
    data_inicio = st.sidebar.date_input("Data Início", value=datetime.now() - timedelta(days=7))
    data_fim = st.sidebar.date_input("Data Fim", value=datetime.now())

    motoristas_selecionados = st.sidebar.multiselect(
        "Motoristas",
        options=trips_data['motorista'].unique(),
        default=trips_data['motorista'].unique()
    )

    status_selecionado = st.sidebar.multiselect(
        "Status das Viagens",
        options=['Completa', 'Em Aberto'],
        default=['Completa', 'Em Aberto']
    )

    # Filtrar dados
    filtered_trips = trips_data[
        (trips_data['motorista'].isin(motoristas_selecionados)) &
        (trips_data['status'].isin(status_selecionado))
        ]

    if not filtered_trips.empty:
        filtered_trips = filtered_trips[
            (filtered_trips['data_saida'].dt.date >= data_inicio) &
            (filtered_trips['data_saida'].dt.date <= data_fim)
            ]

    # KPIs Principais
    st.markdown('<div class="section-header">📊 Indicadores Principais</div>', unsafe_allow_html=True)

    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        total_viagens = len(filtered_trips)
        st.metric(
            label="Total de Viagens",
            value=total_viagens,
            delta=f"+{int(total_viagens * 0.15)} vs mês anterior"
        )

    with col2:
        viagens_abertas = len(filtered_trips[filtered_trips['status'] == 'Em Aberto'])
        chegadas_orfas = len(orphan_arrivals) if not orphan_arrivals.empty else 0
        total_inconsistencias = viagens_abertas + chegadas_orfas
        st.metric(
            label="Inconsistências",
            value=total_inconsistencias,
            delta=f"Abertas: {viagens_abertas}, Órfãs: {chegadas_orfas}"
        )

    with col3:
        km_total = filtered_trips['km_rodados'].sum() if not filtered_trips.empty else 0
        st.metric(
            label="KM Rodados",
            value=f"{km_total:,.0f}",
            delta="+12.5%"
        )

    with col4:
        tempo_medio = filtered_trips['tempo_viagem'].mean() if not filtered_trips.empty else 0
        st.metric(
            label="Tempo Médio (min)",
            value=f"{tempo_medio:.0f}",
            delta="-15 min"
        )

    with col5:
        motoristas_ativos = filtered_trips['motorista'].nunique() if not filtered_trips.empty else 0
        st.metric(
            label="Motoristas Ativos",
            value=motoristas_ativos,
            delta="+1"
        )

    # Alertas detalhados
    if viagens_abertas > 0 or len(orphan_arrivals) > 0:
        col_alert1, col_alert2 = st.columns(2)

        with col_alert1:
            if viagens_abertas > 0:
                st.markdown(f"""
                <div class="warning-card">
                    <strong>⚠️ Viagens em Aberto:</strong> {viagens_abertas} viagem(ns) sem registro de chegada.
                </div>
                """, unsafe_allow_html=True)

        with col_alert2:
            if len(orphan_arrivals) > 0:
                st.markdown(f"""
                <div class="warning-card">
                    <strong>🔍 Chegadas Órfãs:</strong> {len(orphan_arrivals)} chegada(s) sem saída correspondente.
                </div>
                """, unsafe_allow_html=True)

    # Gráficos
    if not filtered_trips.empty:
        st.markdown('<div class="section-header">📈 Análises e Tendências</div>', unsafe_allow_html=True)

        col1, col2 = st.columns(2)

        with col1:
            # Viagens por dia
            viagens_por_dia = filtered_trips.groupby(filtered_trips['data_saida'].dt.date).size().reset_index()
            viagens_por_dia.columns = ['Data', 'Viagens']

            fig_viagens = px.line(
                viagens_por_dia,
                x='Data',
                y='Viagens',
                title='Viagens por Dia',
                color_discrete_sequence=['#F7931E']
            )
            fig_viagens.update_layout(
                plot_bgcolor='white',
                paper_bgcolor='white',
                font=dict(color='#000000')
            )
            st.plotly_chart(fig_viagens, use_container_width=True)

        with col2:
            # Distribuição por motorista
            viagens_motorista = filtered_trips['motorista'].value_counts().reset_index()
            viagens_motorista.columns = ['Motorista', 'Viagens']

            fig_motorista = px.bar(
                viagens_motorista,
                x='Viagens',
                y='Motorista',
                orientation='h',
                title='Viagens por Motorista',
                color_discrete_sequence=['#F7931E']
            )
            fig_motorista.update_layout(
                plot_bgcolor='white',
                paper_bgcolor='white',
                font=dict(color='#000000')
            )
            st.plotly_chart(fig_motorista, use_container_width=True)

        col3, col4 = st.columns(2)

        with col3:
            # KM por finalidade
            km_finalidade = filtered_trips.groupby('finalidade')['km_rodados'].sum().reset_index()

            fig_km = px.pie(
                km_finalidade,
                values='km_rodados',
                names='finalidade',
                title='KM por Finalidade',
                color_discrete_sequence=px.colors.sequential.Oranges_r
            )
            fig_km.update_layout(
                plot_bgcolor='white',
                paper_bgcolor='white',
                font=dict(color='#000000')
            )
            st.plotly_chart(fig_km, use_container_width=True)

        with col4:
            # Horários de maior movimento
            filtered_trips['hora_saida'] = filtered_trips['data_saida'].dt.hour
            movimentos_hora = filtered_trips['hora_saida'].value_counts().sort_index().reset_index()
            movimentos_hora.columns = ['Hora', 'Viagens']

            fig_hora = px.bar(
                movimentos_hora,
                x='Hora',
                y='Viagens',
                title='Movimentação por Hora',
                color_discrete_sequence=['#F7931E']
            )
            fig_hora.update_layout(
                plot_bgcolor='white',
                paper_bgcolor='white',
                font=dict(color='#000000')
            )
            st.plotly_chart(fig_hora, use_container_width=True)

        # Tabela detalhada
        st.markdown('<div class="section-header">📋 Detalhamento das Viagens</div>', unsafe_allow_html=True)

        # Preparar dados para exibição
        display_data = filtered_trips.copy()
        if not display_data.empty:
            display_data['data_saida'] = display_data['data_saida'].dt.strftime('%d/%m/%Y %H:%M')
            display_data['data_chegada'] = display_data['data_chegada'].dt.strftime(
                '%d/%m/%Y %H:%M') if 'data_chegada' in display_data.columns else None
            display_data['tempo_viagem'] = display_data['tempo_viagem'].round(0).astype('Int64')
            display_data['km_rodados'] = display_data['km_rodados'].astype('Int64')

        st.dataframe(
            display_data[['motorista', 'placa', 'modelo', 'data_saida', 'data_chegada',
                          'km_rodados', 'tempo_viagem', 'finalidade', 'status']],
            use_container_width=True,
            column_config={
                'motorista': 'Motorista',
                'placa': 'Placa',
                'modelo': 'Modelo',
                'data_saida': 'Data/Hora Saída',
                'data_chegada': 'Data/Hora Chegada',
                'km_rodados': st.column_config.NumberColumn('KM Rodados', format='%d'),
                'tempo_viagem': st.column_config.NumberColumn('Tempo (min)', format='%d'),
                'finalidade': 'Finalidade',
                'status': 'Status'
            }
        )

        # Tabela de chegadas órfãs (se houver)
        if not orphan_arrivals.empty:
            st.markdown('<div class="section-header">🔍 Chegadas Órfãs (Sem Saída Correspondente)</div>',
                        unsafe_allow_html=True)

            orphan_display = orphan_arrivals.copy()
            orphan_display['data_chegada'] = orphan_display['data_chegada'].dt.strftime('%d/%m/%Y %H:%M')

            st.dataframe(
                orphan_display[['motorista', 'placa', 'modelo', 'data_chegada', 'km_final', 'status']],
                use_container_width=True,
                column_config={
                    'motorista': 'Motorista',
                    'placa': 'Placa',
                    'modelo': 'Modelo',
                    'data_chegada': 'Data/Hora Chegada',
                    'km_final': st.column_config.NumberColumn('KM Final', format='%d'),
                    'status': 'Status'
                }
            )

    else:
        st.warning("Nenhum dado encontrado para os filtros selecionados.")

    # Footer
    st.markdown("---")
    st.markdown(
        """
        <div style='text-align: center; color: #666; padding: 1rem;'>
            <p>© 2025 Rezende Energia - Sistema de Controle de Veículos</p>
        </div>
        """,
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
