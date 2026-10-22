import streamlit as st
import pandas as pd
import plotly.express as px
import os
import glob
import streamlit.components.v1 as components

st.set_page_config(page_title="Análise de Debates Parlamentares", layout="wide", initial_sidebar_state="collapsed")

from idiomas import DICIONARIO

if 'idioma' not in st.session_state:
    st.session_state.idioma = 'PT'

# Atalhos práticos para não escreveres nomes gigantes o tempo todo
lang = st.session_state.idioma
txt = DICIONARIO[lang]

if 'explorar' not in st.session_state:
    st.session_state.explorar = False
if 'ano_temp' not in st.session_state:
    st.session_state.ano_temp = "2026"
if 'ver_global' not in st.session_state:
    st.session_state.ver_global = False
if 'filtro_oradores_tab1' not in st.session_state:
    st.session_state.filtro_oradores_tab1 = ["Todos"]
if 'filtro_oradores_tab3' not in st.session_state:
    st.session_state.filtro_oradores_tab3 = ["Todos"]

import base64

# o icone com o mais (na parte das métricas)
caminho_icone = os.path.join("imagens", "mais.png")

if os.path.exists(caminho_icone):
    with open(caminho_icone, "rb") as image_file:
        img_b64 = base64.b64encode(image_file.read()).decode()
    
    st.markdown(f"""
        <style>
        /* Procura o título do expander e coloca a imagem antes do texto */
        [data-testid="stExpander"] details summary p::before {{
            content: '';
            display: inline-block;
            width: 20px;
            height: 20px; 
            background-image: url('data:image/png;base64,{img_b64}');
            background-size: contain;
            background-repeat: no-repeat;
            background-position: center;
            margin-right: 10px;
            vertical-align: middle;
            margin-top: -2px;
        }}
        </style>
    """, unsafe_allow_html=True)
else:
    st.warning(f"Ícone não encontrado: {caminho_icone}")

# =====================================================================
# 2. CSS CUSTOMIZADO GLOBAL
# =====================================================================
st.markdown("""
    <style>
        .block-container { padding-top: 5rem !important; }
        [data-testid="stSidebarUserContent"] { padding-top: 3rem !important; }

        .main-title {
            font-size: 2.8rem !important; font-weight: 900 !important; letter-spacing: -1.5px !important;
            color: #0f172a !important; margin-top: -1.5rem !important; margin-bottom: 0px !important;
            padding-bottom: 0px !important; line-height: 1.1 !important; font-family: system-ui, -apple-system, sans-serif !important;
        }

        /* BARRA LATERAL */
        [data-testid="stSidebar"] { background-color: #f8fafc !important; border-right: 1px solid #e2e8f0 !important; }
        [data-testid="stSidebar"] h3 { 
            font-size: 1.15rem !important; font-weight: 800 !important; color: #64748b !important; 
            text-transform: uppercase !important; letter-spacing: 1px !important; margin-top: -1.5rem !important; margin-bottom: 1.5rem !important; 
        }
        [data-testid="stSidebar"] .stSelectbox label { font-size: 1.15rem !important; font-weight: 600 !important; color: #0f172a !important; margin-bottom: 0.5rem !important; }
        [data-testid="stSidebar"] div[data-baseweb="select"] > div { background-color: #ffffff !important; border: 1px solid #cbd5e1 !important; border-radius: 6px !important; box-shadow: 0 1px 3px rgba(0,0,0,0.04) !important; transition: all 0.2s ease; }
        [data-testid="stSidebar"] div[data-baseweb="select"] > div:hover { border-color: #94a3b8 !important; }
        /* BOTÕES */
        div.stButton > button[kind="secondary"] { border: 1px solid #cbd5e1 !important; color: #475569 !important; font-weight: 500 !important; font-size: 1.05rem !important; }
        
        div.stButton > button[kind="primary"] {
            background-color: #3B82F6 !important; 
            color: white !important; 
            border: none !important; 
            font-weight: 600 !important; 
            font-size: 1.05rem !important;
            transition: all 0.2s ease !important;
        }
        div.stButton > button[kind="primary"]:hover {
            background-color: #1E3A8A !important; 
        }

        /* SEPARADORES */
        .stTabs [data-baseweb="tab-list"] { gap: 2.5rem; border-bottom: 1px solid #e2e8f0; margin-bottom: 2rem; }
        .stTabs [data-baseweb="tab"] { font-size: 1.1rem !important; font-weight: 500 !important; color: #94a3b8 !important; padding: 10px 0px 15px 0px !important; border-radius: 0 !important; background: transparent !important; border: none !important; }
        .stTabs [data-baseweb="tab"]:hover { color: #3B82F6 !important; }
        .stTabs [data-baseweb="tab"][aria-selected="true"] { color: #1E3A8A !important; font-weight: 700 !important; }
        .stTabs [data-baseweb="tab-highlight"] { background-color: #3B82F6 !important; height: 3px !important; border-radius: 3px 3px 0 0 !important; }
        
        /* FILTROS E MÉTRICAS */
        .main div[data-baseweb="select"] > div { background-color: #ffffff; border-radius: 6px; border: 1px solid #cbd5e1; box-shadow: none; padding: 2px 10px; }
        span[data-baseweb="tag"] { background-color: #f1f5f9 !important; color: #1E3A8A !important; border-radius: 4px !important; font-weight: 600 !important; border: 1px solid #e2e8f0; }
        div[data-testid="metric-container"] { background-color: transparent !important; border: none !important; padding: 10px 0px !important; box-shadow: none !important; border-left: none !important; }
        div[data-testid="metric-container"] label { color: #64748b !important; font-weight: 500 !important; font-size: 1rem !important; }
        div[data-testid="metric-container"] div[data-testid="stMetricValue"] { color: #0f172a !important; font-size: 2.5rem !important; font-weight: 800 !important; letter-spacing: -1px !important; }
    </style>
""", unsafe_allow_html=True)

# Ficheiros dos anos
FICHEIROS_ANO = {
    "2026": "Estado_Da_Nacao_2026.xlsx",
    "2025": "Estado_Da_Nacao_17_julho_2025.xlsx",
    "2024": "Estado_Da_Nacao_17_julho_2024.xlsx",
    "2023": "Estado_Da_Nacao_20_julho_2023.xlsx",
    "2022": "Estado_Da_Nacao_20_julho_2022.xlsx",
    "2021": "Estado_Da_Nacao_21_julho_2021.xlsx"
}

# Passa a ter show_spinner=False
@st.cache_data(show_spinner=False)
def load_data(file_name):
    df = pd.read_excel(file_name)
    return df


# PÁGINA INICIAL - qd abro o site (com uma explicação e tal)
if not st.session_state.explorar:

    col_vazia, col_lang = st.columns([8, 2])
    with col_lang:
        # O seletor correto atualizado (stColumn em vez de column)
        st.markdown("""
        <style>
            div[data-testid="stColumn"]:nth-of-type(2) button {
                border: 2px solid #3B82F6 !important;
                border-radius: 30px !important;
                color: #1E3A8A !important;
                background-color: #eff6ff !important;
                font-weight: 700 !important;
                font-size: 1.1rem !important;
                box-shadow: 0 4px 6px rgba(59, 130, 246, 0.15) !important;
                transition: all 0.3s ease !important;
            }
            div[data-testid="stColumn"]:nth-of-type(2) button:hover {
                background-color: #3B82F6 !important;
                color: white !important;
                transform: translateY(-2px) !important;
                box-shadow: 0 6px 12px rgba(59, 130, 246, 0.25) !important;
            }
        </style>
        """, unsafe_allow_html=True)
        
        texto_botao_lang = "🇺🇸 English" if lang == 'PT' else "🇵🇹 Português"
        if st.button(texto_botao_lang, use_container_width=True):
            st.session_state.idioma = 'EN' if lang == 'PT' else 'PT'
            st.rerun()

    # Ajuste milimétrico para cada idioma
    margem_subtitulo = "-15px" if lang == 'PT' else "50px"

    # Nota o 'f' antes das aspas para permitir a injeção da variável
    st.markdown(f"""
    <style>
        .stApp, [data-testid="stAppViewContainer"] {{ background-color: #f1f5f9 !important; }}
        header[data-testid="stHeader"] {{ display: none !important; }}
        .block-container {{
            background: #ffffff !important; border-radius: 24px !important; padding: 4rem 4rem !important;
            box-shadow: 0 20px 40px rgba(15, 23, 42, 0.08) !important; border: 1px solid rgba(226, 232, 240, 0.8) !important;
            max-width: 950px !important; margin-top: 8vh !important; margin-bottom: 8vh !important;
        }}
        div[data-testid="stSelectbox"], div[data-testid="stButton"] {{ max-width: 350px !important; margin: 0 auto !important; }}
        div[data-testid="stButton"] {{ margin-top: 1.5rem !important; }}
        
        .landing-title {{
            font-size: 3.6rem !important; font-weight: 900 !important; letter-spacing: -2px !important; margin: 0px auto 0px auto !important;
            padding: 0px !important; background: linear-gradient(135deg, #0f172a 0%, #3B82F6 100%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent; line-height: 1.1 !important; text-align: center;
        }}
        
        /* Aqui entra a margem dinâmica calculada acima */
        .landing-subtitle {{
            font-size: 1.4rem !important; color: #64748b !important; font-weight: 500 !important; letter-spacing: -0.5px !important;
            margin: {margem_subtitulo} auto 2.5rem auto !important; 
            padding: 0px !important; text-align: center;
        }}
        
        .landing-text {{ color: #475569; font-size: 1.15rem; line-height: 1.8; text-align: justify; text-align-last: center; margin-bottom: 3rem; }}
        h4.select-title {{ color: #0f172a; font-size: 1.2rem; margin-bottom: 12px; margin-top: 1rem; text-align: center; font-weight: 600; font-family: system-ui, -apple-system, sans-serif; }}
        div[data-baseweb="select"] * {{ font-size: 1.15rem !important; }}
        @keyframes pulse {{ 0% {{ box-shadow: 0 0 0 0 rgba(59, 130, 246, 0.4); }} 70% {{ box-shadow: 0 0 0 15px rgba(59, 130, 246, 0); }} 100% {{ box-shadow: 0 0 0 0 rgba(59, 130, 246, 0); }} }}
    </style>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <h1 class="landing-title">{txt['landing_titulo']}</h1>
    <p class="landing-subtitle">{txt['landing_subtitulo']}</p>
    <div class="landing-text">
        {txt['landing_desc']}
    </div>
    <h4 class="select-title">{txt['landing_select']}</h4>
    """, unsafe_allow_html=True)
    
    ano_escolhido = st.selectbox("Ano", list(FICHEIROS_ANO.keys()), index=list(FICHEIROS_ANO.keys()).index(st.session_state.ano_temp), label_visibility="collapsed")
    
    if st.button(txt['landing_btn'], type="primary", use_container_width=True):
        st.session_state.ano_temp = ano_escolhido
        st.session_state.explorar = True
        st.rerun()

    # LOGOS DA PÁGINA INICIAL
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #94a3b8; font-size: 0.8rem; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 0;'>Apoios e Financiamento</p>", unsafe_allow_html=True)
    
    import base64
    def get_logo_b64(path):
        if os.path.exists(path):
            with open(path, "rb") as f:
                return base64.b64encode(f.read()).decode()
        return None

    b64_cisuc = get_logo_b64("imagens/CISUC.png")
    b64_crai = get_logo_b64("imagens/crai-vertical.png")
    b64_eu = get_logo_b64("imagens/Uniֶao_Europeia.jpg")
    b64_logos = get_logo_b64("imagens/logos_.png")

    html_logos = "<div style='display: flex; justify-content: center; align-items: center; flex-wrap: wrap; gap: 10px; margin-top: 10px; margin-bottom: 20px;'>"
    
    if b64_cisuc:
        html_logos += f"<img src='data:image/png;base64,{b64_cisuc}' style='height: 12px; width: auto; object-fit: contain; opacity: 0.9;'>"
    if b64_crai:
        html_logos += f"<img src='data:image/png;base64,{b64_crai}' style='height: 38px; width: auto; object-fit: contain; opacity: 0.9;'>"
    if b64_eu:
        html_logos += f"<img src='data:image/png;base64,{b64_eu}' style='height: 25px; width: auto; object-fit: contain; opacity: 0.9;'>"
    if b64_logos:
        html_logos += f"<img src='data:image/png;base64,{b64_logos}' style='height: 50px; width: auto; object-fit: contain; opacity: 0.9;'>"
    
    html_logos += "</div>"

    st.markdown(html_logos, unsafe_allow_html=True)

# PÁGINA APÓS A PÁGINA INICIAL
else:
    # --- BARRA LATERAL ---
    # --- BARRA LATERAL ---
    st.sidebar.markdown(f"### {txt['sidebar_contexto']}")
    
    # 1. ESCOLHA DO ANO
    index_ano_atual = list(FICHEIROS_ANO.keys()).index(st.session_state.ano_temp)
    ano_selecionado = st.sidebar.selectbox(txt['sidebar_escolha_ano'], list(FICHEIROS_ANO.keys()), index=index_ano_atual)
    
    # Se o utilizador mudar de ano enquanto está no modo global, sai do modo global automaticamente
    if ano_selecionado != st.session_state.ano_temp:
        st.session_state.ano_temp = ano_selecionado
        st.session_state.ver_global = False 
        st.rerun()
    
    st.sidebar.markdown("<br>", unsafe_allow_html=True)
    
    # 2. O BOTÃO COM A COMPARAÇÃO DOS ANOS
    texto_botao = txt['sidebar_btn_individual'] if st.session_state.ver_global else txt['sidebar_btn_comparacao']
    if st.sidebar.button(texto_botao, use_container_width=True, type="primary" if not st.session_state.ver_global else "secondary"):
        st.session_state.ver_global = not st.session_state.ver_global
        st.rerun()
        
    st.sidebar.markdown("<br><hr style='border-top: 1px solid #e2e8f0;'><br>", unsafe_allow_html=True)
    
    # 3. BOTÃO DE VOLTAR AO INÍCIO
    if st.sidebar.button(txt['btn_voltar_inicio'], type="secondary", use_container_width=True):
        st.session_state.explorar = False
        st.rerun()

    st.sidebar.markdown("<br>" * 6, unsafe_allow_html=True)
    
    # Texto do rodapé com link para a página do parlamento
    st.sidebar.markdown(f"""
        <div style='text-align: center; padding: 15px 10px 5px 10px; border-top: 1px solid #e2e8f0;'>
            <p style='color: #94a3b8; font-size: 0.75rem; font-weight: 600; margin-bottom: 3px; text-transform: uppercase; letter-spacing: 0.5px;'>{txt['sidebar_fonte_titulo']}</p>
            <a href='https://www.parlamento.pt/DAR/Paginas/DAR1Serie.aspx' target='_blank' style='color: #64748b; font-size: 0.8rem; text-decoration: none; font-weight: 500;'>
                {txt['sidebar_fonte_desc']}
            </a>
        </div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <h1 class="landing-title">{txt['titulo_app']}</h1>
    <p class="landing-subtitle">{txt['subtitulo']} </p>
    """, unsafe_allow_html=True)

    ficheiro_a_carregar = FICHEIROS_ANO[ano_selecionado]

    try:
        with st.spinner(txt.get('spinner_individual', 'A processar os dados...')):
            df = load_data(ficheiro_a_carregar)
        
        if 'Speaker' in df.columns:
            import re
            PARTIDOS_VALIDOS = {'PS', 'PSD', 'CH', 'IL', 'PCP', 'BE', 'PAN', 'L', 'CDS-PP', 'JPP', 'PEV'}
            
            def separar_multiplos_oradores(nome):
                nome_str = str(nome).strip()
                if nome_str in ['nan', 'None', '', 'Ações']: return [nome_str]
                if not nome_str.lower().startswith('vozes'): nome_str = re.sub(r'\s*\([^)]*\)', '', nome_str).strip()

                texto_limpo = re.sub(r'^Vozes\s*(da|do|de|das|dos)?\s*\(?', '', nome_str, flags=re.IGNORECASE)
                texto_limpo = re.sub(r'\)$', '', texto_limpo).strip()
                
                if texto_limpo in PARTIDOS_VALIDOS: return [texto_limpo]
                
                if ',' in texto_limpo or ' e ' in texto_limpo:
                    tmp = re.sub(r'\s+e\s+do\s+', '|', texto_limpo)
                    tmp = re.sub(r'\s+e\s+da\s+', '|', tmp)
                    tmp = re.sub(r',\s*do\s+', '|', tmp)
                    tmp = re.sub(r',\s*da\s+', '|', tmp)
                    tmp = re.sub(r'\s+e\s+', '|', tmp)
                    tmp = re.sub(r',\s+', '|', tmp)
                    
                    partes = [p.strip() for p in tmp.split('|') if p.strip()]
                    if all(p in PARTIDOS_VALIDOS for p in partes) and len(partes) > 1: return partes
                return [nome_str]
                
            df['Speaker'] = df['Speaker'].apply(separar_multiplos_oradores)
            df = df.explode('Speaker')
            df = df.reset_index(drop=True)

        if 'Tipo' in df.columns: df['Tipo'] = df['Tipo'].astype(str).str.strip().str.lower()
            
        if 'acoes_simples' in df.columns:
            def mapear_acao(val):
                v = str(val).strip().lower()
                if 'aplausos' in v: return 'Ação: Aplausos'
                if 'protestos' in v: return 'Ação: Protestos'
                return 'Ação: Outros' 
            df['acao_agrupada'] = df['acoes_simples'].apply(mapear_acao)
            
        col_ato = 'speech_act_gerado' if 'speech_act_gerado' in df.columns else ('trueLabel' if 'trueLabel' in df.columns else None)
        if col_ato:
            df[col_ato] = df[col_ato].astype(str).replace(['nan', 'NaN', 'None', 'none', '', ' '], 'Ação: Outros')
            df[col_ato] = df[col_ato].fillna('Ação: Outros')
            if 'Tipo' in df.columns:
                df.loc[df['Tipo'] == 'ação', col_ato] = 'Ação: Outros'
                if 'acoes_simples' in df.columns:
                    mask_acao = (df['Tipo'] == 'ação')
                    df.loc[mask_acao, col_ato] = df.loc[mask_acao, 'acao_agrupada']
            df.loc[df[col_ato] == 'Ações', col_ato] = 'Ação: Outros'

        if lang == 'EN':
            # Traduz os valores dentro da coluna de comportamento
            if col_ato in df.columns:
                df[col_ato] = df[col_ato].replace({
                    'Cooperação': 'Cooperation',
                    'Conflito': 'Conflict',
                    'Ação: Aplausos': 'Action: Applause',
                    'Ação: Protestos': 'Action: Protests',
                    'Ação: Outros': 'Action: Others'
                })
                    
                # E atualizas as variáveis que usas para filtrar/ordens
            categorias_validas = ["Cooperation", "Conflict", "Action: Applause", "Action: Protests", "Action: Others"]
        else:
            categorias_validas = ["Cooperação", "Conflito", "Ação: Aplausos", "Ação: Protestos", "Ação: Outros"]

        # Funções de Filtros
        def atualizar_filtro_tab1():
            selecionados = st.session_state["filtro_oradores_tab1"]
            if any(x in selecionados for x in ["Todos", "All"]) and len(selecionados) > 1:
                if selecionados[-1] in ["Todos", "All"]: 
                    st.session_state["filtro_oradores_tab1"] = [selecionados[-1]]
                else: 
                    st.session_state["filtro_oradores_tab1"] = [s for s in selecionados if s not in ["Todos", "All"]]

        def atualizar_filtro_tab3():
            selecionados = st.session_state["filtro_oradores_tab3"]
            if any(x in selecionados for x in ["Todos", "All"]) and len(selecionados) > 1:
                if selecionados[-1] in ["Todos", "All"]: 
                    st.session_state["filtro_oradores_tab3"] = [selecionados[-1]]
                else: 
                    st.session_state["filtro_oradores_tab3"] = [s for s in selecionados if s not in ["Todos", "All"]]

        if lang == 'EN':
            mapa_cores = {
                "Cooperation": "#10b981",      # Verde Esmeralda
                "Conflict": "#ef4444",         # Vermelho Forte
                "Action: Applause": "#3B82F6", # Laranja (Azul no teu hex original)
                "Action: Protests": "#c084fc", # Roxo
                "Action: Others": "#94a3b8"    # Cinzento
            }
        else:
            mapa_cores = {
                "Cooperação": "#10b981",
                "Conflito": "#ef4444",
                "Ação: Aplausos": "#3B82F6",
                "Ação: Protestos": "#c084fc",
                "Ação: Outros": "#94a3b8"
            }

        # DECISÃO DE NAVEGAÇÃO: MODO GLOBAL vs MODO INDIVIDUAL
        if st.session_state.ver_global:
            
            @st.cache_data(show_spinner=False)
            def carregar_todos_os_anos():
                    lista_dfs = []
                    for ano, ficheiro in FICHEIROS_ANO.items():
                        if os.path.exists(ficheiro):
                            tmp_df = pd.read_excel(ficheiro)
                            tmp_df['Ano_Debate'] = str(ano)
                            
                            if 'Speaker' in tmp_df.columns:
                                tmp_df['Speaker'] = tmp_df['Speaker'].apply(separar_multiplos_oradores)
                                tmp_df = tmp_df.explode('Speaker')
                                tmp_df = tmp_df.reset_index(drop=True)
                                
                            lista_dfs.append(tmp_df)
                            
                    if lista_dfs:
                        df_concat = pd.concat(lista_dfs, ignore_index=True)
                        df_concat = df_concat.sort_values(by='Ano_Debate')
                        return df_concat
                    return pd.DataFrame()

            with st.spinner(txt.get('spinner_global', 'A analisar todos os anos...')):
                df_global = carregar_todos_os_anos()

            if not df_global.empty:
                col_g1, col_g2 = st.columns(2)
                with col_g1:
                    st.markdown(f"<h5 style='text-align: center; color: #0f172a;'>{txt['global_tit_intervencoes']}</h5>", unsafe_allow_html=True)
                    
                    tons_de_azul = ['#bfdbfe', '#93c5fd', '#60a5fa', '#3b82f6', '#2563eb', '#1d4ed8']
                    
                    fig_anos = px.histogram(
                        df_global, 
                        x='Ano_Debate', 
                        color='Ano_Debate', 
                        template="plotly_white",
                        color_discrete_sequence=tons_de_azul
                    )
                    fig_anos.update_layout(showlegend=False, xaxis={'title': txt['eixo_x_ano']}, yaxis={'title': txt['eixo_x_int']})
                    
                    fig_anos.update_traces(
                        hovertemplate=f"<b>{txt['eixo_x_ano']}:</b> %{{x}}<br><b>{txt['eixo_y_contagem']}:</b> %{{y}} {txt['hover_intervencoes']}<extra></extra>"
                    )
                    
                    st.plotly_chart(fig_anos, use_container_width=True)
                    
                    # BOTÃO: MAIS MÉTRICAS (Volume e Duração)
                    with st.expander(txt['expander_volume']):
                        df_tabela = df_global.copy()
                        if 'Tipo' in df_tabela.columns:
                            mask_acao_tab = df_tabela['Tipo'].astype(str).str.strip().str.lower() == 'ação'
                            df_tabela['Categoria_Tabela'] = txt['col_falas']
                            df_tabela.loc[mask_acao_tab, 'Categoria_Tabela'] = txt['col_acoes']
                        else:
                            df_tabela['Categoria_Tabela'] = txt['col_falas']
                            
                        tabela_stats = df_tabela.groupby(['Ano_Debate', 'Categoria_Tabela']).size().unstack(fill_value=0)
                        
                        if txt['col_falas'] not in tabela_stats: tabela_stats[txt['col_falas']] = 0
                        if txt['col_acoes'] not in tabela_stats: tabela_stats[txt['col_acoes']] = 0

                        valores_manuais = {
                            '2021': {txt['col_dialogos']: 1, txt['col_duracao']: '4h21'},
                            '2022': {txt['col_dialogos']: 3, txt['col_duracao']: '4h48'},
                            '2023': {txt['col_dialogos']: 6, txt['col_duracao']: '4h22'},
                            '2024': {txt['col_dialogos']: 11, txt['col_duracao']: '5h16'},
                            '2025': {txt['col_dialogos']: 12, txt['col_duracao']: '5h01'},
                            '2026': {txt['col_dialogos']: 7, txt['col_duracao']: '4h49'}
                        }
                        
                        df_manuais = pd.DataFrame.from_dict(valores_manuais, orient='index')
                        
                        if not tabela_stats.empty:
                            tabela_final = tabela_stats.join(df_manuais)
                            tabela_final.index.name = txt['eixo_x_ano']
                            tabela_final = tabela_final.reset_index()
                            
                            st.dataframe(
                                tabela_final[[txt['eixo_x_ano'], txt['col_falas'], txt['col_acoes'], txt['col_dialogos'], txt['col_duracao']]], 
                                use_container_width=True, 
                                hide_index=True
                            )
                    
                with col_g2:
                    st.markdown(f"<h5 style='text-align: center; color: #0f172a;'>{txt['global_tit_comportamento']}</h5>", unsafe_allow_html=True)
                    col_ato_global = 'speech_act_gerado' if 'speech_act_gerado' in df_global.columns else ('trueLabel' if 'trueLabel' in df_global.columns else None)
                    
                    if col_ato_global:
                        # (O teu código de substituição para 'Ação: Outros' que já está correto continua aqui...)
                        df_global[col_ato_global] = df_global[col_ato_global].astype(str).replace(
                            ['nan', 'NaN', 'None', 'none', '', ' ', 'FALHA_CLASSIFICACAO', 'Outros'], 'Ação: Outros'
                        )

                        if 'Tipo' in df_global.columns: 
                            df_global['Tipo'] = df_global['Tipo'].astype(str).str.strip().str.lower()
                            
                        if 'acoes_simples' in df_global.columns:
                            def mapear_acao_global(val):
                                v = str(val).strip().lower()
                                if 'aplausos' in v: return 'Ação: Aplausos'
                                if 'protestos' in v: return 'Ação: Protestos'
                                return 'Ação: Outros' 
                            df_global['acao_agrupada'] = df_global['acoes_simples'].apply(mapear_acao_global)
                            
                            mask_acao = (df_global['Tipo'] == 'ação')
                            df_global.loc[mask_acao, col_ato_global] = df_global.loc[mask_acao, 'acao_agrupada']
                            
                        df_global.loc[df_global[col_ato_global] == 'Ações', col_ato_global] = 'Ação: Outros'

                        if lang == 'EN':
                            df_global[col_ato_global] = df_global[col_ato_global].replace({
                                'Cooperação': 'Cooperation', 'Conflito': 'Conflict',
                                'Ação: Aplausos': 'Action: Applause', 'Ação: Protestos': 'Action: Protests',
                                'Ação: Outros': 'Action: Others'
                            })
                            categorias_validas = ["Cooperation", "Conflict", "Action: Applause", "Action: Protests", "Action: Others"]
                        else:
                            categorias_validas = ["Cooperação", "Conflito", "Ação: Aplausos", "Ação: Protestos", "Ação: Outros"]

                        df_grafico = df_global[df_global[col_ato_global].isin(categorias_validas)]

                        fig_comp = px.histogram(
                            df_grafico, 
                            x='Ano_Debate', 
                            color=col_ato_global, 
                            barmode='group', 
                            template="plotly_white",
                            color_discrete_map=mapa_cores,
                            category_orders={col_ato_global: categorias_validas}
                        )
                        
                        fig_comp.update_layout(xaxis={'title': txt['eixo_x_ano']}, yaxis={'title': txt['eixo_y_contagem']}, legend_title_text=txt['legenda_comp'])
                        
                        fig_comp.update_traces(
                            hovertemplate=f"<b>{txt['eixo_x_ano']}:</b> %{{x}}<br><b>{txt['legenda_comp']}:</b> %{{data.name}}<br><b>{txt['eixo_y_contagem']}:</b> %{{y}}<extra></extra>"
                        )
                        
                        st.plotly_chart(fig_comp, use_container_width=True)
                        
                        # MAIS MÉTRICAS (RÁCIOS) - COM CORREÇÃO DOS ZEROS!
                        with st.expander(txt['expander_racios']):
                            if col_ato_global: 
                                tabela_comp = df_global.groupby(['Ano_Debate', col_ato_global]).size().unstack(fill_value=0)
                                
                                # ---> AQUI CORRIGIMOS OS ZEROS PROCURANDO PELA LÍNGUA CERTA <---
                                if lang == 'EN':
                                    chaves_racio = ['Cooperation', 'Conflict', 'Action: Applause', 'Action: Protests']
                                    mapa_nomes_racio = {'Cooperation': '#Coop', 'Conflict': '#Conf', 'Action: Applause': '#App', 'Action: Protests': '#Prot'}
                                else:
                                    chaves_racio = ['Cooperação', 'Conflito', 'Ação: Aplausos', 'Ação: Protestos']
                                    mapa_nomes_racio = {'Cooperação': '#Coop', 'Conflito': '#Conf', 'Ação: Aplausos': '#App', 'Ação: Protestos': '#Prot'}

                                for col in chaves_racio:
                                    if col not in tabela_comp.columns:
                                        tabela_comp[col] = 0
                                        
                                tabela_comp = tabela_comp.rename(columns=mapa_nomes_racio)
                                
                                tabela_comp['Conf Ratio'] = (tabela_comp['#Conf'] / (tabela_comp['#Conf'] + tabela_comp['#Coop'])).fillna(0).round(2)
                                tabela_comp['Prot Ratio'] = (tabela_comp['#Prot'] / (tabela_comp['#Prot'] + tabela_comp['#App'])).fillna(0).round(2)
                                
                                tabela_comp = tabela_comp.reset_index()
                                tabela_comp = tabela_comp.rename(columns={'Ano_Debate': txt['eixo_x_ano']})
                                
                                st.dataframe(
                                    tabela_comp[[txt['eixo_x_ano'], '#Coop', '#Conf', 'Conf Ratio', '#App', '#Prot', 'Prot Ratio']], 
                                    use_container_width=True, 
                                    hide_index=True,
                                    column_config={
                                        "#Coop": st.column_config.NumberColumn(
                                            help=txt['help_coop']
                                        ),
                                        "#Conf": st.column_config.NumberColumn(
                                            help=txt['help_conf']
                                        ),
                                        "Conf Ratio": st.column_config.NumberColumn(
                                            help=txt['help_conf_ratio']
                                        ),
                                        "#App": st.column_config.NumberColumn(
                                            help=txt['help_app']
                                        ),
                                        "#Prot": st.column_config.NumberColumn(
                                            help=txt['help_prot']
                                        ),
                                        "Prot Ratio": st.column_config.NumberColumn(
                                            help=txt['help_prot_ratio']
                                        )
                                    }
                                )

                # RÁCIO DE CONFLITO POR PARTIDO (HEATMAP)
                # RÁCIO DE CONFLITO POR PARTIDO (HEATMAP)
                st.markdown("<br><hr style='border-top: 1px solid #e2e8f0; margin: 40px 0;'><br>", unsafe_allow_html=True)
                st.markdown(f"<h4 style='text-align: center; color: #0f172a; font-weight: 700; font-size: 1.4rem; margin-bottom: 5px;'>{txt['racios_heatmap_tit']}</h4>", unsafe_allow_html=True)
                st.markdown(f"""
                    <p style='text-align: center; color: #64748b; font-size: 1rem; margin-bottom: 25px;'>
                    {txt['racios_heatmap_sub']}
                    </p>
                """, unsafe_allow_html=True)
                
                if col_ato_global and 'Speaker' in df_global.columns:
                    df_ratio = df_global.copy()
                    
                    # 1. Função para padronizar os nomes como estão no paper
                    def padronizar_partido(orador):
                        o = str(orador).upper()
                        if 'MINIST' in o or 'SECRET' in o or 'GOVERNO' in o or 'PRIMEIRO' in o: return 'GOV'
                        if 'PRESIDENT' in o: return 'P'
                        if 'CDS' in o: return 'CDS'
                        if o in ['BE', 'PCP', 'PEV', 'L', 'PS', 'PAN', 'JPP', 'IL', 'PSD', 'CH']: return o
                        return 'OUTROS'
                        
                    df_ratio['Partido_Heatmap'] = df_ratio['Speaker'].apply(padronizar_partido)
                    df_ratio = df_ratio[df_ratio['Partido_Heatmap'] != 'OUTROS'] # Excluir ruído
                    
                    # Definir as chaves consoante o idioma
                    if lang == 'EN':
                        nome_coop = 'Cooperation'
                        nome_conf = 'Conflict'
                    else:
                        nome_coop = 'Cooperação'
                        nome_conf = 'Conflito'

                    # Contar as métricas por Partido e Ano
                    tabela_partidos = df_ratio.groupby(['Ano_Debate', 'Partido_Heatmap', col_ato_global]).size().unstack(fill_value=0)

                    for col in [nome_coop, nome_conf]:
                        if col not in tabela_partidos.columns: 
                            tabela_partidos[col] = 0
                                
                    tabela_partidos = tabela_partidos.reset_index()

                    # Calcular o rácio usando as chaves dinâmicas
                    tabela_partidos['Total_Polar'] = tabela_partidos[nome_conf] + tabela_partidos[nome_coop]
                    tabela_partidos['Conf_Ratio'] = tabela_partidos.apply(
                        lambda row: round(row[nome_conf] / row['Total_Polar'], 1) if row['Total_Polar'] > 0 else float('nan'), axis=1
                    )
                    
                    # Transformar numa matriz
                    tabela_pivot = tabela_partidos.pivot(index='Ano_Debate', columns='Partido_Heatmap', values='Conf_Ratio')
                    tabela_pivot.index.name = txt['eixo_x_ano']
                    
                    # Ordenar as colunas
                    ordem_colunas = ['BE', 'PCP', 'PEV', 'L', 'PS', 'PAN', 'JPP', 'GOV', 'IL', 'PSD', 'CDS', 'CH', 'P']
                    colunas_existentes = [c for c in ordem_colunas if c in tabela_pivot.columns]
                    tabela_pivot = tabela_pivot[colunas_existentes]
                    
                    for col in tabela_pivot.columns:
                        tabela_pivot[col] = tabela_pivot[col].apply(
                            lambda x: f"{x:.1f}" if pd.notnull(x) else "-"
                        )
                    
                    def cor_heatmap(val):
                        if val == '-': return 'background-color: transparent; color: #94a3b8; text-align: center;'
                        
                        v = float(val)
                        hue = int((1.0 - v) * 150) 
                        
                        return f'background-color: hsl({hue}, 84%, 55%); color: #0f172a; text-align: center; font-weight: 700;'
                        
                    try:
                        styler = tabela_pivot.style.map(cor_heatmap)
                    except AttributeError:
                        styler = tabela_pivot.style.applymap(cor_heatmap)
                        
                    st.dataframe(styler, use_container_width=True)
                else:
                    st.warning("Não foi possível gerar os dados para a tabela de rácios por partido.")              

                # ==========================================
                #IMAGEM DA EVOLUÇÃO (COM ZOOM)
                # ==========================================
                st.markdown("<br><hr style='border-top: 1px solid #e2e8f0; margin: 40px 0;'><br>", unsafe_allow_html=True)
                st.markdown(f"<h4 style='text-align: center; color: #0f172a; font-weight: 700; font-size: 1.4rem; margin-bottom: 5px;'>{txt['evolucao_tit']}</h4>", unsafe_allow_html=True)
                st.markdown(f"<p style='text-align: center; color: #64748b; font-size: 1rem; margin-bottom: 25px;'>{txt['evolucao_sub']}</p>", unsafe_allow_html=True)

                # Procurar o ficheiro dinamicamente consoante o idioma
                pasta_imagens = "analise_parlamento"
                
                if lang == 'EN':
                    nome_base = "evolucao_grafico_HORIZONTAL_EN"
                else:
                    nome_base = "evolucao_grafico_HORIZONTAL_PT"
                    
                caminho_evolucao = os.path.join(pasta_imagens, f"{nome_base}.png")
                
                # Fallback: Se estiver em Inglês e a imagem EN não existir, carrega a PT
                if lang == 'EN' and not os.path.exists(caminho_evolucao) and not os.path.exists(os.path.join(pasta_imagens, f"{nome_base}.jpg")):
                    nome_base = "evolucao_grafico_HORIZONTAL_PT"
                    caminho_evolucao = os.path.join(pasta_imagens, f"{nome_base}.png")

                if not os.path.exists(caminho_evolucao):
                    caminho_evolucao = os.path.join(pasta_imagens, f"{nome_base}.jpg")

                if os.path.exists(caminho_evolucao):
                    import base64
                    with open(caminho_evolucao, "rb") as img_file:
                        img_b64 = base64.b64encode(img_file.read()).decode()
                        extensao = caminho_evolucao.split('.')[-1]

                    html_zoom_evo = f"""
                    <div style="display: flex; justify-content: center; width: 100%;">
                        <div id="container-zoom-evo" style="width: 100%; max-width: 1200px; height: 500px; overflow: hidden; border: 1px solid #e2e8f0; border-radius: 12px; background: #ffffff; cursor: grab; position: relative; box-shadow: 0 4px 12px rgba(0,0,0,0.05); user-select: none;">
                            <img id="img-evo" src="data:image/{extensao};base64,{img_b64}" style="width: 100%; height: 100%; object-fit: contain; transform-origin: center; transition: transform 0.05s ease-out; pointer-events: none;" title="Duplo clique para zoom rápido. Roda do rato para ajustar. Clique e arraste para explorar." />
                        </div>
                    </div>

                    <script>
                        const containerEvo = document.getElementById('container-zoom-evo');
                        const imgEvo = document.getElementById('img-evo');
                        let scaleEvo = 1; let pointXEvo = 0; let pointYEvo = 0; let panningEvo = false; let startXEvo = 0; let startYEvo = 0;

                        function constrainEvo() {{
                            const maxTranslateX = (containerEvo.clientWidth * (scaleEvo - 1)) / 2;
                            const maxTranslateY = (containerEvo.clientHeight * (scaleEvo - 1)) / 2;
                            if (scaleEvo <= 1) {{ pointXEvo = 0; pointYEvo = 0; }} 
                            else {{ pointXEvo = Math.max(-maxTranslateX, Math.min(maxTranslateX, pointXEvo)); pointYEvo = Math.max(-maxTranslateY, Math.min(maxTranslateY, pointYEvo)); }}
                        }}

                        function setTransformEvo() {{ constrainEvo(); imgEvo.style.transform = `translate(${{pointXEvo}}px, ${{pointYEvo}}px) scale(${{scaleEvo}})`; }}

                        containerEvo.addEventListener('mousedown', function (e) {{
                            if (scaleEvo > 1) {{ panningEvo = true; startXEvo = e.clientX - pointXEvo; startYEvo = e.clientY - pointYEvo; containerEvo.style.cursor = 'grabbing'; }}
                        }});

                        window.addEventListener('mouseup', function () {{ panningEvo = false; containerEvo.style.cursor = scaleEvo > 1 ? 'grab' : 'default'; }});
                        window.addEventListener('mousemove', function (e) {{ if (!panningEvo) return; pointXEvo = e.clientX - startXEvo; pointYEvo = e.clientY - startYEvo; setTransformEvo(); }});

                        containerEvo.addEventListener('wheel', function (e) {{
                            e.preventDefault();
                            if (e.deltaY < 0) scaleEvo *= 1.2; else scaleEvo /= 1.2;
                            scaleEvo = Math.max(1, Math.min(scaleEvo, 6));
                            if (scaleEvo === 1) {{ pointXEvo = 0; pointYEvo = 0; containerEvo.style.cursor = 'default'; }} 
                            else containerEvo.style.cursor = 'grab';
                            setTransformEvo();
                        }});

                        containerEvo.addEventListener('dblclick', function () {{
                            if (scaleEvo === 1) {{ scaleEvo = 2.5; containerEvo.style.cursor = 'grab'; }} 
                            else {{ scaleEvo = 1; pointXEvo = 0; pointYEvo = 0; containerEvo.style.cursor = 'default'; }}
                            setTransformEvo();
                        }});
                    </script>
                    """
                    components.html(html_zoom_evo, height=560)
                else:
                    st.warning(f"A imagem '{nome_base}' não foi encontrada na pasta '{pasta_imagens}'. Verifica a extensão (.png ou .jpg).")
            else:
                st.warning("Não foram encontrados ficheiros correspondentes aos anos para gerar a vista global.")


        # MODO INDIVIDUAL (ANO ESPECÍFICO)
        else:
            # Cria as abas UMA ÚNICA VEZ
            #tab1, tab2, tab3 = st.tabs(["Análise Estatística", "Fluxo", "Explorador de Dados"])
            tab1, tab2, tab3 = st.tabs([txt['tab_estatistica'], txt['tab_fluxo'], txt['tab_dados']])

            # ABA 1: ESTATÍSTICA
            with tab1:
                if 'Speaker' in df.columns:
                    df['Speaker_Filtro'] = df['Speaker'].astype(str).replace(['nan', 'NaN', 'None', 'none', '', ' '], "Ações")
                                        
                    if 'Tipo' in df.columns:
                        df.loc[df['Tipo'] == 'ação', 'Speaker_Filtro'] = "Ações" 
                        if 'acoes_simples' in df.columns:
                            mask_acao = (df['Tipo'] == 'ação')
                            df.loc[mask_acao, 'Speaker_Filtro'] = "Ações (" + df['acao_agrupada'] + ")"

                    oradores_reais = [s for s in df['Speaker_Filtro'].unique() if not str(s).startswith("Ações")]
                    oradores_originais = sorted(oradores_reais)
                    
                    opcoes_filtro = [txt['opcao_todos']] + oradores_originais
                    
                    c_filtro, c_check = st.columns([3, 1])
                    with c_filtro:
                        oradores_tab1 = st.multiselect(txt['filtrar_oradores'], options=opcoes_filtro, default=[txt['opcao_todos']], key="filtro_oradores_tab1", on_change=atualizar_filtro_tab1)
                    with c_check:
                        st.markdown("""<div style='margin-top: 33px;'></div><style>div[data-testid="stCheckbox"] div[data-baseweb="checkbox"] div[data-checked="true"] { background-color: #3B82F6 !important; border-color: #3B82F6 !important; }</style>""", unsafe_allow_html=True)
                        mostrar_acoes_tab1 = st.checkbox(txt['incluir_acoes'], value=True, key="check_acoes_tab1")

                    mask_acoes = (df['Tipo'] == 'ação') | df['Speaker_Filtro'].astype(str).str.startswith(("Ação", "Ações", "Action"))
                    if col_ato: 
                        mask_acoes = mask_acoes | df[col_ato].astype(str).str.startswith(("Ação", "Action"))

                    # Lógica de filtragem (ABA 1)
                    if len(oradores_tab1) == 0 or any(x in oradores_tab1 for x in ["Todos", "All"]): 
                        mask_oradores = pd.Series(True, index=df.index)
                    else: 
                        # 1. Usar a coluna original para não perder quem fez a ação
                        mask_oradores = df['Speaker'].isin(oradores_tab1)

                    if mostrar_acoes_tab1: 
                        # 2. Filtrar falas e ações, mas apenas do orador
                        df_filtrado_tab1 = df[mask_oradores]
                    else: 
                        # 3. Filtrar apenas falas do orador
                        df_filtrado_tab1 = df[mask_oradores & ~mask_acoes]
                else:
                    df_filtrado_tab1 = df

                # KPIs
                kpi1, kpi2, kpi3, kpi4 = st.columns(4)
                kpi1.metric(txt['kpi_registos'], len(df_filtrado_tab1))
                if 'Speaker' in df.columns: kpi2.metric(txt['kpi_entidades'], df_filtrado_tab1['Speaker_Filtro'].nunique())
                if col_ato:
                    if lang == 'EN':
                        total_coop = len(df_filtrado_tab1[df_filtrado_tab1[col_ato].isin(['Cooperation', 'Action: Applause'])])
                        total_conf = len(df_filtrado_tab1[df_filtrado_tab1[col_ato].isin(['Conflict', 'Action: Protests'])])
                    else:
                        total_coop = len(df_filtrado_tab1[df_filtrado_tab1[col_ato].isin(['Cooperação', 'Ação: Aplausos'])])
                        total_conf = len(df_filtrado_tab1[df_filtrado_tab1[col_ato].isin(['Conflito', 'Ação: Protestos'])])
                    
                    kpi3.metric(txt['kpi_coop'], total_coop)
                    kpi4.metric(txt['kpi_conf'], total_conf)

                st.markdown("<br>", unsafe_allow_html=True)
                col_graf1, col_graf2 = st.columns(2)
                
                # Gráfico de Barras
                # Gráfico de Barras
                with col_graf1:
                    st.markdown(f"<h4 style='text-align: center; color: #0f172a; font-weight: 700; font-size: 1.3rem; margin-bottom: 20px;'>{txt['graf_barras_tit']}</h4>", unsafe_allow_html=True)
                    if 'Speaker' in df.columns:
                        df_barras = df_filtrado_tab1[~df_filtrado_tab1['Speaker_Filtro'].astype(str).str.startswith(("Ação", "Ações", "Action"))]
                        
                        if col_ato:
                            # A ordem da legenda tem de acompanhar a língua
                            if lang == 'EN':
                                ordem_legenda = ["Cooperation", "Conflict", "Action: Applause", "Action: Protests", "Action: Others"]
                            else:
                                ordem_legenda = ["Cooperação", "Conflito", "Ação: Aplausos", "Ação: Protestos", "Ação: Outros"]
                            
                            fig_bar = px.histogram(
                                df_barras, 
                                y='Speaker_Filtro', 
                                color=col_ato, 
                                orientation='h',
                                barmode='stack', 
                                template="plotly_white", 
                                color_discrete_map=mapa_cores,
                                category_orders={col_ato: ordem_legenda}
                            )
                            
                            fig_bar.update_layout(
                                showlegend=False, 
                                yaxis={'categoryorder':'total ascending', 'title': txt['eixo_y_orador']}, 
                                xaxis={'title': txt['eixo_x_int']}
                            )
                            
                            # Atualizar o hover com o dicionário e escapar as chavetas do Plotly
                            fig_bar.update_traces(
                                hovertemplate=f"<b>{txt['hover_orador']}:</b> %{{y}}<br><b>%{{data.name}}:</b> %{{x}} {txt['hover_intervencoes']}<extra></extra>"
                            )
                            
                            st.plotly_chart(fig_bar, use_container_width=True)
                        else:
                            st.warning("Coluna de comportamento não encontrada.")

                # Gráfico Circular
                with col_graf2:
                    st.markdown(f"<h4 style='text-align: center; color: #0f172a; font-weight: 700; font-size: 1.3rem; margin-bottom: 20px;'>{txt['graf_pie_tit']}</h4>", unsafe_allow_html=True)
                    if col_ato:
                        fig_pie = px.pie(
                            df_filtrado_tab1, 
                            names=col_ato, 
                            color=col_ato, 
                            hole=0.4, 
                            template="plotly_white", 
                            color_discrete_map=mapa_cores,
                            category_orders={col_ato: ordem_legenda} # Usa a ordem definida acima
                        )
                        
                        # Atualizar o hover do gráfico circular
                        fig_pie.update_traces(
                            textposition='auto', 
                            textinfo='percent', 
                            hovertemplate=f"<b>%{{label}}</b><br>{txt['hover_perc']}: %{{percent}}<br>{txt['hover_int_cap']}: %{{value}}<extra></extra>"
                        )
                        
                        st.plotly_chart(fig_pie, use_container_width=True)

                # CRONOLOGIA DE TENSÃO
                st.markdown("<br><hr style='border-top: 1px solid #e2e8f0; margin: 30px 0;'><br>", unsafe_allow_html=True)
                st.markdown(f"<h4 style='text-align: center; color: #0f172a; font-weight: 700; font-size: 1.4rem; margin-bottom: 5px;'>{txt['cronologia_tit']}</h4>", unsafe_allow_html=True)
                st.markdown(f"<p style='text-align: center; color: #64748b; font-size: 1rem; margin-bottom: 25px;'>{txt['cronologia_sub']}</p>", unsafe_allow_html=True)
                
                if col_ato:
                    df_cronologia = df_filtrado_tab1.reset_index(drop=True).copy()
                    
                    # Criar blocos fixos de 50 em 50
                    tamanho_fatia = 50
                    
                    df_cronologia['Fase_Debate'] = [
                        f"{x * tamanho_fatia + 1} a {(x + 1) * tamanho_fatia}" 
                        for x in (df_cronologia.index // tamanho_fatia)
                    ]
                    
                    # Garantir que o gráfico respeita a ordem cronológica e não a alfabética
                    ordem_das_fases = df_cronologia['Fase_Debate'].unique().tolist()
                    if lang == 'EN':
                        ordem_legenda = ["Cooperation", "Conflict", "Action: Applause", "Action: Protests", "Action: Others"]
                    else:
                        ordem_legenda = ["Cooperação", "Conflito", "Ação: Aplausos", "Ação: Protestos", "Ação: Outros"]
                    
                    # GERAR O GRÁFICO (agrupado pelas fatias de 50)
                    fig_timeline = px.histogram(
                        df_cronologia, 
                        x='Fase_Debate', 
                        color=col_ato, 
                        template="plotly_white",
                        color_discrete_map=mapa_cores,
                        category_orders={
                            col_ato: ordem_legenda,
                            'Fase_Debate': ordem_das_fases
                        }
                    )
                    
                    fig_timeline.update_layout(
                        xaxis={
                            'title': txt['eixo_x_cron'],
                            'showgrid': False,
                            'showticklabels': False 
                        }, 
                        yaxis={'title': txt['eixo_y_cron']},
                        legend_title_text=txt['legenda_comp'],
                        barmode='stack',
                        bargap=0.05,
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1) 
                    )
                    
                    # Atualizar o hover com as chavetas escapadas %{{...}}
                    fig_timeline.update_traces(
                        hovertemplate=f"<b>{txt['hover_fase']}:</b> {txt['hover_int_cap']} %{{x}}<br><b>%{{data.name}}:</b> %{{y}} {txt['hover_registos']}<extra></extra>"
                    )
                    
                    st.plotly_chart(fig_timeline, use_container_width=True)
                else:
                    st.warning("Não foi possível gerar a cronologia.")
                # Matriz de Reação (Heatmap)
                st.markdown("<br><hr style='border-top: 1px solid #e2e8f0; margin: 40px 0;'><br>", unsafe_allow_html=True)
                st.markdown(f"<h4 style='text-align: center; color: #0f172a; font-weight: 700; font-size: 1.4rem; margin-bottom: 5px;'>{txt['matriz_tit']}</h4>", unsafe_allow_html=True)
                st.markdown(f"""
                <p style='text-align: center; color: #475569; font-size: 1.05rem; max-width: 900px; margin: 0 auto 20px auto; line-height: 1.6;'>
                    {txt['matriz_desc_1']}<b>{ano_selecionado}</b>{txt['matriz_desc_2']}
                    <span style='color: #047857; font-weight: 600;'>{txt['matriz_verde']}</span>, 
                    <span style='color: #b91c1c; font-weight: 600;'>{txt['matriz_vermelho']}</span>{txt['matriz_e']}
                    <span style='color: #ca8a04; font-weight: 600;'>{txt['matriz_amarelo']}</span>.
                </p>
                """, unsafe_allow_html=True)
                
                # --- Lógica para tentar carregar a imagem em Inglês (se existir) ---
                pasta_imagens = "analise_parlamento"
                
                if lang == 'EN':
                    nome_heatmap = f"heatmap_indice_POLARIZACAO_EN_{ano_selecionado}.png"
                else:
                    nome_heatmap = f"heatmap_indice_POLARIZACAO_{ano_selecionado}.png"
                
                caminho_completo = os.path.join(pasta_imagens, nome_heatmap)
                
                # Fallback: Se estiver em Inglês mas a imagem _EN.png ainda não existir na pasta, mostra a versão PT original para não dar erro
                if lang == 'EN' and not os.path.exists(caminho_completo):
                    caminho_completo = os.path.join(pasta_imagens, f"heatmap_indice_POLARIZACAO_{ano_selecionado}.png")
                # -----------------------------------------------------------------

                if not os.path.exists(caminho_completo) and os.path.exists(nome_heatmap):
                    caminho_completo = nome_heatmap
                    
                if os.path.exists(caminho_completo):
                    import base64
                    with open(caminho_completo, "rb") as img_file:
                        img_base64 = base64.b64encode(img_file.read()).decode()
                    
                    # (O resto do teu código HTML com os botões de zoom e javascript continua aqui para baixo...)
                    
                    html_zoom = f"""
                    <div style="display: flex; justify-content: center; width: 100%;">
                        <div id="container-zoom" style="width: 100%; max-width: 800px; height: 500px; overflow: hidden; border: 1px solid #e2e8f0; border-radius: 12px; background: #ffffff; cursor: grab; position: relative; box-shadow: 0 4px 12px rgba(0,0,0,0.05); user-select: none;">
                            <img id="img-heatmap" src="data:image/png;base64,{img_base64}" style="width: 100%; height: 100%; object-fit: contain; transform-origin: center; transition: transform 0.05s ease-out; pointer-events: none;" title="Duplo clique para zoom rápido. Roda do rato para ajustar. Clique e arraste para explorar." />
                        </div>
                    </div>

                    <script>
                        const container = document.getElementById('container-zoom');
                        const img = document.getElementById('img-heatmap');
                        let scale = 1; let pointX = 0; let pointY = 0; let panning = false; let startX = 0; let startY = 0;

                        function constrain() {{
                            const maxTranslateX = (container.clientWidth * (scale - 1)) / 2;
                            const maxTranslateY = (container.clientHeight * (scale - 1)) / 2;
                            if (scale <= 1) {{ pointX = 0; pointY = 0; }} 
                            else {{ pointX = Math.max(-maxTranslateX, Math.min(maxTranslateX, pointX)); pointY = Math.max(-maxTranslateY, Math.min(maxTranslateY, pointY)); }}
                        }}

                        function setTransform() {{ constrain(); img.style.transform = `translate(${{pointX}}px, ${{pointY}}px) scale(${{scale}})`; }}

                        container.addEventListener('mousedown', function (e) {{
                            if (scale > 1) {{ panning = true; startX = e.clientX - pointX; startY = e.clientY - pointY; container.style.cursor = 'grabbing'; }}
                        }});

                        window.addEventListener('mouseup', function () {{ panning = false; container.style.cursor = scale > 1 ? 'grab' : 'default'; }});
                        window.addEventListener('mousemove', function (e) {{ if (!panning) return; pointX = e.clientX - startX; pointY = e.clientY - startY; setTransform(); }});

                        container.addEventListener('wheel', function (e) {{
                            e.preventDefault();
                            if (e.deltaY < 0) scale *= 1.2; else scale /= 1.2;
                            scale = Math.max(1, Math.min(scale, 6));
                            if (scale === 1) {{ pointX = 0; pointY = 0; container.style.cursor = 'default'; }} 
                            else container.style.cursor = 'grab';
                            setTransform();
                        }});

                        container.addEventListener('dblclick', function () {{
                            if (scale === 1) {{ scale = 2.5; container.style.cursor = 'grab'; }} 
                            else {{ scale = 1; pointX = 0; pointY = 0; container.style.cursor = 'default'; }}
                            setTransform();
                        }});
                    </script>
                    """
                    components.html(html_zoom, height=560)
                else:
                    st.warning(f"A imagem do heatmap para o ano {ano_selecionado} não foi encontrada na pasta '{pasta_imagens}'.")

                # TOP 10 TEMAS DO ANO
                st.markdown("<br><hr style='border-top: 1px solid #e2e8f0; margin: 30px 0;'><br>", unsafe_allow_html=True)
                st.markdown(f"<h4 style='text-align: center; color: #0f172a; font-weight: 700; font-size: 1.4rem; margin-bottom: 5px;'>{txt['top_temas_tit']} ({ano_selecionado})</h4>", unsafe_allow_html=True)
                st.markdown(f"<p style='text-align: center; color: #64748b; font-size: 1rem; margin-bottom: 25px;'>{txt['top_temas_sub']}</p>", unsafe_allow_html=True)

                if 'Utterance' in df_filtrado_tab1.columns:
                    import re
                    from collections import Counter
                    
                    stopwords_pt = set([
                        "o", "a", "os", "as", "e", "de", "do", "da", "dos", "das", "em", "no", "na", "nos", "nas", 
                        "para", "por", "com", "que", "se", "não", "sim", "mais", "mas", "como", "ao", "aos", "ou", 
                        "um", "uma", "uns", "umas", "sua", "seu", "seus", "suas", "meu", "minha", "nosso", "nossa", 
                        "teu", "tua", "este", "esta", "estes", "estas", "aquele", "aquela", "isso", "isto", "aquilo",
                        "neste", "nesta", "deste", "desta", "esse", "essa", "esses", "essas", "num", "numa",
                        "sr", "sra", "srs", "sras", "senhor", "senhora", "senhores", "senhoras", "deputado", 
                        "deputada", "deputados", "deputadas", "presidente", "ministro", "ministra", "ministros", 
                        "governo", "partido", "portugal", "país", "estado", "nação", "primeiro-ministro", 
                        "portugueses", "portuguesas", "palavra", "debate", "comissão", "socialista", "socialistas",
                        "porque", "também", "já", "muito", "tem", "temos", "ser", "foi", "são", "ter", "sobre", 
                        "quando", "nós", "vós", "eles", "elas", "ele", "ela", "quem", "onde", "anos", "hoje", 
                        "aqui", "agora", "fazer", "dizer", "vez", "tudo", "nada", "só", "está", "estamos", "estão", 
                        "há", "às", "pelo", "pela", "pelos", "pelas", "bem", "mal", "forma", "ainda", "então", "assim", 
                        "qual", "quais", "qualquer", "nem", "sem", "cada", "mesmo", "mesma", "todos", "todas", 
                        "pode", "podem", "podemos", "vamos", "queremos", "devemos", "dar", "estar", "sendo", 
                        "apenas", "dois", "duas", "entre", "primeiro"
                    ])
                    
                    df_ano_ind = df_filtrado_tab1.copy()
                    if 'Tipo' in df_ano_ind.columns:
                        df_ano_ind = df_ano_ind[df_ano_ind['Tipo'] != 'ação']
                        
                    textos = " ".join(df_ano_ind['Utterance'].dropna().astype(str).tolist()).lower()
                    palavras = re.findall(r'\b[a-záàâãéèêíïóôõöúç-]+\b', textos)
                    palavras_limpas = [p.capitalize() for p in palavras if p not in stopwords_pt and len(p) > 3]
                    
                    top_10 = Counter(palavras_limpas).most_common(10)
                    
                    if top_10:
                        termos_df = pd.DataFrame(top_10, columns=['Tema', 'Frequência'])
                        termos_df = termos_df.sort_values('Frequência', ascending=True)
                        
                        # Usamos uma escala de cores 'Cividis' ou 'Blues' mas com boa visibilidade
                        fig_temas = px.bar(
                            termos_df,
                            x='Frequência',
                            y='Tema',
                            orientation='h',
                            text='Frequência',
                            template="plotly_white",
                            color='Frequência',
                            color_continuous_scale='Tealgrn'
                        )
                        
                        fig_temas.update_layout(
                            xaxis={'title': txt['eixo_x_mencoes'], 'showgrid': True, 'gridcolor': '#f1f5f9'}, # <-- Alterado aqui
                            yaxis={'title': '', 'tickfont': {'size': 13, 'color': '#0f172a', 'family': 'sans-serif'}},
                            margin=dict(t=10, b=10, l=10, r=20),
                            height=420,
                            coloraxis_showscale=False
                        )
                        
                        fig_temas.update_traces(
                            texttemplate='%{text}',
                            textposition='outside',
                            textfont=dict(size=12, color='#334155', family='sans-serif'),
                            marker=dict(line=dict(width=0)),
                            hovertemplate=f"<b>{txt['hover_tema']}:</b> %{{y}}<br><b>{txt['hover_freq']}:</b> %{{x}}<extra></extra>"
                        )
                        
                        st.plotly_chart(fig_temas, use_container_width=True)
                    else:
                        st.info("Não existem dados suficientes para extrair temas neste ano.")
                else:
                    st.warning("Coluna 'Utterance' não encontrada.")

            # ABA 2: GRAFO INTERATIVO
            with tab2:
                nome_base_ficheiro = ficheiro_a_carregar.split('.')[0]
                nome_html_esperado = f"Interface_{nome_base_ficheiro}.html"
                caminho_html_encontrado = None
                
                if os.path.exists('./Resultados'):
                    pastas_resultados = sorted(glob.glob('./Resultados/*/'), key=os.path.getmtime, reverse=True)
                    for pasta in pastas_resultados:
                        possivel_caminho = os.path.join(pasta, nome_html_esperado)
                        if os.path.exists(possivel_caminho):
                            caminho_html_encontrado = possivel_caminho
                            break
                
                if caminho_html_encontrado:
                    print(f"👉 Ano {ano_selecionado}: O grafo HTML foi carregado da pasta -> {caminho_html_encontrado}")
                    with open(caminho_html_encontrado, 'r', encoding='utf-8') as f: 
                        html_source = f.read()
                    
                    # CORRIGIR TEXTOS E ERROS DO CÓDIGO NATIVO
                    html_source = html_source.replace("Cooperao", "Cooperação")
                    html_source = html_source.replace("Aco: Aplausos", "Ação: Aplausos")
                    html_source = html_source.replace("Aco: Protestos", "Ação: Protestos")
                    html_source = html_source.replace("Aco: Outros", "Ação: Outros")
                    html_source = html_source.replace("Acoes", "Ações")
                    html_source = re.sub(r'\s*\(\d+\)', '', html_source)
                    
                    
                    html_source = re.sub(r'\.toFixed(?!\()', '.toFixed(2)', html_source)
                                    
                    script_e_css = """
                    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
                    <style>
                        #graph-wrapper { overflow: hidden !important; height: 100vh !important; width: 100vw !important; position: relative !important; }
                        #graph-container { overflow: visible !important; transform-origin: 0 0 !important; width: 100% !important; height: 100% !important; }
                        svg { overflow: visible !important; width: 100% !important; height: 100% !important; }
                        
                        #grafico { margin: 0 !important; width: 100% !important; height: 100vh !important; padding: 0 !important; }
                        #sidebar, #openSidebar, .closebtn, #main-title, #file-subtitle { display: none !important; }
                        
                        /* ESTILOS DA BARRA DE FILTROS E LEGENDA */
                        #nova-barra-limpa { position: absolute; bottom: 20px; left: 20px; display: flex; flex-direction: column; align-items: flex-start; justify-content: center; gap: 12px; z-index: 9999; background: rgba(255, 255, 255, 0.95); padding: 15px 25px; border-radius: 20px; border: 1px solid #e2e8f0; box-shadow: 0 8px 30px rgba(0,0,0,0.08); pointer-events: auto; max-width: calc(100vw - 260px); width: auto; white-space: normal !important; }
                        .linha-um, .linha-dois { display: flex !important; flex-direction: row !important; align-items: center !important; justify-content: flex-start !important; width: 100% !important; flex-wrap: wrap !important; }
                        .linha-um { gap: 15px 30px !important; }
                        .linha-dois { gap: 10px !important; border-top: 1px solid #e2e8f0 !important; padding-top: 12px !important; }
                        .caixa-limpa, .caixa-limpa div { background: transparent !important; box-shadow: none !important; border: none !important; display: flex !important; flex-direction: row !important; align-items: center !important; flex-wrap: wrap !important; gap: 10px !important; margin: 0 !important; padding: 0 !important; }
                        #nova-barra-limpa *, #nova-barra-limpa label { color: #1E3A8A !important; font-family: inherit !important; font-weight: 700 !important; font-size: 14px !important; margin: 0 !important; text-shadow: none !important; }
                        #nova-barra-limpa label { display: flex !important; flex-direction: row !important; align-items: center !important; gap: 5px !important; cursor: pointer; white-space: nowrap !important; }
                        #nova-barra-limpa input[type="text"] { background: #f8fafc !important; border: 1px solid #cbd5e1 !important; padding: 6px 12px !important; border-radius: 20px !important; color: #0f172a !important; max-width: 140px !important; outline: none; }
                        #nova-barra-limpa input[type="range"] { max-width: 100px !important; cursor: pointer; }
                        #nova-barra-limpa input[type="checkbox"], #nova-barra-limpa input[type="radio"] { width: 16px !important; height: 16px !important; cursor: pointer; margin: 0 !important; }
                        .legenda-aberta { position: absolute !important; top: 20px !important; right: 20px !important; display: flex !important; flex-direction: column !important; max-height: 50vh !important; overflow-y: auto !important; opacity: 1 !important; background: rgba(255, 255, 255, 0.95) !important; border: 1px solid #e2e8f0 !important; border-radius: 12px !important; padding: 15px !important; box-shadow: 0 10px 30px rgba(0,0,0,0.1) !important; z-index: 9999 !important; pointer-events: auto !important; min-width: 170px !important; }
                        .legenda-aberta::-webkit-scrollbar { width: 4px; }
                        .legenda-aberta::-webkit-scrollbar-thumb { background-color: #cbd5e1; border-radius: 10px; }
                        .legenda-aberta * { color: #1E3A8A !important; font-size: 12px !important; font-weight: 600 !important; }
                        .legenda-aberta > div { display: flex !important; flex-direction: column !important; gap: 4px !important; margin-bottom: 10px !important; }
                        .legenda-aberta > div > div, .legenda-aberta span { display: flex !important; flex-direction: row !important; align-items: center !important; gap: 6px !important; }
                        .legenda-aberta h1, .legenda-aberta h2, .legenda-aberta h3, .legenda-aberta b, .legenda-aberta strong { text-transform: uppercase !important; margin-bottom: 6px !important; font-size: 11px !important; opacity: 0.8 !important; }
                        .titulo-caption-injetado { font-size: 14px !important; font-weight: 800 !important; text-transform: uppercase !important; border-bottom: 2px solid #e2e8f0 !important; padding-bottom: 6px !important; margin-bottom: 10px !important; opacity: 1 !important; }
                        [title*="Matrix heatmap"], [title*="Matrix Heatmap"] { display: none !important; }
                        @media screen and (max-width: 768px) {
                            .legenda-aberta { top: 5px !important; right: 5px !important; padding: 8px !important; min-width: 120px !important; transform: scale(0.75); transform-origin: top right; max-height: 35vh !important; }
                            #nova-barra-limpa { bottom: 5px !important; left: 5px !important; right: 5px !important; padding: 10px !important; max-width: calc(100vw - 10px) !important; transform: scale(0.8); transform-origin: bottom left; }
                            .linha-um { gap: 5px 10px !important; }
                            .linha-dois { gap: 5px !important; }
                        }
                    </style>
                    <script>
                        setTimeout(function() {
                            let container = document.getElementById("graph-container");
                            let wrapper = document.getElementById("graph-wrapper");
                            
                            if (container && wrapper) {
                                // 1. Matar os eventos defeituosos clonando os botões
                                ['btn-reset', 'btn-zoom-in', 'btn-zoom-out', 'btn-pan', 'btn-zoom-box'].forEach(id => {
                                    let btn = document.getElementById(id);
                                    if(btn) {
                                        let newBtn = btn.cloneNode(true);
                                        btn.parentNode.replaceChild(newBtn, btn);
                                    }
                                });
                                
                                let btnReset = document.getElementById("btn-reset");
                                let btnZoomIn = document.getElementById("btn-zoom-in");
                                let btnZoomOut = document.getElementById("btn-zoom-out");
                                let btnPan = document.getElementById("btn-pan");
                                let btnZoomBox = document.getElementById("btn-zoom-box");
                                
                                // Variáveis de estado imaculadas
                                let zl = 1, tx = 0, ty = 0;
                                let isPan = false, isDrag = false, isBox = false;
                                let sx = 0, sy = 0;
                                let selRect = document.getElementById("selection-rectangle");
                                
                                // Limpar lixo CSS nativo
                                wrapper.style.transform = "";
                                wrapper.style.transformOrigin = "";
                                container.style.transform = `translate(0px, 0px) scale(1)`;
                                container.style.transformOrigin = "0 0";
                                
                                function applyT() {
                                    container.style.transform = `translate(${tx}px, ${ty}px) scale(${zl})`;
                                }
                                
                                function off() {
                                    isPan = false; isBox = false; isDrag = false;
                                    if(btnPan) { btnPan.classList.remove("active"); btnPan.style.backgroundColor = ""; }
                                    if(btnZoomBox) { btnZoomBox.classList.remove("active"); btnZoomBox.style.backgroundColor = ""; }
                                    wrapper.style.cursor = "default";
                                    if(selRect) selRect.style.display = "none";
                                }
                                
                                // Eventos dos botões
                                if(btnReset) btnReset.addEventListener("click", () => { off(); zl=1; tx=0; ty=0; applyT(); });
                                if(btnZoomIn) btnZoomIn.addEventListener("click", () => { zl = Math.min(zl+0.2, 5); applyT(); });
                                if(btnZoomOut) btnZoomOut.addEventListener("click", () => { zl = Math.max(zl-0.2, 0.1); applyT(); });
                                
                                if(btnPan) btnPan.addEventListener("click", () => {
                                    if(isPan) off(); else { off(); isPan = true; btnPan.classList.add("active"); btnPan.style.backgroundColor = "#0055aa"; wrapper.style.cursor = "grab"; }
                                });
                                
                                if(btnZoomBox) btnZoomBox.addEventListener("click", () => {
                                    if(isBox) off();
                                    else { 
                                        off(); isBox = true; btnZoomBox.classList.add("active"); btnZoomBox.style.backgroundColor = "#0055aa"; wrapper.style.cursor = "crosshair"; 
                                        if (!selRect) {
                                            selRect = document.createElement("div");
                                            selRect.id = "selection-rectangle";
                                            selRect.style.position = "absolute";
                                            selRect.style.border = "2px dashed #0055aa";
                                            selRect.style.backgroundColor = "rgba(0, 85, 170, 0.2)";
                                            selRect.style.pointerEvents = "none";
                                            selRect.style.zIndex = "9999";
                                            wrapper.appendChild(selRect);
                                        }
                                    }
                                });
                                
                                // Intercetar e bloquear os eventos defeituosos do ecrã!
                                wrapper.addEventListener("wheel", (e) => {
                                    e.stopPropagation(); e.preventDefault();
                                    let zoomSpeed = 0.1;
                                    if(e.deltaY < 0) zl = Math.min(zl + zoomSpeed, 5);
                                    else zl = Math.max(zl - zoomSpeed, 0.1);
                                    applyT();
                                }, {capture: true, passive: false});
                                
                                wrapper.addEventListener("mousedown", (e) => {
                                    if(isBox || isPan) e.stopPropagation();
                                    if(isBox) {
                                        const r = wrapper.getBoundingClientRect();
                                        sx = (e.clientX - r.left - tx) / zl;
                                        sy = (e.clientY - r.top - ty) / zl;
                                        selRect.style.left = `${sx * zl + tx}px`;
                                        selRect.style.top = `${sy * zl + ty}px`;
                                        selRect.style.width = "0px";
                                        selRect.style.height = "0px";
                                        selRect.style.display = "block";
                                    } else if (isPan) {
                                        isDrag = true;
                                        sx = e.clientX - tx;
                                        sy = e.clientY - ty;
                                        wrapper.style.cursor = "grabbing";
                                    }
                                }, {capture: true});
                                
                                window.addEventListener("mousemove", (e) => {
                                    if(isBox && selRect && selRect.style.display !== "none") {
                                        e.stopPropagation();
                                        const r = wrapper.getBoundingClientRect();
                                        const mx = (e.clientX - r.left - tx) / zl;
                                        const my = (e.clientY - r.top - ty) / zl;
                                        const rx = Math.min(mx, sx);
                                        const ry = Math.min(my, sy);
                                        const rw = Math.abs(mx - sx);
                                        const rh = Math.abs(my - sy);
                                        selRect.style.left = `${rx * zl + tx}px`;
                                        selRect.style.top = `${ry * zl + ty}px`;
                                        selRect.style.width = `${rw * zl}px`;
                                        selRect.style.height = `${rh * zl}px`;
                                    } else if (isPan && isDrag) {
                                        e.stopPropagation();
                                        tx = e.clientX - sx;
                                        ty = e.clientY - sy;
                                        applyT();
                                    }
                                }, {capture: true});
                                
                                window.addEventListener("mouseup", (e) => {
                                    if(isBox && selRect && selRect.style.display !== "none") {
                                        e.stopPropagation();
                                        selRect.style.display = "none";
                                        const r = wrapper.getBoundingClientRect();
                                        const mx = (e.clientX - r.left - tx) / zl;
                                        const my = (e.clientY - r.top - ty) / zl;
                                        const x1 = Math.min(sx, mx);
                                        const y1 = Math.min(sy, my);
                                        const x2 = Math.max(sx, mx);
                                        const y2 = Math.max(sy, my);
                                        const sw = x2 - x1;
                                        const sh = y2 - y1;
                                        if(sw >= 10 && sh >= 10) {
                                            const scX = r.width / sw;
                                            const scY = r.height / sh;
                                            zl = Math.min(scX, scY);
                                            tx = -x1 * zl;
                                            ty = -y1 * zl;
                                            applyT();
                                        }
                                        off();
                                    } else if (isPan) {
                                        isDrag = false;
                                        wrapper.style.cursor = "grab";
                                    }
                                }, {capture: true});
                            }
                        }, 1000);

                        // JS QUE EMBELEZA AS CAIXAS (recolocação e cor)
                        setTimeout(function() {
                            try {
                                let todosElementos = document.querySelectorAll('*');
                                let captionBtn = null;
                                for (let el of todosElementos) { if (el.innerText && el.innerText.trim() === "Caption" && el.tagName !== 'STYLE' && el.tagName !== 'SCRIPT') { captionBtn = el; break; } }
                                if (captionBtn) {
                                    let painelCaption = captionBtn.nextElementSibling;
                                    if (painelCaption) {
                                        painelCaption.innerHTML = painelCaption.innerHTML.replace(/↳/g, '');
                                        document.body.appendChild(painelCaption); 
                                        painelCaption.classList.add('legenda-aberta'); 
                                        let titulo = document.createElement('div'); titulo.innerText = "Caption"; titulo.className = "titulo-caption-injetado"; painelCaption.prepend(titulo);
                                    }
                                    captionBtn.style.display = 'none'; 
                                }
                            } catch (err) {}
                            try {
                                let novaBarra = document.createElement('div'); novaBarra.id = 'nova-barra-limpa';
                                let linha1 = document.createElement('div'); linha1.className = 'linha-um';
                                let linha2 = document.createElement('div'); linha2.className = 'linha-dois';
                                document.body.appendChild(novaBarra); novaBarra.appendChild(linha1); novaBarra.appendChild(linha2);
                                let search = document.querySelector('input[type="text"]');
                                if(search) { let container = search.parentElement; if (!container.innerText.includes("Search")) { container = container.parentElement; } if (!container.innerText.toLowerCase().includes("search")) { let label = document.createElement('span'); label.innerText = "Search:"; label.style.marginRight = "5px"; container.prepend(label); } container.className = 'caixa-limpa'; linha1.appendChild(container); }
                                let slider = document.querySelector('input[type="range"]');
                                if(slider) { let container = slider.parentElement; if (!container.innerText.includes("Threshold")) { container = container.parentElement; } if (!container.innerText.toLowerCase().includes("threshold")) { let label = document.createElement('span'); label.innerText = "Threshold:"; label.style.marginRight = "5px"; container.prepend(label); } container.className = 'caixa-limpa'; linha1.appendChild(container); }
                                let tituloSpeakers = document.createElement('span'); tituloSpeakers.innerText = "Speakers:"; tituloSpeakers.style.marginRight = "10px"; linha2.appendChild(tituloSpeakers);
                                let checkboxes = document.querySelectorAll('input[type="checkbox"], input[type="radio"]');
                                if(checkboxes.length > 0) { let grid = checkboxes[0].closest('div[style*="grid"]') || checkboxes[0].parentElement.parentElement; if(grid) { grid.className = 'caixa-limpa'; linha2.appendChild(grid); } }
                                document.querySelectorAll('h2, h3, label').forEach(el => { if(el.innerText && el.innerText.trim() === "Speakers") el.style.display = "none"; });
                            } catch (err) { }
                        }, 800); 
                        
                        setInterval(function() {
                            try {
                                let divElements = document.querySelectorAll('div');
                                divElements.forEach(div => {
                                    let style = window.getComputedStyle(div);
                                    if (style.position === 'absolute' || style.position === 'fixed') {
                                        let right = parseInt(style.right); let bottom = parseInt(style.bottom);
                                        if (!isNaN(right) && right <= 120 && !isNaN(bottom) && bottom <= 120) {
                                            if (div.children.length >= 3) {
                                                div.style.setProperty('background', 'rgba(255, 255, 255, 0.95)', 'important');
                                                div.style.setProperty('border', '1px solid #e2e8f0', 'important');
                                                div.style.setProperty('box-shadow', '0 8px 30px rgba(0,0,0,0.08)', 'important');
                                                div.style.setProperty('border-radius', '25px', 'important');
                                                div.style.setProperty('padding', '6px 15px', 'important');
                                                div.querySelectorAll('*').forEach(icon => {
                                                    icon.style.setProperty('fill', '#1E3A8A', 'important'); icon.style.setProperty('color', '#1E3A8A', 'important'); icon.style.setProperty('stroke', '#1E3A8A', 'important');
                                                    if (icon.tagName === 'IMG') { icon.style.setProperty('filter', 'invert(16%) sepia(87%) saturate(2256%) hue-rotate(212deg) brightness(92%) contrast(98%)', 'important'); }
                                                });
                                            }
                                        }
                                    }
                                });
                            } catch (err) {}
                            try {
                                document.querySelectorAll('[title*="Heatmap"], [title*="heatmap"]').forEach(botao => {
                                    botao.style.setProperty('display', 'none', 'important');
                                });
                            } catch (err) {}
                        }, 200); 
                    </script>
                    """
                    html_source = html_source + script_e_css
                    components.html(html_source, height=850, scrolling=False)
                else:
                    st.warning(f"O explorador interativo para o ano de {ano_selecionado} ainda não foi gerado.")
            
            # ABA 3: EXPLORADOR DE DADOS
            with tab3:
                if 'Speaker' in df.columns:
                    # Preparar o filtro (mantemos o teu código original para limpar o menu)
                    df['Speaker_Filtro'] = df['Speaker'].astype(str).replace(['nan', 'NaN', 'None', 'none', '', ' '], "Ação: Outros")
                    if 'Tipo' in df.columns:
                        df.loc[df['Tipo'] == 'ação', 'Speaker_Filtro'] = 'Ação: Outros'
                        if 'acoes_simples' in df.columns:
                            mask_acao = (df['Tipo'] == 'ação')
                            df.loc[mask_acao, 'Speaker_Filtro'] = df.loc[mask_acao, 'acao_agrupada']
                            
                    oradores_reais = [s for s in df['Speaker_Filtro'].unique() if not str(s).startswith("Ação")]
                    oradores_originais = sorted(oradores_reais)
                    
                    # Substituir "Todos" pela variável do dicionário
                    opcoes_filtro = [txt['opcao_todos']] + oradores_originais
                    
                    c_filtro, c_check = st.columns([3, 1])
                    with c_filtro:
                        # Substituir o título e o valor default
                        oradores_tab3 = st.multiselect(txt['filtrar_oradores'], options=opcoes_filtro, default=[txt['opcao_todos']], key="filtro_oradores_tab3", on_change=atualizar_filtro_tab3)
                    with c_check:
                        st.markdown("""<div style='margin-top: 33px;'></div><style>div[data-testid="stCheckbox"] div[data-baseweb="checkbox"] div[data-checked="true"] { background-color: #3B82F6 !important; border-color: #3B82F6 !important; }</style>""", unsafe_allow_html=True)
                        # Substituir o texto da checkbox
                        mostrar_acoes_tab3 = st.checkbox(txt['incluir_acoes_sessao'], value=True, key="check_acoes_tab3")
                    
                    mask_acoes = (df['Tipo'] == 'ação') | df['Speaker_Filtro'].astype(str).str.startswith("Ação")
                    if col_ato: mask_acoes = mask_acoes | df[col_ato].astype(str).str.startswith("Ação")
                    
                    # Lógica de filtragem (ABA 3)
                    if len(oradores_tab3) == 0 or any(x in oradores_tab3 for x in ["Todos", "All"]): 
                        mask_oradores = pd.Series(True, index=df.index)
                    else: 
                        mask_oradores = df['Speaker'].isin(oradores_tab3)
                    
                    if txt['opcao_todos'] in oradores_tab3 or not oradores_tab3:
                        if mostrar_acoes_tab3:
                            df_filtrado_tab3 = df # Mostra literalmente tudo
                        else:
                            df_filtrado_tab3 = df[~mask_acoes] # Mostra só falas de todos, sem ações
                    else:
                        if mostrar_acoes_tab3:
                            # Mostra as falas e as ações APENAS do partido selecionado
                            df_filtrado_tab3 = df[mask_oradores] 
                        else:
                            # Mostra APENAS as falas verbais do partido selecionado
                            df_filtrado_tab3 = df[mask_oradores & ~mask_acoes]

                    if lang == 'EN':
                        mapa_colunas = {
                            'dialogue_id': 'Dialogue ID', 'turn_id': 'Turn ID', 'Speaker': 'Speaker',
                            'Utterance': 'Utterance', 'Ação': 'Action', col_ato: 'Behavior'
                        }
                    else:
                        mapa_colunas = {
                            'dialogue_id': 'ID do Diálogo', 'turn_id': 'ID do Turno', 'Speaker': 'Orador',
                            'Utterance': 'Fala', 'Ação': 'Ação', col_ato: 'Comportamento'
                        }
                    
                    colunas_presentes = {k: v for k, v in mapa_colunas.items() if k in df_filtrado_tab3.columns}
                    df_para_mostrar = df_filtrado_tab3[list(colunas_presentes.keys())].rename(columns=colunas_presentes)
                    df_para_mostrar = df_para_mostrar.replace(['None', 'none', 'nan', 'NaN'], '-')
                else:
                    df_para_mostrar = df
                
                st.dataframe(df_para_mostrar, use_container_width=True, height=500, hide_index=True)

    except FileNotFoundError:
        st.error(f"O ficheiro '{ficheiro_a_carregar}' não foi encontrado!")
    except Exception as e:
        st.error(f"Ocorreu um erro inesperado ao carregar os dados: {e}")