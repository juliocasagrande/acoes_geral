import streamlit as st
import yfinance as yf
import pandas as pd
import numpy as np

# Título principal
st.title('Sistema de Monitoramento de Ações')

# Menu de navegação na sidebar
menu = st.sidebar.selectbox("Selecione a Página", ["Ativos S&P 500", "Ativos Ibovespa"])

# Parâmetros de filtragem compartilhados
st.sidebar.header('Parâmetros de Filtragem')
liquidez_minima = st.sidebar.number_input('Liquidez Mínima (USD)', value=0, help="Valor mínimo de liquidez diária (USD) nos últimos 2 meses")
pl_maximo = st.sidebar.number_input('P/L Máximo', value=15, help="Preço/Lucro máximo")
ev_ebitda_maximo = st.sidebar.number_input('EV/EBITDA Máximo', value=12, help="EV/EBITDA máximo")
margem_bruta_minima = st.sidebar.number_input('Margem Bruta Mínima (%)', value=40, help="Margem bruta mínima (%)")
roa_minimo = st.sidebar.number_input('ROA Mínimo (%)', value=5, help="Retorno sobre ativos mínimo (%)")
rendimento_positivo = st.sidebar.checkbox('Rendimento Positivo nos Últimos 12 Meses', value=True)
ebit_positivo = st.sidebar.checkbox('EBIT Positivo', value=True)

# Função para obter tickers do S&P 500
@st.cache_data
def get_sp500_tickers():
    url = 'https://en.wikipedia.org/wiki/List_of_S%26P_500_companies'
    table = pd.read_html(url)
    df = table[0]
    tickers = df['Symbol'].tolist()
    return tickers

# Função para obter tickers do Ibovespa
@st.cache_data
def get_ibovespa_tickers():
    file_path = 'AcoesIndices_2024-11-21.csv'  # Caminho do arquivo na raiz do projeto
    try:
        # Lendo o arquivo
        tickers_data = pd.read_csv(file_path, delimiter=';', skiprows=2, encoding='ISO-8859-1')
        # Extraindo os tickers da coluna "Empresa"
        tickers_list = tickers_data['Empresa'].tolist()
        return tickers_list
    except Exception as e:
        st.error(f"Erro ao carregar os tickers do Ibovespa: {e}")
        return []

# Função para coletar dados financeiros
@st.cache_data
def get_financial_data(tickers):
    financial_data = []
    for ticker in tickers:
        try:
            stock = yf.Ticker(ticker)
            info = stock.info

            # Dados históricos dos últimos dois meses
            hist = stock.history(period='3mo')
            if not hist.empty:
                volume_medio_2m = hist['Volume'].mean()
                preco_medio_2m = hist['Close'].mean()
                liquidez = volume_medio_2m * preco_medio_2m
            else:
                liquidez = 0

            # Rendimento dos últimos 12 meses
            hist_12m = stock.history(period='1y')
            rendimento_12m = ((hist_12m['Close'][-1] - hist_12m['Close'][0]) / hist_12m['Close'][0]) * 100 if not hist_12m.empty else 0

            data = {
                'Ticker': ticker,
                'Liquidez': liquidez,
                'EBIT': info.get('ebitda'),
                'P/L': info.get('trailingPE'),
                'EV/EBITDA': info.get('enterpriseToEbitda'),
                'PSR': info.get('priceToSalesTrailing12Months'),
                'Margem Bruta': info.get('grossMargins') * 100 if info.get('grossMargins') else 0,
                'ROA': info.get('returnOnAssets') * 100 if info.get('returnOnAssets') else 0,
                'Rendimento 12M (%)': rendimento_12m
            }
            financial_data.append(data)
        except Exception as e:
            continue
    return pd.DataFrame(financial_data)

# Função para filtrar ações
def filtrar_acoes(df):
    df = df.dropna(subset=['Liquidez', 'EBIT', 'P/L', 'EV/EBITDA', 'PSR', 'Margem Bruta', 'ROA', 'Rendimento 12M (%)'])

    # Aplicando filtros
    df = df[df['Liquidez'] >= liquidez_minima]
    if ebit_positivo:
        df = df[df['EBIT'] > 0]
    df = df[df['P/L'] <= pl_maximo]
    df = df[df['EV/EBITDA'] <= ev_ebitda_maximo]
    df = df[df['Margem Bruta'] >= margem_bruta_minima]
    df = df[df['ROA'] >= roa_minimo]
    if rendimento_positivo:
        df = df[df['Rendimento 12M (%)'] > 0]

    # Ordenação por PSR
    df = df.sort_values(by='PSR')
    return df

# Página de Ativos do S&P 500
if menu == "Ativos S&P 500":
    st.header("Análise de Ações do S&P 500")
    tickers_list = get_sp500_tickers()
    
    if st.button('Executar Análise para S&P 500'):
        with st.spinner('Coletando dados...'):
            dados_financeiros = get_financial_data(tickers_list)

        st.success('Dados coletados com sucesso!')

        with st.spinner('Aplicando filtros...'):
            resultado = filtrar_acoes(dados_financeiros)

        st.success('Filtros aplicados!')

        st.header('Resultado da Análise')
        st.dataframe(resultado.reset_index(drop=True))

# Página de Ativos do Ibovespa
elif menu == "Ativos Ibovespa":
    st.header("Análise de Ações do Ibovespa")
    tickers_list = get_ibovespa_tickers()
    
    if st.button('Executar Análise para Ibovespa'):
        with st.spinner('Coletando dados...'):
            dados_financeiros = get_financial_data(tickers_list)

        st.success('Dados coletados com sucesso!')

        with st.spinner('Aplicando filtros...'):
            resultado = filtrar_acoes(dados_financeiros)

        st.success('Filtros aplicados!')

        st.header('Resultado da Análise')
        st.dataframe(resultado.reset_index(drop=True))
