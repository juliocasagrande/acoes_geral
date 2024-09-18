import streamlit as st
import yfinance as yf
import pandas as pd
import numpy as np

st.title('Sistema de Monitoramento de Ações dos EUA')

st.sidebar.header('Parâmetros de Filtragem')

# Parâmetros de filtragem com descrições
liquidez_minima = st.sidebar.number_input('Liquidez Mínima (USD)', value=0, help="Valor mínimo de liquidez diária (USD) nos últimos 2 meses, sugerido: 5.000.000")
pl_maximo = st.sidebar.number_input('P/L Máximo', value=15, help="Filtro para o preço/lucro máximo, sugerido: até 15")
ev_ebitda_maximo = st.sidebar.number_input('EV/EBITDA Máximo', value=12, help="Filtro para EV/EBITDA máximo, sugerido: até 12")
margem_bruta_minima = st.sidebar.number_input('Margem Bruta Mínima (%)', value=40, help="Margem bruta mínima para filtrar ações, sugerido: > 40%")
roa_minimo = st.sidebar.number_input('ROA Mínimo (%)', value=5, help="Retorno sobre ativos mínimo, sugerido: > 5%")
rendimento_positivo = st.sidebar.checkbox('Rendimento Positivo nos Últimos 12 Meses', value=True, help="Filtrar ações com rendimento positivo nos últimos 12 meses")
ebit_positivo = st.sidebar.checkbox('EBIT Positivo', value=True, help="Filtrar apenas ações com EBIT positivo")

# Função para obter os tickers da S&P 500
@st.cache_data
def get_sp500_tickers():
    url = 'https://en.wikipedia.org/wiki/List_of_S%26P_500_companies'
    table = pd.read_html(url)
    df = table[0]
    tickers = df['Symbol'].tolist()
    return tickers

tickers_list = get_sp500_tickers()

@st.cache_data
def get_financial_data(tickers):
    financial_data = []
    for ticker in tickers:
        try:
            stock = yf.Ticker(ticker)
            info = stock.info

            # Obtendo dados históricos dos últimos dois meses
            hist = stock.history(period='3mo')
            if not hist.empty:
                volume_medio_2m = hist['Volume'].mean()
                preco_medio_2m = hist['Close'].mean()
                liquidez = volume_medio_2m * preco_medio_2m
            else:
                liquidez = 0

            # Calculando rendimento dos últimos 12 meses
            hist_12m = stock.history(period='1y')
            rendimento_12m = ((hist_12m['Close'][-1] - hist_12m['Close'][0]) / hist_12m['Close'][0]) * 100 if not hist_12m.empty else 0

            data = {
                'Ticker': ticker,
                'Liquidez': liquidez,
                'EBIT': info.get('ebitda'),
                'P/L': info.get('trailingPE'),
                'EV/EBITDA': info.get('enterpriseToEbitda'),
                'PSR': info.get('priceToSalesTrailing12Months'),
                'Margem Bruta': info.get('grossMargins') * 100 if info.get('grossMargins') else 0,  # Convertendo para percentual
                'ROA': info.get('returnOnAssets') * 100 if info.get('returnOnAssets') else 0,  # Convertendo para percentual
                'Rendimento 12M (%)': rendimento_12m
            }
            financial_data.append(data)
        except Exception as e:
            continue
    return pd.DataFrame(financial_data)

def filtrar_acoes(df):
    df = df.dropna(subset=['Liquidez', 'EBIT', 'P/L', 'EV/EBITDA', 'PSR', 'Margem Bruta', 'ROA', 'Rendimento 12M (%)'])

    # Aplicando os filtros definidos
    df = df[df['Liquidez'] >= liquidez_minima]
    if ebit_positivo:
        df = df[df['EBIT'] > 0]
    df = df[df['P/L'] <= pl_maximo]
    df = df[df['EV/EBITDA'] <= ev_ebitda_maximo]
    df = df[df['Margem Bruta'] >= margem_bruta_minima]
    df = df[df['ROA'] >= roa_minimo]
    if rendimento_positivo:
        df = df[df['Rendimento 12M (%)'] > 0]

    # Ordenando pelo PSR
    df = df.sort_values(by='PSR')
    return df

if st.button('Executar Análise'):
    with st.spinner('Coletando dados...'):
        dados_financeiros = get_financial_data(tickers_list)

    st.success('Dados coletados com sucesso!')

    with st.spinner('Aplicando filtros...'):
        resultado = filtrar_acoes(dados_financeiros)

    st.success('Filtros aplicados!')

    st.header('Resultado da Análise')
    st.dataframe(resultado.reset_index(drop=True))
