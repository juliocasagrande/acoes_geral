import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import yfinance as yf
import requests
import os
from datetime import datetime
from ta.trend import MACD
from ta.momentum import RSIIndicator
from ta.volatility import BollingerBands
import urllib3
import openpyxl
from openpyxl.styles import PatternFill, Font
import warnings

# Ignorar avisos de segurança
warnings.filterwarnings("ignore")
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Configurar o proxy manualmente (se necessário)
proxies = {
    "http": "http://ipv4.187.86.222.76.webdefence.global.blackspider.com:80",
    "https": "http://ipv4.187.86.222.76.webdefence.global.blackspider.com:80"
}

# Criar sessão HTTP com o proxy configurado
session = requests.Session()
session.proxies.update(proxies)
session.verify = False  # Desativar a verificação SSL (não recomendado em produção)

# Diretório base para salvar análises
base_dir = f"Análise dia {datetime.now().strftime('%d-%m-%y')}"

# Função para calcular indicadores técnicos
def calcular_indicadores(dados):
    """
    Calcula indicadores técnicos, incluindo SMA, RSI, MACD e Bandas de Bollinger, e atribui notas.
    """
    # Médias Móveis Simples
    dados['SMA20'] = dados['Close'].rolling(window=20).mean()
    dados['SMA50'] = dados['Close'].rolling(window=50).mean()
    dados['SMA'] = np.where(dados['SMA20'] > dados['SMA50'], 1, 0)

    # RSI e sua nota
    rsi = RSIIndicator(close=dados['Close'], window=14)
    dados['RSI'] = rsi.rsi()
    dados['RSI_Nota'] = pd.cut(
        dados['RSI'],
        bins=[-np.inf, 30, 40, 60, 70, np.inf],
        labels=[1, 0.75, 0.5, 0.25, 0],
        right=False
    ).astype(float)

    # MACD e sua nota
    macd = MACD(close=dados['Close'])
    dados['MACD'] = macd.macd()
    dados['MACD_Sinal'] = macd.macd_signal()
    dados['MACD_Nota'] = np.where(dados['MACD'] > dados['MACD_Sinal'], 1, 0)

    # Bandas de Bollinger e sua nota
    bollinger = BollingerBands(close=dados['Close'], window=20, window_dev=2)
    dados['Banda_Superior'] = bollinger.bollinger_hband()
    dados['Banda_Inferior'] = bollinger.bollinger_lband()
    intervalo = dados['Banda_Superior'] - dados['Banda_Inferior']
    dados['Bollinger_Posicao'] = (dados['Close'] - dados['Banda_Inferior']) / intervalo
    dados['Bollinger_Nota'] = pd.cut(
        dados['Bollinger_Posicao'],
        bins=[-np.inf, 0, 0.25, 0.5, 0.75, np.inf],
        labels=[1, 0.75, 0.5, 0.25, 0],
        right=False
    ).astype(float)

    return dados

# Função para calcular rendimentos
def calcular_rendimentos(dados):
    """
    Adiciona colunas com os rendimentos percentuais nos últimos 1, 3 e 6 meses.
    """
    dados['Rendimento_1M'] = (dados['Close'] / dados['Close'].shift(21) - 1) * 100
    dados['Rendimento_3M'] = (dados['Close'] / dados['Close'].shift(63) - 1) * 100
    dados['Rendimento_6M'] = (dados['Close'] / dados['Close'].shift(126) - 1) * 100
    return dados

# Função para criar DataFrame de análise
def criar_dataframe_analise(tickers, mercado):
    """
    Analisa uma lista de tickers, calcula indicadores e retornos e organiza as notas em um DataFrame.
    """
    resultados = []
    for ticker in tickers:
        print(f"Analisando {ticker} no mercado {mercado}...")
        try:
            dados = yf.download(ticker, period='1y', interval='1d', session=session)
            if dados.empty:
                print(f"Nenhum dado encontrado para {ticker}. Pulando.")
                continue

            # Calcular indicadores e rendimentos
            dados = calcular_indicadores(dados)
            dados = calcular_rendimentos(dados)

            # Capturar os dados mais recentes
            ultimos_dados = {
                'Ticker': ticker,
                'SMA': dados['SMA'].iloc[-1],
                'RSI': dados['RSI_Nota'].iloc[-1],
                'MACD': dados['MACD_Nota'].iloc[-1],
                'BOLLINGER': dados['Bollinger_Nota'].iloc[-1],
                'Rendimento_1M': dados['Rendimento_1M'].iloc[-1],
                'Rendimento_3M': dados['Rendimento_3M'].iloc[-1],
                'Rendimento_6M': dados['Rendimento_6M'].iloc[-1]
            }
            resultados.append(ultimos_dados)
        except Exception as e:
            print(f"Erro ao analisar {ticker}: {e}")
    return pd.DataFrame(resultados)

# Função para salvar DataFrames no Excel com formatação
def salvar_todos_dataframe_excel(df_br, df_usa, df_btc):
    """
    Salva os DataFrames em um arquivo Excel e aplica formatações condicionais.
    """

    # Adicionar coluna 'SOMA' e salvar as colunas normalizadas e não normalizadas
    for df in [df_br, df_usa, df_btc]:
        # Salvar colunas de rendimento não normalizadas
        #df['Rendimento_1M_Original'] = df['Rendimento_1M']
        #df['Rendimento_3M_Original'] = df['Rendimento_3M']
        #df['Rendimento_6M_Original'] = df['Rendimento_6M']

        # Normalizar os valores de rendimento para a escala de 0 a 1
        for col in ['Rendimento_1M', 'Rendimento_3M', 'Rendimento_6M']:
            min_val = df[col].min()
            max_val = df[col].max()
            if max_val != min_val:  # Evitar divisão por zero
                df[f'{col}_Norm'] = (df[col] - min_val) / (max_val - min_val)
            else:
                df[f'{col}_Norm'] = 0  # Se todos os valores forem iguais, normalizar para 0

        # Atualizar a coluna SOMA somando os valores normalizados
        df['SOMA'] = (
            df['SMA'] +
            df['RSI'] +
            df['MACD'] +
            df['BOLLINGER'] +
            df['Rendimento_1M_Norm'] +
            df['Rendimento_3M_Norm'] +
            df['Rendimento_6M_Norm']
        )

    # Caminho do arquivo Excel
    arquivo_excel = os.path.join(base_dir, f"Analise_Tickers_{datetime.now().strftime('%d-%m-%y')}.xlsx")
    os.makedirs(base_dir, exist_ok=True)

    # Salvar DataFrames no Excel
    with pd.ExcelWriter(arquivo_excel, engine='openpyxl') as writer:
        df_br.to_excel(writer, sheet_name='Brasil', index=False)
        df_usa.to_excel(writer, sheet_name='EUA', index=False)
        df_btc.to_excel(writer, sheet_name='Criptomoedas', index=False)

    print(f"Arquivo salvo antes da formatação: {arquivo_excel}")

    # Aplicar formatação condicional
    wb = openpyxl.load_workbook(arquivo_excel)
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]

        # Obter índices das colunas
        headers = {cell.value: idx + 1 for idx, cell in enumerate(sheet[1])}

        # Formatação da coluna 'SOMA'
        if 'SOMA' in headers:
            soma_col = [sheet.cell(row=row, column=headers['SOMA']).value for row in range(2, sheet.max_row + 1)]
            soma_col = [v for v in soma_col if isinstance(v, (int, float))]
            if soma_col:
                media = sum(soma_col) / len(soma_col)
                desvio_padrao = (sum((x - media) ** 2 for x in soma_col) / len(soma_col)) ** 0.5
                limite = media + desvio_padrao

                for row in range(2, sheet.max_row + 1):
                    cell = sheet.cell(row=row, column=headers['SOMA'])
                    if isinstance(cell.value, (int, float)) and cell.value < limite:
                        cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                        cell.font = Font(color="FFFFFF", bold=True)

        # Formatação das colunas de rendimento
        for col in ['Rendimento_1M', 'Rendimento_3M', 'Rendimento_6M']:
            if col in headers:
                col_idx = headers[col]
                for row in range(2, sheet.max_row + 1):
                    cell = sheet.cell(row=row, column=col_idx)
                    if isinstance(cell.value, (int, float)):
                        if cell.value > 0:
                            cell.fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
                            cell.font = Font(color="000000", bold=True)
                        elif cell.value < 0:
                            cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                            cell.font = Font(color="FFFFFF", bold=True)

    # Salvar o arquivo Excel com as formatações aplicadas
    wb.save(arquivo_excel)
    print(f"Arquivo salvo com formatações aplicadas: {arquivo_excel}")

# Listas de tickers para análise
lista_br = [
    "LOGN3.SA", "RAIL3.SA", "PTBL3.SA", "ALPK3.SA", "ANIM3.SA", "AURA33.SA",
    "AESB3.SA", "SIMH3.SA", "ZAMP3.SA", "CSNA3.SA", "CVCB3.SA", "VVEO3.SA",
    "QUAL3.SA", "LIGT3.SA", "COGN3.SA", "MOVI3.SA", "CBAV3.SA", "MATD3.SA",
    "ALPA4.SA", "AERI3.SA", "DASA3.SA", "NTCO3.SA", "BRKM5.SA", "PCAR3.SA",
    "AZUL4.SA", "GOLL4.SA", "BHIA3.SA", "SYNE3.SA", "VBBR3.SA", "ETER3.SA",
    "POSI3.SA", "EUCA4.SA", "TPIS3.SA", "BRAP3.SA", "BRAP4.SA", "PINE4.SA",
    "BBAS3.SA", "CMIG4.SA", "HBOR3.SA", "HBRE3.SA", "BRSR6.SA", "TRPL4.SA",
    "JHSF3.SA", "SAPR3.SA", "BMGB4.SA", "SAPR4.SA", "WIZC3.SA", "ALLD3.SA",
    "SAPR11.SA", "TECN3.SA", "ABCB4.SA", "MDNE3.SA", "CMIG3.SA", "ECOR3.SA",
    "VALE3.SA", "DEXP3.SA", "LAVV3.SA", "VLID3.SA", "PETR4.SA", "CAML3.SA",
    "BMEB4.SA", "CYRE3.SA", "SBFG3.SA", "PETR3.SA", "NEOE3.SA", "LEVE3.SA",
    "GOAU3.SA", "GOAU4.SA", "CSMG3.SA", "CPFE3.SA", "COCE5.SA", "SOJA3.SA",
    "POMO3.SA", "ROMI3.SA", "SBSP3.SA", "CSUD3.SA", "PRIO3.SA", "TRIS3.SA",
    "CMIN3.SA", "SMTO3.SA", "GRND3.SA", "KEPL3.SA", "SHUL4.SA", "EGIE3.SA",
    "SANB3.SA", "ENGI11.SA", "PLPL3.SA", "ITUB3.SA", "UGPA3.SA", "JSLG3.SA",
    "VULC3.SA", "BBSE3.SA", "PFRM3.SA", "FIQE3.SA", "ELET3.SA", "BRBI11.SA",
    "BBDC3.SA", "GGBR3.SA", "SANB11.SA", "TAEE3.SA", "TAEE4.SA", "CSED3.SA",
    "TAEE11.SA", "HYPE3.SA", "AGRO3.SA", "CPLE3.SA", "MILS3.SA", "RECV3.SA",
    "USIM3.SA", "GGBR4.SA", "USIM5.SA", "BOBR4.SA", "SANB4.SA", "TTEN3.SA",
    "TGMA3.SA", "ITUB4.SA", "VAMO3.SA", "MELK3.SA", "CSAN3.SA", "ELET6.SA",
    "CPLE6.SA", "BBDC4.SA", "ALUP11.SA", "DIRR3.SA", "INTB3.SA", "MTRE3.SA",
    "MDIA3.SA", "POMO4.SA", "MYPK3.SA", "TUPY3.SA", "PSSA3.SA", "LJQQ3.SA"
]
lista_usa = ["AAPL", "MSFT", "GOOG", "AMZN"]  # Exemplo para EUA
lista_btc = ["BTC-USD", "ETH-USD"]  # Exemplo para criptomoedas

# Criar DataFrames para cada mercado
df_br = criar_dataframe_analise(lista_br, 'Brasil')
df_usa = criar_dataframe_analise(lista_usa, 'EUA')
df_btc = criar_dataframe_analise(lista_btc, 'Criptomoedas')

# Salvar resultados no Excel com as formatações aplicadas
salvar_todos_dataframe_excel(df_br, df_usa, df_btc)

print("Processamento e análise concluídos.")

## ANÁLISE DETALHADA
# Função para plotar gráficos em subplots
def plotar_e_salvar_subplots(dados, ticker, mercado, sentimento):
    """
    Plota todos os gráficos de um ativo em subplots e salva em uma única imagem.
    """
    # Diretório específico para o mercado
    dir_mercado = os.path.join(base_dir, f"Mercado {mercado}")
    os.makedirs(dir_mercado, exist_ok=True)

    # Criar figura com subplots
    fig, axes = plt.subplots(4, 1, figsize=(14, 24))
    fig.subplots_adjust(hspace=0.5)
    fig.suptitle(f"Análise {ticker} - Sentimento atual: {sentimento}", fontsize=16, y=0.95)

    # Gráfico 1: Preço de Fechamento e Médias Móveis
    axes[0].plot(dados['Close'], label='Preço de Fechamento', color='blue')
    axes[0].plot(dados['SMA20'], label='SMA 20', color='red')
    axes[0].plot(dados['SMA50'], label='SMA 50', color='green')
    axes[0].set_title(f'Preço de Fechamento e Médias Móveis - {ticker}')
    axes[0].legend(loc='upper left')
    axes[0].set_xlabel('Data')
    axes[0].set_ylabel('Preço')

    # Identificar cruzamentos de Médias Móveis
    crosses = (dados['SMA20'] > dados['SMA50']) & (dados['SMA20'].shift() <= dados['SMA50'].shift())
    axes[0].scatter(dados.index[crosses], dados['SMA20'][crosses], label='Cruzamento de Alta (SMA20 > SMA50)', color='green', marker='^', zorder=5)
    axes[0].text(dados.index[crosses][-1], dados['SMA20'][crosses][-1], 'ALTA', color='green', fontsize=9)

    crosses_baixa = (dados['SMA20'] < dados['SMA50']) & (dados['SMA20'].shift() >= dados['SMA50'].shift())
    axes[0].scatter(dados.index[crosses_baixa], dados['SMA20'][crosses_baixa], label='Cruzamento de Baixa (SMA20 < SMA50)', color='red', marker='v', zorder=5)
    axes[0].text(dados.index[crosses_baixa][-1], dados['SMA20'][crosses_baixa][-1], 'BAIXA', color='red', fontsize=9)


    axes[0].text(0.5, -0.10, 
                 "Cruzamento de Alta: Quando o SMA20 cruza o SMA50 para cima (SMA20>SMA50)", 
                 fontsize=10, ha='center', va='top', transform=axes[0].transAxes)

    # Gráfico 2: RSI com Eixo Y Secundário
    axes[1].plot(dados['RSI'], label='RSI', color='purple')
    ax2_rsi = axes[1].twinx()
    ax2_rsi.plot(dados['Close'], label='Preço de Fechamento', color='blue', alpha=0.3)
    axes[1].axhline(70, color='red', linestyle='--', label='Sobrecomprado (70)')
    axes[1].axhline(30, color='green', linestyle='--', label='Sobrevendido (30)')
    axes[1].set_title(f'RSI - {ticker}')
    axes[1].legend(loc='upper left')
    ax2_rsi.legend(loc='upper right')
    axes[1].text(0.5, -0.10, 
                 "RSI: Acima de 70 -> Sobrecomprado / Abaixo de 30 -> Sobrevendido.", 
                 fontsize=10, ha='center', va='top', transform=axes[1].transAxes)

    # Gráfico 3: MACD com Eixo Y Secundário
    axes[2].plot(dados['MACD'], label='MACD', color='blue')
    axes[2].plot(dados['MACD_Sinal'], label='Linha de Sinal', color='red')
    ax2_macd = axes[2].twinx()
    ax2_macd.plot(dados['Close'], label='Preço de Fechamento', color='gray', alpha=0.3)
    axes[2].set_title(f'MACD - {ticker}')
    axes[2].legend(loc='upper left')
    ax2_macd.legend(loc='upper right')

    # Identificar cruzamentos de MACD
    macd_compra = (dados['MACD'] > dados['MACD_Sinal']) & (dados['MACD'].shift() <= dados['MACD_Sinal'].shift())
    axes[2].scatter(dados.index[macd_compra], dados['MACD'][macd_compra], label='Sinal de Compra', color='green', marker='^', zorder=5)
    axes[2].text(dados.index[macd_compra][-1], dados['MACD'][macd_compra][-1], 'COMPRA', color='green', fontsize=9)

    macd_venda = (dados['MACD'] < dados['MACD_Sinal']) & (dados['MACD'].shift() >= dados['MACD_Sinal'].shift())
    axes[2].scatter(dados.index[macd_venda], dados['MACD'][macd_venda], label='Sinal de Venda', color='red', marker='v', zorder=5)
    axes[2].text(dados.index[macd_venda][-1], dados['MACD'][macd_venda][-1], 'VENDA', color='red', fontsize=9)


    axes[2].text(0.5, -0.10, 
                 "MACD: Quando a MACD cruza a linha de sinal para cima, pode ser SINAL DE COMPRA", 
                 fontsize=10, ha='center', va='top', transform=axes[2].transAxes)

    # Gráfico 4: Bandas de Bollinger
    axes[3].plot(dados['Close'], label='Preço de Fechamento', color='blue')
    axes[3].plot(dados['Banda_Superior'], label='Banda Superior', color='red', linestyle='--')
    axes[3].plot(dados['Banda_Inferior'], label='Banda Inferior', color='red', linestyle='--')
    axes[3].fill_between(dados.index, dados['Banda_Superior'], dados['Banda_Inferior'], color='grey', alpha=0.1)
    axes[3].set_title(f'Bandas de Bollinger - {ticker}')
    axes[3].legend(loc='upper left')
    axes[3].text(0.5, -0.10, 
                 "Bandas de Bollinger: Toque na Banda Superior -> Sobrecomprado; Inferior -> Sobrevendido.", 
                 fontsize=10, ha='center', va='top', transform=axes[3].transAxes)

    # Salvar o gráfico
    plt.tight_layout(rect=[0, 0.03, 1, 0.95])
    plt.savefig(os.path.join(dir_mercado, f"{ticker}.png"), bbox_inches="tight")
    plt.close()

# Listas de tickers
tickers_brasil = ['AURA33.SA', 'BMEB4.SA', 'BMGB4.SA', 'CMIG4.SA', 'CSMG3.SA', 'CURY3.SA', 'PLPL3.SA', 'SYNE3.SA', 'TECN3.SA', 'TPIS3.SA', 'VLID3.SA']
tickers_usa = ['CMCSA', 'MRO', 'EOG', 'MO', 'FANG', 'SINA']
indices = {'Ibovespa': '^BVSP', 'S&P500': '^GSPC', 'Bitcoin': 'BTC-USD'}


# Análise do Mercado Brasileiro
for ticker in tickers_brasil:
    print(f"Analisando {ticker} no Mercado Brasileiro...")
    try:
        dados = yf.download(ticker, period='1y', interval='1d', session=session)
        if dados.empty:
            print(f"Nenhum dado encontrado para {ticker}. Pulando.\n")
            continue
        dados = calcular_indicadores(dados)
        plotar_e_salvar_subplots(dados, ticker, mercado='Brasileiro', sentimento="Neutro")
    except Exception as e:
        print(f"Erro ao analisar {ticker}: {e}\n")

# Análise do Mercado Americano
for ticker in tickers_usa:
    print(f"Analisando {ticker} no Mercado Americano...")
    try:
        dados = yf.download(ticker, period='1y', interval='1d', session=session)
        if dados.empty:
            print(f"Nenhum dado encontrado para {ticker}. Pulando.\n")
            continue
        dados = calcular_indicadores(dados)
        plotar_e_salvar_subplots(dados, ticker, mercado='Americano', sentimento="Neutro")
    except Exception as e:
        print(f"Erro ao analisar {ticker}: {e}\n")

# Análise de Índices Gerais
for indice, ticker in indices.items():
    print(f"Analisando {indice}...")
    try:
        dados = yf.download(ticker, period='1y', interval='1d', session=session)
        if dados.empty:
            print(f"Nenhum dado encontrado para {indice}. Pulando.\n")
            continue
        dados = calcular_indicadores(dados)
        mercado = 'Brasileiro' if indice == 'Ibovespa' else 'Americano' if indice == 'S&P500' else 'Bitcoin'
        plotar_e_salvar_subplots(dados, indice, mercado=mercado, sentimento="Alta Volatilidade")
    except Exception as e:
        print(f"Erro ao analisar {indice}: {e}\n")
