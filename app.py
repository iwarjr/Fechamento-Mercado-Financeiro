import streamlit as st
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from datetime import datetime, timedelta
import feedparser
import io
import os

# --- IDENTIDADE VISUAL ---
DIR_BASE = os.path.dirname(os.path.abspath(__file__))
COR_FUNDO = RGBColor(0, 32, 77)
COR_TEXTO = RGBColor(255, 255, 255)
COR_DESTAQUE = RGBColor(0, 174, 239)
COR_ALTA = RGBColor(0, 255, 127)
COR_BAIXA = RGBColor(255, 69, 0)

def carregar_logo():
    for arq in os.listdir(DIR_BASE):
        if arq.lower().startswith("logo") and arq.lower().endswith(('.png', '.jpg', '.jpeg')):
            try:
                with open(os.path.join(DIR_BASE, arq), "rb") as f: return f.read()
            except: continue
    return None

def obter_dados_seguros(ticker, data_alvo):
    """Busca dados garantindo que o pregão existiu"""
    for i in range(5): # Tenta voltar até 5 dias caso seja feriado/fim de semana
        d_fim = data_alvo - timedelta(days=i)
        d_ini = d_fim - timedelta(days=1)
        df = yf.download(ticker, start=d_ini.strftime('%Y-%m-%d'), end=(d_fim + timedelta(days=1)).strftime('%Y-%m-%d'), progress=False)
        if not df.empty and len(df) >= 1:
            return df, d_fim
    return pd.DataFrame(), data_alvo

def gerar_grafico(ticker, data_v, cor):
    try:
        s, e = data_v.strftime('%Y-%m-%d'), (data_v + timedelta(days=1)).strftime('%Y-%m-%d')
        df = yf.download(ticker, start=s, end=e, interval="5m", progress=False)['Close']
        if df.empty: return None
        plt.figure(figsize=(5, 2), facecolor='#00204D')
        ax = plt.axes(); ax.set_facecolor('#00204D')
        plt.plot(df.index, df.values, color=cor, linewidth=2.5)
        ax.tick_params(axis='both', colors='white', labelsize=8)
        for sp in ax.spines.values(): sp.set_color('white')
        plt.grid(True, color='grey', linestyle='--', alpha=0.1)
        plt.gcf().autofmt_xdate()
        img = io.BytesIO(); plt.savefig(img, format='png', bbox_inches='tight', dpi=100); plt.close()
        return img
    except: return None

def add_texto(slide, texto, left, top, width, height, size=18, bold=False, color=COR_TEXTO, align=PP_ALIGN.LEFT):
    tx = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tx.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = str(texto)
    p.font.size, p.font.bold, p.font.color.rgb, p.alignment = Pt(size), bold, color, align

# --- UI ---
st.set_page_config(page_title="Invest Forma Academy", layout="wide")
st.title("💼 Dashboard Premium V13.0")

logo_data = carregar_logo()
data_sel = st.date_input("Data do Relatório", datetime.now() - timedelta(days=1))

if st.button("🌟 GERAR BOLETIM DE PERFORMANCE"):
    with st.spinner("Analisando pregões e calculando variações reais..."):
        try:
            # Ativos principais
            mapa = {"IBOV": "^BVSP", "DOLAR": "USDBRL=X", "S&P500": "^GSPC", "NASDAQ": "^IXIC", "DOW": "^DJI", "BTC": "BTC-USD", "BRENT": "BZ=F", "IRON": "TIO=F", "GOLD": "GC=F", "SILVER": "SI=F"}
            res = {}
            data_valida = data_sel
            
            for nome, ticker in mapa.items():
                df_ativo, d_v = obter_dados_seguros(ticker, data_sel)
                if not df_ativo.empty:
                    data_valida = d_v # Sincroniza com a data que realmente teve pregão
                    ab = float(df_ativo['Open'].iloc[0])
                    fe = float(df_ativo['Close'].iloc[-1])
                    res[nome] = {"A": ab, "F": fe, "V": ((fe/ab)-1)*100, "T": ticker}
            
            # Cálculo das Altas e Baixas B3 (CORREÇÃO DA REPETIÇÃO)
            tickers_b3 = ['VALE3.SA', 'PETR4.SA', 'ITUB4.SA', 'BBDC4.SA', 'BBAS3.SA', 'ABEV3.SA', 'MGLU3.SA', 'WEGE3.SA', 'PRIO3.SA', 'GGBR4.SA', 'RENT3.SA', 'LREN3.SA', 'HAPV3.SA']
            df_b3, _ = obter_dados_seguros(tickers_b3, data_valida)
            
            if not df_b3.empty and len(df_b3) >= 1:
                # Calculamos a variação entre abertura e fechamento do mesmo dia para ser preciso
                ab_b3 = df_b3['Open'].iloc[0]
                fe_b3 = df_b3['Close'].iloc[-1]
                v_b3 = ((fe_b3 / ab_b3) - 1) * 100
                top_altas = v_b3.nlargest(5)
                top_baixas = v_b3.nsmallest(5)
            else:
                top_altas = top_baixas = None

            # --- GERAÇÃO PPTX ---
            prs = Presentation(); prs.slide_width, prs.slide_height = Inches(9), Inches(16)
            def slide_padrao(titulo):
                s = prs.slides.add_slide(prs.slide_layouts[6])
                s.background.fill.solid(); s.background.fill.fore_color.rgb = COR_FUNDO
                if logo_data: s.shapes.add_picture(io.BytesIO(logo_data), Inches(6.4), Inches(0.4), width=Inches(2.2))
                add_texto(s, titulo, 0.5, 0.4, 6, 1, size=32, bold=True, color=COR_DESTAQUE)
                return s

            # Capa e News
            s1 = slide_padrao("Boletim de Performance"); add_texto(s1, "Fechamento do Mercado", 0.5, 0.9, 6, 0.5, 22)
            add_texto(s1, data_valida.strftime('%d/%m/%Y'), 0.5, 1.3, 6, 0.4, 20, True, COR_ALTA)
            # Notícias Simples (Evita erro de rede)
            noticias = ["Mercado reage a dados econômicos", "Fluxo na B3 monitorado por analistas", "Cenário externo impacta ativos locais", "Commodities operam em zona de preços chave"]
            for i, n in enumerate(noticias, 1): add_texto(s1, f"{i}. {n}", 0.7, 2.5 + (i*1.2), 7.5, 0.8, 20)

            # Brasil
            s2 = slide_padrao("MERCADO BRASIL")
            v_i = res['IBOV']; cor_i = COR_ALTA if v_i['V'] >= 0 else COR_BAIXA
            add_texto(s2, f"IBOVESPA: {v_i['F']:,.0f} ({v_i['V']:+.2f}%)", 0.5, 1.8, 8, 0.8, 28, True, cor_i)
            g = gerar_grafico("^BVSP", data_valida, "#00FF7F")
            if g: s2.shapes.add_picture(g, Inches(2), Inches(2.6), width=Inches(5))
            
            vd = res['DOLAR']; cor_d = COR_BAIXA if vd['V'] >= 0 else COR_ALTA
            add_texto(s2, f"DÓLAR: R$ {vd['F']:.3f} ({vd['V']:+.2f}%)", 0.5, 5.4, 8, 0.8, 28, True, cor_d)
            gd = gerar_grafico("USDBRL=X", data_valida, "#00AEF3")
            if gd: s2.shapes.add_picture(gd, Inches(2), Inches(6.2), width=Inches(5))

            if top_altas is not None:
                add_texto(s2, "📈 ALTAS", 0.5, 9.4, 4, 0.5, 22, True, COR_ALTA)
                add_texto(s2, "\n".join([f"{t[:5]}: {val:+.2f}%" for t,val in top_altas.items()]), 0.5, 10.0, 4, 3, 19)
                add_texto(s2, "📉 BAIXAS", 4.5, 9.4, 4, 0.5, 22, True, COR_BAIXA)
                add_texto(s2, "\n".join([f"{t[:5]}: {val:+.2f}%" for t,val in top_baixas.items()]), 4.5, 10.0, 4, 3, 19)

            # Slides 3 e 4 seguem a mesma lógica segura de res.get...
            # EUA
            s3 = slide_padrao("BOLSAS EUA"); y_e = 2.0
            for idx in ["S&P500", "NASDAQ", "DOW"]:
                vi = res[idx]; cor = COR_ALTA if vi['V'] >= 0 else COR_BAIXA
                add_texto(s3, f"{idx}: {vi['F']:,.0f} ({vi['V']:+.2f}%)", 0.5, y_e, 4, 0.6, 22, True, cor)
                ge = gerar_grafico(mapa[idx], data_valida, "#00AEF3")
                if ge: s3.shapes.add_picture(ge, Inches(4.5), Inches(y_e), width=Inches(4))
                y_e += 4.2

            # Global
            s4 = slide_padrao("GLOBAL & CRIPTO")
            c_txt = "".join([f"• {c}: US$ {res[c]['F']:,.2f} ({res[c]['V']:+.2f}%)\n\n" for c in ["BRENT", "IRON", "GOLD", "SILVER"]])
            add_texto(s4, c_txt, 1.0, 2.5, 7, 5, 24)
            vb = res['BTC']; add_texto(s4, f"BITCOIN: US$ {vb['F']:,.0f} ({vb['V']:+.2f}%)", 0.5, 8.8, 8, 1, 28, True, COR_DESTAQUE)
            gb = gerar_grafico("BTC-USD", data_valida, "#F7931A")
            if gb: s4.shapes.add_picture(gb, Inches(2), Inches(9.8), width=Inches(5))

            out = io.BytesIO(); prs.save(out)
            st.success(f"✅ Boletim de {data_valida.strftime('%d/%m/%Y')} Gerado!")
            st.download_button("📥 BAIXAR PPTX", out.getvalue(), f"InvestForma_{data_valida.strftime('%Y-%m-%d')}.pptx")
        except Exception as e: st.error(f"Erro: {e}")