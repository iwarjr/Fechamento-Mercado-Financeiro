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

# --- CONFIGURAÇÕES DE DIRETÓRIO ---
DIR_BASE = os.path.dirname(os.path.abspath(__file__))

# --- IDENTIDADE VISUAL INVEST FORMA ---
COR_FUNDO = RGBColor(0, 32, 77)
COR_TEXTO = RGBColor(255, 255, 255)
COR_DESTAQUE = RGBColor(0, 174, 239)
COR_ALTA = RGBColor(0, 255, 127)
COR_BAIXA = RGBColor(255, 69, 0)

def carregar_logo_automatica():
    for arq in os.listdir(DIR_BASE):
        if arq.lower().startswith("logo") and arq.lower().endswith(('.png', '.jpg', '.jpeg')):
            try:
                with open(os.path.join(DIR_BASE, arq), "rb") as f:
                    return f.read()
            except: continue
    return None

def buscar_noticias_pt(dados_resumo):
    feed_url = "https://br.investing.com/rss/news_25.rss"
    try:
        f = feedparser.parse(feed_url)
        noticias = [e.title for e in f.entries[:10]]
        if len(noticias) >= 10: return noticias
    except: pass
    
    # Fallback caso o RSS falhe
    ibov_txt = f"Ibovespa opera próximo aos {dados_resumo.get('IBOV', {}).get('F', 0):,.0f} pontos."
    return [ibov_txt, "Dólar comercial reflete cenário fiscal e juros externos.", "Bolsas americanas operam em volatilidade.", "Criptoativos mantêm níveis de suporte estratégicos.", "Commodities reagem a dados de produção global.", "Mercado aguarda novas decisões de política monetária.", "Fluxo estrangeiro impacta volume financeiro na B3.", "Relatórios de inflação seguem no radar dos investidores.", "Nasdaq e S&P 500 apresentam variações mistas.", "Setor de energia e mineração movimenta o mercado."]

def gerar_grafico_intraday(ticker, data_sel, cor):
    try:
        s, e = data_sel.strftime('%Y-%m-%d'), (data_sel + timedelta(days=1)).strftime('%Y-%m-%d')
        df = yf.download(ticker, start=s, end=e, interval="5m", progress=False)['Close']
        if df.empty or len(df) < 2: return None
        
        plt.figure(figsize=(5, 2), facecolor='#00204D')
        ax = plt.axes(); ax.set_facecolor('#00204D')
        plt.plot(df.index, df.values, color=cor, linewidth=2.5)
        ax.tick_params(axis='both', colors='white', labelsize=8)
        for sp in ax.spines.values(): sp.set_color('white')
        plt.grid(True, color='grey', linestyle='--', alpha=0.1)
        plt.gcf().autofmt_xdate()
        img = io.BytesIO(); plt.savefig(img, format='png', bbox_inches='tight', dpi=120); plt.close()
        return img
    except: return None

def add_texto(slide, texto, left, top, width, height, size=18, bold=False, color=COR_TEXTO, align=PP_ALIGN.LEFT):
    tx = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tx.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = str(texto)
    p.font.size, p.font.bold, p.font.color.rgb, p.alignment = Pt(size), bold, color, align

def aplicar_estilo_base(slide, titulo, logo_bin):
    slide.background.fill.solid(); slide.background.fill.fore_color.rgb = COR_FUNDO
    if logo_bin: slide.shapes.add_picture(io.BytesIO(logo_bin), Inches(6.4), Inches(0.4), width=Inches(2.2))
    add_texto(slide, titulo, 0.5, 0.4, 6, 1, size=32, bold=True, color=COR_DESTAQUE)

# --- UI STREAMLIT ---
st.set_page_config(page_title="Invest Forma Academy", layout="wide")
st.title("💼 Dashboard Premium - Invest Forma Academy")

logo_data = carregar_logo_automatica()
data_sel = st.date_input("Selecione o Dia do Fechamento", datetime.now() - timedelta(days=1))

with st.spinner("Compilando dados..."):
    try:
        s_str, e_str = data_sel.strftime('%Y-%m-%d'), (data_sel + timedelta(days=1)).strftime('%Y-%m-%d')
        ativos_map = {"IBOV": "^BVSP", "DOLAR": "USDBRL=X", "S&P500": "^GSPC", "NASDAQ": "^IXIC", "DOW": "^DJI", "BTC": "BTC-USD", "ETH": "ETH-USD", "BRENT": "BZ=F", "IRON": "TIO=F", "GOLD": "GC=F", "SILVER": "SI=F"}
        
        res = {}
        for nome, ticker in ativos_map.items():
            try:
                d = yf.download(ticker, start=s_str, end=e_str, progress=False)
                if not d.empty:
                    ab, fe = float(d['Open'].iloc[0]), float(d['Close'].iloc[-1])
                    res[nome] = {"A": ab, "F": fe, "V": ((fe/ab)-1)*100, "T": ticker}
                else: res[nome] = {"A": 0, "F": 0, "V": 0, "T": ticker, "ERROR": True}
            except: res[nome] = {"A": 0, "F": 0, "V": 0, "T": ticker, "ERROR": True}

        # B3 Movers
        t_b3 = ['VALE3.SA', 'PETR4.SA', 'ITUB4.SA', 'BBDC4.SA', 'BBAS3.SA', 'ABEV3.SA', 'MGLU3.SA', 'WEGE3.SA', 'PRIO3.SA', 'GGBR4.SA']
        db3 = yf.download(t_b3, start=s_str, end=e_str, progress=False)['Close'].dropna()
        v_b3 = ((db3.iloc[-1] / db3.iloc[0]) - 1) * 100 if not db3.empty else None

        noticias_br = buscar_noticias_pt(res)
        noticias_finais = [st.text_input(f"Destaque {i+1}", value=noticias_br[i] if i < len(noticias_br) else "") for i in range(10)]

        if st.button("🌟 GERAR BOLETIM DE PERFORMANCE"):
            if not logo_data: st.error("Erro: Salve o arquivo 'logo.png' na pasta do app.")
            else:
                prs = Presentation(); prs.slide_width, prs.slide_height = Inches(9), Inches(16)

                # SLIDE 1: CAPA & NEWS
                s1 = prs.slides.add_slide(prs.slide_layouts[6]); aplicar_estilo_base(s1, "Boletim de Performance", logo_data)
                add_texto(s1, "Fechamento do Mercado Financeiro", 0.5, 0.9, 6, 0.5, 22, color=COR_TEXTO)
                add_texto(s1, data_sel.strftime('%d/%m/%Y'), 0.5, 1.3, 6, 0.4, 20, bold=True, color=COR_ALTA)
                y_news = 2.4
                for i, n in enumerate(noticias_finais, 1):
                    if n: add_texto(s1, f"{i}. {n}", 0.7, y_news, 7.5, 0.8, size=18); y_news += 1.15

                # SLIDE 2: BRASIL
                s2 = prs.slides.add_slide(prs.slide_layouts[6]); aplicar_estilo_base(s2, "MERCADO BRASIL", logo_data)
                v = res.get('IBOV', {"A":0, "F":0, "V":0})
                cor = COR_ALTA if v['V'] >= 0 else COR_BAIXA
                add_texto(s2, f"IBOVESPA: {v['F']:,.0f} ({v['V']:+.2f}%)", 0.5, 1.8, 8, 0.8, 28, True, cor)
                gi = gerar_grafico_intraday("^BVSP", data_sel, "#00FF7F")
                if gi: s2.shapes.add_picture(gi, Inches(2), Inches(2.6), width=Inches(5))
                vd = res.get('DOLAR', {"A":0, "F":0, "V":0})
                add_texto(s2, f"DÓLAR: R$ {vd['F']:.3f} ({vd['V']:+.2f}%)", 0.5, 5.4, 8, 0.8, 28, True, COR_BAIXA if vd['V'] >= 0 else COR_ALTA)
                gd = gerar_grafico_intraday("USDBRL=X", data_sel, "#00AEF3")
                if gd: s2.shapes.add_picture(gd, Inches(2), Inches(6.2), width=Inches(5))
                if v_b3 is not None:
                    add_texto(s2, "📈 ALTAS", 0.5, 9.4, 4, 0.5, 22, True, COR_ALTA)
                    add_texto(s2, "\n".join([f"{t[:5]}: {val:+.2f}%" for t,val in v_b3.nlargest(5).items()]), 0.5, 10.0, 4, 3, 19)
                    add_texto(s2, "📉 BAIXAS", 4.5, 9.4, 4, 0.5, 22, True, COR_BAIXA)
                    add_texto(s2, "\n".join([f"{t[:5]}: {val:+.2f}%" for t,val in v_b3.nsmallest(5).items()]), 4.5, 10.0, 4, 3, 19)

                # SLIDE 3: EUA
                s3 = prs.slides.add_slide(prs.slide_layouts[6]); aplicar_estilo_base(s3, "BOLSAS EUA", logo_data)
                y_e = 2.0
                for idx in ["S&P500", "NASDAQ", "DOW"]:
                    v = res.get(idx, {"A":0, "F":0, "V":0})
                    add_texto(s3, f"{idx}: {v['F']:,.0f} ({v['V']:+.2f}%)", 0.5, y_e, 4, 0.6, 22, True, COR_ALTA if v['V'] >= 0 else COR_BAIXA)
                    g = gerar_grafico_intraday(ativos_map[idx], data_sel, "#00AEF3")
                    if g: s3.shapes.add_picture(g, Inches(4.5), Inches(y_e), width=Inches(4))
                    y_e += 4.2

                # SLIDE 4: GLOBAL
                s4 = prs.slides.add_slide(prs.slide_layouts[6]); aplicar_estilo_base(s4, "GLOBAL & CRIPTO", logo_data)
                ct = "".join([f"• {c}: US$ {res.get(c, {}).get('F', 0):,.2f} ({res.get(c, {}).get('V', 0):+.2f}%)\n\n" for c in ["BRENT", "IRON", "GOLD", "SILVER"]])
                add_texto(s4, ct, 1.0, 2.5, 7, 5, 24)
                vb = res.get('BTC', {"F":0, "V":0})
                add_texto(s4, f"BITCOIN: US$ {vb['F']:,.0f} ({vb['V']:+.2f}%)", 0.5, 8.8, 8, 1, 28, True, COR_DESTAQUE)
                gb = gerar_grafico_intraday("BTC-USD", data_sel, "#F7931A")
                if gb: s4.shapes.add_picture(gb, Inches(2), Inches(9.8), width=Inches(5))
                ve = res.get('ETH', {"F":0, "V":0})
                add_texto(s4, f"ETH: US$ {ve['F']:,.0f} ({ve['V']:+.2f}%)", 0.5, 13.5, 8, 1, 24)

                out = io.BytesIO(); prs.save(out); st.success("✅ Boletim Gerado!"); st.download_button("📥 BAIXAR PPTX", out.getvalue(), f"InvestForma_{data_sel}.pptx")
    except Exception as e: st.error(f"Erro Crítico: {e}")