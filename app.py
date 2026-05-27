"""
CBZ Inspector v3 — Mobilny Asystent Kontroli BHP
Mobile-first, PWA-ready, Android/iOS
"""
import streamlit as st
import sqlite3, datetime, json, os, io, base64, smtplib, urllib.parse
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from pathlib import Path
from openai import OpenAI
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm

# ── CONFIG ────────────────────────────────────────────────────
st.set_page_config(
    page_title="CBZ Inspector",
    page_icon="🦺",
    layout="centered",
    initial_sidebar_state="collapsed",
)

DB   = Path("bhp_cbz.db")
DOCX = Path("szablon.docx")

try:
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
except Exception:
    st.error("Brak OPENAI_API_KEY w secrets.toml")
    st.stop()

# ── PWA — wstrzyknij manifest i meta tagi ─────────────────────
st.markdown("""
<link rel="manifest" href="data:application/json;base64,eyJuYW1lIjoiQ0JaIEluc3BlY3RvciIsInNob3J0X25hbWUiOiJDQloiLCJzdGFydF91cmwiOiIuLyIsImRpc3BsYXkiOiJzdGFuZGFsb25lIiwiYmFja2dyb3VuZF9jb2xvciI6IiMwZjJkNGUiLCJ0aGVtZV9jb2xvciI6IiMwZjJkNGUiLCJpY29ucyI6W3sic3JjIjoiZGF0YTppbWFnZS9zdmcreG1sLCUzQ3N2ZyB4bWxucz0naHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmcnIHZpZXdCb3g9JzAgMCAxMDAgMTAwJyUzRSUzQ3RleHQgeT0nLjllbScgZm9udC1zaXplPScxMDAnJTNF8J+mnCUzQy90ZXh0JTNFJTNDJTJGc3ZnJTNFIiwic2l6ZXMiOiIxNTJ4MTUyIn1dfQ==">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="CBZ Inspector">
<meta name="theme-color" content="#0f2d4e">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
""", unsafe_allow_html=True)

# ── CSS MOBILE-FIRST ──────────────────────────────────────────
st.markdown("""<style>
/* === RESET === */
*, *::before, *::after { box-sizing: border-box; }
:root {
  --bg:      #f0f4f8;
  --card:    #ffffff;
  --navy:    #0f2d4e;
  --navy2:   #1a3a5c;
  --orange:  #c85a1e;
  --orange2: #a8441a;
  --text:    #0d1b2a;
  --muted:   #5a7a96;
  --border:  #d0dcea;
  --green:   #1b6e3c;
  --red:     #b71c1c;
  --amber:   #b45309;
  --r:       12px;
  --rs:      8px;
}

/* === UKRYJ STREAMLIT CHROME === */
#MainMenu, footer, header,
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"] { display: none !important; }

/* === TŁO === */
.stApp, [data-testid="stAppViewContainer"],
[data-testid="stMain"], .main,
.block-container {
  background: var(--bg) !important;
  padding: 0 !important;
  max-width: 480px !important;
  margin: 0 auto !important;
}
.block-container { padding: 0 0 80px 0 !important; }

/* === SIDEBAR === */
[data-testid="stSidebar"] {
  background: var(--navy) !important;
  min-width: 240px !important;
  max-width: 260px !important;
}
[data-testid="stSidebar"] * { color: #c8dcf0 !important; }
[data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 { color: #fff !important; }
[data-testid="stSidebar"] .stButton button {
  width: 100% !important; background: rgba(255,255,255,.07) !important;
  border: 1px solid rgba(255,255,255,.12) !important; border-radius: 10px !important;
  color: #ddeaf5 !important; font-weight: 600 !important; font-size: 15px !important;
  padding: 12px 16px !important; text-align: left !important; margin-bottom: 6px !important;
}
[data-testid="stSidebar"] .stButton button:hover {
  background: rgba(200,90,30,.3) !important;
  color: #fff !important;
}

/* === TEKST === */
[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] li,
.stText p { color: var(--text) !important; font-size: 15px !important; }
label, [data-testid="stWidgetLabel"] p {
  color: var(--text) !important; font-weight: 600 !important; font-size: 14px !important;
}
h1 { color: var(--navy) !important; font-size: 22px !important; font-weight: 800 !important; margin: 16px 0 12px !important; }
h2 { color: var(--navy) !important; font-size: 18px !important; font-weight: 700 !important; }
h3 { color: var(--navy2) !important; font-size: 15px !important; font-weight: 700 !important; }

/* === FORMULARZE === */
[data-testid="stTextInput"] input,
[data-testid="stTextArea"] textarea,
[data-testid="stNumberInput"] input {
  background: var(--card) !important; border: 1.5px solid var(--border) !important;
  border-radius: var(--rs) !important; color: var(--text) !important;
  font-size: 15px !important; padding: 10px 14px !important;
}
[data-testid="stSelectbox"] > div > div {
  background: var(--card) !important; border: 1.5px solid var(--border) !important;
  border-radius: var(--rs) !important; color: var(--text) !important;
}

/* === PRZYCISKI === */
.stButton button {
  border-radius: 10px !important; font-weight: 600 !important;
  font-size: 15px !important; padding: 11px 18px !important;
  border: 1.5px solid var(--border) !important;
  background: var(--card) !important; color: var(--navy) !important;
  width: 100% !important; transition: all .15s !important;
  min-height: 48px !important;
}
.stButton button:hover {
  background: var(--navy) !important; color: #fff !important;
  border-color: var(--navy) !important;
}
.stButton button[kind="primary"] {
  background: var(--orange) !important; color: #fff !important;
  border: none !important; font-size: 16px !important;
  font-weight: 700 !important; min-height: 54px !important;
  border-radius: 12px !important;
}
.stButton button[kind="primary"]:hover { background: var(--orange2) !important; }

/* === EXPANDER === */
[data-testid="stExpander"] {
  background: var(--card) !important; border: 1px solid var(--border) !important;
  border-radius: var(--r) !important; margin-bottom: 8px !important;
}
details summary { color: var(--text) !important; font-weight: 600 !important; font-size: 15px !important; }
details summary span { color: var(--text) !important; }

/* === TABY === */
[data-testid="stTabs"] [role="tablist"] {
  background: var(--card) !important; border-radius: 10px !important;
  border: 1px solid var(--border) !important; padding: 4px !important; gap: 4px !important;
}
[data-testid="stTabs"] [role="tab"] {
  border-radius: 8px !important; color: var(--muted) !important;
  font-weight: 600 !important; font-size: 14px !important;
  border: none !important; padding: 8px 14px !important;
}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {
  background: var(--navy) !important; color: #fff !important;
}

/* === KAMERA === */
[data-testid="stCameraInput"] { border-radius: var(--r) !important; overflow: hidden !important; }
[data-testid="stCameraInput"] video,
[data-testid="stCameraInput"] img {
  width: 100% !important; max-height: 55vh !important;
  object-fit: cover !important; border-radius: var(--r) !important;
}
[data-testid="stCameraInput"] button {
  min-height: 52px !important; font-size: 16px !important;
  font-weight: 700 !important; border-radius: 10px !important;
}

/* === INFO BOXY === */
[data-testid="stAlert"] { border-radius: 10px !important; font-size: 14px !important; }

/* === METRIC === */
[data-testid="stMetric"] {
  background: var(--card) !important; border: 1px solid var(--border) !important;
  border-radius: var(--r) !important; padding: 12px 16px !important;
}
[data-testid="stMetricLabel"] p { color: var(--muted) !important; font-size: 12px !important; font-weight: 600 !important; }
[data-testid="stMetricValue"] { color: var(--navy) !important; font-size: 28px !important; font-weight: 800 !important; }

/* === DOLNY PASEK NAWIGACJI === */
.nav-bar {
  display: flex; background: var(--navy);
  border-radius: 16px; margin: 0 0 16px 0;
  overflow: hidden;
}
.nav-item {
  flex: 1; padding: 12px 4px 10px; text-align: center;
  cursor: pointer; border: none; background: transparent;
  color: rgba(255,255,255,.5); font-size: 10px; font-weight: 600;
  text-transform: uppercase; letter-spacing: .05em;
  transition: all .15s; display: flex; flex-direction: column;
  align-items: center; gap: 4px;
}
.nav-item .ni { font-size: 22px; line-height: 1; }
.nav-item.active { color: #fff; background: rgba(200,90,30,.35); }
.nav-item:hover  { color: #fff; }

/* === KARTY BHP === */
.stat-row {
  display: grid; grid-template-columns: 1fr 1fr;
  gap: 10px; margin-bottom: 14px;
}
.stat-box {
  background: var(--card); border-radius: var(--r);
  border: 1px solid var(--border); padding: 14px 16px;
  text-align: center;
}
.stat-num  { font-size: 32px; font-weight: 800; color: var(--navy); line-height: 1; }
.stat-lbl  { font-size: 11px; color: var(--muted); text-transform: uppercase; letter-spacing: .06em; margin-top: 4px; }
.stat-warn { color: var(--orange); }

.action-card {
  background: var(--navy); border-radius: var(--r);
  padding: 20px 20px; margin-bottom: 10px; cursor: pointer;
}
.action-card.orange { background: var(--orange); }
.action-card .ac-icon { font-size: 28px; margin-bottom: 8px; }
.action-card .ac-title { color: #fff; font-size: 18px; font-weight: 800; line-height: 1.2; margin-bottom: 4px; }
.action-card .ac-sub   { color: rgba(255,255,255,.65); font-size: 13px; }

.section-title {
  font-size: 12px; font-weight: 700; color: var(--muted);
  text-transform: uppercase; letter-spacing: .08em;
  margin: 20px 0 10px; padding: 0 2px;
}

.list-item {
  background: var(--card); border-radius: var(--r);
  border: 1px solid var(--border); padding: 14px 16px;
  margin-bottom: 8px; display: flex; align-items: center; gap: 12px;
}
.li-icon  { font-size: 22px; flex-shrink: 0; }
.li-body  { flex: 1; min-width: 0; }
.li-title { font-size: 15px; font-weight: 700; color: var(--text); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.li-sub   { font-size: 12px; color: var(--muted); margin-top: 2px; }
.li-badge { flex-shrink: 0; font-size: 11px; font-weight: 700; padding: 3px 10px; border-radius: 20px; }
.badge-k  { background: #fce8e8; color: var(--red); }
.badge-w  { background: #fff3e0; color: #bf360c; }
.badge-s  { background: #fffde7; color: var(--amber); }
.badge-n  { background: #e8f5e9; color: var(--green); }

hr { border: none; border-top: 1px solid var(--border) !important; margin: 16px 0 !important; }
</style>""", unsafe_allow_html=True)

# ── BAZA DANYCH ────────────────────────────────────────────────
def init_db():
    con = sqlite3.connect(DB)
    con.executescript("""
        CREATE TABLE IF NOT EXISTS firmy (
            id INTEGER PRIMARY KEY AUTOINCREMENT, nazwa TEXT NOT NULL UNIQUE,
            nip TEXT, adres TEXT, kontakt TEXT, email TEXT, uwagi TEXT,
            created_at TEXT DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS protokoly (
            id INTEGER PRIMARY KEY AUTOINCREMENT, firma_id INTEGER,
            nr TEXT, data TEXT, miejsce TEXT, obszar TEXT,
            kontrolujacy TEXT DEFAULT 'Michał Młynarczak',
            osoby TEXT, rodzaj TEXT DEFAULT 'Bieżąca kontrola warunków pracy',
            ograniczenia TEXT DEFAULT 'Brak', status TEXT DEFAULT 'szkic',
            created_at TEXT DEFAULT (datetime('now','localtime'))
        );
        CREATE TABLE IF NOT EXISTS ustalenia (
            id INTEGER PRIMARY KEY AUTOINCREMENT, protokol_id INTEGER,
            obszar TEXT, kategoria TEXT, priorytet TEXT, status TEXT DEFAULT 'Nowe',
            odpowiedzialny TEXT, termin TEXT, ryzyko TEXT, opis TEXT, zalecenie TEXT,
            foto BLOB, created_at TEXT DEFAULT (datetime('now','localtime'))
        );
    """)
    con.commit(); con.close()

init_db()

def db(): return sqlite3.connect(DB)

def q(sql, p=(), one=False):
    with db() as c:
        r = c.execute(sql, p)
        return r.fetchone() if one else r.fetchall()

def qw(sql, p=()):
    with db() as c: c.execute(sql, p); c.commit()

# Pomocnicze
def firmy():     return q("SELECT id,nazwa,nip,adres,kontakt,email FROM firmy ORDER BY nazwa")
def protokoly(fid=None, st_=None):
    s = "SELECT p.id,f.nazwa,p.nr,p.data,p.miejsce,p.status,p.kontrolujacy,(SELECT COUNT(*) FROM ustalenia u WHERE u.protokol_id=p.id) FROM protokoly p JOIN firmy f ON f.id=p.firma_id WHERE 1=1"
    params = []
    if fid:  s+=" AND p.firma_id=?"; params.append(fid)
    if st_:  s+=" AND p.status=?";   params.append(st_)
    return q(s+" ORDER BY p.data DESC", params)
def ustalenia(pid): return q("SELECT * FROM ustalenia WHERE protokol_id=? ORDER BY id",(pid,))
def otwarte():
    return q("""SELECT u.id,f.nazwa,p.nr,p.data,u.obszar,u.priorytet,u.odpowiedzialny,
                       u.termin,u.ryzyko,u.zalecenie,u.status
               FROM ustalenia u JOIN protokoly p ON p.id=u.protokol_id JOIN firmy f ON f.id=p.firma_id
               WHERE u.status!='Zamknięte' ORDER BY u.priorytet DESC,u.termin""")

# ── AI ─────────────────────────────────────────────────────────
def analizuj(foto_bytes):
    b64 = base64.standard_b64encode(foto_bytes).decode()
    r = client.chat.completions.create(
        model="gpt-4o",
        response_format={"type":"json_object"},
        max_tokens=350,
        temperature=0.2,
        messages=[
            {"role":"system","content":"""Jesteś inspektorem BHP. Analizuj zdjęcie i zwróć JSON:
{"zagrozenie":"krótka nazwa (max 8 słów)","opis_niezgodnosci":"opis stanu 2-3 zdania","zalecenie":"zalecenie pokontrolne 2-3 zdania"}
Jeśli brak zagrożeń — podaj zalecenie profilaktyczne."""},
            {"role":"user","content":[
                {"type":"text","text":"Przeanalizuj zdjęcie z kontroli BHP."},
                {"type":"image_url","image_url":{"url":f"data:image/jpeg;base64,{b64}"}}
            ]}
        ]
    )
    raw = r.choices[0].message.content.strip().replace("```json","").replace("```","")
    return json.loads(raw)

# ── GENERUJ WORD ───────────────────────────────────────────────
def gen_word(pid):
    p = q("SELECT p.*,f.nazwa,f.nip,f.adres,f.kontakt,f.email FROM protokoly p JOIN firmy f ON f.id=p.firma_id WHERE p.id=?",(pid,), one=True)
    ust = ustalenia(pid)
    doc = DocxTemplate(str(DOCX))
    pts = []
    for u in ust:
        kp = {"obszar":u[2],"kategoria":u[3],"priorytet":u[4],"status":u[5],"odpowiedzialny":u[6],
              "termin":u[7],"ryzyko":u[8],"opis_stanu":u[9],"zalecenie":u[10],
              "dzialanie_pilne":"Tak" if u[4] in("Wysoki","Krytyczny") else "Nie",
              "wymaga_weryfikacji":"Tak","podpis_zdjecia":"Fot." if u[11] else "Brak"}
        if u[11]: kp["zdjecie"] = InlineImage(doc, io.BytesIO(u[11]), width=Mm(100))
        else: kp["zdjecie"] = ""
        pts.append(kp)
    wk = sum(1 for u in ust if u[4] in("Wysoki","Krytyczny"))
    ctx = {
        "nr_protokolu":p[2],"data_kontroli":p[3],"klient_nazwa":p[11],"klient_nip":p[12] or "",
        "klient_adres":p[13] or "","miejsce_kontroli":p[4] or "","obszar_kontroli":p[5] or "",
        "rodzaj_kontroli":p[8],"data_i_godzina_kontroli":p[3],"kontrolujacy":p[6],
        "osoby_uczestniczace":p[7] or "","osoby_kontrolowane":"-","dokumenty_odniesienia":"Ustalenia",
        "ograniczenia_kontroli":p[9] or "Brak","adresaci_email":p[15] or "-",
        "zakres_kontroli":"BHP i PPOŻ","obszary_sprawdzenia":p[5] or "wg kart",
        "liczba_ustalen":str(len(pts)),"liczba_wysokich_krytycznych":str(wk),
        "obszary_wymagajace_dzialan":"-","termin_przegladu_zalecen":"-",
        "wnioski_koncowe":"Szczegóły w rejestrze zaleceń.",
        "przedstawiciel_zakladu":p[7] or "-","koordynator_realizacji":"-",
        "zatwierdzajacy":p[6],"data_podpisu":p[3],"punkty":pts
    }
    doc.render(ctx)
    bio = io.BytesIO(); doc.save(bio); return bio.getvalue()

# ── GENERUJ PDF ────────────────────────────────────────────────
CN = colors.HexColor("#0f2d4e")
CO = colors.HexColor("#c85a1e")
CL = colors.HexColor("#f0f4f8")
CB = colors.HexColor("#d0dcea")

def gen_pdf(pid):
    p = q("SELECT p.*,f.nazwa,f.nip,f.adres,f.kontakt,f.email FROM protokoly p JOIN firmy f ON f.id=p.firma_id WHERE p.id=?",(pid,),one=True)
    ust = ustalenia(pid)
    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio,pagesize=A4,leftMargin=18*mm,rightMargin=18*mm,topMargin=16*mm,bottomMargin=16*mm)
    ss = getSampleStyleSheet()
    W = doc.width
    def sty(nm,**kw):
        s = ParagraphStyle(nm,parent=ss["Normal"],fontName="Helvetica",fontSize=9,leading=12)
        for k,v in kw.items(): setattr(s,k,v)
        return s
    sh = sty("h",fontName="Helvetica-Bold",fontSize=13,textColor=colors.white)
    sb = sty("b",fontSize=9,leading=13)
    sl = sty("l",fontSize=8,textColor=colors.HexColor("#5a7a96"))
    sc = sty("c",fontName="Helvetica-Bold",fontSize=10,textColor=colors.white)

    def hdr(t):
        tbl = Table([[Paragraph(t,sc)]],[W])
        tbl.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),CN),("TOPPADDING",(0,0),(-1,-1),8),
            ("BOTTOMPADDING",(0,0),(-1,-1),8),("LEFTPADDING",(0,0),(-1,-1),10)]))
        return tbl
    def row2(l1,v1,l2="",v2=""):
        if l2:
            d=[[Paragraph(f"<b>{l1}</b>",sl),Paragraph(str(v1 or ""),sb),Paragraph(f"<b>{l2}</b>",sl),Paragraph(str(v2 or ""),sb)]]
            t=Table(d,[W*.22,W*.28,W*.22,W*.28])
        else:
            d=[[Paragraph(f"<b>{l1}</b>",sl),Paragraph(str(v1 or ""),sb)]]
            t=Table(d,[W*.25,W*.75])
        t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),CL),("GRID",(0,0),(-1,-1),.3,CB),
            ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),("LEFTPADDING",(0,0),(-1,-1),6)]))
        return t

    els = []
    # Nagłówek
    nt = Table([[Paragraph("PROTOKÓŁ BHP / PPOŻ",ParagraphStyle("pt",fontName="Helvetica-Bold",fontSize=15,textColor=colors.white)),
                 Paragraph(f"Nr: {p[2]}<br/>Data: {p[3]}",ParagraphStyle("nr",fontName="Helvetica",fontSize=9,textColor=colors.HexColor("#a8c4e0"),alignment=2))]],[W*.6,W*.4])
    nt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),CN),("TOPPADDING",(0,0),(-1,-1),14),
        ("BOTTOMPADDING",(0,0),(-1,-1),14),("LEFTPADDING",(0,0),(0,0),14),("RIGHTPADDING",(-1,-1),(-1,-1),14),("VALIGN",(0,0),(-1,-1),"MIDDLE")]))
    els += [nt,Spacer(1,5*mm),hdr("1. Dane kontroli"),
            row2("Klient",p[11],"NIP",p[12]),row2("Adres",p[13],"Miejsce",p[4]),
            row2("Kontrolujący",p[6],"Osoby",p[7]),Spacer(1,5*mm),hdr("2. Ustalenia")]

    BG = {"Krytyczny":colors.HexColor("#fce8e8"),"Wysoki":colors.HexColor("#fff3e0"),
          "Średni":colors.HexColor("#fffde7"),"Niski":colors.HexColor("#e8f5e9")}
    for i,u in enumerate(ust,1):
        bg = BG.get(u[4],colors.white)
        hh = Table([[Paragraph(f"Ustalenie {i} · {u[3]} · {u[2]}",sty("uh",fontName="Helvetica-Bold",fontSize=9,textColor=CN)),
                     Paragraph(f"<b>{u[4].upper()}</b>",sty("up",fontName="Helvetica-Bold",fontSize=9,textColor=CO,alignment=2))]],[W*.6,W*.4])
        hh.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),bg),("GRID",(0,0),(-1,-1),.3,CB),
            ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),("LEFTPADDING",(0,0),(0,0),8),("RIGHTPADDING",(-1,-1),(-1,-1),8)]))
        body_items = [Paragraph(f"<b>Zagrożenie:</b> {u[8] or ''}",sb),Spacer(1,2*mm),
                      Paragraph("<b>Stan faktyczny:</b>",sl),Paragraph(u[9] or "",sb),
                      Spacer(1,2*mm),Paragraph("<b>Zalecenie:</b>",sl),Paragraph(u[10] or "",sb)]
        if u[11]:
            ct = Table([[RLImage(io.BytesIO(u[11]),48*mm,36*mm),body_items]],[50*mm,W-50*mm])
        else:
            ct = Table([[body_items]],[W])
        ct.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),colors.white),("GRID",(0,0),(-1,-1),.3,CB),
            ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),("LEFTPADDING",(0,0),(-1,-1),8),("VALIGN",(0,0),(-1,-1),"TOP")]))
        mt = Table([[Paragraph(f"Odp: {u[6] or '–'}",sl),Paragraph(f"Termin: {u[7] or '–'}",sl),Paragraph(f"Status: {u[5]}",sl)]],[W*.4,W*.3,W*.3])
        mt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),CL),("GRID",(0,0),(-1,-1),.3,CB),
            ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),("LEFTPADDING",(0,0),(-1,-1),8)]))
        els += [hh,ct,mt,Spacer(1,4*mm)]

    # Tabela zbiorcza
    els += [Spacer(1,4*mm),hdr("3. Rejestr zaleceń")]
    td = [["Lp.","Obszar","Niezgodność","Zalecenie","Odp.","Termin","Prior.","Status"]]
    for i,u in enumerate(ust,1):
        td.append([str(i),u[2][:20] or "",(u[9] or "")[:40],(u[10] or "")[:40],u[6][:15] or "",u[7] or "",u[4],u[5]])
    tt = Table(td,[8*mm,25*mm,50*mm,50*mm,22*mm,18*mm,16*mm,16*mm])
    tt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),CN),("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),7.5),
        ("GRID",(0,0),(-1,-1),.3,CB),("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,CL]),
        ("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3),("LEFTPADDING",(0,0),(-1,-1),4),("VALIGN",(0,0),(-1,-1),"TOP")]))
    els.append(tt)
    doc.build(els)
    return bio.getvalue()

# ── EMAIL ──────────────────────────────────────────────────────
def wyslij(do, temat, tresc, att_bytes, att_name):
    try:
        host=st.secrets.get("SMTP_HOST",""); port=int(st.secrets.get("SMTP_PORT",587))
        user=st.secrets.get("SMTP_USER",""); pwd=st.secrets.get("SMTP_PASS","")
    except: return False,"Brak konfiguracji SMTP"
    if not host: return False,"Uzupełnij SMTP w secrets.toml"
    msg = MIMEMultipart(); msg["From"]=user; msg["To"]=do; msg["Subject"]=temat
    msg.attach(MIMEText(tresc,"plain","utf-8"))
    p = MIMEBase("application","octet-stream"); p.set_payload(att_bytes)
    encoders.encode_base64(p); p.add_header("Content-Disposition",f'attachment; filename="{att_name}"')
    msg.attach(p)
    try:
        with smtplib.SMTP(host,port) as s: s.starttls(); s.login(user,pwd); s.send_message(msg)
        return True,"OK"
    except Exception as e: return False,str(e)

# ── SESSION STATE ──────────────────────────────────────────────
for k,v in [("ekran","menu"),("protokol_id",None),("cam_key",0),("podekran","metryczka")]:
    if k not in st.session_state: st.session_state[k] = v

# ── SIDEBAR ───────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🦺 CBZ Inspector")
    st.markdown("---")
    for lbl, ekr in [("🏠  Menu","menu"),("🏢  Firmy","firmy"),
                      ("📋  Nowy protokół","nowy"),("📁  Archiwum","archiwum"),
                      ("🔄  Rekontrola","rekontrola")]:
        if st.button(lbl, key=f"sb_{ekr}"):
            st.session_state.ekran = ekr
            st.session_state.podekran = "metryczka"
            st.rerun()

# ── PASEK NAWIGACJI (bottom nav jako HTML) ─────────────────────
def nav_bar():
    e = st.session_state.ekran
    def cls(n): return "nav-item active" if e==n else "nav-item"
    items = [
        ("menu",   "🏠", "Menu"),
        ("nowy",   "📋", "Kontrola"),
        ("archiwum","📁","Archiwum"),
        ("rekontrola","🔄","Rekontrola"),
        ("firmy",  "🏢", "Firmy"),
    ]
    btns = "".join(f'<button class="{cls(k)}" onclick="window.parent.document.querySelector(\'[data-testid=\"stSidebar\"] [data-testid=\"stButton\"]:nth-child({i+1})\')?.click()"><span class="ni">{ic}</span>{lbl}</button>'
                   for i,(k,ic,lbl) in enumerate(items))
    st.markdown(f'<div class="nav-bar">{btns}</div>', unsafe_allow_html=True)

# ── HELPER: karta ──────────────────────────────────────────────
def karta(html): st.markdown(f'<div style="background:#fff;border-radius:12px;border:1px solid #d0dcea;padding:16px;margin-bottom:10px">{html}</div>', unsafe_allow_html=True)
def section(t):  st.markdown(f'<div class="section-title">{t}</div>', unsafe_allow_html=True)
def badge(p):
    m = {"Krytyczny":"badge-k","Wysoki":"badge-w","Średni":"badge-s","Niski":"badge-n"}
    return f'<span class="li-badge {m.get(p,"badge-n")}">{p}</span>'

def go(ekr, **kw):
    st.session_state.ekran = ekr
    for k,v in kw.items(): st.session_state[k] = v
    st.rerun()

# ══════════════════════════════════════════════════════════════
# EKRAN: MENU
# ══════════════════════════════════════════════════════════════
if st.session_state.ekran == "menu":
    pc = len(protokoly())
    fc = len(firmy())
    oc = len(otwarte())
    pz = len(protokoly(st_="zamknięty"))

    # Hero header
    st.markdown(f"""
<div style="background:linear-gradient(160deg,#0f2d4e 0%,#1a3a5c 100%);
            border-radius:16px;padding:22px 20px 20px;margin-bottom:16px">
  <div style="color:#fff;font-size:11px;font-weight:700;letter-spacing:.1em;
              text-transform:uppercase;opacity:.6;margin-bottom:4px">CBZ · Centrum Bezpiecznego Zatrudnienia</div>
  <div style="color:#fff;font-size:24px;font-weight:800;line-height:1.1;margin-bottom:16px">
    Asystent kontroli BHP 🦺
  </div>
  <div style="display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:8px">
    <div style="background:rgba(255,255,255,.1);border-radius:10px;padding:10px 6px;text-align:center">
      <div style="font-size:22px;font-weight:800;color:#fff;line-height:1">{pc}</div>
      <div style="font-size:9px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:.06em;margin-top:3px">Protokołów</div>
    </div>
    <div style="background:rgba(255,255,255,.1);border-radius:10px;padding:10px 6px;text-align:center">
      <div style="font-size:22px;font-weight:800;color:#fff;line-height:1">{pz}</div>
      <div style="font-size:9px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:.06em;margin-top:3px">Zamkniętych</div>
    </div>
    <div style="background:rgba(255,255,255,.1);border-radius:10px;padding:10px 6px;text-align:center">
      <div style="font-size:22px;font-weight:800;color:#fff;line-height:1">{fc}</div>
      <div style="font-size:9px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:.06em;margin-top:3px">Firm</div>
    </div>
    <div style="background:rgba(255,255,255,{'.2' if oc>0 else '.1'});
                border-radius:10px;padding:10px 6px;text-align:center;
                {'border:1px solid rgba(200,90,30,.7)' if oc>0 else ''}">
      <div style="font-size:22px;font-weight:800;color:{'#f5a623' if oc>0 else '#fff'};line-height:1">{oc}</div>
      <div style="font-size:9px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:.06em;margin-top:3px">Zalecenia</div>
    </div>
  </div>
</div>""", unsafe_allow_html=True)

    # Główny przycisk
    if st.button("📋  Zacznij nową kontrolę →", type="primary", use_container_width=True):
        go("nowy", podekran="metryczka")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    # Kafelki 2x2
    c1, c2 = st.columns(2)
    with c1:
        karta(f'<div style="font-size:24px;margin-bottom:8px">📁</div><div style="font-size:15px;font-weight:700;color:#0f2d4e">Archiwum</div><div style="font-size:12px;color:#5a7a96;margin-top:3px">{pc} protokołów</div>')
        if st.button("Otwórz", key="m_a", use_container_width=True): go("archiwum")
    with c2:
        karta(f'<div style="font-size:24px;margin-bottom:8px">🔄</div><div style="font-size:15px;font-weight:700;color:#0f2d4e">Rekontrola</div><div style="font-size:12px;color:{"#c85a1e" if oc>0 else "#5a7a96"};margin-top:3px">{"⚠ "+str(oc)+" otwartych" if oc>0 else "Brak otwartych"}</div>')
        if st.button("Otwórz", key="m_r", use_container_width=True): go("rekontrola")
    c3, c4 = st.columns(2)
    with c3:
        karta(f'<div style="font-size:24px;margin-bottom:8px">🏢</div><div style="font-size:15px;font-weight:700;color:#0f2d4e">Firmy</div><div style="font-size:12px;color:#5a7a96;margin-top:3px">{fc} w bazie</div>')
        if st.button("Otwórz", key="m_f", use_container_width=True): go("firmy")
    with c4:
        karta('<div style="font-size:24px;margin-bottom:8px">ℹ️</div><div style="font-size:15px;font-weight:700;color:#0f2d4e">O aplikacji</div><div style="font-size:12px;color:#5a7a96;margin-top:3px">CBZ Inspector v3</div>')
        st.markdown('<div style="font-size:12px;color:#5a7a96;padding:4px 0">Michał Młynarczak<br>centrum-bhp.com.pl</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# EKRAN: FIRMY
# ══════════════════════════════════════════════════════════════
elif st.session_state.ekran == "firmy":
    col_back, col_title = st.columns([1, 4])
    with col_back:
        if st.button("← Wróć"): go("menu")
    with col_title:
        st.markdown("## 🏢 Firmy")

    with st.expander("➕ Dodaj / edytuj firmę", expanded=not bool(firmy())):
        with st.form("ff"):
            fn  = st.text_input("Nazwa firmy *")
            c1,c2 = st.columns(2)
            with c1: fnip = st.text_input("NIP"); feml = st.text_input("E-mail")
            with c2: fkon = st.text_input("Kontakt"); fadr = st.text_area("Adres",height=80)
            if st.form_submit_button("💾 Zapisz", use_container_width=True):
                if fn:
                    qw("INSERT INTO firmy(nazwa,nip,adres,kontakt,email) VALUES(?,?,?,?,?) ON CONFLICT(nazwa) DO UPDATE SET nip=excluded.nip,adres=excluded.adres,kontakt=excluded.kontakt,email=excluded.email",(fn,fnip,fadr,fkon,feml))
                    st.success(f"✅ {fn}"); st.rerun()
                else: st.error("Podaj nazwę")

    fl = firmy()
    if fl:
        section(f"{len(fl)} firm w bazie")
        for fid, fnaz, fnip, fadr, fkon, feml in fl:
            with st.expander(f"🏢 {fnaz}"):
                c1,c2 = st.columns(2)
                with c1: st.write(f"**NIP:** {fnip or '–'}"); st.write(f"**Adres:** {fadr or '–'}")
                with c2: st.write(f"**Kontakt:** {fkon or '–'}"); st.write(f"**E-mail:** {feml or '–'}")
                if st.button("📋 Nowy protokół", key=f"fp_{fid}", use_container_width=True):
                    st.session_state["firma_id_preset"] = fid; go("nowy", podekran="metryczka")
    else:
        st.info("Brak firm. Dodaj pierwszą firmę powyżej.")

# ══════════════════════════════════════════════════════════════
# EKRAN: NOWY PROTOKÓŁ
# ══════════════════════════════════════════════════════════════
elif st.session_state.ekran == "nowy":

    # ── METRYCZKA ──────────────────────────────────────────────
    if st.session_state.podekran == "metryczka":
        col_back, col_title = st.columns([1, 4])
        with col_back:
            if st.button("← Wróć"): go("menu")
        with col_title:
            st.markdown("## 📋 Nowa kontrola")

        fl = firmy()
        nazwy = [f[1] for f in fl]
        preset_id = st.session_state.get("firma_id_preset")
        idx = 0
        if preset_id:
            ids = [f[0] for f in fl]
            if preset_id in ids: idx = ids.index(preset_id) + 1

        wyb = st.selectbox("Firma *", ["– wybierz –"] + nazwy, index=idx)
        if wyb == "– wybierz –":
            st.info("Wybierz firmę lub dodaj ją w module Firmy.")
            if st.button("➕ Dodaj firmę", use_container_width=True): go("firmy")
        else:
            fr = next(f for f in fl if f[1]==wyb)
            fid = fr[0]
            rok = datetime.date.today().year; mies = datetime.date.today().month
            c1,c2 = st.columns(2)
            with c1:
                nr   = st.text_input("Nr protokołu", value=f"CBZ/BHP/{rok}/{mies:02d}/001")
                data = st.date_input("Data", datetime.date.today())
                kontr= st.text_input("Kontrolujący", value="Michał Młynarczak")
            with c2:
                miej = st.text_input("Miejsce / wydział", placeholder="np. Hala W-2")
                obs  = st.text_input("Obszar", placeholder="np. Spawalnia")
                osob = st.text_input("Osoby obecne", value=fr[4] or "")
            rod  = st.text_input("Rodzaj kontroli", value="Bieżąca kontrola warunków pracy")
            ogr  = st.text_input("Ograniczenia", value="Brak")

            if st.button("✅ Utwórz i przejdź do ustaleń →", type="primary", use_container_width=True):
                qw("INSERT INTO protokoly(firma_id,nr,data,miejsce,obszar,kontrolujacy,osoby,rodzaj,ograniczenia) VALUES(?,?,?,?,?,?,?,?,?)",
                   (fid,nr,str(data),miej,obs,kontr,osob,rod,ogr))
                pid = q("SELECT last_insert_rowid()",one=True)[0]
                st.session_state.protokol_id = pid
                st.session_state["firma_id_preset"] = None
                st.session_state.podekran = "ustalenia"
                st.rerun()

    # ── USTALENIA ──────────────────────────────────────────────
    elif st.session_state.podekran == "ustalenia":
        pid   = st.session_state.protokol_id
        prow  = q("SELECT p.*,f.nazwa FROM protokoly p JOIN firmy f ON f.id=p.firma_id WHERE p.id=?",(pid,),one=True)
        ust_l = ustalenia(pid)

        col_back, col_title = st.columns([1, 4])
        with col_back:
            if st.button("← Menu"): go("menu")
        with col_title:
            st.markdown(f"## 📸 {prow[12]}")
        st.caption(f"{prow[2]}  ·  {prow[3]}  ·  {len(ust_l)} ustaleń")

        # Liczniki kompaktowe
        c1,c2,c3 = st.columns(3)
        wys = sum(1 for u in ust_l if u[4] in("Wysoki","Krytyczny"))
        c1.metric("Ustaleń", len(ust_l))
        c2.metric("Wysoki/Krytyczy", wys)
        c3.metric("Status", prow[10].upper())
        st.markdown("---")

        # KAMERA
        st.markdown("### 📷 Zdjęcie")
        if f"ai_ok_{st.session_state.cam_key}" not in st.session_state:
            st.session_state[f"ai_ok_{st.session_state.cam_key}"] = False

        foto = st.camera_input("Zrób zdjęcie — AI automatycznie przeanalizuje",
                                key=f"cam_{st.session_state.cam_key}")

        if foto and not st.session_state.get(f"ai_ok_{st.session_state.cam_key}"):
            with st.spinner("🤖 Analizuję zdjęcie..."):
                try:
                    ai = analizuj(foto.getvalue())
                    st.session_state["ai_z"] = ai.get("zagrozenie","")
                    st.session_state["ai_o"] = ai.get("opis_niezgodnosci","")
                    st.session_state["ai_zal"] = ai.get("zalecenie","")
                    st.session_state[f"ai_ok_{st.session_state.cam_key}"] = True
                    st.success("✅ Analiza gotowa — sprawdź i zatwierdź pola")
                except Exception as ex:
                    st.warning(f"AI niedostępne: {ex}")

        st.markdown("### ✏️ Dane ustalenia")
        c1,c2 = st.columns(2)
        with c1:
            ob_u = st.text_input("Lokalizacja", placeholder="np. Tokarka nr 3")
            kat  = st.selectbox("Kategoria",["BHP","PPOŻ","Maszyny","Drogi/5S","Chemia","ŚOI","Inne"])
            pri  = st.selectbox("Priorytet",["Niski","Średni","Wysoki","Krytyczny"])
        with c2:
            odp  = st.text_input("Odpowiedzialny", placeholder="Mistrz zmiany")
            ter  = st.date_input("Termin", datetime.date.today()+datetime.timedelta(days=7))
            sta  = st.selectbox("Status",["Nowe","W trakcie","Zamknięte"])

        ryz = st.text_input("Zagrożenie",       value=st.session_state.get("ai_z",""))
        opi = st.text_area("Stan faktyczny",     value=st.session_state.get("ai_o",""), height=100)
        zal = st.text_area("Zalecenie",          value=st.session_state.get("ai_zal",""), height=100)

        if st.button("➕ Zapisz ustalenie", type="primary", use_container_width=True):
            if opi and zal:
                fb = foto.getvalue() if foto else None
                qw("INSERT INTO ustalenia(protokol_id,obszar,kategoria,priorytet,status,odpowiedzialny,termin,ryzyko,opis,zalecenie,foto) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                   (pid,ob_u,kat,pri,sta,odp,str(ter),ryz,opi,zal,fb))
                ck = st.session_state.cam_key
                st.session_state.cam_key = ck+1
                for k in ["ai_z","ai_o","ai_zal",f"ai_ok_{ck}"]: st.session_state.pop(k,None)
                st.success("✅ Zapisano!"); st.rerun()
            else:
                st.error("Uzupełnij 'Stan faktyczny' i 'Zalecenie'")

        # Lista dodanych
        if ust_l:
            st.markdown("---")
            section(f"Dodane ustalenia ({len(ust_l)})")
            for u in ust_l:
                with st.expander(f"{badge(u[4])} {u[2] or 'Ustalenie'} — {u[3]}", expanded=False):
                    st.write(f"**Zagrożenie:** {u[8]}")
                    st.write(f"**Opis:** {u[9]}")
                    st.write(f"**Zalecenie:** {u[10]}")
                    if u[11]: st.image(u[11], width=260)

        st.markdown("---")
        c1,c2 = st.columns(2)
        with c1:
            if st.button("🏁 Zamknij i eksportuj", use_container_width=True):
                qw("UPDATE protokoly SET status='zamknięty' WHERE id=?",(pid,))
                st.session_state.podekran = "eksport"; st.rerun()
        with c2:
            if st.button("💾 Zapisz szkic", use_container_width=True): go("menu")

    # ── EKSPORT ────────────────────────────────────────────────
    elif st.session_state.podekran == "eksport":
        pid  = st.session_state.protokol_id
        prow = q("SELECT p.*,f.nazwa,f.nip,f.adres,f.kontakt,f.email FROM protokoly p JOIN firmy f ON f.id=p.firma_id WHERE p.id=?",(pid,),one=True)
        ust_l = ustalenia(pid)

        col_back, col_title = st.columns([1, 4])
        with col_back:
            if st.button("← Menu"): go("menu")
        with col_title:
            st.markdown("## 📤 Eksport")

        st.success(f"✅ Protokół zamknięty — {len(ust_l)} ustaleń")
        karta(f'<b>{prow[2]}</b><br><span style="color:#5a7a96;font-size:13px">{prow[12]} · {prow[3]}</span>')

        tab1, tab2, tab3 = st.tabs(["📄 Word", "📕 PDF", "📧 E-mail"])

        with tab1:
            if st.button("🔧 Generuj Word", use_container_width=True):
                with st.spinner("Generowanie..."):
                    wb = gen_word(pid)
                fn = f"Protokol_{prow[12].replace(' ','_')}_{prow[3]}.docx"
                st.download_button("⬇️ Pobierz .docx", wb, fn,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True)

        with tab2:
            if st.button("🔧 Generuj PDF", use_container_width=True):
                with st.spinner("Generowanie PDF..."):
                    pb = gen_pdf(pid)
                fn = f"Protokol_{prow[12].replace(' ','_')}_{prow[3]}.pdf"
                st.download_button("⬇️ Pobierz PDF", pb, fn, "application/pdf", use_container_width=True)

        with tab3:
            adr = st.text_input("E-mail odbiorcy", value=prow[16] or "")
            tem = st.text_input("Temat", value=f"Protokół BHP {prow[2]} — {prow[12]}")
            tre = st.text_area("Treść", value=f"Dzień dobry,\n\nW załączeniu protokół BHP z {prow[3]}, {prow[12]}.\n\nZ poważaniem\n{prow[6]}", height=120)
            fmt = st.radio("Załącznik", ["PDF","Word"], horizontal=True)
            c1,c2 = st.columns(2)
            with c1:
                if st.button("📧 Wyślij", type="primary", use_container_width=True):
                    with st.spinner("Wysyłanie..."):
                        att = gen_pdf(pid) if fmt=="PDF" else gen_word(pid)
                        ext = "pdf" if fmt=="PDF" else "docx"
                        ok,info = wyslij(adr,tem,tre,att,f"Protokol_{prow[3]}.{ext}")
                    st.success("✅ Wysłano!") if ok else st.error(f"❌ {info}")
            with c2:
                ml = f"mailto:{adr}?subject={urllib.parse.quote(tem)}&body={urllib.parse.quote(tre)}"
                st.link_button("📬 Klient pocztowy", ml, use_container_width=True)

# ══════════════════════════════════════════════════════════════
# EKRAN: ARCHIWUM
# ══════════════════════════════════════════════════════════════
elif st.session_state.ekran == "archiwum":
    col_back, col_title = st.columns([1, 4])
    with col_back:
        if st.button("← Wróć"): go("menu")
    with col_title:
        st.markdown("## 📁 Archiwum")

    fl = firmy(); nazwy = [f[1] for f in fl]
    c1,c2 = st.columns(2)
    with c1: ff = st.selectbox("Firma", ["Wszystkie"]+nazwy)
    with c2: fs = st.selectbox("Status", ["Wszystkie","szkic","zamknięty"])
    fid_f = next((f[0] for f in fl if f[1]==ff), None) if ff!="Wszystkie" else None
    st_f  = None if fs=="Wszystkie" else fs
    pl = protokoly(fid_f, st_f)

    if not pl:
        st.info("Brak protokołów.")
    else:
        section(f"{len(pl)} protokołów")
        for p in pl:
            ico = "🔒" if p[5]=="zamknięty" else "✏️"
            with st.expander(f"{ico} {p[2]}  ·  {p[1]}  ·  {p[7]} ust."):
                c1,c2 = st.columns(2)
                with c1: st.write(f"**Firma:** {p[1]}"); st.write(f"**Data:** {p[3]}"); st.write(f"**Status:** {p[5]}")
                with c2: st.write(f"**Miejsce:** {p[4] or '–'}"); st.write(f"**Kontrolujący:** {p[6]}")
                ca,cb,cc = st.columns(3)
                with ca:
                    if st.button("📄 Word", key=f"ew_{p[0]}", use_container_width=True):
                        w = gen_word(p[0])
                        st.download_button("⬇️ .docx", w,
                            f"Protokol_{p[1].replace(' ','_')}_{p[3]}.docx",
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"dw_{p[0]}")
                with cb:
                    if st.button("📕 PDF", key=f"ep_{p[0]}", use_container_width=True):
                        pb = gen_pdf(p[0])
                        st.download_button("⬇️ PDF", pb,
                            f"Protokol_{p[1].replace(' ','_')}_{p[3]}.pdf",
                            "application/pdf", key=f"dp_{p[0]}")
                with cc:
                    if st.button("✏️ Edytuj", key=f"ed_{p[0]}", use_container_width=True):
                        st.session_state.protokol_id = p[0]
                        st.session_state.ekran = "nowy"
                        st.session_state.podekran = "ustalenia"
                        st.rerun()

# ══════════════════════════════════════════════════════════════
# EKRAN: REKONTROLA
# ══════════════════════════════════════════════════════════════
elif st.session_state.ekran == "rekontrola":
    col_back, col_title = st.columns([1, 4])
    with col_back:
        if st.button("← Wróć"): go("menu")
    with col_title:
        st.markdown("## 🔄 Rekontrola")

    ot = otwarte()
    if not ot:
        st.success("✅ Wszystkie zalecenia zamknięte!")
    else:
        for pri_gr in ["Krytyczny","Wysoki","Średni","Niski"]:
            gr = [z for z in ot if z[5]==pri_gr]
            if not gr: continue
            section(f"{pri_gr} — {len(gr)} zaleceń")
            for z in gr:
                prze = False
                try:
                    prze = datetime.date.fromisoformat(z[7]) < datetime.date.today()
                except: pass
                with st.expander(f"{'🚨' if prze else '⚠️'} {z[1]}  ·  {z[4] or '–'}  {badge(z[5])}", expanded=(pri_gr=="Krytyczny")):
                    c1,c2 = st.columns(2)
                    with c1:
                        st.write(f"**Zagrożenie:** {z[8] or '–'}")
                        st.write(f"**Zalecenie:** {z[9]}")
                        st.write(f"**Protokół:** {z[2]} z {z[3]}")
                    with c2:
                        st.write(f"**Odp.:** {z[6] or '–'}")
                        if prze: st.error(f"Termin: {z[7]} ⚠ MINĄŁ")
                        else: st.write(f"**Termin:** {z[7] or '–'}")
                    ns = st.selectbox("Status",["Nowe","W trakcie","Zamknięte"],
                                       index=["Nowe","W trakcie","Zamknięte"].index(z[10])
                                       if z[10] in ["Nowe","W trakcie","Zamknięte"] else 0,
                                       key=f"rs_{z[0]}")
                    if ns != z[10]:
                        if st.button(f"💾 Zapisz → {ns}", key=f"rsv_{z[0]}", use_container_width=True):
                            qw("UPDATE ustalenia SET status=? WHERE id=?",(ns,z[0]))
                            st.success("✅ Zaktualizowano"); st.rerun()
