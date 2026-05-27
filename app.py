"""
Mobilny Asystent Kontroli BHP — CBZ v2.0
Michał Młynarczak / Centrum Bezpiecznego Zatrudnienia

Stack: Streamlit + SQLite + Anthropic Claude + docxtpl + reportlab
"""

import streamlit as st
import sqlite3
import datetime
import json
import os
import io
import base64
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from pathlib import Path

from openai import OpenAI
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                 Table, TableStyle, Image as RLImage)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm

# ─────────────────────────────────────────────
# KONFIGURACJA
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="BHP CBZ",
    page_icon="🦺",
    layout="wide",
    initial_sidebar_state="expanded",
)

DB_PATH = Path("bhp_cbz.db")
SZABLON_DOCX = Path("szablon.docx")

# Wczytaj klucz API
try:
    api_key = st.secrets["OPENAI_API_KEY"]
except Exception:
    st.error("🚨 Brak klucza API. Dodaj OPENAI_API_KEY do .streamlit/secrets.toml")
    st.stop()

AI_CLIENT = OpenAI(api_key=api_key)

# ─────────────────────────────────────────────
# BAZA DANYCH — inicjalizacja
# ─────────────────────────────────────────────
def init_db():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.executescript("""
        CREATE TABLE IF NOT EXISTS firmy (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            nazwa       TEXT NOT NULL UNIQUE,
            nip         TEXT,
            adres       TEXT,
            kontakt     TEXT,
            email       TEXT,
            uwagi       TEXT,
            created_at  TEXT DEFAULT (datetime('now','localtime'))
        );

        CREATE TABLE IF NOT EXISTS protokoly (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            firma_id        INTEGER REFERENCES firmy(id),
            nr_protokolu    TEXT,
            data_kontroli   TEXT,
            miejsce         TEXT,
            obszar          TEXT,
            kontrolujacy    TEXT DEFAULT 'Michał Młynarczak',
            osoby_obecne    TEXT,
            rodzaj_kontroli TEXT DEFAULT 'Bieżąca kontrola warunków pracy',
            ograniczenia    TEXT DEFAULT 'Brak',
            status          TEXT DEFAULT 'szkic',  -- szkic / zamknięty
            created_at      TEXT DEFAULT (datetime('now','localtime'))
        );

        CREATE TABLE IF NOT EXISTS ustalenia (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            protokol_id     INTEGER REFERENCES protokoly(id),
            obszar          TEXT,
            kategoria       TEXT,
            priorytet       TEXT,
            status          TEXT DEFAULT 'Nowe',
            odpowiedzialny  TEXT,
            termin          TEXT,
            ryzyko          TEXT,
            opis_stanu      TEXT,
            zalecenie       TEXT,
            foto_bytes      BLOB,
            created_at      TEXT DEFAULT (datetime('now','localtime'))
        );
    """)
    con.commit()
    con.close()

init_db()

# ─────────────────────────────────────────────
# POMOCNICZE FUNKCJE DB
# ─────────────────────────────────────────────
def db():
    return sqlite3.connect(DB_PATH)

def pobierz_firmy():
    with db() as con:
        rows = con.execute("SELECT id, nazwa, nip, adres, kontakt, email FROM firmy ORDER BY nazwa").fetchall()
    return rows

def zapisz_firme(nazwa, nip, adres, kontakt, email, uwagi=""):
    with db() as con:
        con.execute("""
            INSERT INTO firmy(nazwa,nip,adres,kontakt,email,uwagi)
            VALUES(?,?,?,?,?,?)
            ON CONFLICT(nazwa) DO UPDATE SET
              nip=excluded.nip, adres=excluded.adres,
              kontakt=excluded.kontakt, email=excluded.email, uwagi=excluded.uwagi
        """, (nazwa, nip, adres, kontakt, email, uwagi))
        con.commit()

def nowy_protokol(firma_id, nr, data, miejsce, obszar, kontrolujacy, osoby, rodzaj, ograniczenia):
    with db() as con:
        cur = con.execute("""
            INSERT INTO protokoly(firma_id,nr_protokolu,data_kontroli,miejsce,obszar,
                                   kontrolujacy,osoby_obecne,rodzaj_kontroli,ograniczenia)
            VALUES(?,?,?,?,?,?,?,?,?)
        """, (firma_id, nr, data, miejsce, obszar, kontrolujacy, osoby, rodzaj, ograniczenia))
        con.commit()
        return cur.lastrowid

def pobierz_protokoly(firma_id=None, status=None):
    sql = """
        SELECT p.id, f.nazwa, p.nr_protokolu, p.data_kontroli,
               p.miejsce, p.status, p.kontrolujacy,
               (SELECT COUNT(*) FROM ustalenia u WHERE u.protokol_id=p.id) as liczba_ust
        FROM protokoly p
        JOIN firmy f ON f.id=p.firma_id
        WHERE 1=1
    """
    params = []
    if firma_id:
        sql += " AND p.firma_id=?"; params.append(firma_id)
    if status:
        sql += " AND p.status=?"; params.append(status)
    sql += " ORDER BY p.data_kontroli DESC"
    with db() as con:
        return con.execute(sql, params).fetchall()

def pobierz_protokol(prot_id):
    with db() as con:
        row = con.execute("""
            SELECT p.*, f.nazwa, f.nip, f.adres, f.kontakt, f.email
            FROM protokoly p JOIN firmy f ON f.id=p.firma_id
            WHERE p.id=?
        """, (prot_id,)).fetchone()
    return row

def zapisz_ustalenie(prot_id, obszar, kat, prior, status, odp, termin,
                     ryzyko, opis, zalecenie, foto_bytes=None):
    with db() as con:
        cur = con.execute("""
            INSERT INTO ustalenia(protokol_id,obszar,kategoria,priorytet,status,
                                   odpowiedzialny,termin,ryzyko,opis_stanu,zalecenie,foto_bytes)
            VALUES(?,?,?,?,?,?,?,?,?,?,?)
        """, (prot_id, obszar, kat, prior, status, odp, str(termin),
              ryzyko, opis, zalecenie, foto_bytes))
        con.commit()
        return cur.lastrowid

def pobierz_ustalenia(prot_id):
    with db() as con:
        return con.execute(
            "SELECT * FROM ustalenia WHERE protokol_id=? ORDER BY id",
            (prot_id,)
        ).fetchall()

def pobierz_otwarte_zalecenia():
    """Dla modułu rekontroli — wszystkie niezamknięte"""
    with db() as con:
        return con.execute("""
            SELECT u.id, f.nazwa, p.nr_protokolu, p.data_kontroli,
                   u.obszar, u.priorytet, u.odpowiedzialny, u.termin,
                   u.ryzyko, u.zalecenie, u.status
            FROM ustalenia u
            JOIN protokoly p ON p.id=u.protokol_id
            JOIN firmy f ON f.id=p.firma_id
            WHERE u.status != 'Zamknięte'
            ORDER BY u.priorytet DESC, u.termin ASC
        """).fetchall()

def aktualizuj_status_ustalenia(ust_id, nowy_status):
    with db() as con:
        con.execute("UPDATE ustalenia SET status=? WHERE id=?", (nowy_status, ust_id))
        con.commit()

def zamknij_protokol(prot_id):
    with db() as con:
        con.execute("UPDATE protokoly SET status='zamknięty' WHERE id=?", (prot_id,))
        con.commit()

# ─────────────────────────────────────────────
# AI — ANALIZA ZDJĘCIA (Claude Vision)
# ─────────────────────────────────────────────
def analizuj_zdjecie_claude(foto_bytes: bytes) -> dict:
    b64 = base64.standard_b64encode(foto_bytes).decode()
    response = AI_CLIENT.chat.completions.create(
        model="gpt-4o",
        response_format={"type": "json_object"},
        max_tokens=400,
        messages=[
            {
                "role": "system",
                "content": """Jesteś głównym specjalistą BHP z 20-letnim doświadczeniem w przemyśle.
Analizujesz zdjęcia z inspekcji BHP/PPOŻ i identyfikujesz niezgodności.
Odpowiedz WYŁĄCZNIE czystym JSON:
{
  "zagrozenie": "krótka nazwa problemu (max 8 słów)",
  "opis_niezgodnosci": "techniczny opis stanu faktycznego (2-4 zdania)",
  "zalecenie": "konkretne zalecenie pokontrolne (2-3 zdania)"
}
Jeśli nie widać niezgodności, zaproponuj zalecenie profilaktyczne."""
            },
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": "Przeanalizuj to zdjęcie z kontroli BHP."},
                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}}
                ]
            }
        ],
        temperature=0.2,
    )
    raw = response.choices[0].message.content.strip()
    raw = raw.replace("```json", "").replace("```", "").strip()
    return json.loads(raw)

# ─────────────────────────────────────────────
# GENEROWANIE WORD
# ─────────────────────────────────────────────
def generuj_word(prot_id) -> bytes:
    prot = pobierz_protokol(prot_id)
    ustalenia = pobierz_ustalenia(prot_id)

    # prot: id,firma_id,nr,data,miejsce,obszar,kontrolujacy,osoby,rodzaj,ogr,status,created,nazwa_f,nip,adres,kontakt,email
    doc = DocxTemplate(str(SZABLON_DOCX))

    sformatowane = []
    for u in ustalenia:
        # u: id,prot_id,obszar,kat,prior,status,odp,termin,ryzyko,opis,zalecenie,foto,created
        kopia = {
            "obszar": u[2], "kategoria": u[3], "priorytet": u[4],
            "status": u[5], "odpowiedzialny": u[6], "termin": u[7],
            "ryzyko": u[8], "opis_stanu": u[9], "zalecenie": u[10],
            "dzialanie_pilne": "Tak" if u[4] in ("Wysoki","Krytyczny") else "Nie",
            "wymaga_weryfikacji": "Tak",
            "podpis_zdjecia": "Fot. dołączona" if u[11] else "Brak zdjęcia",
        }
        if u[11]:
            stream = io.BytesIO(u[11])
            kopia["zdjecie"] = InlineImage(doc, stream, width=Mm(100))
        else:
            kopia["zdjecie"] = ""
        sformatowane.append(kopia)

    wysok_kryt = sum(1 for u in ustalenia if u[4] in ("Wysoki","Krytyczny"))

    ctx = {
        "nr_protokolu": prot[2],
        "data_kontroli": prot[3],
        "klient_nazwa": prot[12],
        "klient_nip": prot[13] or "",
        "klient_adres": prot[14] or "",
        "miejsce_kontroli": prot[4] or "",
        "obszar_kontroli": prot[5] or "",
        "rodzaj_kontroli": prot[8],
        "data_i_godzina_kontroli": prot[3],
        "kontrolujacy": prot[6],
        "osoby_uczestniczace": prot[7] or "",
        "osoby_kontrolowane": "-",
        "dokumenty_odniesienia": "Ustalenia pokontrolne",
        "ograniczenia_kontroli": prot[9] or "Brak",
        "adresaci_email": prot[16] or "-",
        "zakres_kontroli": "BHP i PPOŻ",
        "obszary_sprawdzenia": prot[5] or "Zgodnie z kartami ustaleń",
        "liczba_ustalen": str(len(sformatowane)),
        "liczba_wysokich_krytycznych": str(wysok_kryt),
        "obszary_wymagajace_dzialan": "-",
        "termin_przegladu_zalecen": "-",
        "wnioski_koncowe": "Szczegóły zamieszczono w tabeli rejestru zaleceń.",
        "przedstawiciel_zakladu": prot[7] or "-",
        "koordynator_realizacji": "-",
        "zatwierdzajacy": prot[6],
        "data_podpisu": prot[3],
        "punkty": sformatowane,
    }

    doc.render(ctx)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# ─────────────────────────────────────────────
# GENEROWANIE PDF
# ─────────────────────────────────────────────
KOLOR_CIEMNY = colors.HexColor("#0f2d4e")
KOLOR_AKCENT = colors.HexColor("#c85a1e")
KOLOR_JASNY  = colors.HexColor("#f4f7fb")
KOLOR_LINIA  = colors.HexColor("#dde4ee")

KOLOR_WYSOKI    = colors.HexColor("#fff3e0")
KOLOR_KRYTYCZNY = colors.HexColor("#ffebee")
KOLOR_SREDNI    = colors.HexColor("#fffde7")
KOLOR_NISKI     = colors.HexColor("#e8f5e9")

def kolor_priorytetu(p):
    return {"Krytyczny": KOLOR_KRYTYCZNY, "Wysoki": KOLOR_WYSOKI,
            "Średni": KOLOR_SREDNI, "Niski": KOLOR_NISKI}.get(p, colors.white)

def generuj_pdf(prot_id) -> bytes:
    prot = pobierz_protokol(prot_id)
    ustalenia = pobierz_ustalenia(prot_id)
    bio = io.BytesIO()

    doc = SimpleDocTemplate(bio, pagesize=A4,
                             leftMargin=20*mm, rightMargin=20*mm,
                             topMargin=18*mm, bottomMargin=18*mm)
    styles = getSampleStyleSheet()

    # Style
    def st_h1():
        return ParagraphStyle("h1", parent=styles["Normal"],
                               fontName="Helvetica-Bold", fontSize=13,
                               textColor=colors.white, spaceAfter=4)
    def st_h2():
        return ParagraphStyle("h2", parent=styles["Normal"],
                               fontName="Helvetica-Bold", fontSize=10,
                               textColor=KOLOR_CIEMNY, spaceAfter=3)
    def st_body():
        return ParagraphStyle("body", parent=styles["Normal"],
                               fontName="Helvetica", fontSize=9,
                               leading=13, spaceAfter=2)
    def st_small():
        return ParagraphStyle("small", parent=styles["Normal"],
                               fontName="Helvetica", fontSize=8,
                               textColor=colors.HexColor("#555"), leading=11)
    def st_cap():
        return ParagraphStyle("cap", parent=styles["Normal"],
                               fontName="Helvetica-Bold", fontSize=10,
                               textColor=colors.white)

    W = doc.width

    def naglowek_tabela(tekst):
        t = Table([[Paragraph(tekst, st_cap())]], colWidths=[W])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), KOLOR_CIEMNY),
            ("TOPPADDING",  (0,0), (-1,-1), 8),
            ("BOTTOMPADDING",(0,0),(-1,-1), 8),
            ("LEFTPADDING", (0,0), (-1,-1), 10),
        ]))
        return t

    def wiersz_danych(etykieta, wartosc, etykieta2=None, wartosc2=None):
        if etykieta2:
            row = [[Paragraph(f"<b>{etykieta}</b>", st_small()),
                    Paragraph(str(wartosc or ""), st_body()),
                    Paragraph(f"<b>{etykieta2}</b>", st_small()),
                    Paragraph(str(wartosc2 or ""), st_body())]]
            t = Table(row, colWidths=[W*0.22, W*0.28, W*0.22, W*0.28])
        else:
            row = [[Paragraph(f"<b>{etykieta}</b>", st_small()),
                    Paragraph(str(wartosc or ""), st_body())]]
            t = Table(row, colWidths=[W*0.25, W*0.75])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), KOLOR_JASNY),
            ("GRID",       (0,0), (-1,-1), 0.3, KOLOR_LINIA),
            ("TOPPADDING", (0,0), (-1,-1), 4),
            ("BOTTOMPADDING",(0,0),(-1,-1), 4),
            ("LEFTPADDING", (0,0), (-1,-1), 6),
        ]))
        return t

    elems = []

    # ── Nagłówek ──
    nagl = Table([[
        Paragraph("PROTOKÓŁ BHP / PPOŻ", ParagraphStyle("pt", fontName="Helvetica-Bold",
                   fontSize=16, textColor=colors.white)),
        Paragraph(f"Nr: {prot[2]}<br/>Data: {prot[3]}",
                   ParagraphStyle("nr", fontName="Helvetica", fontSize=9,
                                   textColor=colors.HexColor("#a8c4e0"), alignment=2)),
    ]], colWidths=[W*0.65, W*0.35])
    nagl.setStyle(TableStyle([
        ("BACKGROUND", (0,0),(-1,-1), KOLOR_CIEMNY),
        ("TOPPADDING", (0,0),(-1,-1), 14),
        ("BOTTOMPADDING",(0,0),(-1,-1), 14),
        ("LEFTPADDING", (0,0),(0,0), 14),
        ("RIGHTPADDING",(-1,-1),(-1,-1), 14),
        ("VALIGN", (0,0),(-1,-1), "MIDDLE"),
    ]))
    elems += [nagl, Spacer(1, 6*mm)]

    # ── Dane kontroli ──
    elems.append(naglowek_tabela("1. Dane kontroli"))
    elems.append(wiersz_danych("Klient / zakład", prot[12], "NIP", prot[13]))
    elems.append(wiersz_danych("Adres", prot[14], "Miejsce", prot[4]))
    elems.append(wiersz_danych("Kontrolujący", prot[6], "Osoby obecne", prot[7]))
    elems.append(wiersz_danych("Rodzaj kontroli", prot[8], "Ograniczenia", prot[9] or "Brak"))
    elems.append(Spacer(1, 5*mm))

    # ── Ustalenia ──
    elems.append(naglowek_tabela("2. Ustalenia i zalecenia"))
    elems.append(Spacer(1, 2*mm))

    for i, u in enumerate(ustalenia, 1):
        # u: id,prot_id,obszar,kat,prior,status,odp,termin,ryzyko,opis,zalecenie,foto,created
        bg = kolor_priorytetu(u[4])

        karta = []
        # Wiersz nagłówka karty
        hdr = Table([[
            Paragraph(f"Ustalenie {i}  ·  {u[3]}  ·  {u[2]}",
                       ParagraphStyle("uh", fontName="Helvetica-Bold", fontSize=9,
                                       textColor=KOLOR_CIEMNY)),
            Paragraph(f"<b>PRIORYTET: {u[4].upper()}</b>",
                       ParagraphStyle("up", fontName="Helvetica-Bold", fontSize=9,
                                       textColor=KOLOR_AKCENT, alignment=2)),
        ]], colWidths=[W*0.6, W*0.4])
        hdr.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1), bg),
            ("TOPPADDING",(0,0),(-1,-1), 5),
            ("BOTTOMPADDING",(0,0),(-1,-1), 5),
            ("LEFTPADDING",(0,0),(0,0), 8),
            ("RIGHTPADDING",(-1,-1),(-1,-1), 8),
            ("GRID",(0,0),(-1,-1), 0.3, KOLOR_LINIA),
        ]))
        karta.append(hdr)

        # Zdjęcie + treść
        has_foto = bool(u[11])
        if has_foto:
            foto_stream = io.BytesIO(u[11])
            foto_img = RLImage(foto_stream, width=50*mm, height=38*mm)
            tresc_col = [
                Paragraph(f"<b>Zagrożenie:</b> {u[8] or ''}", st_body()),
                Spacer(1,2*mm),
                Paragraph(f"<b>Stan faktyczny:</b>", st_small()),
                Paragraph(u[9] or "", st_body()),
                Spacer(1,2*mm),
                Paragraph(f"<b>Zalecenie:</b>", st_small()),
                Paragraph(u[10] or "", st_body()),
            ]
            row_content = [[foto_img, tresc_col]]
            t_content = Table(row_content, colWidths=[53*mm, W-53*mm])
        else:
            tresc_col = [
                Paragraph(f"<b>Zagrożenie:</b> {u[8] or ''}", st_body()),
                Spacer(1,2*mm),
                Paragraph(f"<b>Stan faktyczny:</b>", st_small()),
                Paragraph(u[9] or "", st_body()),
                Spacer(1,2*mm),
                Paragraph(f"<b>Zalecenie:</b>", st_small()),
                Paragraph(u[10] or "", st_body()),
            ]
            row_content = [[tresc_col]]
            t_content = Table(row_content, colWidths=[W])

        t_content.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1), colors.white),
            ("GRID",(0,0),(-1,-1), 0.3, KOLOR_LINIA),
            ("TOPPADDING",(0,0),(-1,-1), 6),
            ("BOTTOMPADDING",(0,0),(-1,-1), 6),
            ("LEFTPADDING",(0,0),(-1,-1), 8),
            ("VALIGN",(0,0),(-1,-1), "TOP"),
        ]))
        karta.append(t_content)

        # Wiersz: odp / termin / status
        meta = Table([[
            Paragraph(f"Odpowiedzialny: {u[6] or '–'}", st_small()),
            Paragraph(f"Termin: {u[7] or '–'}", st_small()),
            Paragraph(f"Status: {u[5]}", st_small()),
        ]], colWidths=[W*0.4, W*0.3, W*0.3])
        meta.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1), KOLOR_JASNY),
            ("GRID",(0,0),(-1,-1), 0.3, KOLOR_LINIA),
            ("TOPPADDING",(0,0),(-1,-1), 4),
            ("BOTTOMPADDING",(0,0),(-1,-1), 4),
            ("LEFTPADDING",(0,0),(-1,-1), 8),
        ]))
        karta.append(meta)
        elems += karta + [Spacer(1, 4*mm)]

    # ── Tabela zbiorcza ──
    elems.append(Spacer(1, 4*mm))
    elems.append(naglowek_tabela("3. Rejestr zaleceń — zestawienie zbiorcze"))
    elems.append(Spacer(1, 2*mm))

    tab_data = [["Lp.", "Obszar", "Niezgodność", "Zalecenie", "Odp.", "Termin", "Prior.", "Status"]]
    for i, u in enumerate(ustalenia, 1):
        tab_data.append([
            str(i), u[2][:20] or "", (u[9] or "")[:40],
            (u[10] or "")[:40], u[6][:15] or "", u[7] or "",
            u[4], u[5],
        ])

    tab = Table(tab_data, colWidths=[
        8*mm, 28*mm, 50*mm, 50*mm, 25*mm, 20*mm, 18*mm, 17*mm
    ])
    tab.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0), KOLOR_CIEMNY),
        ("TEXTCOLOR",(0,0),(-1,0), colors.white),
        ("FONTNAME",(0,0),(-1,0), "Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1), 7.5),
        ("GRID",(0,0),(-1,-1), 0.3, KOLOR_LINIA),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, KOLOR_JASNY]),
        ("TOPPADDING",(0,0),(-1,-1), 3),
        ("BOTTOMPADDING",(0,0),(-1,-1), 3),
        ("LEFTPADDING",(0,0),(-1,-1), 4),
        ("VALIGN",(0,0),(-1,-1),"TOP"),
        ("WORDWRAP",(0,0),(-1,-1), True),
    ]))
    elems.append(tab)
    elems.append(Spacer(1, 8*mm))

    # ── Podpisy ──
    elems.append(naglowek_tabela("4. Potwierdzenie"))
    podp = Table([[
        Paragraph("Sporządził", st_small()),
        Paragraph("Przedstawiciel zakładu", st_small()),
        Paragraph("Zatwierdził", st_small()),
    ],[
        Paragraph(f"<b>{prot[6]}</b><br/>.........................", st_body()),
        Paragraph(f"<b>{prot[7] or '...'}</b><br/>.........................", st_body()),
        Paragraph(f"<b>{prot[6]}</b><br/>.........................", st_body()),
    ],[
        Paragraph(f"Data: {prot[3]}", st_small()),
        Paragraph(f"Data: {prot[3]}", st_small()),
        Paragraph(f"Data: {prot[3]}", st_small()),
    ]], colWidths=[W/3, W/3, W/3])
    podp.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1), 0.3, KOLOR_LINIA),
        ("TOPPADDING",(0,0),(-1,-1), 5),
        ("BOTTOMPADDING",(0,0),(-1,-1), 5),
        ("LEFTPADDING",(0,0),(-1,-1), 8),
        ("ROWBACKGROUNDS",(0,0),(-1,-1),[KOLOR_JASNY, colors.white, KOLOR_JASNY]),
    ]))
    elems.append(podp)

    doc.build(elems)
    return bio.getvalue()

# ─────────────────────────────────────────────
# WYSYŁKA E-MAIL
# ─────────────────────────────────────────────
def wyslij_email(do, temat, tresc, zalacznik_bytes, nazwa_pliku):
    try:
        smtp_host = st.secrets.get("SMTP_HOST", "")
        smtp_port = int(st.secrets.get("SMTP_PORT", 587))
        smtp_user = st.secrets.get("SMTP_USER", "")
        smtp_pass = st.secrets.get("SMTP_PASS", "")
    except Exception:
        return False, "Brak konfiguracji SMTP w secrets.toml"

    if not smtp_host or not smtp_user:
        return False, "Uzupełnij SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS w secrets.toml"

    msg = MIMEMultipart()
    msg["From"] = smtp_user
    msg["To"] = do
    msg["Subject"] = temat
    msg.attach(MIMEText(tresc, "plain", "utf-8"))

    part = MIMEBase("application", "octet-stream")
    part.set_payload(zalacznik_bytes)
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{nazwa_pliku}"')
    msg.attach(part)

    try:
        with smtplib.SMTP(smtp_host, smtp_port) as s:
            s.starttls()
            s.login(smtp_user, smtp_pass)
            s.send_message(msg)
        return True, "Wysłano"
    except Exception as e:
        return False, str(e)

# ─────────────────────────────────────────────
# CSS — MOBILNY STYL
# ─────────────────────────────────────────────
st.markdown("""
<style>
/* Ogólne */
[data-testid="stAppViewContainer"] { background: #f4f7fb; }
[data-testid="stSidebar"] { background: #0f2d4e !important; }
[data-testid="stSidebar"] * { color: #c8d8ea !important; }
[data-testid="stSidebar"] .stButton > button {
    width: 100%; background: rgba(255,255,255,.07);
    border: 1px solid rgba(255,255,255,.15); border-radius: 10px;
    color: #fff !important; font-weight: 600; margin-bottom: 4px;
    text-align: left; padding: 10px 14px; font-size: 15px;
}
[data-testid="stSidebar"] .stButton > button:hover {
    background: rgba(200,90,30,.25); border-color: rgba(200,90,30,.5);
}
/* Karty */
.cbz-card {
    background: #fff; border-radius: 14px;
    border: 1px solid #dde4ee; padding: 18px 20px; margin-bottom: 14px;
}
/* Priorytet badge */
.badge-krytyczny { background:#ffebee; color:#c62828; font-weight:700;
    border-radius:6px; padding:3px 10px; font-size:12px; }
.badge-wysoki { background:#fff3e0; color:#e65100; font-weight:700;
    border-radius:6px; padding:3px 10px; font-size:12px; }
.badge-sredni { background:#fffde7; color:#f57f17; font-weight:700;
    border-radius:6px; padding:3px 10px; font-size:12px; }
.badge-niski { background:#e8f5e9; color:#2e7d32; font-weight:700;
    border-radius:6px; padding:3px 10px; font-size:12px; }
/* Kamera — większa */
[data-testid="stCameraInput"] video,
[data-testid="stCameraInput"] img {
    width: 100% !important; max-height: 65vh !important; object-fit: cover;
    border-radius: 12px;
}
/* Przyciski główne */
.stButton > button[kind="primary"] {
    background: #c85a1e !important; border: none !important;
    border-radius: 12px !important; font-size: 16px !important;
    font-weight: 700 !important; padding: 14px !important;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# NAWIGACJA — SIDEBAR
# ─────────────────────────────────────────────
if "ekran" not in st.session_state:
    st.session_state.ekran = "menu"
if "aktywny_protokol" not in st.session_state:
    st.session_state.aktywny_protokol = None
if "aktywna_firma" not in st.session_state:
    st.session_state.aktywna_firma = None

with st.sidebar:
    st.markdown("## 🦺 CBZ Inspector")
    st.markdown("---")

    if st.button("🏠  Menu główne"):
        st.session_state.ekran = "menu"
        st.rerun()
    if st.button("🏢  Firmy / klienci"):
        st.session_state.ekran = "firmy"
        st.rerun()
    if st.button("📋  Nowy protokół"):
        st.session_state.ekran = "nowy_protokol"
        st.rerun()
    if st.button("📁  Archiwum protokołów"):
        st.session_state.ekran = "archiwum"
        st.rerun()
    if st.button("🔄  Rekontrola"):
        st.session_state.ekran = "rekontrola"
        st.rerun()

    st.markdown("---")
    firmy = pobierz_firmy()
    if firmy:
        st.markdown("**Ostatnie firmy:**")
        for fid, fnazwa, *_ in firmy[:5]:
            if st.button(f"  {fnazwa[:22]}", key=f"sf_{fid}"):
                st.session_state.aktywna_firma = fid
                st.session_state.ekran = "nowy_protokol"
                st.rerun()

# ─────────────────────────────────────────────
# EKRAN: MENU GŁÓWNE
# ─────────────────────────────────────────────
if st.session_state.ekran == "menu":
    st.title("🦺 CBZ Inspector")
    st.markdown("### Mobilny Asystent Kontroli BHP")
    st.markdown("---")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="cbz-card">', unsafe_allow_html=True)
        st.markdown("### 📋 Nowa kontrola")
        st.markdown("Stwórz nowy protokół z analizą AI zdjęć")
        if st.button("Zacznij kontrolę →", use_container_width=True):
            st.session_state.ekran = "nowy_protokol"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="cbz-card">', unsafe_allow_html=True)
        st.markdown("### 🔄 Rekontrola")
        otwarte = pobierz_otwarte_zalecenia()
        st.markdown(f"Otwartych zaleceń: **{len(otwarte)}**")
        if st.button("Przejdź do rekontroli →", use_container_width=True):
            st.session_state.ekran = "rekontrola"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="cbz-card">', unsafe_allow_html=True)
        st.markdown("### 📁 Archiwum")
        protokoly = pobierz_protokoly()
        st.markdown(f"Protokołów w bazie: **{len(protokoly)}**")
        if st.button("Przeglądaj archiwum →", use_container_width=True):
            st.session_state.ekran = "archiwum"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="cbz-card">', unsafe_allow_html=True)
        st.markdown("### 🏢 Baza firm")
        firmy = pobierz_firmy()
        st.markdown(f"Firm w bazie: **{len(firmy)}**")
        if st.button("Zarządzaj firmami →", use_container_width=True):
            st.session_state.ekran = "firmy"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# EKRAN: FIRMY
# ─────────────────────────────────────────────
elif st.session_state.ekran == "firmy":
    st.title("🏢 Baza klientów")

    with st.expander("➕ Dodaj / edytuj firmę", expanded=not bool(pobierz_firmy())):
        with st.form("form_firma"):
            c1, c2 = st.columns(2)
            with c1:
                fnazwa  = st.text_input("Pełna nazwa firmy *")
                fnip    = st.text_input("NIP")
                femail  = st.text_input("E-mail (do wysyłki protokołów)")
            with c2:
                fadres  = st.text_area("Adres", height=80)
                fkontakt = st.text_input("Osoba kontaktowa")
                fuwagi  = st.text_area("Uwagi", height=80)

            if st.form_submit_button("💾 Zapisz firmę", use_container_width=True):
                if fnazwa:
                    zapisz_firme(fnazwa, fnip, fadres, fkontakt, femail, fuwagi)
                    st.success(f"✅ Zapisano: {fnazwa}")
                    st.rerun()
                else:
                    st.error("Podaj nazwę firmy")

    firmy = pobierz_firmy()
    if firmy:
        st.markdown(f"**{len(firmy)} firm w bazie:**")
        for fid, fnazwa, fnip, fadres, fkontakt, femail in firmy:
            with st.expander(f"🏢 {fnazwa}"):
                c1, c2 = st.columns(2)
                with c1:
                    st.write(f"**NIP:** {fnip or '–'}")
                    st.write(f"**Adres:** {fadres or '–'}")
                with c2:
                    st.write(f"**Kontakt:** {fkontakt or '–'}")
                    st.write(f"**E-mail:** {femail or '–'}")
                if st.button(f"📋 Nowy protokół dla tej firmy", key=f"np_{fid}"):
                    st.session_state.aktywna_firma = fid
                    st.session_state.ekran = "nowy_protokol"
                    st.rerun()
    else:
        st.info("Brak firm w bazie. Dodaj pierwszą firmę powyżej.")

# ─────────────────────────────────────────────
# EKRAN: NOWY PROTOKÓŁ
# ─────────────────────────────────────────────
elif st.session_state.ekran == "nowy_protokol":

    # Podekran: metryczka lub ustalenia
    if "podekran" not in st.session_state:
        st.session_state.podekran = "metryczka"

    # ── METRYCZKA ──
    if st.session_state.podekran == "metryczka":
        st.title("📋 Nowy protokół — Dane")

        firmy = pobierz_firmy()
        nazwy = [f[1] for f in firmy]
        firma_idx = 0
        if st.session_state.aktywna_firma:
            ids = [f[0] for f in firmy]
            if st.session_state.aktywna_firma in ids:
                firma_idx = ids.index(st.session_state.aktywna_firma)

        wybrana = st.selectbox("Firma / klient *", ["– wybierz lub dodaj –"] + nazwy,
                                index=firma_idx + 1 if nazwy else 0)

        if wybrana == "– wybierz lub dodaj –":
            st.info("Wybierz firmę lub przejdź do modułu Firm żeby ją dodać.")
            if st.button("➕ Dodaj nową firmę"):
                st.session_state.ekran = "firmy"
                st.rerun()
        else:
            firma_row = next(f for f in firmy if f[1] == wybrana)
            fid = firma_row[0]

            rok = datetime.date.today().year
            mies = datetime.date.today().month
            nr_prop = f"CBZ/BHP/{rok}/{mies:02d}/001"

            c1, c2 = st.columns(2)
            with c1:
                nr   = st.text_input("Nr protokołu", value=nr_prop)
                data = st.date_input("Data kontroli", datetime.date.today())
                kontr = st.text_input("Kontrolujący", value="Michał Młynarczak")
            with c2:
                miejsce = st.text_input("Miejsce / wydział", placeholder="np. Hala W-2")
                obszar  = st.text_input("Obszar / zakres", placeholder="np. Spawalnia")
                osoby   = st.text_input("Osoby obecne", value=firma_row[4] or "")

            rodz = st.text_input("Rodzaj kontroli", value="Bieżąca kontrola warunków pracy")
            ogr  = st.text_input("Ograniczenia", value="Brak")

            if st.button("✅ Utwórz protokół i przejdź do ustaleń →",
                         type="primary", use_container_width=True):
                pid = nowy_protokol(fid, nr, str(data), miejsce, obszar,
                                    kontr, osoby, rodz, ogr)
                st.session_state.aktywny_protokol = pid
                st.session_state.podekran = "ustalenia"
                st.rerun()

    # ── USTALENIA ──
    elif st.session_state.podekran == "ustalenia":
        prot_id = st.session_state.aktywny_protokol
        prot = pobierz_protokol(prot_id)
        ust_lista = pobierz_ustalenia(prot_id)

        st.title(f"📸 Ustalenia — {prot[12]}")
        st.caption(f"Protokół {prot[2]}  ·  {prot[3]}  ·  Ustaleń: {len(ust_lista)}")

        # Liczniki
        col_a, col_b, col_c = st.columns(3)
        wysok = sum(1 for u in ust_lista if u[4] in ("Wysoki","Krytyczny"))
        col_a.metric("Ustaleń", len(ust_lista))
        col_b.metric("Wysoki/Krytyczny", wysok)
        col_c.metric("Status", prot[10].upper())

        st.markdown("---")

        # Forma nowego ustalenia
        with st.container():
            st.subheader("➕ Nowe ustalenie")

            # Kamera — duża
            if "cam_key" not in st.session_state:
                st.session_state.cam_key = 0

            foto = st.camera_input(
                "📷 Zdjęcie (pełny ekran — AI przeanalizuje automatycznie)",
                key=f"cam_{st.session_state.cam_key}",
                help="Zrób zdjęcie — AI automatycznie wypełni opis i zalecenie"
            )

            if foto and f"ai_done_{st.session_state.cam_key}" not in st.session_state:
                with st.spinner("🤖 Claude analizuje zagrożenia na zdjęciu..."):
                    try:
                        ai = analizuj_zdjecie_claude(foto.getvalue())
                        st.session_state["ai_zagr"] = ai.get("zagrozenie","")
                        st.session_state["ai_opis"] = ai.get("opis_niezgodnosci","")
                        st.session_state["ai_zal"]  = ai.get("zalecenie","")
                        st.session_state[f"ai_done_{st.session_state.cam_key}"] = True
                        st.success("✅ Analiza gotowa — sprawdź i zatwierdź pola poniżej")
                    except Exception as e:
                        st.warning(f"⚠ AI niedostępne: {e}")
                        st.session_state["ai_zagr"] = ""
                        st.session_state["ai_opis"] = ""
                        st.session_state["ai_zal"]  = ""

            c1, c2 = st.columns(2)
            with c1:
                obszar_u  = st.text_input("Lokalizacja / stanowisko", placeholder="np. Tokarka nr 3")
                kategoria = st.selectbox("Kategoria",
                    ["BHP","PPOŻ","Maszyny","Drogi / 5S","Chemia","ŚOI","Inne"])
                priorytet = st.selectbox("Priorytet",
                    ["Niski","Średni","Wysoki","Krytyczny"])
            with c2:
                odp    = st.text_input("Odpowiedzialny", placeholder="np. Mistrz zmiany")
                termin = st.date_input("Termin",
                    datetime.date.today() + datetime.timedelta(days=7))
                status_u = st.selectbox("Status", ["Nowe","W trakcie","Zamknięte"])

            ryzyko = st.text_input("Zagrożenie (skrót)",
                value=st.session_state.get("ai_zagr",""))
            opis   = st.text_area("Stan faktyczny — niezgodność",
                value=st.session_state.get("ai_opis",""), height=110)
            zal    = st.text_area("Zalecenie / działanie korygujące",
                value=st.session_state.get("ai_zal",""), height=110)

            if st.button("➕ Zapisz ustalenie do protokołu",
                         type="primary", use_container_width=True):
                if opis and zal:
                    foto_bytes = foto.getvalue() if foto else None
                    zapisz_ustalenie(prot_id, obszar_u, kategoria, priorytet,
                                     status_u, odp, termin, ryzyko, opis, zal,
                                     foto_bytes)
                    # Reset stanu
                    st.session_state.cam_key += 1
                    for k in ["ai_zagr","ai_opis","ai_zal"]:
                        st.session_state.pop(k, None)
                    # Usuń ai_done dla poprzedniego klucza
                    old_key = f"ai_done_{st.session_state.cam_key - 1}"
                    st.session_state.pop(old_key, None)
                    st.success(f"✅ Ustalenie {len(ust_lista)+1} zapisane!")
                    st.rerun()
                else:
                    st.error("Uzupełnij 'Stan faktyczny' oraz 'Zalecenie'")

        # Lista dodanych ustaleń
        if ust_lista:
            st.markdown("---")
            st.subheader(f"📋 Dodane ustalenia ({len(ust_lista)})")
            for u in ust_lista:
                badge_cls = {
                    "Krytyczny":"badge-krytyczny",
                    "Wysoki":"badge-wysoki",
                    "Średni":"badge-sredni",
                    "Niski":"badge-niski"
                }.get(u[4],"badge-niski")
                with st.expander(f"**{u[2] or 'Ustalenie'}** — {u[3]}  "
                                  f'<span class="{badge_cls}">{u[4]}</span>',
                                  expanded=False):
                    st.write(f"**Zagrożenie:** {u[8]}")
                    st.write(f"**Opis:** {u[9]}")
                    st.write(f"**Zalecenie:** {u[10]}")
                    st.write(f"Odpowiedzialny: {u[6]}  ·  Termin: {u[7]}")
                    if u[11]:
                        st.image(u[11], width=280)

        st.markdown("---")
        st.subheader("🚀 Finalizacja")

        c1, c2, c3 = st.columns(3)
        with c1:
            if st.button("➕ Kolejne ustalenie", use_container_width=True):
                st.rerun()
        with c2:
            if st.button("🏁 Zamknij protokół", use_container_width=True):
                zamknij_protokol(prot_id)
                st.session_state.podekran = "eksport"
                st.rerun()
        with c3:
            if st.button("📁 Zapisz szkic i wróć", use_container_width=True):
                st.session_state.ekran = "menu"
                st.session_state.podekran = "metryczka"
                st.rerun()

    # ── EKSPORT ──
    elif st.session_state.podekran == "eksport":
        prot_id = st.session_state.aktywny_protokol
        prot = pobierz_protokol(prot_id)
        ust_lista = pobierz_ustalenia(prot_id)

        st.title("📤 Eksport protokołu")
        st.success(f"✅ Protokół {prot[2]} zamknięty — {len(ust_lista)} ustaleń")

        tab1, tab2, tab3 = st.tabs(["📄 Word (.docx)", "📕 PDF", "📧 E-mail"])

        with tab1:
            st.markdown("**Protokół w formacie Word** — edytowalny szablon CBZ")
            if st.button("🔧 Generuj Word", use_container_width=True):
                with st.spinner("Generowanie..."):
                    word_bytes = generuj_word(prot_id)
                fn = f"Protokol_{prot[12].replace(' ','_')}_{prot[3]}.docx"
                st.download_button("⬇️ Pobierz .docx", word_bytes, fn,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True)

        with tab2:
            st.markdown("**Protokół PDF** — gotowy do druku i wysyłki")
            if st.button("🔧 Generuj PDF", use_container_width=True):
                with st.spinner("Generowanie PDF..."):
                    pdf_bytes = generuj_pdf(prot_id)
                fn_pdf = f"Protokol_{prot[12].replace(' ','_')}_{prot[3]}.pdf"
                st.download_button("⬇️ Pobierz PDF", pdf_bytes, fn_pdf,
                    mime="application/pdf", use_container_width=True)

        with tab3:
            st.markdown("**Wyślij protokół e-mailem**")
            adresat = st.text_input("Adres e-mail odbiorcy",
                value=prot[16] or "")
            temat = st.text_input("Temat",
                value=f"Protokół BHP {prot[2]} — {prot[12]}")
            tresc_mail = st.text_area("Treść wiadomości",
                value=f"Dzień dobry,\n\nW załączeniu przesyłam protokół z kontroli BHP "
                      f"przeprowadzonej w dniu {prot[3]} w zakładzie {prot[12]}.\n\n"
                      f"Proszę o zapoznanie się z zaleceniami i potwierdzenie odbioru.\n\n"
                      f"Z poważaniem\n{prot[6]}\nCentrum Bezpiecznego Zatrudnienia",
                height=140)
            fmt = st.radio("Format załącznika", ["PDF", "Word (.docx)"],
                horizontal=True)

            col_a, col_b = st.columns(2)
            with col_a:
                if st.button("📧 Wyślij e-mail", type="primary",
                             use_container_width=True):
                    with st.spinner("Wysyłanie..."):
                        if fmt == "PDF":
                            att = generuj_pdf(prot_id)
                            fn_att = f"Protokol_{prot[3]}.pdf"
                        else:
                            att = generuj_word(prot_id)
                            fn_att = f"Protokol_{prot[3]}.docx"
                        ok, info = wyslij_email(adresat, temat, tresc_mail,
                                                 att, fn_att)
                    if ok:
                        st.success("✅ E-mail wysłany!")
                    else:
                        st.error(f"❌ Błąd: {info}")
                        st.info("Skonfiguruj SMTP w .streamlit/secrets.toml")
            with col_b:
                # mailto: link jako fallback
                import urllib.parse
                mailto_url = (
                    f"mailto:{adresat}"
                    f"?subject={urllib.parse.quote(temat)}"
                    f"&body={urllib.parse.quote(tresc_mail)}"
                )
                st.link_button("📬 Otwórz klienta pocztowego",
                               mailto_url, use_container_width=True)

        st.markdown("---")
        if st.button("🏠 Wróć do menu głównego", use_container_width=True):
            st.session_state.ekran = "menu"
            st.session_state.podekran = "metryczka"
            st.session_state.aktywny_protokol = None
            st.rerun()

# ─────────────────────────────────────────────
# EKRAN: ARCHIWUM
# ─────────────────────────────────────────────
elif st.session_state.ekran == "archiwum":
    st.title("📁 Archiwum protokołów")

    firmy = pobierz_firmy()
    filtr_firma = st.selectbox("Filtruj po firmie",
        ["Wszystkie"] + [f[1] for f in firmy])
    filtr_status = st.selectbox("Status",
        ["Wszystkie","szkic","zamknięty"])

    fid_filtr = None
    if filtr_firma != "Wszystkie":
        fid_filtr = next(f[0] for f in firmy if f[1] == filtr_firma)

    st_filtr = None if filtr_status == "Wszystkie" else filtr_status
    protokoly = pobierz_protokoly(fid_filtr, st_filtr)

    if not protokoly:
        st.info("Brak protokołów spełniających kryteria.")
    else:
        st.markdown(f"**{len(protokoly)} protokołów:**")
        for p in protokoly:
            # p: id, nazwa, nr, data, miejsce, status, kontrolujący, liczba_ust
            icon = "🔒" if p[5] == "zamknięty" else "✏️"
            with st.expander(
                f"{icon} {p[2]} — {p[1]}  |  {p[3]}  |  {p[7]} ustaleń",
                expanded=False
            ):
                c1, c2 = st.columns(2)
                with c1:
                    st.write(f"**Firma:** {p[1]}")
                    st.write(f"**Miejsce:** {p[4] or '–'}")
                    st.write(f"**Status:** {p[5]}")
                with c2:
                    st.write(f"**Kontrolujący:** {p[6]}")
                    st.write(f"**Ustaleń:** {p[7]}")

                c_a, c_b, c_c = st.columns(3)
                with c_a:
                    if st.button("📄 Eksport Word", key=f"ew_{p[0]}",
                                 use_container_width=True):
                        w = generuj_word(p[0])
                        fn = f"Protokol_{p[1].replace(' ','_')}_{p[3]}.docx"
                        st.download_button("⬇️ Pobierz .docx", w, fn,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"dw_{p[0]}")
                with c_b:
                    if st.button("📕 Eksport PDF", key=f"ep_{p[0]}",
                                 use_container_width=True):
                        pdf = generuj_pdf(p[0])
                        fn = f"Protokol_{p[1].replace(' ','_')}_{p[3]}.pdf"
                        st.download_button("⬇️ Pobierz PDF", pdf, fn,
                            mime="application/pdf",
                            key=f"dp_{p[0]}")
                with c_c:
                    if st.button("📋 Kontynuuj / dodaj ustalenia",
                                 key=f"ko_{p[0]}", use_container_width=True):
                        st.session_state.aktywny_protokol = p[0]
                        st.session_state.ekran = "nowy_protokol"
                        st.session_state.podekran = "ustalenia"
                        st.rerun()

# ─────────────────────────────────────────────
# EKRAN: REKONTROLA
# ─────────────────────────────────────────────
elif st.session_state.ekran == "rekontrola":
    st.title("🔄 Rekontrola — otwarte zalecenia")

    otwarte = pobierz_otwarte_zalecenia()

    if not otwarte:
        st.success("✅ Wszystkie zalecenia zamknięte!")
    else:
        # Grupuj wg priorytetu
        for priorytet_gr in ["Krytyczny","Wysoki","Średni","Niski"]:
            gr = [z for z in otwarte if z[5] == priorytet_gr]
            if not gr:
                continue

            badge_cls = {
                "Krytyczny":"badge-krytyczny","Wysoki":"badge-wysoki",
                "Średni":"badge-sredni","Niski":"badge-niski"
            }[priorytet_gr]
            st.markdown(
                f'<span class="{badge_cls}">{priorytet_gr.upper()}</span>'
                f'  — {len(gr)} zaleceń',
                unsafe_allow_html=True
            )

            for z in gr:
                # z: id,firma,nr_prot,data,obszar,priorytet,odp,termin,ryzyko,zalecenie,status
                przeterminowane = False
                try:
                    t = datetime.date.fromisoformat(z[7])
                    przeterminowane = t < datetime.date.today()
                except Exception:
                    pass

                label = f"{'🚨' if przeterminowane else '⚠️'} {z[1]} · {z[4] or 'brak obszaru'}"
                with st.expander(label, expanded=priorytet_gr == "Krytyczny"):
                    c1, c2 = st.columns([2,1])
                    with c1:
                        st.write(f"**Zagrożenie:** {z[8] or '–'}")
                        st.write(f"**Zalecenie:** {z[9]}")
                        st.write(f"**Protokół:** {z[2]} z {z[3]}")
                    with c2:
                        st.write(f"**Odpowiedzialny:** {z[6] or '–'}")
                        if przeterminowane:
                            st.error(f"Termin: {z[7]} ⚠ MINĄŁ!")
                        else:
                            st.write(f"**Termin:** {z[7] or '–'}")
                        st.write(f"**Status:** {z[10]}")

                    nowy_status = st.selectbox(
                        "Zmień status",
                        ["Nowe","W trakcie","Zamknięte"],
                        index=["Nowe","W trakcie","Zamknięte"].index(z[10])
                        if z[10] in ["Nowe","W trakcie","Zamknięte"] else 0,
                        key=f"st_{z[0]}"
                    )
                    if nowy_status != z[10]:
                        if st.button(f"💾 Zapisz status → {nowy_status}",
                                     key=f"sv_{z[0]}", use_container_width=True):
                            aktualizuj_status_ustalenia(z[0], nowy_status)
                            st.success("✅ Status zaktualizowany")
                            st.rerun()

            st.markdown("---")

