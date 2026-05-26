import streamlit as st
import datetime
import json
import os
import io
import base64
from openai import OpenAI
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

# ==========================================
# KONFIGURACJA I AUTORYZACJA AI
# ==========================================
st.set_page_config(page_title="Asystent Kontroli BHP", page_icon="🦺", layout="centered")
PLIK_BAZY = "klienci.json"

try:
    klucz_api = st.secrets["OPENAI_API_KEY"]
    client = OpenAI(api_key=klucz_api)
except Exception as e:
    st.error("🚨 Błąd: Nie znaleziono klucza API! Upewnij się, że masz plik .streamlit/secrets.toml")
    st.stop()

if "punkty" not in st.session_state:
    st.session_state.punkty = []

# ==========================================
# FUNKCJE POMOCNICZE
# ==========================================
def wczytaj_klientow():
    if os.path.exists(PLIK_BAZY):
        with open(PLIK_BAZY, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def zapisz_klienta(nazwa, nip, adres, kontakt):
    klienci = wczytaj_klientow()
    klienci[nazwa] = {"nip": nip, "adres": adres, "kontakt": kontakt}
    with open(PLIK_BAZY, "w", encoding="utf-8") as f:
        json.dump(klienci, f, ensure_ascii=False, indent=4)

def analizuj_zdjecie(foto_bytes):
    base64_image = base64.b64encode(foto_bytes).decode('utf-8')
    # Zmodyfikowany prompt uodparniający AI na "normalne" zdjęcia
    prompt_systemowy = """Jesteś głównym specjalistą ds. BHP.
    Twoim zadaniem jest analiza zdjęcia i identyfikacja niezgodności BHP/PPOŻ.
    Nawet jeśli na zdjęciu są tylko ludzie, biuro, lub nie ma ewidentnych zagrożeń, MUSISZ zachować swoją rolę i wymyślić profilaktyczne zalecenie.
    Musisz odpowiedzieć WYŁĄCZNIE czystym formatem JSON, np:
    {
      "zagrozenie": "Krótka nazwa problemu lub Brak zagrożeń",
      "opis_niezgodnosci": "Rzeczowy, techniczny opis tego co widać.",
      "zalecenie": "Krótkie zalecenie pokontrolne lub profilaktyczne."
    }"""

    # Kluczowa zmiana: response_format={"type": "json_object"}
    response = client.chat.completions.create(
        model="gpt-4o",
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": prompt_systemowy},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": "Przeanalizuj to zdjęcie i zwróć obiekt JSON."},
                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}}
                ]
            }
        ],
        max_tokens=300,
        temperature=0.2
    )
    
    wynik = response.choices[0].message.content.strip()
    return json.loads(wynik)

# ==========================================
# BOCZNE MENU - PROFILE FIRM
# ==========================================
with st.sidebar:
    st.header("🏢 Baza Klientów")
    with st.form("nowy_klient"):
        nowa_nazwa = st.text_input("Pełna nazwa firmy")
        nowy_nip = st.text_input("NIP")
        nowy_adres = st.text_area("Adres siedziby / zakładu")
        nowy_kontakt = st.text_input("Osoba kontaktowa")
        if st.form_submit_button("Zapisz profil"):
            if nowa_nazwa:
                zapisz_klienta(nowa_nazwa, nowy_nip, nowy_adres, nowy_kontakt)
                st.success(f"Dodano firmę: {nowa_nazwa}")

# ==========================================
# 1. METRYCZKA KONTROLI
# ==========================================
st.title("🦺 Mobilny Asystent Kontroli")

klienci_w_bazie = wczytaj_klientow()
lista_nazw = ["-- Wybierz klienta z bazy --"] + list(klienci_w_bazie.keys())

st.header("1. Metryczka Kontroli")
wybrany_klient = st.selectbox("Wybierz profil zakładu", lista_nazw)

domyslny_nip = klienci_w_bazie.get(wybrany_klient, {}).get("nip", "") if wybrany_klient != "-- Wybierz klienta z bazy --" else ""
domyslny_adres = klienci_w_bazie.get(wybrany_klient, {}).get("adres", "") if wybrany_klient != "-- Wybierz klienta z bazy --" else ""
domyslny_uczestnik = klienci_w_bazie.get(wybrany_klient, {}).get("kontakt", "") if wybrany_klient != "-- Wybierz klienta z bazy --" else ""

col1, col2 = st.columns(2)
with col1:
    klient_nazwa = st.text_input("Nazwa klienta", value=wybrany_klient if wybrany_klient != "-- Wybierz klienta z bazy --" else "")
    klient_nip = st.text_input("NIP", value=domyslny_nip)
with col2:
    klient_adres = st.text_input("Adres", value=domyslny_adres)
    miejsce_kontroli = st.text_input("Obszar/Wydział kontroli", placeholder="np. Hala W-2")

osoby_obecne = st.text_input("Osoby uczestniczące", value=domyslny_uczestnik)
data_kontroli = st.date_input("Data kontroli", datetime.date.today())
kontrolujacy = st.text_input("Osoba kontrolująca", value="Michał Młynarczak")

st.markdown("---")

# ==========================================
# 2. REJESTRACJA USTALEŃ (Z MODUŁEM AI)
# ==========================================
st.header("2. Karta Ustalenia (Analiza AI 🤖)")

if "klucz_aparatu" not in st.session_state:
    st.session_state.klucz_aparatu = 0

with st.container():
    foto = st.camera_input("Zrób zdjęcie (AI automatycznie je przeanalizuje)", key=f"aparat_{st.session_state.klucz_aparatu}")
    
    if foto is not None:
        if "aktualne_foto" not in st.session_state or st.session_state.aktualne_foto != foto.name:
            st.session_state.aktualne_foto = foto.name 
            
            with st.spinner("🤖 Wirtualny Inspektor analizuje zagrożenia na zdjęciu..."):
                try:
                    dane_ai = analizuj_zdjecie(foto.getvalue())
                    st.session_state.ai_zagrozenie = dane_ai.get("zagrozenie", "Inne zagrożenie")
                    st.session_state.ai_opis = dane_ai.get("opis_niezgodnosci", "")
                    st.session_state.ai_zalecenie = dane_ai.get("zalecenie", "")
                    st.success("✅ Analiza zakończona! Sprawdź i zatwierdź opisy w formularzu poniżej.")
                except Exception as e:
                    st.error(f"Wystąpił problem z połączeniem do AI. ({e})")
                    st.session_state.ai_zagrozenie = ""
                    st.session_state.ai_opis = ""
                    st.session_state.ai_zalecenie = ""

    ai_zagr = st.session_state.get("ai_zagrozenie", "")
    ai_op = st.session_state.get("ai_opis", "")
    ai_zal = st.session_state.get("ai_zalecenie", "")

    col3, col4 = st.columns(2)
    with col3:
        obszar = st.text_input("Lokalizacja usterki", placeholder="np. Tokarka nr 3")
        kategoria = st.selectbox("Kategoria", ["BHP", "PPOŻ", "Maszyny", "Drogi / 5S", "Chemia", "ŚOI", "Inne"])
        priorytet = st.selectbox("Priorytet", ["Niski", "Średni", "Wysoki", "Krytyczny"])
    with col4:
        odpowiedzialny = st.text_input("Odpowiedzialny za realizację", placeholder="np. Mistrz zmiany")
        termin = st.date_input("Termin wykonania", datetime.date.today() + datetime.timedelta(days=7))
        status = st.selectbox("Status", ["Nowe", "W trakcie", "Zamknięte"])

    ryzyko = st.text_input("Krótkie zagrożenie / problem", value=ai_zagr)
    opis_stanu = st.text_area("Opis stanu faktycznego (niezgodność)", value=ai_op, height=100)
    zalecenie = st.text_area("Zalecenie / działanie korygujące", value=ai_zal, height=100)

    if st.button("➕ Zapisz ten punkt do protokołu", use_container_width=True):
        if opis_stanu and zalecenie:
            nowy_punkt = {
                "foto_obiekt": foto,
                "obszar": obszar,
                "kategoria": kategoria,
                "priorytet": priorytet,
                "status": status,
                "odpowiedzialny": odpowiedzialny,
                "termin": str(termin),
                "dzialanie_pilne": "Tak" if priorytet in ["Wysoki", "Krytyczny"] else "Nie",
                "wymaga_weryfikacji": "Tak",
                "podpis_zdjecia": "Fot. dołączona" if foto else "Brak zdjęcia",
                "opis_stanu": opis_stanu,
                "ryzyko": ryzyko,
                "zalecenie": zalecenie
            }
            st.session_state.punkty.append(nowy_punkt)
            
            st.session_state.klucz_aparatu += 1 
            if "aktualne_foto" in st.session_state: del st.session_state.aktualne_foto
            if "ai_zagrozenie" in st.session_state: del st.session_state.ai_zagrozenie
            if "ai_opis" in st.session_state: del st.session_state.ai_opis
            if "ai_zalecenie" in st.session_state: del st.session_state.ai_zalecenie
            
            st.rerun() 
        else:
            st.error("Uzupełnij przynajmniej 'Opis stanu' oraz 'Zalecenie'!")

st.info(f"📍 Aktualna liczba ustaleń oczekujących na wygenerowanie w protokole: {len(st.session_state.punkty)}")

# ==========================================
# 3. GENEROWANIE RAPORTU WORD
# ==========================================
st.markdown("---")
st.header("3. Finał - Generowanie raportu")

if st.button("🚀 Wygeneruj Protokół Word z zebranych punktów", type="primary", use_container_width=True):
    if not st.session_state.punkty:
        st.warning("Twój protokół jest pusty! Dodaj przynajmniej jedną niezgodność na górze.")
    else:
        try:
            doc = DocxTemplate("szablon.docx")
            
            sformatowane_punkty = []
            for pkt in st.session_state.punkty:
                kopia_pkt = pkt.copy()
                if pkt["foto_obiekt"] is not None:
                    image_stream = io.BytesIO(pkt["foto_obiekt"].getvalue())
                    kopia_pkt["zdjecie"] = InlineImage(doc, image_stream, width=Mm(100))
                else:
                    kopia_pkt["zdjecie"] = ""
                sformatowane_punkty.append(kopia_pkt)
                
            kontekst = {
                "nr_protokolu": f"CBZ/BHP/{datetime.date.today().year}/{datetime.date.today().month:02d}",
                "data_kontroli": str(data_kontroli),
                "klient_nazwa": klient_nazwa,
                "klient_nip": klient_nip,
                "klient_adres": klient_adres,
                "miejsce_kontroli": miejsce_kontroli,
                "obszar_kontroli": miejsce_kontroli,
                "rodzaj_kontroli": "Bieżąca kontrola warunków pracy",
                "data_i_godzina_kontroli": str(data_kontroli),
                "kontrolujacy": kontrolujacy,
                "osoby_uczestniczace": osoby_obecne,
                "osoby_kontrolowane": "-",
                "dokumenty_odniesienia": "Ustalenia pokontrolne",
                "ograniczenia_kontroli": "Brak",
                "adresaci_email": "-",
                "zakres_kontroli": "BHP i PPOŻ",
                "obszary_sprawdzenia": "Zgodnie z kartami ustaleń",
                "liczba_ustalen": str(len(sformatowane_punkty)),
                "liczba_wysokich_krytycznych": str(sum(1 for p in sformatowane_punkty if p["priorytet"] in ["Wysoki", "Krytyczny"])),
                "obszary_wymagajace_dzialan": "-",
                "termin_przegladu_zalecen": "-",
                "wnioski_koncowe": "Szczegóły zamieszczono w tabeli rejestru zaleceń.",
                "przedstawiciel_zakladu": osoby_obecne,
                "koordynator_realizacji": "-",
                "zatwierdzajacy": kontrolujacy,
                "data_podpisu": str(data_kontroli),
                "punkty": sformatowane_punkty
            }
            
            doc.render(kontekst)
            
            bio = io.BytesIO()
            doc.save(bio)
            
            st.success("✅ Protokół został wygenerowany pomyślnie!")
            st.download_button(
                label="⬇️ Pobierz gotowy plik Word (.docx)",
                data=bio.getvalue(),
                file_name=f"Protokol_BHP_{klient_nazwa.replace(' ', '_')}_{data_kontroli}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"Wystąpił błąd podczas budowania pliku: {e}")