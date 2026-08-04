import os
import io
import time
import json
import re
import openpyxl
import pdfplumber
import streamlit as st
import traceback
from datetime import datetime
from google import genai
from google.genai import types
from google.genai.errors import APIError

# ==============================================================================
# SPEICHERBASIERTES (RAM) LOGGING SYSTEM
# ==============================================================================
if "canli_loglar" not in st.session_state:
    st.session_state["canli_loglar"] = []

def log_ekle(seviye, mesaj):
    zaman = datetime.now().strftime("%H:%M:%S")
    st.session_state["canli_loglar"].append(f"[{zaman}] [{seviye}] {mesaj}")

def log_ve_hata_yaz(hata_nesnesi, dosya_adi="General", islem_adi="Unbekannter Vorgang"):
    # temp_X_ Präfix aus Dateinamen entfernen
    temiz_dosya_adi = re.sub(r'^temp_\d+_', '', dosya_adi)
    hata_turu = type(hata_nesnesi).__name__
    hata_mesaji = str(hata_nesnesi)
    tam_iz = traceback.format_exc()

    log_mesaji = f"FEHLER | Datei: {temiz_dosya_adi} | Prozess: {islem_adi} | Typ: {hata_turu} | Nachricht: {hata_mesaji}"
    log_ekle("ERROR", log_mesaji)

    return {
        "zeit": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "datei": temiz_dosya_adi,
        "prozess": islem_adi,
        "fehler_typ": hata_turu,
        "fehler_nachricht": hata_mesaji,
        "traceback": tam_iz
    }

# ==============================================================================
# STREAMLIT SEITEN-EINSTELLUNGEN & PASSWORT-SCHUTZ
# ==============================================================================
st.set_page_config(page_title="LV Passiver Brandschutz Analysator", layout="wide")

UYGULAMA_SIFRESI = "Yangin2026"

if "dogrulandi" not in st.session_state:
    st.session_state["dogrulandi"] = False

# --- ANMELDUNG / PASSWORT-EINGABE ---
if not st.session_state["dogrulandi"]:
    girilen_sifre = st.text_input("🔒 Passwort", type="password", key="login_pass")
    
    if girilen_sifre:
        if girilen_sifre == UYGULAMA_SIFRESI:
            st.session_state["dogrulandi"] = True
            log_ekle("INFO", "Erfolgreich am System angemeldet.")
            st.rerun()
        else:
            st.error("❌ Falsches Passwort eingegeben.")
            
    st.stop()

# ==============================================================================
# HAUPTANWENDUNG (NACH ERFOLGREICHER ANMELDUNG)
# ==============================================================================

# --- SEITENLEISTE (LOG-PANEL) ---
with st.sidebar:
    with st.popover("⚙️ Systemprotokolle", use_container_width=True):
        st.write("#### 📜 Live-Prozessverlauf")
        if st.session_state["canli_loglar"]:
            log_metni = "\n\n".join(st.session_state["canli_loglar"])
            st.text_area(
                label="Log-Stream",
                value=log_metni,
                height=350,
                disabled=True,
                label_visibility="collapsed"
            )
        else:
            st.info("Noch keine Protokolle vorhanden.")

# --- HAUPTBILDSCHIRM TITEL UND BESCHREIBUNG ---
st.title("LV Passiver Brandschutz Analysator")
st.write("Extrahiert automatisch passive Brandschutzprodukte (FSU) aus deutschen Leistungsverzeichnissen (LV) und erzeugt einen Excel-Bericht.")

# API KEY HOLEN
SECILEN_API_KEY = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY", ""))

PROMPT = """
Sen katı kurallara sahip, hayal kurmayan ve ASLA uydurma bilgi üretmeyen bir Alman ihale dosyası (LV) denetçisisin.
GÖREVİN: Belgeyi poz poz incelemek ve SADECE gerçek FSU (Fire Stop Unit / Pasif Yangın Durdurucu / Abschottungssystem) ürünlerini seçmektir.

🎯 ADIM 1 - ÖNCE TÜM BELGEYİ İKİ KEZ TARA (ZİHNİNDE, ÇIKTIYA YAZMA):
Tur 1: Belgedeki TÜM poz numaralarını (Poz-Nr., örn: 01.01.0140) ve bunlara karşılık gelen Menge/Stück (adet) değerlerini uçtan uca çıkar ve eşleştir.
ÖNEMLİ: Bir pozun ürün açıklaması bir sayfada bitip, "Menge" (adet) bilgisi bir SONRAKİ sayfada, tablonun devamında veya sayfa altında/üstünde ayrı olarak yazılmış olabilir. Poz numarasını (tam olarak, nokta dahil) eşleştirerek doğru adedi bul; asla bir sonraki/önceki farklı pozun adedini karıştırma. Poz numarası tam eşleşmesi olmadan adet ATAMA.
KRİTİK KARIŞTIRMA UYARISI: Bir poz açıklamasında birden fazla sayı geçebilir (DIN norm numarası, ürün kodu/Artikelnummer, yangın sınıfı+onay kodu gibi "K 90-18017", Nenngröße vb.). Bunların HİÇBİRİ adet (Menge) DEĞİLDİR. Adet sadece "Menge", "Stück", "Anzahl" başlıklı sütunda veya "St", "Stk", "Stück" birimiyle birlikte yazan sayıdır. Emin değilsen adet için "-" yaz, tahmin etme.
BU BELGENİN TİPİK KALIBI: Adet genelde pozun açıklama metninin EN SONUNDA, "Zulassungs-Nr. [onay no]." ifadesinden HEMEN SONRA, bir sonraki poz başlamadan önce "XX,000 St" gibi yazılıdır (örn: "...Zulassungs-Nr. Z-41.3-368. 17,000 St" → doğru adet 17, birim St). Bu kısım sayfa sınırını aşarsa, sonraki sayfanın başındaki metni önceki pozun devamı olarak oku.
Tur 2: Sadece FSU niteliğindeki pozları seç (aşağıdaki tanıma bak) ve Tur 1'de çıkardığın doğru adet değeriyle birleştirerek JSON'a yaz.

🎯 FSU (PASİF YANGIN DURDURUCU) NEDİR - SADECE BUNLARI SEÇ (BEYAZ LİSTE):
Bir ürün ancak aşağıdaki gibi bir yapı elemanından (duvar, döşeme, tavan) geçen kablo/boru/kanal/yük taşıyıcı elemanın etrafındaki AÇIKLIĞI yangına karşı mühürleyen/kapatan bir "Abschottung" (durdurucu/tıkaç) sistemiyse FSU'dur:
- Kabelschott / Kabelabschottung (kablo geçiş durdurucu)
- Rohrschott / Rohrabschottung (boru geçiş durdurucu)
- Kombischott / Kombiabschottung (kombine geçiş durdurucu)
- Weichschott, Hartschott, Brandschott (yumuşak/sert dolgu durdurucu)
- Brandschutzklappe (yangın damperi/kelebek vanası - kanal içine monte edilen, F/K sınıfı olan)
- Brandschutzmanschette / Rohrmanschette / Brandmanschette (yanıcı boru manşonu)
- Brandschutzkissen (yangın yastığı/torbası)
- Deckenschott / Bodenschott / Wandschott (tavan/döşeme/duvar geçiş durdurucusu — dikkat: melez parçalar dahi mühürleme işlevindeyse dahil et)
- İçinde "Brandschutz" + "Schott/Klappe/Manschette/Kissen/Abschottung" birlikte geçen ürünler
Bu ürünlerin GENELDE bir yangın dayanım sınıfı (F30, F90, K30, K90, S90, R30 vb.) ve bir yapı onay/Zulassungsnummer'ı (Z-19.xx, Z-41.xx, abP, ETA vb.) olur.

🚫 KESİNLİKLE FSU DEĞİL - ASLA SEÇME:
Genel havalandırma/HVAC ekipmanı FSU DEĞİLDİR:
- Ventilatoreinsatz, Montagehalterung, Zweitraumset, Innenblende, Rohbauset, Fassadenblende
- Wetterschutzgitter, Luftdurchlass, Volumenstromregelgerät/VVS Regelgerät, Zusatzschalldämpfer
- Sadece çatı/duvar geçişi için genel kılıf/manto olup yangın mühürleme işlevi belirtilmeyen "Dachdurchführung", "Wanddurchführung", "Dachschneidemantel", "Lüftungsleitung"

🚨 UYDURMA / HALÜSİNASYON YASAĞI:
- **uygulama_alani**: Metinde ürünün binanın neresine takılacağı açıkça yazmıyorsa "-", yazmıyorsa uydurma.

📏 DİĞER KURALLAR:
- **urun_olcusu**: SADECE ürünün fiziksel boyutunu/çapını yaz (Örn: 1000x700x500 mm, DN100).
- **uretici**: "KA2-EU-DE" gibi kodlar üretici değildir. Üretici sadece gerçek firma adıdır.
- **sertifika_ref**: Sadece kısa onay numarasını yaz.

Çıktıyı SADECE şu JSON yapısında ver, başka hiçbir açıklama yapma:
{
  "urunler": [
    {
      "poz_no": "...", "urun_adi": "...", "urun_olcusu": "...", "kategori": "...",
      "adet": "...", "birim": "...", "uretici": "...", "model_tip": "...",
      "yangin_dayanimi": "...", "malzeme_turu": "...", "uygulama_alani": "...", "sertifika_ref": "..."
    }
  ]
}
"""

PROMPT_MIKTAR = """
Bu bir Alman ihale dosyası (LV - Leistungsverzeichnis). Görevin SADECE şudur:
Belgedeki HER TEK pozisyon (Poz-Nr.) için Menge (miktar/adet) ve Einheit (birim, örn: St, m, m², kg) değerlerini eksiksiz çıkarmak.

⚠️ ÖNEMLİ NOT - TERS OKUMA MODELİ:
Bu belge SAYFALARI TERS SIRADA sunulmıştır. 

🎯 BU BELGENİN YAPISI:
Adet bilgisi serbest metin içinde, genellikle o pozun metninin EN SONUNDA geçiyor ("...Zulassungs-Nr. Z-41.3-368. 17,000 St").

DİĞER KURALLAR:
- Ürün açıklaması, kategori, üretici vb. YAZMA — sadece poz no + miktar + birim eşleşmesi.
- Poz numarasını TAM olarak (nokta dahil) yaz.
- DIN norm numaralarını, sınıf kodlarını adet sanma.
- Emin olmadığın poz için adet değerini "-" yap.

Çıktıyı SADECE şu JSON yapısında ver:
{
  "pozlar": [
    {"poz_no": "...", "adet": "...", "birim": "..."}
  ]
}
"""

BEYAZ_LISTE = [
    "kabelschott", "kabelabschottung", "rohrschott", "rohrabschottung",
    "kombischott", "kombiabschottung", "weichschott", "hartschott", "brandschott",
    "brandschutzklappe", "brandschutzmanschette", "rohrmanschette", "brandmanschette",
    "brandschutzkissen", "deckenschott", "bodenschott", "wandschott",
    "abschottung", "abschottungssystem", "feuerschutzabschluss", "fsu",
]

YASAKLI_KELIMELER = [
    "wetterschutzgitter", "luftdurchlass", "volumenstromregelgerät", "regelgerät",
    "schalldämpfer", "ventilatoreinsatz", "montagehalterung", "zweitraumset",
    "innenblende", "rohbauset", "fassadenblende", "dachschneidemantel", "lüftungsleitung",
]

def temizle(ham_urunler, miktar_sozlugu=None):
    miktar_sozlugu = miktar_sozlugu or {}
    temiz_urunler = []
    for u in ham_urunler:
        metin = (
            str(u.get("poz_no", "")) + " " +
            str(u.get("urun_adi", "")) + " " +
            str(u.get("kategori", "")) + " " +
            str(u.get("model_tip", "")) + " " +
            str(u.get("malzeme_turu", ""))
        ).lower()

        if not any(beyaz in metin for beyaz in BEYAZ_LISTE):
            continue

        if any(yasak in metin for yasak in YASAKLI_KELIMELER):
            continue

        olcu = str(u.get("urun_olcusu", "")).strip()
        olcu_lower = olcu.lower()
        if "din" in olcu_lower or re.match(r'^(k|f|s|r|t|i)\s?\d{2,3}[\s-]', olcu_lower):
            olcu = ""

        if not olcu or olcu == "..." or olcu.lower() == "none":
            aranan_metin = str(u.get("model_tip", "")) + " " + str(u.get("urun_adi", ""))
            aranan_metin_temiz = re.sub(r'\b(DIN|EN|ISO)\s?\d+\b', '', aranan_metin, flags=re.IGNORECASE)
            bulunan = re.search(r'\b(DN)?\d{2,4}(?:[xX]\d{2,4})*\b', aranan_metin_temiz)
            u["urun_olcusu"] = bulunan.group(0) if bulunan else "-"
        else:
            u["urun_olcusu"] = olcu

        uretici = str(u.get("uretici", "")).strip().upper()
        if "KA2" in uretici or "EU" in uretici or len(uretici) > 25 or not uretici:
            u["uretici"] = "-"

        uygulama = str(u.get("uygulama_alani", "")).strip().lower()
        if "massivwände" in uygulama or "leichtbauwände" in uygulama or len(uygulama) > 50 or not uygulama:
            u["uygulama_alani"] = "-"

        sertifika = str(u.get("sertifika_ref", "")).strip()
        if len(sertifika) > 30:
            u["sertifika_ref"] = sertifika[:30] + "..."

        poz_no = str(u.get("poz_no", "")).strip()
        if poz_no in miktar_sozlugu:
            u["adet"] = miktar_sozlugu[poz_no]["adet"]
            if miktar_sozlugu[poz_no]["birim"]:
                u["birim"] = miktar_sozlugu[poz_no]["birim"]

        temiz_urunler.append(u)
    return temiz_urunler


def pdf_ters_metin_cikar(pdf_yolu: str) -> str:
    sayfalar = []
    try:
        with pdfplumber.open(pdf_yolu) as pdf:
            for sayfa in reversed(pdf.pages):
                metin = sayfa.extract_text(layout=False) or ""
                temiz_satirlar = []
                for satir in metin.split("\n"):
                    s = satir.strip()
                    if re.match(r'^Seite\s+\d+', s, re.IGNORECASE):
                        continue
                    if re.match(r'^(OZ\s+BESCHREIBUNG|Alle\s+Einzelpreise)', s, re.IGNORECASE):
                        continue
                    temiz_satirlar.append(satir)
                sayfalar.append("\n".join(temiz_satirlar))
        return "\n".join(sayfalar)
    except Exception as e:
        log_ve_hata_yaz(e, dosya_adi=os.path.basename(pdf_yolu), islem_adi="pdf_ters_metin_cikar")
        raise e


def miktar_tablosu_cikar(dosya, client, durum_alani, orijinal_pdf_yolu=None):
    ham_dosya_adi = os.path.basename(orijinal_pdf_yolu) if orijinal_pdf_yolu else "Unbekanntes PDF"
    dosya_adi = re.sub(r'^temp_\d+_', '', ham_dosya_adi)
    
    log_ekle("INFO", f"[{dosya_adi}] Schritt 1/2: Mengentabelle wird gescannt...")
    
    ters_metin = None
    if orijinal_pdf_yolu:
        try:
            ters_metin = pdf_ters_metin_cikar(orijinal_pdf_yolu)
        except Exception:
            durum_alani.warning(f"⚠️ Umgekehrter Text konnte nicht erstellt werden, Original-PDF wird verwendet.")

    max_deneme = 3
    for deneme in range(max_deneme):
        try:
            icerik = [ters_metin, PROMPT_MIKTAR] if ters_metin else [dosya, PROMPT_MIKTAR]

            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=icerik,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.0
                )
            )
            if response.text:
                data = json.loads(response.text)
                pozlar = data.get("pozlar", [])
                sozluk = {}
                for p in pozlar:
                    poz_no = str(p.get("poz_no", "")).strip()
                    adet = str(p.get("adet", "")).strip()
                    birim = str(p.get("birim", "")).strip()
                    if poz_no and adet and adet != "-":
                        sozluk[poz_no] = {"adet": adet, "birim": birim}
                return sozluk
            return {}
        except Exception as e:
            log_ve_hata_yaz(e, dosya_adi=dosya_adi, islem_adi=f"miktar_tablosu_cikar (Versuch {deneme+1})")
            err_str = str(e)
            if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str or "503" in err_str or "UNAVAILABLE" in err_str:
                wait_time = (deneme + 1) * 15
                log_ekle("WARNING", f"[{dosya_adi}] API-Überlastung/Rate-Limit. Wartezeit: {wait_time}s...")
                time.sleep(wait_time)
            else:
                return {}
    return {}


def analiz_et_guvenli(pdf_yolu, client, durum_alani):
    ham_dosya_adi = os.path.basename(pdf_yolu)
    dosya_adi = re.sub(r'^temp_\d+_', '', ham_dosya_adi)

    try:
        dosya = client.files.upload(file=pdf_yolu)
        while dosya.state.name == "PROCESSING":
            time.sleep(2)
            dosya = client.files.get(name=dosya.name)
    except Exception as e:
        log_ve_hata_yaz(e, dosya_adi=dosya_adi, islem_adi="client.files.upload")
        raise e

    durum_alani.info(f"🔢 Mengentabelle wird extrahiert: {dosya_adi}...")
    miktar_sozlugu = miktar_tablosu_cikar(dosya, client, durum_alani, orijinal_pdf_yolu=pdf_yolu)

    log_ekle("INFO", f"[{dosya_adi}] Schritt 2/2: FSU-Produktanalyse läuft...")
    
    max_deneme = 5
    for deneme in range(max_deneme):
        try:
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=[dosya, PROMPT],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.0
                )
            )
            if response.text:
                data = json.loads(response.text)
                ham_urunler = data.get("urunler", [])
                temiz = temizle(ham_urunler, miktar_sozlugu)
                
                log_ekle("SUCCESS", f"[{dosya_adi}] Gefilterte FSU-Anzahl: {len(temiz)}")
                return temiz
            return []

        except Exception as e:
            log_ve_hata_yaz(e, dosya_adi=dosya_adi, islem_adi=f"generate_content (Versuch {deneme+1})")
            err_str = str(e)
            if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str or "503" in err_str or "UNAVAILABLE" in err_str:
                wait_time = (deneme + 1) * 15
                log_ekle("WARNING", f"[{dosya_adi}] Server beschäftigt/Limit erreicht. Warte {wait_time} Sekunden...")
                time.sleep(wait_time)
            else:
                if deneme == max_deneme - 1:
                    raise e
                time.sleep(3)

    return []

# ==============================================================================
# DATEI-UPLOAD UND PROZESS-BEDIENFELD
# ==============================================================================
yuklenen_dosyalar = st.file_uploader(
    "LV-Dateien (PDF) auswählen",
    type=["pdf"],
    accept_multiple_files=True
)

if yuklenen_dosyalar:
    st.info(f"📄 {len(yuklenen_dosyalar)} Datei(en) ausgewählt.")

    if not SECILEN_API_KEY:
        st.error("⚠️ GEMINI_API_KEY fehlt! Bitte überprüfen Sie Ihre secrets.toml Datei.")
    elif st.button("🚀 Analyse starten", type="primary"):
        log_ekle("INFO", "ANALYSE GESTARTET...")
        
        st.session_state["analiz_sonuclari"] = []
        client = genai.Client(api_key=SECILEN_API_KEY)

        # Deutshe Spaltenüberschriften für Excel
        spalten_de = [
            "PDF-Datei", "Pos.-Nr.", "Produktname", "Abmessung", "Kategorie", 
            "Menge", "Einheit", "Hersteller", "Modell/Typ", "Feuerwiderstand", 
            "Material", "Anwendungsbereich", "Zertifikat-Ref"
        ]

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Brandschutzprodukte"
        ws.append(spalten_de)

        tum_urunler = []
        genel_progress = st.progress(0)

        for i, yuklenen_dosya in enumerate(yuklenen_dosyalar):
            durum_alani = st.empty()
            durum_alani.info(f"⏳ In Bearbeitung: {yuklenen_dosya.name} ({i+1}/{len(yuklenen_dosyalar)})")

            gecici_yol = f"temp_{i}_{yuklenen_dosya.name}"
            try:
                with open(gecici_yol, "wb") as f:
                    f.write(yuklenen_dosya.getbuffer())

                urunler = analiz_et_guvenli(gecici_yol, client, durum_alani)
                for u in urunler:
                    satir = [
                        yuklenen_dosya.name, str(u.get("poz_no", "")), str(u.get("urun_adi", "")),
                        str(u.get("urun_olcusu", "")), str(u.get("kategori", "")), str(u.get("adet", "")),
                        str(u.get("birim", "")), str(u.get("uretici", "")), str(u.get("model_tip", "")),
                        str(u.get("yangin_dayanimi", "")), str(u.get("malzeme_turu", "")),
                        str(u.get("uygulama_alani", "")), str(u.get("sertifika_ref", ""))
                    ]
                    ws.append(satir)
                    tum_urunler.append(dict(zip(spalten_de, satir)))
                
                durum_alani.success(f"✓ {yuklenen_dosya.name}: {len(urunler)} FSU-Produkt(e) gefunden.")

            except Exception as genel_hata:
                log_ve_hata_yaz(genel_hata, dosya_adi=yuklenen_dosya.name, islem_adi="Datei-Analyse-Schleife")
                st.error(f"🚨 Fehler beim Verarbeiten von **{yuklenen_dosya.name}**.")
            finally:
                if os.path.exists(gecici_yol):
                    os.remove(gecici_yol)

            genel_progress.progress((i + 1) / len(yuklenen_dosyalar))

        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        
        st.session_state["analiz_sonuclari"] = tum_urunler
        st.session_state["excel_data"] = excel_buffer.getvalue()
        
        log_ekle("INFO", "VERARBEITUNG ABGESCHLOSSEN...")
        st.rerun()

# --- ERGEBNISANZEIGE ---
if st.session_state.get("analiz_sonuclari"):
    tum_urunler = st.session_state["analiz_sonuclari"]
    excel_bytes = st.session_state.get("excel_data")

    st.success(f"🎉 Vorgang abgeschlossen! Insgesamt {len(tum_urunler)} FSU-Produkte identifiziert.")
    st.subheader("📊 Analyseteergebnisse")
    st.dataframe(tum_urunler, use_container_width=True)

    if excel_bytes:
        st.download_button(
            label="📥 Excel-Bericht herunterladen (.xlsx)",
            data=excel_bytes,
            file_name="Brandschutz_Produkte_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
