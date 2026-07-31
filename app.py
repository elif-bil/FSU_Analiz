import os
import io
import time
import json
import re
import tempfile
import openpyxl
import pdfplumber
import streamlit as st
from google import genai
from google.genai import types

# --- STREAMLIT ARAYÜZ AYARLARI ---
st.set_page_config(page_title="Brandschutz LV Analiz Sistemi", layout="wide")

# --- BASİT ŞİFRE KORUMASI ---
UYGULAMA_SIFRESI = "Yangin2026"

def sifre_kontrolu():
    def sifre_girildi():
        if st.session_state.get("girilen_sifre") == UYGULAMA_SIFRESI:
            st.session_state["dogrulandi"] = True
            del st.session_state["girilen_sifre"]
        else:
            st.session_state["dogrulandi"] = False

    if st.session_state.get("dogrulandi", False):
        return True

    st.text_input("🔒 Şifre", type="password", on_change=sifre_girildi, key="girilen_sifre")
    if "dogrulandi" in st.session_state and not st.session_state["dogrulandi"]:
        st.error("❌ Yanlış şifre, tekrar deneyin.")
    return False

if not sifre_kontrolu():
    st.stop()

st.title("🔥 Alman İhale Dosyası (LV) Pasif Yangın Analiz Sistemi — FSU Ürünleri")
st.markdown(
    "Bir veya birden fazla LV (PDF) dosyası yükleyin. Sistem her dosyayı tarayıp "
    "**sadece FSU ürünlerini** ayıklayacak ve sonuçları listeleyip Excel olarak indirmenizi sağlayacaktır."
)

# --- API KEY ---
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
- Deckenschott / Bodenschott / Wandschott (tavan/döşeme/duvar geçiş durdurucusu — dikkat: bunlar "Wand"/"Dach" kelimesi geçse bile FSU'dur çünkü asıl işlev geçiş mühürlemesidir, bunu ele geçirme)
- İçinde "Brandschutz" + "Schott/Klappe/Manschette/Kissen/Abschottung" birlikte geçen ürünler
Bu ürünlerin GENELDE (her zaman değil) bir yangın dayanım sınıfı (F30, F90, K30, K90, S90, R30 vb.) ve bir yapı onay/Zulassungsnummer'ı (Z-19.xx, Z-41.xx, abP, ETA vb.) olur — bu bilgiler varsa güçlü bir FSU işaretidir, ama olmaması otomatik ret sebebi değildir, önce yukarıdaki tanıma bak.

🚫 KESİNLİKLE FSU DEĞİL - ASLA SEÇME:
Genel havalandırma/HVAC ekipmanı FSU DEĞİLDİR, yangınla ilgisi olmayan hava akışını sağlayan/yönlendiren parçalardır:
- Ventilatoreinsatz, Montagehalterung, Zweitraumset, Innenblende, Rohbauset, Fassadenblende (genel montaj/adaptör parçaları)
- Wetterschutzgitter (hava koşulu koruma ızgarası), Luftdurchlass (hava difüzörü), Volumenstromregelgerät/VVS Regelgerät (debi kontrol), Zusatzschalldämpfer (susturucu)
- Sadece çatı/duvar geçişi için genel kılıf/manto olup yangın mühürleme işlevi belirtilmeyen "Dachdurchführung", "Wanddurchführung", "Dachschneidemantel", "Lüftungsleitung" (bunlar sadece montaj kılıfıdır, mühürleme sistemi değildir — ancak açıkça "Brandschutz-Wanddurchführung" gibi yangın onayı belirtilen bir ürünse bu istisnadır, dikkatli oku)
Şüphen varsa ve ürün açıkça yukarıdaki beyaz listedeki bir kategoriye girmiyorsa, DAHİL ETME.

🚨 UYDURMA / HALÜSİNASYON YASAĞI:
- **uygulama_alani**: Metinde ürünün binanın neresine takılacağı açıkça yazmıyorsa (örn: Kat 2, Büro 104 vb.), mahalleri veya duvar tiplerini (Massivwände, Leichtbauwände vb.) UYDURMA. Yazmıyorsa "-" bırak.

📏 DİĞER KURALLAR:
- **urun_olcusu**: SADECE ürünün fiziksel boyutunu/çapını yaz (Örn: 1000x700x500 mm, DN100, çap 125 mm). Bu bilgi genelde "Nenngröße", "Abmessung", "Durchmesser", "Baugröße" ifadelerinin yanında veya ürün adının sonunda (örn: "Deckenschott 100" → 100) geçer.
  KESİNLİKLE ÖLÇÜ SAYMA: DIN/EN/ISO norm referans numaraları (örn: "DIN 18017", "DIN 24145", "EN 1366-3"), yangın dayanım sınıfı + onay kodu birleşimleri (örn: "K 90-18017", "F90-Z-19.11-123" — bunların hepsi yangin_dayanimi/sertifika_ref alanına gider, ölçü değildir), Artikelnummer/ürün kodları. Bu tür sayıları gördüğünde ölçü alanına ASLA yazma. Metinde gerçek bir fiziksel boyut geçmiyorsa "-" bırak.
- **uretici**: "KA2-EU-DE" veya benzeri standart kodları ASLA üretici olamaz. Üretici sadece TROX, Hilti, Helios vb. gerçek firma adıdır. Emin değilsen "-" yap.
- **sertifika_ref**: Devasa norm zincirlerini yazma, sadece kısa onay numarasını yaz.

TÜM SAYFALARI TARA VE HİÇBİR SATIRI ATLAMA.

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
Bu belge SAYFALARI TERS SIRADA sunulmuştur (son sayfa önce, ilk sayfa sonda). Bu kasıtlı yapılmıştır çünkü adet bilgisi her pozun metninin SONUNDA gelir — ters okumada bu bilgi artık her pozun metninin BAŞINDA yer alır ve çok daha kolay tespit edilir.
Okuma yönün değişmiş olsa da poz numaralarını TAM olarak (nokta dahil) çıkar ve doğru adet/birimle eşleştir.

🎯 BU BELGENİN YAPISI:
Bu belgede adet bilgisi AYRI BİR TABLO SÜTUNUNDA değil, her pozun uzun açıklama metninin (Langtext) EN SONUNDA (normal okuma sırasında), serbest metin içinde geçiyor. Tipik sıralama şöyledir:
  ... ürün teknik açıklaması ... Feuerwiderstandsklasse [sınıf] ... Zulassungs-Nr. [onay no]. [ADET],000 [BİRİM]
Yani doğru adet, genellikle o pozun paragrafındaki EN SON sayıdır ve hemen "Zulassungs-Nr." referansından SONRA, bir sonraki Poz-Nr. başlamadan HEMEN ÖNCE gelir. Bu son sayı çoğu zaman "St", "m", "m²" gibi bir birimle birlikte yazılıdır.

GERÇEK ÖRNEK (bu belgeden, birebir bu kalıba dikkat et):
Normal sırada metin şöyle akıyor: "...Wartungsfrei und ohne Inspektionsauflagen. Feuerwiderstandsklasse K 90-18017, Zulassungs-Nr. Z-41.3-368. 17,000 St"
Burada:
- "K 90-18017" → bu bir SINIF KODUDUR, adet DEĞİLDİR.
- "Z-41.3-368" → bu bir ONAY NUMARASIDIR, adet DEĞİLDİR.
- "17,000 St" → İŞTE DOĞRU ADET BUDUR (17, birim: St).
Ters sırada bu bilgi ("17,000 St ... Zulassungs-Nr. Z-41.3-368 ...") ilgili pozun BAŞINA yakın gelecektir — bu sayede adet bilgisi çok daha kolay yakalanır.

DİĞER KURALLAR:
- Ürün açıklaması, kategori, üretici vb. YAZMA — sadece poz no + miktar + birim eşleşmesi istiyorum.
- Poz numarasını TAM olarak (nokta dahil, örn: 01.01.0140) yaz, kısaltma.
- Bir pozun paragrafındaki DIN norm numaraları (DIN 18017 gibi), sınıf+onay kodları (K 90-18017, Z-41.3-368 gibi), Artikelnummer gibi ARA sayıları asla adet sanma. Adet, o pozun metninin GERÇEK SONUNDA (normal okumada), genelde bir birim kısaltmasıyla (St/m/m²/kg/Stk) birlikte yazan sayıdır.
- Emin olmadığın bir poz için adet değerini boş bırak ("-" yaz), ASLA tahmin etme veya başka bir pozun değerini kopyalama.
- TÜM belgeyi tara, hiçbir pozu atlama.

Çıktıyı SADECE şu JSON yapısında ver, başka hiçbir açıklama yapma:
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
    """
    PDF sayfalarını TERS sırayla okur ve tek bir metin bloğu döndürür.
    Amaç: Her pozun SONUNDA gelen adet bilgisi, ters okumada metnin
    BAŞINA gelir ve model tarafından çok daha kolay tespit edilir.
    pypdf/pikepdf gerektirmez; sadece pdfplumber kullanır.
    """
    sayfalar = []
    with pdfplumber.open(pdf_yolu) as pdf:
        for sayfa in reversed(pdf.pages):
            metin = sayfa.extract_text(layout=False) or ""
            # Sayfa numarası ve başlık satırlarını temizle
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


def miktar_tablosu_cikar(dosya, client, durum_alani, orijinal_pdf_yolu=None):
    """
    Tersten okuma yöntemi:
    PDF sayfaları pdfplumber ile TERS sırayla okunur ve metin string olarak
    Gemini'ye gönderilir. Bu sayede her pozun SONUNDA gelen adet bilgisi,
    modelin gördüğü metnin BAŞINA gelir ve çok daha kolay yakalanır.
    Ters metin gönderilemezse orijinal PDF dosyasıyla fallback yapar.
    """
    # Ters metin oluştur
    ters_metin = None
    if orijinal_pdf_yolu:
        try:
            ters_metin = pdf_ters_metin_cikar(orijinal_pdf_yolu)
        except Exception as e:
            durum_alani.warning(f"⚠️ Ters metin oluşturulamadı, orijinal PDF kullanılıyor: {e}")

    max_deneme = 3
    for deneme in range(max_deneme):
        try:
            # Ters metin varsa text olarak, yoksa orijinal PDF dosyasıyla gönder
            if ters_metin:
                icerik = [ters_metin, PROMPT_MIKTAR]
            else:
                icerik = [dosya, PROMPT_MIKTAR]

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
            hata_mesaji = str(e)
            if "429" in hata_mesaji or "RESOURCE_EXHAUSTED" in hata_mesaji:
                bekleme_suresi = (deneme + 1) * 30
                durum_alani.warning(f"⚠️ Miktar tablosu için kota bekleniyor ({bekleme_suresi}sn)...")
                time.sleep(bekleme_suresi)
            else:
                durum_alani.warning(f"⚠️ Miktar tablosu çıkarılamadı: {e}. Ürün detaylarındaki adet kullanılacak.")
                return {}
    return {}


def analiz_et_guvenli(pdf_yolu, client, durum_alani):
    dosya = client.files.upload(file=pdf_yolu)
    while dosya.state.name == "PROCESSING":
        time.sleep(3)
        dosya = client.files.get(name=dosya.name)

    durum_alani.info(f"🔢 Miktar tablosu çıkarılıyor (tersten okuma): {os.path.basename(pdf_yolu)}...")
    miktar_sozlugu = miktar_tablosu_cikar(dosya, client, durum_alani, orijinal_pdf_yolu=pdf_yolu)

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

                with st.expander(f"🔍 Debug bilgisi: {os.path.basename(pdf_yolu)}"):
                    st.write(f"Miktar tablosunda bulunan poz sayısı: **{len(miktar_sozlugu)}**")
                    if miktar_sozlugu:
                        st.write("Miktar tablosundan örnek (ilk 10 poz):")
                        st.json(dict(list(miktar_sozlugu.items())[:10]))
                    st.write(f"Gemini'nin bulduğu ham ürün sayısı: **{len(ham_urunler)}**")
                    st.write(f"Filtre sonrası kalan ürün sayısı: **{len(temiz)}**")

                return temiz
            return []

        except json.JSONDecodeError as e:
            durum_alani.error(f"JSON ayrıştırma hatası: {e}")
            return []

        except Exception as e:
            hata_mesaji = str(e)
            if "429" in hata_mesaji or "RESOURCE_EXHAUSTED" in hata_mesaji:
                bekleme_suresi = (deneme + 1) * 30
                durum_alani.warning(f"⚠️ Kota sınırı (429) aşıldı. {bekleme_suresi} saniye bekleniyor ({deneme + 1}/{max_deneme})...")
                time.sleep(bekleme_suresi)
            else:
                durum_alani.error(f"Hata: {e}")
                return []

    durum_alani.error("❌ Maksimum deneme sınırına ulaşıldı, bu dosya atlandı.")
    return []


yuklenen_dosyalar = st.file_uploader(
    "LV (PDF) Dosyalarını Seçin (birden fazla seçebilirsiniz)",
    type=["pdf"],
    accept_multiple_files=True
)

if yuklenen_dosyalar:
    st.info(f"📄 {len(yuklenen_dosyalar)} dosya yüklendi.")

    if not SECILEN_API_KEY:
        st.warning("⚠️ API key bulunamadı. Lütfen secrets.toml dosyasına GEMINI_API_KEY değerini ekleyin.")
    elif st.button("🚀 Analizi Başlat", type="primary"):
        client = genai.Client(api_key=SECILEN_API_KEY)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Yangin Urunleri"
        ws.append(["PDF", "Poz", "Ürün Adı", "Ölçü", "Kategori", "Adet", "Birim",
                   "Üretici", "Model/Tip", "Dayanım", "Malzeme", "Uygulama", "Sertifika"])

        tum_urunler = []
        genel_progress = st.progress(0)

        for i, yuklenen_dosya in enumerate(yuklenen_dosyalar):
            durum_alani = st.empty()
            durum_alani.info(f"⏳ İşleniyor: {yuklenen_dosya.name} ({i+1}/{len(yuklenen_dosyalar)})")

            gecici_yol = f"temp_{i}_{yuklenen_dosya.name}"
            with open(gecici_yol, "wb") as f:
                f.write(yuklenen_dosya.getbuffer())

            try:
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
                    tum_urunler.append(dict(zip(
                        ["PDF", "Poz", "Ürün Adı", "Ölçü", "Kategori", "Adet", "Birim",
                         "Üretici", "Model/Tip", "Dayanım", "Malzeme", "Uygulama", "Sertifika"],
                        satir
                    )))
                durum_alani.success(f"✓ {yuklenen_dosya.name}: {len(urunler)} FSU ürünü bulundu.")
            finally:
                if os.path.exists(gecici_yol):
                    os.remove(gecici_yol)

            genel_progress.progress((i + 1) / len(yuklenen_dosyalar))

            if i < len(yuklenen_dosyalar) - 1:
                time.sleep(5)

        if tum_urunler:
            st.success(f"🎉 Tüm dosyalar tamamlandı! Toplam {len(tum_urunler)} FSU ürünü bulundu.")
            st.subheader("📊 Bulunan FSU Ürünlerinin Önizlemesi")
            st.dataframe(tum_urunler, use_container_width=True)

            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            excel_buffer.seek(0)

            st.download_button(
                label="📥 Excel Çıktısını İndir (.xlsx)",
                data=excel_buffer,
                file_name="Yangin_Urunleri_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ Dosyalar tarandı ancak uygun FSU ürünü bulunamadı.")
