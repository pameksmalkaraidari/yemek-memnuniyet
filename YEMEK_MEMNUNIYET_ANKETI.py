import re
import io
import pandas as pd
from datetime import datetime
from zoneinfo import ZoneInfo
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from streamlit_cookies_manager import EncryptedCookieManager

# =========================================================
# SAYFA YAPILANDIRMASI
# =========================================================
st.set_page_config(
    page_title="Pameks Yemek ve Geri Bildirim Portalı",
    page_icon="🍽️",
    layout="centered",
)

# =========================================================
# GOOGLE SHEETS BAĞLANTISI
# =========================================================
# secrets.toml içine (Streamlit Cloud > Settings > Secrets) şunlar eklenmeli:
#
# sheet_id = "Sheet URL'sindeki /d/ ile /edit arasındaki kod"
#
# [gcp_service_account]
# ... (servis hesabı JSON'undaki tüm alanlar)
#
# Not: Sheet dosyasını servis hesabının client_email'iyle "Düzenleyen"
# olarak paylaşmayı unutma. Sadece Google Sheets API yeterli.

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

WS_MENU = "GunlukMenu"
WS_ON = "OnOylama"
WS_SON = "SonDegerlendirme"
WS_ONERI = "MenuOneri"

TARIH_FORMAT = "%d.%m.%Y"
TR_TZ = ZoneInfo("Europe/Istanbul")

# Aynı cihazdan/tarayıcıdan tekrar oy kullanımını engellemek için kullanılan
# çerez adları. Değerler, o cihazda hangi günlerin zaten oylandığını virgülle
# ayrılmış tarih listesi olarak tutar. Bu bir güvenlik önlemi DEĞİL, sıradan
# tekrar tıklamaları engelleyen bir sürtünme katmanıdır — çerezini silen ya
# da gizli sekme kullanan biri yine de tekrar oy kullanabilir.
COOKIE_ON_OYLAMA = "yemek_on_oylanan_gunler"
COOKIE_SON_DEGERLENDIRME = "yemek_son_degerlendirilen_gunler"
COOKIE_ONERI = "yemek_oneri_yapilan_gunler"
COOKIE_GECERLILIK_GUN = 120


def simdi_tr() -> datetime:
    """Sunucu hangi saat diliminde çalışırsa çalışsın, her zaman İstanbul
    yerel saatini döndürür. 'Bugün' hesaplamalarında tutarlılık için tüm
    tarih/saat işlemleri bu fonksiyon üzerinden yapılmalı."""
    return datetime.now(TR_TZ)


def tarih_araligina_filtrele(df: pd.DataFrame, baslangic, bitis) -> pd.DataFrame:
    """df'in 'Tarih' sütununu (DD.MM.YYYY metni) ayrıştırıp [baslangic, bitis]
    (date nesneleri, iki uç dahil) aralığındaki satırları döndürür."""
    if df is None or df.empty or "Tarih" not in df.columns:
        return pd.DataFrame(columns=(df.columns if df is not None else []))
    d = df.copy()
    d["_t"] = pd.to_datetime(d["Tarih"], format=TARIH_FORMAT, errors="coerce").dt.date
    d = d[(d["_t"] >= baslangic) & (d["_t"] <= bitis)]
    return d.drop(columns=["_t"])


def araligin_menu_metni(df_menu: pd.DataFrame, tarih_str: str) -> str:
    """Belirli bir günün menüsünü 'Kalem: Yemek | Kalem: Yemek' şeklinde tek
    satırlık metne çevirir."""
    if df_menu is None or df_menu.empty or "Tarih" not in df_menu.columns:
        return ""
    eslesen = df_menu[df_menu["Tarih"] == tarih_str]
    if eslesen.empty:
        return ""
    satir = eslesen.iloc[0]
    parcalar = [
        f"{kolon}: {satir[kolon]}"
        for kolon in df_menu.columns
        if kolon != "Tarih" and str(satir[kolon]).strip()
    ]
    return " | ".join(parcalar)


def ozet_excel_raporu(
    baslangic,
    bitis,
    df_on: pd.DataFrame,
    df_son: pd.DataFrame,
    df_oneri: pd.DataFrame,
    df_menu: pd.DataFrame,
) -> bytes:
    """Seçilen tarih aralığı (baslangic/bitis: date, iki uç dahil) için genel
    özet, günlük menü + ön oylama, yemek bazlı performans, günlük detay,
    kötü değerlendirmeler, öneri/şikâyetler ve menü sayfalarından oluşan
    çok sayfalı bir Excel raporu (bytes) üretir."""
    on_ar = tarih_araligina_filtrele(df_on, baslangic, bitis)
    son_ar = tarih_araligina_filtrele(df_son, baslangic, bitis)
    oneri_ar = tarih_araligina_filtrele(df_oneri, baslangic, bitis)
    menu_ar = tarih_araligina_filtrele(df_menu, baslangic, bitis)

    if not son_ar.empty:
        son_ar = son_ar.copy()
        son_ar["_puan_sayi"] = pd.to_numeric(son_ar["Degerlendirme"], errors="coerce")
        son_ar["_kategori"] = son_ar["Degerlendirme"].apply(puan_kategorisi)
        son_ar["_yemek"] = son_ar.apply(
            lambda s: s["UrunAdi"] if str(s.get("UrunAdi", "")).strip() else s["Kalem"], axis=1
        )

    def _tarihe_gore_sirala(df, artan=True):
        df = df.copy()
        df["_s"] = pd.to_datetime(df["Tarih"], format=TARIH_FORMAT, errors="coerce")
        df = df.sort_values("_s", ascending=artan).drop(columns=["_s"])
        return df

    tampon = io.BytesIO()
    with pd.ExcelWriter(tampon, engine="openpyxl") as yazici:
        # --- 1. Genel Özet ---
        toplam_on = len(on_ar)
        iyi = int((on_ar["Degerlendirme"] == "İyi").sum()) if not on_ar.empty else 0
        kotu_on = int((on_ar["Degerlendirme"] == "Kötü").sum()) if not on_ar.empty else 0
        iyi_orani = round(iyi / toplam_on * 100, 1) if toplam_on else 0

        toplam_son = len(son_ar)
        ort_puan = round(son_ar["_puan_sayi"].mean(), 2) if not son_ar.empty else 0
        kotu_son = int((son_ar["_kategori"] == "Kötü").sum()) if not son_ar.empty else 0
        kotu_orani = round(kotu_son / toplam_son * 100, 1) if toplam_son else 0

        pd.DataFrame([
            {"Alan": "Rapor Aralığı", "Değer": f"{baslangic.strftime('%d.%m.%Y')} - {bitis.strftime('%d.%m.%Y')}"},
            {"Alan": "Toplam Ön Oylama", "Değer": toplam_on},
            {"Alan": "Ön Oylama İyi", "Değer": iyi},
            {"Alan": "Ön Oylama Kötü", "Değer": kotu_on},
            {"Alan": "Ön Oylama İyi Oranı (%)", "Değer": iyi_orani},
            {"Alan": "Toplam Ürün Değerlendirmesi", "Değer": toplam_son},
            {"Alan": "Ortalama Puan (1-5)", "Değer": ort_puan},
            {"Alan": "Kötü Değerlendirme Sayısı (1-2 Puan)", "Değer": kotu_son},
            {"Alan": "Kötü Oranı (%)", "Değer": kotu_orani},
            {"Alan": "Toplam Öneri / Şikâyet", "Değer": len(oneri_ar)},
        ]).to_excel(yazici, sheet_name="Genel Özet", index=False)

        # --- 2. Menü Ön Değerlendirme (gün + o günün menüsü + ön oylama) ---
        tum_tarihler = sorted(
            set(menu_ar["Tarih"]) | set(on_ar["Tarih"] if not on_ar.empty else []),
            key=lambda t: datetime.strptime(t, TARIH_FORMAT),
        )
        if not tum_tarihler:
            pd.DataFrame([{"Sonuç": "Bu tarih aralığında menü ya da ön oylama verisi yok."}]).to_excel(
                yazici, sheet_name="Menü Ön Değerlendirme", index=False
            )
        else:
            satirlar = []
            for t in tum_tarihler:
                gunluk_on = on_ar[on_ar["Tarih"] == t] if not on_ar.empty else pd.DataFrame()
                g_iyi = int((gunluk_on["Degerlendirme"] == "İyi").sum())
                g_kotu = int((gunluk_on["Degerlendirme"] == "Kötü").sum())
                g_toplam = len(gunluk_on)
                satirlar.append({
                    "Tarih": t,
                    "Gün": gun_adi(t),
                    "Menü": araligin_menu_metni(df_menu, t),
                    "İyi": g_iyi,
                    "Kötü": g_kotu,
                    "Toplam Oy": g_toplam,
                    "İyi Oranı (%)": round(g_iyi / g_toplam * 100, 1) if g_toplam else "",
                })
            pd.DataFrame(satirlar).to_excel(yazici, sheet_name="Menü Ön Değerlendirme", index=False)

        # --- 3. Yemek Bazlı Genel Performans (en kötüden en iyiye) ---
        if son_ar.empty:
            pd.DataFrame([{"Sonuç": "Bu tarih aralığında değerlendirme verisi yok."}]).to_excel(
                yazici, sheet_name="Yemek Performansı", index=False
            )
        else:
            perf = son_ar.groupby("_yemek").agg(
                Değerlendirme_Sayısı=("_puan_sayi", "count"),
                Ortalama_Puan=("_puan_sayi", "mean"),
            )
            kotu_sayilar = son_ar[son_ar["_kategori"] == "Kötü"].groupby("_yemek").size()
            perf["Kötü_Sayısı"] = kotu_sayilar.reindex(perf.index, fill_value=0)
            perf["Kötü_Oranı"] = (perf["Kötü_Sayısı"] / perf["Değerlendirme_Sayısı"] * 100).round(1)
            perf["Ortalama_Puan"] = perf["Ortalama_Puan"].round(2)
            perf = perf.sort_values("Ortalama_Puan")
            perf.index.name = "Yemek"
            perf.reset_index().rename(columns={
                "Değerlendirme_Sayısı": "Değerlendirme Sayısı",
                "Ortalama_Puan": "Ortalama Puan",
                "Kötü_Sayısı": "Kötü Sayısı",
                "Kötü_Oranı": "Kötü Oranı (%)",
            }).to_excel(yazici, sheet_name="Yemek Performansı", index=False)

        # --- 4. Günlük Detay (her gün, her yemek) ---
        if son_ar.empty:
            pd.DataFrame([{"Sonuç": "Veri yok."}]).to_excel(yazici, sheet_name="Günlük Detay", index=False)
        else:
            detay = son_ar.groupby(["Tarih", "Kalem", "_yemek"]).agg(
                Ortalama_Puan=("_puan_sayi", "mean"),
                Değerlendirme_Sayısı=("_puan_sayi", "count"),
            ).reset_index()
            detay["Ortalama_Puan"] = detay["Ortalama_Puan"].round(2)
            detay = detay.rename(columns={
                "Kalem": "Öğün Kalemi", "_yemek": "Yemek",
                "Ortalama_Puan": "Ortalama Puan", "Değerlendirme_Sayısı": "Değerlendirme Sayısı",
            })
            _tarihe_gore_sirala(detay).to_excel(yazici, sheet_name="Günlük Detay", index=False)

        # --- 5. Kötü Değerlendirmeler ve Açıklamalar ---
        kotuler = son_ar[son_ar["_kategori"] == "Kötü"] if not son_ar.empty else pd.DataFrame()
        if kotuler.empty:
            pd.DataFrame([{"Sonuç": "Bu tarih aralığında 1-2 puan verilen ürün yok."}]).to_excel(
                yazici, sheet_name="Kötü Değerlendirmeler", index=False
            )
        else:
            kotuler_sirali = kotuler[["Tarih", "Kalem", "UrunAdi", "Degerlendirme", "Aciklama"]].rename(
                columns={"Kalem": "Öğün Kalemi", "UrunAdi": "Yemek", "Degerlendirme": "Puan", "Aciklama": "Açıklama"}
            )
            _tarihe_gore_sirala(kotuler_sirali, artan=False).to_excel(
                yazici, sheet_name="Kötü Değerlendirmeler", index=False
            )

        # --- 6. Öneriler / Şikâyetler ---
        if oneri_ar.empty:
            pd.DataFrame([{"Sonuç": "Bu tarih aralığında öneri/şikâyet yok."}]).to_excel(
                yazici, sheet_name="Öneriler-Şikayetler", index=False
            )
        else:
            oneri_sirali = oneri_ar[["Tarih", "Oneri"]].rename(columns={"Oneri": "Öneri / Şikâyet"})
            _tarihe_gore_sirala(oneri_sirali, artan=False).to_excel(yazici, sheet_name="Öneriler-Şikayetler", index=False)

        # --- 7. Menü (Referans) ---
        if menu_ar.empty:
            pd.DataFrame([{"Sonuç": "Bu tarih aralığında menü verisi yok."}]).to_excel(
                yazici, sheet_name="Menü (Referans)", index=False
            )
        else:
            _tarihe_gore_sirala(menu_ar).to_excel(yazici, sheet_name="Menü (Referans)", index=False)

    return tampon.getvalue()


@st.cache_resource(show_spinner=False)
def sheets_baglantisi_al():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPES
    )
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(st.secrets["sheet_id"])
    return spreadsheet


@st.cache_resource(show_spinner=False)
def worksheet_al(_spreadsheet, isim, kolonlar):
    try:
        ws = _spreadsheet.worksheet(isim)
    except gspread.exceptions.WorksheetNotFound:
        ws = _spreadsheet.add_worksheet(title=isim, rows=1000, cols=max(len(kolonlar), 5))
        ws.append_row(kolonlar)
    return ws


@st.cache_data(ttl=60, show_spinner=False)
def df_oku(_worksheet, worksheet_adi: str) -> pd.DataFrame:
    """gspread'in get_all_records() fonksiyonu, sayfada boş veya birbirini
    tekrar eden başlık hücreleri olduğunda GSpreadException fırlatır.
    Bunun yerine ham veriyi çekip başlıkları kendimiz temizliyoruz, böylece
    sayfa elle düzenlenmiş / bozulmuş olsa bile uygulama çökmez."""
    try:
        degerler = _worksheet.get_all_values()
    except Exception as e:
        st.warning(
            f"'{worksheet_adi}' sayfası okunurken bir sorun oluştu, boş kabul edildi: {e}"
        )
        return pd.DataFrame()

    if not degerler:
        return pd.DataFrame()

    ham_basliklar = degerler[0]
    satirlar = degerler[1:]

    # Boş / tekrarlı başlıkları benzersiz hale getir
    basliklar = []
    gorulme = {}
    for i, b in enumerate(ham_basliklar):
        b = (b or "").strip()
        if not b:
            b = f"Sutun{i + 1}"
        if b in gorulme:
            gorulme[b] += 1
            b = f"{b}_{gorulme[b]}"
        else:
            gorulme[b] = 0
        basliklar.append(b)

    # Satırların hücre sayısını başlık sayısına eşitle (eksikse boş, fazlaysa kes)
    kolon_sayisi = len(basliklar)
    duzeltilmis_satirlar = []
    for satir in satirlar:
        if len(satir) < kolon_sayisi:
            satir = satir + [""] * (kolon_sayisi - len(satir))
        elif len(satir) > kolon_sayisi:
            satir = satir[:kolon_sayisi]
        duzeltilmis_satirlar.append(satir)

    # Tamamen boş satırları at
    duzeltilmis_satirlar = [s for s in duzeltilmis_satirlar if any(str(x).strip() for x in s)]

    return pd.DataFrame(duzeltilmis_satirlar, columns=basliklar)


def satir_ekle(worksheet, satir: list):
    worksheet.append_row(satir, value_input_option="USER_ENTERED")



def menuyu_yukle(worksheet, df: pd.DataFrame):
    """Mevcut menünün üzerine tamamen yeni menüyü yazar."""
    df = df.copy()
    df["Tarih"] = pd.to_datetime(df["Tarih"], dayfirst=True, errors="coerce").dt.strftime(TARIH_FORMAT)
    df = df.dropna(subset=["Tarih"])
    df = df.fillna("").astype(str)
    worksheet.clear()
    worksheet.update([df.columns.values.tolist()] + df.values.tolist())
    df_oku.clear()


def gemo_sablonunu_ayikla(raw: pd.DataFrame) -> pd.DataFrame:
    """GEMO Gıda tarzı haftalık ızgara şablonunu (gün sütunları, altında
    çorba/ana yemek/yardımcı/içecek satırları) uzun formata çevirir."""
    satir_sayisi, kolon_sayisi = raw.shape

    # Bir satırda en az 2 tarih hücresi varsa o satır "tarih satırı"dır
    tarih_satirlari = []
    for r in range(satir_sayisi):
        adet = sum(isinstance(v, (pd.Timestamp, datetime)) for v in raw.iloc[r])
        if adet >= 2:
            tarih_satirlari.append(r)

    etiketler = ["Çorba", "Ana Yemek", "Yardımcı", "İçecek/Tatlı", "Ekstra 1", "Ekstra 2", "Ekstra 3"]
    kayitlar = []

    for i, r in enumerate(tarih_satirlari):
        sonraki_r = tarih_satirlari[i + 1] if i + 1 < len(tarih_satirlari) else satir_sayisi
        for c in range(kolon_sayisi):
            deger = raw.iat[r, c]
            if not isinstance(deger, (pd.Timestamp, datetime)):
                continue
            tarih = pd.Timestamp(deger).strftime(TARIH_FORMAT)
            kalemler = []
            for rr in range(r + 1, sonraki_r):
                hucre = raw.iat[rr, c]
                if pd.notna(hucre) and str(hucre).strip() and str(hucre).strip().lower() != "kcal":
                    kalemler.append(str(hucre).strip())
            if kalemler:
                kayit = {"Tarih": tarih}
                for idx, kalem in enumerate(kalemler):
                    etiket = etiketler[idx] if idx < len(etiketler) else f"Kalem {idx + 1}"
                    kayit[etiket] = kalem
                kayitlar.append(kayit)

    return pd.DataFrame(kayitlar)


def bugunun_menusu(df_menu: pd.DataFrame):
    if df_menu.empty or "Tarih" not in df_menu.columns:
        return None
    bugun = simdi_tr().strftime(TARIH_FORMAT)
    eslesen = df_menu[df_menu["Tarih"] == bugun]
    if eslesen.empty:
        return None
    return eslesen.iloc[0]


GUN_ADLARI_TR = [
    "Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar",
]


def gun_adi(tarih_str: str) -> str:
    """'DD.MM.YYYY' formatındaki bir tarih metninden Türkçe gün adını döndürür.
    Tarih ayrıştırılamazsa boş string döner."""
    try:
        tarih = datetime.strptime(tarih_str, TARIH_FORMAT)
        return GUN_ADLARI_TR[tarih.weekday()]
    except (ValueError, TypeError):
        return ""


def puan_kategorisi(puan) -> str:
    """1-5 arası sayısal puanı raporlarda kullanılan Kötü/Orta/Güzel
    kategorisine çevirir. 1-2: Kötü, 3: Orta, 4-5: Güzel."""
    try:
        p = int(puan)
    except (ValueError, TypeError):
        return "Orta"
    if p <= 2:
        return "Kötü"
    if p == 3:
        return "Orta"
    return "Güzel"


# Bağlantıyı kur
baglanti_hatasi = None
spreadsheet = None
try:
    spreadsheet = sheets_baglantisi_al()
except Exception as e:
    baglanti_hatasi = str(e)

if baglanti_hatasi:
    st.error(
        "⚠️ Google Sheets bağlantısı kurulamadı. Lütfen `secrets.toml` "
        "içindeki `gcp_service_account` ve `sheet_id` ayarlarını kontrol edin."
    )
    with st.expander("Hata detayı"):
        st.code(baglanti_hatasi)
    st.stop()

try:
    ws_menu = worksheet_al(spreadsheet, WS_MENU, ["Tarih"])
    ws_on = worksheet_al(spreadsheet, WS_ON, ["Tarih", "Degerlendirme", "KayitZamani"])
    ws_son = worksheet_al(spreadsheet, WS_SON, ["Tarih", "Kalem", "UrunAdi", "Degerlendirme", "Aciklama", "KayitZamani"])
    ws_oneri = worksheet_al(spreadsheet, WS_ONERI, ["Tarih", "Oneri", "KayitZamani"])
except Exception as e:
    st.error(
        "⚠️ Google Sheets ile iletişimde bir sorun oluştu. Bu genellikle Google'ın "
        "kısa süreli istek sınırına (rate limit) takılmaktan kaynaklanır — birkaç "
        "saniye bekleyip sayfayı yenilemeyi dene. Sorun devam ederse şifreni ve "
        "sheet paylaşım ayarlarını kontrol et."
    )
    with st.expander("Hata detayı"):
        st.code(str(e))
    st.stop()

# =========================================================
# ARAYÜZ
# =========================================================
# streamlit_cookies_controller yerine streamlit-cookies-manager kullanılıyor.
# Bu kütüphane, çerezlerin tarayıcıya gerçekten yazılıp yazılmadığını
# ready()/save() ile açıkça kontrol etmemizi sağlıyor — bu, önceki
# kütüphanenin sessizce yazmayı kaybettiği (ve kısıtlamanın hiç
# çalışmamış gibi görünmesine yol açan) zamanlama sorununu ortadan
# kaldırıyor.
#
# secrets.toml içine (opsiyonel ama önerilir) şu satırı ekleyebilirsin:
# cookies_sifre = "uzun-rastgele-bir-metin"
cookies = EncryptedCookieManager(
    prefix="pameks_yemek_anketi/",
    password=st.secrets.get(
        "cookies_sifre", "pameks-yemek-anketi-varsayilan-sifre-degistir"
    ),
)

# Bileşen tarayıcıdaki çerezleri Python tarafına aktarana kadar (bir
# defalık, en fazla birkaç rerun süren bir yükleme) bekliyoruz.
if not cookies.ready():
    st.info("Yükleniyor, lütfen bekleyin...")
    st.stop()


def cihazda_oylanan_gunleri_oku(cookie_adi: str) -> set:
    """İlgili çerezden, bu cihazda daha önce oylanmış/değerlendirilmiş
    günlerin tarih kümesini okur."""
    ham = cookies.get(cookie_adi)
    if not ham:
        return set()
    return {t for t in str(ham).split(",") if t}


def cihaza_oylanan_gun_ekle(cookie_adi: str, mevcut_kume: set, yeni_tarih: str):
    """Kümeye yeni tarihi ekleyip aynı çerez adına geri yazar. save()
    çağrısı, bir sonraki rerun'u beklemeden yazmayı hemen tarayıcıya
    gönderir — böylece st.rerun() ile yarış durumu oluşmaz."""
    guncel = mevcut_kume | {yeni_tarih}
    cookies[cookie_adi] = ",".join(sorted(guncel))
    cookies.save()


def cihaz_bugun_kullandi_mi(cookie_adi: str, gun: str) -> bool:
    """Bu cihaz, belirtilen gün için bu formu daha önce kullandı mı?
    Önce oturum durumuna (session_state) bakar — çerez tarayıcıya yazılırken
    kısa bir gecikme olabildiği için, aynı oturumda hemen art arda tekrar
    gönderimi engellemenin tek güvenilir yolu budur. Çerez ise farklı bir
    oturumda (tarayıcı kapatılıp açıldığında) devreye girer."""
    oturum_anahtari = f"_cihaz_kullanildi_{cookie_adi}_{gun}"
    if st.session_state.get(oturum_anahtari, False):
        return True
    return gun in cihazda_oylanan_gunleri_oku(cookie_adi)


def cihazi_bugun_icin_isaretle(cookie_adi: str, mevcut_kume: set, gun: str):
    """Hem oturum durumunu hem de çerezi günceller."""
    oturum_anahtari = f"_cihaz_kullanildi_{cookie_adi}_{gun}"
    st.session_state[oturum_anahtari] = True
    cihaza_oylanan_gun_ekle(cookie_adi, mevcut_kume, gun)


st.markdown(
    """
    <style>
    /* Sekmeler dar ekranlarda (telefon) kaydırma yerine alt satıra kaysın */
    div[data-baseweb="tab-list"] {
        flex-wrap: wrap !important;
        gap: 4px;
        row-gap: 6px;
    }
    button[data-baseweb="tab"] {
        white-space: normal !important;
        font-size: 13px;
        padding: 8px 10px;
        flex: 1 1 auto;
        min-width: 0;
    }
    @media (max-width: 480px) {
        button[data-baseweb="tab"] {
            font-size: 12px;
            padding: 6px 8px;
        }
        button[data-baseweb="tab"] p {
            font-size: 12px;
        }
    }
    .gun-adi-etiket {
        font-size: 17px;
        font-weight: 700;
        color: #444444;
        margin-top: -8px;
        margin-bottom: 4px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🍽️ Pameks Giyim San. A.Ş. Yemek ve Menü Görüş Bildirim Sistemi")
st.markdown("Görüşleriniz menüleri birlikte iyileştirmemiz için bizim çok değerli.")
st.divider()

df_menu = df_oku(ws_menu, WS_MENU)
gunun_menusu = bugunun_menusu(df_menu)
bugun_str = simdi_tr().strftime(TARIH_FORMAT)

secim = st.radio(
    "Gitmek istediğin bölümü seç",
    ["✅ Günün Menüsünü Değerlendirme", "📅 Aylık Menü Oylama", "💡 Dilek-Şikâyet-Öneri"],
    key="ana_bolum_secimi",
)
st.divider()

# --- TAB 1: Aylık Menü & Ön Oylama ---
if secim == "📅 Aylık Menü Oylama":
    if df_menu.empty:
        st.info("Sisteme henüz bir menü yüklenmemiş.")
    else:
        # Tarihe göre kronolojik sırala
        df_menu_sirali = df_menu.copy()
        df_menu_sirali["_siralama"] = pd.to_datetime(
            df_menu_sirali["Tarih"], format=TARIH_FORMAT, errors="coerce"
        )
        df_menu_sirali = df_menu_sirali.sort_values("_siralama").reset_index(drop=True)

        # Geçmişte kalan günleri (bugünden önceki) oylamadan çıkar
        bugun_tarihi = pd.to_datetime(bugun_str, format=TARIH_FORMAT)
        df_menu_sirali = df_menu_sirali[df_menu_sirali["_siralama"] >= bugun_tarihi].reset_index(drop=True)

        # Bu cihazda daha önce (önceki bir oturumda dahil) oylanmış günleri çıkar
        cihazda_oylanan = cihazda_oylanan_gunleri_oku(COOKIE_ON_OYLAMA)
        # Çerez yazma isteği tarayıcıya ulaşmadan hemen aynı oturumda tekrar
        # oy kullanılmasını da engellemek için, bu oturumda oylanan günleri
        # ayrıca session_state'te tutuyoruz (çereze ek güvence).
        if "_on_oylama_bu_oturum" not in st.session_state:
            st.session_state._on_oylama_bu_oturum = set()
        oylanmis_gunler = cihazda_oylanan | st.session_state._on_oylama_bu_oturum

        df_menu_sirali_tum = df_menu_sirali.copy()  # tablo görünümü için orijinali sakla
        df_menu_sirali = df_menu_sirali[~df_menu_sirali["Tarih"].isin(oylanmis_gunler)].reset_index(drop=True)

        gun_listesi = df_menu_sirali["Tarih"].tolist()
        toplam_gun_sayisi = len(df_menu_sirali_tum)
        kalan_gun_sayisi = len(gun_listesi)
        oylanan_gun_sayisi = toplam_gun_sayisi - kalan_gun_sayisi

        if not gun_listesi:
            st.success("🎉 Bu ayki tüm günleri oyladın, teşekkürler!")
        else:
            if toplam_gun_sayisi:
                st.progress(
                    oylanan_gun_sayisi / toplam_gun_sayisi,
                    text=f"Gün {oylanan_gun_sayisi + 1} / {toplam_gun_sayisi}",
                )

            # Liste zaten oylanmış günleri dışarıda bıraktığı için gösterilecek
            # gün her zaman listede kalan İLK gündür. Ayrıca bir indeks sayacı
            # tutup elle ilerletmiyoruz — bu, listenin kendisinin küçülmesiyle
            # birlikte günlerin ikişer ikişer atlanmasına neden oluyordu.
            secili_satir = df_menu_sirali.iloc[0]
            secili_tarih = secili_satir["Tarih"]

            st.markdown(f"## {secili_tarih}")
            st.markdown(f'<div class="gun-adi-etiket">📆 {gun_adi(secili_tarih)}</div>', unsafe_allow_html=True)
            for kolon in df_menu.columns:
                if kolon != "Tarih" and str(secili_satir[kolon]).strip():
                    st.markdown(f"**{kolon}:** {secili_satir[kolon]}")

            st.write("")
            c1, c2 = st.columns(2)

            def _oy_kaydet_ve_ilerle(tarih, deger):
                try:
                    satir_ekle(ws_on, [tarih, deger, simdi_tr().strftime("%d.%m.%Y %H:%M:%S")])
                    st.session_state._on_oylama_bu_oturum.add(tarih)
                    cihaza_oylanan_gun_ekle(COOKIE_ON_OYLAMA, cihazda_oylanan, tarih)
                    st.rerun()
                except Exception as e:
                    st.error(f"Kayıt sırasında hata oluştu: {e}")

            if c1.button("👍 İyi", use_container_width=True, key=f"iyi_{secili_tarih}"):
                _oy_kaydet_ve_ilerle(secili_tarih, "İyi")
            if c2.button("👎 Kötü", use_container_width=True, key=f"kotu_{secili_tarih}"):
                _oy_kaydet_ve_ilerle(secili_tarih, "Kötü")

        st.divider()
        with st.expander("📅 Aylık Menü Tablosunu Gör"):
            st.dataframe(
                df_menu_sirali_tum.drop(columns=["_siralama"]),
                use_container_width=True,
                hide_index=True,
            )


# --- TAB 2: Yemek Sonrası Değerlendirme ---
if secim == "✅ Günün Menüsünü Değerlendirme":
    st.subheader(f"Bugünkü Yemekler Nasıldı? — {bugun_str}")
    st.markdown(f'<div class="gun-adi-etiket">📆 {gun_adi(bugun_str)}</div>', unsafe_allow_html=True)

    cihazda_degerlendirilen = cihazda_oylanan_gunleri_oku(COOKIE_SON_DEGERLENDIRME)

    if gunun_menusu is None:
        st.info("Bugün için sisteme henüz bir menü girilmemiş.")
    elif cihaz_bugun_kullandi_mi(COOKIE_SON_DEGERLENDIRME, bugun_str):
        st.success("✅ Bu cihazdan bugün için değerlendirmeni zaten gönderdin. Teşekkürler!")
    else:
        urun_kolonlari = [
            k for k in df_menu.columns
            if k != "Tarih" and str(gunun_menusu[k]).strip()
        ]

        st.markdown("Her ürünü ayrı ayrı **1 (çok kötü) - 5 (çok iyi)** arasında puanla. **1** veya **2** verirsen nedenini bize açıklar mısın? Bu bizim sorunu düzeltmemize yardımcı olacaktır.")
        st.divider()

        degerlendirmeler = {}
        aciklamalar = {}

        for kolon in urun_kolonlari:
            urun_adi = gunun_menusu[kolon]
            st.markdown(f"**{kolon}: {urun_adi}**")
            secim = st.radio(
                f"{kolon} nasıldı?",
                [1, 2, 3, 4, 5],
                index=3,
                horizontal=True,
                key=f"degerlendirme_{bugun_str}_{kolon}",
                label_visibility="collapsed",
            )
            degerlendirmeler[kolon] = secim

            if secim in (1, 2):
                aciklama = st.text_input(
                    f"{kolon} için neden beğenmedin?",
                    key=f"aciklama_{bugun_str}_{kolon}",
                    placeholder="Örn: Fazla tuzluydu, soğuktu vb.",
                )
                aciklamalar[kolon] = aciklama

            st.divider()

        if st.button("Değerlendirmeyi Gönder", key="son_degerlendirme_gonder"):
            eksik_aciklama_var = any(
                degerlendirmeler[k] in (1, 2) and not aciklamalar.get(k, "").strip()
                for k in urun_kolonlari
            )
            if eksik_aciklama_var:
                st.warning("1 veya 2 puan verdiğin ürün(ler) için lütfen kısa bir açıklama yaz.")
            else:
                try:
                    zaman = simdi_tr().strftime("%d.%m.%Y %H:%M:%S")
                    for kolon in urun_kolonlari:
                        satir_ekle(
                            ws_son,
                            [
                                bugun_str,
                                kolon,
                                gunun_menusu[kolon],
                                degerlendirmeler[kolon],
                                aciklamalar.get(kolon, ""),
                                zaman,
                            ],
                        )
                    cihazi_bugun_icin_isaretle(COOKIE_SON_DEGERLENDIRME, cihazda_degerlendirilen, bugun_str)
                    st.success("Teşekkürler! Değerlendirmeniz yönetim ekibine iletildi.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Kayıt sırasında hata oluştu: {e}")

# --- TAB 3: Menü Önerisi ---
if secim == "💡 Dilek-Şikâyet-Öneri":
    st.subheader("Menü Önerin Var mı? Veya Herhangi bir konuda şikâyetin var mı?")

    cihazda_oneri_yapilan = cihazda_oylanan_gunleri_oku(COOKIE_ONERI)

    if cihaz_bugun_kullandi_mi(COOKIE_ONERI, bugun_str):
        st.success("✅ Bu cihazdan bugün için zaten bir öneri/şikâyet gönderdin. Teşekkürler!")
    else:
        with st.form("oneri_form"):
            oneri_metni = st.text_area(
                "Önerini yaz:",
                placeholder="Örn: Ayda bir kere mercimek köftesi olabilir. Veya yemekte herhangi bir cisim çıktı gibi...",
            )
            oneri_submit = st.form_submit_button("Öneriyi Gönder")
            if oneri_submit:
                if oneri_metni.strip():
                    try:
                        satir_ekle(
                            ws_oneri,
                            [bugun_str, oneri_metni, simdi_tr().strftime("%d.%m.%Y %H:%M:%S")],
                        )
                        cihazi_bugun_icin_isaretle(COOKIE_ONERI, cihazda_oneri_yapilan, bugun_str)
                        st.success("Teşekkürler! Önerin kaydedildi.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Kayıt sırasında hata oluştu: {e}")
                else:
                    st.warning("Lütfen bir öneri yazın.")

# =========================================================
# YÖNETİCİ PANELİ
# =========================================================
yonetici_erisimi = "admin" in st.query_params

if yonetici_erisimi:
    with st.expander("🔒 Yönetici Paneli (Yetkili Girişi)", expanded=False):

        if "admin_giris_yapildi" not in st.session_state:
            st.session_state.admin_giris_yapildi = False

        if not st.session_state.admin_giris_yapildi:
            girilen_sifre = st.text_input("Yönetici Şifresi", type="password", key="admin_sifre_girisi")
            if st.button("Giriş Yap"):
                if girilen_sifre == st.secrets.get("admin_sifre", None):
                    st.session_state.admin_giris_yapildi = True
                    st.rerun()
                else:
                    st.error("Şifre hatalı.")
            st.stop()

        cikis_col = st.columns([4, 1])[1]
        with cikis_col:
            if st.button("Çıkış Yap"):
                st.session_state.admin_giris_yapildi = False
                st.rerun()

        admin_tab1, admin_tab2 = st.tabs(["📤 Aylık Menü Yükle", "📊 Raporlar"])

        # --- Menü Yükleme ---
        with admin_tab1:
            st.markdown(
                "GEMO Gıda'nın haftalık ızgara formatındaki Excel dosyasını olduğu gibi "
                "yükleyebilirsin (gün sütunları, altında çorba/ana yemek/yardımcı/içecek "
                "satırları). Ayrıca düz bir **Tarih** sütunlu CSV de kabul edilir. "
                "Yüklediğinde mevcut menünün **tamamen üzerine yazılır**."
            )
            yuklenen_dosya = st.file_uploader(
                "Aylık menü dosyasını seç", type=["csv", "xlsx", "xls"]
            )

            yeni_menu_df = None

            if yuklenen_dosya is not None:
                try:
                    if yuklenen_dosya.name.endswith(".csv"):
                        yeni_menu_df = pd.read_csv(yuklenen_dosya)
                        if "Tarih" not in yeni_menu_df.columns:
                            st.error("Dosyada 'Tarih' adında bir sütun bulunamadı.")
                            yeni_menu_df = None
                    else:
                        excel_dosyasi = pd.ExcelFile(yuklenen_dosya)
                        sayfa_secimi = st.selectbox(
                            "Hangi sayfayı içe aktarmak istiyorsun?",
                            excel_dosyasi.sheet_names,
                        )
                        ham_veri = excel_dosyasi.parse(sayfa_secimi, header=None)

                        # Önce GEMO ızgara şablonu olarak dene
                        ayiklanan = gemo_sablonunu_ayikla(ham_veri)

                        if not ayiklanan.empty:
                            yeni_menu_df = ayiklanan
                            st.success(
                                f"Şablon otomatik tanındı — {len(ayiklanan)} günlük menü tespit edildi."
                            )
                        else:
                            # Şablon tanınmazsa düz tablo (ilk satır başlık) olarak dene
                            duz_deneme = excel_dosyasi.parse(sayfa_secimi)
                            if "Tarih" in duz_deneme.columns:
                                yeni_menu_df = duz_deneme
                            else:
                                st.error(
                                    "Bu sayfadan menü tespit edilemedi. Format farklı olabilir "
                                    "veya sayfa boş/kalori listesi gibi başka bir içerik olabilir."
                                )

                    if yeni_menu_df is not None:
                        st.markdown("**Önizleme:**")
                        st.dataframe(yeni_menu_df, use_container_width=True)

                        if st.button("📤 Menüyü Sisteme Yükle (Üzerine Yazar)"):
                            menuyu_yukle(ws_menu, yeni_menu_df)
                            st.success("Menü başarıyla yüklendi!")
                            st.rerun()
                except Exception as e:
                    st.error(f"Dosya okunurken hata oluştu: {e}")

            if not df_menu.empty:
                st.divider()
                st.markdown("**Sistemde şu anda kayıtlı menü:**")
                st.dataframe(df_menu, use_container_width=True)

        # --- Raporlar ---
        with admin_tab2:
            def _basliklari_onar(worksheet, beklenen_kolonlar):
                """Sayfanın 1. satırındaki başlık etiketlerini kodun beklediği
                kanonik isimlerle değiştirir. Veriler sütunlara pozisyonel olarak
                (A, B, C... sırasına göre) yazıldığı için bu işlem mevcut veriyi
                BOZMAZ; sadece yanlış/eski başlık metnini düzeltir. Beklenenden
                fazla (kullanılmayan) başlık sütunu varsa temizler."""
                mevcut_degerler = worksheet.get_all_values()
                mevcut_sutun_sayisi = len(mevcut_degerler[0]) if mevcut_degerler else 0
                yeni_baslik_satiri = list(beklenen_kolonlar)
                if mevcut_sutun_sayisi > len(beklenen_kolonlar):
                    yeni_baslik_satiri += [""] * (mevcut_sutun_sayisi - len(beklenen_kolonlar))
                worksheet.update("A1", [yeni_baslik_satiri])

            with st.expander("🛠️ Sayfa Başlıklarını Onar", expanded=False):
                st.caption(
                    "Eğer yukarıda 'beklenen sütun bulunamadı' uyarısı görüyorsan, "
                    "bu genellikle sayfanın 1. satırındaki başlık etiketlerinin eski/yanlış "
                    "olmasından kaynaklanır. Aşağıdaki buton **veriyi bozmadan**, sadece "
                    "başlık satırını kodun beklediği doğru isimlerle değiştirir."
                )
                if st.button("🛠️ Tüm Sayfaların Başlıklarını Onar"):
                    try:
                        _basliklari_onar(ws_on, ["Tarih", "Degerlendirme", "KayitZamani"])
                        _basliklari_onar(
                            ws_son,
                            ["Tarih", "Kalem", "UrunAdi", "Degerlendirme", "Aciklama", "KayitZamani"],
                        )
                        _basliklari_onar(ws_oneri, ["Tarih", "Oneri", "KayitZamani"])
                        df_oku.clear()
                        st.success("Başlıklar onarıldı! Sayfa yenileniyor...")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Başlıklar onarılırken hata oluştu: {e}")

            def _sutun_normalize(s: str) -> str:
                """Türkçe karakterleri, boşlukları ve büyük/küçük harf farkını
                yok sayarak karşılaştırma yapabilmek için bir sütun adını sadeleştirir.
                Örn: 'Değerlendirme', 'DEĞERLENDİRME ', 'değerlendirme' -> 'degerlendirme'"""
                s = str(s).strip().lower()
                donusum = {
                    "ı": "i", "i̇": "i", "İ": "i",
                    "ğ": "g", "ü": "u", "ş": "s", "ö": "o", "ç": "c",
                }
                for eski, yeni in donusum.items():
                    s = s.replace(eski, yeni)
                return re.sub(r"[^a-z0-9]", "", s)

            def beklenen_sutunlari_garanti_et(df, beklenen_kolonlar, sayfa_adi):
                """Sheet'teki gerçek başlıklar kodun beklediğiyle birebir uyuşmuyorsa
                (fazladan/eksik boşluk, Türkçe karakter farkı, farklı büyük/küçük harf vb.)
                önce normalize ederek eşleştirmeyi dener ve gerçek sütunu beklenen isme
                yeniden adlandırır. Yine de bulunamayan sütunlar için KeyError yerine
                kullanıcıya net bir uyarı gösterir ve boş sütun ekleyerek çökmeyi önler."""
                if df.empty:
                    return df

                gercek_norm_harita = {_sutun_normalize(c): c for c in df.columns}
                yeniden_adlandir = {}
                hala_eksik = []
                for beklenen in beklenen_kolonlar:
                    if beklenen in df.columns:
                        continue
                    norm = _sutun_normalize(beklenen)
                    if norm in gercek_norm_harita:
                        yeniden_adlandir[gercek_norm_harita[norm]] = beklenen
                    else:
                        hala_eksik.append(beklenen)

                if yeniden_adlandir:
                    df = df.rename(columns=yeniden_adlandir)

                if hala_eksik:
                    st.warning(
                        f"⚠️ **{sayfa_adi}** sayfasında beklenen sütun(lar) bulunamadı: "
                        f"{', '.join(hala_eksik)}. Google Sheet'teki başlık satırını kontrol et "
                        f"(farklı bir isim kullanılmış ya da hücre boş olabilir)."
                    )
                    with st.expander(f"'{sayfa_adi}' sayfasında bulunan gerçek sütunlar"):
                        st.write(list(df.columns))
                    for k in hala_eksik:
                        df[k] = ""
                return df

            df_on = beklenen_sutunlari_garanti_et(
                df_oku(ws_on, WS_ON), ["Tarih", "Degerlendirme", "KayitZamani"], WS_ON
            )
            df_son = beklenen_sutunlari_garanti_et(
                df_oku(ws_son, WS_SON),
                ["Tarih", "Kalem", "UrunAdi", "Degerlendirme", "Aciklama", "KayitZamani"],
                WS_SON,
            )
            df_oneri = beklenen_sutunlari_garanti_et(
                df_oku(ws_oneri, WS_ONERI), ["Tarih", "Oneri", "KayitZamani"], WS_ONERI
            )

            if df_on.empty and df_son.empty:
                st.info("Henüz hiç oylama verisi yok.")
            else:
                if not df_son.empty:
                    df_son = df_son.copy()
                    df_son["_puan_sayi"] = pd.to_numeric(df_son["Degerlendirme"], errors="coerce")
                    df_son["_kategori"] = df_son["Degerlendirme"].apply(puan_kategorisi)

                st.markdown("### 📊 Özet")
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Ön Oylama Sayısı", len(df_on))
                iyi_orani = (
                    (df_on["Degerlendirme"] == "İyi").mean() * 100 if not df_on.empty else 0
                )
                m2.metric("Ön Oylama İyi Oranı", f"%{iyi_orani:.0f}" if not df_on.empty else "—")
                m3.metric(
                    "Ortalama Puan (1-5)",
                    f"{df_son['_puan_sayi'].mean():.2f}" if not df_son.empty else "—",
                )
                kotu_orani = (
                    (df_son["_kategori"] == "Kötü").mean() * 100 if not df_son.empty else 0
                )
                m4.metric("Kötü Oranı (1-2 Puan)", f"%{kotu_orani:.0f}" if not df_son.empty else "—")

                st.divider()
                st.markdown("### 📈 Ön Oylama Trend (Zaman İçinde İyi Oranı)")
                if not df_on.empty:
                    on_ort = df_on.groupby("Tarih")["Degerlendirme"].apply(
                        lambda x: (x == "İyi").mean() * 100
                    )
                    on_ort.index = pd.to_datetime(on_ort.index, format=TARIH_FORMAT, errors="coerce")
                    st.line_chart(on_ort.sort_index())
                else:
                    st.info("Henüz ön oylama verisi yok.")

                st.divider()
                st.markdown("### 🍽️ Ürün Bazında Dağılım (Yemek Sonrası, 1-2: Kötü, 3: Orta, 4-5: Güzel)")
                if df_son.empty:
                    st.info("Henüz ürün değerlendirmesi yok.")
                else:
                    pivot = df_son.pivot_table(
                        index="Kalem", columns="_kategori", values="Tarih",
                        aggfunc="count", fill_value=0,
                    )
                    for sutun in ["Güzel", "Orta", "Kötü"]:
                        if sutun not in pivot.columns:
                            pivot[sutun] = 0
                    pivot = pivot[["Güzel", "Orta", "Kötü"]]
                    st.bar_chart(pivot)

                    st.divider()
                    st.markdown("### ⚠️ Kötü Değerlendirmeler ve Açıklamalar (1-2 Puan)")
                    kotuler = df_son[df_son["_kategori"] == "Kötü"]
                    if kotuler.empty:
                        st.success("1-2 puan verilen ürün yok.")
                    else:
                        st.dataframe(
                            kotuler[["Tarih", "Kalem", "UrunAdi", "Degerlendirme", "Aciklama"]]
                            .sort_values("Tarih", ascending=False),
                            use_container_width=True,
                            hide_index=True,
                        )
                        csv_kotu = kotuler[["Tarih", "Kalem", "UrunAdi", "Degerlendirme", "Aciklama"]].to_csv(
                            index=False
                        ).encode("utf-8-sig")
                        st.download_button(
                            "⬇️ Kötü Değerlendirmeleri CSV Olarak İndir",
                            data=csv_kotu,
                            file_name=f"kotu_degerlendirmeler_{simdi_tr().strftime('%d_%m_%Y')}.csv",
                            mime="text/csv",
                        )

                st.divider()
                st.markdown("### 💡 Menü Önerileri")
                if df_oneri.empty:
                    st.info("Henüz öneri yok.")
                else:
                    st.dataframe(
                        df_oneri[["Tarih", "Oneri"]].sort_values("Tarih", ascending=False),
                        use_container_width=True,
                    )

                st.divider()
                st.markdown("### 📊 Raporu Excel Olarak İndir")

                rapor_araligi_turu = st.radio(
                    "Rapor aralığı",
                    ["Günlük", "Haftalık", "Aylık", "Özel Tarih Aralığı"],
                    horizontal=True,
                    key="excel_rapor_araligi_turu",
                )

                bugun_tarihi = simdi_tr().date()

                if rapor_araligi_turu == "Günlük":
                    secilen_gun = st.date_input(
                        "Hangi günün raporu indirilsin?",
                        value=bugun_tarihi,
                        format="DD.MM.YYYY",
                        key="excel_rapor_gun_secimi",
                    )
                    rapor_baslangic = rapor_bitis = secilen_gun

                elif rapor_araligi_turu == "Haftalık":
                    referans_gun = st.date_input(
                        "Bu haftanın herhangi bir günü",
                        value=bugun_tarihi,
                        format="DD.MM.YYYY",
                        key="excel_rapor_hafta_secimi",
                    )
                    rapor_baslangic = referans_gun - pd.Timedelta(days=referans_gun.weekday())
                    rapor_bitis = rapor_baslangic + pd.Timedelta(days=6)
                    st.caption(f"Seçilen hafta: {rapor_baslangic.strftime('%d.%m.%Y')} – {rapor_bitis.strftime('%d.%m.%Y')} (Pazartesi–Pazar)")

                elif rapor_araligi_turu == "Aylık":
                    referans_gun = st.date_input(
                        "Bu ayın herhangi bir günü",
                        value=bugun_tarihi,
                        format="DD.MM.YYYY",
                        key="excel_rapor_ay_secimi",
                    )
                    rapor_baslangic = referans_gun.replace(day=1)
                    sonraki_ay = (rapor_baslangic + pd.Timedelta(days=32)).replace(day=1)
                    rapor_bitis = sonraki_ay - pd.Timedelta(days=1)
                    st.caption(f"Seçilen ay: {rapor_baslangic.strftime('%d.%m.%Y')} – {rapor_bitis.strftime('%d.%m.%Y')}")

                else:  # Özel Tarih Aralığı
                    aralik = st.date_input(
                        "Başlangıç ve bitiş tarihini seç",
                        value=(bugun_tarihi.replace(day=1), bugun_tarihi),
                        format="DD.MM.YYYY",
                        key="excel_rapor_ozel_aralik",
                    )
                    if isinstance(aralik, tuple) and len(aralik) == 2:
                        rapor_baslangic, rapor_bitis = aralik
                    else:
                        rapor_baslangic = rapor_bitis = aralik

                try:
                    excel_verisi = ozet_excel_raporu(
                        rapor_baslangic, rapor_bitis, df_on, df_son, df_oneri, df_menu
                    )
                    st.download_button(
                        "⬇️ Excel Raporunu İndir",
                        data=excel_verisi,
                        file_name=(
                            f"yemek_raporu_{rapor_baslangic.strftime('%d_%m_%Y')}"
                            f"_{rapor_bitis.strftime('%d_%m_%Y')}.xlsx"
                        ),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                except Exception as e:
                    st.error(f"Excel raporu oluşturulurken hata oluştu: {e}")

                st.divider()
                if st.button("🔄 Verileri Yenile"):
                    st.rerun()
