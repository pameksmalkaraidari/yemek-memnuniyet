import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from datetime import datetime
from zoneinfo import ZoneInfo
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

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


def simdi_tr() -> datetime:
    """Sunucu hangi saat diliminde çalışırsa çalışsın, her zaman İstanbul
    yerel saatini döndürür. 'Bugün' hesaplamalarında tutarlılık için tüm
    tarih/saat işlemleri bu fonksiyon üzerinden yapılmalı."""
    return datetime.now(TR_TZ)


def mail_gonder(konu: str, govde_html: str):
    """secrets.toml içindeki [email] ayarlarını kullanarak HTML gövdeli
    bir e-posta gönderir. Ayarlar eksikse veya gönderim başarısız olursa
    açıklayıcı bir hata fırlatır (çağıran taraf yakalayıp gösterir)."""
    email_ayarlari = st.secrets.get("email", None)
    if not email_ayarlari:
        raise RuntimeError(
            "secrets.toml içinde '[email]' bölümü bulunamadı. Lütfen gonderen_email, "
            "gonderen_sifre, alici_epostalar, smtp_sunucu ve smtp_port ayarlarını ekle."
        )

    gonderen = email_ayarlari["gonderen_email"]
    sifre = email_ayarlari["gonderen_sifre"]
    aliciler = email_ayarlari["alici_epostalar"]
    if isinstance(aliciler, str):
        aliciler = [a.strip() for a in aliciler.split(",") if a.strip()]
    smtp_sunucu = email_ayarlari.get("smtp_sunucu", "smtp.gmail.com")
    smtp_port = int(email_ayarlari.get("smtp_port", 465))

    mesaj = MIMEMultipart("alternative")
    mesaj["Subject"] = konu
    mesaj["From"] = gonderen
    mesaj["To"] = ", ".join(aliciler)
    mesaj.attach(MIMEText(govde_html, "html", "utf-8"))

    with smtplib.SMTP_SSL(smtp_sunucu, smtp_port) as sunucu:
        sunucu.login(gonderen, sifre)
        sunucu.sendmail(gonderen, aliciler, mesaj.as_string())


def gunluk_ozet_html(tarih_str: str, df_on: pd.DataFrame, df_son: pd.DataFrame, df_oneri: pd.DataFrame) -> str:
    """Belirtilen tarih için ön oylama, ürün değerlendirmesi ve öneri
    verilerinden sade bir HTML e-posta gövdesi üretir."""
    gun = gun_adi(tarih_str)
    parcalar = [f"<h2>🍽️ Günlük Yemek Raporu — {tarih_str} {gun}</h2>"]

    # Ön oylama özeti
    on_bugun = df_on[df_on["Tarih"] == tarih_str] if not df_on.empty and "Tarih" in df_on.columns else pd.DataFrame()
    if on_bugun.empty:
        parcalar.append("<p><b>Ön Oylama:</b> Bugün için ön oylama verisi yok.</p>")
    else:
        iyi = (on_bugun["Degerlendirme"] == "İyi").sum()
        kotu = (on_bugun["Degerlendirme"] == "Kötü").sum()
        toplam = len(on_bugun)
        oran = (iyi / toplam * 100) if toplam else 0
        parcalar.append(
            f"<p><b>Ön Oylama:</b> {toplam} oy — 👍 {iyi} İyi, 👎 {kotu} Kötü "
            f"(İyi oranı: %{oran:.0f})</p>"
        )

    # Ürün bazlı değerlendirme özeti
    son_bugun = df_son[df_son["Tarih"] == tarih_str] if not df_son.empty and "Tarih" in df_son.columns else pd.DataFrame()
    if son_bugun.empty:
        parcalar.append("<p><b>Yemek Sonrası Değerlendirme:</b> Bugün için değerlendirme verisi yok.</p>")
    else:
        pivot = son_bugun.pivot_table(
            index="Kalem", columns="Degerlendirme", values="Tarih", aggfunc="count", fill_value=0
        )
        for sutun in ["Güzel", "Orta", "Kötü"]:
            if sutun not in pivot.columns:
                pivot[sutun] = 0
        pivot = pivot[["Güzel", "Orta", "Kötü"]]
        parcalar.append("<p><b>Yemek Sonrası Değerlendirme:</b></p>")
        parcalar.append(pivot.to_html(border=1))

        kotuler = son_bugun[son_bugun["Degerlendirme"] == "Kötü"]
        if not kotuler.empty:
            parcalar.append("<p><b>⚠️ Kötü Değerlendirmeler:</b></p>")
            parcalar.append(
                kotuler[["Kalem", "UrunAdi", "Aciklama"]].to_html(index=False, border=1)
            )

    # Öneriler
    oneri_bugun = df_oneri[df_oneri["Tarih"] == tarih_str] if not df_oneri.empty and "Tarih" in df_oneri.columns else pd.DataFrame()
    if not oneri_bugun.empty:
        parcalar.append("<p><b>💡 Bugünkü Menü Önerileri:</b></p>")
        parcalar.append(oneri_bugun[["Oneri"]].to_html(index=False, border=1))

    return "<br>".join(parcalar)


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

tab_degerlendirme, tab_on_oylama, tab3 = st.tabs(
    ["✅ Günün Menüsünü Değerlendirme", "📅 Aylık Menü Oylama", "💡 Dilek-Şikâyet-Öneri"]
)

# --- TAB 1: Aylık Menü & Ön Oylama ---
with tab_on_oylama:
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

        gun_listesi = df_menu_sirali["Tarih"].tolist()

        if "on_oylama_indeks" not in st.session_state:
            st.session_state.on_oylama_indeks = 0

        indeks = st.session_state.on_oylama_indeks

        if indeks >= len(gun_listesi):
            st.success("🎉 Bu ayki tüm günleri oyladın, teşekkürler!")
            if st.button("Baştan Oylamaya Başla"):
                st.session_state.on_oylama_indeks = 0
                st.rerun()
        else:
            st.progress(indeks / len(gun_listesi), text=f"Gün {indeks + 1} / {len(gun_listesi)}")

            secili_satir = df_menu_sirali.iloc[indeks]
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
                    st.session_state.on_oylama_indeks += 1
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
                df_menu_sirali.drop(columns=["_siralama"]),
                use_container_width=True,
                hide_index=True,
            )


# --- TAB 2: Yemek Sonrası Değerlendirme ---
with tab_degerlendirme:
    st.subheader(f"Bugünkü Yemekler Nasıldı? — {bugun_str}")
    st.markdown(f'<div class="gun-adi-etiket">📆 {gun_adi(bugun_str)}</div>', unsafe_allow_html=True)

    if gunun_menusu is None:
        st.info("Bugün için sisteme henüz bir menü girilmemiş.")
    else:
        urun_kolonlari = [
            k for k in df_menu.columns
            if k != "Tarih" and str(gunun_menusu[k]).strip()
        ]

        st.markdown("Her ürünü ayrı ayrı değerlendir. **Kötü** olduğunu düşünüyorsan nedenini bize açıklar mısın? Bu bizim sorunu düzeltmemize yardımcı olacaktır.")
        st.divider()

        degerlendirmeler = {}
        aciklamalar = {}

        for kolon in urun_kolonlari:
            urun_adi = gunun_menusu[kolon]
            st.markdown(f"**{kolon}: {urun_adi}**")
            secim = st.radio(
                f"{kolon} nasıldı?",
                ["Güzel", "Orta", "Kötü"],
                index=1,
                horizontal=True,
                key=f"degerlendirme_{bugun_str}_{kolon}",
                label_visibility="collapsed",
            )
            degerlendirmeler[kolon] = secim

            if secim == "Kötü":
                aciklama = st.text_input(
                    f"{kolon} için neden beğenmedin?",
                    key=f"aciklama_{bugun_str}_{kolon}",
                    placeholder="Örn: Fazla tuzluydu, soğuktu vb.",
                )
                aciklamalar[kolon] = aciklama

            st.divider()

        if st.button("Değerlendirmeyi Gönder", key="son_degerlendirme_gonder"):
            eksik_aciklama_var = any(
                degerlendirmeler[k] == "Kötü" and not aciklamalar.get(k, "").strip()
                for k in urun_kolonlari
            )
            if eksik_aciklama_var:
                st.warning("Kötü olarak işaretlediğin ürün(ler) için lütfen kısa bir açıklama yaz.")
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
                    st.success("Teşekkürler! Değerlendirmeniz yönetim ekibine iletildi.")
                except Exception as e:
                    st.error(f"Kayıt sırasında hata oluştu: {e}")

# --- TAB 3: Menü Önerisi ---
with tab3:
    st.subheader("Menü Önerin Var mı? Veya Herhangi bir konuda şikâyetin var mı?")
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
                    st.success("Teşekkürler! Önerin kaydedildi.")
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
                st.markdown("### 📊 Özet")
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Ön Oylama Sayısı", len(df_on))
                iyi_orani = (
                    (df_on["Degerlendirme"] == "İyi").mean() * 100 if not df_on.empty else 0
                )
                m2.metric("Ön Oylama İyi Oranı", f"%{iyi_orani:.0f}" if not df_on.empty else "—")
                m3.metric("Toplam Ürün Değerlendirmesi", len(df_son))
                kotu_orani = (
                    (df_son["Degerlendirme"] == "Kötü").mean() * 100 if not df_son.empty else 0
                )
                m4.metric("Kötü Oranı", f"%{kotu_orani:.0f}" if not df_son.empty else "—")

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
                st.markdown("### 🍽️ Ürün Bazında Dağılım (Yemek Sonrası)")
                if df_son.empty:
                    st.info("Henüz ürün değerlendirmesi yok.")
                else:
                    pivot = df_son.pivot_table(
                        index="Kalem", columns="Degerlendirme", values="Tarih",
                        aggfunc="count", fill_value=0,
                    )
                    for sutun in ["Güzel", "Orta", "Kötü"]:
                        if sutun not in pivot.columns:
                            pivot[sutun] = 0
                    pivot = pivot[["Güzel", "Orta", "Kötü"]]
                    st.bar_chart(pivot)

                    st.divider()
                    st.markdown("### ⚠️ Kötü Değerlendirmeler ve Açıklamalar")
                    kotuler = df_son[df_son["Degerlendirme"] == "Kötü"]
                    if kotuler.empty:
                        st.success("Kötü olarak işaretlenen ürün yok.")
                    else:
                        st.dataframe(
                            kotuler[["Tarih", "Kalem", "UrunAdi", "Aciklama"]]
                            .sort_values("Tarih", ascending=False),
                            use_container_width=True,
                            hide_index=True,
                        )
                        csv_kotu = kotuler[["Tarih", "Kalem", "UrunAdi", "Aciklama"]].to_csv(
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
                st.markdown("### 📧 Günlük Raporu Mail Gönder")
                gonderilecek_tarih = st.date_input(
                    "Hangi günün raporu gönderilsin?",
                    value=simdi_tr().date(),
                    format="DD.MM.YYYY",
                    key="mail_tarih_secimi",
                )
                gonderilecek_tarih_str = gonderilecek_tarih.strftime(TARIH_FORMAT)

                if st.button("📧 Bu Günün Raporunu Mail Gönder"):
                    try:
                        govde = gunluk_ozet_html(gonderilecek_tarih_str, df_on, df_son, df_oneri)
                        mail_gonder(
                            f"🍽️ Günlük Yemek Raporu — {gonderilecek_tarih_str}", govde
                        )
                        st.success("Rapor e-posta ile gönderildi!")
                    except Exception as e:
                        st.error(f"Mail gönderilirken hata oluştu: {e}")

                st.divider()
                if st.button("🔄 Verileri Yenile"):
                    st.rerun()
