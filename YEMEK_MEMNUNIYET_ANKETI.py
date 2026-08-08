import pandas as pd
from datetime import datetime
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


@st.cache_resource(show_spinner=False)
def sheets_baglantisi_al():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPES
    )
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(st.secrets["sheet_id"])
    return spreadsheet


def worksheet_al(spreadsheet, isim, kolonlar):
    try:
        ws = spreadsheet.worksheet(isim)
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=isim, rows=1000, cols=max(len(kolonlar), 5))
        ws.append_row(kolonlar)
    return ws


def df_oku(worksheet) -> pd.DataFrame:
    kayitlar = worksheet.get_all_records()
    return pd.DataFrame(kayitlar)


def satir_ekle(worksheet, satir: list):
    worksheet.append_row(satir, value_input_option="USER_ENTERED")


def menuyu_yukle(worksheet, df: pd.DataFrame):
    """Mevcut menünün üzerine tamamen yeni menüyü yazar."""
    df = df.copy()
    df["Tarih"] = pd.to_datetime(df["Tarih"], dayfirst=True, errors="coerce").dt.strftime(TARIH_FORMAT)
    df = df.dropna(subset=["Tarih"])
    worksheet.clear()
    worksheet.update([df.columns.values.tolist()] + df.astype(str).values.tolist())


def bugunun_menusu(df_menu: pd.DataFrame):
    if df_menu.empty or "Tarih" not in df_menu.columns:
        return None
    bugun = datetime.today().strftime(TARIH_FORMAT)
    eslesen = df_menu[df_menu["Tarih"] == bugun]
    if eslesen.empty:
        return None
    return eslesen.iloc[0]


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

ws_menu = worksheet_al(spreadsheet, WS_MENU, ["Tarih"])
ws_on = worksheet_al(spreadsheet, WS_ON, ["Tarih", "Puan", "KayitZamani"])
ws_son = worksheet_al(spreadsheet, WS_SON, ["Tarih", "Puan", "Eksiklikler", "KayitZamani"])
ws_oneri = worksheet_al(spreadsheet, WS_ONERI, ["Tarih", "Oneri", "KayitZamani"])

# =========================================================
# ARAYÜZ
# =========================================================
st.title("🍽️ Yemek ve Geri Bildirim Sistemi")
st.markdown("Görüşleriniz menüleri birlikte iyileştirmemiz için bizim çok değerli.")
st.divider()

df_menu = df_oku(ws_menu)
gunun_menusu = bugunun_menusu(df_menu)
bugun_str = datetime.today().strftime(TARIH_FORMAT)

tab1, tab2, tab3 = st.tabs(
    ["📋 Bugünün Menüsü & Ön Oylama", "✅ Yemek Sonrası Değerlendirme", "💡 Menü Önerisi"]
)

# --- TAB 1: Ön Oylama ---
with tab1:
    st.subheader(f"Bugünün Menüsü — {bugun_str}")

    if gunun_menusu is None:
        st.info("Bugün için sisteme henüz bir menü girilmemiş.")
    else:
        for kolon in df_menu.columns:
            if kolon != "Tarih" and str(gunun_menusu[kolon]).strip():
                st.markdown(f"**{kolon}:** {gunun_menusu[kolon]}")

        st.divider()
        with st.form("on_oylama_form"):
            on_puan = st.slider(
                "Bu menüyü ne kadar merakla bekliyorsun / beğendin mi?",
                min_value=1, max_value=5, value=4,
                help="1: Hiç heyecanlı değilim, 5: Çok heyecanlıyım",
            )
            on_submit = st.form_submit_button("Oyumu Gönder")
            if on_submit:
                try:
                    satir_ekle(ws_on, [bugun_str, on_puan, datetime.now().strftime("%d.%m.%Y %H:%M:%S")])
                    st.success("Teşekkürler! Oyun kaydedildi.")
                except Exception as e:
                    st.error(f"Kayıt sırasında hata oluştu: {e}")

# --- TAB 2: Yemek Sonrası Değerlendirme ---
with tab2:
    st.subheader(f"Bugünkü Yemekler Nasıldı? — {bugun_str}")

    if gunun_menusu is None:
        st.info("Bugün için sisteme henüz bir menü girilmemiş.")
    else:
        for kolon in df_menu.columns:
            if kolon != "Tarih" and str(gunun_menusu[kolon]).strip():
                st.markdown(f"**{kolon}:** {gunun_menusu[kolon]}")

        st.divider()
        with st.form("son_degerlendirme_form"):
            son_puan = st.slider(
                "Genel Memnuniyet Puanı",
                min_value=1, max_value=5, value=4,
                help="1: Çok Kötü, 5: Çok İyi",
            )
            eksiklikler = st.text_area(
                "Eksiklikler / öneriler:",
                placeholder="Örn: Tuz oranı biraz daha dengeli olabilir...",
            )
            son_submit = st.form_submit_button("Değerlendirmeyi Gönder")
            if son_submit:
                try:
                    satir_ekle(
                        ws_son,
                        [bugun_str, son_puan, eksiklikler, datetime.now().strftime("%d.%m.%Y %H:%M:%S")],
                    )
                    st.success("Teşekkürler! Değerlendirmeniz yönetim ekibine iletildi.")
                except Exception as e:
                    st.error(f"Kayıt sırasında hata oluştu: {e}")

# --- TAB 3: Menü Önerisi ---
with tab3:
    st.subheader("Menü Önerin Var mı?")
    with st.form("oneri_form"):
        oneri_metni = st.text_area(
            "Önerini yaz:",
            placeholder="Örn: Ayda bir kere mercimek köftesi olabilir...",
        )
        oneri_submit = st.form_submit_button("Öneriyi Gönder")
        if oneri_submit:
            if oneri_metni.strip():
                try:
                    satir_ekle(
                        ws_oneri,
                        [bugun_str, oneri_metni, datetime.now().strftime("%d.%m.%Y %H:%M:%S")],
                    )
                    st.success("Teşekkürler! Önerin kaydedildi.")
                except Exception as e:
                    st.error(f"Kayıt sırasında hata oluştu: {e}")
            else:
                st.warning("Lütfen bir öneri yazın.")

# =========================================================
# YÖNETİCİ PANELİ
# =========================================================
with st.expander("🔒 Yönetici Paneli (Yetkili Girişi)", expanded=False):

    admin_tab1, admin_tab2 = st.tabs(["📤 Aylık Menü Yükle", "📊 Raporlar"])

    # --- Menü Yükleme ---
    with admin_tab1:
        st.markdown(
            "Excel/CSV dosyanda bir **Tarih** sütunu ve istediğin kadar yemek "
            "sütunu (Çorba, Ana Yemek, Ek Yemek, Tatlı vb.) olmalı. "
            "Yüklediğinde mevcut menünün **tamamen üzerine yazılır**."
        )
        yuklenen_dosya = st.file_uploader(
            "Aylık menü dosyasını seç", type=["csv", "xlsx", "xls"]
        )
        if yuklenen_dosya is not None:
            try:
                if yuklenen_dosya.name.endswith(".csv"):
                    yeni_menu_df = pd.read_csv(yuklenen_dosya)
                else:
                    yeni_menu_df = pd.read_excel(yuklenen_dosya)

                if "Tarih" not in yeni_menu_df.columns:
                    st.error("Dosyada 'Tarih' adında bir sütun bulunamadı.")
                else:
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
        df_on = df_oku(ws_on)
        df_son = df_oku(ws_son)
        df_oneri = df_oku(ws_oneri)

        if df_on.empty and df_son.empty:
            st.info("Henüz hiç oylama verisi yok.")
        else:
            st.markdown("### 📊 Özet")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Ön Oylama Sayısı", len(df_on))
            m2.metric(
                "Ortalama Ön Puan",
                f"{df_on['Puan'].astype(float).mean():.2f}" if not df_on.empty else "—",
            )
            m3.metric("Son Değerlendirme Sayısı", len(df_son))
            m4.metric(
                "Ortalama Son Puan",
                f"{df_son['Puan'].astype(float).mean():.2f}" if not df_son.empty else "—",
            )

            st.divider()
            st.markdown("### 📈 Zaman İçinde Ortalama Puanlar")

            grafik_df = pd.DataFrame()
            if not df_on.empty:
                on_ort = df_on.groupby("Tarih")["Puan"].apply(lambda x: pd.to_numeric(x).mean())
                grafik_df["Ön Oylama"] = on_ort
            if not df_son.empty:
                son_ort = df_son.groupby("Tarih")["Puan"].apply(lambda x: pd.to_numeric(x).mean())
                grafik_df["Son Değerlendirme"] = son_ort

            if not grafik_df.empty:
                grafik_df.index = pd.to_datetime(grafik_df.index, format=TARIH_FORMAT, errors="coerce")
                grafik_df = grafik_df.sort_index()
                st.line_chart(grafik_df)

            st.divider()
            st.markdown("### 📝 Yemek Sonrası Yorumlar / Eksiklikler")
            if df_son.empty or df_son["Eksiklikler"].str.strip().eq("").all():
                st.info("Henüz yorum yok.")
            else:
                yorumlu = df_son[df_son["Eksiklikler"].str.strip() != ""]
                st.dataframe(
                    yorumlu[["Tarih", "Puan", "Eksiklikler"]].sort_values("Tarih", ascending=False),
                    use_container_width=True,
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
            if st.button("🔄 Verileri Yenile"):
                st.rerun()
