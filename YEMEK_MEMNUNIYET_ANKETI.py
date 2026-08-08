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
# .streamlit/secrets.toml içine (veya Streamlit Cloud > Settings > Secrets)
# aşağıdaki yapıda bir servis hesabı bilgisi ve sheet ID'si eklenmeli:
#
# sheet_id = "1AbCdEfGhIjKlMnOpQrStUvWxYz..."
#   (Sheet'in tarayıcı adres çubuğundaki uzun kod: .../d/BURASI/edit)
#
# [gcp_service_account]
# type = "service_account"
# project_id = "..."
# private_key_id = "..."
# private_key = "..."
# client_email = "...@....iam.gserviceaccount.com"
# ...
#
# Not: İlgili Google Sheet dosyasını bu servis hesabının e-postasıyla
# (client_email) "Düzenleyen" olarak paylaşmayı unutma.
#
# Sadece Google Sheets API yeterli — sheet ID ile açtığımız için
# Google Drive API'yi ayrıca etkinleştirmene gerek yok.

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
]

WORKSHEET_NAME = "Kayitlar"
KOLONLAR = ["Tarih", "Puan", "Kategori", "Yorum"]


@st.cache_resource(show_spinner=False)
def sheets_baglantisi_al():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPES
    )
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(st.secrets["sheet_id"])

    try:
        worksheet = spreadsheet.worksheet(WORKSHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(
            title=WORKSHEET_NAME, rows=1000, cols=len(KOLONLAR)
        )
        worksheet.append_row(KOLONLAR)

    return worksheet


def verileri_oku(worksheet) -> pd.DataFrame:
    kayitlar = worksheet.get_all_records()
    df = pd.DataFrame(kayitlar)
    if df.empty:
        df = pd.DataFrame(columns=KOLONLAR)
    return df


def kayit_ekle(worksheet, tarih, puan, kategori, yorum):
    worksheet.append_row(
        [str(tarih), int(puan), kategori, yorum],
        value_input_option="USER_ENTERED",
    )


# Bağlantıyı kurmayı dene; sorun olursa kullanıcıyı bilgilendir
baglanti_hatasi = None
worksheet = None
try:
    worksheet = sheets_baglantisi_al()
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

# =========================================================
# ARAYÜZ
# =========================================================
st.title("🍽️ Yemek ve Geri Bildirim Sistemi")
st.markdown(
    "Görüşleriniz menüleri birlikte iyileştirmemiz için bizim çok değerli."
)
st.divider()

tab1, tab2 = st.tabs(
    ["📝 Bugünün Yemeğini Değerlendir", "💡 Haftaya Ne Olsun? (Anket)"]
)

with tab1:
    st.subheader("Bugünkü Yemekler Nasıl Harikaydı?")
    with st.form("daily_feedback_form"):
        yemek_tarihi = st.date_input("Tarih", datetime.today())
        puan = st.slider(
            "Genel Memnuniyet Puanı",
            min_value=1,
            max_value=5,
            value=4,
            help="1: Çok Kötü, 5: Çok İyi",
        )
        kategori = st.selectbox(
            "Hangi konuda geri bildirim vermek istiyorsunuz?",
            ["Genel", "Çorba", "Ana Yemek", "Pilav/Makarna/Yan Ürün", "Porsiyon/Sıcaklık"],
        )
        yorum = st.text_area(
            "Eklemek istediğiniz öneri veya eleştiri:",
            placeholder="Örn: Tuz oranı biraz daha dengeli olabilir...",
        )

        submitted = st.form_submit_button("Geri Bildirimi Gönder")
        if submitted:
            try:
                kayit_ekle(worksheet, yemek_tarihi, puan, kategori, yorum)
                st.success(
                    "Teşekkürler! Geri bildiriminiz yönetim ekibine iletildi."
                )
            except Exception as e:
                st.error(f"Kayıt eklenirken bir hata oluştu: {e}")

with tab2:
    st.subheader("Gelecek Hafta Menüsünü Seçelim")
    st.markdown("Hangi ana yemekleri daha sık görmek istersiniz? Oy verin.")

    with st.form("poll_form"):
        secim = st.radio(
            "Gelecek hafta menüsünde favoriniz hangisi olsun?",
            [
                "Etli / Tavuklu Sebze Yemekleri",
                "Fırın / Izgara Çeşitleri",
                "Zeytinyağlı Çeşitliliği",
                "Geleneksel Ev Yemekleri (Kurufasulye vb.)",
            ],
        )
        poll_submitted = st.form_submit_button("Oyla")
        if poll_submitted:
            st.success(
                "Oyunuz kaydedilmiştir. Teşekkür ederiz!"
            )

# =========================================================
# YÖNETİCİ PANELİ (Filtreleme, Metrikler ve Grafikler)
# =========================================================
with st.expander("🔒 Yönetici Paneli (Yetkili Girişi)", expanded=False):

    df = verileri_oku(worksheet)

    if df.empty:
        st.info("Henüz herhangi bir geri bildirim gönderilmedi.")
    else:
        df["Tarih"] = pd.to_datetime(df["Tarih"])
        df["Puan"] = pd.to_numeric(df["Puan"], errors="coerce")

        st.markdown("### 🔍 Filtreler")
        col1, col2 = st.columns(2)

        with col1:
            tarih_araligi = st.date_input(
                "Tarih Aralığı",
                value=(df["Tarih"].min().date(), df["Tarih"].max().date()),
                min_value=df["Tarih"].min().date(),
                max_value=df["Tarih"].max().date(),
            )

        with col2:
            secilen_kategoriler = st.multiselect(
                "Kategori",
                options=sorted(df["Kategori"].unique()),
                default=sorted(df["Kategori"].unique()),
            )

        puan_araligi = st.slider(
            "Puan Aralığı",
            min_value=1,
            max_value=5,
            value=(1, 5),
        )

        if isinstance(tarih_araligi, tuple) and len(tarih_araligi) == 2:
            baslangic, bitis = tarih_araligi
        else:
            baslangic = bitis = tarih_araligi

        filtreli_df = df[
            (df["Tarih"].dt.date >= baslangic)
            & (df["Tarih"].dt.date <= bitis)
            & (df["Kategori"].isin(secilen_kategoriler))
            & (df["Puan"].between(puan_araligi[0], puan_araligi[1]))
        ]

        st.divider()

        if filtreli_df.empty:
            st.warning("Seçilen filtrelere uyan kayıt bulunamadı.")
        else:
            st.markdown("### 📊 Özet")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Toplam Geri Bildirim", len(filtreli_df))
            m2.metric("Ortalama Puan", f"{filtreli_df['Puan'].mean():.2f}")
            m3.metric("En Sık Kategori", filtreli_df["Kategori"].mode()[0])
            memnuniyet_orani = (filtreli_df["Puan"] >= 4).mean() * 100
            m4.metric("Memnuniyet Oranı (4-5 Puan)", f"%{memnuniyet_orani:.0f}")

            st.divider()

            st.markdown("### 📈 Grafikler")
            g1, g2 = st.columns(2)

            with g1:
                st.markdown("**Kategoriye Göre Ortalama Puan**")
                kategori_ort = (
                    filtreli_df.groupby("Kategori")["Puan"]
                    .mean()
                    .sort_values(ascending=False)
                )
                st.bar_chart(kategori_ort)

            with g2:
                st.markdown("**Puan Dağılımı**")
                puan_dagilimi = (
                    filtreli_df["Puan"].value_counts().sort_index()
                )
                st.bar_chart(puan_dagilimi)

            st.markdown("**Zaman İçinde Ortalama Puan (Günlük)**")
            gunluk_ort = (
                filtreli_df.groupby(filtreli_df["Tarih"].dt.date)["Puan"]
                .mean()
            )
            st.line_chart(gunluk_ort)

            st.divider()

            st.markdown("### 📋 Detaylı Kayıtlar")
            st.dataframe(
                filtreli_df.sort_values("Tarih", ascending=False),
                use_container_width=True,
            )

            csv_veri = filtreli_df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "⬇️ Filtrelenmiş Veriyi CSV Olarak İndir",
                data=csv_veri,
                file_name=f"yemek_geri_bildirim_{datetime.today().strftime('%d_%m_%Y')}.csv",
                mime="text/csv",
            )

            if st.button("🔄 Verileri Yenile"):
                st.rerun()
