import streamlit as st
import pandas as pd
import sqlite3
import os
import io
import time
from datetime import datetime

st.set_page_config(page_title="Tahsilat ve Cari Takip", page_icon="📈", layout="wide")

DB_FILE = "tahsilat.db"

# ==========================================
# VERİTABANI (SQLite) İŞLEMLERİ
# ==========================================
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS takip (
                    kod TEXT PRIMARY KEY, 
                    isim TEXT, 
                    telefon TEXT, 
                    bakiye REAL, 
                    durum TEXT, 
                    ozel_durum TEXT, 
                    tarih TEXT
                )''')
    c.execute('''CREATE TABLE IF NOT EXISTS loglar (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, 
                    cari_kod TEXT, 
                    tarih_saat TEXT, 
                    not_metni TEXT
                )''')
    c.execute('''CREATE TABLE IF NOT EXISTS ana_liste (
                    kod TEXT PRIMARY KEY, 
                    isim TEXT, 
                    telefon TEXT, 
                    bakiye REAL
                )''')
    conn.commit()
    conn.close()

init_db()

def get_db_connection():
    return sqlite3.connect(DB_FILE)

# --- YARDIMCI FONKSİYONLAR ---
def bakiye_temizle(deger):
    """Excel'den gelen verinin sayı mı yoksa metin mi olduğunu anlar ve doğru çevirir."""
    if pd.isna(deger): 
        return 0.0
    # Eğer zaten matematiksel bir sayıysa doğrudan kabul et (HATA BURADAN KAYNAKLANIYORDU)
    if isinstance(deger, (int, float)): 
        return float(deger)
    
    # Metin olarak geldiyse temizle
    try:
        temiz = str(deger).replace(' TL', '').replace('₺', '').strip()
        # İçinde hem nokta hem virgül varsa (örn: 1.234,56)
        if '.' in temiz and ',' in temiz:
            temiz = temiz.replace('.', '').replace(',', '.')
        # Sadece virgül varsa (örn: 1234,56)
        elif ',' in temiz:
            temiz = temiz.replace(',', '.')
        return float(temiz)
    except ValueError:
        return 0.0

def bakiye_formatla(deger):
    try: return f"{float(deger):,.2f} TL".replace(",", "X").replace(".", ",").replace("X", ".")
    except: return str(deger)

# ==========================================
# YAN MENÜ (YEDEKLEME)
# ==========================================
with st.sidebar:
    st.header("⚙️ Sistem İşlemleri")
    st.info("Bulut sisteminde (Streamlit Cloud) verilerin silinmemesi için gün sonunda .db yedeğinizi alın.")
    
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "rb") as f:
            st.download_button(label="📥 Veritabanını Yedekle (.db)", data=f, file_name="tahsilat_yedek.db", mime="application/octet-stream")
    
    st.markdown("---")
    st.markdown("**Yedekten Geri Yükle**")
    yedek_dosya = st.file_uploader("Yedek .db Dosyasını Seçin", type=["db"])
    if yedek_dosya is not None:
        if st.button("🔄 Yedeği Yükle"):
            with open(DB_FILE, "wb") as f:
                f.write(yedek_dosya.getbuffer())
            st.success("Veritabanı başarıyla yüklendi!")
            st.rerun()

st.title("📈 Netsis Tahsilat ve Cari Takip Sistemi")

tab1, tab2 = st.tabs(["📋 Tüm Cariler (Ana Ekran)", "🔍 Tahsilat Takip Sayfası"])

# ==========================================
# 1. SEKME: ANA EKRAN
# ==========================================
with tab1:
    st.markdown("### Excel'den Cari Yükle")
    uploaded_file = st.file_uploader("Netsis Excel Raporunu Seçin", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        if st.button("Verileri Aktar ve Listeyi Yenile", type="primary"):
            df = pd.read_excel(uploaded_file)
            
            conn = get_db_connection()
            c = conn.cursor()
            
            c.execute("DELETE FROM ana_liste")
            
            c.execute("SELECT kod FROM takip")
            takipteki_kodlar = [row[0] for row in c.fetchall()]
            
            sayac = 0
            for index, row in df.iterrows():
                c_isim = str(row.get("Cari İsim", ""))
                if pd.isna(row.get("Cari İsim")) or not c_isim.strip(): continue
                
                c_kod = str(row.get("Cari Kod", "-"))
                if c_kod in takipteki_kodlar: continue
                
                c_tel = str(row.get("Telefon", "")) if pd.notna(row.get("Telefon")) else ""
                # Düzeltilmiş temizleme fonksiyonu devrede:
                bakiye_val = bakiye_temizle(row.get("Borç Bak.", 0.0))
                
                c.execute("INSERT INTO ana_liste (kod, isim, telefon, bakiye) VALUES (?, ?, ?, ?)", 
                          (c_kod, c_isim, c_tel, bakiye_val))
                sayac += 1
                
            conn.commit()
            conn.close()
            st.success(f"{sayac} adet cari başarıyla yüklendi.")
            st.rerun()

    st.markdown("---")
    conn = get_db_connection()
    df_ana = pd.read_sql_query("SELECT * FROM ana_liste", conn)
    conn.close()
    
    if not df_ana.empty:
        col1, col2, col3 = st.columns(3)
        min_b = col1.number_input("Min Bakiye (TL)", value=0.0)
        max_b = col2.number_input("Max Bakiye (TL)", value=9999999.0)
        arama = col3.text_input("Cari İsim veya Kod Ara").lower()
        
        mask = (df_ana["bakiye"] >= min_b) & (df_ana["bakiye"] <= max_b)
        if arama: 
            mask = mask & (df_ana["isim"].str.lower().str.contains(arama) | df_ana["kod"].str.lower().str.contains(arama))
        
        df_gosterim = df_ana[mask].copy()
        df_gosterim.rename(columns={"kod": "Cari Kod", "isim": "Cari İsim", "telefon": "Telefon", "bakiye": "Bakiye"}, inplace=True)
        df_gosterim.insert(0, "Seç", False)
        
        edited_df = st.data_editor(
            df_gosterim,
            hide_index=True,
            column_config={
                "Seç": st.column_config.CheckboxColumn("Seç", default=False),
                "Bakiye": st.column_config.NumberColumn("Bakiye (TL)", format="%.2f")
            },
            disabled=["Cari Kod", "Cari İsim", "Telefon", "Bakiye"],
            use_container_width=True,
            height=400
        )
        
        secilenler = edited_df[edited_df["Seç"] == True]
        if st.button("Seçilenleri Takibe Aktar ➔", type="primary") and not secilenler.empty:
            conn = get_db_connection()
            c = conn.cursor()
            for index, row in secilenler.iterrows():
                kod = row["Cari Kod"]
                c.execute("INSERT OR REPLACE INTO takip (kod, isim, telefon, bakiye, durum, ozel_durum, tarih) VALUES (?, ?, ?, ?, ?, ?, ?)",
                          (kod, row["Cari İsim"], row["Telefon"], row["Bakiye"], "Beklemede", "", ""))
                c.execute("DELETE FROM ana_liste WHERE kod=?", (kod,))
                
                ilk_log = f"Sistem: Cari takibe alındı. (Bakiye: {row['Bakiye']:,.2f} TL)"
                zaman = datetime.now().strftime("%d.%m.%Y %H:%M")
                c.execute("INSERT INTO loglar (cari_kod, tarih_saat, not_metni) VALUES (?, ?, ?)", (kod, zaman, ilk_log))
                
            conn.commit()
            conn.close()
            st.success("Seçilen cariler takibe aktarıldı!")
            st.rerun()

# ==========================================
# 2. SEKME: TAHSİLAT TAKİP VE LOG SİSTEMİ
# ==========================================
with tab2:
    conn = get_db_connection()
    df_takip = pd.read_sql_query("SELECT * FROM takip", conn)
    
    if not df_takip.empty:
        toplam_bakiye = df_takip['bakiye'].sum()
        odenen_bakiye = df_takip[df_takip['durum'] == 'Ödedi']['bakiye'].sum()
        bekleyen_sayisi = len(df_takip[df_takip['durum'] == 'Beklemede'])
        
        st.markdown("### 📊 Genel Durum Özeti")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Takipteki Cari Sayısı", len(df_takip))
        m2.metric("Henüz Aranmayan", bekleyen_sayisi)
        m3.metric("Tahsil Edilen (Ödedi)", f"{odenen_bakiye:,.2f} TL")
        m4.metric("Kalan Toplam Alacak", f"{(toplam_bakiye - odenen_bakiye):,.2f} TL")
        st.markdown("---")
        
        col_f1, col_f2 = st.columns(2)
        durumlar = ["Tümü", "Beklemede", "Arandı", "Ödedi", "Dönmedi"]
        secili_durum = col_f1.selectbox("Durum Filtresi", durumlar)
        
        temiz_ozeller = [str(x) for x in df_takip["ozel_durum"].unique() if x]
        ozel_durumlar = ["Tümü"] + sorted(temiz_ozeller)
        secili_ozel = col_f2.selectbox("Özel Durum Filtresi", ozel_durumlar)
        
        df_gosterim = df_takip.copy()
        if secili_durum != "Tümü": df_gosterim = df_gosterim[df_gosterim["durum"] == secili_durum]
        if secili_ozel != "Tümü": df_gosterim = df_gosterim[df_gosterim["ozel_durum"] == secili_ozel]
        
        df_gosterim.rename(columns={"kod": "Cari Kod", "isim": "Cari İsim", "telefon": "Telefon", 
                                    "bakiye": "Bakiye", "durum": "Durum", "ozel_durum": "Özel Durum", "tarih": "Tarih"}, inplace=True)
        df_gosterim.insert(0, "Seç", False)
        
        df_gosterim["Tarih"] = pd.to_datetime(df_gosterim["Tarih"], format="%d.%m.%Y", errors="coerce").dt.date
        
        st.info("Hızlı Güncelleme: Durum, Özel Durum ve Tarih hücrelerini tablodan direkt değiştirebilirsiniz.")
        edited_takip = st.data_editor(
            df_gosterim,
            hide_index=True,
            column_config={
                "Seç": st.column_config.CheckboxColumn("Sil", default=False),
                "Durum": st.column_config.SelectboxColumn("Durum", options=["Beklemede", "Arandı", "Ödedi", "Dönmedi"]),
                "Tarih": st.column_config.DateColumn("Tarih", format="DD.MM.YYYY"),
                "Bakiye": st.column_config.NumberColumn("Bakiye (TL)", format="%.2f")
            },
            disabled=["Cari Kod", "Cari İsim", "Telefon", "Bakiye"],
            use_container_width=True,
            height=350
        )
        
        degisiklik_yapildi = False
        c = conn.cursor()
        for index, row in edited_takip.iterrows():
            kod = row["Cari Kod"]
            orj_row = df_takip[df_takip["kod"] == kod].iloc[0]
            
            yeni_tarih = row["Tarih"].strftime("%d.%m.%Y") if pd.notna(row["Tarih"]) else ""
            yeni_ozel = str(row["Özel Durum"]) if pd.notna(row["Özel Durum"]) else ""
            
            if (orj_row["durum"] != row["Durum"] or 
                orj_row["ozel_durum"] != yeni_ozel or 
                orj_row["tarih"] != yeni_tarih):
                
                c.execute("UPDATE takip SET durum=?, ozel_durum=?, tarih=? WHERE kod=?", 
                          (row["Durum"], yeni_ozel, yeni_tarih, kod))
                degisiklik_yapildi = True
                
        if degisiklik_yapildi:
            conn.commit()
            st.toast("✅ Tablo güncellemeleri veritabanına kaydedildi!", icon="💾")
            time.sleep(1)
            st.rerun()

        silinecekler = edited_takip[edited_takip["Seç"] == True]
        if st.button("❌ Seçilenleri Takipten Çıkar ve Ana Ekrana Aktar"):
            for index, row in silinecekler.iterrows():
                kod = row["Cari Kod"]
                if not kod.startswith("MANUEL-"):
                    c.execute("INSERT OR REPLACE INTO ana_liste (kod, isim, telefon, bakiye) VALUES (?, ?, ?, ?)",
                              (kod, row["Cari İsim"], row["Telefon"], row["Bakiye"]))
                c.execute("DELETE FROM takip WHERE kod=?", (kod,))
                c.execute("DELETE FROM loglar WHERE cari_kod=?", (kod,)) 
            conn.commit()
            st.success("Seçilen cariler takipten çıkarıldı.")
            st.rerun()

        st.markdown("---")
        
        # ==========================================
        # LOG SİSTEMİ (GÖRÜŞME GEÇMİŞİ)
        # ==========================================
        st.markdown("### 📝 Cari Detay ve Görüşme Geçmişi")
        
        cari_liste = df_takip.apply(lambda x: f"{x['isim']} ({x['kod']})", axis=1).tolist()
        secilen_cari_str = st.selectbox("Görüşme detaylarını görmek için bir cari seçin:", cari_liste)
        
        if secilen_cari_str:
            secilen_kod = secilen_cari_str.split("(")[-1].replace(")", "")
            cari_detay = df_takip[df_takip['kod'] == secilen_kod].iloc[0]
            
            log_col1, log_col2 = st.columns([1, 2])
            
            with log_col1:
                st.info(f"**Bakiye:** {cari_detay['bakiye']:,.2f} TL\n\n**Telefon:** {cari_detay['telefon']}")
                with st.form(key=f"form_{secilen_kod}", clear_on_submit=True):
                    yeni_not = st.text_area("Bu cari için yeni bir görüşme notu ekleyin:", height=100)
                    if st.form_submit_button("Notu Veritabanına Kaydet"):
                        if yeni_not.strip():
                            zaman = datetime.now().strftime("%d.%m.%Y %H:%M")
                            c.execute("INSERT INTO loglar (cari_kod, tarih_saat, not_metni) VALUES (?, ?, ?)", 
                                      (secilen_kod, zaman, yeni_not.strip()))
                            conn.commit()
                            st.success("Not başarıyla eklendi!")
                            st.rerun()
                        else:
                            st.warning("Lütfen boş bir not kaydetmeyin.")
            
            with log_col2:
                st.markdown(f"**{cari_detay['isim']} - Geçmiş Görüşmeler**")
                df_log = pd.read_sql_query("SELECT tarih_saat, not_metni FROM loglar WHERE cari_kod=? ORDER BY id DESC", conn, params=(secilen_kod,))
                
                if not df_log.empty:
                    for index, row in df_log.iterrows():
                        st.markdown(f"🗓️ **{row['tarih_saat']}**")
                        st.markdown(f"> {row['not_metni']}")
                        st.divider()
                else:
                    st.write("Bu cari için henüz bir görüşme notu bulunmuyor.")
                    
    else:
        st.info("Takip listeniz şu an boş. Ana ekrandan cari aktarabilirsiniz.")
        
    conn.close()
