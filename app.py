import streamlit as st
import pandas as pd
import psycopg2
import os
import io
import time
import urllib.parse
from datetime import datetime
import plotly.express as px

st.set_page_config(page_title="Tahsilat ve Cari Takip", page_icon="📈", layout="wide")

# ==========================================
# BULUT VERİTABANI BAĞLANTI AYARLARI
# ==========================================
try:
    DB_URL = st.secrets["DB_URL"]
    if "sslmode=require" not in DB_URL:
        separator = "&" if "?" in DB_URL else "?"
        DB_URL += f"{separator}sslmode=require"
except Exception as e:
    st.error("Bağlantı linki bulunamadı. Lütfen Streamlit Cloud 'Secrets' bölümüne DB_URL'i eklediğinizden emin olun.")
    st.stop()

@st.cache_resource
def init_db():
    try:
        conn = psycopg2.connect(DB_URL)
        c = conn.cursor()
        c.execute('''CREATE TABLE IF NOT EXISTS takip (
                        kod TEXT PRIMARY KEY, 
                        isim TEXT, 
                        telefon TEXT, 
                        bakiye NUMERIC, 
                        durum TEXT, 
                        ozel_durum TEXT, 
                        tarih TEXT
                    )''')
        c.execute('''CREATE TABLE IF NOT EXISTS loglar (
                        id SERIAL PRIMARY KEY, 
                        cari_kod TEXT, 
                        tarih_saat TEXT, 
                        not_metni TEXT
                    )''')
        c.execute('''CREATE TABLE IF NOT EXISTS ana_liste (
                        kod TEXT PRIMARY KEY, 
                        isim TEXT, 
                        telefon TEXT, 
                        bakiye NUMERIC
                    )''')
        conn.commit()
        c.close()
        conn.close()
    except Exception as e:
        st.error(f"Veritabanı oluşturulamadı. Detay: {e}")

init_db()

def get_db_connection():
    return psycopg2.connect(DB_URL)

# --- YARDIMCI FONKSİYONLAR ---
def bakiye_temizle(deger):
    if pd.isna(deger): return 0.0
    if isinstance(deger, (int, float)): return float(deger)
    try:
        temiz = str(deger).replace(' TL', '').replace('₺', '').strip()
        if '.' in temiz and ',' in temiz: temiz = temiz.replace('.', '').replace(',', '.')
        elif ',' in temiz: temiz = temiz.replace(',', '.')
        return float(temiz)
    except ValueError: return 0.0

def bakiye_formatla(deger):
    try: return f"{float(deger):,.2f} TL".replace(",", "X").replace(".", ",").replace("X", ".")
    except: return str(deger)

def whatsapp_link_olustur(telefon, isim, bakiye):
    if not telefon or pd.isna(telefon): return None
    temiz_tel = "".join(filter(str.isdigit, str(telefon)))
    if len(temiz_tel) == 10: temiz_tel = "90" + temiz_tel
    elif len(temiz_tel) == 11 and temiz_tel.startswith("0"): temiz_tel = "9" + temiz_tel
    if len(temiz_tel) < 10: return None
    
    mesaj = f"Merhaba {isim}, sistemimizde {bakiye:,.2f} TL tutarında bakiyeniz bulunmaktadır. İyi çalışmalar dileriz."
    mesaj_kodlu = urllib.parse.quote(mesaj)
    return f"https://web.whatsapp.com/send?phone={temiz_tel}&text={mesaj_kodlu}"

# ==========================================
# ANA UYGULAMA ARAYÜZÜ
# ==========================================
with st.sidebar:
    st.header("⚙️ Sistem Durumu")
    st.success("🟢 Canlı Bulut Veritabanına Bağlı")
    st.write("Verileriniz anlık olarak buluta kaydedilmektedir. Uygulamadan güvenle çıkabilirsiniz.")

st.title("📈 Netsis Tahsilat ve Cari Takip Sistemi")

tab1, tab2, tab3 = st.tabs(["📋 Tüm Cariler (Ana Ekran)", "🔍 Tahsilat Takip Sayfası", "📊 Raporlar ve Analiz"])

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
            guncellenen_sayac = 0 
            
            for index, row in df.iterrows():
                c_isim = str(row.get("Cari İsim", ""))
                if pd.isna(row.get("Cari İsim")) or not c_isim.strip(): continue
                
                c_kod = str(row.get("Cari Kod", "")).strip()
                if c_kod == "nan" or c_kod == "-" or not c_kod:
                    c_kod = f"KODSUZ-{index}-{int(time.time())}"
                
                bakiye_val = bakiye_temizle(row.get("Borç Bak.", 0.0))
                
                if c_kod in takipteki_kodlar:
                    c.execute("UPDATE takip SET bakiye=%s WHERE kod=%s", (bakiye_val, c_kod))
                    guncellenen_sayac += 1
                    continue
                
                c_tel = str(row.get("Telefon", "")) if pd.notna(row.get("Telefon")) else ""
                
                c.execute("""
                    INSERT INTO ana_liste (kod, isim, telefon, bakiye) VALUES (%s, %s, %s, %s)
                    ON CONFLICT (kod) DO UPDATE SET 
                    isim=EXCLUDED.isim, telefon=EXCLUDED.telefon, bakiye=EXCLUDED.bakiye
                """, (c_kod, c_isim, c_tel, bakiye_val))
                sayac += 1
                
            conn.commit()
            c.close()
            conn.close()
            
            mesaj = f"{sayac} adet yeni cari ana listeye eklendi."
            if guncellenen_sayac > 0: mesaj += f"\n\n✅ Takipteki {guncellenen_sayac} adet carinin bakiyesi güncellendi!"
            st.success(mesaj)
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
        
        df_ana["bakiye"] = pd.to_numeric(df_ana["bakiye"], errors='coerce')
        mask = (df_ana["bakiye"] >= min_b) & (df_ana["bakiye"] <= max_b)
        if arama: mask = mask & (df_ana["isim"].str.lower().str.contains(arama) | df_ana["kod"].str.lower().str.contains(arama))
        
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
            try:
                conn = get_db_connection()
                c = conn.cursor()
                for index, row in secilenler.iterrows():
                    # Veri Tiplerini Güvence Altına Alma (HATA ÇÖZÜMÜ BURADA)
                    kod = str(row["Cari Kod"])
                    isim = str(row["Cari İsim"])
                    telefon = str(row["Telefon"])
                    bakiye = float(row["Bakiye"]) if pd.notna(row["Bakiye"]) else 0.0
                    
                    c.execute("""
                        INSERT INTO takip (kod, isim, telefon, bakiye, durum, ozel_durum, tarih) 
                        VALUES (%s, %s, %s, %s, %s, %s, %s) 
                        ON CONFLICT (kod) DO UPDATE SET 
                        isim=EXCLUDED.isim, telefon=EXCLUDED.telefon, bakiye=EXCLUDED.bakiye
                    """, (kod, isim, telefon, bakiye, "Beklemede", "", ""))
                    
                    c.execute("DELETE FROM ana_liste WHERE kod=%s", (kod,))
                    ilk_log = f"Sistem: Cari takibe alındı. (Bakiye: {bakiye:,.2f} TL)"
                    zaman = datetime.now().strftime("%d.%m.%Y %H:%M")
                    c.execute("INSERT INTO loglar (cari_kod, tarih_saat, not_metni) VALUES (%s, %s, %s)", (kod, zaman, ilk_log))
                
                conn.commit()
                c.close()
                conn.close()
                st.success("Seçilen cariler takibe aktarıldı!")
                time.sleep(1) # Kullanıcının başarı mesajını görebilmesi için minik bir bekleme
                st.rerun()
            except Exception as e:
                st.error(f"Aktarım sırasında bir hata oluştu: {e}")

# ==========================================
# 2. SEKME: TAHSİLAT TAKİP VE LOG
# ==========================================
with tab2:
    conn = get_db_connection()
    df_takip = pd.read_sql_query("SELECT * FROM takip", conn)
    
    if not df_takip.empty:
        df_takip["bakiye"] = pd.to_numeric(df_takip["bakiye"], errors='coerce')
        
        bugun = datetime.now().date()
        vadesi_gecen_cariler = []
        bugun_aranacaklar = []
        
        for index, row in df_takip.iterrows():
            if row["durum"] == "Ödedi": continue
            if row["tarih"]:
                try:
                    tarih_obj = datetime.strptime(row["tarih"], "%d.%m.%Y").date()
                    if tarih_obj < bugun: vadesi_gecen_cariler.append(row)
                    elif tarih_obj == bugun: bugun_aranacaklar.append(row)
                except: pass
                
        if vadesi_gecen_cariler or bugun_aranacaklar:
            with st.expander("🚨 AKILLI UYARILAR (İlgilenilmesi Gerekenler)", expanded=True):
                if vadesi_gecen_cariler:
                    toplam_vade_gecen = sum(float(r['bakiye']) for r in vadesi_gecen_cariler)
                    st.error(f"**VADESİ GEÇEN {len(vadesi_gecen_cariler)} CARİ VAR!** (Toplam Risk: {toplam_vade_gecen:,.2f} TL)")
                    for r in vadesi_gecen_cariler:
                        st.write(f"⚠️ {r['isim']} | Tarih: {r['tarih']} | Bakiye: {float(r['bakiye']):,.2f} TL")
                if bugun_aranacaklar:
                    toplam_bugun = sum(float(r['bakiye']) for r in bugun_aranacaklar)
                    st.warning(f"**BUGÜN ARANACAK {len(bugun_aranacaklar)} CARİ VAR!** (Toplam Beklenti: {toplam_bugun:,.2f} TL)")
                    for r in bugun_aranacaklar:
                        st.write(f"📞 {r['isim']} | Bakiye: {float(r['bakiye']):,.2f} TL")
            st.markdown("---")

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
        
        df_gosterim["WhatsApp"] = df_gosterim.apply(lambda r: whatsapp_link_olustur(r['telefon'], r['isim'], r['bakiye']), axis=1)
        
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
                "Bakiye": st.column_config.NumberColumn("Bakiye (TL)", format="%.2f"),
                "WhatsApp": st.column_config.LinkColumn("WhatsApp İletişim", display_text="💬 Mesaj Gönder")
            },
            disabled=["Cari Kod", "Cari İsim", "Telefon", "Bakiye", "WhatsApp"],
            use_container_width=True,
            height=350
        )
        
        degisiklik_yapildi = False
        c = conn.cursor()
        for index, row in edited_takip.iterrows():
            kod = str(row["Cari Kod"])
            orj_row = df_takip[df_takip["kod"] == kod].iloc[0]
            yeni_tarih = row["Tarih"].strftime("%d.%m.%Y") if pd.notna(row["Tarih"]) else ""
            yeni_ozel = str(row["Özel Durum"]) if pd.notna(row["Özel Durum"]) else ""
            yeni_durum = str(row["Durum"])
            
            if (orj_row["durum"] != yeni_durum or orj_row["ozel_durum"] != yeni_ozel or orj_row["tarih"] != yeni_tarih):
                c.execute("UPDATE takip SET durum=%s, ozel_durum=%s, tarih=%s WHERE kod=%s", (yeni_durum, yeni_ozel, yeni_tarih, kod))
                degisiklik_yapildi = True
                
        if degisiklik_yapildi:
            conn.commit()
            st.toast("✅ Tablo güncellemeleri buluta kaydedildi!", icon="☁️")
            time.sleep(1)
            st.rerun()

        silinecekler = edited_takip[edited_takip["Seç"] == True]
        if st.button("❌ Seçilenleri Takipten Çıkar ve Ana Ekrana Aktar"):
            try:
                for index, row in silinecekler.iterrows():
                    kod = str(row["Cari Kod"])
                    isim = str(row["Cari İsim"])
                    telefon = str(row["Telefon"])
                    bakiye = float(row["Bakiye"]) if pd.notna(row["Bakiye"]) else 0.0
                    
                    if not kod.startswith("MANUEL-") and not kod.startswith("KODSUZ-"):
                        c.execute("""
                            INSERT INTO ana_liste (kod, isim, telefon, bakiye) VALUES (%s, %s, %s, %s)
                            ON CONFLICT (kod) DO UPDATE SET isim=EXCLUDED.isim, telefon=EXCLUDED.telefon, bakiye=EXCLUDED.bakiye
                        """, (kod, isim, telefon, bakiye))
                    c.execute("DELETE FROM takip WHERE kod=%s", (kod,))
                    c.execute("DELETE FROM loglar WHERE cari_kod=%s", (kod,)) 
                conn.commit()
                st.success("Seçilen cariler takipten çıkarıldı.")
                time.sleep(1)
                st.rerun()
            except Exception as e:
                st.error(f"Çıkarma işlemi sırasında bir hata oluştu: {e}")
            
        c.close()

        st.markdown("---")
        st.markdown("### 📝 Cari Detay ve Görüşme Geçmişi")
        cari_liste = df_takip.apply(lambda x: f"{x['isim']} ({x['kod']})", axis=1).tolist()
        secilen_cari_str = st.selectbox("Görüşme detaylarını görmek için bir cari seçin:", cari_liste)
        
        if secilen_cari_str:
            secilen_kod = secilen_cari_str.split("(")[-1].replace(")", "")
            cari_detay = df_takip[df_takip['kod'] == secilen_kod].iloc[0]
            log_col1, log_col2 = st.columns([1, 2])
            
            with log_col1:
                st.info(f"**Bakiye:** {float(cari_detay['bakiye']):,.2f} TL\n\n**Telefon:** {cari_detay['telefon']}")
                with st.form(key=f"form_{secilen_kod}", clear_on_submit=True):
                    yeni_not = st.text_area("Bu cari için yeni bir görüşme notu ekleyin:", height=100)
                    if st.form_submit_button("Notu Buluta Kaydet"):
                        if yeni_not.strip():
                            zaman = datetime.now().strftime("%d.%m.%Y %H:%M")
                            c = conn.cursor()
                            c.execute("INSERT INTO loglar (cari_kod, tarih_saat, not_metni) VALUES (%s, %s, %s)", (secilen_kod, zaman, yeni_not.strip()))
                            conn.commit()
                            c.close()
                            st.success("Not başarıyla eklendi!")
                            st.rerun()
                        else: st.warning("Lütfen boş bir not kaydetmeyin.")
            
            with log_col2:
                st.markdown(f"**{cari_detay['isim']} - Geçmiş Görüşmeler**")
                df_log = pd.read_sql_query(f"SELECT tarih_saat, not_metni FROM loglar WHERE cari_kod='{secilen_kod}' ORDER BY id DESC", conn)
                if not df_log.empty:
                    for index, row in df_log.iterrows():
                        st.markdown(f"🗓️ **{row['tarih_saat']}**")
                        st.markdown(f"> {row['not_metni']}")
                        st.divider()
                else: st.write("Bu cari için henüz bir görüşme notu bulunmuyor.")
    else:
        st.info("Takip listeniz şu an boş. Ana ekrandan cari aktarabilirsiniz.")
    conn.close()

# ==========================================
# 3. SEKME: GÖRSEL RAPORLAR
# ==========================================
with tab3:
    st.markdown("### 📊 Gelişmiş Finans ve Tahsilat Raporları")
    conn = get_db_connection()
    df_rapor = pd.read_sql_query("SELECT * FROM takip", conn)
    conn.close()
    
    if not df_rapor.empty:
        df_rapor["bakiye"] = pd.to_numeric(df_rapor["bakiye"], errors='coerce')
        r_col1, r_col2 = st.columns(2)
        
        with r_col1:
            st.markdown("#### Paranın Durum Dağılımı")
            df_pie = df_rapor.groupby("durum")["bakiye"].sum().reset_index()
            fig_pie = px.pie(df_pie, values='bakiye', names='durum', 
                             hole=0.4, color_discrete_sequence=px.colors.qualitative.Set2)
            fig_pie.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with r_col2:
            st.markdown("#### Planlanan Nakit Akışı (Tarihe Göre)")
            df_nakit = df_rapor[(df_rapor["durum"] != "Ödedi") & (df_rapor["tarih"] != "")]
            if not df_nakit.empty:
                df_nakit["tarih_obj"] = pd.to_datetime(df_nakit["tarih"], format="%d.%m.%Y", errors='coerce')
                df_nakit = df_nakit.dropna(subset=["tarih_obj"]).sort_values("tarih_obj")
                
                df_bar = df_nakit.groupby("tarih")["bakiye"].sum().reset_index()
                
                fig_bar = px.bar(df_bar, x='tarih', y='bakiye', text='bakiye',
                                 labels={'tarih': 'Ödeme Tarihi', 'bakiye': 'Beklenen Tutar (TL)'},
                                 color_discrete_sequence=['#1f77b4'])
                fig_bar.update_traces(texttemplate='%{text:,.2f} TL', textposition='outside')
                st.plotly_chart(fig_bar, use_container_width=True)
            else:
                st.info("Tarihi planlanmış aktif bir tahsilat bulunmuyor.")
                
        st.markdown("#### 🎯 Özel Durumlara Göre Risk Analizi")
        df_ozel = df_rapor[df_rapor["ozel_durum"] != ""].groupby("ozel_durum")["bakiye"].sum().reset_index()
        if not df_ozel.empty:
            df_ozel = df_ozel.sort_values(by="bakiye", ascending=False)
            df_ozel.rename(columns={"ozel_durum": "Sizin Eklediğiniz Durum", "bakiye": "İçerideki Toplam Tutar (TL)"}, inplace=True)
            st.dataframe(df_ozel.style.format({"İçerideki Toplam Tutar (TL)": "{:,.2f} TL"}), use_container_width=True)
        else:
            st.write("Henüz bir 'Özel Durum' etiketi kullanmadınız.")
    else:
        st.info("Rapor oluşturulabilmesi için Takip listenize veri eklemelisiniz.")
