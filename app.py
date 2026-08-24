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
# SESSION STATE (HAFIZA) TANIMLAMALARI
# ==========================================
if "secili_uyari_kod" not in st.session_state:
    st.session_state["secili_uyari_kod"] = None

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
        c.execute('''CREATE TABLE IF NOT EXISTS gorevler (
                        id SERIAL PRIMARY KEY, 
                        gorev_adi TEXT, 
                        tarih TEXT, 
                        durum TEXT, 
                        notlar TEXT
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

def whatsapp_link_olustur(telefon, isim, bakiye):
    if not telefon or pd.isna(telefon): return None
    temiz_tel = "".join(filter(str.isdigit, str(telefon)))
    if len(temiz_tel) == 10: temiz_tel = "90" + temiz_tel
    elif len(temiz_tel) == 11 and temiz_tel.startswith("0"): temiz_tel = "9" + temiz_tel
    if len(temiz_tel) < 10: return None
    mesaj = f"Merhaba {isim}, sistemimizde {bakiye:,.2f} TL tutarında bakiyeniz bulunmaktadır. İyi çalışmalar dileriz."
    mesaj_kodlu = urllib.parse.quote(mesaj)
    return f"https://wa.me/{temiz_tel}?text={mesaj_kodlu}"

def create_ics_file(baslik, aciklama, tarih_str):
    try:
        dt = datetime.strptime(tarih_str, "%d.%m.%Y")
        dt_start = dt.strftime("%Y%m%dT090000")
        dt_end = dt.strftime("%Y%m%dT093000")
        ics_icerik = f"""BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Tahsilat Sistemi//TR
BEGIN:VEVENT
SUMMARY:{baslik}
DESCRIPTION:{aciklama}
DTSTART:{dt_start}
DTEND:{dt_end}
BEGIN:VALARM
TRIGGER:-PT15M
ACTION:DISPLAY
DESCRIPTION:Hatırlatma
END:VALARM
END:VEVENT
END:VCALENDAR"""
        return ics_icerik
    except:
        return None

def arama_temizle(deger):
    if pd.isna(deger): return ""
    metin = str(deger)
    degisim = {"I": "ı", "İ": "i", "Ş": "ş", "Ç": "ç", "Ö": "ö", "Ü": "ü", "Ğ": "ğ"}
    for b, k in degisim.items():
        metin = metin.replace(b, k)
    metin = metin.lower()
    degisim_fuzzy = {"ı": "i", "ğ": "g", "ü": "u", "ş": "s", "ö": "o", "ç": "c"}
    for b, k in degisim_fuzzy.items():
        metin = metin.replace(b, k)
    return metin.strip()

# ==========================================
# ANA UYGULAMA ARAYÜZÜ
# ==========================================
with st.sidebar:
    st.header("⚙️ Sistem Durumu")
    st.success("🟢 Canlı Bulut Veritabanına Bağlı")
    st.write("Verileriniz anlık olarak buluta kaydedilmektedir. Uygulamadan güvenle çıkabilirsiniz.")

st.title("📈 Netsis Tahsilat ve Cari Takip Sistemi")

tab1, tab2, tab3, tab4 = st.tabs([
    "📋 Tüm Cariler (Ana Ekran)", 
    "🔍 Tahsilat Takip Sayfası", 
    "📊 Raporlar ve Analiz",
    "✅ Görev Yöneticisi"
])

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
            sayac, guncellenen_sayac = 0, 0
            
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
                    ON CONFLICT (kod) DO UPDATE SET isim=EXCLUDED.isim, telefon=EXCLUDED.telefon, bakiye=EXCLUDED.bakiye
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
        arama = col3.text_input("Cari İsim veya Kod Ara").strip()
        
        df_ana["bakiye"] = pd.to_numeric(df_ana["bakiye"], errors='coerce')
        mask = (df_ana["bakiye"] >= min_b) & (df_ana["bakiye"] <= max_b)
        
        if arama: 
            arama_temiz = arama_temizle(arama)
            isim_mask = df_ana["isim"].apply(arama_temizle).str.contains(arama_temiz)
            kod_mask = df_ana["kod"].apply(arama_temizle).str.contains(arama_temiz)
            mask = mask & (isim_mask | kod_mask)
        
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
                    kod = str(row["Cari Kod"])
                    isim = str(row["Cari İsim"])
                    telefon = str(row["Telefon"])
                    bakiye = float(row["Bakiye"]) if pd.notna(row["Bakiye"]) else 0.0
                    
                    c.execute("""
                        INSERT INTO takip (kod, isim, telefon, bakiye, durum, ozel_durum, tarih) 
                        VALUES (%s, %s, %s, %s, %s, %s, %s) 
                        ON CONFLICT (kod) DO UPDATE SET isim=EXCLUDED.isim, telefon=EXCLUDED.telefon, bakiye=EXCLUDED.bakiye
                    """, (kod, isim, telefon, bakiye, "Beklemede", "", ""))
                    c.execute("DELETE FROM ana_liste WHERE kod=%s", (kod,))
                    
                    zaman = datetime.now().strftime("%d.%m.%Y %H:%M")
                    c.execute("INSERT INTO loglar (cari_kod, tarih_saat, not_metni) VALUES (%s, %s, %s)", (kod, zaman, f"Sistem: Cari takibe alındı. (Bakiye: {bakiye:,.2f} TL)"))
                conn.commit()
                c.close()
                conn.close()
                st.success("Seçilen cariler takibe aktarıldı!")
                time.sleep(1) 
                st.rerun()
            except Exception as e:
                st.error(f"Aktarım sırasında bir hata oluştu: {e}")

# ==========================================
# 2. SEKME: TAHSİLAT TAKİP VE LOG
# ==========================================
with tab2:
    with st.expander("➕ Manuel Cari Ekle (Excel'de Olmayanlar İçin)", expanded=False):
        with st.form("manuel_cari_form", clear_on_submit=True):
            col_m1, col_m2 = st.columns(2)
            m_isim = col_m1.text_input("Cari İsim *")
            m_bakiye = col_m2.number_input("Bakiye (TL)", min_value=0.0, step=100.0)
            
            col_m3, col_m4 = st.columns(2)
            m_tel = col_m3.text_input("Telefon (İsteğe Bağlı)")
            m_kod = col_m4.text_input("Cari Kod (Boş bırakırsanız sistem otomatik üretir)")
            
            if st.form_submit_button("Listeye Ekle ve Takibe Al"):
                if m_isim.strip():
                    if not m_kod.strip():
                        m_kod = f"MANUEL-{int(time.time())}"
                    try:
                        conn = get_db_connection()
                        c = conn.cursor()
                        c.execute("SELECT kod FROM takip WHERE kod=%s", (m_kod,))
                        if c.fetchone():
                            st.error("Bu Cari Kod zaten kullanımda. Lütfen farklı bir kod girin.")
                        else:
                            c.execute("""
                                INSERT INTO takip (kod, isim, telefon, bakiye, durum, ozel_durum, tarih) 
                                VALUES (%s, %s, %s, %s, %s, %s, %s)
                            """, (m_kod, m_isim, m_tel, m_bakiye, "Beklemede", "Manuel Eklendi", ""))
                            zaman = datetime.now().strftime("%d.%m.%Y %H:%M")
                            c.execute("INSERT INTO loglar (cari_kod, tarih_saat, not_metni) VALUES (%s, %s, %s)", (m_kod, zaman, f"Sistem: Manuel olarak eklendi. (Bakiye: {m_bakiye:,.2f} TL)"))
                            conn.commit()
                            st.success(f"'{m_isim}' başarıyla takip listesine eklendi!")
                            time.sleep(1)
                            st.rerun()
                        c.close()
                        conn.close()
                    except Exception as e:
                        st.error(f"Ekleme hatası: {e}")
                else:
                    st.error("Lütfen Cari İsim giriniz!")

    st.markdown("---")
    
    conn = get_db_connection()
    df_takip = pd.read_sql_query("SELECT * FROM takip", conn)
    
    if not df_takip.empty:
        df_takip["bakiye"] = pd.to_numeric(df_takip["bakiye"], errors='coerce')
        bugun = datetime.now().date()
        vadesi_gecen_cariler, bugun_aranacaklar = [], []
        
        for index, row in df_takip.iterrows():
            if row["durum"] == "Ödedi": continue
            if row["tarih"]:
                try:
                    tarih_obj = datetime.strptime(row["tarih"], "%d.%m.%Y").date()
                    if tarih_obj < bugun: vadesi_gecen_cariler.append(row)
                    elif tarih_obj == bugun: bugun_aranacaklar.append(row)
                except: pass
                
        # --- YENİ: TIKLANABİLİR AKILLI UYARILAR ---
        if vadesi_gecen_cariler or bugun_aranacaklar:
            with st.expander("🚨 AKILLI UYARILAR (İlgilenilmesi Gerekenler)", expanded=True):
                if vadesi_gecen_cariler:
                    st.error(f"**VADESİ GEÇEN {len(vadesi_gecen_cariler)} CARİ VAR!** (Risk: {sum(float(r['bakiye']) for r in vadesi_gecen_cariler):,.2f} TL)")
                    for r in vadesi_gecen_cariler: 
                        w1, w2 = st.columns([5, 1])
                        w1.write(f"⚠️ **{r['isim']}** | Tarih: {r['tarih']} | Bakiye: {float(r['bakiye']):,.2f} TL")
                        if w2.button("🔍 İncele", key=f"vade_{r['kod']}", use_container_width=True):
                            st.session_state["secili_uyari_kod"] = r['kod']
                            st.rerun()
                            
                if bugun_aranacaklar:
                    st.warning(f"**BUGÜN ARANACAK {len(bugun_aranacaklar)} CARİ VAR!** (Beklenti: {sum(float(r['bakiye']) for r in bugun_aranacaklar):,.2f} TL)")
                    for r in bugun_aranacaklar: 
                        w1, w2 = st.columns([5, 1])
                        w1.write(f"📞 **{r['isim']}** | Bakiye: {float(r['bakiye']):,.2f} TL")
                        if w2.button("🔍 İncele", key=f"bugun_{r['kod']}", use_container_width=True):
                            st.session_state["secili_uyari_kod"] = r['kod']
                            st.rerun()
            st.markdown("---")

        st.markdown("### 📊 Genel Durum Özeti")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Takipteki Cari Sayısı", len(df_takip))
        m2.metric("Henüz Aranmayan", len(df_takip[df_takip['durum'] == 'Beklemede']))
        m3.metric("Tahsil Edilen", f"{df_takip[df_takip['durum'] == 'Ödedi']['bakiye'].sum():,.2f} TL")
        m4.metric("Kalan Alacak", f"{(df_takip['bakiye'].sum() - df_takip[df_takip['durum'] == 'Ödedi']['bakiye'].sum()):,.2f} TL")
        st.markdown("---")
        
        col_f1, col_f2, col_f3 = st.columns(3)
        secili_durum = col_f1.selectbox("Durum Filtresi", ["Tümü", "Beklemede", "Arandı", "Ödedi", "Dönmedi"])
        temiz_ozeller = [str(x) for x in df_takip["ozel_durum"].unique() if x]
        secili_ozel = col_f2.selectbox("Özel Durum Filtresi", ["Tümü"] + sorted(temiz_ozeller))
        arama_takip = col_f3.text_input("Cari İsim Ara 🔍", placeholder="Örn: fatı, Fati...").strip()
        
        df_gosterim = df_takip.copy()
        if secili_durum != "Tümü": df_gosterim = df_gosterim[df_gosterim["durum"] == secili_durum]
        if secili_ozel != "Tümü": df_gosterim = df_gosterim[df_gosterim["ozel_durum"] == secili_ozel]
        
        if arama_takip:
            a_t = arama_temizle(arama_takip)
            isim_mask = df_gosterim["isim"].apply(arama_temizle).str.contains(a_t)
            df_gosterim = df_gosterim[isim_mask]
        
        df_gosterim["WhatsApp"] = df_gosterim.apply(lambda r: whatsapp_link_olustur(r['telefon'], r['isim'], r['bakiye']), axis=1)
        df_gosterim.rename(columns={"kod": "Cari Kod", "isim": "Cari İsim", "telefon": "Telefon", "bakiye": "Bakiye", "durum": "Durum", "ozel_durum": "Özel Durum", "tarih": "Tarih"}, inplace=True)
        df_gosterim.insert(0, "Seç", False)
        df_gosterim["Tarih"] = pd.to_datetime(df_gosterim["Tarih"], format="%d.%m.%Y", errors="coerce").dt.date
        
        st.info("💡 **Hızlı Düzenleme:** Tablodan istediğiniz kadar değişikliği hızlıca yapın, ardından aşağıdaki butona basıp kaydedin.")
        
        edited_takip = st.data_editor(
            df_gosterim, hide_index=True,
            column_config={
                "Seç": st.column_config.CheckboxColumn("Seç / İncele", default=False),
                "Durum": st.column_config.SelectboxColumn("Durum", options=["Beklemede", "Arandı", "Ödedi", "Dönmedi"]),
                "Tarih": st.column_config.DateColumn("Tarih", format="DD.MM.YYYY"),
                "Bakiye": st.column_config.NumberColumn("Bakiye (TL)", format="%.2f"),
                "WhatsApp": st.column_config.LinkColumn("WhatsApp İletişim", display_text="💬 Mesaj Gönder")
            },
            disabled=["Cari Kod", "Cari İsim", "Telefon", "Bakiye", "WhatsApp"],
            use_container_width=True, height=350
        )
        
        secili_satirlar = edited_takip[edited_takip["Seç"] == True]
        
        # Eğer kullanıcı tablodan kutucuk işaretlerse veya filtre atarsa, uyarı tıklamasını sıfırla
        if len(secili_satirlar) > 0 or len(df_gosterim) == 1:
            st.session_state["secili_uyari_kod"] = None
        
        degisiklik_var_mi = False
        for index, row in edited_takip.iterrows():
            kod = str(row["Cari Kod"])
            orj_row = df_takip[df_takip["kod"] == kod].iloc[0]
            yeni_tarih = row["Tarih"].strftime("%d.%m.%Y") if pd.notna(row["Tarih"]) else ""
            yeni_ozel = str(row["Özel Durum"]) if pd.notna(row["Özel Durum"]) else ""
            yeni_durum = str(row["Durum"])
            
            if (str(orj_row["durum"]) != yeni_durum or str(orj_row["ozel_durum"]) != yeni_ozel or str(orj_row["tarih"]) != yeni_tarih):
                degisiklik_var_mi = True
                break
        
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            btn_label = "🔴 Değişiklikler Var! Tıkla ve Kaydet" if degisiklik_var_mi else "🟢 Tüm Veriler Güncel (Değişiklik Yok)"
            btn_type = "primary" if degisiklik_var_mi else "secondary"
            kaydet_basildi = st.button(btn_label, type=btn_type, disabled=not degisiklik_var_mi, use_container_width=True, key="btn_cari_kaydet")
            
        with col_btn2:
            sil_basildi = st.button("❌ Seçilenleri Takipten Çıkar ve Ana Ekrana Aktar", use_container_width=True, key="btn_cari_sil")
        
        if kaydet_basildi:
            c = conn.cursor()
            zaman_simdi = datetime.now().strftime("%d.%m.%Y %H:%M")
            for index, row in edited_takip.iterrows():
                kod = str(row["Cari Kod"])
                orj_row = df_takip[df_takip["kod"] == kod].iloc[0]
                yeni_tarih = row["Tarih"].strftime("%d.%m.%Y") if pd.notna(row["Tarih"]) else ""
                yeni_ozel = str(row["Özel Durum"]) if pd.notna(row["Özel Durum"]) else ""
                eski_durum = str(orj_row["durum"])
                yeni_durum = str(row["Durum"])
                
                if (eski_durum != yeni_durum or str(orj_row["ozel_durum"]) != yeni_ozel or str(orj_row["tarih"]) != yeni_tarih):
                    c.execute("UPDATE takip SET durum=%s, ozel_durum=%s, tarih=%s WHERE kod=%s", (yeni_durum, yeni_ozel, yeni_tarih, kod))
                    
                    if eski_durum != yeni_durum:
                        log_mesaji = f"Sistem: Durum güncellendi ({eski_durum} ➔ {yeni_durum})"
                        c.execute("INSERT INTO loglar (cari_kod, tarih_saat, not_metni) VALUES (%s, %s, %s)", (kod, zaman_simdi, log_mesaji))
                        
            conn.commit()
            c.close()
            st.success("✅ Tablodaki tüm güncellemeler tek seferde buluta kaydedildi!")
            time.sleep(1)
            st.rerun()

        if sil_basildi:
            if not secili_satirlar.empty:
                c = conn.cursor()
                for index, row in secili_satirlar.iterrows():
                    kod = str(row["Cari Kod"])
                    if not kod.startswith("MANUEL-") and not kod.startswith("KODSUZ-"):
                        c.execute("INSERT INTO ana_liste (kod, isim, telefon, bakiye) VALUES (%s, %s, %s, %s) ON CONFLICT (kod) DO UPDATE SET isim=EXCLUDED.isim, telefon=EXCLUDED.telefon, bakiye=EXCLUDED.bakiye", 
                                  (kod, str(row["Cari İsim"]), str(row["Telefon"]), float(row["Bakiye"]) if pd.notna(row["Bakiye"]) else 0.0))
                    c.execute("DELETE FROM takip WHERE kod=%s", (kod,))
                    c.execute("DELETE FROM loglar WHERE cari_kod=%s", (kod,)) 
                conn.commit()
                c.close()
                st.success("Seçilen cariler takipten çıkarıldı.")
                time.sleep(1)
                st.rerun()
            else: 
                st.warning("Lütfen takipten çıkarmak için tablodan birilerini seçin.")

        st.markdown("---")
        
        # --- CARİ DETAY PANELİ AKILLI AÇILIŞ MANTIĞI ---
        aktif_cari_kod = None
        
        if len(secili_satirlar) == 1:
            aktif_cari_kod = str(secili_satirlar.iloc[0]["Cari Kod"])
        elif len(secili_satirlar) > 1:
            st.warning("⚠️ Görüşme detaylarını görmek için tablodan sadece **BİR** kişinin kutucuğunu işaretleyin.")
        elif len(df_gosterim) == 1:
            aktif_cari_kod = str(df_gosterim.iloc[0]["Cari Kod"])
        elif st.session_state["secili_uyari_kod"]:
            aktif_cari_kod = st.session_state["secili_uyari_kod"]
            if aktif_cari_kod in df_takip["kod"].values:
                isim_uyari = df_takip[df_takip["kod"] == aktif_cari_kod].iloc[0]["isim"]
                col_u1, col_u2 = st.columns([4, 1])
                col_u1.success(f"🚨 Akıllı Uyarılardan **{isim_uyari}** inceleniyor.")
                if col_u2.button("❌ İncelemeyi Kapat", use_container_width=True):
                    st.session_state["secili_uyari_kod"] = None
                    st.rerun()
        else:
            cari_liste = ["-- Lütfen tablodan bir cari işaretleyin veya arama yapın --"] + df_takip.apply(lambda x: f"{x['isim']} ({x['kod']})", axis=1).tolist()
            secilen_cari_str = st.selectbox("Manuel Cari Seçimi (İsteğe Bağlı):", cari_liste)
            if secilen_cari_str != "-- Lütfen tablodan bir cari işaretleyin veya arama yapın --":
                aktif_cari_kod = secilen_cari_str.split("(")[-1].replace(")", "")
        
        # Eğer aktif bir cari varsa paneli aç
        if aktif_cari_kod and aktif_cari_kod in df_takip['kod'].values:
            cari_detay = df_takip[df_takip['kod'] == aktif_cari_kod].iloc[0]
            log_col1, log_col2 = st.columns([1, 2])
            with log_col1:
                st.info(f"**Cari:** {cari_detay['isim']}\n\n**Bakiye:** {float(cari_detay['bakiye']):,.2f} TL\n\n**Telefon:** {cari_detay['telefon']}")
                
                st.markdown("📅 **Takvim Hatırlatıcısı**")
                hatirlatma_tarihi = st.date_input("Hangi gün arayacaksınız?", format="DD.MM.YYYY", key=f"date_{aktif_cari_kod}")
                if hatirlatma_tarihi:
                    ics_baslik = f"☎️ {cari_detay['isim']} Tahsilat Araması"
                    ics_aciklama = f"Bakiye: {float(cari_detay['bakiye']):,.2f} TL\\nTelefon: {cari_detay['telefon']}\\nDurum: {cari_detay['durum']}"
                    ics_data = create_ics_file(ics_baslik, ics_aciklama, hatirlatma_tarihi.strftime("%d.%m.%Y"))
                    
                    if ics_data:
                        st.download_button(
                            label="🗓️ iPhone Takvimine Hatırlatıcı Ekle",
                            data=ics_data,
                            file_name=f"Tahsilat_{cari_detay['isim'].replace(' ', '_')}.ics",
                            mime="text/calendar",
                            use_container_width=True
                        )
                st.markdown("---")

                with st.form(key=f"form_{aktif_cari_kod}", clear_on_submit=True):
                    yeni_not = st.text_area("Bu cari için yeni bir görüşme notu ekleyin:", height=100)
                    if st.form_submit_button("Notu Buluta Kaydet"):
                        if yeni_not.strip():
                            zaman = datetime.now().strftime("%d.%m.%Y %H:%M")
                            c = conn.cursor()
                            c.execute("INSERT INTO loglar (cari_kod, tarih_saat, not_metni) VALUES (%s, %s, %s)", (aktif_cari_kod, zaman, yeni_not.strip()))
                            conn.commit()
                            c.close()
                            st.success("Not başarıyla eklendi!")
                            st.rerun()
                        else: st.warning("Lütfen boş bir not kaydetmeyin.")
            
            with log_col2:
                st.markdown(f"**Geçmiş Görüşmeler ve Durum Logları**")
                df_log = pd.read_sql_query(f"SELECT id, tarih_saat, not_metni FROM loglar WHERE cari_kod='{aktif_cari_kod}' ORDER BY id DESC", conn)
                
                if not df_log.empty:
                    df_log.insert(0, "Sil", False)
                    df_log.rename(columns={"tarih_saat": "Tarih / Saat", "not_metni": "Görüşme Notu"}, inplace=True)
                    
                    st.caption("Aşağıdaki notların üzerine tıklayarak değiştirebilir veya sol taraftan seçerek silebilirsiniz.")
                    
                    edited_log = st.data_editor(
                        df_log, hide_index=True,
                        column_config={
                            "id": None, 
                            "Sil": st.column_config.CheckboxColumn("Sil", default=False, width="small"),
                            "Tarih / Saat": st.column_config.TextColumn("Tarih / Saat", disabled=True, width="medium"),
                            "Görüşme Notu": st.column_config.TextColumn("Görüşme Notu", width="large")
                        },
                        use_container_width=True, height=300
                    )
                    
                    degisiklik_log_var = False
                    for index, row in edited_log.iterrows():
                        if row["Sil"] == True:
                            degisiklik_log_var = True
                            break
                        
                        orj_not = str(df_log[df_log["id"] == row["id"]].iloc[0]["Görüşme Notu"])
                        yeni_not = str(row["Görüşme Notu"])
                        if orj_not != yeni_not:
                            degisiklik_log_var = True
                            break
                            
                    btn_log_lbl = "🔴 Not Değişikliklerini Kaydet" if degisiklik_log_var else "🟢 Geçmiş Güncel"
                    btn_log_type = "primary" if degisiklik_log_var else "secondary"
                    
                    if st.button(btn_log_lbl, type=btn_log_type, disabled=not degisiklik_log_var, key=f"btn_log_{aktif_cari_kod}", use_container_width=True):
                        c = conn.cursor()
                        for index, row in edited_log.iterrows():
                            log_id = row["id"]
                            if row["Sil"] == True:
                                c.execute("DELETE FROM loglar WHERE id=%s", (log_id,))
                            else:
                                orj_not = str(df_log[df_log["id"] == log_id].iloc[0]["Görüşme Notu"])
                                yeni_not = str(row["Görüşme Notu"])
                                if orj_not != yeni_not:
                                    c.execute("UPDATE loglar SET not_metni=%s WHERE id=%s", (yeni_not, log_id))
                        conn.commit()
                        c.close()
                        st.success("✅ Görüşme notları başarıyla güncellendi/silindi!")
                        time.sleep(1)
                        st.rerun()
                else: 
                    st.write("Bu cari için henüz bir görüşme notu veya durum değişikliği bulunmuyor.")
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
            fig_pie = px.pie(df_pie, values='bakiye', names='durum', hole=0.4, color_discrete_sequence=px.colors.qualitative.Set2)
            fig_pie.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with r_col2:
            st.markdown("#### Planlanan Nakit Akışı (Tarihe Göre)")
            df_nakit = df_rapor[(df_rapor["durum"] != "Ödedi") & (df_rapor["tarih"] != "")]
            if not df_nakit.empty:
                df_nakit["tarih_obj"] = pd.to_datetime(df_nakit["tarih"], format="%d.%m.%Y", errors='coerce')
                df_nakit = df_nakit.dropna(subset=["tarih_obj"]).sort_values("tarih_obj")
                fig_bar = px.bar(df_nakit.groupby("tarih")["bakiye"].sum().reset_index(), x='tarih', y='bakiye', text='bakiye', labels={'tarih': 'Ödeme Tarihi', 'bakiye': 'Beklenen Tutar (TL)'})
                fig_bar.update_traces(texttemplate='%{text:,.2f} TL', textposition='outside')
                st.plotly_chart(fig_bar, use_container_width=True)
            else: st.info("Tarihi planlanmış aktif bir tahsilat bulunmuyor.")
                
        st.markdown("#### 🎯 Özel Durumlara Göre Risk Analizi")
        df_ozel = df_rapor[df_rapor["ozel_durum"] != ""].groupby("ozel_durum")["bakiye"].sum().reset_index()
        if not df_ozel.empty:
            df_ozel = df_ozel.sort_values(by="bakiye", ascending=False)
            df_ozel.rename(columns={"ozel_durum": "Sizin Eklediğiniz Durum", "bakiye": "İçerideki Toplam Tutar (TL)"}, inplace=True)
            st.dataframe(df_ozel.style.format({"İçerideki Toplam Tutar (TL)": "{:,.2f} TL"}), use_container_width=True)
        else: st.write("Henüz bir 'Özel Durum' etiketi kullanmadınız.")
    else: st.info("Rapor oluşturulabilmesi için Takip listenize veri eklemelisiniz.")

# ==========================================
# 4. SEKME: GÖREV YÖNETİCİSİ
# ==========================================
with tab4:
    st.markdown("### ✅ Günlük İşler ve Görev Yöneticisi")
    
    with st.expander("➕ Yeni Görev Ekle", expanded=True):
        with st.form("yeni_gorev_form", clear_on_submit=True):
            g_col1, g_col2 = st.columns(2)
            g_adi = g_col1.text_input("Görev Adı / Yapılacak İş *")
            g_tarih = g_col2.date_input("Hedef Tarih", value=None, format="DD.MM.YYYY")
            g_not = st.text_area("Detay / Açıklama")
            
            if st.form_submit_button("Görevi Kaydet"):
                if g_adi.strip():
                    tarih_str = g_tarih.strftime("%d.%m.%Y") if g_tarih else ""
                    conn = get_db_connection()
                    c = conn.cursor()
                    c.execute("INSERT INTO gorevler (gorev_adi, tarih, durum, notlar) VALUES (%s, %s, %s, %s)",
                              (g_adi.strip(), tarih_str, "Bekliyor", g_not.strip()))
                    conn.commit()
                    c.close()
                    conn.close()
                    st.success("Görev başarıyla eklendi!")
                    time.sleep(1)
                    st.rerun()
                else:
                    st.error("Lütfen bir görev adı girin!")

    st.markdown("---")
    conn = get_db_connection()
    df_gorev = pd.read_sql_query("SELECT id, gorev_adi, tarih, durum, notlar FROM gorevler ORDER BY id DESC", conn)
    
    if not df_gorev.empty:
        df_gorev.insert(0, "Seç", False)
        df_gorev.rename(columns={"gorev_adi": "Görev Adı", "tarih": "Tarih", "durum": "Durum", "notlar": "Notlar"}, inplace=True)
        df_gorev["Tarih"] = pd.to_datetime(df_gorev["Tarih"], format="%d.%m.%Y", errors="coerce").dt.date
        
        edited_gorev = st.data_editor(
            df_gorev, hide_index=True,
            column_config={
                "id": None, 
                "Seç": st.column_config.CheckboxColumn("Seç", default=False),
                "Durum": st.column_config.SelectboxColumn("Durum", options=["Bekliyor", "Devam Ediyor", "Tamamlandı", "İptal"]),
                "Tarih": st.column_config.DateColumn("Tarih", format="DD.MM.YYYY")
            },
            disabled=["Görev Adı"],
            use_container_width=True, height=300
        )
        
        silinecek_gorevler = edited_gorev[edited_gorev["Seç"] == True]
        
        degisiklik_gorev_var = False
        for index, row in edited_gorev.iterrows():
            g_id = row["id"]
            orj_row = df_gorev[df_gorev["id"] == g_id].iloc[0]
            y_durum = str(row["Durum"])
            y_not = str(row["Notlar"]) if pd.notna(row["Notlar"]) else ""
            y_tarih = row["Tarih"].strftime("%d.%m.%Y") if pd.notna(row["Tarih"]) else ""
            
            if (orj_row["Durum"] != y_durum or str(orj_row["Notlar"]) != y_not or orj_row["Tarih"] != row["Tarih"]):
                degisiklik_gorev_var = True
                break
        
        col_gbtn1, col_gbtn2 = st.columns(2)
        with col_gbtn1:
            g_lbl = "🔴 Görevlerde Değişiklik Var! Tıkla ve Kaydet" if degisiklik_gorev_var else "🟢 Tüm Görevler Güncel"
            g_type = "primary" if degisiklik_gorev_var else "secondary"
            kaydet_g_basildi = st.button(g_lbl, type=g_type, disabled=not degisiklik_gorev_var, use_container_width=True, key="btn_gorev_kaydet")
            
        with col_gbtn2:
            sil_g_basildi = st.button("❌ Seçilen Görevleri Sil", use_container_width=True, key="btn_gorev_sil")
            
        if kaydet_g_basildi:
            c = conn.cursor()
            for index, row in edited_gorev.iterrows():
                g_id = row["id"]
                orj_row = df_gorev[df_gorev["id"] == g_id].iloc[0]
                y_durum = str(row["Durum"])
                y_not = str(row["Notlar"]) if pd.notna(row["Notlar"]) else ""
                y_tarih = row["Tarih"].strftime("%d.%m.%Y") if pd.notna(row["Tarih"]) else ""
                
                if (orj_row["Durum"] != y_durum or str(orj_row["Notlar"]) != y_not or orj_row["Tarih"] != row["Tarih"]):
                    c.execute("UPDATE gorevler SET durum=%s, notlar=%s, tarih=%s WHERE id=%s", (y_durum, y_not, y_tarih, g_id))
            conn.commit()
            c.close()
            st.success("✅ Tüm görev güncellemeleri tek seferde kaydedildi!")
            time.sleep(1)
            st.rerun()

        if sil_g_basildi:
            if not silinecek_gorevler.empty:
                c = conn.cursor()
                for index, row in silinecek_gorevler.iterrows():
                    c.execute("DELETE FROM gorevler WHERE id=%s", (row["id"],))
                conn.commit()
                c.close()
                st.success("Seçilen görevler başarıyla silindi.")
                time.sleep(1)
                st.rerun()
            else:
                st.warning("Lütfen silmek için tablodan bir görev seçin.")
                
        if len(silinecek_gorevler) == 1:
            st.markdown("---")
            secili_g_id = silinecek_gorevler.iloc[0]["id"]
            secili_g_adi = str(silinecek_gorevler.iloc[0]["Görev Adı"])
            secili_g_not = str(silinecek_gorevler.iloc[0]["Notlar"])
            
            st.info(f"📌 **{secili_g_adi}** görevini takvime ekleyebilirsiniz.")
            g_hedef_tarih = st.date_input("Görev Takvim Tarihi", format="DD.MM.YYYY", key=f"g_date_{secili_g_id}")
            
            if g_hedef_tarih:
                g_ics_data = create_ics_file(f"✅ Görev: {secili_g_adi}", secili_g_not, g_hedef_tarih.strftime("%d.%m.%Y"))
                if g_ics_data:
                    st.download_button(
                        label="🗓️ Bu Görevi iPhone Takvimine Ekle",
                        data=g_ics_data,
                        file_name=f"Gorev_{secili_g_id}.ics",
                        mime="text/calendar",
                        use_container_width=True
                    )
    else:
        st.info("Harika! Bekleyen veya tamamlanmış hiçbir göreviniz yok.")
    conn.close()
