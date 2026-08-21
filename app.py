import streamlit as st
import pandas as pd
import json
import os
import io

st.set_page_config(page_title="Tahsilat ve Cari Takip", page_icon="📈", layout="wide")

TAKIP_DOSYASI = "takip_verileri.json"
ANA_LISTE_DOSYASI = "ana_liste_verileri.json"

def bakiye_temizle(deger):
    try:
        temiz = str(deger).replace(' TL', '').replace('₺', '').strip()
        temiz = temiz.replace('.', '').replace(',', '.')
        return float(temiz)
    except ValueError:
        return 0.0

def bakiye_formatla(deger):
    try:
        return f"{float(deger):,.2f} TL".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(deger)

def verileri_yukle():
    if os.path.exists(TAKIP_DOSYASI):
        with open(TAKIP_DOSYASI, 'r', encoding='utf-8') as f: return json.load(f)
    return {}

def ana_liste_yukle():
    if os.path.exists(ANA_LISTE_DOSYASI):
        with open(ANA_LISTE_DOSYASI, 'r', encoding='utf-8') as f: return json.load(f)
    return []

def verileri_kaydet():
    with open(TAKIP_DOSYASI, 'w', encoding='utf-8') as f:
        json.dump(st.session_state.takip, f, ensure_ascii=False, indent=4)

def ana_liste_kaydet():
    with open(ANA_LISTE_DOSYASI, 'w', encoding='utf-8') as f:
        json.dump(st.session_state.ana_liste, f, ensure_ascii=False, indent=4)

if 'takip' not in st.session_state: st.session_state.takip = verileri_yukle()
if 'ana_liste' not in st.session_state: st.session_state.ana_liste = ana_liste_yukle()

# --- YAN MENÜ (SİSTEM VE YEDEKLEME) ---
with st.sidebar:
    st.header("⚙️ Sistem İşlemleri")
    st.info("Bulut sisteminde verilerinizin kaybolmaması için gün sonunda yedeğinizi indirin.")
    
    # JSON İndirme
    if st.session_state.takip:
        json_data = json.dumps(st.session_state.takip, ensure_ascii=False, indent=4)
        st.download_button(label="📥 Takip Verilerini Yedekle", data=json_data, file_name="takip_yedek.json", mime="application/json")
    
    # JSON Yükleme
    st.markdown("---")
    st.markdown("**Yedekten Geri Yükle**")
    yedek_dosya = st.file_uploader("Yedek JSON Dosyasını Seçin", type=["json"])
    if yedek_dosya is not None:
        if st.button("🔄 Yedeği Yükle"):
            st.session_state.takip = json.load(yedek_dosya)
            verileri_kaydet()
            st.success("Yedek başarıyla yüklendi!")
            st.rerun()

st.title("📈 Netsis Tahsilat ve Cari Takip Sistemi")

tab1, tab2 = st.tabs(["📋 Tüm Cariler (Ana Ekran)", "🔍 Tahsilat Takip Sayfası"])

# --- 1. SEKME ---
with tab1:
    st.markdown("### Excel'den Cari Yükle")
    uploaded_file = st.file_uploader("Netsis Excel Raporunu Seçin", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        if st.button("Verileri Aktar ve Listeyi Yenile", type="primary"):
            df = pd.read_excel(uploaded_file)
            st.session_state.ana_liste = []
            s_kod, s_isim, s_tel, s_bakiye = "Cari Kod", "Cari İsim", "Telefon", "Borç Bak."
            
            sayac = 0
            for index, row in df.iterrows():
                c_isim = str(row.get(s_isim, ""))
                if pd.isna(row.get(s_isim)) or not c_isim.strip(): continue
                c_kod = str(row.get(s_kod, "-"))
                if c_kod in st.session_state.takip: continue
                c_tel = str(row.get(s_tel, "")) if pd.notna(row.get(s_tel)) else ""
                bakiye_val = row.get(s_bakiye, 0.0) if pd.notna(row.get(s_bakiye)) else 0.0
                
                satir = {
                    "Seç": False,
                    "Cari Kod": c_kod,
                    "Cari İsim": c_isim,
                    "Telefon": c_tel,
                    "Bakiye": bakiye_formatla(bakiye_val),
                    "_Bakiye_Num": bakiye_temizle(bakiye_val)
                }
                st.session_state.ana_liste.append(satir)
                sayac += 1
                
            ana_liste_kaydet()
            st.success(f"{sayac} adet cari başarıyla yüklendi.")
            st.rerun()

    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    min_b = col1.number_input("Min Bakiye (TL)", value=0.0)
    max_b = col2.number_input("Max Bakiye (TL)", value=9999999.0)
    arama = col3.text_input("Cari İsim veya Kod Ara").lower()
    
    if st.session_state.ana_liste:
        df_ana = pd.DataFrame(st.session_state.ana_liste)
        mask = (df_ana["_Bakiye_Num"] >= min_b) & (df_ana["_Bakiye_Num"] <= max_b)
        if arama: mask = mask & (df_ana["Cari İsim"].str.lower().str.contains(arama) | df_ana["Cari Kod"].str.lower().str.contains(arama))
        
        df_gosterim = df_ana[mask].drop(columns=["_Bakiye_Num"])
        
        st.write("💡 *İpucu: Bir hücreyi kopyalamak için üzerine tıklayıp Ctrl+C yapabilirsiniz.*")
        edited_df = st.data_editor(
            df_gosterim,
            hide_index=True,
            column_config={"Seç": st.column_config.CheckboxColumn("Seç", default=False)},
            disabled=["Cari Kod", "Cari İsim", "Telefon", "Bakiye"],
            use_container_width=True,
            height=400
        )
        
        secilenler = edited_df[edited_df["Seç"] == True]
        if st.button("Seçilenleri Takibe Aktar ➔", type="primary") and not secilenler.empty:
            kalan_liste = []
            for item in st.session_state.ana_liste:
                if item["Cari Kod"] in secilenler["Cari Kod"].values:
                    st.session_state.takip[item["Cari Kod"]] = {
                        "Cari İsim": item["Cari İsim"],
                        "Telefon": item["Telefon"],
                        "Bakiye": item["Bakiye"],
                        "Durum": "Beklemede",
                        "Özel Durum": "",
                        "Tarih": "",
                        "Not": ""
                    }
                else: kalan_liste.append(item)
                    
            st.session_state.ana_liste = kalan_liste
            verileri_kaydet()
            ana_liste_kaydet()
            st.success("Seçilen cariler takibe aktarıldı!")
            st.rerun()

# --- 2. SEKME ---
with tab2:
    if st.session_state.takip:
        takip_listesi = []
        toplam_bakiye = 0.0
        odenen_bakiye = 0.0
        bekleyen_sayisi = 0
        
        for kod, veri in st.session_state.takip.items():
            bakiye_sayi = bakiye_temizle(veri.get("Bakiye", "0"))
            durum = veri.get("Durum", "Beklemede")
            
            toplam_bakiye += bakiye_sayi
            if durum == "Ödedi": odenen_bakiye += bakiye_sayi
            if durum == "Beklemede": bekleyen_sayisi += 1
            
            takip_listesi.append({
                "Seç": False,
                "Cari Kod": kod,
                "Cari İsim": veri.get("Cari İsim", ""),
                "Telefon": veri.get("Telefon", ""),
                "Bakiye": veri.get("Bakiye", ""),
                "Durum": durum,
                "Özel Durum": veri.get("Özel Durum", ""),
                "Tarih": veri.get("Tarih", ""),
                "Not": veri.get("Not", "")
            })
            
        df_takip = pd.DataFrame(takip_listesi)
        
        # DASHBOARD (Bilgi Kartları)
        st.markdown("### 📊 Genel Durum Özeti")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Takipteki Cari Sayısı", len(df_takip))
        m2.metric("Henüz Aranmayan (Beklemede)", bekleyen_sayisi)
        m3.metric("Tahsil Edilen (Ödedi)", f"{odenen_bakiye:,.2f} TL")
        m4.metric("Kalan Toplam Alacak", f"{(toplam_bakiye - odenen_bakiye):,.2f} TL")
        st.markdown("---")
        
        col_f1, col_f2 = st.columns(2)
        durumlar = ["Tümü", "Beklemede", "Arandı", "Ödedi", "Dönmedi"]
        secili_durum = col_f1.selectbox("Durum Filtresi", durumlar)
        
        ozel_durumlar = ["Tümü"] + sorted(list(set(df_takip["Özel Durum"][df_takip["Özel Durum"] != ""])))
        secili_ozel = col_f2.selectbox("Özel Durum Filtresi", ozel_durumlar)
        
        if secili_durum != "Tümü": df_takip = df_takip[df_takip["Durum"] == secili_durum]
        if secili_ozel != "Tümü": df_takip = df_takip[df_takip["Özel Durum"] == secili_ozel]
        
        st.info("✏️ **Durum, Özel Durum, Tarih ve Not** hücrelerine çift tıklayarak doğrudan tablo üzerinden düzenleme yapabilirsiniz. Değişiklikler otomatik kaydedilir.")
        
        edited_takip = st.data_editor(
            df_takip,
            hide_index=True,
            column_config={
                "Seç": st.column_config.CheckboxColumn("Silmek İçin Seç", default=False),
                "Durum": st.column_config.SelectboxColumn("Durum", options=["Beklemede", "Arandı", "Ödedi", "Dönmedi"]),
            },
            disabled=["Cari Kod", "Cari İsim", "Telefon", "Bakiye"],
            use_container_width=True,
            height=500
        )
        
        # Excel İndirme Butonu
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_takip.drop(columns=["Seç"]).to_excel(writer, index=False, sheet_name='Tahsilat_Listesi')
        st.download_button(label="📄 Güncel Listeyi Excel Olarak İndir", data=output.getvalue(), file_name="Tahsilat_Takip_Raporu.xlsx", mime="application/vnd.ms-excel")
        
        degisiklik_var = False
        for index, row in edited_takip.iterrows():
            kod = row["Cari Kod"]
            eski = st.session_state.takip[kod]
            if (eski.get("Durum", "Beklemede") != row["Durum"] or 
                eski.get("Özel Durum", "") != row["Özel Durum"] or 
                eski.get("Tarih", "") != row["Tarih"] or 
                eski.get("Not", "") != row["Not"]):
                
                st.session_state.takip[kod]["Durum"] = row["Durum"]
                st.session_state.takip[kod]["Özel Durum"] = row["Özel Durum"]
                st.session_state.takip[kod]["Tarih"] = row["Tarih"]
                st.session_state.takip[kod]["Not"] = row["Not"]
                degisiklik_var = True
                
        if degisiklik_var: verileri_kaydet()
            
        silinecekler = edited_takip[edited_takip["Seç"] == True]
        if st.button("❌ Seçilenleri Takipten Çıkar ve Ana Ekrana Aktar"):
            if not silinecekler.empty:
                for index, row in silinecekler.iterrows():
                    kod = row["Cari Kod"]
                    mevcut_kodlar = [item["Cari Kod"] for item in st.session_state.ana_liste]
                    if kod not in mevcut_kodlar:
                        st.session_state.ana_liste.append({
                            "Seç": False, "Cari Kod": kod, "Cari İsim": row["Cari İsim"],
                            "Telefon": row["Telefon"], "Bakiye": row["Bakiye"],
                            "_Bakiye_Num": bakiye_temizle(row["Bakiye"])
                        })
                    del st.session_state.takip[kod]
                verileri_kaydet()
                ana_liste_kaydet()
                st.success("Seçilen cariler takipten çıkarıldı.")
                st.rerun()
            else: st.warning("Lütfen silmek için tablodan seçim yapın.")
    else:
        st.info("Takip listeniz şu an boş. Ana ekrandan cari aktarabilirsiniz.")
