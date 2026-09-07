import streamlit as st
import pandas as pd
import datetime
import io
import calendar
from supabase import create_client, Client

st.set_page_config(page_title="Cafe Yönetim", layout="wide")

# --- SUPABASE BAĞLANTISI ---
url = st.secrets["SUPABASE_URL"]
key = st.secrets["SUPABASE_KEY"]
supabase: Client = create_client(url, key)

# --- YARDIMCI FONKSİYONLAR ---
def safe_float(val):
    if pd.isna(val): return 0.0
    try:
        return float(str(val).replace(',', '.').replace(' ', '').strip())
    except ValueError:
        return 0.0

def excel_indir(df):
    output = io.BytesIO()
    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Veriler')
        return output.getvalue(), "xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    except ModuleNotFoundError:
        return df.to_csv(index=False).encode('utf-8-sig'), "csv", "text/csv"

# --- ZAMAN DAMGASI (GELİŞTİRİLMİŞ PARSER) ---
def format_single_date(d):
    if pd.isna(d) or str(d).strip() in ["", "NaT", "nan", "None"]: 
        return "-"
    try:
        dt = pd.to_datetime(d)
        if dt.tzinfo is None:
            dt = dt.tz_localize('UTC')
        return dt.tz_convert('Europe/Istanbul').strftime('%d.%m.%Y %H:%M')
    except Exception:
        return "-"

def zaman_sutunlari_ekle(df, mevcut_sutunlar):
    ek_sutunlar = mevcut_sutunlar.copy()
    if 'created_at' in df.columns:
        df['İşlenme Zamanı'] = df['created_at'].apply(format_single_date)
        ek_sutunlar.append('İşlenme Zamanı')
    if 'updated_at' in df.columns:
        df['Düzenleme Zamanı'] = df['updated_at'].apply(format_single_date)
        ek_sutunlar.append('Düzenleme Zamanı')
    return df, ek_sutunlar

# --- GÜVENLİ VERİTABANI FONKSİYONLARI ---
def db_oku(sorgu):
    try:
        sonuc = sorgu.execute()
        return sonuc.data if sonuc.data else []
    except Exception as e:
        st.error(f"⚠️ Veri Okuma Hatası: {e}")
        return []

def db_yaz(sorgu):
    try:
        sorgu.execute()
        return True
    except Exception as e:
        st.error(f"⚠️ İşlem Başarısız Oldu. Hata Detayı: {e}")
        return False

# --- MERKEZİ BİLDİRİM SİSTEMİ ---
def bildirim_goster():
    if "genel_mesaj" in st.session_state:
        tur, metin = st.session_state["genel_mesaj"]
        if tur == "success":
            st.success(metin)
            st.toast(metin, icon="✅")
        elif tur == "error":
            st.error(metin)
            st.toast(metin, icon="❌")
        elif tur == "warning":
            st.warning(metin)
            st.toast(metin, icon="⚠️")
        elif tur == "info":
            st.info(metin)
            st.toast(metin, icon="ℹ️")
        del st.session_state["genel_mesaj"]

# --- CALLBACK FONKSİYONLARI ---
def platform_kaydet_cb(plat_adi):
    tarih = st.session_state[f"{plat_adi}_tarih"]
    online = st.session_state[f"{plat_adi}_on"]
    kapida = st.session_state[f"{plat_adi}_kap"]
    
    if online <= 0 and kapida <= 0:
        st.session_state["genel_mesaj"] = ("warning", "Lütfen en az bir tutar girin.")
        return

    mevcut = db_oku(supabase.table("platform_satis").select("id").eq("platform", plat_adi).eq("tarih", str(tarih)))
    if mevcut:
        st.session_state["genel_mesaj"] = ("error", f"⚠️ {tarih} tarihi için {plat_adi} satışı zaten girilmiş! Değiştirmek için Düzenle panelini kullanın.")
        return

    hata = False
    for o_tip, tutar in [("Online", online), ("Kapıda Ödeme", kapida)]:
        if tutar > 0:
            ayar_getir = db_oku(supabase.table("ayarlar").select("*").eq("platform", plat_adi).eq("odeme_tipi", o_tip))
            if len(ayar_getir) > 0:
                ayar = ayar_getir[0]
                
                k_tutari = round(tutar * (float(ayar['komisyon']) / 100), 2)
                s_tutari = round((tutar / 1.10) * (float(ayar['stopaj']) / 100), 2)
                kes = round(k_tutari + s_tutari, 2)
                net_t = round(tutar - kes if o_tip == "Online" else -kes, 2)
                
                t_tarihi = tarih + datetime.timedelta(days=int(ayar['vade']))
                
                veri = {
                    "tarih": str(tarih), "platform": plat_adi, "odeme_tipi": o_tip, 
                    "brut": round(tutar, 2), "komisyon_tutari": k_tutari, "stopaj_tutari": s_tutari, 
                    "kesinti": kes, "net": net_t, "tahsilat_tarihi": str(t_tarihi), "durum": "Bekliyor"
                }
                if not db_yaz(supabase.table("platform_satis").insert(veri)):
                    hata = True
            else:
                st.session_state["genel_mesaj"] = ("error", f"⚠️ Lütfen önce Ayarlar'dan '{o_tip}' için oranları kaydedin!")
                return
    if not hata:
        st.session_state["genel_mesaj"] = ("success", f"{plat_adi} satışları başarıyla kaydedildi!")
        st.session_state[f"{plat_adi}_on"] = 0.0
        st.session_state[f"{plat_adi}_kap"] = 0.0

def ciro_kaydet_cb():
    tarih = st.session_state["ciro_tarih"]
    kasa = st.session_state["ciro_kasa"]
    nakit = st.session_state["ciro_nakit"]
    kredi = st.session_state["ciro_kredi"]
    pavo_n = st.session_state["ciro_pavo_n"]
    pavo_k = st.session_state["ciro_pavo_k"]
    odenmez = st.session_state["ciro_odenmez"]
    
    mevcut = db_oku(supabase.table("ciro").select("id").eq("tarih", str(tarih)))
    if mevcut:
        st.session_state["genel_mesaj"] = ("error", f"⚠️ {tarih} tarihi için Dükkan Cirosu zaten girilmiş! Düzenlemek için paneli kullanın.")
        return
        
    veri = {"tarih": str(tarih), "kasa": kasa, "nakit": nakit, "kredi_karti": kredi, "pavo_nakit": pavo_n, "pavo_kredi": pavo_k, "odenmez": odenmez}
    if db_yaz(supabase.table("ciro").insert(veri)):
        st.session_state["genel_mesaj"] = ("success", "Dükkan Cirosu başarıyla kaydedildi!")
        st.session_state["ciro_nakit"] = 0.0
        st.session_state["ciro_kredi"] = 0.0
        st.session_state["ciro_pavo_n"] = 0.0
        st.session_state["ciro_pavo_k"] = 0.0
        st.session_state["ciro_odenmez"] = 0.0

def puantaj_kaydet_cb():
    tarih = st.session_state["puantaj_tarih"]
    isim = st.session_state["puantaj_isim"]
    durum = st.session_state["puantaj_durum"]
    mesai = st.session_state["puantaj_mesai"]
    
    mevcut = db_oku(supabase.table("puantaj").select("id").eq("personel_adi", isim).eq("tarih", str(tarih)))
    if mevcut:
        st.session_state["genel_mesaj"] = ("error", f"⚠️ {tarih} tarihinde {isim} için zaten giriş yapılmış!")
        return

    if durum == "Haftalık İzin":
        h_baslangic = tarih - datetime.timedelta(days=tarih.weekday())
        h_bitis = h_baslangic + datetime.timedelta(days=6)
        sorgu = supabase.table("puantaj").select("id").eq("personel_adi", isim).eq("durum", "Haftalık İzin").gte("tarih", str(h_baslangic)).lte("tarih", str(h_bitis))
        if db_oku(sorgu):
            st.session_state["genel_mesaj"] = ("error", f"⚠️ İŞLEM REDDEDİLDİ: {isim} bu hafta içinde zaten Haftalık İzin kullanmış!")
            return

    if durum == "Yıllık İzin":
        personel_bilgi = db_oku(supabase.table("personeller").select("yillik_izin_hakki").eq("isim", isim))
        izin_hakki = float(personel_bilgi[0].get('yillik_izin_hakki', 0)) if personel_bilgi else 0
        kullanilan_izinler = db_oku(supabase.table("puantaj").select("id").eq("personel_adi", isim).eq("durum", "Yıllık İzin"))
        if len(kullanilan_izinler) >= izin_hakki:
            st.session_state["genel_mesaj"] = ("error", f"⚠️ İŞLEM REDDEDİLDİ: Yıllık İzin hakkı kalmamıştır!")
            return
        
    veri = {"tarih": str(tarih), "personel_adi": isim, "durum": durum, "fazla_mesai_saati": mesai}
    if db_yaz(supabase.table("puantaj").insert(veri)):
        st.session_state["genel_mesaj"] = ("success", f"{isim} için puantaj kaydedildi!")
        st.session_state["puantaj_mesai"] = 0.0
        st.session_state["puantaj_durum"] = "Tam Gün"

def masraf_kaydet_cb():
    tarih = st.session_state["masraf_tarih"]
    tip = st.session_state["masraf_tipi"]
    aciklama = st.session_state["masraf_aciklama"]
    tutar = st.session_state["masraf_tutar"]
    odeme = st.session_state["masraf_odeme"]
    
    if not aciklama or tutar <= 0:
        st.session_state["genel_mesaj"] = ("warning", "Lütfen açıklama ve tutar girin.")
        return
        
    veri = {"tarih": str(tarih), "masraf_tipi": tip, "aciklama": aciklama, "tutar": tutar, "odeme_tipi": odeme}
    if db_yaz(supabase.table("masraf").insert(veri)):
        bankalar_db = db_oku(supabase.table("banka_hesaplari").select("*"))
        banka_liste = [b['isim'] for b in bankalar_db] if bankalar_db else []
        
        if odeme.startswith("Cari - "):
            cari_adi = odeme.replace("Cari - ", "")
            db_yaz(supabase.table("cari_islemler").insert({"tarih": str(tarih), "cari_adi": cari_adi, "islem_tipi": "Gelen Fatura (Bize Borç Yazar)", "tutar": tutar, "aciklama": f"Masraf: {aciklama}", "odeme_tipi": "- Yok -"}))
            st.session_state["genel_mesaj"] = ("success", f"Masraf kaydedildi ve {cari_adi} hesabına borç işlendi!")
        elif odeme in banka_liste:
            db_yaz(supabase.table("banka_islemleri").insert({"tarih": str(tarih), "hesap_adi": odeme, "islem_tipi": "Para Çıkışı (Masraf)", "tutar": tutar, "aciklama": f"Masraf: {aciklama}"}))
            st.session_state["genel_mesaj"] = ("success", f"Masraf kaydedildi ve {odeme} hesabından düşüldü!")
        else:
            st.session_state["genel_mesaj"] = ("success", "Masraf başarıyla kaydedildi!")
            
        st.session_state["masraf_aciklama"] = ""
        st.session_state["masraf_tutar"] = 0.0

def cari_islem_kaydet_cb():
    tarih = st.session_state["cari_islem_tarih"]
    cari = st.session_state["cari_islem_adi"]
    islem = st.session_state["cari_islem_tipi"]
    tutar = st.session_state["cari_islem_tutar"]
    aciklama = st.session_state["cari_islem_aciklama"]
    odeme = st.session_state.get("cari_islem_odeme", "- Yok -")
    
    if tutar <= 0:
        st.session_state["genel_mesaj"] = ("warning", "Lütfen 0'dan büyük bir tutar girin.")
        return

    if islem == "Ödeme Yaptık (Borç Düşer)" and odeme == "- Yok -":
        st.session_state["genel_mesaj"] = ("warning", "Lütfen ödemenin nereden yapıldığını seçin!")
        return
        
    veri = {"tarih": str(tarih), "cari_adi": cari, "islem_tipi": islem, "tutar": tutar, "aciklama": aciklama, "odeme_tipi": odeme if islem == "Ödeme Yaptık (Borç Düşer)" else "- Yok -"}
    
    if db_yaz(supabase.table("cari_islemler").insert(veri)):
        bankalar_db = db_oku(supabase.table("banka_hesaplari").select("*"))
        banka_liste = [b['isim'] for b in bankalar_db] if bankalar_db else []
        
        if islem == "Ödeme Yaptık (Borç Düşer)" and odeme in banka_liste:
            db_yaz(supabase.table("banka_islemleri").insert({
                "tarih": str(tarih), "hesap_adi": odeme, "islem_tipi": "Para Çıkışı", "karsi_hesap": cari, "tutar": tutar, "aciklama": f"Cari Ödemesi: {aciklama}"
            }))
            
        st.session_state["genel_mesaj"] = ("success", f"{cari} firması için {islem} kaydedildi!")
        st.session_state["cari_islem_tutar"] = 0.0
        st.session_state["cari_islem_aciklama"] = ""

# --- GİRİŞ EKRANI ---
if 'giris_yapildi' not in st.session_state:
    st.session_state.giris_yapildi = False

if not st.session_state.giris_yapildi:
    st.title("☕ Cafe Yönetim Sistemi - Giriş")
    with st.form("login_form"):
        kullanici = st.text_input("Kullanıcı Adı")
        sifre = st.text_input("Şifre", type="password")
        if st.form_submit_button("Giriş Yap"):
            k = kullanici.strip().lower()
            s = sifre.strip()
            if (k == "admin" and s == "admin123") or (k == "user" and s == "user123"):
                st.session_state.giris_yapildi = True
                st.session_state.rol = "admin" if k == "admin" else "user"
                st.rerun()
            else:
                st.error("Hatalı Giriş! Şifrenizi kontrol edin.")
    st.stop() 

# --- SOL MENÜ ---
st.sidebar.title(f"Hoşgeldin, {st.session_state.rol}")
menu = st.sidebar.radio("Menü", [
    "Adisyo (Excel) İçe Aktar",
    "Günlük Dükkan Cirosu", 
    "Yemek Sepeti Yönetimi", 
    "Trendyol Yönetimi", 
    "Banka & Kart Yönetimi",
    "Masraf Girişi", 
    "Cari (Tedarikçi) Yönetimi",
    "Kasa Yönetimi (Virman)",
    "Personel & Puantaj",
    "Raporlar"
])

if st.sidebar.button("Çıkış Yap"):
    st.session_state.giris_yapildi = False
    st.rerun()

# --- PLATFORM FONKSİYONU ---
def platform_sayfasi(platform_adi):
    st.header(f"📦 {platform_adi} Yönetimi")
    bildirim_goster()
    
    alt_menu = st.radio(
        "İşlem Seçin", 
        ["💰 Satış Girişi", "🕒 Tahsilat Takibi", "⚙️ Sisteme Öğret"], 
        horizontal=True, 
        label_visibility="collapsed",
        key=f"{platform_adi}_alt_menu"
    )
    st.markdown("---")
    
    if alt_menu == "💰 Satış Girişi":
        st.subheader("Satışları Gir")
        st.date_input("Satış Tarihi", datetime.date.today(), key=f"{platform_adi}_tarih")
        col1, col2 = st.columns(2)
        with col1: st.number_input("Online Ödeme Cirosu (₺)", min_value=0.0, key=f"{platform_adi}_on")
        with col2: st.number_input("Kapıda Ödeme Cirosu (₺)", min_value=0.0, key=f"{platform_adi}_kap")
        
        st.button("Satışları Kaydet", on_click=platform_kaydet_cb, args=(platform_adi,), type="primary")

        st.divider()
        with st.expander(f"✏️ {platform_adi} Kaydını Düzenle veya Sil", expanded=False):
            satislar_db = db_oku(supabase.table("platform_satis").select("*").eq("platform", platform_adi).order("tarih", desc=True))
            if satislar_db:
                secenekler_ps = {f"{s['tarih']} | {s.get('odeme_tipi','')} | Brüt: {s.get('brut',0)} ₺ | Durum: {s.get('durum','')} (ID: {s['id']})": s for s in satislar_db}
                secilen_ps_str = st.selectbox("İşlem Yapılacak Kaydı Seçin", ["Lütfen bir kayıt seçin..."] + list(secenekler_ps.keys()), key=f"{platform_adi}_duz_select")
                if secilen_ps_str != "Lütfen bir kayıt seçin...":
                    secilen_ps = secenekler_ps[secilen_ps_str]
                    with st.form(f"{platform_adi}_duzenle_form"):
                        try: ps_tarih = datetime.datetime.strptime(secilen_ps['tarih'], '%Y-%m-%d').date()
                        except: ps_tarih = datetime.date.today()
                        y_ps_tarih = st.date_input("Tarih", value=ps_tarih)
                        
                        odm_val = secilen_ps.get('odeme_tipi', 'Online')
                        if odm_val not in ["Online", "Kapıda Ödeme"]: odm_val = "Online"
                        y_ps_odeme = st.selectbox("Ödeme Tipi", ["Online", "Kapıda Ödeme"], index=["Online", "Kapıda Ödeme"].index(odm_val))
                        
                        y_ps_brut = st.number_input("Brüt Tutar (₺)", min_value=0.0, value=float(secilen_ps.get('brut', 0)))
                        
                        c_gun, c_sil = st.columns(2)
                        with c_gun:
                            if st.form_submit_button("Güncelle"):
                                ayar_getir = db_oku(supabase.table("ayarlar").select("*").eq("platform", platform_adi).eq("odeme_tipi", y_ps_odeme))
                                if len(ayar_getir) > 0:
                                    ayar = ayar_getir[0]
                                    k_tutari = round(y_ps_brut * (float(ayar['komisyon']) / 100), 2)
                                    s_tutari = round((y_ps_brut / 1.10) * (float(ayar['stopaj']) / 100), 2)
                                    kes = round(k_tutari + s_tutari, 2)
                                    net_t = round(y_ps_brut - kes if y_ps_odeme == "Online" else -kes, 2)
                                    t_tarihi = y_ps_tarih + datetime.timedelta(days=int(ayar['vade']))
                                    guncel_veri = {"tarih": str(y_ps_tarih), "odeme_tipi": y_ps_odeme, "brut": round(y_ps_brut, 2), "komisyon_tutari": k_tutari, "stopaj_tutari": s_tutari, "kesinti": kes, "net": net_t, "tahsilat_tarihi": str(t_tarihi)}
                                    
                                    if db_yaz(supabase.table("platform_satis").update(guncel_veri).eq("id", secilen_ps['id'])):
                                        st.session_state.genel_mesaj = ("success", "Satış kaydı güncellendi!")
                                        st.rerun()
                                else: st.error("Önce ayarları yapın.")
                        with c_sil:
                            if st.form_submit_button("🗑️ Sil"):
                                if db_yaz(supabase.table("platform_satis").delete().eq("id", secilen_ps['id'])):
                                    st.session_state.genel_mesaj = ("info", "Satış kaydı silindi!")
                                    st.rerun()

    elif alt_menu == "🕒 Tahsilat Takibi":
        bekleyenler = db_oku(supabase.table("platform_satis").select("*").eq("platform", platform_adi).eq("durum", "Bekliyor"))
        if bekleyenler:
            df = pd.DataFrame(bekleyenler)
            for col in ['brut', 'komisyon_tutari', 'stopaj_tutari', 'net', 'odeme_tipi', 'tahsilat_tarihi']:
                if col not in df.columns:
                    if col in ['odeme_tipi', 'tahsilat_tarihi']: df[col] = ""
                    else: df[col] = 0.0

            df['brut'] = pd.to_numeric(df['brut'], errors='coerce').fillna(0).round(2)
            df['komisyon_tutari'] = pd.to_numeric(df['komisyon_tutari'], errors='coerce').fillna(0).round(2)
            df['stopaj_tutari'] = pd.to_numeric(df['stopaj_tutari'], errors='coerce').fillna(0).round(2)
            df['net'] = pd.to_numeric(df['net'], errors='coerce').fillna(0).round(2)
            
            z_goster = st.toggle("⏱️ İşlenme Zamanlarını Göster", key=f"{platform_adi}_bek_zg")
            gosterim_sutunlar = ['id', 'tarih', 'odeme_tipi', 'brut', 'komisyon_tutari', 'stopaj_tutari', 'net', 'tahsilat_tarihi']
            if z_goster: df, gosterim_sutunlar = zaman_sutunlari_ekle(df, gosterim_sutunlar)
            
            df.insert(0, "Seç", False)
            gosterim_sutunlar.insert(0, "Seç")
            
            disabled_cols = ['tarih', 'odeme_tipi', 'brut', 'komisyon_tutari', 'stopaj_tutari', 'net', 'tahsilat_tarihi', 'İşlenme Zamanı', 'Düzenleme Zamanı']
            
            edited_df = st.data_editor(
                df[gosterim_sutunlar],
                column_config={"Seç": st.column_config.CheckboxColumn("Tik (Seç)", default=False), "id": None, "brut": st.column_config.NumberColumn("Brüt", format="%.2f ₺"), "komisyon_tutari": st.column_config.NumberColumn("Komisyon", format="%.2f ₺"), "stopaj_tutari": st.column_config.NumberColumn("Stopaj", format="%.2f ₺"), "net": st.column_config.NumberColumn("Net Yatan", format="%.2f ₺")},
                disabled=disabled_cols,
                hide_index=True, use_container_width=True, key=f"{platform_adi}_editor"
            )
            secilenler = edited_df[edited_df["Seç"] == True]
            col1, col2 = st.columns(2)
            with col1: st.write(f"Tüm Bekleyenlerin Toplamı: **{df['net'].sum():,.2f} ₺**")
            with col2: st.success(f"✔️ Seçtiklerinin Toplamı: **{secilenler['net'].sum():,.2f} ₺**")
            
            if not secilenler.empty:
                st.markdown("---")
                st.subheader("💳 Tahsilat ve Banka Aktarımı")
                bankalar_db = db_oku(supabase.table("banka_hesaplari").select("*"))
                banka_isimleri = [b['isim'] for b in bankalar_db] if bankalar_db else []
                
                if not banka_isimleri:
                    st.warning("⚠️ Lütfen önce 'Banka & Kart Yönetimi' sayfasından en az bir banka hesabı ekleyin.")
                else:
                    c_t1, c_t2 = st.columns(2)
                    with c_t1:
                        tahsilat_tarihi = st.date_input("Paranın Bankaya Yattığı Tarih", datetime.date.today(), key=f"{platform_adi}_tah_tar")
                    with c_t2:
                        secilen_banka = st.selectbox("Paranın Yattığı Banka Hesabı", banka_isimleri, key=f"{platform_adi}_tah_bnk")
                    
                    toplam_net = round(float(secilenler['net'].sum()), 2)
                    if st.button(f"✅ Seçili Satışları 'ÖDENDİ' Yap ve Hesaba Aktar ({toplam_net:,.2f} ₺)", type="primary", key=f"{platform_adi}_odeme_btn"):
                        for idx in secilenler["id"]:
                            db_yaz(supabase.table("platform_satis").update({"durum": "Ödendi"}).eq("id", int(idx)))
                        if toplam_net > 0:
                            db_yaz(supabase.table("banka_islemleri").insert({
                                "tarih": str(tahsilat_tarihi), "hesap_adi": secilen_banka, "islem_tipi": "Para Girişi", "tutar": toplam_net, "aciklama": f"{platform_adi} Tahsilatı"
                            }))
                        st.session_state.genel_mesaj = ("success", "Tahsilat başarıyla aktarıldı!")
                        st.rerun()
        else:
            st.success("Bekleyen alacağınız bulunmuyor.")

        odenenler = db_oku(supabase.table("platform_satis").select("*").eq("platform", platform_adi).eq("durum", "Ödendi"))
        if odenenler:
            st.divider()
            st.subheader("✅ Tahsil Edilenler")
            df_odenen = pd.DataFrame(odenenler).sort_values(by="tarih", ascending=False)
            
            for col in ['brut', 'net', 'odeme_tipi', 'tahsilat_tarihi']:
                if col not in df_odenen.columns:
                    df_odenen[col] = 0.0 if col in ['brut', 'net'] else ""

            df_odenen['net'] = pd.to_numeric(df_odenen['net'], errors='coerce').fillna(0).round(2)
            df_odenen['brut'] = pd.to_numeric(df_odenen['brut'], errors='coerce').fillna(0).round(2)
            
            z_goster2 = st.toggle("⏱️ İşlenme Zamanlarını Göster", key=f"{platform_adi}_tah_zg")
            gosterim_odn = ['tarih', 'odeme_tipi', 'brut', 'net', 'tahsilat_tarihi']
            if z_goster2: df_odenen, gosterim_odn = zaman_sutunlari_ekle(df_odenen, gosterim_odn)
            
            st.dataframe(df_odenen[gosterim_odn], hide_index=True, use_container_width=True)
            
            with st.expander("↩️ Tahsilatı Geri Al (Yanlış Aktarımlar İçin)", expanded=False):
                st.info("💡 Yanlışlıkla 'Ödendi' işaretlediğiniz kayıtları tekrar 'Bekliyor' durumuna alabilirsiniz. (Not: Bankaya yansıyan toplu tutarı 'Banka & Kart Yönetimi' sayfasından da silmeyi veya düzeltmeyi unutmayın.)")
                secenekler_o = {f"{o['tarih']} | {o.get('odeme_tipi','')} | Brüt: {o['brut']} ₺ | Net: {o['net']} ₺ (ID:{o['id']})": o for _, o in df_odenen.iterrows()}
                sec_o_str = st.selectbox("Geri Alınacak Kaydı Seçin", ["Lütfen seçin..."] + list(secenekler_o.keys()), key=f"{platform_adi}_gerial")
                if sec_o_str != "Lütfen seçin...":
                    sec_o = secenekler_o[sec_o_str]
                    if st.button("İşlemi Geri Al (Bekleyenlere Taşı)", key=f"{platform_adi}_gerial_btn"):
                        db_yaz(supabase.table("platform_satis").update({"durum": "Bekliyor"}).eq("id", sec_o['id']))
                        st.session_state.genel_mesaj = ("success", "Tahsilat başarıyla geri alındı, işlem tekrar listeye eklendi!")
                        st.rerun()

    elif alt_menu == "⚙️ Sisteme Öğret":
        st.subheader("Komisyon Ayarları")
        with st.form(f"{platform_adi}_ayar_form"):
            y_o_kom = st.number_input("Online Komisyon (%)", value=10.0)
            y_o_stop = st.number_input("Online Stopaj (%)", value=0.0)
            y_o_vade = st.number_input("Online Vade", value=1, step=1)
            y_k_kom = st.number_input("Kapıda Komisyon (%)", value=10.0)
            y_k_stop = st.number_input("Kapıda Stopaj (%)", value=0.0)
            y_k_vade = st.number_input("Kapıda Vade", value=1, step=1)
            if st.form_submit_button("Ayarları Kaydet"):
                if db_yaz(supabase.table("ayarlar").delete().eq("platform", platform_adi)):
                    db_yaz(supabase.table("ayarlar").insert([
                        {"platform": platform_adi, "odeme_tipi": "Online", "komisyon": y_o_kom, "stopaj": y_o_stop, "vade": y_o_vade},
                        {"platform": platform_adi, "odeme_tipi": "Kapıda Ödeme", "komisyon": y_k_kom, "stopaj": y_k_stop, "vade": y_k_vade}
                    ]))
                    st.session_state.genel_mesaj = ("success", "Ayarlar güncellendi!")
                    st.rerun()

# --- MENÜ İÇERİKLERİ ---
if menu == "Adisyo (Excel) İçe Aktar":
    st.header("📥 Adisyo Excel İçe Aktar (Günlük Satışlar)")
    bildirim_goster()
    
    st.info("Adisyo'dan indirdiğiniz Excel satış raporunu yükleyin. Sistem, 'Geliş Kanalı' ve belirlediğiniz ödeme yöntemlerini analiz edip Ciro ve Platform kayıtlarınızı otomatik ayırır.")
    
    uploaded_file = st.file_uploader("Adisyo Satış Raporu Yükle (Excel)", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            df_ad = pd.read_excel(uploaded_file)
            st.success("Excel başarıyla okundu! Lütfen aşağıdaki sütunları seçin:")
            cols = ["- Yok -"] + df_ad.columns.tolist()
            
            def match_col(keywords, exclude=None):
                for i, c in enumerate(cols):
                    c_lower = str(c).lower()
                    if exclude and exclude in c_lower: continue
                    for kw in keywords:
                        if kw in c_lower: return i
                return 0
                
            st.markdown("### Sütun Eşleştirme (Otomatik Tanınanları Kontrol Edin)")
            c1, c2, c3, c4 = st.columns(4)
            with c1: c_kanal = st.selectbox("Geliş Kanalı Sütunu", cols, index=match_col(["kanal", "gelis"]))
            with c2: c_nakit = st.selectbox("Nakit Ödeme Sütunu", cols, index=match_col(["nakit"], exclude="pavo"))
            with c3: c_kk = st.selectbox("Kredi Kartı Sütunu", cols, index=match_col(["kredi kartı", "kredi", "kart"], exclude="pavo"))
            with c4: c_pavo_nakit = st.selectbox("Pavo Nakit Sütunu", cols, index=match_col(["pavo nakit"]))
            
            c5, c6, c7, c8 = st.columns(4)
            with c5: c_pavo_kk = st.selectbox("Pavo Kredi Kartı Sütunu", cols, index=match_col(["pavo kredi", "pavo kart"]))
            with c6: c_ys_on = st.selectbox("YS Online Sütunu", cols, index=match_col(["ys online", "yemek"]))
            with c7: c_ty_on = st.selectbox("Trendyol Online Sütunu", cols, index=match_col(["trendyol online", "ty online"]))
            
            st.divider()
            d1, d2 = st.columns(2)
            with d1: islem_tarihi = st.date_input("İşlem Tarihi (Bu Satışlar Hangi Güne Ait?)", datetime.date.today())
            with d2: hedef_kasa = st.selectbox("Dükkan Nakitleri Hangi Kasaya Eklensin?", ["Kasa 1", "Kasa 2"])
            
            if st.button("Verileri Analiz Et ve Önizleme Oluştur", type="primary"):
                if c_kanal == "- Yok -":
                    st.error("Lütfen 'Geliş Kanalı' sütununu mutlaka seçin!")
                else:
                    c_n = 0.0; c_k = 0.0; c_pn = 0.0; c_pk = 0.0
                    ys_on = 0.0; ys_kap = 0.0
                    ty_on = 0.0; ty_kap = 0.0
                    
                    for idx, row in df_ad.iterrows():
                        kanal = str(row[c_kanal]).lower() if c_kanal != "- Yok -" else ""
                        n_val = safe_float(row[c_nakit]) if c_nakit != "- Yok -" else 0.0
                        k_val = safe_float(row[c_kk]) if c_kk != "- Yok -" else 0.0
                        pn_val = safe_float(row[c_pavo_nakit]) if c_pavo_nakit != "- Yok -" else 0.0
                        pk_val = safe_float(row[c_pavo_kk]) if c_pavo_kk != "- Yok -" else 0.0
                        ys_o_val = safe_float(row[c_ys_on]) if c_ys_on != "- Yok -" else 0.0
                        ty_o_val = safe_float(row[c_ty_on]) if c_ty_on != "- Yok -" else 0.0
                        
                        ys_on += ys_o_val
                        ty_on += ty_o_val
                        
                        c_n += n_val
                        c_k += k_val
                        c_pn += pn_val
                        c_pk += pk_val
                        
                        if "yemek sepeti" in kanal or "deliveryhero" in kanal or "ys" in kanal:
                            ys_kap += (n_val + k_val)
                        elif "trendyol" in kanal or "ty" in kanal:
                            ty_kap += (n_val + k_val)
                            
                    st.session_state['adisyo_ciro'] = [{"Tarih": str(islem_tarihi), "Kasa": hedef_kasa, "Nakit": round(c_n,2), "Kredi Kartı": round(c_k,2), "Pavo Nakit": round(c_pn,2), "Pavo Kredi": round(c_pk,2), "Ödenmez": 0.0}]
                    st.session_state['adisyo_ys'] = [{"Tarih": str(islem_tarihi), "Online Ödeme": round(ys_on,2), "Kapıda Ödeme": round(ys_kap,2)}]
                    st.session_state['adisyo_ty'] = [{"Tarih": str(islem_tarihi), "Online Ödeme": round(ty_on,2), "Kapıda Ödeme": round(ty_kap,2)}]
                    
        except Exception as e:
            st.error(f"Dosya okuma veya ayrıştırma hatası: {e}")
            
    if 'adisyo_ciro' in st.session_state:
        st.divider()
        st.subheader("📝 İşlem Önizlemesi ve Düzenleme")
        st.info("Aşağıdaki veriler Excel'den ayrıştırıldı. Gerekirse kutuların içindeki rakamlara tıklayarak elle son düzeltmeleri yapabilirsiniz.")
        
        st.write("### 🏠 Günlük Dükkan Cirosu (Platform Ödemeleri Dahil Tüm Nakit/Kartlar)")
        df_ciro_edit = st.data_editor(pd.DataFrame(st.session_state['adisyo_ciro']), hide_index=True, use_container_width=True, key="edit_ciro")
        
        c1, c2 = st.columns(2)
        with c1:
            st.write("### 🍔 Yemek Sepeti Satışları")
            df_ys_edit = st.data_editor(pd.DataFrame(st.session_state['adisyo_ys']), hide_index=True, use_container_width=True, key="edit_ys")
        with c2:
            st.write("### 🛍️ Trendyol Satışları")
            df_ty_edit = st.data_editor(pd.DataFrame(st.session_state['adisyo_ty']), hide_index=True, use_container_width=True, key="edit_ty")
            
        if st.button("✅ Tablolardaki Verileri Onayla ve Sisteme Aktar", type="primary"):
            c_row = df_ciro_edit.iloc[0]
            tar_c = str(c_row['Tarih'])
            
            if db_oku(supabase.table("ciro").select("id").eq("tarih", tar_c)):
                st.error(f"HATA: {tar_c} tarihi için Dükkan Cirosu zaten girilmiş! Eski kaydı silmeden yenisini aktaramazsınız.")
            else:
                db_yaz(supabase.table("ciro").insert({
                    "tarih": tar_c, "kasa": c_row['Kasa'], "nakit": float(c_row['Nakit']), "kredi_karti": float(c_row['Kredi Kartı']),
                    "pavo_nakit": float(c_row['Pavo Nakit']), "pavo_kredi": float(c_row['Pavo Kredi']), "odenmez": float(c_row['Ödenmez'])
                }))
                
                def plat_islet(p_adi, r_data):
                    tar_p = str(r_data['Tarih'])
                    if db_oku(supabase.table("platform_satis").select("id").eq("platform", p_adi).eq("tarih", tar_p)):
                        st.warning(f"Uyarı: {p_adi} için {tar_p} tarihinde zaten satış kaydı var, bu platform atlandı.")
                        return
                        
                    for o_tip, tutar in [("Online", float(r_data['Online Ödeme'])), ("Kapıda Ödeme", float(r_data['Kapıda Ödeme']))]:
                        if tutar > 0:
                            ayarlar = db_oku(supabase.table("ayarlar").select("*").eq("platform", p_adi).eq("odeme_tipi", o_tip))
                            if ayarlar:
                                a = ayarlar[0]
                                k_tut = round(tutar * (float(a['komisyon']) / 100), 2)
                                s_tut = round((tutar / 1.10) * (float(a['stopaj']) / 100), 2)
                                kes = round(k_tut + s_tut, 2)
                                net = round(tutar - kes if o_tip == "Online" else -kes, 2)
                                t_tar = pd.to_datetime(tar_p).date() + datetime.timedelta(days=int(a['vade']))
                                
                                db_yaz(supabase.table("platform_satis").insert({
                                    "tarih": tar_p, "platform": p_adi, "odeme_tipi": o_tip, 
                                    "brut": tutar, "komisyon_tutari": k_tut, "stopaj_tutari": s_tut, 
                                    "kesinti": kes, "net": net, "tahsilat_tarihi": str(t_tar), "durum": "Bekliyor"
                                }))
                            else:
                                st.error(f"HATA: {p_adi} ({o_tip}) için komisyon ayarı bulunamadığından sisteme aktarılamadı!")

                plat_islet("Yemek Sepeti", df_ys_edit.iloc[0])
                plat_islet("Trendyol", df_ty_edit.iloc[0])
                
                st.session_state.genel_mesaj = ("success", "Tüm Adisyo verileri başarıyla sisteme entegre edildi!")
                del st.session_state['adisyo_ciro']
                del st.session_state['adisyo_ys']
                del st.session_state['adisyo_ty']
                st.rerun()

elif menu == "Günlük Dükkan Cirosu":
    st.header("Günlük Dükkan Cirosu")
    bildirim_goster()
    
    c1, c2 = st.columns(2)
    with c1:
        st.date_input("Tarih", datetime.date.today(), key="ciro_tarih")
        st.selectbox("Hedef Kasa", ["Kasa 1", "Kasa 2"], key="ciro_kasa")
        st.number_input("Nakit (₺)", min_value=0.0, key="ciro_nakit")
        st.number_input("Kredi Kartı (₺)", min_value=0.0, key="ciro_kredi")
    with c2:
        st.number_input("Pavo Nakit (₺)", min_value=0.0, key="ciro_pavo_n")
        st.number_input("Pavo Kredi (₺)", min_value=0.0, key="ciro_pavo_k")
        st.number_input("Ödenmez (₺)", min_value=0.0, key="ciro_odenmez")
        
    st.button("Ciro Kaydet", on_click=ciro_kaydet_cb, type="primary")

    st.divider()
    with st.expander("✏️ Ciro Kaydını Düzenle veya Sil", expanded=False):
        tum_cirolar = db_oku(supabase.table("ciro").select("*").order("tarih", desc=True))
        if tum_cirolar:
            secenekler_c = {f"{c['tarih']} | {c['kasa']} | Nakit: {c.get('nakit',0)} ₺ | KK: {c.get('kredi_karti',0)} ₺ (ID: {c['id']})": c for c in tum_cirolar}
            secilen_c_str = st.selectbox("İşlem Yapılacak Ciroyu Seçin", ["Lütfen seçin..."] + list(secenekler_c.keys()))
            if secilen_c_str != "Lütfen seçin...":
                secilen_c = secenekler_c[secilen_c_str]
                with st.form("c_duz_form"):
                    try: c_tarih = datetime.datetime.strptime(secilen_c['tarih'], '%Y-%m-%d').date()
                    except: c_tarih = datetime.date.today()
                    y_c_tarih = st.date_input("Tarih", value=c_tarih)
                    y_kasa = st.selectbox("Hedef Kasa", ["Kasa 1", "Kasa 2"], index=["Kasa 1", "Kasa 2"].index(secilen_c.get('kasa', 'Kasa 1') if secilen_c.get('kasa') in ["Kasa 1", "Kasa 2"] else "Kasa 1"))
                    y_nakit = st.number_input("Nakit", value=float(secilen_c.get('nakit', 0)))
                    y_kredi = st.number_input("KK", value=float(secilen_c.get('kredi_karti', 0)))
                    y_pavo_n = st.number_input("Pavo Nakit", value=float(secilen_c.get('pavo_nakit', 0)))
                    y_pavo_k = st.number_input("Pavo KK", value=float(secilen_c.get('pavo_kredi', 0)))
                    y_odenmez = st.number_input("Ödenmez", value=float(secilen_c.get('odenmez', 0)))
                    c_gun, c_sil = st.columns(2)
                    with c_gun:
                        if st.form_submit_button("Güncelle"):
                            if db_yaz(supabase.table("ciro").update({"tarih": str(y_c_tarih), "kasa": y_kasa, "nakit": y_nakit, "kredi_karti": y_kredi, "pavo_nakit": y_pavo_n, "pavo_kredi": y_pavo_k, "odenmez": y_odenmez}).eq("id", secilen_c['id'])):
                                st.session_state.genel_mesaj = ("success", "Ciro güncellendi!")
                            st.rerun()
                    with c_sil:
                        if st.form_submit_button("Sil"):
                            if db_yaz(supabase.table("ciro").delete().eq("id", secilen_c['id'])):
                                st.session_state.genel_mesaj = ("info", "Ciro kaydı silindi!")
                            st.rerun()

    st.subheader("📋 Geçmiş Ciro Kayıtları")
    cirolar = db_oku(supabase.table("ciro").select("*").order("tarih", desc=True))
    if cirolar:
        z_goster_c = st.toggle("⏱️ Kayıt (İşlenme) Zamanlarını Göster", key="ciro_zg")
        df_ciro = pd.DataFrame(cirolar)
        for col in ['nakit', 'kredi_karti', 'pavo_nakit', 'pavo_kredi', 'odenmez']:
            if col not in df_ciro.columns: df_ciro[col] = 0.0
            df_ciro[col] = pd.to_numeric(df_ciro[col], errors='coerce').fillna(0).round(2)
            
        gosterim_ciro = ['tarih', 'kasa', 'nakit', 'kredi_karti', 'pavo_nakit', 'pavo_kredi', 'odenmez']
        if z_goster_c: df_ciro, gosterim_ciro = zaman_sutunlari_ekle(df_ciro, gosterim_ciro)
            
        st.dataframe(df_ciro[gosterim_ciro], hide_index=True, use_container_width=True)

elif menu == "Yemek Sepeti Yönetimi":
    platform_sayfasi("Yemek Sepeti")

elif menu == "Trendyol Yönetimi":
    platform_sayfasi("Trendyol")

elif menu == "Banka & Kart Yönetimi":
    st.header("💳 Banka ve Kredi Kartı Yönetimi")
    bildirim_goster()
    
    alt_menu = st.radio(
        "İşlem Seçin", 
        ["💵 İşlem Girişi", "📊 Bakiyeler ve Hesap Ekstresi", "📂 Excel İçe Aktar", "⚙️ Hesap / Kart Ekle"], 
        horizontal=True, 
        label_visibility="collapsed"
    )
    st.markdown("---")
    
    if alt_menu == "⚙️ Hesap / Kart Ekle":
        st.subheader("Sisteme Yeni Banka veya Kart Ekle")
        with st.form("banka_ekle_form"):
            b_isim = st.text_input("Hesap / Kart Adı")
            b_tip = st.selectbox("Türü", ["Banka Hesabı", "Kredi Kartı"])
            if st.form_submit_button("Ekle"):
                if b_isim.strip():
                    if db_yaz(supabase.table("banka_hesaplari").insert({"isim": b_isim.strip(), "tip": b_tip})):
                        st.session_state.genel_mesaj = ("success", "Hesap başarıyla eklendi!")
                    st.rerun()
        st.divider()
        bankalar_db = db_oku(supabase.table("banka_hesaplari").select("*"))
        if bankalar_db:
            with st.expander("✏️ Hesap Adı Düzenle veya Sil", expanded=False):
                sec_b_str = st.selectbox("İşlem Yapılacak Hesabı Seçin", ["Lütfen seçin..."] + [b['isim'] for b in bankalar_db])
                if sec_b_str != "Lütfen seçin...":
                    sec_b = next(b for b in bankalar_db if b['isim'] == sec_b_str)
                    with st.form("banka_isim_duzenle"):
                        y_isim = st.text_input("Hesap Adı", value=sec_b['isim'])
                        c_g, c_s = st.columns(2)
                        with c_g:
                            if st.form_submit_button("İsmi Güncelle"):
                                db_yaz(supabase.table("banka_hesaplari").update({"isim": y_isim}).eq("id", sec_b['id']))
                                db_yaz(supabase.table("banka_islemleri").update({"hesap_adi": y_isim}).eq("hesap_adi", sec_b['isim']))
                                db_yaz(supabase.table("banka_islemleri").update({"karsi_hesap": y_isim}).eq("karsi_hesap", sec_b['isim']))
                                st.session_state.genel_mesaj = ("success", "Hesap adı güncellendi!")
                                st.rerun()
                        with c_s:
                            if st.form_submit_button("Sil"):
                                if db_yaz(supabase.table("banka_hesaplari").delete().eq("id", sec_b['id'])):
                                    st.session_state.genel_mesaj = ("info", "Hesap sistemden silindi!")
                                st.rerun()

    elif alt_menu == "💵 İşlem Girişi":
        st.subheader("Banka veya Kredi Kartı İşlemi Ekle")
        bankalar_db = db_oku(supabase.table("banka_hesaplari").select("*"))
        if bankalar_db:
            banka_isimleri = [b['isim'] for b in bankalar_db]
            b_tarih = st.date_input("Tarih", datetime.date.today(), key="b_tar")
            c1, c2 = st.columns(2)
            with c1:
                b_hesap = st.selectbox("İşlem Yapılacak Hesap / Kart", banka_isimleri, key="b_hesap")
                b_tip = st.selectbox("İşlem Tipi", ["Açılış", "Para Girişi", "Para Çıkışı", "Bankalar Arası Virman", "Kredi Kartı Borç Ödemesi"], key="b_tip")
            with c2:
                b_tutar = st.number_input("Tutar (₺)", min_value=0.0, key="b_tut")
                b_ack = st.text_input("Açıklama", key="b_ack")
            
            b_karsi = None
            if b_tip in ["Bankalar Arası Virman", "Kredi Kartı Borç Ödemesi"]:
                diger_hesaplar = [b for b in banka_isimleri if b != st.session_state.b_hesap]
                b_karsi = st.selectbox("Karşı Hesap / Ödenen Kart", diger_hesaplar, key="b_karsi")
                
            if st.button("İşlemi Kaydet", type="primary"):
                if b_tutar > 0:
                    veri = {"tarih": str(b_tarih), "hesap_adi": st.session_state.b_hesap, "islem_tipi": b_tip, "tutar": float(b_tutar), "aciklama": b_ack}
                    if b_karsi: veri["karsi_hesap"] = b_karsi
                    if db_yaz(supabase.table("banka_islemleri").insert(veri)):
                        st.session_state.genel_mesaj = ("success", "Banka işlemi başarıyla kaydedildi!")
                    st.rerun()
                else: st.warning("Tutar 0'dan büyük olmalıdır.")
                    
            st.divider()
            with st.expander("✏️ Geçmiş İşlemi Düzenle/Sil", expanded=False):
                islemler_b = db_oku(supabase.table("banka_islemleri").select("*").order("tarih", desc=True))
                if islemler_b:
                    secenekler_bi = {f"{i['tarih']} | {i['hesap_adi']} | {i['islem_tipi']} | {i['tutar']} ₺ (ID: {i['id']})": i for i in islemler_b}
                    sec_bi_str = st.selectbox("İşlem Seçin", ["Seçiniz..."] + list(secenekler_bi.keys()))
                    if sec_bi_str != "Seçiniz...":
                        sec_bi = secenekler_bi[sec_bi_str]
                        with st.form("islem_b_duz"):
                            try: y_b_tarih = datetime.datetime.strptime(sec_bi['tarih'], '%Y-%m-%d').date()
                            except: y_b_tarih = datetime.date.today()
                            y_b_tarih_val = st.date_input("Tarih", value=y_b_tarih)
                            y_b_tutar = st.number_input("Tutar", value=float(sec_bi['tutar']))
                            y_b_ack = st.text_input("Açıklama", value=sec_bi.get('aciklama', ''))
                            cg, cs = st.columns(2)
                            with cg:
                                if st.form_submit_button("Güncelle"):
                                    if db_yaz(supabase.table("banka_islemleri").update({"tarih": str(y_b_tarih_val), "tutar": y_b_tutar, "aciklama": y_b_ack}).eq("id", sec_bi['id'])):
                                        st.session_state.genel_mesaj = ("success", "Banka işlemi güncellendi!")
                                    st.rerun()
                            with cs:
                                if st.form_submit_button("Sil"):
                                    if db_yaz(supabase.table("banka_islemleri").delete().eq("id", sec_bi['id'])):
                                        st.session_state.genel_mesaj = ("info", "Banka işlemi silindi!")
                                    st.rerun()

    elif alt_menu == "📊 Bakiyeler ve Hesap Ekstresi":
        st.subheader("Güncel Hesap Bakiyeleri ve Kart Borçları")
        islemler_b = db_oku(supabase.table("banka_islemleri").select("*"))
        hesaplar_db = db_oku(supabase.table("banka_hesaplari").select("*"))
        
        banka_isimleri_tam = []
        if hesaplar_db and islemler_b:
            banka_isimleri_tam = [b['isim'] for b in hesaplar_db]
            hesap_sozluk = {b['isim']: {'Tip': b['tip'], 'Bakiye (Eksi İse Borç)': 0.0} for b in hesaplar_db}
            for i in islemler_b:
                h = i['hesap_adi']
                tip = i['islem_tipi']
                tut = float(i['tutar'])
                kh = i.get('karsi_hesap')
                if h in hesap_sozluk:
                    if tip in ["Açılış", "Para Girişi"]: hesap_sozluk[h]['Bakiye (Eksi İse Borç)'] += tut
                    elif tip in ["Para Çıkışı", "Para Çıkışı (Masraf)", "Bankalar Arası Virman", "Kredi Kartı Borç Ödemesi"]: hesap_sozluk[h]['Bakiye (Eksi İse Borç)'] -= tut
                if kh and kh in hesap_sozluk:
                    if tip == "Bankalar Arası Virman": hesap_sozluk[kh]['Bakiye (Eksi İse Borç)'] += tut
                    elif tip == "Kredi Kartı Borç Ödemesi": hesap_sozluk[kh]['Bakiye (Eksi İse Borç)'] += tut 
            
            df_bakiye = pd.DataFrame.from_dict(hesap_sozluk, orient='index').reset_index().rename(columns={'index': 'Hesap / Kart Adı'})
            df_bakiye['Bakiye (Eksi İse Borç)'] = df_bakiye['Bakiye (Eksi İse Borç)'].round(2)
            st.dataframe(df_bakiye, hide_index=True, use_container_width=True)
            
            st.divider()
            st.subheader("🧾 Hesap Ekstresi (Bakiyeli Rapor)")
            z_goster_e = st.toggle("⏱️ Kayıt (İşlenme) Zamanlarını Göster", key="banka_eks_zg")
            
            if banka_isimleri_tam:
                secili_ekstre_hesabi = st.selectbox("Ekstresini Görmek İstediğiniz Hesabı Seçin", ["Lütfen seçin..."] + banka_isimleri_tam)
                if secili_ekstre_hesabi != "Lütfen seçin...":
                    hesap_hareketleri = []
                    for i in islemler_b:
                        if i['hesap_adi'] == secili_ekstre_hesabi:
                            if i['islem_tipi'] in ["Açılış", "Para Girişi"]:
                                hesap_hareketleri.append({"id": i['id'], "tarih": i['tarih'], "islem": i['islem_tipi'], "aciklama": i.get('aciklama',''), "karsi_hesap": i.get('karsi_hesap',''), "Giriş": float(i['tutar']), "Çıkış": 0.0, "created_at": i.get('created_at'), "updated_at": i.get('updated_at')})
                            else:
                                hesap_hareketleri.append({"id": i['id'], "tarih": i['tarih'], "islem": i['islem_tipi'], "aciklama": i.get('aciklama',''), "karsi_hesap": i.get('karsi_hesap',''), "Giriş": 0.0, "Çıkış": float(i['tutar']), "created_at": i.get('created_at'), "updated_at": i.get('updated_at')})
                        elif i.get('karsi_hesap') == secili_ekstre_hesabi and i['islem_tipi'] in ["Bankalar Arası Virman", "Kredi Kartı Borç Ödemesi"]:
                            hesap_hareketleri.append({"id": i['id'], "tarih": i['tarih'], "islem": f"{i['islem_tipi']} (Gelen)", "aciklama": i.get('aciklama',''), "karsi_hesap": i['hesap_adi'], "Giriş": float(i['tutar']), "Çıkış": 0.0, "created_at": i.get('created_at'), "updated_at": i.get('updated_at')})
                            
                    if hesap_hareketleri:
                        df_e = pd.DataFrame(hesap_hareketleri)
                        df_e['tarih_dt'] = pd.to_datetime(df_e['tarih'])
                        df_e['is_acilis'] = df_e['islem'].apply(lambda x: 0 if x == 'Açılış' else 1)
                        df_e = df_e.sort_values(by=["tarih_dt", "is_acilis", "id"], ascending=[True, True, True]).reset_index(drop=True)
                        
                        bakiye_list = []
                        bakiye = 0.0
                        for idx, r in df_e.iterrows():
                            bakiye = round(bakiye + r['Giriş'] - r['Çıkış'], 2)
                            bakiye_list.append(bakiye)
                        
                        df_e['Bakiye'] = bakiye_list
                        df_e['tarih'] = df_e['tarih_dt'].dt.date
                        df_e = df_e.sort_values(by=["tarih_dt", "is_acilis", "id"], ascending=[False, False, False]).drop(columns=['tarih_dt', 'is_acilis', 'id'])
                        
                        df_e['Giriş'] = df_e['Giriş'].round(2)
                        df_e['Çıkış'] = df_e['Çıkış'].round(2)
                        
                        gosterim_ekstre = ['tarih', 'islem', 'karsi_hesap', 'aciklama', 'Giriş', 'Çıkış', 'Bakiye']
                        if z_goster_e: df_e, gosterim_ekstre = zaman_sutunlari_ekle(df_e, gosterim_ekstre)
                        
                        st.dataframe(df_e[gosterim_ekstre], hide_index=True, use_container_width=True)
                        dosya_e, uzanti_e, mime_e = excel_indir(df_e[gosterim_ekstre])
                        st.download_button(f"📥 {secili_ekstre_hesabi} Ekstresini İndir", data=dosya_e, file_name=f"{secili_ekstre_hesabi}_Ekstresi.{uzanti_e}", mime=mime_e, key="dl_ekstre")
                    else:
                        st.info("Bu hesaba ait kayıt bulunmuyor.")

            st.divider()
            st.subheader("📋 Tüm Banka ve Kart Hareketleri (Genel Döküm)")
            z_goster_b = st.toggle("⏱️ Kayıt (İşlenme) Zamanlarını Göster", key="banka_genel_zg")
            
            banka_dokum_genel = []
            for i in islemler_b:
                banka_dokum_genel.append({
                    "tarih": i['tarih'], 
                    "hesap_adi": i['hesap_adi'], 
                    "islem_tipi": i['islem_tipi'], 
                    "karsi_hesap": i.get('karsi_hesap', ''), 
                    "tutar": float(i['tutar']), 
                    "aciklama": i.get('aciklama', ''),
                    "created_at": i.get('created_at'),
                    "updated_at": i.get('updated_at')
                })
                    
            df_islem_b = pd.DataFrame(banka_dokum_genel)
            if not df_islem_b.empty:
                df_islem_b['tutar'] = df_islem_b['tutar'].round(2)
                df_islem_b['tarih'] = pd.to_datetime(df_islem_b['tarih']).dt.date
                
                with st.expander("🔍 Genel Dökümü Filtrele", expanded=True):
                    c1, c2, c3, c4 = st.columns(4)
                    with c1: t_aralik_b = st.date_input("Tarih Aralığı", [df_islem_b['tarih'].min(), df_islem_b['tarih'].max()], key="filt_b_tar")
                    with c2: sec_hesap = st.multiselect("Hesap / Kart Seç", df_islem_b['hesap_adi'].unique().tolist(), key="filt_b_hesap")
                    with c3: sec_islem_b = st.multiselect("İşlem Tipi", df_islem_b['islem_tipi'].unique().tolist(), key="filt_b_tip")
                    with c4: ara_b = st.text_input("Açıklama Ara", key="filt_b_ara")
                
                if len(t_aralik_b) == 2: df_islem_b = df_islem_b[(df_islem_b['tarih'] >= t_aralik_b[0]) & (df_islem_b['tarih'] <= t_aralik_b[1])]
                elif len(t_aralik_b) == 1: df_islem_b = df_islem_b[df_islem_b['tarih'] == t_aralik_b[0]]
                
                if sec_hesap: df_islem_b = df_islem_b[df_islem_b['hesap_adi'].isin(sec_hesap)]
                if sec_islem_b: df_islem_b = df_islem_b[df_islem_b['islem_tipi'].isin(sec_islem_b)]
                if ara_b: df_islem_b = df_islem_b[df_islem_b['aciklama'].str.contains(ara_b, case=False, na=False)]

                gosterim_b = ['tarih', 'hesap_adi', 'islem_tipi', 'karsi_hesap', 'tutar', 'aciklama']
                if z_goster_b: df_islem_b, gosterim_b = zaman_sutunlari_ekle(df_islem_b, gosterim_b)

                st.dataframe(df_islem_b[gosterim_b].sort_values("tarih", ascending=False), hide_index=True, use_container_width=True)
                st.info(f"📊 Ekranda filtrelenen toplam işlem sayısı: **{len(df_islem_b)}** | Toplam Tutar: **{df_islem_b['tutar'].sum():,.2f} ₺**")
                
                dosya_b, uzanti_b, mime_b = excel_indir(df_islem_b[gosterim_b])
                st.download_button("📥 Filtrelenmiş Dökümü Excel'e İndir", data=dosya_b, file_name=f"Tum_Banka_Hareketleri.{uzanti_b}", mime=mime_b, key="dl_banka")

    elif alt_menu == "📂 Excel İçe Aktar":
        st.subheader("📂 Banka Ekstresi (Excel) İçe Aktar ve Öğret")
        
        excel_menu = st.radio(
            "İşlem Seçin", 
            ["📤 Excel Yükle ve Aktar", "🧠 Kelime Kuralları (Öğret)"], 
            horizontal=True, 
            label_visibility="collapsed",
            key="banka_excel_menu"
        )
        st.markdown("---")
        
        if excel_menu == "🧠 Kelime Kuralları (Öğret)":
            tipler_db = db_oku(supabase.table("masraf_tipleri").select("*"))
            masraf_tipleri = [t['tip_adi'] for t in tipler_db] if tipler_db else ["Genel Masraf"]
            cariler_db = db_oku(supabase.table("cariler").select("*"))
            cari_liste = [c['isim'] for c in cariler_db] if cariler_db else []
            
            with st.form("kural_ekle_form"):
                k_kelime = st.text_input("Açıklamada Geçen Kelime (Örn: BİM, POS)")
                k_islem = st.selectbox("Bunu Hangi İşlem Olarak Tanısın?", ["Masraf", "Cari Ödeme (Para Çıkışı)", "Para Girişi", "Para Çıkışı"])
                k_hedef = st.selectbox("Hedef / Alt Kategori (Sadece Masraf ve Cari İçin)", ["- Yok -"] + masraf_tipleri + cari_liste)
                if st.form_submit_button("Kuralı Öğret"):
                    if k_kelime.strip():
                        if db_yaz(supabase.table("banka_kurallari").insert({"kelime": k_kelime.strip(), "islem_tipi": k_islem, "hedef": k_hedef})):
                            st.session_state.genel_mesaj = ("success", "Kural başarıyla öğretildi!")
                        st.rerun()
                        
            st.divider()
            kurallar_db = db_oku(supabase.table("banka_kurallari").select("*"))
            if kurallar_db:
                for k in kurallar_db:
                    col1, col2 = st.columns([4, 1])
                    with col1: st.write(f"Kelime: **{k['kelime']}** ➡️ İşlem: **{k['islem_tipi']}** | Hedef: **{k['hedef']}**")
                    with col2:
                        if st.button("Sil", key=f"del_k_{k['id']}"):
                            if db_yaz(supabase.table("banka_kurallari").delete().eq("id", k['id'])):
                                st.session_state.genel_mesaj = ("info", "Kural silindi!")
                            st.rerun()

        elif excel_menu == "📤 Excel Yükle ve Aktar":
            bankalar_db = db_oku(supabase.table("banka_hesaplari").select("*"))
            banka_isimleri = [b['isim'] for b in bankalar_db] if bankalar_db else []
            
            if not banka_isimleri:
                st.warning("Lütfen önce sisteme bir banka hesabı ekleyin.")
            else:
                h_secim = st.selectbox("Hangi Hesaba Aktarılacak?", banka_isimleri)
                uploaded_file = st.file_uploader("Banka Ekstresi Yükle (Excel veya CSV)", type=["xlsx", "xls", "csv"])
                
                if uploaded_file is not None:
                    try:
                        if uploaded_file.name.endswith(".csv"): df_yuk = pd.read_csv(uploaded_file)
                        else: df_yuk = pd.read_excel(uploaded_file)
                        
                        cols = ["Yok"] + df_yuk.columns.tolist()
                        c1, c2, c3, c4 = st.columns(4)
                        with c1: col_tar = st.selectbox("Tarih Sütunu", cols, index=1 if len(cols)>1 else 0)
                        with c2: col_ack = st.selectbox("Açıklama Sütunu", cols, index=2 if len(cols)>2 else 0)
                        with c3: col_cik = st.selectbox("Çıkan Tutar (Borç)", cols)
                        with c4: col_gir = st.selectbox("Giren Tutar (Alacak)", cols)
                        
                        if st.button("Verileri İncele ve Kuralları Uygula"):
                            if col_tar == "Yok" or col_ack == "Yok":
                                st.error("Tarih ve Açıklama sütunları zorunludur!")
                            else:
                                kurallar_db = db_oku(supabase.table("banka_kurallari").select("*"))
                                preview = []
                                for i, row in df_yuk.iterrows():
                                    t_val = row[col_tar]
                                    ack_val = str(row[col_ack])
                                    g_tutar = safe_float(row[col_gir]) if col_gir != "Yok" else 0.0
                                    c_tutar = safe_float(row[col_cik]) if col_cik != "Yok" else 0.0
                                    
                                    if g_tutar > 0:
                                        tutar = g_tutar
                                        i_tip = "Para Girişi"
                                    elif c_tutar > 0:
                                        tutar = c_tutar
                                        i_tip = "Para Çıkışı"
                                    else: continue
                                        
                                    h_def = ""
                                    if kurallar_db:
                                        for kr in kurallar_db:
                                            if str(kr['kelime']).lower() in ack_val.lower():
                                                i_tip = kr['islem_tipi']
                                                h_def = kr['hedef']
                                                break
                                    try: tar_str = str(pd.to_datetime(t_val).date())
                                    except: tar_str = str(datetime.date.today())
                                    
                                    preview.append({"İşle": True, "Tarih": tar_str, "Açıklama": ack_val, "Tutar": tutar, "İşlem Tipi": i_tip, "Hedef Kategori": h_def if h_def != "- Yok -" else ""})
                                st.session_state['excel_preview'] = preview

                    except Exception as e:
                        st.error(f"Dosya okuma hatası: {e}")
                        
                if 'excel_preview' in st.session_state:
                    df_p = pd.DataFrame(st.session_state['excel_preview'])
                    edited_df = st.data_editor(df_p, column_config={"İşle": st.column_config.CheckboxColumn("İşle", default=True), "İşlem Tipi": st.column_config.SelectboxColumn("İşlem Tipi", options=["Para Girişi", "Para Çıkışı", "Masraf", "Cari Ödeme (Para Çıkışı)"])}, hide_index=True, use_container_width=True)
                    
                    if st.button("✅ Seçili İşlemleri Sisteme Kaydet", type="primary"):
                        secilen_rows = edited_df[edited_df['İşle'] == True]
                        basarili = 0
                        for idx, r in secilen_rows.iterrows():
                            tip = r['İşlem Tipi']
                            tar = r['Tarih']
                            tut = float(r['Tutar'])
                            ack = r['Açıklama']
                            hedef = r['Hedef Kategori']
                            
                            if tip == "Masraf":
                                db_yaz(supabase.table("masraf").insert({"tarih": tar, "masraf_tipi": hedef, "aciklama": ack, "tutar": tut, "odeme_tipi": h_secim}))
                                db_yaz(supabase.table("banka_islemleri").insert({"tarih": tar, "hesap_adi": h_secim, "islem_tipi": "Para Çıkışı (Masraf)", "tutar": tut, "aciklama": f"Masraf: {ack}"}))
                            elif tip == "Cari Ödeme (Para Çıkışı)":
                                db_yaz(supabase.table("cari_islemler").insert({"tarih": tar, "cari_adi": hedef, "islem_tipi": "Ödeme Yaptık (Borç Düşer)", "tutar": tut, "aciklama": f"Banka: {ack}", "odeme_tipi": h_secim}))
                                db_yaz(supabase.table("banka_islemleri").insert({"tarih": tar, "hesap_adi": h_secim, "islem_tipi": "Para Çıkışı", "karsi_hesap": hedef, "tutar": tut, "aciklama": f"Cari Ödemesi: {ack}"}))
                            else:
                                db_yaz(supabase.table("banka_islemleri").insert({"tarih": tar, "hesap_adi": h_secim, "islem_tipi": tip, "tutar": tut, "aciklama": ack}))
                            
                            basarili += 1
                                
                        del st.session_state['excel_preview']
                        st.session_state.genel_mesaj = ("success", f"{basarili} işlem başarıyla kaydedildi!")
                        st.rerun()

elif menu == "Masraf Girişi":
    st.header("Masraf Girişi")
    bildirim_goster()
    tipler_db = db_oku(supabase.table("masraf_tipleri").select("*"))
    tipler = [t['tip_adi'] for t in tipler_db] if tipler_db else ["Genel Masraf"]

    with st.expander("⚙️ Yeni Masraf Tipi Tanımla", expanded=False):
        with st.form("masraf_tipi_form"):
            yeni_tip = st.text_input("Masraf Tipi Adı")
            if st.form_submit_button("Ekle"):
                if yeni_tip.strip():
                    if db_yaz(supabase.table("masraf_tipleri").insert({"tip_adi": yeni_tip.strip()})):
                        st.session_state.genel_mesaj = ("success", "Masraf Tipi eklendi!")
                    st.rerun()
    st.divider()
        
    bankalar_db = db_oku(supabase.table("banka_hesaplari").select("*"))
    banka_liste = [b['isim'] for b in bankalar_db] if bankalar_db else []
    cariler_db = db_oku(supabase.table("cariler").select("*"))
    cari_liste = [f"Cari - {c['isim']}" for c in cariler_db] if cariler_db else []
    odeme_yontemleri = ["Nakit - Kasa 1", "Nakit - Kasa 2"] + banka_liste + cari_liste

    c1, c2 = st.columns(2)
    with c1:
        st.date_input("Tarih", datetime.date.today(), key="masraf_tarih")
        st.selectbox("Masraf Tipi", tipler, key="masraf_tipi")
        st.text_input("Açıklama", key="masraf_aciklama")
    with c2:
        st.number_input("Tutar (₺)", min_value=0.0, key="masraf_tutar")
        st.selectbox("Nereden Ödendi?", odeme_yontemleri, key="masraf_odeme")
    st.button("Masrafı Kaydet", on_click=masraf_kaydet_cb, type="primary")

    st.divider()
    with st.expander("✏️ Masraf Düzenle veya Sil", expanded=False):
        tum_masraflar = db_oku(supabase.table("masraf").select("*").order("tarih", desc=True))
        if tum_masraflar:
            secenekler = {f"{m['tarih']} | {m.get('masraf_tipi','Genel')} | {m.get('aciklama','')} | {m['tutar']} ₺ (ID: {m['id']})": m for m in tum_masraflar}
            secilen_m_str = st.selectbox("İşlem Yapılacak Masrafı Seçin", ["Lütfen seçin..."] + list(secenekler.keys()))
            if secilen_m_str != "Lütfen seçin...":
                secilen_m = secenekler[secilen_m_str]
                with st.form("masraf_duz_form"):
                    try: m_tarih = datetime.datetime.strptime(secilen_m['tarih'], '%Y-%m-%d').date()
                    except: m_tarih = datetime.date.today()
                    y_tarih = st.date_input("Tarih", value=m_tarih)
                    try: t_idx = tipler.index(secilen_m.get('masraf_tipi', 'Genel Masraf'))
                    except: t_idx = 0
                    y_tip = st.selectbox("Masraf Tipi", tipler, index=t_idx)
                    y_aciklama = st.text_input("Açıklama", value=secilen_m.get('aciklama',''))
                    y_tutar = st.number_input("Tutar (₺)", value=float(secilen_m.get('tutar', 0)))
                    
                    o_val = secilen_m.get('odeme_tipi', 'Nakit - Kasa 1')
                    if o_val not in odeme_yontemleri: o_val = odeme_yontemleri[0]
                    y_odeme_y = st.selectbox("Nereden Ödendi?", odeme_yontemleri, index=odeme_yontemleri.index(o_val))
                    
                    c_gun, c_sil = st.columns(2)
                    with c_gun:
                        if st.form_submit_button("Güncelle"):
                            if str(secilen_m.get('odeme_tipi','')).startswith("Cari - "):
                                eski_c_adi = str(secilen_m.get('odeme_tipi')).replace("Cari - ", "")
                                db_yaz(supabase.table("cari_islemler").delete().eq("cari_adi", eski_c_adi).eq("tarih", str(secilen_m['tarih'])).ilike("aciklama", f"Masraf: {secilen_m.get('aciklama','')}%"))
                            elif str(secilen_m.get('odeme_tipi','')) in banka_liste:
                                db_yaz(supabase.table("banka_islemleri").delete().eq("hesap_adi", secilen_m['odeme_tipi']).eq("tarih", str(secilen_m['tarih'])).eq("islem_tipi", "Para Çıkışı (Masraf)").ilike("aciklama", f"Masraf: {secilen_m.get('aciklama','')}%"))
                            
                            db_yaz(supabase.table("masraf").update({"tarih": str(y_tarih), "masraf_tipi": y_tip, "aciklama": y_aciklama, "tutar": y_tutar, "odeme_tipi": y_odeme_y}).eq("id", secilen_m['id']))
                            
                            if str(y_odeme_y).startswith("Cari - "):
                                yeni_c_adi = y_odeme_y.replace("Cari - ", "")
                                db_yaz(supabase.table("cari_islemler").insert({"tarih": str(y_tarih), "cari_adi": yeni_c_adi, "islem_tipi": "Gelen Fatura (Bize Borç Yazar)", "tutar": y_tutar, "aciklama": f"Masraf: {y_aciklama}"}))
                            elif y_odeme_y in banka_liste:
                                db_yaz(supabase.table("banka_islemleri").insert({"tarih": str(y_tarih), "hesap_adi": y_odeme_y, "islem_tipi": "Para Çıkışı (Masraf)", "tutar": y_tutar, "aciklama": f"Masraf: {y_aciklama}"}))
                            
                            st.session_state.genel_mesaj = ("success", "Masraf güncellendi!")
                            st.rerun()
                    with c_sil:
                        if st.form_submit_button("Sil"):
                            if db_yaz(supabase.table("masraf").delete().eq("id", secilen_m['id'])):
                                if str(secilen_m.get('odeme_tipi','')).startswith("Cari - "):
                                    sil_c_adi = str(secilen_m.get('odeme_tipi')).replace("Cari - ", "")
                                    db_yaz(supabase.table("cari_islemler").delete().eq("cari_adi", sil_c_adi).eq("tarih", str(secilen_m['tarih'])).ilike("aciklama", f"Masraf: {secilen_m.get('aciklama','')}%"))
                                elif str(secilen_m.get('odeme_tipi','')) in banka_liste:
                                    db_yaz(supabase.table("banka_islemleri").delete().eq("hesap_adi", secilen_m['odeme_tipi']).eq("tarih", str(secilen_m['tarih'])).eq("islem_tipi", "Para Çıkışı (Masraf)").ilike("aciklama", f"Masraf: {secilen_m.get('aciklama','')}%"))
                                st.session_state.genel_mesaj = ("info", "Masraf tamamen silindi!")
                                st.rerun()

    st.subheader("📋 Masraf Kayıtları ve Filtreleme")
    masraflar = db_oku(supabase.table("masraf").select("*").order("tarih", desc=True))
    if masraflar:
        z_goster_m = st.toggle("⏱️ Kayıt (İşlenme) Zamanlarını Göster", key="masraf_zg")
        df_masraf = pd.DataFrame(masraflar)
        if 'tutar' not in df_masraf.columns: df_masraf['tutar'] = 0.0
        if 'odeme_tipi' not in df_masraf.columns: df_masraf['odeme_tipi'] = ""
        if 'aciklama' not in df_masraf.columns: df_masraf['aciklama'] = ""
        
        df_masraf['tutar'] = pd.to_numeric(df_masraf['tutar'], errors='coerce').fillna(0).round(2)
        df_masraf['masraf_tipi'] = df_masraf.get('masraf_tipi', 'Genel').fillna('Genel Masraf')
        df_masraf['tarih'] = pd.to_datetime(df_masraf['tarih']).dt.date
        with st.expander("🔍 Filtreleme Seçenekleri", expanded=True):
            c1, c2, c3, c4 = st.columns(4)
            with c1: tarih_araligi = st.date_input("Tarih Aralığı", [df_masraf['tarih'].min(), df_masraf['tarih'].max()])
            with c2: sec_tip = st.multiselect("Masraf Tipi", df_masraf['masraf_tipi'].unique().tolist())
            with c3: sec_odeme = st.multiselect("Nereden Ödendi?", df_masraf['odeme_tipi'].unique().tolist())
            with c4: aranan = st.text_input("Açıklama Ara")
        if len(tarih_araligi) == 2: df_masraf = df_masraf[(df_masraf['tarih'] >= tarih_araligi[0]) & (df_masraf['tarih'] <= tarih_araligi[1])]
        elif len(tarih_araligi) == 1: df_masraf = df_masraf[df_masraf['tarih'] == tarih_araligi[0]]
        if sec_tip: df_masraf = df_masraf[df_masraf['masraf_tipi'].isin(sec_tip)]
        if sec_odeme: df_masraf = df_masraf[df_masraf['odeme_tipi'].isin(sec_odeme)]
        if aranan: df_masraf = df_masraf[df_masraf['aciklama'].str.contains(aranan, case=False, na=False)]
        
        gosterim_mas = ['tarih', 'masraf_tipi', 'aciklama', 'tutar', 'odeme_tipi']
        if z_goster_m: df_masraf, gosterim_mas = zaman_sutunlari_ekle(df_masraf, gosterim_mas)
        
        st.dataframe(df_masraf[gosterim_mas], hide_index=True, use_container_width=True)
        st.info(f"📊 Toplam Tutar: **{df_masraf['tutar'].sum():,.2f} ₺**")

elif menu == "Cari (Tedarikçi) Yönetimi":
    st.header("🏢 Cari (Tedarikçi) Yönetimi")
    bildirim_goster()
    
    alt_menu = st.radio(
        "İşlem Seçin", 
        ["📈 Fatura & Ödeme Girişi", "📋 Cari Ekstre", "⚙️ Tedarikçi Ekle"], 
        horizontal=True, 
        label_visibility="collapsed"
    )
    st.markdown("---")
    
    if alt_menu == "⚙️ Tedarikçi Ekle":
        with st.form("cari_ekle_form"):
            yeni_cari = st.text_input("Tedarikçi Firma / Kişi Adı")
            if st.form_submit_button("Ekle"):
                if yeni_cari.strip():
                    if db_yaz(supabase.table("cariler").insert({"isim": yeni_cari.strip()})):
                        st.session_state.genel_mesaj = ("success", "Cari eklendi!")
                    st.rerun()
        st.divider()
        cariler_db = db_oku(supabase.table("cariler").select("*"))
        if cariler_db:
            with st.expander("✏️ Cari Adı Düzenle veya Sil", expanded=False):
                sec_c_str = st.selectbox("İşlem Yapılacak Cariyi Seçin", ["Lütfen seçin..."] + [c['isim'] for c in cariler_db])
                if sec_c_str != "Lütfen seçin...":
                    sec_c = next(c for c in cariler_db if c['isim'] == sec_c_str)
                    with st.form("cari_isim_duzenle"):
                        y_isim = st.text_input("Firma Adı", value=sec_c['isim'])
                        c_g, c_s = st.columns(2)
                        with c_g:
                            if st.form_submit_button("Güncelle"):
                                db_yaz(supabase.table("cariler").update({"isim": y_isim}).eq("id", sec_c['id']))
                                db_yaz(supabase.table("cari_islemler").update({"cari_adi": y_isim}).eq("cari_adi", sec_c['isim']))
                                st.session_state.genel_mesaj = ("success", "Cari güncellendi!")
                                st.rerun()
                        with c_s:
                            if st.form_submit_button("Sil"):
                                if db_yaz(supabase.table("cariler").delete().eq("id", sec_c['id'])):
                                    st.session_state.genel_mesaj = ("info", "Cari silindi!")
                                st.rerun()

    elif alt_menu == "📈 Fatura & Ödeme Girişi":
        cariler_db = db_oku(supabase.table("cariler").select("*"))
        if cariler_db:
            c1, c2 = st.columns(2)
            with c1:
                st.date_input("İşlem Tarihi", datetime.date.today(), key="cari_islem_tarih")
                st.selectbox("Cari (Firma) Seçin", [c['isim'] for c in cariler_db], key="cari_islem_adi")
                st.selectbox("İşlem Tipi", ["Gelen Fatura (Bize Borç Yazar)", "Ödeme Yaptık (Borç Düşer)"], key="cari_islem_tipi")
            with c2:
                st.number_input("Tutar (₺)", min_value=0.0, key="cari_islem_tutar")
                
                bankalar_db = db_oku(supabase.table("banka_hesaplari").select("*"))
                b_liste = [b['isim'] for b in bankalar_db] if bankalar_db else []
                odeme_yontemleri = ["- Yok -", "Nakit - Kasa 1", "Nakit - Kasa 2"] + b_liste
                st.selectbox("Nereden Ödendi? (Sadece Ödeme Yaptıysanız Seçin)", odeme_yontemleri, key="cari_islem_odeme")
                
                st.text_input("Açıklama / Fatura No", key="cari_islem_aciklama")
                st.button("İşlemi Kaydet", on_click=cari_islem_kaydet_cb, type="primary")
            st.divider()
            with st.expander("✏️ Geçmiş İşlemi Düzenle/Sil", expanded=False):
                islemler = db_oku(supabase.table("cari_islemler").select("*").order("tarih", desc=True))
                if islemler:
                    secenekler_i = {f"{i['tarih']} | {i['cari_adi']} | {i['islem_tipi']} | {i['tutar']} ₺ (ID: {i['id']})": i for i in islemler}
                    sec_i_str = st.selectbox("İşlem Seçin", ["Seçiniz..."] + list(secenekler_i.keys()))
                    if sec_i_str != "Seçiniz...":
                        sec_i = secenekler_i[sec_i_str]
                        with st.form("islem_duz"):
                            y_tarih = st.date_input("Tarih", value=datetime.datetime.strptime(sec_i['tarih'], '%Y-%m-%d').date())
                            
                            islem_tipi_val = sec_i.get('islem_tipi', 'Gelen Fatura (Bize Borç Yazar)')
                            if islem_tipi_val not in ["Gelen Fatura (Bize Borç Yazar)", "Ödeme Yaptık (Borç Düşer)"]: islem_tipi_val = "Gelen Fatura (Bize Borç Yazar)"
                            y_tip = st.selectbox("İşlem Tipi", ["Gelen Fatura (Bize Borç Yazar)", "Ödeme Yaptık (Borç Düşer)"], index=["Gelen Fatura (Bize Borç Yazar)", "Ödeme Yaptık (Borç Düşer)"].index(islem_tipi_val))
                            
                            y_tutar = st.number_input("Tutar", value=float(sec_i.get('tutar', 0)))
                            
                            odm_val = sec_i.get('odeme_tipi', '- Yok -')
                            if odm_val not in odeme_yontemleri: odm_val = "- Yok -"
                            y_odeme = st.selectbox("Nereden Ödendi?", odeme_yontemleri, index=odeme_yontemleri.index(odm_val))
                            
                            y_ack = st.text_input("Açıklama", value=sec_i.get('aciklama', ''))
                            cg, cs = st.columns(2)
                            with cg:
                                if st.form_submit_button("Güncelle"):
                                    if str(sec_i.get('islem_tipi')) == "Ödeme Yaptık (Borç Düşer)" and str(sec_i.get('odeme_tipi')) in b_liste:
                                        db_yaz(supabase.table("banka_islemleri").delete().eq("hesap_adi", sec_i['odeme_tipi']).eq("tarih", str(sec_i['tarih'])).eq("islem_tipi", "Para Çıkışı").ilike("aciklama", f"Cari Ödemesi: {sec_i.get('aciklama', '')}%"))
                                    
                                    db_yaz(supabase.table("cari_islemler").update({"tarih": str(y_tarih), "islem_tipi": y_tip, "tutar": y_tutar, "aciklama": y_ack, "odeme_tipi": y_odeme if y_tip == "Ödeme Yaptık (Borç Düşer)" else "- Yok -"}).eq("id", sec_i['id']))
                                    
                                    if y_tip == "Ödeme Yaptık (Borç Düşer)" and y_odeme in b_liste:
                                        db_yaz(supabase.table("banka_islemleri").insert({"tarih": str(y_tarih), "hesap_adi": y_odeme, "islem_tipi": "Para Çıkışı", "karsi_hesap": sec_i['cari_adi'], "tutar": y_tutar, "aciklama": f"Cari Ödemesi: {y_ack}"}))
                                    st.session_state.genel_mesaj = ("success", "İşlem güncellendi!")
                                    st.rerun()
                            with cs:
                                if st.form_submit_button("Sil"):
                                    if str(sec_i.get('islem_tipi')) == "Ödeme Yaptık (Borç Düşer)" and str(sec_i.get('odeme_tipi')) in b_liste:
                                        db_yaz(supabase.table("banka_islemleri").delete().eq("hesap_adi", sec_i['odeme_tipi']).eq("tarih", str(sec_i['tarih'])).eq("islem_tipi", "Para Çıkışı").ilike("aciklama", f"Cari Ödemesi: {sec_i.get('aciklama', '')}%"))
                                    
                                    if db_yaz(supabase.table("cari_islemler").delete().eq("id", sec_i['id'])):
                                        st.session_state.genel_mesaj = ("info", "İşlem tamamen silindi!")
                                    st.rerun()

    elif alt_menu == "📋 Cari Ekstre":
        islemler = db_oku(supabase.table("cari_islemler").select("*"))
        if islemler:
            df_i = pd.DataFrame(islemler)
            if 'tutar' not in df_i.columns: df_i['tutar'] = 0.0
            if 'odeme_tipi' not in df_i.columns: df_i['odeme_tipi'] = ""
            if 'aciklama' not in df_i.columns: df_i['aciklama'] = ""
            
            df_i['tutar'] = pd.to_numeric(df_i['tutar'], errors='coerce').fillna(0).round(2)
            
            fatura_toplam = df_i[df_i['islem_tipi'] == 'Gelen Fatura (Bize Borç Yazar)'].groupby('cari_adi')['tutar'].sum()
            odeme_toplam = df_i[df_i['islem_tipi'] == 'Ödeme Yaptık (Borç Düşer)'].groupby('cari_adi')['tutar'].sum()
            bakiye_df = pd.DataFrame({'Toplam Fatura Tutarı': fatura_toplam, 'Ödenen Tutar': odeme_toplam}).fillna(0)
            
            bakiye_df['Toplam Fatura Tutarı'] = bakiye_df['Toplam Fatura Tutarı'].round(2)
            bakiye_df['Ödenen Tutar'] = bakiye_df['Ödenen Tutar'].round(2)
            bakiye_df['KALAN BORCUMUZ'] = (bakiye_df['Toplam Fatura Tutarı'] - bakiye_df['Ödenen Tutar']).round(2)
            st.dataframe(bakiye_df.reset_index(), hide_index=True, use_container_width=True)
            
            st.divider()
            st.subheader("Tüm Cari Hareketler Dökümü")
            z_goster_c2 = st.toggle("⏱️ Kayıt (İşlenme) Zamanlarını Göster", key="cari_zg")
            df_i['odeme_tipi'] = df_i.get('odeme_tipi', '- Yok -')
            df_i['tarih'] = pd.to_datetime(df_i['tarih']).dt.date
            
            with st.expander("🔍 Cari Filtreleme Paneli", expanded=True):
                c1, c2, c3, c4 = st.columns(4)
                with c1: t_aralik_c = st.date_input("Tarih Aralığı", [df_i['tarih'].min(), df_i['tarih'].max()], key="filt_c_tar")
                with c2: sec_cari = st.multiselect("Cari (Firma) Seç", df_i['cari_adi'].unique().tolist(), key="filt_c_cari")
                with c3: sec_tip_c = st.multiselect("İşlem Tipi", df_i['islem_tipi'].unique().tolist(), key="filt_c_tip")
                with c4: ara_c = st.text_input("Açıklama Ara", key="filt_c_ara")
                
            if len(t_aralik_c) == 2: df_i = df_i[(df_i['tarih'] >= t_aralik_c[0]) & (df_i['tarih'] <= t_aralik_c[1])]
            elif len(t_aralik_c) == 1: df_i = df_i[df_i['tarih'] == t_aralik_c[0]]
            if sec_cari: df_i = df_i[df_i['cari_adi'].isin(sec_cari)]
            if sec_tip_c: df_i = df_i[df_i['islem_tipi'].isin(sec_tip_c)]
            if ara_c: df_i = df_i[df_i['aciklama'].str.contains(ara_c, case=False, na=False)]
            
            gosterim_cari = ['tarih', 'cari_adi', 'islem_tipi', 'tutar', 'odeme_tipi', 'aciklama']
            if z_goster_c2: df_i, gosterim_cari = zaman_sutunlari_ekle(df_i, gosterim_cari)
            
            st.dataframe(df_i[gosterim_cari].sort_values("tarih", ascending=False), hide_index=True, use_container_width=True)
            
            dosya_c, uzanti_c, mime_c = excel_indir(df_i[gosterim_cari])
            st.download_button("📥 Filtrelenmiş Dökümü Excel'e İndir", data=dosya_c, file_name=f"Cari_Hareketleri.{uzanti_c}", mime=mime_c, key="dl_cari")

elif menu == "Kasa Yönetimi (Virman)":
    st.header("Kasa Yönetimi ve Virman")
    bildirim_goster()
    secilen = st.date_input("İşlem Tarihi", datetime.date.today())
    
    st.subheader("🏦 Kasa - Banka Arası Transfer")
    bankalar_db = db_oku(supabase.table("banka_hesaplari").select("*"))
    banka_liste = [b['isim'] for b in bankalar_db] if bankalar_db else []
    
    if banka_liste:
        with st.form("kasa_banka_transfer"):
            c1, c2, c3, c4 = st.columns(4)
            with c1: kb_yon = st.selectbox("İşlem Yönü", ["Kasadan/Havuzdan Bankaya Yatırma", "Bankadan Kasaya Çekme"])
            with c2: kb_kasa = st.selectbox("Hangi Kasa / Havuz?", ["Kasa 1", "Kasa 2", "POS Havuzu", "Pavo Havuzu"])
            with c3: kb_banka = st.selectbox("Hangi Banka?", banka_liste)
            with c4: 
                kb_tutar = st.number_input("Tutar (₺)", min_value=0.0)
                st.markdown("<br>", unsafe_allow_html=True)
                if st.form_submit_button("Transferi Gerçekleştir", type="primary"):
                    if kb_tutar > 0:
                        if kb_yon == "Bankadan Kasaya Çekme" and kb_kasa in ["POS Havuzu", "Pavo Havuzu"]:
                            st.error("Bankadan POS veya Pavo havuzuna para çekilemez!")
                        else:
                            if kb_yon == "Kasadan/Havuzdan Bankaya Yatırma":
                                db_yaz(supabase.table("kasa_islemleri").insert({"tarih": str(secilen), "islem_tipi": "Bankaya Yatırılan", "gonderen": kb_kasa, "tutar": kb_tutar}))
                                db_yaz(supabase.table("banka_islemleri").insert({"tarih": str(secilen), "hesap_adi": kb_banka, "islem_tipi": "Para Girişi", "karsi_hesap": kb_kasa, "tutar": kb_tutar, "aciklama": f"{kb_kasa}'dan Yatırılan"}))
                                st.session_state.genel_mesaj = ("success", "Para bankaya aktarıldı!")
                            else:
                                db_yaz(supabase.table("banka_islemleri").insert({"tarih": str(secilen), "hesap_adi": kb_banka, "islem_tipi": "Para Çıkışı", "karsi_hesap": kb_kasa, "tutar": kb_tutar, "aciklama": f"{kb_kasa}'ya Çekilen"}))
                                db_yaz(supabase.table("kasa_islemleri").insert({"tarih": str(secilen), "islem_tipi": "Bankadan Çekilen", "alan": kb_kasa, "tutar": kb_tutar}))
                                st.session_state.genel_mesaj = ("success", "Para kasaya çekildi!")
                            st.rerun()
    else:
        st.info("💡 Kasa ile Banka arasında transfer yapmak için lütfen önce 'Banka & Kart Yönetimi' sayfasından Banka Hesabı ekleyin.")

    st.divider()

    st.subheader("Dışarıdan Kasaya Nakit Ekle (Sermaye / Bozukluk)")
    with st.form("nakit_ekle_form"):
        c1, c2, c3 = st.columns(3)
        with c1: k_secim = st.selectbox("Hangi Kasa?", ["Kasa 1", "Kasa 2"])
        with c2: k_tutar = st.number_input("Tutar (₺)", min_value=0.0)
        with c3:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.form_submit_button("Kasaya Ekle"):
                if k_tutar > 0:
                    if db_yaz(supabase.table("kasa_islemleri").insert({"tarih": str(secilen), "islem_tipi": "Para Girişi (Sermaye)", "alan": k_secim, "tutar": k_tutar})):
                        st.session_state.genel_mesaj = ("success", "Nakit başarıyla eklendi!")
                    st.rerun()

    st.subheader("Kasalar Arası Virman")
    with st.form("virman_form"):
        c1, c2, c3, c4 = st.columns(4)
        with c1: gonderen = st.selectbox("Gönderen", ["Kasa 1", "Kasa 2"])
        with c2: alan = st.selectbox("Alan", ["Kasa 2", "Kasa 1"])
        with c3: tutar_v = st.number_input("Tutar (₺)", min_value=0.0)
        with c4:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.form_submit_button("Virman Yap"):
                if gonderen != alan and tutar_v > 0:
                    if db_yaz(supabase.table("kasa_islemleri").insert({"tarih": str(secilen), "islem_tipi": "Virman", "gonderen": gonderen, "alan": alan, "tutar": tutar_v})):
                        st.session_state.genel_mesaj = ("success", "Virman başarıyla yapıldı!")
                    st.rerun()

    st.subheader("Kasa Sayım Farkı (Eksik / Fazla)")
    with st.form("kasa_fark_form"):
        c1, c2, c3, c4 = st.columns(4)
        with c1: f_kasa = st.selectbox("Kasa?", ["Kasa 1", "Kasa 2"])
        with c2: f_durum = st.selectbox("Durum", ["Eksik Çıktı", "Fazla Çıktı"])
        with c3: f_tutar = st.number_input("Tutar", min_value=0.0)
        with c4:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.form_submit_button("Farkı İşle"):
                islem = "Eksik" if f_durum == "Eksik Çıktı" else "Fazla"
                if db_yaz(supabase.table("kasa_islemleri").insert({"tarih": str(secilen), "islem_tipi": islem, "gonderen": f_kasa if islem=="Eksik" else None, "alan": f_kasa if islem=="Fazla" else None, "tutar": f_tutar})):
                    st.session_state.genel_mesaj = ("success", "Kasa farkı işlendi!")
                st.rerun()

    st.divider()
    with st.expander("✏️ Sermaye, Virman ve Fark Düzenle", expanded=False):
        gecmis_islemler = db_oku(supabase.table("kasa_islemleri").select("*").in_("islem_tipi", ["Virman", "Eksik", "Fazla", "Açılış", "Para Girişi (Sermaye)", "Bankaya Yatırılan", "Bankadan Çekilen"]).order("tarih", desc=True))
        if gecmis_islemler:
            secenekler_k = {}
            for i in gecmis_islemler:
                if i['islem_tipi'] in ['Açılış', 'Para Girişi (Sermaye)']: lbl = f"{i['tarih']} | SERMAYE/AÇILIŞ | {i.get('alan')} | {i.get('tutar',0)} ₺"
                elif i['islem_tipi'] == 'Virman': lbl = f"{i['tarih']} | VİRMAN | {i.get('gonderen')} -> {i.get('alan')} | {i.get('tutar',0)} ₺"
                elif i['islem_tipi'] == 'Eksik': lbl = f"{i['tarih']} | EKSİK ÇIKTI | {i.get('gonderen')} | {i.get('tutar',0)} ₺"
                elif i['islem_tipi'] == 'Fazla': lbl = f"{i['tarih']} | FAZLA ÇIKTI | {i.get('alan')} | {i.get('tutar',0)} ₺"
                elif i['islem_tipi'] == 'Bankaya Yatırılan': lbl = f"{i['tarih']} | BANKAYA YATIRILAN | {i.get('gonderen')} | {i.get('tutar',0)} ₺"
                elif i['islem_tipi'] == 'Bankadan Çekilen': lbl = f"{i['tarih']} | BANKADAN ÇEKİLEN | {i.get('alan')} | {i.get('tutar',0)} ₺"
                secenekler_k[f"{lbl} (ID: {i['id']})"] = i
                
            sec_k_str = st.selectbox("İşlem Seçin", ["Lütfen seçin..."] + list(secenekler_k.keys()))
            if sec_k_str != "Lütfen seçin...":
                sec_k = secenekler_k[sec_k_str]
                with st.form("k_duz_form"):
                    try: k_tar = datetime.datetime.strptime(sec_k['tarih'], '%Y-%m-%d').date()
                    except: k_tar = datetime.date.today()
                    y_tar = st.date_input("Tarih", value=k_tar)
                    
                    islem_turleri = ["Virman", "Eksik", "Fazla", "Para Girişi (Sermaye)", "Bankaya Yatırılan", "Bankadan Çekilen"]
                    i_tip_val = sec_k.get('islem_tipi', 'Virman')
                    if i_tip_val not in islem_turleri: i_tip_val = "Virman"
                    y_islem = st.selectbox("İşlem Tipi", islem_turleri, index=islem_turleri.index(i_tip_val))
                    
                    k_list = ["Kasa 1", "Kasa 2", "POS Havuzu", "Pavo Havuzu"]
                    c_k1, c_k2 = st.columns(2)
                    with c_k1:
                        g_val = sec_k.get('gonderen', 'Kasa 1')
                        if g_val not in k_list: g_val = k_list[0]
                        y_gon = st.selectbox("Gönderen (Veya Eksik) Kasa/Havuz", k_list, index=k_list.index(g_val))
                    with c_k2:
                        a_val = sec_k.get('alan', 'Kasa 2')
                        if a_val not in k_list: a_val = k_list[1]
                        y_aln = st.selectbox("Alan (Veya Fazla) Kasa/Havuz", k_list, index=k_list.index(a_val))
                        
                    y_tutar = st.number_input("Tutar (₺)", min_value=0.0, value=float(sec_k.get('tutar', 0)))
                    
                    c_gun, c_sil = st.columns(2)
                    with c_gun:
                        if st.form_submit_button("Güncelle"):
                            if y_islem == "Virman": veri = {"tarih": str(y_tar), "islem_tipi": y_islem, "gonderen": y_gon, "alan": y_aln, "tutar": y_tutar}
                            elif y_islem in ["Eksik", "Bankaya Yatırılan"]: veri = {"tarih": str(y_tar), "islem_tipi": y_islem, "gonderen": y_gon, "alan": None, "tutar": y_tutar}
                            elif y_islem in ["Fazla", "Para Girişi (Sermaye)", "Bankadan Çekilen"]: veri = {"tarih": str(y_tar), "islem_tipi": y_islem, "gonderen": None, "alan": y_aln, "tutar": y_tutar}
                            db_yaz(supabase.table("kasa_islemleri").update(veri).eq("id", sec_k['id']))
                            st.session_state.genel_mesaj = ("success", "Kasa işlemi güncellendi!")
                            st.rerun()
                    with c_sil:
                        if st.form_submit_button("Sil"):
                            if sec_k['islem_tipi'] == "Bankaya Yatırılan":
                                db_yaz(supabase.table("banka_islemleri").delete().eq("tarih", sec_k['tarih']).eq("tutar", sec_k['tutar']).eq("aciklama", f"{sec_k.get('gonderen')}['dan Yatırılan]"))
                            elif sec_k['islem_tipi'] == "Bankadan Çekilen":
                                db_yaz(supabase.table("banka_islemleri").delete().eq("tarih", sec_k['tarih']).eq("tutar", sec_k['tutar']).eq("aciklama", f"{sec_k.get('alan')}['ya Çekilen]"))
                                
                            if db_yaz(supabase.table("kasa_islemleri").delete().eq("id", sec_k['id'])):
                                st.session_state.genel_mesaj = ("info", "Kasa işlemi silindi!")
                            st.rerun()

    cirolar_all = db_oku(supabase.table("ciro").select("*"))
    if cirolar_all is None: cirolar_all = []
    
    masraflar_all = db_oku(supabase.table("masraf").select("*"))
    if masraflar_all is None: masraflar_all = []
    
    islemler_all = db_oku(supabase.table("kasa_islemleri").select("*"))
    if islemler_all is None: islemler_all = []
    
    cari_islemler_all = db_oku(supabase.table("cari_islemler").select("*"))
    if cari_islemler_all is None: cari_islemler_all = []

    def kasa_durumu(k_adi):
        g_c = [c for c in cirolar_all if c['tarih'] < str(secilen) and c.get('kasa') == k_adi]
        g_m = [m for m in masraflar_all if m['tarih'] < str(secilen) and m.get('odeme_tipi') == f"Nakit - {k_adi}"]
        g_co = [co for co in cari_islemler_all if co['tarih'] < str(secilen) and co.get('islem_tipi') == 'Ödeme Yaptık (Borç Düşer)' and co.get('odeme_tipi') == f"Nakit - {k_adi}"]
        g_i = [i for i in islemler_all if i['tarih'] < str(secilen)]

        devir = sum([(float(c.get('nakit', 0)) + float(c.get('pavo_nakit', 0))) for c in g_c])
        devir -= sum([float(m.get('tutar', 0)) for m in g_m])
        devir -= sum([float(co.get('tutar', 0)) for co in g_co]) 
        devir += sum([float(i.get('tutar', 0)) for i in g_i if i.get('islem_tipi') in ['Açılış', 'Para Girişi (Sermaye)', 'Bankadan Çekilen'] and i.get('alan') == k_adi])
        devir += sum([float(i.get('tutar', 0)) for i in g_i if i.get('islem_tipi') == 'Virman' and i.get('alan') == k_adi])
        devir -= sum([float(i.get('tutar', 0)) for i in g_i if i.get('islem_tipi') == 'Virman' and i.get('gonderen') == k_adi])
        devir -= sum([float(i.get('tutar', 0)) for i in g_i if i.get('islem_tipi') in ['Eksik', 'Bankaya Yatırılan'] and i.get('gonderen') == k_adi])
        devir += sum([float(i.get('tutar', 0)) for i in g_i if i.get('islem_tipi') == 'Fazla' and i.get('alan') == k_adi])

        b_c = [c for c in cirolar_all if c['tarih'] == str(secilen) and c.get('kasa') == k_adi]
        b_m = [m for m in masraflar_all if m['tarih'] == str(secilen) and m.get('odeme_tipi') == f"Nakit - {k_adi}"]
        b_co = [co for co in cari_islemler_all if co['tarih'] == str(secilen) and co.get('islem_tipi') == 'Ödeme Yaptık (Borç Düşer)' and co.get('odeme_tipi') == f"Nakit - {k_adi}"]
        b_i = [i for i in islemler_all if i['tarih'] == str(secilen)]

        b_giris = sum([(float(c.get('nakit', 0)) + float(c.get('pavo_nakit', 0))) for c in b_c])
        b_cikis = sum([float(m.get('tutar', 0)) for m in b_m])
        b_cari_odeme = sum([float(co.get('tutar', 0)) for co in b_co])
        b_ekle = sum([float(i.get('tutar', 0)) for i in b_i if i.get('islem_tipi') in ['Açılış', 'Para Girişi (Sermaye)'] and i.get('alan') == k_adi])
        b_vg = sum([float(i.get('tutar', 0)) for i in b_i if i.get('islem_tipi') == 'Virman' and i.get('alan') == k_adi])
        b_vc = sum([float(i.get('tutar', 0)) for i in b_i if i.get('islem_tipi') == 'Virman' and i.get('gonderen') == k_adi])
        b_eksik = sum([float(i.get('tutar', 0)) for i in b_i if i.get('islem_tipi') == 'Eksik' and i.get('gonderen') == k_adi])
        b_fazla = sum([float(i.get('tutar', 0)) for i in b_i if i.get('islem_tipi') == 'Fazla' and i.get('alan') == k_adi])
        
        b_bankaya_yatan = sum([float(i.get('tutar', 0)) for i in b_i if i.get('islem_tipi') == 'Bankaya Yatırılan' and i.get('gonderen') == k_adi])
        b_bankadan_cekilen = sum([float(i.get('tutar', 0)) for i in b_i if i.get('islem_tipi') == 'Bankadan Çekilen' and i.get('alan') == k_adi])

        gun_sonu = devir + b_giris + b_ekle + b_bankadan_cekilen + b_vg + b_fazla - b_cikis - b_cari_odeme - b_bankaya_yatan - b_vc - b_eksik
        return round(devir,2), round(b_giris,2), round(b_ekle,2), round(b_cikis,2), round(b_cari_odeme,2), round(b_vg,2), round(b_vc,2), round(b_eksik,2), round(b_fazla,2), round(b_bankaya_yatan,2), round(b_bankadan_cekilen,2), round(gun_sonu,2)

    def havuz_durumu(h_adi):
        g_c = [c for c in cirolar_all if c['tarih'] < str(secilen)]
        g_i = [i for i in islemler_all if i['tarih'] < str(secilen)]
        if h_adi == "POS Havuzu": devir_ciro = sum([float(c.get('kredi_karti', 0)) for c in g_c])
        else: devir_ciro = sum([float(c.get('pavo_kredi', 0)) for c in g_c])
        devir_cikis = sum([float(i.get('tutar', 0)) for i in g_i if i.get('islem_tipi') == 'Bankaya Yatırılan' and i.get('gonderen') == h_adi])
        devir = devir_ciro - devir_cikis

        b_c = [c for c in cirolar_all if c['tarih'] == str(secilen)]
        b_i = [i for i in islemler_all if i['tarih'] == str(secilen)]
        if h_adi == "POS Havuzu": bugun_ciro = sum([float(c.get('kredi_karti', 0)) for c in b_c])
        else: bugun_ciro = sum([float(c.get('pavo_kredi', 0)) for c in b_c])
        bugun_cikis = sum([float(i.get('tutar', 0)) for i in b_i if i.get('islem_tipi') == 'Bankaya Yatırılan' and i.get('gonderen') == h_adi])
        
        net = devir + bugun_ciro - bugun_cikis
        return round(devir,2), round(bugun_ciro,2), round(bugun_cikis,2), round(net,2)

    st.divider()
    st.subheader("📊 Günün Kasa Özetleri")
    
    k1_devir, k1_giris, k1_ekle, k1_cikis, k1_cari_odeme, k1_vg, k1_vc, k1_eksik, k1_fazla, k1_bankaya_yatan, k1_bankadan_cekilen, k1_net = kasa_durumu("Kasa 1")
    k2_devir, k2_giris, k2_ekle, k2_cikis, k2_cari_odeme, k2_vg, k2_vc, k2_eksik, k2_fazla, k2_bankaya_yatan, k2_bankadan_cekilen, k2_net = kasa_durumu("Kasa 2")

    col1, col2 = st.columns(2)
    with col1:
        st.info("### KASA 1")
        st.write(f"**Dünden Devreden:** {k1_devir:,.2f} ₺")
        st.write(f"Bugün Eklenen (Sermaye): + {k1_ekle:,.2f} ₺")
        st.write(f"Bugünkü Ciro (Nakit): + {k1_giris:,.2f} ₺")
        st.write(f"Bankadan Çekilen: + {k1_bankadan_cekilen:,.2f} ₺")
        st.write(f"Kasa Fazlası: + {k1_fazla:,.2f} ₺")
        st.write(f"Bugünkü Masraflar: - {k1_cikis:,.2f} ₺")
        st.write(f"Cari Ödemeleri (Nakit): - {k1_cari_odeme:,.2f} ₺")
        st.write(f"Bankaya Yatan: - {k1_bankaya_yatan:,.2f} ₺")
        st.write(f"Kasa Eksiği: - {k1_eksik:,.2f} ₺")
        st.write(f"Virman (Gelen - Giden): {(k1_vg - k1_vc):,.2f} ₺")
        st.metric("GÜN SONU KASADA OLMASI GEREKEN", f"{k1_net:,.2f} ₺")
    with col2:
        st.success("### KASA 2")
        st.write(f"**Dünden Devreden:** {k2_devir:,.2f} ₺")
        st.write(f"Bugün Eklenen (Sermaye): + {k2_ekle:,.2f} ₺")
        st.write(f"Bugünkü Ciro (Nakit): + {k2_giris:,.2f} ₺")
        st.write(f"Bankadan Çekilen: + {k2_bankadan_cekilen:,.2f} ₺")
        st.write(f"Kasa Fazlası: + {k2_fazla:,.2f} ₺")
        st.write(f"Bugünkü Masraflar: - {k2_cikis:,.2f} ₺")
        st.write(f"Cari Ödemeleri (Nakit): - {k2_cari_odeme:,.2f} ₺")
        st.write(f"Bankaya Yatan: - {k2_bankaya_yatan:,.2f} ₺")
        st.write(f"Kasa Eksiği: - {k2_eksik:,.2f} ₺")
        st.write(f"Virman (Gelen - Giden): {(k2_vg - k2_vc):,.2f} ₺")
        st.metric("GÜN SONU KASADA OLMASI GEREKEN", f"{k2_net:,.2f} ₺")

    st.divider()
    c3, c4 = st.columns(2)
    pos_devir, pos_giris, pos_cikis, pos_net = havuz_durumu("POS Havuzu")
    pavo_devir, pavo_giris, pavo_cikis, pavo_net = havuz_durumu("Pavo Havuzu")
    
    with c3:
        st.warning("💳 POS (KREDİ KARTI) HAVUZU")
        st.write(f"**Dünden Devreden:** {pos_devir:,.2f} ₺")
        st.write(f"Bugünkü Cirodan Gelen: + {pos_giris:,.2f} ₺")
        st.write(f"Bankaya Aktarılan: - {pos_cikis:,.2f} ₺")
        st.metric("HAVUZDA BEKLEYEN BİRİKİM", f"{pos_net:,.2f} ₺")
        
    with c4:
        st.error("💳 PAVO KREDİ HAVUZU")
        st.write(f"**Dünden Devreden:** {pavo_devir:,.2f} ₺")
        st.write(f"Bugünkü Cirodan Gelen: + {pavo_giris:,.2f} ₺")
        st.write(f"Bankaya Aktarılan: - {pavo_cikis:,.2f} ₺")
        st.metric("HAVUZDA BEKLEYEN BİRİKİM", f"{pavo_net:,.2f} ₺")

    # --- KASA HAREKETLERİ MENÜSÜ ---
    st.divider()
    st.subheader("📋 Kasa ve Havuz Dökümleri")
    
    kasa_dokum_menu = st.radio(
        "Döküm Tipi Seçin", 
        ["🧾 Bakiyeli Hesap Ekstresi", "📋 Tüm Genel Hareketler"], 
        horizontal=True, 
        label_visibility="collapsed"
    )
    
    if kasa_dokum_menu == "🧾 Bakiyeli Hesap Ekstresi":
        secili_ekstre = st.selectbox("Ekstresini Görmek İstediğiniz Kasa / Havuz", ["Lütfen seçin...", "Kasa 1", "Kasa 2", "POS Havuzu", "Pavo Havuzu"])
        
        if secili_ekstre != "Lütfen seçin...":
            hesap_hareketleri = []
            
            for c in cirolar_all:
                if secili_ekstre in ["Kasa 1", "Kasa 2"]:
                    if c.get('kasa') == secili_ekstre:
                        tut = float(c.get('nakit', 0)) + float(c.get('pavo_nakit', 0))
                        if tut > 0:
                            hesap_hareketleri.append({"id": c['id'], "tarih": c['tarih'], "islem": "Ciro Girişi", "aciklama": "Günlük Nakit Ciro", "Giriş": tut, "Çıkış": 0.0, "created_at": c.get('created_at'), "updated_at": c.get('updated_at')})
                elif secili_ekstre == "POS Havuzu":
                    tut = float(c.get('kredi_karti', 0))
                    if tut > 0:
                        hesap_hareketleri.append({"id": c['id'], "tarih": c['tarih'], "islem": "Ciro Girişi", "aciklama": "Kredi Kartı Cirosu", "Giriş": tut, "Çıkış": 0.0, "created_at": c.get('created_at'), "updated_at": c.get('updated_at')})
                elif secili_ekstre == "Pavo Havuzu":
                    tut = float(c.get('pavo_kredi', 0))
                    if tut > 0:
                        hesap_hareketleri.append({"id": c['id'], "tarih": c['tarih'], "islem": "Ciro Girişi", "aciklama": "Pavo Kredi Cirosu", "Giriş": tut, "Çıkış": 0.0, "created_at": c.get('created_at'), "updated_at": c.get('updated_at')})

            if secili_ekstre in ["Kasa 1", "Kasa 2"]:
                for m in masraflar_all:
                    if str(m.get('odeme_tipi', '')) == f"Nakit - {secili_ekstre}":
                        tut = float(m.get('tutar', 0))
                        if tut > 0:
                            hesap_hareketleri.append({"id": m['id'], "tarih": m['tarih'], "islem": "Masraf Çıkışı", "aciklama": m.get('aciklama', ''), "Giriş": 0.0, "Çıkış": tut, "created_at": m.get('created_at'), "updated_at": m.get('updated_at')})

            if secili_ekstre in ["Kasa 1", "Kasa 2"]:
                for co in cari_islemler_all:
                    if co.get('islem_tipi') == 'Ödeme Yaptık (Borç Düşer)' and str(co.get('odeme_tipi', '')) == f"Nakit - {secili_ekstre}":
                        tut = float(co.get('tutar', 0))
                        if tut > 0:
                            hesap_hareketleri.append({"id": co['id'], "tarih": co['tarih'], "islem": "Cari Ödemesi", "aciklama": f"Firma: {co.get('cari_adi', '')} - {co.get('aciklama', '')}", "Giriş": 0.0, "Çıkış": tut, "created_at": co.get('created_at'), "updated_at": co.get('updated_at')})

            for i in islemler_all:
                tip = i.get('islem_tipi', '')
                tut = float(i.get('tutar', 0))
                g = i.get('gonderen')
                a = i.get('alan')
                
                if tut > 0:
                    if tip in ['Açılış', 'Para Girişi (Sermaye)', 'Bankadan Çekilen', 'Fazla'] and a == secili_ekstre:
                        hesap_hareketleri.append({"id": i['id'], "tarih": i['tarih'], "islem": tip, "aciklama": i.get('aciklama', '-'), "Giriş": tut, "Çıkış": 0.0, "created_at": i.get('created_at'), "updated_at": i.get('updated_at')})
                    elif tip in ['Eksik', 'Bankaya Yatırılan'] and g == secili_ekstre:
                        hesap_hareketleri.append({"id": i['id'], "tarih": i['tarih'], "islem": tip, "aciklama": i.get('aciklama', '-'), "Giriş": 0.0, "Çıkış": tut, "created_at": i.get('created_at'), "updated_at": i.get('updated_at')})
                    elif tip == 'Virman':
                        if a == secili_ekstre:
                            hesap_hareketleri.append({"id": i['id'], "tarih": i['tarih'], "islem": "Virman Girişi", "aciklama": f"Gönderen: {g}", "Giriş": tut, "Çıkış": 0.0, "created_at": i.get('created_at'), "updated_at": i.get('updated_at')})
                        if g == secili_ekstre:
                            hesap_hareketleri.append({"id": i['id'], "tarih": i['tarih'], "islem": "Virman Çıkışı", "aciklama": f"Alıcı: {a}", "Giriş": 0.0, "Çıkış": tut, "created_at": i.get('created_at'), "updated_at": i.get('updated_at')})

            if hesap_hareketleri:
                df_e = pd.DataFrame(hesap_hareketleri)
                df_e['tarih_dt'] = pd.to_datetime(df_e['tarih'])
                df_e['is_acilis'] = df_e['islem'].apply(lambda x: 0 if 'Açılış' in x or 'Sermaye' in x else 1)
                
                df_e = df_e.sort_values(by=["tarih_dt", "is_acilis", "id"], ascending=[True, True, True]).reset_index(drop=True)
                
                bakiye_list = []
                bakiye = 0.0
                for idx, r in df_e.iterrows():
                    bakiye = round(bakiye + r['Giriş'] - r['Çıkış'], 2)
                    bakiye_list.append(bakiye)
                
                df_e['Bakiye'] = bakiye_list
                df_e['tarih'] = df_e['tarih_dt'].dt.date
                
                df_e = df_e.sort_values(by=["tarih_dt", "is_acilis", "id"], ascending=[False, False, False]).drop(columns=['tarih_dt', 'is_acilis', 'id'])
                
                df_e['Giriş'] = df_e['Giriş'].round(2)
                df_e['Çıkış'] = df_e['Çıkış'].round(2)
                df_e['Bakiye'] = df_e['Bakiye'].round(2)

                z_goster_ke = st.toggle("⏱️ Kayıt (İşlenme) Zamanlarını Göster", key="kasa_ekstre_zg")
                gosterim_ekstre = ['tarih', 'islem', 'aciklama', 'Giriş', 'Çıkış', 'Bakiye']
                if z_goster_ke: df_e, gosterim_ekstre = zaman_sutunlari_ekle(df_e, gosterim_ekstre)
                
                st.dataframe(df_e[gosterim_ekstre], hide_index=True, use_container_width=True)
                dosya_e, uzanti_e, mime_e = excel_indir(df_e[gosterim_ekstre])
                st.download_button(f"📥 {secili_ekstre} Ekstresini İndir", data=dosya_e, file_name=f"{secili_ekstre}_Ekstresi.{uzanti_e}", mime=mime_e, key="dl_kasa_ekstre")
            else:
                st.info("Bu kasa/havuz için henüz hareket bulunmuyor.")

    elif kasa_dokum_menu == "📋 Tüm Genel Hareketler":
        kasa_dokum = []
        
        if cirolar_all:
            for c in cirolar_all:
                # Nakit Kasası
                n_tutar = float(c.get('nakit', 0)) + float(c.get('pavo_nakit', 0))
                if n_tutar > 0:
                    kasa_dokum.append({"Tarih": c['tarih'], "Kasa": c.get('kasa'), "İşlem": "Ciro Girişi", "Yön": "Giriş", "Tutar": n_tutar, "Açıklama": "Günlük Nakit Ciro", "created_at": c.get('created_at'), "updated_at": c.get('updated_at')})
                
                # POS Havuzu (Yeni Eklendi)
                kk_tutar = float(c.get('kredi_karti', 0))
                if kk_tutar > 0:
                    kasa_dokum.append({"Tarih": c['tarih'], "Kasa": "POS Havuzu", "İşlem": "Ciro Girişi", "Yön": "Giriş", "Tutar": kk_tutar, "Açıklama": "Kredi Kartı Cirosu", "created_at": c.get('created_at'), "updated_at": c.get('updated_at')})
                    
                # Pavo Havuzu (Yeni Eklendi)
                pk_tutar = float(c.get('pavo_kredi', 0))
                if pk_tutar > 0:
                    kasa_dokum.append({"Tarih": c['tarih'], "Kasa": "Pavo Havuzu", "İşlem": "Ciro Girişi", "Yön": "Giriş", "Tutar": pk_tutar, "Açıklama": "Pavo Kredi Cirosu", "created_at": c.get('created_at'), "updated_at": c.get('updated_at')})
                    
        if masraflar_all:
            for m in masraflar_all:
                o_tipi = str(m.get('odeme_tipi', ''))
                if o_tipi.startswith("Nakit - "):
                    k_adi = o_tipi.replace("Nakit - ", "")
                    kasa_dokum.append({"Tarih": m['tarih'], "Kasa": k_adi, "İşlem": "Masraf Çıkışı", "Yön": "Çıkış", "Tutar": float(m.get('tutar',0)), "Açıklama": m.get('aciklama', ''), "created_at": m.get('created_at'), "updated_at": m.get('updated_at')})
                    
        if cari_islemler_all:
            for co in cari_islemler_all:
                o_tipi = str(co.get('odeme_tipi', ''))
                if o_tipi.startswith("Nakit - "):
                    k_adi = o_tipi.replace("Nakit - ", "")
                    kasa_dokum.append({"Tarih": co['tarih'], "Kasa": k_adi, "İşlem": "Cari Ödemesi", "Yön": "Çıkış", "Tutar": float(co.get('tutar',0)), "Açıklama": f"Firma: {co.get('cari_adi', '')} - {co.get('aciklama', '')}", "created_at": co.get('created_at'), "updated_at": co.get('updated_at')})
                    
        if islemler_all:
            for i in islemler_all:
                tip = i.get('islem_tipi', '')
                tut = float(i.get('tutar', 0))
                g = i.get('gonderen')
                a = i.get('alan')
                
                if tip in ['Açılış', 'Para Girişi (Sermaye)', 'Bankadan Çekilen']:
                    kasa_dokum.append({"Tarih": i['tarih'], "Kasa": a, "İşlem": tip, "Yön": "Giriş", "Tutar": tut, "Açıklama": "-", "created_at": i.get('created_at'), "updated_at": i.get('updated_at')})
                elif tip == 'Eksik':
                    kasa_dokum.append({"Tarih": i['tarih'], "Kasa": g, "İşlem": tip, "Yön": "Çıkış", "Tutar": tut, "Açıklama": "Sayım Eksiği", "created_at": i.get('created_at'), "updated_at": i.get('updated_at')})
                elif tip == 'Fazla':
                    kasa_dokum.append({"Tarih": i['tarih'], "Kasa": a, "İşlem": tip, "Yön": "Giriş", "Tutar": tut, "Açıklama": "Sayım Fazlası", "created_at": i.get('created_at'), "updated_at": i.get('updated_at')})
                elif tip == 'Bankaya Yatırılan':
                    kasa_dokum.append({"Tarih": i['tarih'], "Kasa": g, "İşlem": tip, "Yön": "Çıkış", "Tutar": tut, "Açıklama": "-", "created_at": i.get('created_at'), "updated_at": i.get('updated_at')})
                elif tip == 'Virman':
                    kasa_dokum.append({"Tarih": i['tarih'], "Kasa": g, "İşlem": "Virman Çıkışı", "Yön": "Çıkış", "Tutar": tut, "Açıklama": f"Alıcı: {a}", "created_at": i.get('created_at'), "updated_at": i.get('updated_at')})
                    kasa_dokum.append({"Tarih": i['tarih'], "Kasa": a, "İşlem": "Virman Girişi", "Yön": "Giriş", "Tutar": tut, "Açıklama": f"Gönderen: {g}", "created_at": i.get('created_at'), "updated_at": i.get('updated_at')})
                    
        if kasa_dokum:
            df_kd = pd.DataFrame(kasa_dokum)
            df_kd['Tarih'] = pd.to_datetime(df_kd['Tarih']).dt.date
            z_goster_k = st.toggle("⏱️ Kayıt (İşlenme) Zamanlarını Göster", key="kasa_zg")
            
            with st.expander("🔍 Kasa Filtreleme Paneli", expanded=True):
                f1, f2, f3 = st.columns(3)
                with f1: tar_aralik_k = st.date_input("Tarih Aralığı", [df_kd['Tarih'].min(), df_kd['Tarih'].max()], key="kd_tar")
                with f2: sec_kasa_k = st.multiselect("Kasa Seç", df_kd['Kasa'].unique().tolist(), key="kd_kasa")
                with f3: sec_islem_k = st.multiselect("İşlem Tipi", df_kd['İşlem'].unique().tolist(), key="kd_islem")
                
            if len(tar_aralik_k) == 2: df_kd = df_kd[(df_kd['Tarih'] >= tar_aralik_k[0]) & (df_kd['Tarih'] <= tar_aralik_k[1])]
            elif len(tar_aralik_k) == 1: df_kd = df_kd[df_kd['Tarih'] == tar_aralik_k[0]]
            if sec_kasa_k: df_kd = df_kd[df_kd['Kasa'].isin(sec_kasa_k)]
            if sec_islem_k: df_kd = df_kd[df_kd['İşlem'].isin(sec_islem_k)]
            
            gosterim_kasa = ['Tarih', 'Kasa', 'İşlem', 'Yön', 'Tutar', 'Açıklama']
            if z_goster_k: df_kd, gosterim_kasa = zaman_sutunlari_ekle(df_kd, gosterim_kasa)
            
            st.dataframe(df_kd[gosterim_kasa].sort_values("Tarih", ascending=False), hide_index=True, use_container_width=True)
            dosya_kd, uz_kd, mi_kd = excel_indir(df_kd[gosterim_kasa])
            st.download_button("📥 Kasa Dökümünü Excel'e İndir", data=dosya_kd, file_name=f"Kasa_Hareketleri.{uz_kd}", mime=mi_kd, key="dl_kasa")
        else:
            st.info("Kayıtlı kasa hareketi bulunmuyor.")

elif menu == "Personel & Puantaj":
    st.header("👥 Personel, İzin ve Maaş Yönetimi")
    bildirim_goster()
    
    alt_menu = st.radio(
        "İşlem Seçin", 
        ["📝 Puantaj Girişi", "📋 Filtreli Geçmiş Kayıtlar", "💰 Maaş Hesaplama", "⚙️ Personel Yönetimi"], 
        horizontal=True, 
        label_visibility="collapsed"
    )
    st.markdown("---")
    
    if alt_menu == "⚙️ Personel Yönetimi":
        st.subheader("Sisteme Yeni Personel Ekle")
        with st.form("personel_ekle_form"):
            col1, col2 = st.columns(2)
            with col1:
                yeni_personel = st.text_input("Personel Adı Soyadı")
                yeni_tarih = st.date_input("İşe Başlama Tarihi", datetime.date.today())
            with col2:
                yeni_maas = st.number_input("Aylık Net Maaşı (₺)", min_value=0.0)
                yeni_izin_hakki = st.number_input("Yıllık İzin Hakkı (Gün)", min_value=0.0, step=1.0)
                
            if st.form_submit_button("Ekle"):
                if yeni_personel.strip():
                    if db_yaz(supabase.table("personeller").insert({"isim": yeni_personel.strip(), "maas": yeni_maas, "ise_baslama_tarihi": str(yeni_tarih), "yillik_izin_hakki": yeni_izin_hakki})):
                        st.session_state.genel_mesaj = ("success", "Personel eklendi!")
                    st.rerun()
                    
        st.divider()
        st.subheader("Personel Listesi ve İzin Hakları")
        personel_listesi = db_oku(supabase.table("personeller").select("*"))
        tum_puantaj_izin = db_oku(supabase.table("puantaj").select("personel_adi").eq("durum", "Yıllık İzin"))
        
        if personel_listesi:
            izin_kullanim = {}
            if tum_puantaj_izin:
                for p_izin in tum_puantaj_izin:
                    izin_kullanim[p_izin['personel_adi']] = izin_kullanim.get(p_izin['personel_adi'], 0) + 1
                    
            pers_tablo = []
            for p in personel_listesi:
                hak = float(p.get('yillik_izin_hakki', 0))
                kull = izin_kullanim.get(p['isim'], 0)
                pers_tablo.append({
                    "Personel Adı": p['isim'],
                    "İşe Başlama": p.get('ise_baslama_tarihi', ''),
                    "Aylık Maaş": f"{float(p.get('maas', 0)):,.2f} ₺",
                    "Tanımlı İzin": hak,
                    "Kullanılan İzin": kull,
                    "Kalan İzin": max(0, hak - kull)
                })
            st.dataframe(pd.DataFrame(pers_tablo), hide_index=True, use_container_width=True)

        st.divider()
        with st.expander("✏️ Personel Düzenle veya Sil", expanded=False):
            if personel_listesi:
                sec_pers_str = st.selectbox("Seç", ["Seç..."] + [p['isim'] for p in personel_listesi])
                if sec_pers_str != "Seç...":
                    sec_pers = next(p for p in personel_listesi if p['isim'] == sec_pers_str)
                    with st.form("p_duz"):
                        y_isim = st.text_input("Adı", value=sec_pers['isim'])
                        c1, c2 = st.columns(2)
                        with c1:
                            try: p_tar = datetime.datetime.strptime(sec_pers.get('ise_baslama_tarihi', str(datetime.date.today())), '%Y-%m-%d').date()
                            except: p_tar = datetime.date.today()
                            y_tarih = st.date_input("Başlama Tarihi", value=p_tar)
                            y_izin_hakki = st.number_input("Yıllık İzin Hakkı (Gün)", value=float(sec_pers.get('yillik_izin_hakki', 0.0)))
                        with c2:
                            y_maas = st.number_input("Maaş (₺)", value=float(sec_pers.get('maas', 0.0)))
                            st.markdown("<br>", unsafe_allow_html=True)
                            
                        cg, cs = st.columns(2)
                        with cg:
                            if st.form_submit_button("Güncelle"):
                                db_yaz(supabase.table("personeller").update({"isim": y_isim, "maas": y_maas, "ise_baslama_tarihi": str(y_tarih), "yillik_izin_hakki": y_izin_hakki}).eq("id", sec_pers['id']))
                                db_yaz(supabase.table("puantaj").update({"personel_adi": y_isim}).eq("personel_adi", sec_pers['isim']))
                                st.session_state.genel_mesaj = ("success", "Personel güncellendi!")
                                st.rerun()
                        with cs:
                            if st.form_submit_button("Sil"):
                                if db_yaz(supabase.table("personeller").delete().eq("id", sec_pers['id'])):
                                    st.session_state.genel_mesaj = ("info", "Personel silindi!")
                                st.rerun()

    elif alt_menu == "📝 Puantaj Girişi":
        personel_listesi = db_oku(supabase.table("personeller").select("*"))
        if personel_listesi:
            c1, c2 = st.columns(2)
            with c1:
                st.date_input("Tarih", datetime.date.today(), key="puantaj_tarih")
                st.selectbox("Personel", [p['isim'] for p in personel_listesi], key="puantaj_isim")
            with c2:
                st.selectbox("Durum", ["Tam Gün", "Yarım Gün", "Haftalık İzin", "Yıllık İzin", "Ücretsiz İzin", "Raporlu", "Gelmedi"], key="puantaj_durum")
                st.number_input("Mesai (Saat)", min_value=0.0, step=0.5, key="puantaj_mesai")
            st.button("Kaydet", on_click=puantaj_kaydet_cb, type="primary")

            st.divider()
            with st.expander("✏️ Puantaj Düzenle/Sil", expanded=False):
                tum_puantaj = db_oku(supabase.table("puantaj").select("*").order("tarih", desc=True))
                if tum_puantaj:
                    secenekler_p = {f"{p['tarih']} | {p.get('personel_adi','')} | {p.get('durum','')} (ID: {p['id']})": p for p in tum_puantaj}
                    secilen_p_str = st.selectbox("Kayıt Seç", ["Seçiniz..."] + list(secenekler_p.keys()))
                    if secilen_p_str != "Seçiniz...":
                        secilen_p = secenekler_p[secilen_p_str]
                        with st.form("puantaj_duz"):
                            try: y_tar = datetime.datetime.strptime(secilen_p['tarih'], '%Y-%m-%d').date()
                            except: y_tar = datetime.date.today()
                            y_tarih = st.date_input("Tarih", value=y_tar)
                            
                            p_isim_val = secilen_p.get('personel_adi')
                            p_isim_list = [pr['isim'] for pr in personel_listesi]
                            if p_isim_val not in p_isim_list: p_isim_val = p_isim_list[0]
                            y_isim = st.selectbox("Personel", p_isim_list, index=p_isim_list.index(p_isim_val))
                            
                            durumlar = ["Tam Gün", "Yarım Gün", "Haftalık İzin", "Yıllık İzin", "Ücretsiz İzin", "Raporlu", "Gelmedi"]
                            d_val = secilen_p.get('durum', 'Tam Gün')
                            if d_val not in durumlar: d_val = "Tam Gün"
                            y_durum = st.selectbox("Durum", durumlar, index=durumlar.index(d_val))
                            
                            y_mesai = st.number_input("Mesai", value=float(secilen_p.get('fazla_mesai_saati', 0)))
                            
                            cg, cs = st.columns(2)
                            with cg:
                                if st.form_submit_button("Güncelle"):
                                    hata_var = False
                                    if y_durum == "Yıllık İzin" and secilen_p.get('durum') != "Yıllık İzin":
                                        p_bilgi = next((p for p in personel_listesi if p['isim'] == y_isim), None)
                                        i_hakki = float(p_bilgi.get('yillik_izin_hakki', 0)) if p_bilgi else 0
                                        kull_izinler = db_oku(supabase.table("puantaj").select("id").eq("personel_adi", y_isim).eq("durum", "Yıllık İzin"))
                                        if len(kull_izinler) >= i_hakki:
                                            st.error("Bu personelin yıllık izin hakkı kalmamıştır!")
                                            hata_var = True
                                            
                                    if y_durum == "Haftalık İzin" and secilen_p.get('durum') != "Haftalık İzin" and not hata_var:
                                        h_baslangic = y_tarih - datetime.timedelta(days=y_tarih.weekday())
                                        h_bitis = h_baslangic + datetime.timedelta(days=6)
                                        sorgu_h = supabase.table("puantaj").select("id").eq("personel_adi", y_isim).eq("durum", "Haftalık İzin").gte("tarih", str(h_baslangic)).lte("tarih", str(h_bitis))
                                        kull_haftalik = db_oku(sorgu_h)
                                        if kull_haftalik:
                                            st.error("Bu personel ilgili hafta zaten Haftalık İzin kullanmış!")
                                            hata_var = True
                                            
                                    if not hata_var:
                                        if db_yaz(supabase.table("puantaj").update({"tarih": str(y_tarih), "personel_adi": y_isim, "durum": y_durum, "fazla_mesai_saati": y_mesai}).eq("id", secilen_p['id'])):
                                            st.session_state.genel_mesaj = ("success", "Puantaj güncellendi!")
                                        st.rerun()
                            with cs:
                                if st.form_submit_button("Sil"):
                                    if db_yaz(supabase.table("puantaj").delete().eq("id", secilen_p['id'])):
                                        st.session_state.genel_mesaj = ("info", "Puantaj silindi!")
                                    st.rerun()

    elif alt_menu == "📋 Filtreli Geçmiş Kayıtlar":
        st.subheader("Geçmiş Puantaj ve İzin Kayıtları")
        puantajlar = db_oku(supabase.table("puantaj").select("*").order("tarih", desc=True))
        if puantajlar:
            z_goster_p = st.toggle("⏱️ Kayıt (İşlenme) Zamanlarını Göster", key="puantaj_zg")
            df_puantaj = pd.DataFrame(puantajlar)
            df_puantaj['tarih'] = pd.to_datetime(df_puantaj['tarih']).dt.date
            
            with st.expander("🔍 Detaylı Filtreleme Paneli", expanded=True):
                c1, c2, c3 = st.columns(3)
                with c1:
                    t_aralik = st.date_input("Tarih Aralığı Seç", [df_puantaj['tarih'].min(), df_puantaj['tarih'].max()])
                with c2:
                    secili_pers = st.multiselect("Personel Seç", df_puantaj['personel_adi'].unique().tolist())
                with c3:
                    secili_durum = st.multiselect("Durum Seç (Örn: Haftalık İzin, Yıllık İzin)", df_puantaj['durum'].unique().tolist())
            
            if len(t_aralik) == 2: df_puantaj = df_puantaj[(df_puantaj['tarih'] >= t_aralik[0]) & (df_puantaj['tarih'] <= t_aralik[1])]
            elif len(t_aralik) == 1: df_puantaj = df_puantaj[df_puantaj['tarih'] == t_aralik[0]]
            
            if secili_pers: df_puantaj = df_puantaj[df_puantaj['personel_adi'].isin(secili_pers)]
            if secili_durum: df_puantaj = df_puantaj[df_puantaj['durum'].isin(secili_durum)]
            
            gosterim_puan = ['tarih', 'personel_adi', 'durum', 'fazla_mesai_saati']
            if z_goster_p: df_puantaj, gosterim_puan = zaman_sutunlari_ekle(df_puantaj, gosterim_puan)
            
            st.dataframe(df_puantaj[gosterim_puan], hide_index=True, use_container_width=True)
            st.info(f"📊 Ekranda filtrelenen toplam kayıt sayısı: **{len(df_puantaj)}**")
            
            dosya_p, uzanti_p, mime_p = excel_indir(df_puantaj[gosterim_puan])
            st.download_button(label="📥 Filtrelenmiş Kayıtları Excel'e İndir", data=dosya_p, file_name=f"Puantaj_Raporu.{uzanti_p}", mime=mime_p)

    elif alt_menu == "💰 Maaş Hesaplama":
        st.subheader("Aylık Maaş ve Mesai Hesaplama")
        st.info("💡 **Bilgi:** Maaş kesinti üzerinden değil; geldiği ve hak ettiği günler üzerinden hesaplanır (SGK Standartları). Tam Gün, Haftalık İzin ve Yıllık İzin = **1 Gün**. Yarım Gün = **0.5 Gün**.")
        
        aylar = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
        col1, col2 = st.columns(2)
        with col1: secilen_ay = st.selectbox("Hesaplanacak Ay", aylar, index=datetime.date.today().month - 1)
        with col2: secilen_yil = st.number_input("Yıl", value=datetime.date.today().year, step=1)
        ay_index = aylar.index(secilen_ay) + 1
        
        if st.button("Maaşları Hesapla", type="primary"):
            tum_personel = db_oku(supabase.table("personeller").select("*"))
            tum_puantaj = db_oku(supabase.table("puantaj").select("*"))
            if tum_personel and tum_puantaj:
                df_p = pd.DataFrame(tum_puantaj)
                df_p['tarih'] = pd.to_datetime(df_p['tarih'])
                df_ay = df_p[(df_p['tarih'].dt.month == ay_index) & (df_p['tarih'].dt.year == secilen_yil)]
                
                hesap_listesi = []
                for pers in tum_personel:
                    isim = pers['isim']
                    maas = float(pers.get('maas', 0.0))
                    if maas > 0:
                        gunluk_ucret = maas / 30
                        saatlik_ucret = maas / 225
                        mesai_saatlik_ucret = saatlik_ucret * 1.5
                        
                        pers_puantaj = df_ay[df_ay['personel_adi'] == isim]
                        
                        tam_gun = len(pers_puantaj[pers_puantaj['durum'] == 'Tam Gün'])
                        h_izin = len(pers_puantaj[pers_puantaj['durum'] == 'Haftalık İzin'])
                        y_izin = len(pers_puantaj[pers_puantaj['durum'] == 'Yıllık İzin'])
                        yarim_gun = len(pers_puantaj[pers_puantaj['durum'] == 'Yarım Gün'])
                        
                        odenecek_gun = tam_gun + h_izin + y_izin + (yarim_gun * 0.5)
                        
                        _, aydaki_gun_sayisi = calendar.monthrange(secilen_yil, ay_index)
                        if aydaki_gun_sayisi in [28, 29] and odenecek_gun == aydaki_gun_sayisi:
                            odenecek_gun = 30
                        if odenecek_gun > 30:
                            odenecek_gun = 30
                            
                        hakedis_maas = odenecek_gun * gunluk_ucret
                        toplam_mesai = pd.to_numeric(pers_puantaj['fazla_mesai_saati'], errors='coerce').fillna(0).sum()
                        mesai_tutari = toplam_mesai * mesai_saatlik_ucret
                        net_odenecek = hakedis_maas + mesai_tutari
                        
                        hesap_listesi.append({
                            "Personel": isim, 
                            "Kök Maaş": f"{maas:,.2f} ₺", 
                            "Ödenecek Gün": f"{odenecek_gun} Gün", 
                            "Hak Ediş": f"{hakedis_maas:,.2f} ₺", 
                            "Mesai Süresi": f"{toplam_mesai} Saat", 
                            "Mesai Ücreti": f"+{mesai_tutari:,.2f} ₺",
                            "Net Ödenecek": f"{net_odenecek:,.2f} ₺"
                        })
                
                if hesap_listesi: st.dataframe(pd.DataFrame(hesap_listesi), hide_index=True, use_container_width=True)
                else: st.warning("Bu ay için puantaj işlemi görmüş kayıtlı personel bulunamadı.")
            else:
                st.warning("Henüz puantaj kaydı bulunmuyor.")

elif menu == "Raporlar":
    st.header("📊 Sistem Raporları ve Excel Çıktıları")
    z_goster_r = st.toggle("⏱️ Kayıt (İşlenme) Zamanlarını Tüm Tablolarda Göster", key="rapor_zg")
    
    st.subheader("1. Yemek Sepeti ve Trendyol Satış Raporu")
    sat = db_oku(supabase.table("platform_satis").select("*"))
    if sat: 
        df_sat = pd.DataFrame(sat)
        for col in ['platform', 'odeme_tipi', 'brut', 'komisyon_tutari', 'stopaj_tutari', 'net', 'tahsilat_tarihi', 'durum']:
            if col not in df_sat.columns: df_sat[col] = 0.0 if 'tutari' in col or col in ['brut', 'net'] else ''
            
        df_sat = df_sat.sort_values(by="tarih", ascending=False)
        gosterim_sat = ['tarih', 'platform', 'odeme_tipi', 'brut', 'komisyon_tutari', 'stopaj_tutari', 'net', 'tahsilat_tarihi', 'durum']
        if z_goster_r: df_sat, gosterim_sat = zaman_sutunlari_ekle(df_sat, gosterim_sat)
            
        st.dataframe(df_sat[gosterim_sat], hide_index=True, use_container_width=True)
        dosya_sat, uzanti_sat, mime_sat = excel_indir(df_sat[gosterim_sat])
        st.download_button(label="📥 Platform Satışlarını Excel'e İndir", data=dosya_sat, file_name=f"Platform_Satislar_Raporu.{uzanti_sat}", mime=mime_sat)
    
    st.divider()
    
    st.subheader("2. Günlük Dükkan Cirosu Raporu")
    cir = db_oku(supabase.table("ciro").select("*"))
    if cir: 
        df_cir = pd.DataFrame(cir)
        for col in ['kasa', 'nakit', 'kredi_karti', 'pavo_nakit', 'pavo_kredi', 'odenmez']:
            if col not in df_cir.columns: df_cir[col] = 0.0 if col != 'kasa' else 'Kasa 1'
            
        df_cir = df_cir.sort_values(by="tarih", ascending=False)
        gosterim_cir = ['tarih', 'kasa', 'nakit', 'kredi_karti', 'pavo_nakit', 'pavo_kredi', 'odenmez']
        if z_goster_r: df_cir, gosterim_cir = zaman_sutunlari_ekle(df_cir, gosterim_cir)
            
        st.dataframe(df_cir[gosterim_cir], hide_index=True, use_container_width=True)
        dosya_cir, uzanti_cir, mime_cir = excel_indir(df_cir[gosterim_cir])
        st.download_button(label="📥 Dükkan Cirosunu Excel'e İndir", data=dosya_cir, file_name=f"Dukkan_Cirosu_Raporu.{uzanti_cir}", mime=mime_cir)

    st.divider()

    st.subheader("3. Tüm Masraflar Raporu")
    masraflar_r = db_oku(supabase.table("masraf").select("*").order("tarih", desc=True))
    if masraflar_r:
        df_masraf_r = pd.DataFrame(masraflar_r)
        for col in ['masraf_tipi', 'aciklama', 'tutar', 'odeme_tipi']:
            if col not in df_masraf_r.columns: df_masraf_r[col] = 0.0 if col == 'tutar' else ''
            
        df_masraf_r['masraf_tipi'] = df_masraf_r.get('masraf_tipi', 'Genel').fillna('Genel Masraf')
        gosterim_mas = ['tarih', 'masraf_tipi', 'aciklama', 'tutar', 'odeme_tipi']
        if z_goster_r: df_masraf_r, gosterim_mas = zaman_sutunlari_ekle(df_masraf_r, gosterim_mas)
            
        st.dataframe(df_masraf_r[gosterim_mas], hide_index=True, use_container_width=True)
        dosya_mas, uzanti_mas, mime_mas = excel_indir(df_masraf_r[gosterim_mas])
        st.download_button(label="📥 Tüm Masrafları Excel'e İndir", data=dosya_mas, file_name=f"Masraflar_Raporu.{uzanti_mas}", mime=mime_mas)
