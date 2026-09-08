import pandas as pd
import psycopg2
from datetime import datetime

# ==========================================
# ⚙️ AYARLAR (BURALARI KENDİNE GÖRE DOLDUR)
# ==========================================
# 1. Netsis'in her gün raporu kaydettiği veya senin koyduğun Excel dosyasının tam yolu
DOSYA_YOLU = r"C:\Users\KullaniciAdi\Desktop\Tahsilat_Raporu.xlsx"

# 2. Supabase Veritabanı Linkin (Streamlit'teki DB_URL ile aynı)
DB_URL = "postgres://kullaniciadi:sifre@aws-0-eu-central-1.pooler.supabase.com:6543/postgres?sslmode=require"

# ==========================================
# 🤖 ROBOTUN MOTORU
# ==========================================
def bakiye_temizle(deger):
    if pd.isna(deger): return 0.0
    if isinstance(deger, (int, float)): return float(deger)
    try:
        temiz = str(deger).replace(' TL', '').replace('₺', '').strip()
        if '.' in temiz and ',' in temiz: temiz = temiz.replace('.', '').replace(',', '.')
        elif ',' in temiz: temiz = temiz.replace(',', '.')
        return float(temiz)
    except ValueError: return 0.0

print(f"[{datetime.now().strftime('%H:%M:%S')}] 🤖 Masaüstü Robotu Çalıştı!")
print(f"📂 Dosya okunuyor: {DOSYA_YOLU}")

try:
    df = pd.read_excel(DOSYA_YOLU)
    
    conn = psycopg2.connect(DB_URL)
    c = conn.cursor()
    
    # Mevcut takipteki carilerin bakiyelerini alalım (Değişim logu için)
    c.execute("SELECT kod, bakiye FROM takip")
    takip_dict = {row[0]: float(row[1]) if row[1] is not None else 0.0 for row in c.fetchall()}
    
    c.execute("DELETE FROM ana_liste")
    
    sayac = 0
    guncellenen = 0
    zaman_simdi = datetime.now().strftime("%d.%m.%Y %H:%M")
    
    for index, row in df.iterrows():
        c_isim = str(row.get("Cari İsim", ""))
        if pd.isna(row.get("Cari İsim")) or not c_isim.strip(): continue
        
        c_kod = str(row.get("Cari Kod", "")).strip()
        if c_kod == "nan" or c_kod == "-" or not c_kod:
            c_kod = f"KODSUZ-{index}-{datetime.now().strftime('%H%M%S')}"
            
        bakiye_val = bakiye_temizle(row.get("Borç Bak.", 0.0))
        c_tel = str(row.get("Telefon", "")) if pd.notna(row.get("Telefon")) else ""
        
        # Eğer bu kişi TAKİP LİSTESİNDEYSE bakiyesini ve logunu güncelle
        if c_kod in takip_dict:
            eski_bakiye = takip_dict[c_kod]
            if abs(eski_bakiye - bakiye_val) > 0.01:
                c.execute("UPDATE takip SET bakiye=%s WHERE kod=%s", (bakiye_val, c_kod))
                log_mesaji = f"Sistem: Bilgisayardaki Excel'den bakiye güncellendi ({eski_bakiye:,.2f} TL ➔ {bakiye_val:,.2f} TL)"
                c.execute("INSERT INTO loglar (cari_kod, tarih_saat, not_metni) VALUES (%s, %s, %s)", (c_kod, zaman_simdi, log_mesaji))
                guncellenen += 1
            continue
            
        # TAKİP LİSTESİNDE DEĞİLSE, Ana Listeye Ekle
        c.execute("""
            INSERT INTO ana_liste (kod, isim, telefon, bakiye) VALUES (%s, %s, %s, %s)
            ON CONFLICT (kod) DO UPDATE SET isim=EXCLUDED.isim, telefon=EXCLUDED.telefon, bakiye=EXCLUDED.bakiye
        """, (c_kod, c_isim, c_tel, bakiye_val))
        sayac += 1

    conn.commit()
    c.close()
    conn.close()
    
    print(f"✅ İŞLEM BAŞARILI! {sayac} kişi ana listeye eklendi.")
    print(f"✅ Takipteki {guncellenen} kişinin bakiyesi güncellendi.")

except FileNotFoundError:
    print(f"❌ HATA: {DOSYA_YOLU} adresinde Excel dosyası bulunamadı!")
except Exception as e:
    print(f"❌ BEKLENMEYEN HATA: {e}")
