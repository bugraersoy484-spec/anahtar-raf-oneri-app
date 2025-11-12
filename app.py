import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl.chart import BarChart, Reference

st.set_page_config(page_title="Anahtar Raf Öneri Sistemi (Grup Bazlı, No'dan Öneri)", layout="centered")
st.title("🔑 Anahtar Raf Öneri Sistemi")
st.markdown("Her anahtar için `ANAHTAR` sayfasındaki **B (No)** sütununu baz alarak o gruptan en az dolu rafı önerir.")

uploaded_file = st.file_uploader("📂 Lütfen Excel dosyanı yükle (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        stok_df = pd.read_excel(uploaded_file, sheet_name="STOK")
        anahtar_df = pd.read_excel(uploaded_file, sheet_name="ANAHTAR")

        stok = stok_df.copy()
        anahtar = anahtar_df.copy()

        # Grup kodunu çıkar (örnek: 001A -> 001). Grup string olarak tutulur (başındaki sıfırlar korunur)
        stok["Grup"] = stok["Raf Bilgisi"].astype(str).str.extract(r"(\d+)")
        # Temizle (eğer herhangi bir NaN varsa stringe çevir)
        stok["Grup"] = stok["Grup"].fillna("").astype(str)

        # Rafları grup ve raf bilgisine göre sırala
        stok = stok.sort_values(by=["Grup", "Raf Bilgisi"]).reset_index(drop=True)

        # Hazırla
        onerilen_raf = []

        # Ön hesap: grup listesi
        mevcut_gruplar = set(stok["Grup"].unique())

        # Döngü: anahtardaki her satır için öneri üret
        for idx in anahtar.index:
            # Öncelikle kullanıcının verdiği 'No' sütununu al (B sütunu). Sütun adı farklıysa 'No' yerine uygun ismi kullan.
            kullanici_no = None
            if "No" in anahtar.columns:
                val = anahtar.loc[idx, "No"]
                # NaN kontrolü
                if pd.notna(val):
                    # stringe çevir ve trim
                    kullanici_no = str(val).strip()
                    # Bazı hücrelerde ondalık/float gelebilir (ör. 1.0) -> formatla 3 haneli gibi bırakma, kullanıcının verdiği hali kullan
                    # Eğer kullanıcı 1 yazdıysa '1' olur; stok grup '001' ise birebir eşleşme olmaz -> bu yüzden iki ihtimali kontrol edeceğiz.
            # Hedef grup kararı:
            hedef_grup = None

            # 1) Eğer kullanici_no varsa, dene doğrudan eşleşme
            if kullanici_no:
                if kullanici_no in mevcut_gruplar:
                    hedef_grup = kullanici_no
                else:
                    # stok'ta grup '001' şeklindeyse ve kullanıcı '1' yazdıysa bunu '001' ile eşleştirmeye çalış
                    # En uzun grup uzunluğunu al (ör: stokta '001' gibi 3 haneli olabilir)
                    # Burada mantık: kullanıcı '1' yazdıysa gruplarda son olarak eşleşen grup olacak şekilde genişletmeye çalış.
                    # Basit yaklaşım: stok'taki herhangi bir grup stringinin sonu kullanıcı_no ile bitiyorsa eşleştir.
                    for g in mevcut_gruplar:
                        if g.endswith(kullanici_no):
                            hedef_grup = g
                            break
            # 2) Eğer hedef_grup hala None -> fallback: en az dolu grup (önceki mantık)
            if hedef_grup is None:
                grup_toplam = stok.groupby("Grup")["Raftaki Adet"].sum()
                # Grup indeksleri boş string (""), NaN vs olabilir; filtrele boş olanları eğer gerekliyse
                grup_toplam = grup_toplam.drop(labels=[""], errors="ignore")
                if len(grup_toplam) > 0:
                    hedef_grup = grup_toplam.idxmin()
                else:
                    # Eğer hiç grup yoksa (olağan dışı) direk en az dolu raf genelinde seç
                    min_idx_genel = stok["Raftaki Adet"].idxmin()
                    min_raf_bilgisi = stok.loc[min_idx_genel, "Raf Bilgisi"]
                    onerilen_raf.append(min_raf_bilgisi)
                    stok.loc[min_idx_genel, "Raftaki Adet"] += 1
                    continue  # sonraki anahtara geç

            # Hedef grubun içindeki rafları al
            alt_raflar = stok[stok["Grup"] == hedef_grup]
            if alt_raflar.empty:
                # beklenmedik durum: hedef grup yoksa fallback genel en az dolu raf
                min_idx_genel = stok["Raftaki Adet"].idxmin()
                min_raf_bilgisi = stok.loc[min_idx_genel, "Raf Bilgisi"]
                onerilen_raf.append(min_raf_bilgisi)
                stok.loc[min_idx_genel, "Raftaki Adet"] += 1
            else:
                # grup içinden en az dolu rafı seç
                min_idx = alt_raflar["Raftaki Adet"].idxmin()
                min_raf_bilgisi = stok.loc[min_idx, "Raf Bilgisi"]
                onerilen_raf.append(min_raf_bilgisi)
                stok.loc[min_idx, "Raftaki Adet"] += 1

        # Anahtar tablosuna önerileri ekle
        anahtar["Önerilen Raf"] = onerilen_raf

        # Kolon isimlendirme ve temizlik (isteğe göre)
        if "Yeni Raf" in anahtar.columns:
            anahtar.rename(columns={"Yeni Raf": "Kullanıcı Örneği Raf"}, inplace=True)

        if "No" not in anahtar.columns:
            st.warning("Not: 'ANAHTAR' sayfasında 'No' sütunu bulunamadı; grup bilgisi yoksa otomatik dengeleme yapıldı.")

        # Doluluk oranı hesapla
        max_capacity = stok["Raftaki Adet"].max()
        stok["Doluluk (%)"] = (stok["Raftaki Adet"] / max_capacity * 100).round(1) if max_capacity > 0 else 0

        # Özet oluştur
        ozet_data = {
            "Toplam Raf Sayısı": [len(stok)],
            "Toplam Grup Sayısı": [stok["Grup"].nunique()],
            "Toplam Anahtar Sayısı (Güncel)": [stok["Raftaki Adet"].sum()],
            "Yeni Eklenen Anahtar Sayısı": [len(anahtar)]
        }
        ozet_df = pd.DataFrame(ozet_data)
        doluluk_sirali = stok.sort_values(by="Raftaki Adet", ascending=False).reset_index(drop=True)

        # Excel çıktısı oluştur
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            stok.to_excel(writer, index=False, sheet_name="STOK_GUNCEL")
            anahtar.to_excel(writer, index=False, sheet_name="ANAHTAR_ONERI")
            ozet_df.to_excel(writer, index=False, sheet_name="OZET")
            doluluk_sirali.to_excel(writer, index=False, sheet_name="RAF_DOLULUK_SIRALAMA")

            wb = writer.book
            ws = wb["RAF_DOLULUK_SIRALAMA"]
            chart = BarChart()
            chart.title = "Raf Doluluk Oranları (Güncel Adet)"
            chart.x_axis.title = "Raf Bilgisi"
            chart.y_axis.title = "Anahtar Adedi"

            row_count = len(doluluk_sirali)
            cats = Reference(ws, min_col=1, min_row=2, max_row=row_count + 1)
            data = Reference(ws, min_col=2, min_row=1, max_row=row_count + 1)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            chart.height = 15
            chart.width = 30
            ws.add_chart(chart, "E2")

        st.success("✅ Öneriler hesaplandı — 'ANAHTAR_ONERI' sayfasını indirin.")

        st.download_button(
            label="💾 Sonuç Excel Dosyasını İndir",
            data=output.getvalue(),
            file_name="anahtar_raf_oneri_grafikli.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.subheader("📊 Özet Bilgiler")
        st.dataframe(ozet_df)

    except Exception as e:
        st.error(f"Hata oluştu: {e}")
