import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference

st.set_page_config(page_title="Anahtar Raf Öneri Sistemi", layout="centered")

st.title("🔑 Anahtar Raf Öneri Sistemi")
st.markdown("Excel dosyanı yükle, sistem raf önerilerini otomatik hesaplasın.")

uploaded_file = st.file_uploader("📂 Lütfen Excel dosyanı yükle (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Excel dosyasını oku
        stok_df = pd.read_excel(uploaded_file, sheet_name="STOK")
        anahtar_df = pd.read_excel(uploaded_file, sheet_name="ANAHTAR")

        # Kopyalar
        stok = stok_df.copy()
        anahtar = anahtar_df.copy()

        # Raf sıralama
        stok = stok.sort_values(by="Raf Bilgisi").reset_index(drop=True)

        # Raf önerisi hesaplama
        onerilen_raf = []
        for _ in anahtar.index:
            min_idx = stok["Raftaki Adet"].idxmin()
            min_raf_bilgisi = stok.loc[min_idx, "Raf Bilgisi"]
            onerilen_raf.append(min_raf_bilgisi)
            stok.loc[min_idx, "Raftaki Adet"] += 1

        anahtar["Önerilen Raf"] = onerilen_raf

        if "Yeni Raf" in anahtar.columns:
            anahtar.rename(columns={"Yeni Raf": "Kullanıcı Örneği Raf"}, inplace=True)
        if "No" in anahtar.columns:
            anahtar.drop(columns=["No"], inplace=True)

        # Doluluk oranı
        max_capacity = stok["Raftaki Adet"].max()
        stok["Doluluk (%)"] = (stok["Raftaki Adet"] / max_capacity * 100).round(1) if max_capacity > 0 else 0

        # Özet tablo
        ozet_data = {
            "Toplam Raf Sayısı": [len(stok)],
            "Toplam Anahtar Sayısı (Güncel)": [stok["Raftaki Adet"].sum()],
            "Yeni Eklenen Anahtar Sayısı": [len(anahtar)],
            "En Dolu Raf": [stok.loc[stok["Raftaki Adet"].idxmax(), "Raf Bilgisi"]],
            "En Boş Raf": [stok.loc[stok["Raftaki Adet"].idxmin(), "Raf Bilgisi"]],
            "Doluluk Farkı (Max - Min Adet)": [stok["Raftaki Adet"].max() - stok["Raftaki Adet"].min()]
        }
        ozet_df = pd.DataFrame(ozet_data)

        doluluk_sirali = stok.sort_values(by="Raftaki Adet", ascending=False).reset_index(drop=True)

        # Sonuç Excel oluştur
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

        st.success("✅ Raf önerileri başarıyla hesaplandı!")

        # İndirilebilir dosya oluştur
        st.download_button(
            label="💾 Sonuç Excel Dosyasını İndir",
            data=output.getvalue(),
            file_name="anahtar_raf_oneri_grafikli.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Ek bilgi görüntüleme
        st.subheader("📊 Özet Bilgiler")
        st.dataframe(ozet_df)

    except Exception as e:
        st.error(f"Hata oluştu: {e}")
