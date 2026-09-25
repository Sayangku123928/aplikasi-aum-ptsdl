import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="AUM PTSdL",
    page_icon="📋",
    layout="wide"
)

st.title("📋 Aplikasi Pengolahan AUM PTSdL")
st.write("Upload file Excel hasil Google Form untuk mengolah jawaban AUM PTSdL.")

st.subheader("1. Upload Data AUM PTSdL")

file = st.file_uploader(
    "Pilih file Excel hasil Google Form",
    type=["xlsx", "xls"]
)

if file is not None:

    try:
        df = pd.read_excel(file)

        st.success("File berhasil dibaca!")

        st.write("Jumlah responden:", len(df))
        st.write("Jumlah kolom:", len(df.columns))

        st.subheader("2. Data Asli")
        st.dataframe(df.head())

        # Pilihan jawaban AUM
        pilihan = {
            "jarang": 1,
            "kadang-kadang": 2,
            "kadang kadang": 2,
            "sering": 3,
            "pada umumnya": 4,
            "pada umumnya ": 4,
            "selalu": 5
        }

        # Mencari kolom jawaban AUM
        kolom_aum = []

        for kolom in df.columns:

            nilai = df[kolom].astype(str).str.lower().str.strip()

            jumlah = nilai.isin(pilihan.keys()).sum()

            if jumlah > 0:
                kolom_aum.append(kolom)

        st.write("Kolom jawaban AUM yang terdeteksi:", len(kolom_aum))

        if len(kolom_aum) == 0:

            st.error(
                "Jawaban AUM tidak terdeteksi. "
                "Pastikan isi Excel menggunakan pilihan Jarang, "
                "Kadang-kadang, Sering, Pada umumnya, atau Selalu."
            )

        else:

            # Mengubah jawaban menjadi angka
            data_skor = df[kolom_aum].copy()

            for kolom in data_skor.columns:

                data_skor[kolom] = (
                    data_skor[kolom]
                    .astype(str)
                    .str.lower()
                    .str.strip()
                    .map(pilihan)
                )

            st.subheader("3. Hasil Konversi Jawaban")

            st.write(
                "Jarang = 1 | Kadang-kadang = 2 | Sering = 3 | "
                "Pada umumnya = 4 | Selalu = 5"
            )

            st.dataframe(data_skor.head())

            # Jumlah butir
            jumlah_butir = len(kolom_aum)

            st.subheader("4. Ringkasan")

            st.write("Jumlah responden:", len(df))
            st.write("Jumlah butir terdeteksi:", jumlah_butir)

            # Total dan rata-rata setiap responden
            total_skor = data_skor.sum(axis=1, skipna=True)
            rata_rata = data_skor.mean(axis=1, skipna=True)

            hasil = pd.DataFrame()

            # Cari kolom identitas
            kolom_nama = None

            for kolom in df.columns:

                nama = str(kolom).lower()

                if "nama" in nama:
                    kolom_nama = kolom
                    break

            if kolom_nama is not None:
                hasil["Nama"] = df[kolom_nama]

            hasil["Jumlah Butir Terisi"] = data_skor.notna().sum(axis=1)
            hasil["Total Skor"] = total_skor
            hasil["Rata-rata Skor"] = rata_rata.round(2)

            st.subheader("5. Hasil Pengolahan Per Responden")

            st.dataframe(hasil)

            # Download hasil
            output = BytesIO()

            with pd.ExcelWriter(output, engine="openpyxl") as writer:

                df.to_excel(
                    writer,
                    sheet_name="Data Asli",
                    index=False
                )

                data_skor.to_excel(
                    writer,
                    sheet_name="Data Skor",
                    index=False
                )

                hasil.to_excel(
                    writer,
                    sheet_name="Hasil Per Responden",
                    index=False
                )

            st.download_button(
                label="⬇️ Download Hasil Pengolahan Excel",
                data=output.getvalue(),
                file_name="Hasil_AUM_PTSdL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:

        st.error("File belum dapat diproses.")

        st.write("Pesan error:")
        st.code(str(e))
