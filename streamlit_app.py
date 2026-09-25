import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(
    page_title="AUM PTSdL",
    page_icon="📋",
    layout="wide"
)

st.title("📋 Aplikasi Pengolahan AUM PTSdL")
st.write("Pengolahan data AUM PTSdL dari Excel Google Form")


# ==============================
# FUNGSI MENGUBAH JAWABAN
# ==============================

def ubah_ke_skor(nilai):

    if pd.isna(nilai):
        return None

    teks = str(nilai).strip()

    # Kalau jawabannya sudah angka
    try:
        angka = float(teks)

        if angka in [1, 2, 3, 4, 5]:
            return int(angka)

    except:
        pass

    # Membaca angka di awal jawaban
    cocok = re.match(r"^\s*([1-5])", teks)

    if cocok:
        return int(cocok.group(1))

    # Kalau tidak ada angka, membaca kata
    teks = teks.lower()

    if "jarang" in teks:
        return 1

    if "kadang" in teks:
        return 2

    if "sering" in teks:
        return 3

    if "pada umumnya" in teks:
        return 4

    if "selalu" in teks:
        return 5

    return None


# ==============================
# UPLOAD EXCEL
# ==============================

st.subheader("📁 Upload Data AUM")

file = st.file_uploader(
    "Upload file Excel hasil Google Form",
    type=["xlsx"]
)


if file is not None:

    try:

        df = pd.read_excel(
            file,
            engine="openpyxl"
        )

        st.success("✅ File Excel berhasil dibaca!")

        st.write("Jumlah responden:", len(df))
        st.write("Jumlah kolom:", len(df.columns))


        # ==============================
        # TAMPILKAN DATA ASLI
        # ==============================

        st.subheader("Data Asli")

        st.dataframe(
            df.head(5),
            use_container_width=True
        )


        # ==============================
        # MENCARI KOLOM AUM
        # ==============================

        kolom_aum = []

        for kolom in df.columns:

            jumlah_terbaca = (
                df[kolom]
                .apply(ubah_ke_skor)
                .notna()
                .sum()
            )

            # Kolom dianggap kolom AUM
            # jika minimal ada jawaban yang bisa dibaca
            if jumlah_terbaca >= 1:
                kolom_aum.append(kolom)


        st.subheader("Pemeriksaan Data")

        st.write(
            "Kolom yang terdeteksi sebagai data AUM:",
            len(kolom_aum)
        )


        # ==============================
        # AMBIL 165 BUTIR TERAKHIR/RELEVAN
        # ==============================

        if len(kolom_aum) < 165:

            st.warning(
                "Aplikasi belum menemukan 165 butir AUM. "
                "Kita perlu melihat struktur kolom Excel kamu."
            )

            st.write("Kolom yang terdeteksi:")

            st.write(kolom_aum)

            st.stop()


        # Ambil 165 kolom AUM
        kolom_aum = kolom_aum[-165:]


        # ==============================
        # KONVERSI KE ANGKA
        # ==============================

        data_skor = df[kolom_aum].copy()

        for kolom in data_skor.columns:

            data_skor[kolom] = (
                data_skor[kolom]
                .apply(ubah_ke_skor)
            )


        st.success(
            "✅ 165 butir AUM berhasil ditemukan!"
        )


        # ==============================
        # TAMPILKAN HASIL KONVERSI
        # ==============================

        st.subheader("Hasil Konversi Jawaban")

        st.write(
            "1 = Jarang | "
            "2 = Kadang-kadang | "
            "3 = Sering | "
            "4 = Pada umumnya | "
            "5 = Selalu"
        )

        st.dataframe(
            data_skor.head(10),
            use_container_width=True
        )


        # ==============================
        # HASIL TOTAL PER SISWA
        # ==============================

        total = data_skor.sum(
            axis=1,
            skipna=True
        )

        rata = data_skor.mean(
            axis=1,
            skipna=True
        )

        hasil = pd.DataFrame()

        # Cari kolom nama
        kolom_nama = None

        for kolom in df.columns:

            nama = str(kolom).lower()

            if (
                "nama" in nama
                or "name" in nama
            ):
                kolom_nama = kolom
                break


        if kolom_nama is not None:

            hasil["Nama"] = df[kolom_nama]

        else:

            hasil["Responden"] = range(
                1,
                len(df) + 1
            )


        hasil["Jumlah Butir Terisi"] = (
            data_skor.notna().sum(axis=1)
        )

        hasil["Total Skor"] = total

        hasil["Rata-rata"] = rata.round(2)


        # ==============================
        # TAMPILKAN HASIL
        # ==============================

        st.subheader(
            "📊 Hasil Pengolahan"
        )

        st.dataframe(
            hasil,
            use_container_width=True
        )


        # ==============================
        # DOWNLOAD
        # ==============================

        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            data_skor.to_excel(
                writer,
                sheet_name="Data Skor",
                index=False
            )

            hasil.to_excel(
                writer,
                sheet_name="Hasil",
                index=False
            )


        st.download_button(
            label="⬇️ Download Hasil Excel",
            data=output.getvalue(),
            file_name="Hasil_AUM_PTSdL.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )


    except Exception as e:

        st.error(
            "Terjadi kesalahan saat memproses file:"
        )

        st.code(str(e))
