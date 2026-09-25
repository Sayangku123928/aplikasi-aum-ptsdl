import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="AUM PTSdL",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Aplikasi Analisis AUM PTSdL")
st.write("Pengolahan data AUM PTSdL dari hasil Google Form.")

# =========================================================
# PEMBAGIAN 165 BUTIR AUM PTSdL
# =========================================================

GROUPS = {
    "P – Prasyarat Penguasaan Materi": [
        1, 2, 3, 4, 5,
        31, 32, 33, 34, 35,
        61, 62, 63, 64, 65,
        91, 92, 93, 94, 95
    ],

    "T – Teknik / Keterampilan Belajar": [
        6, 7, 8, 9, 10, 11, 12, 13, 14, 15,
        36, 37, 38, 39, 40, 41, 42, 43, 44, 45,
        66, 67, 68, 69, 70, 71, 72, 73, 74, 75,
        96, 97, 98, 99, 100, 101, 102, 103, 104, 105,
        106, 107, 108, 109, 110,
        121, 122, 123, 124, 125, 126, 127, 128, 129, 130,
        131, 132, 133, 134, 135,
        146, 147, 148, 149, 150, 151, 152, 153, 154, 155,
        156, 157, 158, 159, 160
    ],

    "S – Sarana Belajar": [
        16, 17, 18, 19, 20,
        46, 47, 48, 49, 50,
        76, 77, 78, 79, 80
    ],

    "D – Keadaan Diri Pribadi": [
        21, 22, 23, 24, 25,
        51, 52, 53, 54, 55,
        81, 82, 83, 84, 85,
        111, 112, 113, 114, 115,
        136, 137, 138, 139, 140,
        161, 162, 163, 164, 165
    ],

    "L – Lingkungan Sosial-Emosional": [
        26, 27, 28, 29, 30,
        56, 57, 58, 59, 60,
        86, 87, 88, 89, 90,
        116, 117, 118, 119, 120,
        141, 142, 143, 144, 145
    ]
}


# =========================================================
# KONVERSI JAWABAN
# =========================================================

def ubah_jawaban(nilai):

    if pd.isna(nilai):
        return None

    teks = str(nilai).strip().lower()

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

    # Kalau ternyata sudah berupa angka
    try:
        angka = float(teks)
        if angka in [1, 2, 3, 4, 5]:
            return int(angka)
    except:
        pass

    return None


# =========================================================
# UPLOAD FILE
# =========================================================

st.subheader("📁 Upload Data AUM")

file = st.file_uploader(
    "Pilih file Excel hasil Google Form",
    type=["xlsx"]
)

if file is None:

    st.info(
        "Silakan upload file SMA (Jawaban) dari Google Form."
    )

else:

    # =====================================================
    # MEMBACA EXCEL
    # =====================================================

    try:

        df = pd.read_excel(
            file,
            engine="openpyxl"
        )

    except Exception as e:

        st.error("File Excel belum dapat dibaca.")

        st.code(str(e))

        st.warning(
            "Pastikan file yang diupload adalah file .xlsx "
            "hasil Google Form."
        )

        st.stop()


    st.success(
        f"File berhasil dibaca! "
        f"Ditemukan {len(df)} responden dan {len(df.columns)} kolom."
    )


    # =====================================================
    # CEK FORMAT
    # =====================================================

    if len(df.columns) < 178:

        st.error(
            f"Jumlah kolom terdeteksi {len(df.columns)}. "
            "File AUM yang digunakan seharusnya memiliki "
            "13 kolom identitas + 165 butir AUM = 178 kolom."
        )

        st.stop()


    # =====================================================
    # IDENTITAS
    # =====================================================

    identitas = df.iloc[:, :13].copy()

    # 165 pertanyaan dimulai dari kolom ke-14
    pertanyaan = df.iloc[:, 13:178].copy()


    # =====================================================
    # KONVERSI JAWABAN KE ANGKA
    # =====================================================

    skor = pertanyaan.map(ubah_jawaban)


    # =====================================================
    # CEK JUMLAH BUTIR
    # =====================================================

    if skor.shape[1] != 165:

        st.error(
            f"Jumlah butir yang terbaca adalah "
            f"{skor.shape[1]}, bukan 165."
        )

        st.stop()


    st.success(
        "165 butir AUM berhasil ditemukan dan diproses."
    )


    # =====================================================
    # PREVIEW DATA
    # =====================================================

    st.subheader("👀 Preview Data")

    preview = pd.concat(
        [identitas, skor],
        axis=1
    )

    st.dataframe(
        preview.head(10),
        use_container_width=True
    )


    # =====================================================
    # HASIL PER KELOMPOK
    # =====================================================

    hasil = identitas.copy()

    for nama_kelompok, nomor_butir in GROUPS.items():

        kolom = [
            pertanyaan.iloc[:, nomor - 1].name
            for nomor in nomor_butir
        ]

        data_kelompok = skor[kolom]

        hasil[nama_kelompok + " - Total Skor"] = (
            data_kelompok.sum(axis=1)
        )

        hasil[nama_kelompok + " - Rata-rata"] = (
            data_kelompok.mean(axis=1).round(2)
        )

        hasil[nama_kelompok + " - Butir ≥ 3"] = (
            (data_kelompok >= 3).sum(axis=1)
        )


    # =====================================================
    # HASIL
    # =====================================================

    st.subheader("📊 Hasil Pengolahan AUM")

    st.dataframe(
        hasil,
        use_container_width=True,
        height=450
    )


    # =====================================================
    # RINGKASAN P / T / S / D / L
    # =====================================================

    st.subheader("📈 Ringkasan Bidang")

    ringkasan = []

    for nama_kelompok, nomor_butir in GROUPS.items():

        kolom = [
            pertanyaan.iloc[:, nomor - 1].name
            for nomor in nomor_butir
        ]

        data_kelompok = skor[kolom]

        ringkasan.append({
            "Bidang": nama_kelompok,
            "Jumlah Butir": len(nomor_butir),
            "Rata-rata Skor": round(
                data_kelompok.mean().mean(), 2
            ),
            "Rata-rata Butir ≥ 3": round(
                (data_kelompok >= 3).sum(axis=1).mean(), 2
            )
        })

    ringkasan_df = pd.DataFrame(ringkasan)

    st.dataframe(
        ringkasan_df,
        use_container_width=True
    )


    # =====================================================
    # PILIH SISWA
    # =====================================================

    st.subheader("👤 Hasil Per Siswa")

    nama_kolom = "NAMA"

    if nama_kolom in df.columns:

        daftar_nama = (
            df[nama_kolom]
            .fillna("Tanpa Nama")
            .astype(str)
            .tolist()
        )

        nama_pilihan = st.selectbox(
            "Pilih nama siswa:",
            daftar_nama
        )

        posisi = daftar_nama.index(
            nama_pilihan
        )

        detail = []

        for nama_kelompok, nomor_butir in GROUPS.items():

            kolom = [
                pertanyaan.iloc[:, nomor - 1].name
                for nomor in nomor_butir
            ]

            data_siswa = skor.iloc[posisi][kolom]

            detail.append({
                "Bidang": nama_kelompok,
                "Jumlah Butir": len(nomor_butir),
                "Total Skor": int(
                    data_siswa.sum()
                ),
                "Rata-rata": round(
                    data_siswa.mean(), 2
                ),
                "Butir Skor ≥ 3": int(
                    (data_siswa >= 3).sum()
                )
            })

        detail_df = pd.DataFrame(detail)

        st.dataframe(
            detail_df,
            use_container_width=True
        )


    # =====================================================
    # DOWNLOAD EXCEL
    # =====================================================

    st.subheader("⬇️ Download Hasil")

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        hasil.to_excel(
            writer,
            index=False,
            sheet_name="Hasil Per Siswa"
        )

        ringkasan_df.to_excel(
            writer,
            index=False,
            sheet_name="Ringkasan PTSdL"
        )

        preview.to_excel(
            writer,
            index=False,
            sheet_name="Data Skor"
        )

    st.download_button(
        label="📥 Download Hasil Pengolahan AUM",
        data=output.getvalue(),
        file_name="Hasil_Pengolahan_AUM_PTSdL.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
