import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="AUM PTSdL",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Aplikasi Analisis AUM PTSdL")
st.caption("Pengolahan otomatis data AUM PTSdL dari Excel hasil Google Form")

# =========================
# KELOMPOK BUTIR AUM
# =========================

GROUPS = {
    "P – Prasyarat Penguasaan Materi": [
        1,2,3,4,5,31,32,33,34,35,
        61,62,63,64,65,91,92,93,94,95
    ],

    "T – Teknik/Keterampilan Belajar": (
        list(range(6,16)) +
        list(range(36,46)) +
        list(range(66,76)) +
        list(range(96,111)) +
        list(range(121,136)) +
        list(range(146,161))
    ),

    "S – Sarana Belajar": [
        16,17,18,19,20,
        46,47,48,49,50,
        76,77,78,79,80
    ],

    "D – Keadaan Diri Pribadi": [
        21,22,23,24,25,
        51,52,53,54,55,
        81,82,83,84,85,
        111,112,113,114,115,
        136,137,138,139,140,
        161,162,163,164,165
    ],

    "L – Lingkungan Sosial-Emosional": [
        26,27,28,29,30,
        56,57,58,59,60,
        86,87,88,89,90,
        116,117,118,119,120,
        141,142,143,144,145
    ]
}

# =========================
# KONVERSI JAWABAN
# =========================

SCORE_MAP = {
    "1. Jarang": 1,
    "2. Kadang-kadang": 2,
    "3. Sering": 3,
    "4. Pada umumnya": 4,
    "5. Selalu": 5,
    "5. selalu": 5
}

# =========================
# UPLOAD FILE
# =========================

uploaded_file = st.file_uploader(
    "Upload file Excel jawaban AUM PTSdL",
    type=["xlsx", "xls"]
)

if uploaded_file:

    try:
        df = pd.read_excel(uploaded_file)

    except Exception as e:
        st.error(f"File Excel tidak dapat dibaca: {e}")
        st.stop()

    if df.shape[1] < 178:
        st.error(
            f"Kolom pada file hanya {df.shape[1]}. "
            "Format yang diharapkan memiliki 13 kolom identitas + 165 butir AUM."
        )
        st.stop()

    st.success(f"File berhasil dibaca: {len(df)} responden.")

    # 165 butir AUM berada setelah 13 kolom identitas
    question_cols = list(df.columns[13:178])

    if len(question_cols) != 165:
        st.error("165 kolom butir AUM tidak terdeteksi dengan benar.")
        st.stop()

    # =========================
    # UBAH JAWABAN MENJADI ANGKA
    # =========================

    numeric = df[question_cols].applymap(
        lambda x: SCORE_MAP.get(
            str(x).strip(),
            pd.to_numeric(x, errors="coerce")
        )
    )

    # =========================
    # DATA JAWABAN
    # =========================

    st.subheader("1. Data Jawaban")

    st.dataframe(
        df.iloc[:, :13].join(numeric),
        use_container_width=True,
        height=350
    )

    # =========================
    # HASIL PER SISWA
    # =========================

    st.subheader("2. Hasil Per Bidang AUM")

    result = df.iloc[:, :13].copy()

    for group, numbers in GROUPS.items():

        cols = [
            question_cols[n - 1]
            for n in numbers
        ]

        group_scores = numeric[cols]

        result[f"{group} - Total"] = (
            group_scores.sum(axis=1)
        )

        result[f"{group} - Rata-rata"] = (
            group_scores.mean(axis=1).round(2)
        )

        result[f"{group} - Skor ≥ 3"] = (
            (group_scores >= 3).sum(axis=1)
        )

    st.dataframe(
        result,
        use_container_width=True,
        height=450
    )

    # =========================
    # RINGKASAN BIDANG
    # =========================

    st.subheader("3. Ringkasan Bidang")

    bidang_rows = []

    for group, numbers in GROUPS.items():

        cols = [
            question_cols[n - 1]
            for n in numbers
        ]

        scores = numeric[cols]

        bidang_rows.append({
            "Bidang": group,
            "Jumlah Butir": len(numbers),
            "Rata-rata Semua Responden":
                round(scores.mean().mean(), 2),
            "Rata-rata Jumlah Skor":
                round(scores.sum(axis=1).mean(), 2),
            "Rata-rata Butir Skor ≥ 3":
                round((scores >= 3).sum(axis=1).mean(), 2)
        })

    bidang_df = pd.DataFrame(bidang_rows)

    st.dataframe(
        bidang_df,
        use_container_width=True
    )

    # =========================
    # HASIL SISWA
    # =========================

    st.subheader("4. Hasil Tiap Siswa")

    nama_col = df.columns[3]

    pilihan = st.selectbox(
        "Pilih siswa",
        df[nama_col].astype(str).tolist()
    )

    idx = df[nama_col].astype(str).tolist().index(
        pilihan
    )

    siswa = result.iloc[idx]

    detail = []

    for group, numbers in GROUPS.items():

        detail.append({
            "Bidang": group,
            "Jumlah Butir": len(numbers),
            "Total Skor":
                siswa[f"{group} - Total"],
            "Rata-rata":
                siswa[f"{group} - Rata-rata"],
            "Butir dengan Skor ≥ 3":
                siswa[f"{group} - Skor ≥ 3"]
        })

    detail_df = pd.DataFrame(detail)

    st.dataframe(
        detail_df,
        use_container_width=True
    )

    st.info(
        "Catatan: hasil ini merupakan pengolahan skor "
        "frekuensi jawaban AUM. Interpretasi sebagai "
        "masalah/kebutuhan siswa tetap perlu mengikuti "
        "pedoman AUM PTSdL dan penilaian konselor."
    )

    # =========================
    # DOWNLOAD EXCEL
    # =========================

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        result.to_excel(
            writer,
            index=False,
            sheet_name="Hasil Per Siswa"
        )

        bidang_df.to_excel(
            writer,
            index=False,
            sheet_name="Ringkasan Bidang"
        )

        detail_df.to_excel(
            writer,
            index=False,
            sheet_name="Siswa Terpilih"
        )

    st.download_button(
        "⬇️ Download Hasil Excel",
        data=output.getvalue(),
        file_name="Hasil_Pengolahan_AUM_PTSdL.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:

    st.info(
        "Silakan upload file SMA (Jawaban) dari Google Form "
        "untuk mulai mengolah data."
    )
