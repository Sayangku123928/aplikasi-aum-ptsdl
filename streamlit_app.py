import streamlit as st
import pandas as pd
import re
from io import BytesIO

# =========================================================
# CONFIG
# =========================================================
st.set_page_config(
    page_title="AUM PTSdL",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =========================================================
# STYLE
# =========================================================
st.markdown("""
<style>
    .stApp {
        background: #fff8fc;
    }

    .main-title {
        background: linear-gradient(135deg, #f7a8d8, #c9a7ff);
        padding: 28px 32px;
        border-radius: 22px;
        color: white;
        margin-bottom: 24px;
        box-shadow: 0 8px 24px rgba(190, 120, 180, .18);
    }

    .main-title h1 {
        margin: 0;
        font-size: 34px;
    }

    .main-title p {
        margin: 8px 0 0;
        font-size: 16px;
        opacity: .95;
    }

    .metric-card {
        background: white;
        border-radius: 18px;
        padding: 18px;
        border: 1px solid #f1ddec;
        box-shadow: 0 5px 18px rgba(170, 100, 160, .08);
        text-align: center;
    }

    .metric-label {
        color: #8d6a87;
        font-size: 14px;
        font-weight: 600;
    }

    .metric-value {
        color: #4f3b4d;
        font-size: 28px;
        font-weight: 800;
        margin-top: 4px;
    }

    .section-title {
        color: #5d3d58;
        font-size: 22px;
        font-weight: 800;
        margin: 18px 0 12px;
    }

    .info-box {
        background: #fff0f8;
        border-left: 5px solid #e99acb;
        padding: 14px 18px;
        border-radius: 12px;
        color: #67495f;
        margin: 12px 0 20px;
    }

    div[data-testid="stDataFrame"] {
        border-radius: 14px;
        overflow: hidden;
    }

    .field-card {
        background: white;
        border-radius: 18px;
        padding: 18px;
        border: 1px solid #ead9ee;
        text-align: center;
        box-shadow: 0 4px 16px rgba(150, 100, 170, .07);
    }

    .field-name {
        font-size: 20px;
        font-weight: 800;
        color: #6d4b69;
    }

    .field-score {
        font-size: 26px;
        font-weight: 800;
        color: #b35d9f;
        margin: 5px 0;
    }

    .field-items {
        font-size: 12px;
        color: #907b8d;
    }
</style>
""", unsafe_allow_html=True)

# =========================================================
# HEADER
# =========================================================
st.markdown("""
<div class="main-title">
    <h1>📋 Aplikasi Pengolahan AUM PTSdL</h1>
    <p>Pengolahan data AUM dari hasil Google Form secara otomatis</p>
</div>
""", unsafe_allow_html=True)

# =========================================================
# KUNCI BIDANG 1–165
# =========================================================
kunci_bidang = {}

# P
for n in (
    list(range(1, 6))
    + list(range(31, 36))
    + list(range(61, 66))
    + list(range(91, 96))
):
    kunci_bidang[n] = "P"

# T
for n in (
    list(range(6, 16))
    + list(range(36, 46))
    + list(range(66, 76))
    + list(range(96, 111))
    + list(range(121, 136))
):
    kunci_bidang[n] = "T"

# S
for n in (
    list(range(16, 21))
    + list(range(46, 51))
    + list(range(76, 81))
    + list(range(146, 161))
):
    kunci_bidang[n] = "S"

# D
for n in (
    list(range(21, 26))
    + list(range(51, 56))
    + list(range(81, 86))
    + list(range(111, 116))
    + list(range(136, 141))
    + list(range(161, 166))
):
    kunci_bidang[n] = "D"

# L
for n in (
    list(range(26, 31))
    + list(range(56, 61))
    + list(range(86, 91))
    + list(range(116, 121))
    + list(range(141, 146))
):
    kunci_bidang[n] = "L"

# Validasi
if len(kunci_bidang) != 165:
    st.error(
        f"Kunci bidang tidak lengkap: "
        f"{len(kunci_bidang)} dari 165 butir."
    )
    st.stop()

jumlah_butir_bidang = {
    bidang: list(kunci_bidang.values()).count(bidang)
    for bidang in ["P", "T", "S", "D", "L"]
}

# =========================================================
# FUNGSI KONVERSI JAWABAN KE SKOR
# =========================================================
def ubah_ke_skor(nilai):

    if pd.isna(nilai):
        return None

    teks = str(nilai).strip()

    # Jika sudah berupa angka 1–5
    try:
        angka = float(teks)

        if angka in [1, 2, 3, 4, 5]:
            return int(angka)

    except:
        pass

    # Jika jawaban diawali angka
    cocok = re.match(r"^\s*([1-5])", teks)

    if cocok:
        return int(cocok.group(1))

    # Jika berupa teks
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


# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:

    st.markdown("## 📁 Input Data")

    file = st.file_uploader(
        "Upload Excel hasil Google Form",
        type=["xlsx"]
    )

    st.markdown("---")

    st.markdown("### 📌 Skala Jawaban")

    st.write("1 = Jarang")
    st.write("2 = Kadang-kadang")
    st.write("3 = Sering")
    st.write("4 = Pada umumnya")
    st.write("5 = Selalu")

    st.markdown("---")

    st.markdown("### 🗂️ Pembagian Butir")

    for bidang in ["P", "T", "S", "D", "L"]:

        st.write(
            f"**{bidang}** — "
            f"{jumlah_butir_bidang[bidang]} butir"
        )


# =========================================================
# JIKA BELUM UPLOAD
# =========================================================
if file is None:

    st.markdown("""
    <div class="info-box">
        👋 <b>Selamat datang!</b><br>
        Silakan upload file Excel hasil Google Form
        melalui menu di sebelah kiri untuk mulai
        mengolah data AUM.
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)

    with c1:

        st.markdown("""
        <div class="metric-card">
            <div class="metric-label">Jumlah Butir</div>
            <div class="metric-value">165</div>
        </div>
        """, unsafe_allow_html=True)

    with c2:

        st.markdown("""
        <div class="metric-card">
            <div class="metric-label">Bidang</div>
            <div class="metric-value">5</div>
        </div>
        """, unsafe_allow_html=True)

    with c3:

        st.markdown("""
        <div class="metric-card">
            <div class="metric-label">Skala</div>
            <div class="metric-value">1–5</div>
        </div>
        """, unsafe_allow_html=True)

    st.stop()


# =========================================================
# BACA FILE EXCEL
# =========================================================
try:

    df = pd.read_excel(
        file,
        engine="openpyxl"
    )

except Exception as e:

    st.error("❌ File Excel tidak dapat dibaca.")

    st.code(str(e))

    st.stop()


# =========================================================
# DETEKSI KOLOM AUM
# =========================================================
kolom_aum = []

for kolom in df.columns:

    jumlah_terbaca = (
        df[kolom]
        .apply(ubah_ke_skor)
        .notna()
        .sum()
    )

    if jumlah_terbaca >= 1:

        kolom_aum.append(kolom)


if len(kolom_aum) < 165:

    st.error(
        f"❌ Baru ditemukan {len(kolom_aum)} "
        "kolom yang terbaca sebagai AUM. "
        "Dibutuhkan minimal 165 butir."
    )

    st.write(kolom_aum)

    st.stop()


# Ambil 165 kolom AUM terakhir
kolom_aum = kolom_aum[-165:]


# =========================================================
# KONVERSI KE SKOR
# =========================================================
data_skor = df[kolom_aum].copy()

for kolom in data_skor.columns:

    data_skor[kolom] = (
        data_skor[kolom]
        .apply(ubah_ke_skor)
    )


# =========================================================
# DETEKSI NAMA RESPONDEN
# =========================================================
kolom_nama = None

for kolom in df.columns:

    nama = str(kolom).lower()

    if "nama" in nama or "name" in nama:

        kolom_nama = kolom

        break


if kolom_nama is not None:

    nama_responden = (
        df[kolom_nama]
        .fillna("Tanpa Nama")
        .astype(str)
    )

else:

    nama_responden = pd.Series(
        [
            f"Responden {i + 1}"
            for i in range(len(df))
        ]
    )


# =========================================================
# TOTAL DAN RATA-RATA
# =========================================================
total = data_skor.sum(
    axis=1,
    skipna=True
)

rata = data_skor.mean(
    axis=1,
    skipna=True
)


# =========================================================
# SKOR PER BIDANG
# =========================================================
skor_bidang = {}

for bidang in ["P", "T", "S", "D", "L"]:

    nomor_soal = [
        no
        for no, b in kunci_bidang.items()
        if b == bidang
    ]

    kolom_bidang = [
        data_skor.columns[no - 1]
        for no in nomor_soal
    ]

    skor_bidang[bidang] = (
        data_skor[kolom_bidang]
        .sum(
            axis=1,
            skipna=True
        )
    )


# =========================================================
# MEMBUAT HASIL
# =========================================================
hasil = pd.DataFrame()

hasil["Nama"] = nama_responden

hasil["Jumlah Butir Terisi"] = (
    data_skor
    .notna()
    .sum(axis=1)
)

# Skor P/T/S/D/L
for bidang in ["P", "T", "S", "D", "L"]:

    hasil[bidang] = skor_bidang[bidang]


# Rata-rata P/T/S/D/L
for bidang in ["P", "T", "S", "D", "L"]:

    hasil[f"Rata-rata {bidang}"] = (
        hasil[bidang]
        / jumlah_butir_bidang[bidang]
    ).round(2)


hasil["Total Skor"] = total

hasil["Rata-rata"] = rata.round(2)


# =========================================================
# DASHBOARD
# =========================================================
st.markdown(
    '<div class="section-title">'
    '📊 Dashboard'
    '</div>',
    unsafe_allow_html=True
)

c1, c2, c3, c4 = st.columns(4)


with c1:

    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">
                👥 Responden
            </div>
            <div class="metric-value">
                {len(df)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c2:

    st.markdown(
        """
        <div class="metric-card">
            <div class="metric-label">
                📝 Butir
            </div>
            <div class="metric-value">
                165
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c3:

    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">
                📋 Rata-rata Butir Terisi
            </div>
            <div class="metric-value">
                {hasil["Jumlah Butir Terisi"].mean():.0f}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c4:

    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">
                ⭐ Rata-rata Umum
            </div>
            <div class="metric-value">
                {hasil["Rata-rata"].mean():.2f}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# RINGKASAN BIDANG
# =========================================================
st.markdown(
    '<div class="section-title">'
    '🗂️ Ringkasan Bidang'
    '</div>',
    unsafe_allow_html=True
)

cards = st.columns(5)


for i, bidang in enumerate(
    ["P", "T", "S", "D", "L"]
):

    nilai = (
        hasil[f"Rata-rata {bidang}"]
        .mean()
    )

    with cards[i]:

        st.markdown(
            f"""
            <div class="field-card">

                <div class="field-name">
                    {bidang}
                </div>

                <div class="field-score">
                    {nilai:.2f}
                </div>

                <div class="field-items">
                    {jumlah_butir_bidang[bidang]} butir
                </div>

            </div>
            """,
            unsafe_allow_html=True
        )


# =========================================================
# GRAFIK
# =========================================================
st.markdown(
    '<div class="section-title">'
    '📈 Grafik Rata-rata Bidang'
    '</div>',
    unsafe_allow_html=True
)

rata_bidang = pd.DataFrame({

    "Bidang": [
        "P",
        "T",
        "S",
        "D",
        "L"
    ],

    "Rata-rata": [

        hasil["Rata-rata P"].mean(),

        hasil["Rata-rata T"].mean(),

        hasil["Rata-rata S"].mean(),

        hasil["Rata-rata D"].mean(),

        hasil["Rata-rata L"].mean(),

    ]

}).set_index("Bidang")


st.bar_chart(rata_bidang)


# =========================================================
# PILIH RESPONDEN
# =========================================================
st.markdown(
    '<div class="section-title">'
    '👤 Detail Responden'
    '</div>',
    unsafe_allow_html=True
)

pilihan_nama = hasil["Nama"].tolist()


nama_pilihan = st.selectbox(
    "Pilih responden:",
    pilihan_nama
)


detail = (
    hasil[
        hasil["Nama"] == nama_pilihan
    ]
    .iloc[0]
)


d1, d2, d3, d4, d5 = st.columns(5)


for col, bidang in zip(
    [d1, d2, d3, d4, d5],
    ["P", "T", "S", "D", "L"]
):

    with col:

        st.metric(
            bidang,
            f"{detail[f'Rata-rata {bidang}']:.2f}",
            f"{int(detail[bidang])} skor"
        )


st.info(
    f"📌 {nama_pilihan} mengisi "
    f"{int(detail['Jumlah Butir Terisi'])} "
    "dari 165 butir."
)


# =========================================================
# TABEL HASIL
# =========================================================
st.markdown(
    '<div class="section-title">'
    '📋 Tabel Hasil Semua Responden'
    '</div>',
    unsafe_allow_html=True
)

kolom_tampil = [

    "Nama",

    "Jumlah Butir Terisi",

    "P",
    "T",
    "S",
    "D",
    "L",

    "Rata-rata P",
    "Rata-rata T",
    "Rata-rata S",
    "Rata-rata D",
    "Rata-rata L",

    "Total Skor",

    "Rata-rata"

]


st.dataframe(
    hasil[kolom_tampil],
    use_container_width=True,
    hide_index=True
)


# =========================================================
# DATA SKOR 165 BUTIR
# =========================================================
with st.expander(
    "🔢 Lihat Data Skor 165 Butir"
):

    st.dataframe(
        data_skor,
        use_container_width=True,
        hide_index=True
    )


# =========================================================
# KUNCI BIDANG
# =========================================================
with st.expander(
    "🗝️ Lihat Kunci Bidang 1–165"
):

    tabel_kunci = pd.DataFrame({

        "No. Butir":
            list(range(1, 166)),

        "Bidang": [
            kunci_bidang[i]
            for i in range(1, 166)
        ]

    })


    st.dataframe(
        tabel_kunci,
        use_container_width=True,
        hide_index=True,
        height=450
    )


# =========================================================
# DOWNLOAD EXCEL
# =========================================================
st.markdown(
    '<div class="section-title">'
    '📥 Download Hasil'
    '</div>',
    unsafe_allow_html=True
)

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

    rata_bidang.reset_index().to_excel(
        writer,
        sheet_name="Ringkasan Bidang",
        index=False
    )

    tabel_kunci.to_excel(
        writer,
        sheet_name="Kunci Bidang",
        index=False
    )


st.download_button(

    label="⬇️ Download Hasil Excel",

    data=output.getvalue(),

    file_name="Hasil_AUM_PTSdL.xlsx",

    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),

    use_container_width=True
)


st.markdown("---")

st.caption(
    "Aplikasi Pengolahan AUM PTSdL • "
    "165 Butir • P / T / S / D / L"
)
