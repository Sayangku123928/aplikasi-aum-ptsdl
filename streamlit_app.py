import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

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

# =========================================================
# VALIDASI KUNCI
# =========================================================
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
# FUNGSI NAMA SHEET EXCEL
# =========================================================
def buat_nama_sheet(nama, nama_terpakai):

    nama = str(nama).strip()

    if not nama or nama.lower() == "nan":
        nama = "Tanpa Nama"

    # Karakter yang tidak boleh digunakan Excel
    nama = re.sub(r'[\[\]\:\*\?\/\\]', '', nama)

    # Maksimal nama sheet Excel = 31 karakter
    nama = nama[:31].strip()

    if not nama:
        nama = "Siswa"

    nama_awal = nama
    nomor = 2

    while nama in nama_terpakai:

        tambahan = f" ({nomor})"

        nama = (
            nama_awal[:31 - len(tambahan)]
            + tambahan
        )

        nomor += 1

    nama_terpakai.add(nama)

    return nama


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
tabel_kunci = pd.DataFrame({

    "No. Butir":
        list(range(1, 166)),

    "Bidang": [
        kunci_bidang[i]
        for i in range(1, 166)
    ]

})

with st.expander(
    "🗝️ Lihat Kunci Bidang 1–165"
):

    st.dataframe(
        tabel_kunci,
        use_container_width=True,
        hide_index=True,
        height=450
    )


# =========================================================
# DOWNLOAD EXCEL
# 1 SISWA = 1 SHEET
# =========================================================
st.markdown(
    '<div class="section-title">'
    '📥 Download Hasil Excel'
    '</div>',
    unsafe_allow_html=True
)

st.info(
    "📌 Setiap siswa akan dibuatkan 1 sheet sendiri. "
    "Nama siswa digunakan sebagai nama sheet."
)

output = BytesIO()

with pd.ExcelWriter(
    output,
    engine="openpyxl"
) as writer:

    workbook = writer.book

    nama_terpakai = set()

    # =====================================================
    # BUAT SHEET UNTUK SETIAP SISWA
    # =====================================================
    for index in range(len(df)):

        nama_siswa = (
            str(nama_responden.iloc[index])
            .strip()
        )

        sheet_name = buat_nama_sheet(
            nama_siswa,
            nama_terpakai
        )

        worksheet = workbook.create_sheet(
            title=sheet_name
        )

        # =================================================
        # JUDUL
        # =================================================
        worksheet["A1"] = "HASIL AUM PTSdL"
        worksheet["A1"].font = Font(
            bold=True,
            size=16
        )

        worksheet.merge_cells(
            "A1:D1"
        )

        worksheet["A1"].alignment = Alignment(
            horizontal="center"
        )

        # =================================================
        # IDENTITAS SISWA
        # =================================================
        worksheet["A3"] = "Nama Siswa"
        worksheet["B3"] = nama_siswa

        worksheet["A4"] = "Jumlah Butir Terisi"
        worksheet["B4"] = int(
            data_skor.iloc[index].notna().sum()
        )

        worksheet["A5"] = "Total Skor"
        worksheet["B5"] = float(
            total.iloc[index]
        )

        worksheet["A6"] = "Rata-rata"
        worksheet["B6"] = round(
            float(rata.iloc[index]),
            2
        )

        # =================================================
        # REKAP BIDANG
        # =================================================
        worksheet["A8"] = "REKAP BIDANG"

        worksheet["A9"] = "Bidang"
        worksheet["B9"] = "Skor"
        worksheet["C9"] = "Rata-rata"
        worksheet["D9"] = "Jumlah Butir"

        for kolom in ["A9", "B9", "C9", "D9"]:

            worksheet[kolom].font = Font(
                bold=True
            )

        baris = 10

        for bidang in ["P", "T", "S", "D", "L"]:

            skor = float(
                hasil.iloc[index][bidang]
            )

            rata_bidang_siswa = round(
                skor / jumlah_butir_bidang[bidang],
                2
            )

            worksheet.cell(
                row=baris,
                column=1,
                value=bidang
            )

            worksheet.cell(
                row=baris,
                column=2,
                value=skor
            )

            worksheet.cell(
                row=baris,
                column=3,
                value=rata_bidang_siswa
            )

            worksheet.cell(
                row=baris,
                column=4,
                value=jumlah_butir_bidang[bidang]
            )

            baris += 1

        # =================================================
        # TABEL 165 BUTIR
        # =================================================
        baris_awal = 17

        worksheet.cell(
            row=baris_awal,
            column=1,
            value="DATA 165 BUTIR"
        )

        worksheet.cell(
            row=baris_awal + 1,
            column=1,
            value="No. Butir"
        )

        worksheet.cell(
            row=baris_awal + 1,
            column=2,
            value="Jawaban"
        )

        worksheet.cell(
            row=baris_awal + 1,
            column=3,
            value="Skor"
        )

        worksheet.cell(
            row=baris_awal + 1,
            column=4,
            value="Bidang"
        )

        for kolom in range(1, 5):

            worksheet.cell(
                row=baris_awal + 1,
                column=kolom
            ).font = Font(
                bold=True
            )

        # =================================================
        # ISI 165 BUTIR
        # =================================================
        for nomor in range(1, 166):

            baris_data = (
                baris_awal + 1 + nomor
            )

            kolom_asli = kolom_aum[
                nomor - 1
            ]

            jawaban_asli = df.iloc[
                index
            ][kolom_asli]

            skor = data_skor.iloc[
                index,
                nomor - 1
            ]

            bidang = kunci_bidang[
                nomor
            ]

            worksheet.cell(
                row=baris_data,
                column=1,
                value=nomor
            )

            if pd.isna(jawaban_asli):

                worksheet.cell(
                    row=baris_data,
                    column=2,
                    value=""
                )

            else:

                worksheet.cell(
                    row=baris_data,
                    column=2,
                    value=str(jawaban_asli)
                )

            if pd.isna(skor):

                worksheet.cell(
                    row=baris_data,
                    column=3,
                    value=""
                )

            else:

                worksheet.cell(
                    row=baris_data,
                    column=3,
                    value=int(skor)
                )

            worksheet.cell(
                row=baris_data,
                column=4,
                value=bidang
            )

        # =================================================
        # FORMAT SHEET
        # =================================================

        # Header utama
        for cell in worksheet[1]:

            cell.font = Font(
                bold=True,
                size=16
            )

        # Header tabel
        for row in [
            9,
            baris_awal + 1
        ]:

            for cell in worksheet[row]:

                cell.font = Font(
                    bold=True
                )

        # Alignment
        for row in worksheet.iter_rows():

            for cell in row:

                cell.alignment = Alignment(
                    vertical="center"
                )

        # Lebar kolom
        worksheet.column_dimensions["A"].width = 22
        worksheet.column_dimensions["B"].width = 35
        worksheet.column_dimensions["C"].width = 15
        worksheet.column_dimensions["D"].width = 15

        # Freeze tabel
        worksheet.freeze_panes = (
            f"A{baris_awal + 2}"
        )

        # Border
        thin = Side(
            style="thin",
            color="D9C7D8"
        )

        border = Border(
            left=thin,
            right=thin,
            top=thin,
            bottom=thin
        )

        # Border rekap
        for row in worksheet.iter_rows(
            min_row=9,
            max_row=14,
            min_col=1,
            max_col=4
        ):

            for cell in row:

                cell.border = border

        # Border data 165
        for row in worksheet.iter_rows(
            min_row=baris_awal + 1,
            max_row=baris_awal + 166,
            min_col=1,
            max_col=4
        ):

            for cell in row:

                cell.border = border


# =========================================================
# DOWNLOAD BUTTON
# =========================================================
st.download_button(

    label="⬇️ Download Excel — 1 Siswa = 1 Sheet",

    data=output.getvalue(),

    file_name="Hasil_AUM_PTSdL_Per_Siswa.xlsx",

    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),

    use_container_width=True
)


# =========================================================
# FOOTER
# =========================================================
st.markdown("---")

st.caption(
    "Aplikasi Pengolahan AUM PTSdL • "
    "165 Butir • P / T / S / D / L"
)
