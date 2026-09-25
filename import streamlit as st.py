import os
import json
import re
import urllib.request
from difflib import get_close_matches
import streamlit as st
import pandas as pd
import plotly.express as px


# =============================================================================
# 1. KONFIGURASI HALAMAN & TEMA EDITORIAL MODERN (LIGHT MODE)
# =============================================================================

st.set_page_config(
    page_title="Youth Vulnerability Index Dashboard",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Poppins', sans-serif;
}

.stApp {
    background:
        radial-gradient(circle at 8% 15%, rgba(99,102,241,0.16) 0%, transparent 40%),
        radial-gradient(circle at 92% 10%, rgba(244,63,94,0.10) 0%, transparent 38%),
        radial-gradient(circle at 15% 90%, rgba(245,158,11,0.10) 0%, transparent 35%),
        radial-gradient(circle at 90% 85%, rgba(139,92,246,0.16) 0%, transparent 40%),
        linear-gradient(160deg, #faf8ff 0%, #f3f0ff 100%);
    background-attachment: fixed;
    color: #1e293b;
}

[data-testid="stSidebar"] {
    background-color: #ffffff;
    border-right: 1px solid rgba(15,23,42,0.08);
}

/* Rapikan tag multiselect di sidebar agar tidak terpotong/tumpang tindih tombol X */
[data-baseweb="tag"] {
    max-width: 100% !important;
    height: auto !important;
    white-space: normal !important;
    padding: 5px 8px !important;
    margin: 3px 4px 3px 0 !important;
    background: linear-gradient(90deg, #6366f1, #8b5cf6) !important;
    border-radius: 8px !important;
}

[data-baseweb="tag"] span {
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: unset !important;
    word-break: break-word;
    line-height: 1.3 !important;
    padding-right: 4px !important;
}

[data-baseweb="tag"] svg {
    fill: #ffffff !important;
}

[data-baseweb="select"] > div {
    border-radius: 10px !important;
}

[data-baseweb="select"]:focus-within > div {
    border-color: #8b5cf6 !important;
    box-shadow: 0 0 0 1px #8b5cf6 !important;
}

h1 {
    background: linear-gradient(90deg, #4f46e5 0%, #a855f7 55%, #ec4899 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    font-weight: 800 !important;
    letter-spacing: -0.5px;
}

h2, h3, h4, h5 {
    color: #0f172a !important;
    font-weight: 600 !important;
}

p, span, label, .stCaption, [data-testid="stCaptionContainer"] {
    color: #475569;
}

[data-testid="stMetric"] {
    background: #ffffff;
    border: 1px solid rgba(15,23,42,0.08);
    border-left: 4px solid #8b5cf6;
    border-radius: 14px;
    padding: 16px 18px;
    box-shadow: 0 2px 10px rgba(15,23,42,0.06);
    transition: transform 0.15s ease, box-shadow 0.15s ease;
    min-height: 108px;
}

[data-testid="stMetric"]:hover {
    transform: translateY(-3px);
    box-shadow: 0 8px 20px rgba(15,23,42,0.10);
}

/* Warna berbeda tiap kartu KPI baris atas, sesuai makna datanya */
.st-key-kpi_row div[data-testid="column"]:nth-of-type(1) [data-testid="stMetric"] {
    border-left-color: #6366f1;
}

.st-key-kpi_row div[data-testid="column"]:nth-of-type(1) [data-testid="stMetricValue"] {
    color: #4f46e5 !important;
}

.st-key-kpi_row div[data-testid="column"]:nth-of-type(2) [data-testid="stMetric"] {
    border-left-color: #a855f7;
}

.st-key-kpi_row div[data-testid="column"]:nth-of-type(2) [data-testid="stMetricValue"] {
    color: #9333ea !important;
}

.st-key-kpi_row div[data-testid="column"]:nth-of-type(3) [data-testid="stMetric"] {
    border-left-color: #f43f5e;
}

.st-key-kpi_row div[data-testid="column"]:nth-of-type(3) [data-testid="stMetricValue"] {
    color: #e11d48 !important;
}

.st-key-kpi_row div[data-testid="column"]:nth-of-type(4) [data-testid="stMetric"] {
    border-left-color: #10b981;
}

.st-key-kpi_row div[data-testid="column"]:nth-of-type(4) [data-testid="stMetricValue"] {
    color: #059669 !important;
}

[data-testid="stMetricValue"] {
    color: #7c3aed;
    font-weight: 700 !important;
    font-size: 1.35rem !important;
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: unset !important;
    line-height: 1.25 !important;
    word-break: break-word;
}

[data-testid="stMetricValue"] > div {
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: unset !important;
}

[data-testid="stMetricLabel"] {
    color: #64748b !important;
    white-space: normal !important;
    overflow: visible !important;
}

[data-testid="stMetricDelta"] {
    white-space: normal !important;
    overflow: visible !important;
}

[data-testid="stContainer"],
div[data-testid="stVerticalBlockBorderWrapper"] {
    background: #ffffff;
    border: 1px solid rgba(15,23,42,0.06);
    border-radius: 14px;
    box-shadow: 0 2px 10px rgba(15,23,42,0.05);
}

[data-testid="stDataFrame"] {
    background: #ffffff;
    border-radius: 10px;
}

hr {
    border-color: rgba(15,23,42,0.1);
}
</style>
""", unsafe_allow_html=True)

PLOT_FONT_COLOR = "#1e293b"
PLOT_GRID_COLOR = "rgba(15,23,42,0.12)"


# =============================================================================
# 2. LOAD DATA DARI EXCEL & GENERASI KEBIJAKAN SPESIFIK
# =============================================================================

@st.cache_data
def load_data():
    file_path = r"C:\Users\firda\Downloads\YOUTH-LENS_hasil_akhir.xlsx"

    if not os.path.exists(file_path):
        file_path = "YOUTH-LENS_hasil_akhir.xlsx"

    df_raw = pd.read_excel(file_path)

    return df_raw


try:
    df = load_data()

except Exception as e:
    st.error(f"Gagal membaca file data: {e}")
    st.stop()


# =============================================================================
# 3. LOAD GEOJSON 38 PROVINSI
# =============================================================================

GEOJSON_LOCAL_PATH = "indonesia-38-provinces.geojson"

# Tetap disimpan, tetapi untuk sekarang tidak digunakan.
GEOJSON_URL = (
    "https://raw.githubusercontent.com/denyherianto/"
    "indonesia-geojson-topojson-maps-with-38-provinces/main/"
    "GeoJSON/indonesia-38-provinces.geojson"
)


# Beberapa nama provinsi di data sumber kadang berbeda penulisan
# dengan nama resmi di peta.
PROV_NAME_ALIAS = {
    "DI Yogyakarta": "Daerah Istimewa Yogyakarta",
    "D.I. Yogyakarta": "Daerah Istimewa Yogyakarta",
    "DI. Yogyakarta": "Daerah Istimewa Yogyakarta",
    "DIY": "Daerah Istimewa Yogyakarta",
    "Yogyakarta": "Daerah Istimewa Yogyakarta",
    "Jakarta": "DKI Jakarta",
    "DKI": "DKI Jakarta",
    "Bangka Belitung": "Kepulauan Bangka Belitung",
    "Kep. Bangka Belitung": "Kepulauan Bangka Belitung",
    "Kep Bangka Belitung": "Kepulauan Bangka Belitung",
    "Babel": "Kepulauan Bangka Belitung",
    "Kepri": "Kepulauan Riau",
    "Kep. Riau": "Kepulauan Riau",
    "Kep Riau": "Kepulauan Riau",
}


def _normalize_name(s: str) -> str:
    s = str(s).strip().lower()
    s = s.replace(".", "")
    s = re.sub(r"\s+", " ", s)

    return s


@st.cache_data(show_spinner=False)
def load_geojson():

    # SEKARANG DIPAKSA MENGGUNAKAN FILE LOKAL
    with open(GEOJSON_LOCAL_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


try:
    geojson_provinsi = load_geojson()
    geojson_ok = True

except Exception as e:
    geojson_provinsi = None
    geojson_ok = False

    st.error(f"Gagal membaca GeoJSON: {e}")


if geojson_ok:

    _geo_props_names = [
        f["properties"]["PROVINSI"]
        for f in geojson_provinsi["features"]
    ]

    _geo_norm_lookup = {
        _normalize_name(n): n
        for n in _geo_props_names
    }

    _alias_norm_lookup = {
        _normalize_name(k): v
        for k, v in PROV_NAME_ALIAS.items()
    }


    def to_geo_name(nama_provinsi: str) -> str:

        norm = _normalize_name(nama_provinsi)

        # 1. Coba alias manual dulu
        if norm in _alias_norm_lookup:
            return _alias_norm_lookup[norm]

        # 2. Coba cocokkan setelah dinormalisasi
        if norm in _geo_norm_lookup:
            return _geo_norm_lookup[norm]

        # 3. Coba pencocokan mirip
        mirip = get_close_matches(
            norm,
            _geo_norm_lookup.keys(),
            n=1,
            cutoff=0.87
        )

        if mirip:
            return _geo_norm_lookup[mirip[0]]

        # 4. Jika tidak cocok, biarkan apa adanya
        return nama_provinsi


    df["Provinsi_Geo"] = df["Provinsi"].apply(to_geo_name)

else:

    df["Provinsi_Geo"] = df["Provinsi"]


# =============================================================================
# 4. GENERASI REKOMENDASI KEBIJAKAN
# =============================================================================

def get_detailed_policy(prov, tipologi, cluster, lisa, yvi):

    if cluster == 1:

        return (
            "1. Intervensi Afirmatif Terpadu: Program wajib belajar "
            "berbasis beasiswa penuh dan asrama bagi pemuda pedalaman.\n"
            "2. Akselerasi Infrastruktur Digital: Pembangunan BTS 4G USO "
            "dan penyediaan ruang belajar digital komunitas gratis.\n"
            "3. Penanganan Lintas Batas (Hotspot Spasial): Kolaborasi "
            "antardaerah di kawasan Papua untuk penyetaraan rasio guru "
            "dan tenaga pelatih vokasi vokasional."
        )

    elif cluster == 3:

        if "High-High" in str(lisa):

            return (
                "1. Program Transisi Pendidikan-Pekerjaan: Pembentukan "
                "Balai Latihan Kerja (BLK) Maritim dan Komoditas Lokal "
                "terpadu antardaerah tetangga.\n"
                "2. Program Padat Karya Muda: Insentif proyek infrastruktur "
                "daerah yang memprioritaskan tenaga kerja pemuda lokal "
                "usia 16-24 tahun.\n"
                "3. Subsidi Akses Permodalan Wirausaha: Skema kredit pemuda "
                "mikro tanpa agunan untuk menekan angka NEET di zona "
                "hotspot rentan."
            )

        elif "Low-High" in str(lisa):

            return (
                "1. Fasilitas Pusat Logistik Ketenagakerjaan: Memanfaatkan "
                "posisi strategis daerah sebagai penghubung pusat "
                "pertumbuhan ekonomi kawasan sekitar.\n"
                "2. Program Link and Match Industri Khusus: Kerjasama vokasi "
                "dengan industri pengolahan/jasa untuk menyerap pemuda "
                "terdidik.\n"
                "3. Buffer Kebijakan Spasial: Mencegah limpahan pengangguran "
                "musiman dari daerah hotspot tetangga melalui registrasi "
                "ketenagakerjaan digital."
            )

        elif "Low-Low" in str(lisa):

            return (
                "1. Upskilling Berbasis Teknologi Maju: Pelatihan vokasi "
                "industri manufaktur berteknologi tinggi, otomasi, dan "
                "ekonomi digital.\n"
                "2. Insentif Retensi Tenaga Kerja Formal: Pengurangan pajak "
                "daerah bagi industri yang merekrut pemuda fresh graduate "
                "dengan status kontrak formal berkelanjutan.\n"
                "3. Inkubator Bisnis Kreatif Pemuda: Penguatan ekosistem "
                "startup dan sertifikasi keahlian terstandarisasi nasional."
            )

        else:

            return (
                "1. Revitalisasi Sekolah Menengah Kejuruan (SMK): "
                "Penyelarasan kurikulum kejuruan dengan kebutuhan riil "
                "pasar kerja lokal.\n"
                "2. Perluasan Program Magang Bersertifikat: Kolaborasi "
                "dinas ketenagakerjaan daerah dan sektor swasta lokal "
                "dengan subsidi uang saku.\n"
                "3. Sentra Kewirausahaan Pemuda: Pendampingan mentoring "
                "bisnis dan digital marketing bagi pelaku usaha muda "
                "sektor jasa/ritel."
            )

    elif cluster == 2:

        return (
            "1. Replikasi Praktik Baik (Knowledge Sharing): Menjadi pusat "
            "transfer inovasi tata kelola kepemudaan bagi provinsi-provinsi "
            "di sekitarnya.\n"
            "2. Penguatan Karier Berkelanjutan: Program beasiswa riset, "
            "kepemimpinan pemuda tingkat lanjut, dan pengembangan talenta "
            "global.\n"
            "3. Pemeliharaan Ketahanan Sosioekonomi: Pengawasan terhadap "
            "potensi disparitas pendapatan perkotaan agar pemuda rentan "
            "marjinal tetap terproteksi."
        )

    else:

        if "Low-Low" in str(lisa):

            return (
                "1. Penguatan Ekosistem Inovasi & Pendidikan Tinggi: "
                "Optimalisasi iklim belajar dan budaya untuk mendorong "
                "daya saing pemuda ke level nasional.\n"
                "2. Perlindungan Pekerja Lepas (Gig Economy): Skema "
                "jaminan sosial ketenagakerjaan bagi pemuda yang bekerja "
                "di sektor kreatif dan pariwisata.\n"
                "3. Pencegahan Underemployment: Fasilitasi penempatan "
                "kerja yang sesuai dengan tingkat kualifikasi pendidikan."
            )

        else:

            return (
                "1. Intervensi Multisektoral Seimbang: Penataan kebijakan "
                "terkoordinasi antara dinas pendidikan, kesehatan, dan "
                "tenaga kerja tanpa dominasi sektor tunggal.\n"
                "2. Sistem Monitoring Berkala (Early Warning): Pemantauan "
                "indikator kerentanan pemuda per semester guna mengantisipasi "
                "pergeseran menuju klaster berisiko.\n"
                "3. Program Pembinaan Komunitas Pemuda: Penguatan peran "
                "organisasi karang taruna dan forum kepemudaan dalam "
                "pemberdayaan ekonomi desa."
            )


df["Rekomendasi_Spesifik"] = df.apply(
    lambda r: get_detailed_policy(
        r["Provinsi"],
        r["Tipologi"],
        r["Cluster"],
        r["LISA_quadrant"],
        r["YVI"]
    ),
    axis=1
)


# =============================================================================
# 5. SIDEBAR & FILTER
# =============================================================================

with st.sidebar:

    st.markdown("### Filter Analisis")

    st.caption(
        "Youth Vulnerability Index (YVI) & Spatial Clustering"
    )

    st.markdown("---")

    all_prov = [
        "Semua Provinsi"
    ] + sorted(df["Provinsi"].tolist())

    selected_prov = st.selectbox(
        "Wilayah Spesifik:",
        all_prov
    )

    tipologi_list = df["Tipologi"].unique().tolist()

    selected_tipologi = st.multiselect(
        "Tipologi Klaster:",
        options=tipologi_list,
        default=tipologi_list
    )

    lisa_list = df["LISA_quadrant"].unique().tolist()

    selected_lisa = st.multiselect(
        "Kuadran LISA:",
        options=lisa_list,
        default=lisa_list
    )


filtered_df = df[
    (df["Tipologi"].isin(selected_tipologi)) &
    (df["LISA_quadrant"].isin(selected_lisa))
]


if selected_prov != "Semua Provinsi":

    filtered_df = filtered_df[
        filtered_df["Provinsi"] == selected_prov
    ]


# =============================================================================
# 6. HEADER UTAMA & KPI METRIK
# =============================================================================

mean_val = (
    filtered_df["YVI"].mean()
    if not filtered_df.empty
    else 0
)

rentan = (
    filtered_df.sort_values(
        by="YVI",
        ascending=False
    ).iloc[0]
    if not filtered_df.empty
    else None
)

tangguh = (
    filtered_df.sort_values(
        by="YVI",
        ascending=True
    ).iloc[0]
    if not filtered_df.empty
    else None
)


r_nama = (
    rentan["Provinsi"]
    if rentan is not None
    else "-"
)

r_val = (
    f"YVI: {rentan['YVI']:.1f}"
    if rentan is not None
    else "-"
)

t_nama = (
    tangguh["Provinsi"]
    if tangguh is not None
    else "-"
)

t_val = (
    f"YVI: {tangguh['YVI']:.1f}"
    if tangguh is not None
    else "-"
)


st.title(
    "Dashboard Analisis Youth Vulnerability Index (YVI)"
)

st.caption(
    "Pemantauan disparitas kerentanan pemuda antardaerah di Indonesia "
    "berbasis indeks komposit, reduksi dimensi (PCA), klasterisasi "
    "K-Means, dan autokorelasi spasial (LISA)."
)

st.markdown("---")


with st.container(key="kpi_row"):

    col1, col2, col3, col4 = st.columns(4)

    col1.metric(
        "Cakupan Wilayah",
        f"{len(filtered_df)} / {len(df)} Provinsi"
    )

    col2.metric(
        "Rata-Rata YVI",
        f"{mean_val:.2f}"
    )

    col3.metric(
        "Kerentanan Tertinggi",
        r_nama,
        r_val
    )

    col4.metric(
        "Kerentanan Terendah",
        t_nama,
        t_val
    )


# =============================================================================
# 7. TAB UTAMA
# =============================================================================

tab1, tab2, tab3 = st.tabs([
    "Peringkat & Disparitas YVI",
    "Tipologi Klaster (K-Means)",
    "Pola Spasial & Rekomendasi Kebijakan"
])


color_map = {
    "Digitally Excluded & Education Deprived": "#f43f5e",
    "Employment Insecure": "#f59e0b",
    "Moderate / No Dominant Domain": "#8b5cf6",
    "Relatively Resilient": "#10b981"
}


YVI_SCALE = [
    "#1E1B4B",  
    "#312E81",
    "#4338CA",
    "#6366F1", 
    "#7C3AED",  
    "#A855F7",  
    "#C026D3",  
    "#E11D73",  
    "#F43F5E"   
]


# =============================================================================
# 8. TAB 1: PERINGKAT YVI
# =============================================================================

with tab1:

    st.subheader(
        "Distribusi Peringkat Youth Vulnerability Index"
    )

    st.caption(
        "Peringkat kerentanan komposit dari tingkat tertinggi "
        "(paling rentan) hingga terendah."
    )

    if filtered_df.empty:

        st.warning(
            "Tidak ada provinsi yang memenuhi kriteria filter."
        )

    else:

        df_rank = filtered_df.sort_values(
            by="YVI",
            ascending=True
        )

        fig_bar = px.bar(
            df_rank,
            x="YVI",
            y="Provinsi",
            orientation="h",
            color="YVI",
            color_continuous_scale=YVI_SCALE,
            hover_data={
                "Tipologi": True,
                "LISA_quadrant": True,
                "YVI": ":.2f"
            },
            height=max(
                500,
                len(df_rank) * 22
            )
        )

        fig_bar.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color=PLOT_FONT_COLOR),
            xaxis=dict(
                showgrid=True,
                gridcolor=PLOT_GRID_COLOR,
                title="Skor Indeks (0 = Tangguh, 100 = Paling Rentan)"
            ),
            yaxis=dict(title=""),
            coloraxis_colorbar=dict(
                title="Skor YVI"
            ),
            margin=dict(
                l=10,
                r=10,
                t=20,
                b=10
            )
        )

        st.plotly_chart(
            fig_bar,
            width="stretch"
        )


    st.info(
        "Kawasan Papua menunjukkan rentang disparitas yang signifikan: "
        "Papua Pegunungan (100.0), Papua Tengah (83.9), dan Papua Selatan "
        "(69.1) menempati kelompok indeks tertinggi, sedangkan Papua (38.9) "
        "dan Papua Barat Daya (30.7) mencatatkan skor yang jauh lebih rendah. "
        "Hal ini memperlihatkan bahwa kedekatan geografis tidak otomatis "
        "menghasilkan tingkat kerentanan yang homogen."
    )


# =============================================================================
# 9. TAB 2: TIPOLOGI KLASTER
# =============================================================================

with tab2:

    st.subheader(
        "Tipologi Kerentanan Pemuda Berdasarkan Klasterisasi K-Means"
    )

    st.caption(
        "Pengelompokan 38 provinsi berdasarkan skor 3 komponen utama "
        "PCA (Silhouette Score: 0,376)."
    )

    col_pie, col_box = st.columns([5, 5])


    with col_pie:

        tip_count = (
            filtered_df["Tipologi"]
            .value_counts()
            .reset_index()
        )

        tip_count.columns = [
            "Tipologi",
            "Jumlah"
        ]

        fig_pie = px.pie(
            tip_count,
            names="Tipologi",
            values="Jumlah",
            hole=0.5,
            color="Tipologi",
            color_discrete_map=color_map,
            title="Proporsi Jumlah Provinsi per Tipologi"
        )

        fig_pie.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            font=dict(color=PLOT_FONT_COLOR),
            margin=dict(
                t=40,
                b=10,
                l=10,
                r=10
            )
        )

        fig_pie.update_traces(
            textposition="inside",
            textinfo="percent+value"
        )

        st.plotly_chart(
            fig_pie,
            width="stretch"
        )


    with col_box:

        fig_box = px.box(
            filtered_df,
            x="Tipologi",
            y="YVI",
            color="Tipologi",
            points="all",
            hover_name="Provinsi",
            color_discrete_map=color_map,
            title="Sebaran Skor YVI per Tipologi"
        )

        fig_box.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color=PLOT_FONT_COLOR),
            yaxis=dict(
                showgrid=True,
                gridcolor=PLOT_GRID_COLOR,
                title="Skor YVI"
            ),
            xaxis=dict(title=""),
            showlegend=False,
            margin=dict(
                t=40,
                b=10,
                l=10,
                r=10
            )
        )

        st.plotly_chart(
            fig_box,
            width="stretch"
        )


    st.write(
        "##### Ringkasan Statistik Tipologi"
    )

    summary_tabel = (
        filtered_df
        .groupby("Tipologi")
        .agg(
            Jumlah_Provinsi=("Provinsi", "count"),
            Rata_Rata_YVI=("YVI", "mean"),
            Min_YVI=("YVI", "min"),
            Max_YVI=("YVI", "max")
        )
        .reset_index()
    )

    summary_tabel["Rata_Rata_YVI"] = (
        summary_tabel["Rata_Rata_YVI"]
        .round(2)
    )

    st.dataframe(
        summary_tabel,
        width="stretch"
    )


# =============================================================================
# 10. TAB 3: SPASIAL LISA & KEBIJAKAN
# =============================================================================

with tab3:

    st.subheader(
        "Autokorelasi Spasial Lokal (LISA) & Arah Intervensi Kebijakan"
    )

    st.caption(
        "Identifikasi aglomerasi spasial untuk memetakan "
        "penularan kerentanan antardaerah."
    )

    st.markdown(
        "##### Peta Sebaran Tipologi K-Means Youth Vulnerability Antarprovinsi"
    )


    if filtered_df.empty:

        st.warning(
            "Tidak ada provinsi yang memenuhi kriteria filter."
        )

    elif not geojson_ok:

        st.warning(
            "Peta tidak dapat dimuat "
            "(gagal mengambil data batas wilayah). "
            "Coba refresh halaman."
        )

    else:

        # =============================================================
        # PETA
        # =============================================================

        fig_map = px.choropleth_map(
            filtered_df,
            geojson=geojson_provinsi,
            locations="Provinsi_Geo",
            featureidkey="properties.PROVINSI",
            color="Tipologi",
            color_discrete_map=color_map,
            range_color=(0, 100),
            hover_name="Provinsi",
            hover_data={
                "Tipologi": True,
                "LISA_quadrant": True,
                "YVI": ":.2f",
                "Provinsi_Geo": False
            },
            center={"lat": -2.5, "lon": 118},
            zoom=3.5,
            map_style="carto-positron",
            opacity=0.75
        )
        
        fig_map.update_layout(
            margin=dict(l=0, r=0, t=10, b=0),
            coloraxis_colorbar=dict(title="Skor YVI"),
            height=500
        )
        
        st.plotly_chart(
            fig_map,
            width="stretch",
            key="yvi_map"
        )
    
        # =============================================================
        # CEK MATCHING NAMA PROVINSI
        # =============================================================

        nama_teridentifikasi = {
            f["properties"]["PROVINSI"]
            for f in geojson_provinsi["features"]
        }

        prov_cocok = [
            p
            for p in filtered_df["Provinsi_Geo"]
            if p in nama_teridentifikasi
        ]

        prov_tak_ketemu = sorted(
            set(filtered_df["Provinsi_Geo"])
            - nama_teridentifikasi
        )


        st.caption(
            f"Berhasil dipetakan: "
            f"{len(prov_cocok)} dari "
            f"{len(filtered_df)} provinsi."
        )


        if prov_tak_ketemu:

            st.caption(
                "Catatan: provinsi berikut belum cocok "
                "dengan nama di batas wilayah peta, "
                "jadi tidak tampil di peta "
                "(namun tetap tampil di chart lain): "
                + ", ".join(prov_tak_ketemu)
            )


    st.markdown("---")


    # =============================================================================
    # LISA
    # =============================================================================

    col_lisa_chart, col_lisa_text = st.columns([5, 5])


    with col_lisa_chart:

        lisa_count = (
            filtered_df["LISA_quadrant"]
            .value_counts()
            .reset_index()
        )

        lisa_count.columns = [
            "Status LISA",
            "Jumlah"
        ]

        fig_lisa = px.bar(
            lisa_count,
            x="Status LISA",
            y="Jumlah",
            color="Status LISA",
            color_discrete_map={
                "High-High (hotspot rentan)": "#f43f5e",
                "Low-Low (coldspot/aman)": "#10b981",
                "Low-High": "#f59e0b",
                "Tidak signifikan": "#94a3b8"
            }
        )

        fig_lisa.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color=PLOT_FONT_COLOR),
            yaxis=dict(
                showgrid=True,
                gridcolor=PLOT_GRID_COLOR,
                title="Jumlah Provinsi"
            ),
            xaxis=dict(title=""),
            showlegend=False,
            margin=dict(
                t=10,
                b=10,
                l=10,
                r=10
            )
        )

        st.plotly_chart(
            fig_lisa,
            width="stretch"
        )


    with col_lisa_text:

        st.markdown("""
        **Pola Hubungan Antarwilayah:**

        * **High-High (Hotspot Rentan):** Terpusat di wilayah Papua Pegunungan,
          Papua Tengah, Papua Selatan, Papua, dan Maluku. Pola ini mengindikasikan
          adanya perangkap kerentanan spasial (*spatial poverty trap*) yang
          memerlukan intervensi terpadu lintas perbatasan.

        * **Low-Low (Coldspot Aman):** Terkonsentrasi di Pulau Jawa-Bali
          (DI Yogyakarta, Bali, Jawa Barat, Jawa Tengah, Jawa Timur).

        * **Low-High (Spatial Outlier):** Teridentifikasi pada
          **Papua Barat Daya**, yaitu wilayah dengan tingkat kerentanan
          relatif lebih rendah di tengah kawasan hotspot rentan.

        * **Tidak Signifikan:** Mencakup mayoritas provinsi (27 dari 38),
          yang berarti tingkat kerentanan pemuda di wilayah tersebut tidak
          menunjukkan pola asosiasi spasial yang cukup kuat dengan provinsi
          tetangganya, baik ke arah rentan maupun tangguh.
        """)


    # =============================================================================
    # REKOMENDASI KEBIJAKAN
    # =============================================================================

    st.markdown("---")

    st.subheader(
        "Tinjauan Kebijakan Intervensi per Provinsi"
    )


    target_prov = st.selectbox(
        "Pilih Provinsi Sasaran:",
        sorted(df["Provinsi"].unique())
    )


    row = df[
        df["Provinsi"] == target_prov
    ].iloc[0]


    with st.container(border=True):

        st.markdown(
            f"#### Provinsi: {row['Provinsi']}"
        )

        st.caption(
            f"Tipologi: **{row['Tipologi']}** | "
            f"Status Spasial: **{row['LISA_quadrant']}**"
        )


        col_m1, col_m2, col_m3 = st.columns(3)


        col_m1.metric(
            "Skor YVI",
            f"{row['YVI']:.2f}"
        )


        col_m2.metric(
            "Klaster K-Means",
            f"Cluster {row['Cluster']}"
        )


        col_m3.metric(
            "Kuadran LISA",
            f"{row['LISA_quadrant']}"
        )


        st.markdown(
            "**Program Intervensi Terarah (Action Plan):**"
        )


        st.info(
            row["Rekomendasi_Spesifik"]
        )
