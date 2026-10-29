"""
Dashboard Layanan Pelanggan – Gudeg Sekar Arum (data simulasi 2025)
Jalankan:  streamlit run app.py
"""
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# ─────────────────────────────── Konfigurasi ───────────────────────────────
st.set_page_config(
    page_title="Dashboard Layanan Pelanggan · Gudeg Sekar Arum",
    page_icon="🍛",
    layout="wide",
)

FILE_BAWAAN = Path(__file__).parent / "data_gudeg.xlsx"
HARI = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
BULAN = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]
WARNA_OUTLET = {
    "Malioboro": "#F0C24B",
    "Wijilan": "#FFFFFF",
    "Kaliurang": "#7FA3D6",
    "Godean": "#B08D57",
}
WARNA_KANAL = {
    "Makan di tempat": "#F0C24B",
    "Bawa pulang": "#FFFFFF",
    "Pesan antar online": "#7FA3D6",
}
SKALA_EMAS = ["#1B3A66", "#8A7A4A", "#D4AF37", "#FFE9A0"]  # navy → emas
WARNA_PELANGGAN = {"Wisatawan": "#F0C24B", "Lokal": "#DCE6F5"}
KOLOM_WAJIB = [
    "Tanggal", "Hari", "Jenis_Hari", "Outlet", "Kanal_Penjualan", "Jumlah_Transaksi",
    "Transaksi_Wisatawan", "Transaksi_Lokal", "Omzet_Bersih_Rp", "Diskon_Rp",
    "Rating_Rata_rata", "Jumlah_Ulasan",
]

st.markdown(
    """
    <style>
    /* ===== Palet tema navy (ubah di sini jika ingin mengganti) ===== */
    :root {
        --latar-1: #0A1A33;       /* latar halaman (gradasi atas) */
        --latar-2: #10294C;       /* latar halaman (gradasi bawah) */
        --latar-sidebar: #081530; /* panel samping */
        --kartu: #132B4F;         /* kartu metrik, grafik, kolom input */
        --garis: #2A4A7F;         /* garis tepi */
        --aksen: #D4AF37;         /* emas: tombol, chip, aksen */
        --aksen-terang: #F0C24B;  /* emas terang: hover, tab aktif */
        --teks-di-emas: #0A1A33;  /* teks di atas tombol/chip emas */
        --teks: #FFFFFF;          /* teks utama (putih) */
        --teks-redup: #B9C7DD;    /* teks keterangan */
    }
    .stApp {
        background: linear-gradient(160deg, var(--latar-1) 0%, var(--latar-2) 100%) fixed;
        color: var(--teks);
    }
    [data-testid="stHeader"] {background: transparent;}
    [data-testid="stToolbar"], [data-testid="stToolbar"] * {color: var(--teks);}
    .block-container {padding-top: 2rem;}

    /* ----- Teks ----- */
    .stApp h1, .stApp h2, .stApp h3,
    [data-testid="stMarkdownContainer"] p, [data-testid="stMarkdownContainer"] li,
    [data-testid="stMarkdownContainer"] strong,
    [data-testid="stWidgetLabel"], [data-testid="stWidgetLabel"] *,
    [data-testid="stRadio"] label, [data-testid="stRadio"] label * {color: var(--teks);}
    [data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] * {color: var(--teks-redup);}

    /* ----- Panel samping ----- */
    [data-testid="stSidebar"] {background-color: var(--latar-sidebar); border-right: 1px solid var(--garis);}
    [data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"] button,
    [data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"] svg {color: var(--teks);}
    [data-testid="stSidebar"] [data-baseweb="select"] > div,
    [data-testid="stSidebar"] [data-baseweb="input"],
    [data-testid="stSidebar"] [data-baseweb="base-input"] {
        background-color: var(--kartu) !important; border-color: var(--garis) !important;
    }
    [data-testid="stSidebar"] input {color: var(--teks) !important; -webkit-text-fill-color: var(--teks) !important;}
    [data-testid="stSidebar"] [data-baseweb="select"] svg,
    [data-testid="stSidebar"] [data-baseweb="input"] svg {fill: var(--teks-redup) !important; color: var(--teks-redup) !important;}
    [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
        background-color: var(--kartu); border: 1px dashed var(--garis);
    }
    [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] * {color: var(--teks-redup);}
    [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button {
        background-color: var(--aksen) !important; border-color: var(--aksen) !important; color: var(--teks-di-emas) !important;
    }
    [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button * {color: var(--teks-di-emas) !important;}

    /* ----- Kartu metrik ----- */
    [data-testid="stMetric"] {
        background: var(--kartu); border: 1px solid var(--garis);
        border-radius: 12px; padding: 14px 16px;
        border-top: 3px solid var(--aksen);
    }
    [data-testid="stMetricLabel"], [data-testid="stMetricLabel"] * {color: var(--teks-redup) !important;}
    [data-testid="stMetricValue"], [data-testid="stMetricValue"] * {color: #FFFFFF !important;}

    /* ----- Tombol ----- */
    .stButton > button, .stDownloadButton > button {
        background-color: var(--aksen); border: 1px solid var(--aksen); border-radius: 8px;
    }
    .stButton > button *, .stDownloadButton > button * {color: var(--teks-di-emas) !important; font-weight: 600;}
    .stButton > button:hover, .stDownloadButton > button:hover,
    .stButton > button:focus, .stDownloadButton > button:focus,
    .stButton > button:active, .stDownloadButton > button:active {
        background-color: var(--aksen-terang); border-color: var(--aksen-terang);
    }

    /* ----- Tab ----- */
    button[data-baseweb="tab"], button[data-baseweb="tab"] * {color: var(--teks-redup) !important;}
    button[data-baseweb="tab"][aria-selected="true"], button[data-baseweb="tab"][aria-selected="true"] * {color: var(--aksen-terang) !important;}
    [data-baseweb="tab-highlight"] {background-color: var(--aksen-terang) !important;}
    [data-baseweb="tab-border"] {background-color: var(--garis) !important;}

    /* ----- Chip filter & slider ----- */
    span[data-baseweb="tag"] {background-color: var(--aksen) !important;}
    span[data-baseweb="tag"], span[data-baseweb="tag"] * {color: var(--teks-di-emas) !important;}
    [data-testid="stSlider"] div[role="slider"] {background-color: var(--aksen-terang) !important;}
    [data-testid="stSlider"] [data-testid="stThumbValue"], [data-testid="stSlider"] [data-testid="stTickBarMin"],
    [data-testid="stSlider"] [data-testid="stTickBarMax"] {color: var(--teks) !important;}

    /* ----- Grafik, tabel, kotak info ----- */
    [data-testid="stPlotlyChart"] {
        border: 1px solid var(--garis); border-radius: 12px; overflow: hidden;
    }
    [data-testid="stDataFrame"] {
        border: 1px solid var(--garis); border-radius: 12px; overflow: hidden;
        filter: invert(0.92) hue-rotate(180deg);
    }
    [data-testid="stAlert"] {background-color: #17335E; border: 1px solid var(--garis); border-radius: 12px;}
    [data-testid="stAlert"] * {color: var(--teks);}
    </style>
    """,
    unsafe_allow_html=True,
)


# ─────────────────────────────── Fungsi bantu ──────────────────────────────
def angka(n) -> str:
    """Format angka gaya Indonesia: 1.234.567"""
    return f"{n:,.0f}".replace(",", ".")


def desimal(n, d=2) -> str:
    return f"{n:,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")


@st.cache_data(show_spinner="Memuat data…")
def muat_data(isi_file: bytes) -> pd.DataFrame:
    df = pd.read_excel(BytesIO(isi_file), sheet_name="Data_Penjualan")
    kurang = [k for k in KOLOM_WAJIB if k not in df.columns]
    if kurang:
        raise ValueError(f"Kolom berikut tidak ditemukan di sheet Data_Penjualan: {', '.join(kurang)}")
    df["Tanggal"] = pd.to_datetime(df["Tanggal"])
    df["Bulan_No"] = df["Tanggal"].dt.month
    df["Bulan_Singkat"] = pd.Categorical(
        df["Bulan_No"].map(lambda m: BULAN[m - 1]), BULAN, ordered=True
    )
    df["Hari"] = pd.Categorical(df["Hari"], HARI, ordered=True)
    # Rating dirata-ratakan dengan bobot jumlah ulasan (hari tanpa ulasan tidak dihitung)
    ada_rating = df["Rating_Rata_rata"].notna()
    df["Skor_Ulasan"] = df["Rating_Rata_rata"].fillna(0) * df["Jumlah_Ulasan"]
    df["Ulasan_Berating"] = np.where(ada_rating, df["Jumlah_Ulasan"], 0)
    return df


def ringkas(d: pd.DataFrame, by) -> pd.DataFrame:
    """Agregasi dengan rating tertimbang ulasan."""
    g = (
        d.groupby(by, observed=True)
        .agg(
            Skor=("Skor_Ulasan", "sum"),
            Ulasan=("Ulasan_Berating", "sum"),
            Transaksi=("Jumlah_Transaksi", "sum"),
            Wisatawan=("Transaksi_Wisatawan", "sum"),
            Lokal=("Transaksi_Lokal", "sum"),
            Omzet=("Omzet_Bersih_Rp", "sum"),
        )
        .reset_index()
    )
    g["Rating"] = g["Skor"] / g["Ulasan"].replace(0, np.nan)
    g["Pct_Wisatawan"] = g["Wisatawan"] / g["Transaksi"].replace(0, np.nan) * 100
    g["Nilai_per_Transaksi"] = g["Omzet"] / g["Transaksi"].replace(0, np.nan)
    return g


def rating_total(d: pd.DataFrame) -> float:
    u = d["Ulasan_Berating"].sum()
    return d["Skor_Ulasan"].sum() / u if u else np.nan


def rapikan(fig, tinggi=360, legenda=True):
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="#132B4F",
        plot_bgcolor="#132B4F",
        font=dict(color="#EAF2FF"),
        height=tinggi,
        margin=dict(l=10, r=10, t=50, b=10),
        legend_title_text="",
        showlegend=legenda,
        title_font_size=15,
    )
    fig.update_xaxes(gridcolor="#2A4A7F", zerolinecolor="#2A4A7F")
    fig.update_yaxes(gridcolor="#2A4A7F", zerolinecolor="#2A4A7F")
    return fig


# ─────────────────────────────── Sumber data ───────────────────────────────
st.sidebar.title("🍛 Gudeg Sekar Arum")
st.sidebar.caption("Dashboard Tim Layanan Pelanggan · 2025")

unggahan = st.sidebar.file_uploader("Ganti sumber data (opsional)", type=["xlsx"])
if unggahan is not None:
    isi = unggahan.getvalue()
elif FILE_BAWAAN.exists():
    isi = FILE_BAWAAN.read_bytes()
else:
    st.error("File data_gudeg.xlsx tidak ditemukan. Unggah file Excel lewat panel samping.")
    st.stop()

try:
    data = muat_data(isi)
except Exception as e:  # noqa: BLE001
    st.error(f"Data tidak dapat dibaca: {e}")
    st.stop()

# ─────────────────────────────── Filter ────────────────────────────────────
st.sidebar.header("Filter")
tgl_min, tgl_max = data["Tanggal"].min().date(), data["Tanggal"].max().date()
rentang = st.sidebar.date_input(
    "Periode", value=(tgl_min, tgl_max), min_value=tgl_min, max_value=tgl_max
)
outlet_pilih = st.sidebar.multiselect(
    "Outlet", sorted(data["Outlet"].unique()), default=sorted(data["Outlet"].unique())
)
kanal_pilih = st.sidebar.multiselect(
    "Kanal penjualan", sorted(data["Kanal_Penjualan"].unique()),
    default=sorted(data["Kanal_Penjualan"].unique()),
)
jenis_pilih = st.sidebar.multiselect(
    "Jenis hari", ["Hari kerja", "Akhir pekan", "Libur nasional"],
    default=["Hari kerja", "Akhir pekan", "Libur nasional"],
)

if isinstance(rentang, (tuple, list)) and len(rentang) == 2:
    awal, akhir = rentang
else:
    awal = akhir = rentang[0] if isinstance(rentang, (tuple, list)) else rentang

df = data[
    data["Tanggal"].dt.date.between(awal, akhir)
    & data["Outlet"].isin(outlet_pilih)
    & data["Kanal_Penjualan"].isin(kanal_pilih)
    & data["Jenis_Hari"].isin(jenis_pilih)
]

st.sidebar.divider()
st.sidebar.caption(
    "Data fiktif hasil simulasi. Rating dihitung sebagai rata-rata **tertimbang jumlah ulasan**; "
    "hari tanpa ulasan tidak dihitung sebagai nol."
)

# ─────────────────────────────── Judul & KPI ───────────────────────────────
st.title("Dashboard Layanan Pelanggan")
st.caption(
    f"Gudeg Sekar Arum · {awal:%d %b %Y} – {akhir:%d %b %Y} · "
    f"{len(df):,} catatan harian".replace(",", ".")
)

if df.empty:
    st.warning("Tidak ada data untuk kombinasi filter ini. Longgarkan filter di panel samping.")
    st.stop()

r_all = rating_total(data)
r_sel = rating_total(df)
trx = df["Jumlah_Transaksi"].sum()
ulasan = df["Jumlah_Ulasan"].sum()
pct_wisata = df["Transaksi_Wisatawan"].sum() / trx * 100 if trx else 0
harian = df.groupby("Tanggal")["Jumlah_Transaksi"].sum()

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Total transaksi", angka(trx))
k2.metric(
    "Rating rata-rata",
    f"{desimal(r_sel)} ★" if not np.isnan(r_sel) else "–",
    delta=f"{r_sel - r_all:+.2f} vs seluruh data".replace(".", ",") if not np.isnan(r_sel) else None,
)
k3.metric("Total ulasan masuk", angka(ulasan))
k4.metric("Ulasan per 100 transaksi", desimal(ulasan / trx * 100, 1) if trx else "–")
k5.metric("Pelanggan wisatawan", f"{desimal(pct_wisata, 1)}%")

# ─────────────────────────────── Temuan otomatis ───────────────────────────
temuan = []
per_outlet = ringkas(df, "Outlet").dropna(subset=["Rating"])
if len(per_outlet) > 1:
    rendah = per_outlet.sort_values("Rating").iloc[0]
    tinggi = per_outlet.sort_values("Rating").iloc[-1]
    temuan.append(
        f"**{rendah['Outlet']}** punya rating terendah ({desimal(rendah['Rating'])}), "
        f"sedangkan **{tinggi['Outlet']}** tertinggi ({desimal(tinggi['Rating'])}). "
        f"{rendah['Outlet']} juga menangani {desimal(rendah['Transaksi'] / per_outlet['Transaksi'].sum() * 100, 0)}% "
        f"dari seluruh transaksi."
    )
per_jenis = ringkas(df, "Jenis_Hari").dropna(subset=["Rating"]).set_index("Jenis_Hari")
if "Hari kerja" in per_jenis.index and len(per_jenis) > 1:
    kerja = per_jenis.loc["Hari kerja", "Rating"]
    lain = df[df["Jenis_Hari"] != "Hari kerja"]
    if rating_total(lain) == rating_total(lain):  # bukan NaN
        temuan.append(
            f"Rating turun saat akhir pekan/libur nasional: {desimal(rating_total(lain))} "
            f"dibanding {desimal(kerja)} pada hari kerja."
        )
rata_hari = (
    df.groupby(["Tanggal", "Hari"], observed=True)["Jumlah_Transaksi"].sum()
    .groupby("Hari", observed=True).mean()
)
if len(rata_hari.dropna()) > 1:
    sibuk = rata_hari.idxmax()
    temuan.append(
        f"**{sibuk}** adalah hari paling ramai (rata-rata {angka(rata_hari.max())} transaksi/hari "
        f"di semua outlet terpilih) — siapkan tim lebih banyak."
    )
per_outlet_w = per_outlet.sort_values("Pct_Wisatawan", ascending=False)
if len(per_outlet_w) > 1:
    w = per_outlet_w.iloc[0]
    l = per_outlet_w.iloc[-1]
    temuan.append(
        f"{w['Outlet']} didominasi wisatawan ({desimal(w['Pct_Wisatawan'], 0)}%), "
        f"sementara {l['Outlet']} mayoritas pelanggan lokal ({desimal(100 - l['Pct_Wisatawan'], 0)}%)."
    )
if temuan:
    with st.container(border=True):
        st.markdown("**💡 Temuan utama pada filter saat ini**")
        for t in temuan:
            st.markdown(f"- {t}")

# ─────────────────────────────── Tab ───────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs(
    ["⭐ Kepuasan pelanggan", "👥 Keramaian & staffing", "🧳 Profil pelanggan",
     "🚨 Perlu tindak lanjut", "📋 Data"]
)

# ── Tab 1: Kepuasan ──
with tab1:
    c1, c2 = st.columns(2)
    g = ringkas(df, "Outlet").dropna(subset=["Rating"]).sort_values("Rating")
    fig = px.bar(
        g, x="Rating", y="Outlet", orientation="h", text=g["Rating"].round(2),
        color="Outlet", color_discrete_map=WARNA_OUTLET,
        title="Rating rata-rata per outlet",
        hover_data={"Ulasan": ":,.0f", "Rating": ":.2f", "Outlet": False},
    )
    if not np.isnan(r_sel):
        fig.add_vline(x=r_sel, line_dash="dash", line_color="#F0C24B",
                      annotation_text=f"rata-rata {r_sel:.2f}", annotation_position="top")
    fig.update_xaxes(range=[3.5, 5])
    fig.update_traces(textposition="outside")
    fig.update_layout(
        font=dict(color="#FFFFFF"),          
        title=dict(font=dict(color="#FFFFFF")) 
    )
    c1.plotly_chart(rapikan(fig, legenda=False), width="stretch")

    g = ringkas(df, ["Outlet", "Kanal_Penjualan"]).dropna(subset=["Rating"])
    fig = px.bar(
        g, x="Outlet", y="Rating", color="Kanal_Penjualan", barmode="group",
        color_discrete_map=WARNA_KANAL, text=g["Rating"].round(2),
        title="Rating per outlet dan kanal",
    )
    fig.update_yaxes(range=[3.5, 5])
    fig.update_traces(textposition="outside")
    fig.update_layout(
        font=dict(color="#FFFFFF"),          
        title=dict(font=dict(color="#FFFFFF")), 
        legend=dict(font=dict(color="#FFFFFF"))
    )
    c2.plotly_chart(rapikan(fig), width="stretch")

    c3, c4 = st.columns(2)
    g = ringkas(df, ["Bulan_Singkat", "Outlet"]).dropna(subset=["Rating"])
    fig = px.line(
        g, x="Bulan_Singkat", y="Rating", color="Outlet", markers=True,
        color_discrete_map=WARNA_OUTLET, 
        title="Tren rating bulanan per outlet",
        category_orders={"Bulan_Singkat": BULAN},
    )
    fig.update_xaxes(title="")
    fig.update_layout(
        font=dict(color="#FFFFFF"),
        title=dict(font=dict(color="#FFFFFF")),
        legend=dict(font=dict(color="#FFFFFF"))
    )
    c3.plotly_chart(rapikan(fig), width="stretch")

    g = ringkas(df, ["Hari", "Outlet"]).dropna(subset=["Rating"])
    pv = g.pivot(index="Hari", columns="Outlet", values="Rating").reindex(HARI)
    fig = px.imshow(
        pv, text_auto=".2f", aspect="auto", color_continuous_scale=SKALA_EMAS,
        title="Rating menurut hari & outlet (emas = tinggi, navy = rendah)",
    )
    fig.update_xaxes(title="")
    fig.update_yaxes(title="")
    fig.update_coloraxes(colorbar_title="Rating")
    fig.update_layout(
        font=dict(color="#FFFFFF"),
        title=dict(font=dict(color="#FFFFFF")),
        coloraxis_colorbar=dict(
            title=dict(font=dict(color="#FFFFFF")),
            tickfont=dict(color="#FFFFFF")
        )
    )
    c4.plotly_chart(rapikan(fig), width="stretch")

    g = ringkas(df, ["Jenis_Hari", "Outlet"]).dropna(subset=["Rating"])
    fig = px.bar(
        g, x="Jenis_Hari", y="Rating", color="Outlet", barmode="group",
        color_discrete_map=WARNA_OUTLET, title="Rating menurut jenis hari",
        category_orders={"Jenis_Hari": ["Hari kerja", "Akhir pekan", "Libur nasional"]},
    )
    fig.update_yaxes(range=[3.5, 5])
    fig.update_xaxes(title="")
    fig.update_layout(
        font=dict(color="#FFFFFF"),             
        title=dict(font=dict(color="#FFFFFF")),    
        legend=dict(font=dict(color="#FFFFFF"))  
    )
    st.plotly_chart(rapikan(fig, 330), width="stretch")

    # ── Tambahan: volume ulasan per bulan ──
    st.markdown("### Volume ulasan per bulan")
    c7, c8 = st.columns(2)

    # Total ulasan per bulan, ditumpuk menurut outlet.
    vol = (
        df.groupby(["Bulan_Singkat", "Outlet"], observed=True)["Jumlah_Ulasan"]
        .sum().reset_index()
    )
    fig = px.bar(
        vol, x="Bulan_Singkat", y="Jumlah_Ulasan", color="Outlet",
        color_discrete_map=WARNA_OUTLET, title="Jumlah ulasan per bulan menurut outlet",
        category_orders={"Bulan_Singkat": BULAN},
        labels={"Jumlah_Ulasan": "Jumlah ulasan", "Bulan_Singkat": "Bulan"},
    )
    fig.update_layout(barmode="stack")
    fig.update_xaxes(title="")
    fig.update_layout(
        font=dict(color="#FFFFFF"),
        title=dict(font=dict(color="#FFFFFF")),
        legend=dict(font=dict(color="#FFFFFF"))
    )
    c7.plotly_chart(rapikan(fig), width="stretch")

    # Ulasan per 100 transaksi: volume ulasan relatif terhadap keramaian bulan itu.
    rel = (
        df.groupby("Bulan_Singkat", observed=True)
        .agg(Ulasan=("Jumlah_Ulasan", "sum"), Transaksi=("Jumlah_Transaksi", "sum"))
        .reset_index()
    )
    rel["Ulasan_per_100"] = rel["Ulasan"] / rel["Transaksi"].replace(0, np.nan) * 100
    fig = px.line(
        rel, x="Bulan_Singkat", y="Ulasan_per_100", markers=True,
        title="Ulasan per 100 transaksi tiap bulan",
        category_orders={"Bulan_Singkat": BULAN},
        labels={"Ulasan_per_100": "Ulasan per 100 transaksi", "Bulan_Singkat": "Bulan"},
        color_discrete_sequence=["#F0C24B"],
    )
    fig.update_xaxes(title="")
    fig.update_layout(
        font=dict(color="#FFFFFF"),
        title=dict(font=dict(color="#FFFFFF")),
        legend=dict(font=dict(color="#FFFFFF"))
    )
    c8.plotly_chart(rapikan(fig, legenda=False), width="stretch")

    if not rel.empty and rel["Ulasan"].sum() > 0:
        bln_max = rel.loc[rel["Ulasan"].idxmax()]
        bln_min = rel.loc[rel["Ulasan"].idxmin()]
        st.caption(
            f"Ulasan terbanyak pada bulan {bln_max['Bulan_Singkat']} ({angka(bln_max['Ulasan'])} ulasan) "
            f"dan tersedikit pada bulan {bln_min['Bulan_Singkat']} ({angka(bln_min['Ulasan'])} ulasan) "
            "pada filter saat ini. Grafik kanan membantu membedakan kenaikan ulasan karena pelanggan "
            "makin banyak dari kenaikan karena pelanggan makin aktif memberi ulasan."
        )

    # ── Tambahan Fase 3: kategori distribusi dan hubungan ──
    st.markdown("### Analisis tambahan untuk Fase 3")
    c5, c6 = st.columns(2)

    # Grafik distribusi: boxplot rating yang tersedia per outlet.
    data_box = df.dropna(subset=["Rating_Rata_rata"]).copy()
    if not data_box.empty:
        fig = px.box(
            data_box, x="Outlet", y="Rating_Rata_rata", color="Outlet",
            color_discrete_map=WARNA_OUTLET, points="outliers",
            title="Sebaran rating pelanggan menurut outlet",
            labels={"Rating_Rata_rata": "Rating", "Outlet": "Outlet"},
        )
        fig.update_yaxes(range=[1, 5])
        fig.update_layout(
                font=dict(color="#FFFFFF"),             
                title=dict(font=dict(color="#FFFFFF")),    
                legend=dict(font=dict(color="#FFFFFF"))  
            )
        c5.plotly_chart(rapikan(fig, legenda=False), width="stretch")
        c5.caption("Boxplot menunjukkan median, rentang antarkuartil, dan nilai yang menyimpang. Catatan rating adalah agregat, bukan ulasan individu.")
    else:
        c5.info("Belum ada rating pada filter yang dipilih untuk membuat boxplot.")

    # Grafik hubungan: setiap titik adalah satu catatan pada data terpilih.
    data_scatter = df.dropna(subset=["Rating_Rata_rata"]).copy()
    if not data_scatter.empty:
        fig = px.scatter(
            data_scatter, x="Jumlah_Transaksi", y="Rating_Rata_rata",
            color="Outlet", color_discrete_map=WARNA_OUTLET,
            size="Jumlah_Ulasan", size_max=18, opacity=0.65,
            hover_data=["Tanggal", "Kanal_Penjualan", "Jumlah_Ulasan"],
            title="Hubungan jumlah transaksi dengan rating",
            labels={"Jumlah_Transaksi": "Jumlah transaksi", "Rating_Rata_rata": "Rating", "Jumlah_Ulasan": "Jumlah ulasan"},
        )
        fig.update_yaxes(range=[1, 5])
        fig.update_layout(
                font=dict(color="#FFFFFF"),             
                title=dict(font=dict(color="#FFFFFF")),    
                legend=dict(font=dict(color="#FFFFFF"))  
            )
        c6.plotly_chart(rapikan(fig), width="stretch")
        c6.caption("Grafik ini menunjukkan hubungan, bukan bukti bahwa banyak transaksi menyebabkan rating turun atau naik.")
    else:
        c6.info("Belum ada rating pada filter yang dipilih untuk membuat scatter plot.")

# ── Tab 2: Keramaian ──
with tab2:
    st.caption("Gunakan tab ini untuk mengatur jadwal tim dan mengantisipasi lonjakan pertanyaan/keluhan.")
    harian_outlet = (
        df.groupby(["Tanggal", "Hari", "Jenis_Hari", "Bulan_Singkat", "Outlet"], observed=True)
        ["Jumlah_Transaksi"].sum().reset_index()
    )
    c1, c2 = st.columns(2)
    pv = (
        harian_outlet.groupby(["Hari", "Outlet"], observed=True)["Jumlah_Transaksi"].mean()
        .unstack().reindex(HARI)
    )
    fig = px.imshow(
        pv, text_auto=".0f", aspect="auto", color_continuous_scale=SKALA_EMAS,
        title="Rata-rata transaksi per hari: hari × outlet",
    )
    fig.update_xaxes(title="")
    fig.update_layout(
        font=dict(color="#FFFFFF"),
        title=dict(font=dict(color="#FFFFFF")),
        legend=dict(font=dict(color="#FFFFFF"))
    )
    fig.update_yaxes(title="")
    fig.update_coloraxes(colorbar_title="Transaksi")
    c1.plotly_chart(rapikan(fig), width="stretch")

    g = harian_outlet.groupby(["Jenis_Hari", "Outlet"], observed=True)["Jumlah_Transaksi"].mean().reset_index()
    fig = px.bar(
        g, x="Jenis_Hari", y="Jumlah_Transaksi", color="Outlet", barmode="group",
        color_discrete_map=WARNA_OUTLET, 
        title="Rata-rata transaksi per hari menurut jenis hari",
        category_orders={"Jenis_Hari": ["Hari kerja", "Akhir pekan", "Libur nasional"]},
        labels={"Jumlah_Transaksi": "Transaksi / hari"},
    )
    fig.update_xaxes(title="")
    fig.update_layout(
        font=dict(color="#FFFFFF"),
        title=dict(font=dict(color="#FFFFFF")),
        legend=dict(font=dict(color="#FFFFFF"))
    )
    c2.plotly_chart(rapikan(fig), width="stretch")

    mode = st.radio("Tren harian", ["Total semua outlet", "Per outlet"], horizontal=True)
    if mode == "Total semua outlet":
        s = harian.reset_index()
        s["Rata-rata 7 hari"] = s["Jumlah_Transaksi"].rolling(7, min_periods=1).mean()
        fig = px.line(
            s, x="Tanggal", y=["Jumlah_Transaksi", "Rata-rata 7 hari"],
            title="Transaksi harian (garis tebal = rata-rata bergerak 7 hari)",
            color_discrete_sequence=["#6F86A8", "#F0C24B"],
        )
        fig.for_each_trace(lambda t: t.update(name="Harian" if t.name == "Jumlah_Transaksi" else t.name))
    else:
        s = harian_outlet.sort_values("Tanggal").copy()
        s["Transaksi"] = s.groupby("Outlet")["Jumlah_Transaksi"].transform(
            lambda x: x.rolling(7, min_periods=1).mean()
        )
        fig = px.line(
            s, x="Tanggal", y="Transaksi", color="Outlet", color_discrete_map=WARNA_OUTLET,
            title="Transaksi harian per outlet (rata-rata bergerak 7 hari)",
        )
    fig.update_yaxes(title="Transaksi")
    fig.update_layout(
        font=dict(color="#FFFFFF"),
        title=dict(font=dict(color="#FFFFFF")),
        legend=dict(font=dict(color="#FFFFFF"))
    )
    st.plotly_chart(rapikan(fig, 380), width="stretch")

    c3, c4 = st.columns(2)
    g = df.groupby(["Bulan_Singkat", "Kanal_Penjualan"], observed=True)["Jumlah_Transaksi"].sum().reset_index()
    fig = px.bar(
        g, x="Bulan_Singkat", y="Jumlah_Transaksi", color="Kanal_Penjualan",
        color_discrete_map=WARNA_KANAL, 
        title="Transaksi per bulan menurut kanal",
        category_orders={"Bulan_Singkat": BULAN}, labels={"Jumlah_Transaksi": "Transaksi"},
    )
    fig.update_xaxes(title="")
    fig.update_layout(
        font=dict(color="#FFFFFF"),
        title=dict(font=dict(color="#FFFFFF")),
        legend=dict(font=dict(color="#FFFFFF"))
    )
    c3.plotly_chart(rapikan(fig), width="stretch")

    top = (
        df.groupby(["Tanggal", "Hari", "Jenis_Hari"], observed=True)["Jumlah_Transaksi"].sum()
        .nlargest(10).reset_index()
    )
    top["Tanggal"] = top["Tanggal"].dt.strftime("%d %b %Y")
    top.columns = ["Tanggal", "Hari", "Jenis hari", "Total transaksi"]
    c4.markdown("**10 hari tersibuk** (semua outlet terpilih)")
    c4.dataframe(top, hide_index=True, width="stretch", height=370)

    # ── Tambahan: omzet per outlet sebagai konteks aktivitas ──
    st.markdown("### Omzet per outlet sebagai konteks aktivitas")
    om = ringkas(df, "Outlet")
    om["Omzet_Miliar"] = om["Omzet"] / 1e9
    om = om.sort_values("Omzet_Miliar", ascending=True)  # terbesar tampil di atas
    if om["Omzet"].sum() > 0:
        fig = px.bar(
            om, x="Omzet_Miliar", y="Outlet", orientation="h",
            color="Outlet", color_discrete_map=WARNA_OUTLET,
            text=[desimal(v) for v in om["Omzet_Miliar"]],
            title="Konteks beban interaksi: omzet bersih per outlet",
            labels={"Omzet_Miliar": "Omzet bersih (miliar Rp)", "Outlet": ""},
            category_orders={"Outlet": list(om["Outlet"])},
            hover_data={"Omzet_Miliar": ":.2f", "Transaksi": ":,.0f", "Outlet": False},
        )
        fig.update_traces(textposition="outside", cliponaxis=False)
        fig.update_xaxes(title="Omzet bersih (miliar Rp)")
        fig.update_layout(
            font=dict(color="#FFFFFF"),
            title=dict(font=dict(color="#FFFFFF")),
            legend=dict(font=dict(color="#FFFFFF"))
        )
        st.plotly_chart(rapikan(fig, 340, legenda=False), width="stretch")
        teratas = om.iloc[-1]
        terendah = om.iloc[0]
        if len(om) > 1:
            st.caption(
                f"{teratas['Outlet']} memiliki omzet bersih tertinggi ({desimal(teratas['Omzet_Miliar'])} miliar Rp) "
                f"dan {terendah['Outlet']} terendah ({desimal(terendah['Omzet_Miliar'])} miliar Rp). "
                "Gunakan sebagai konteks perencanaan koordinasi: omzet tidak sama dengan volume keluhan."
            )
    else:
        st.info("Tidak ada data omzet pada filter saat ini.")

# ── Tab 3: Profil pelanggan ──
with tab3:
    c1, c2 = st.columns(2)
    g = ringkas(df, "Outlet")
    gl = g.melt(id_vars="Outlet", value_vars=["Wisatawan", "Lokal"], var_name="Pelanggan", value_name="Jml")
    fig = px.bar(
        gl, x="Outlet", y="Jml", color="Pelanggan",
        color_discrete_map=WARNA_PELANGGAN, text_auto=".0f",
        title="Komposisi wisatawan vs lokal per outlet (%)",
        category_orders={"Outlet": list(g.sort_values("Pct_Wisatawan", ascending=False)["Outlet"])},
    )
    fig.update_layout(barnorm="percent")
    fig.update_yaxes(title="%")
    fig.update_layout(
        font=dict(color="#FFFFFF"),
        title=dict(font=dict(color="#FFFFFF")),
        legend=dict(font=dict(color="#FFFFFF"))
    )
    c1.plotly_chart(rapikan(fig), width="stretch")

    g = ringkas(df, ["Bulan_Singkat", "Outlet"])
    fig = px.line(
        g, x="Bulan_Singkat", y="Pct_Wisatawan", color="Outlet", markers=True,
        color_discrete_map=WARNA_OUTLET, 
        title="Persentase transaksi wisatawan per bulan",
        category_orders={"Bulan_Singkat": BULAN}, labels={"Pct_Wisatawan": "% wisatawan"},
    )
    fig.update_xaxes(title="")
    fig.update_layout(
        font=dict(color="#FFFFFF"),
        title=dict(font=dict(color="#FFFFFF")),
        legend=dict(font=dict(color="#FFFFFF"))
    )
    c2.plotly_chart(rapikan(fig), width="stretch")

    c3, c4 = st.columns(2)
    g = ringkas(df, ["Outlet", "Kanal_Penjualan"])
    fig = px.bar(
        g, x="Outlet", y="Transaksi", color="Kanal_Penjualan",
        color_discrete_map=WARNA_KANAL, text_auto=".0f",
        title="Bauran kanal per outlet (%)",
    )
    fig.update_layout(barnorm="percent")
    fig.update_yaxes(title="%")
    fig.update_layout(
        font=dict(color="#FFFFFF"),
        title=dict(font=dict(color="#FFFFFF")),
        legend=dict(font=dict(color="#FFFFFF"))
    )
    c3.plotly_chart(rapikan(fig), width="stretch")

    g = ringkas(df, ["Kanal_Penjualan", "Outlet"]).dropna(subset=["Nilai_per_Transaksi"])
    fig = px.bar(
        g, x="Kanal_Penjualan", y="Nilai_per_Transaksi", color="Outlet", barmode="group",
        color_discrete_map=WARNA_OUTLET, 
        title="Rata-rata nilai belanja per transaksi (Rp)",
        labels={"Nilai_per_Transaksi": "Rp / transaksi"},
    )
    fig.update_xaxes(title="")
    fig.update_layout(
        font=dict(color="#FFFFFF"),
        title=dict(font=dict(color="#FFFFFF")),
        legend=dict(font=dict(color="#FFFFFF"))
    )
    c4.plotly_chart(rapikan(fig), width="stretch")

    online = df[df["Kanal_Penjualan"] == "Pesan antar online"]
    if len(online):
        hari_promo = (online["Diskon_Rp"] > 0).sum()
        st.info(
            f"**Promo pesan antar online:** terdapat potongan harga pada {angka(hari_promo)} dari "
            f"{angka(len(online))} catatan harian online "
            f"({desimal(hari_promo / len(online) * 100, 0)}%), total diskon Rp {angka(online['Diskon_Rp'].sum())}. "
            "Berguna untuk mengantisipasi pertanyaan seputar promo aplikasi."
        )

# ── Tab 4: Tindak lanjut ──
with tab4:
    st.caption(
        "Daftar hari/outlet/kanal dengan rating di bawah ambang. "
        "Gunakan untuk menelusuri ulasan asli dan menghubungi tim outlet."
    )
    ambang = st.slider("Tampilkan catatan dengan rating di bawah", 3.6, 4.6, 4.0, 0.1)
    rendah = df[(df["Rating_Rata_rata"] < ambang) & (df["Jumlah_Ulasan"] > 0)].copy()
    a, b, c = st.columns(3)
    a.metric("Catatan rating rendah", angka(len(rendah)))
    a2 = df[df["Jumlah_Ulasan"] > 0]
    b.metric("Dari seluruh catatan berulasan", f"{desimal(len(rendah) / max(len(a2), 1) * 100, 1)}%")
    c.metric("Ulasan terdampak", angka(rendah["Jumlah_Ulasan"].sum()))

    if rendah.empty:
        st.success("Tidak ada catatan di bawah ambang ini pada filter saat ini. 🎉")
    else:
        c1, c2 = st.columns(2)
        g = rendah.groupby(["Outlet", "Kanal_Penjualan"]).size().reset_index(name="Jumlah")
        fig = px.bar(
            g, x="Outlet", y="Jumlah", color="Kanal_Penjualan", color_discrete_map=WARNA_KANAL,
            title="Catatan rating rendah per outlet & kanal",
        )
        fig.update_layout(
            font=dict(color="#FFFFFF"),
            title=dict(font=dict(color="#FFFFFF")),
            legend=dict(font=dict(color="#FFFFFF"))
        )
        c1.plotly_chart(rapikan(fig), width="stretch")
        g = rendah.groupby(["Bulan_Singkat", "Outlet"], observed=True).size().reset_index(name="Jumlah")
        fig = px.bar(
            g, x="Bulan_Singkat", y="Jumlah", color="Outlet", color_discrete_map=WARNA_OUTLET,
            title="Kapan rating rendah terjadi?", category_orders={"Bulan_Singkat": BULAN},
        )
        fig.update_xaxes(title="")
        fig.update_layout(
        font=dict(color="#FFFFFF"),
        title=dict(font=dict(color="#FFFFFF")),
        legend=dict(font=dict(color="#FFFFFF"))
    )
        c2.plotly_chart(rapikan(fig), width="stretch")

        tabel = rendah.sort_values(["Rating_Rata_rata", "Jumlah_Ulasan"], ascending=[True, False])[
            ["Tanggal", "Hari", "Jenis_Hari", "Outlet", "Kanal_Penjualan",
             "Rating_Rata_rata", "Jumlah_Ulasan", "Jumlah_Transaksi"]
        ].copy()
        tabel["Tanggal"] = tabel["Tanggal"].dt.strftime("%d %b %Y")
        tabel.columns = ["Tanggal", "Hari", "Jenis hari", "Outlet", "Kanal", "Rating", "Ulasan", "Transaksi"]
        st.dataframe(tabel, hide_index=True, width="stretch", height=380)
        st.download_button(
            "⬇️ Unduh daftar (CSV)", tabel.to_csv(index=False).encode("utf-8-sig"),
            file_name="rating_rendah.csv", mime="text/csv",
        )

# ── Tab 5: Data ──
with tab5:
    st.caption("Data setelah filter diterapkan.")
    tampil = df.drop(columns=["Skor_Ulasan", "Ulasan_Berating", "Bulan_Singkat", "Bulan_No"], errors="ignore")
    st.dataframe(tampil, hide_index=True, width="stretch", height=480)
    st.download_button(
        "⬇️ Unduh data terfilter (CSV)", tampil.to_csv(index=False).encode("utf-8-sig"),
        file_name="data_terfilter.csv", mime="text/csv",
    )

# ───────────────────── Tambahan Fase 2, 4, dan 5 ──────────────────────────
st.divider()
with st.expander("🧪 Fase 2 · Pemeriksaan kualitas data dan transformasi", expanded=False):
    st.markdown("Bagian ini menjalankan pemeriksaan pada data sumber yang berhasil dimuat. Hasil mengikuti file yang sedang digunakan, bukan angka contoh.")
    checks = []
    checks.append({"Pemeriksaan": "Jumlah baris", "Hasil": angka(len(data)), "Status": "Informasi"})
    checks.append({"Pemeriksaan": "Jumlah kolom", "Hasil": angka(len(data.columns)), "Status": "Informasi"})
    checks.append({"Pemeriksaan": "Baris duplikat penuh", "Hasil": angka(data.duplicated().sum()), "Status": "Perlu ditinjau" if data.duplicated().any() else "Tidak ditemukan"})
    missing = data.isna().sum()
    for col in data.columns:
        if int(missing[col]) > 0:
            checks.append({"Pemeriksaan": f"Nilai kosong: {col}", "Hasil": angka(int(missing[col])), "Status": "Perlu ditinjau"})

    if "Jumlah_Transaksi" in data.columns and {"Transaksi_Wisatawan", "Transaksi_Lokal"}.issubset(data.columns):
        mismatch_trx = (data["Jumlah_Transaksi"] != data["Transaksi_Wisatawan"] + data["Transaksi_Lokal"]).sum()
        checks.append({"Pemeriksaan": "Konsistensi transaksi", "Hasil": angka(int(mismatch_trx)), "Status": "Perlu ditinjau" if mismatch_trx else "Sesuai"})
    if {"Omzet_Bersih_Rp", "Omzet_Gudeg_Rp", "Omzet_Oleh_Oleh_Rp", "Diskon_Rp"}.issubset(data.columns):
        expected = data["Omzet_Gudeg_Rp"] + data["Omzet_Oleh_Oleh_Rp"] - data["Diskon_Rp"]
        mismatch_omzet = ~np.isclose(data["Omzet_Bersih_Rp"], expected, equal_nan=True)
        checks.append({"Pemeriksaan": "Konsistensi omzet bersih", "Hasil": angka(int(mismatch_omzet.sum())), "Status": "Perlu ditinjau" if mismatch_omzet.any() else "Sesuai"})
    if {"Rating_Rata_rata", "Jumlah_Ulasan"}.issubset(data.columns):
        rating_tanpa_ulasan = (data["Rating_Rata_rata"].notna() & (data["Jumlah_Ulasan"].fillna(0) == 0)).sum()
        ulasan_tanpa_rating = (data["Rating_Rata_rata"].isna() & (data["Jumlah_Ulasan"].fillna(0) > 0)).sum()
        checks.append({"Pemeriksaan": "Rating terisi tetapi tanpa ulasan", "Hasil": angka(int(rating_tanpa_ulasan)), "Status": "Perlu ditinjau" if rating_tanpa_ulasan else "Sesuai"})
        checks.append({"Pemeriksaan": "Ulasan ada tetapi rating kosong", "Hasil": angka(int(ulasan_tanpa_rating)), "Status": "Perlu ditinjau" if ulasan_tanpa_rating else "Sesuai"})
    if "Omzet_Bersih_Rp" in data.columns:
        checks.append({"Pemeriksaan": "Catatan omzet bersih nol", "Hasil": angka(int((data["Omzet_Bersih_Rp"] == 0).sum())), "Status": "Perlu ditinjau" if (data["Omzet_Bersih_Rp"] == 0).any() else "Tidak ditemukan"})
        omzet = data["Omzet_Bersih_Rp"].dropna()
        if len(omzet) >= 4:
            q1, q3 = omzet.quantile([0.25, 0.75])
            iqr = q3 - q1
            batas_bawah, batas_atas = q1 - 1.5 * iqr, q3 + 1.5 * iqr
            ekstrem = ((data["Omzet_Bersih_Rp"] < batas_bawah) | (data["Omzet_Bersih_Rp"] > batas_atas)).sum()
            checks.append({"Pemeriksaan": "Catatan omzet di luar batas IQR", "Hasil": angka(int(ekstrem)), "Status": "Tinjau konteks, jangan otomatis hapus"})

    hasil_cek = pd.DataFrame(checks)
    st.dataframe(hasil_cek, hide_index=True, width="stretch")
    st.download_button(
        "⬇️ Unduh hasil pemeriksaan kualitas data (CSV)",
        hasil_cek.to_csv(index=False).encode("utf-8-sig"),
        file_name="hasil_pemeriksaan_kualitas_data.csv", mime="text/csv",
        key="unduh_audit_fase2",
    )

    st.markdown("#### Variabel turunan untuk analisis layanan")
    turunan = df.copy()
    turunan["Nilai_per_Transaksi_Rp"] = turunan["Omzet_Bersih_Rp"] / turunan["Jumlah_Transaksi"].replace(0, np.nan)
    turunan["Persentase_Wisatawan"] = turunan["Transaksi_Wisatawan"] / turunan["Jumlah_Transaksi"].replace(0, np.nan) * 100
    turunan["Rata_bergerak_7_hari"] = turunan.sort_values("Tanggal").groupby("Outlet")["Jumlah_Transaksi"].transform(lambda s: s.rolling(7, min_periods=1).mean())
    st.caption("Nilai per transaksi, persentase wisatawan, dan rata-rata bergerak 7 hari dihitung pada data yang sedang difilter. Nilai kosong tidak diganti nol.")
    st.dataframe(turunan[["Tanggal", "Outlet", "Kanal_Penjualan", "Jumlah_Transaksi", "Nilai_per_Transaksi_Rp", "Persentase_Wisatawan", "Rata_bergerak_7_hari"]].head(100), hide_index=True, width="stretch")
    st.download_button(
        "⬇️ Unduh data dengan variabel turunan (CSV)",
        turunan.to_csv(index=False).encode("utf-8-sig"),
        file_name="data_transformasi_layanan.csv", mime="text/csv",
        key="unduh_transformasi_fase2",
    )

with st.expander("📖 Fase 5 · Cerita data dan rekomendasi", expanded=False):
    rating_outlet = ringkas(df, "Outlet").dropna(subset=["Rating"]).sort_values("Rating")
    rating_kanal = ringkas(df, "Kanal_Penjualan").dropna(subset=["Rating"]).sort_values("Rating")
    st.markdown("### Ringkasan eksekutif berdasarkan filter saat ini")
    st.write(
        f"Pada periode {awal:%d %b %Y} sampai {akhir:%d %b %Y}, data terpilih mencakup "
        f"{angka(len(df))} catatan dan {angka(trx)} transaksi, dengan {angka(ulasan)} ulasan "
        f"serta rating berbobot {desimal(r_sel) if not np.isnan(r_sel) else 'tidak tersedia'} dari 5. "
        "Dashboard ini ditujukan untuk membantu Tim Layanan Pelanggan menentukan area yang perlu ditinjau, bukan menetapkan penyebab keluhan tanpa bukti tambahan."
    )
    if not rating_outlet.empty:
        rendah_outlet = rating_outlet.iloc[0]
        tinggi_outlet = rating_outlet.iloc[-1]
        st.markdown(f"- **Prioritas outlet:** {rendah_outlet['Outlet']} memiliki rating berbobot terendah ({desimal(rendah_outlet['Rating'])}); {tinggi_outlet['Outlet']} tertinggi ({desimal(tinggi_outlet['Rating'])}).")
    if not rating_kanal.empty:
        rendah_kanal = rating_kanal.iloc[0]
        st.markdown(f"- **Prioritas kanal:** {rendah_kanal['Kanal_Penjualan']} memiliki rating berbobot terendah ({desimal(rendah_kanal['Rating'])}) pada filter saat ini.")
    if "Jenis_Hari" in df.columns:
        jenis_rating = ringkas(df, "Jenis_Hari").dropna(subset=["Rating"]).sort_values("Rating")
        if not jenis_rating.empty:
            j = jenis_rating.iloc[0]
            st.markdown(f"- **Periode untuk ditinjau:** {j['Jenis_Hari']} memiliki rating berbobot terendah ({desimal(j['Rating'])}). Periksa data operasional sebelum menyimpulkan penyebabnya.")
    st.markdown("#### Tindakan yang disarankan")
    st.markdown(
        "1. Telusuri ulasan asli untuk catatan rating rendah.\n"
        "2. Bandingkan pola rating dengan volume transaksi, outlet, kanal, dan jenis hari.\n"
        "3. Periksa catatan operasional seperti waktu tunggu atau ketepatan pesanan jika tersedia.\n"
        "4. Catat tindakan perbaikan dan pantau perubahan rating pada periode berikutnya."
    )
    laporan = [
        "LAPORAN RINGKAS DASHBOARD LAYANAN PELANGGAN - GUDEG SEKAR ARUM",
        f"Periode: {awal:%d %b %Y} - {akhir:%d %b %Y}",
        f"Jumlah catatan: {angka(len(df))}",
        f"Total transaksi: {angka(trx)}",
        f"Total ulasan: {angka(ulasan)}",
        f"Rating berbobot: {desimal(r_sel) if not np.isnan(r_sel) else 'tidak tersedia'} / 5",
        "",
        "Ringkasan:",
        f"Rating berbobot dihitung berdasarkan jumlah ulasan pada filter terpilih.",
        "Data merupakan simulasi. Perbedaan rating tidak membuktikan penyebab keluhan.",
        "",
        "Rekomendasi:",
        "1. Telusuri ulasan asli pada catatan rating rendah.",
        "2. Periksa kesiapan layanan pada periode ramai.",
        "3. Bandingkan pengalaman menurut outlet dan kanal.",
        "4. Dokumentasikan tindak lanjut dan evaluasi berkala.",
    ]
    st.download_button(
        "⬇️ Unduh ringkasan cerita data (TXT)",
        "\n".join(laporan).encode("utf-8"),
        file_name="ringkasan_cerita_data_layanan.txt", mime="text/plain",
        key="unduh_story_fase5",
    )

with st.expander("🧑‍🤝‍🧑 Fase 4 · Checklist uji keterbacaan dashboard", expanded=False):
    st.markdown("Gunakan checklist ini saat minimal dua orang di luar kelompok mencoba dashboard. Tandai hanya setelah pengujian benar-benar dilakukan.")
    st.checkbox("Penguji 1 berhasil menemukan outlet dengan rating terendah", key="uji1_outlet")
    st.checkbox("Penguji 1 berhasil menemukan catatan rating di bawah ambang", key="uji1_rendah")
    st.checkbox("Penguji 2 berhasil menggunakan filter outlet dan periode", key="uji2_filter")
    st.checkbox("Penguji 2 memahami arti rating berbobot dan batasan data simulasi", key="uji2_paham")
    catatan_uji = st.text_area("Catatan kendala dan perbaikan setelah pengujian", placeholder="Isi berdasarkan hasil uji nyata, misalnya label sulit dibaca atau filter kurang jelas.")
    st.download_button(
        "⬇️ Unduh catatan uji keterbacaan (TXT)",
        ("CHECKLIST UJI KETERBACAAN DASHBOARD\n"
         f"Penguji 1 - menemukan outlet terendah: {'Ya' if st.session_state.get('uji1_outlet') else 'Belum ditandai'}\n"
         f"Penguji 1 - menemukan rating rendah: {'Ya' if st.session_state.get('uji1_rendah') else 'Belum ditandai'}\n"
         f"Penguji 2 - menggunakan filter: {'Ya' if st.session_state.get('uji2_filter') else 'Belum ditandai'}\n"
         f"Penguji 2 - memahami rating berbobot: {'Ya' if st.session_state.get('uji2_paham') else 'Belum ditandai'}\n\nCatatan: {catatan_uji}\n").encode("utf-8"),
        file_name="catatan_uji_keterbacaan.txt", mime="text/plain",
        key="unduh_uji_fase4",
    )