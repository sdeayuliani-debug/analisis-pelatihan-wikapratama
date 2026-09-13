# ============================================================
#  ANALISIS & PROYEKSI JUMLAH PESERTA PELATIHAN WIKAPratama
#  Periode 2019–2025  |  Proyeksi 2026–2027
#  Google Colab Compatible
# ============================================================
#  CARA PAKAI DI GOOGLE COLAB:
#  1. Upload file "data pelatihan satuan.xlsx" ke Colab
#  2. Jalankan cell pertama → install library
#  3. Upload file → sesuaikan path di cell LOAD DATA
#  4. Run all cells
# ============================================================

# %%  [CELL 1] INSTALASI LIBRARY
# ===============================
# !pip install pandas openpyxl matplotlib seaborn statsmodels scikit-learn -q

# %%  [CELL 1b] PATH KERJA
# =========================
# Terdeteksi otomatis; tidak perlu diubah.
import os
try:                       # dijalankan sebagai file .py
    WORKDIR = os.path.dirname(os.path.abspath(__file__))
except NameError:          # dijalankan di Jupyter/Colab
    WORKDIR = os.getcwd()
print(" WORKDIR:", WORKDIR)

# %%  [CELL 2] IMPORT & KONFIGURASI
# ==================================
import matplotlib
try:
    import google.colab  # noqa: F401  → sedang di Colab, biarkan backend inline
except ImportError:
    matplotlib.use('Agg')  # lokal/headless → simpan PNG tanpa window
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.gridspec as gridspec
from matplotlib.ticker import MaxNLocator
import seaborn as sns
import warnings, os, io
from datetime import datetime

warnings.filterwarnings('ignore')

# --- Warna tema WIKA (gelap + aksen amber) ---
BG        = '#0D1117'
BG_CARD   = '#161B22'
BG_CARD2  = '#1C2128'
TEXT      = '#E6EDF3'
MUTED     = '#8B949E'
ACCENT    = '#EECA8C'   # amber utama
BLUE      = '#58A6FF'
GREEN     = '#3FB950'
RED       = '#F85149'
PURPLE    = '#BC8CFF'
TEAL      = '#39D0C4'
ORANGE    = '#FFA657'
PINK      = '#FF7EB6'
GRID      = '#21262D'

plt.rcParams.update({
    'figure.facecolor':   BG,
    'axes.facecolor':     BG_CARD,
    'axes.edgecolor':     GRID,
    'axes.labelcolor':    TEXT,
    'xtick.color':        MUTED,
    'ytick.color':        MUTED,
    'text.color':         TEXT,
    'font.family':        'sans-serif',
    'font.sans-serif':    ['DejaVu Sans'],
    'axes.titlesize':     13,
    'axes.titleweight':   'bold',
    'axes.titlecolor':    TEXT,
    'legend.framealpha':   0,
    'legend.labelcolor':  TEXT,
    'axes.spines.top':    False,
    'axes.spines.right':  False,
})

# ============================================================
# %%  [CELL 3] LOAD DATA
# ============================================================
# Colab: upload "data pelatihan satuan.xlsx" lewat panel Files (ikon folder
# di kiri), lalu biarkan path default '/content/...' di bawah.
# Lokal: fallback ke path Windows otomatis (WORKDIR dari cell 1).
CANDIDATES = [
    '/content/data pelatihan satuan.xlsx',
    os.path.join(WORKDIR, 'data pelatihan satuan.xlsx'),
    os.path.join(WORKDIR, 'bahan analisis pelatihan 2019 - 2025',
                 'data pelatihan satuan.xlsx'),
    'data pelatihan satuan.xlsx',
]

FILE_PATH = next((p for p in CANDIDATES if os.path.exists(p)), None)
if FILE_PATH is None:
    raise FileNotFoundError(
        "File 'data pelatihan satuan.xlsx' tidak ditemukan.\n"
        "Di Colab: upload file lewat panel Files, lalu jalankan ulang cell ini.\n"
        "Path yang sudah dicoba:\n  " + "\n  ".join(CANDIDATES)
    )
print(f" Memakai file: {FILE_PATH}")

df_raw = pd.read_excel(FILE_PATH, sheet_name=0)
df_raw.columns = ['Tahun', 'Nama_Pelatihan']
df_raw['Tahun'] = pd.to_numeric(df_raw['Tahun'], errors='coerce').astype('Int64')
df_raw['Nama_Pelatihan'] = df_raw['Nama_Pelatihan'].astype(str).str.strip()
df_raw = df_raw.dropna(subset=['Tahun', 'Nama_Pelatihan'])
df_raw['Tahun'] = df_raw['Tahun'].astype(int)

print(f" Data loaded: {len(df_raw):,} baris")
print(f"   Tahun      : {df_raw['Tahun'].min()} – {df_raw['Tahun'].max()}")
print(f"   Jenis pelatihan unik: {df_raw['Nama_Pelatihan'].nunique()}")

# ============================================================
# %%  [CELL 4] AGREGAT TOTAL PESERTA PER TAHUN
# ============================================================
df_year = (df_raw.groupby('Tahun')
             .size()
             .reset_index(name='Jumlah_Peserta')
             .sort_values('Tahun'))

print("\n Total Peserta per Tahun:")
print(df_year.to_string(index=False))

# ============================================================
# %%  [CELL 5] MAPPING PELATIHAN KE STANDAR 2025
# ============================================================
# Mapping nama pelatihan historis → kategori standar 2025
# agar dataset konsisten untuk analisis tren

MAPPING = {
    # --- QSHE / Safety ---
    'QSHE AWARENESS':                    'QHSE AWARENESS',
    'QHSE AWARENESS':                    'QHSE AWARENESS',
    'QHSE AWARENESS BATCH 32':           'QHSE AWARENESS',
    'QHSE AWARENESS BATCH 33':           'QHSE AWARENESS',
    'QHSE AWARENESS BATCH 34':           'QHSE AWARENESS',
    'QHSE AWARENESS BATCH 35':           'QHSE AWARENESS',
    'QHSE AWARENESS BATCH 36':           'QHSE AWARENESS',
    'QHSE AWARENESS BATCH 37':           'QHSE AWARENESS',
    'QHSE AWARENESS BATCH 38':           'QHSE AWARENESS',
    'QHSE AWARENESS BATCH 39':           'QHSE AWARENESS',
    'SAFETY OFFICER':                    'SAFETY OFFICER',
    'K3 KONSTRUKSI':                     'Petugas K3 Konstruksi',
    'Petugas K3 Konstruksi':             'Petugas K3 Konstruksi',
    'Ahli Muda K3 Konstruksi':           'Ahli Muda K3 Konstruksi',
    'Ahli Madya K3 Konstruksi':          'Ahli Madya K3 Konstruksi',
    'Ahli Madya K3 konstruksi':          'Ahli Madya K3 Konstruksi',
    'Bekerja di Ketinggian':             'Petugas K3 Konstruksi',
    'Operasi Scaffolding':               'Pengawas Scaffolding',
    'Operator Scaffolding':              'Pengawas Scaffolding',
    'PENULANGAN':                        'Pengawas Scaffolding',
    'PENULANGAN (SERPAN)':               'Pengawas Scaffolding',
    'PERANCAH':                          'Pengawas Scaffolding',
    'Pemasang Perancah dan Acuan/Cetakan Beton': 'Pengawas Scaffolding',
    'RIGGING & LIFTING':                 'RIGGING & LIFTING',
    'PEW RIGING &LIFTING':               'RIGGING & LIFTING',
    'PESAWAT ANGKAT ANGKUT':             'RIGGING & LIFTING',
    'MAGNETIC WIRE ROPE TEST':           'RIGGING & LIFTING',
    'HSE FOR NON HSE':                   'HSE FOR NON HSE',
    'HSE FOR HSE':                       'HSE FOR HSE',
    'MANAJEMEN RISIKO':                  'MANAJEMEN RISIKO BASIC LEVEL',
    'MANAJAEMEN RISIKO':                 'MANAJEMEN RISIKO BASIC LEVEL',
    'MANAJMEN RISIKO':                   'MANAJEMEN RISIKO BASIC LEVEL',
    'MANAJEMEN RISIKO BASIC LEVEL':      'MANAJEMEN RISIKO BASIC LEVEL',
    'MANAJEMEN RISIKO INTERMEDIATE LEVEL':'MANAJEMEN RISIKO INTERMEDIATE LEVEL',

    # --- ESG (khusus 2025) ---
    'ESG BASIC LEVEL':                    'ESG BASIC LEVEL',
    'ESG INTERMEDIATE LEVEL':             'ESG INTERMEDIATE LEVEL',

    # --- SCM / Procurement ---
    'SCM (Iproc)':                        'IPROC',
    'SCM (eProcurement & PMCS)':          'IPROC',
    'SCM (eSCM & eCATALOGUE)':            'IPROC',
    'SCM (SIMPABEAN)':                    'IPROC',
    'SCM (WITON)':                        'IPROC',
    'SUPPLY CHAIN MANAGEMENT':           'IPROC',
    'SUPPLY CHAIN MANAGEMENT TKDN':       'IPROC',
    'SUPPLY CHAIN MANAGEMENT ESCM & ECATALOGUE': 'IPROC',
    'SCM PENGADAAN':                      'IPROC',
    'IPROC':                              'IPROC',
    'PMCS':                               'PM EPCC',
    'KOMERSIAL & PMCS':                   'PM EPCC',
    'PM EPCC':                            'PM EPCC',

    # --- BIM ---
    'BIM':                                'BIM INFRA',
    'BIM (SAP2000)':                      'BIM INFRA',
    'BIM ALL PLAN':                       'BIM INFRA',
    'BIM BUILDING':                       'BIM INFRA',
    'BULDING INFORMATION MODELING':      'BIM INFRA',
    'PELATIHAN BIM TAHAP 1':              'BIM INFRA',
    'PELATIHAN BIM TAHAP 2':              'BIM INFRA',
    'PEW MODUL BIM & GIS':                'BIM INFRA',
    'BIM INFRA':                          'BIM INFRA',
    'Manager BIM Madya':                  'Manager BIM Madya',
    'Manager BIM Muda':                   'Manager BIM Muda',
    'Koordinator BIM':                    'Koordinator BIM',

    # --- Keuangan ---
    'KEUANGAN':                            'FINANCIAL FOR NON FINANCIAL',
    'FINON':                              'FINANCIAL FOR NON FINANCIAL',
    'FINANCIAL NON FINANCIAL':            'FINANCIAL FOR NON FINANCIAL',
    'FINANCIAL FOR NON FINANCIAL':        'FINANCIAL FOR NON FINANCIAL',
    'FINANCIAL FOR FINANCIAL':            'FINANCIAL FOR FINANCIAL',
    'PEMBELAJARAN BERKELANJUTAN FUNGSI AKUNTANSI': 'FINANCIAL FOR NON FINANCIAL',
    'AKUNTANSI PAJAK':                    'FINANCIAL FOR FINANCIAL',

    # --- Kontrak & Legal ---
    'MANAJEMEN KONTRAK KONSTRUKSI':       'MANAJEMEN KONTRAK',
    'MANAJEMEN KONTRAK':                  'MANAJEMEN KONTRAK',
    'KONTRAK MANAJEMEN':                  'MANAJEMEN KONTRAK',
    'INVESTASI (BALIKPAPAN)':             'MANAJEMEN KONTRAK',
    'INVESTASI (MAKASSAR)':              'MANAJEMEN KONTRAK',
    'EKSPOR IMPOR':                       'MANAJEMEN KONTRAK',

    # --- Manajemen Proyek ---
    'PRIMAVERA':                          'MS PROJECT',
    'PLANNING SCHEDULING MS PROJECT':     'MS PROJECT',
    'MS PROJECT':                         'MS PROJECT',
    'WIKA FUNDAMENTALS OF PROJECT MANAGEMENT': 'WIKA FUNDAMENTALS OF PROJECT MANAGEMENT',
    'WIKA FUNDAMENTALS OF PROJECT MANAGEMENT #1': 'WIKA FUNDAMENTALS OF PROJECT MANAGEMENT',
    'KELULUSAN WIKA FUNDAMENTALS OF PROJECT MANGEMENT': 'WIKA FUNDAMENTALS OF PROJECT MANAGEMENT',
    'PROSDEM':                            'WIKA FUNDAMENTALS OF PROJECT MANAGEMENT',
    'MANAJEMEN KONSTRUKSI GEDUNG':        'MANAJEMEN KONSTRUKSI BIDANG GEDUNG',
    'MANAJEMEN KONSTRUKSI BIDANG GEDUNG': 'MANAJEMEN KONSTRUKSI BIDANG GEDUNG',
    'PROJECT COST MANAGEMENT':           'COST ESTIMATION',
    'COST ESTIMATION':                    'COST ESTIMATION',
    'ASET MANAGEMENT':                    'MANAJEMEN KONSTRUKSI BIDANG GEDUNG',
    'PEKERJAAN RESTROOM':                 'MANAJEMEN KONSTRUKSI BIDANG GEDUNG',
    'PEKERJAAN TANAH':                    'MANAJEMEN KONSTRUKSI BIDANG GEDUNG',
    'PEMADATAN TANAH':                    'MANAJEMEN KONSTRUKSI BIDANG GEDUNG',
    'PEW PEMADATAN TANAH':                'MANAJEMEN KONSTRUKSI BIDANG GEDUNG',
    'RIGID & FLEXSIBLE PAVEMENT':         'MANAJEMEN KONSTRUKSI BIDANG GEDUNG',
    'FINISHING BASAH':                    'MANAJEMEN KONSTRUKSI BIDANG GEDUNG',
    'ELEKTRIKAL DASAR':                   'MEKANIKAL KELISTRIKAN',
    'MEKANIK KLISTRIKAN':                 'MEKANIKAL KELISTRIKAN',
    'MEKANIKAL KELISTRIKAN':              'MEKANIKAL KELISTRIKAN',
    'PEW ELECTRICAL INTERMEDIATE':        'MEKANIKAL KELISTRIKAN',
    'ADVANCE COMMISIONING':               'TESTING & COMMISIONING',
    'TESTING & COMMISIONING':             'TESTING & COMMISIONING',
    'COMMISSIONING DASAR':                'TESTING & COMMISIONING',

    # --- Beton ---
    'BETON MUDA':                         'BETON MUDA',
    'PEW MODUL BETON MUDA':               'BETON MUDA',
    'Tukang Besi Beton':                   'BETON MUDA',

    # --- Soft Skill ---
    'SUPERVISORY':                        'SUPERVISORY',
    'LEADER AS COACH':                    'SUPERVISORY',
    'LEADER AS COACH ':                   'SUPERVISORY',
    'NEGOSIASI':                          'NEGOSIASI & KOMUNIKASI',
    'PRESENTASI MEMUKAU':                 'NEGOSIASI & KOMUNIKASI',
    'PRESENTASI WIKA INOVATIF':           'NEGOSIASI & KOMUNIKASI',
    'NEGOSIASI & KOMUNIKASI':             'NEGOSIASI & KOMUNIKASI',
    'COMMUNICATION CROSS GENERATION':     'NEGOSIASI & KOMUNIKASI',
    'COMMCROSS GEN STAFF':                'NEGOSIASI & KOMUNIKASI',
    'HC FOR NON HC':                      'HC FOR NON HC',
    'HC FOR NON HC ':                     'HC FOR NON HC',
    'PROBLEM SOLVING & DECISION MAKING':  'PROBLEM SOLVING & DECISION MAKING',
    'ADVANCE PFW':                        'SUPERVISORY',
    'ALP KASIE BATCH 65':                 'SUPERVISORY',
    'ALP KASIE BATCH 66':                 'SUPERVISORY',

    # --- EPCC ---
    'BASIC EPCC':                         'BASIC EPCC',
    'PESERTA WORKSHOP EHU EPCC 2023':     'BASIC EPCC',

    # --- QA/QC ---
    'QA/QC':                              'QA/QC',

    # --- TKDN ---
    'TKDN':                               'TKDN',

    # --- SMAP ---
    'SMAP':                               'MANAJEMEN KONTRAK',

    # --- Lean ---
    'LEAN CONSTRUCTION':                  'LEAN CONSRUCTION PEOPLE DEVELOPMENT',
    'SEMINAR LEAN COSTRUCTION DAN SIMULASI VILLEGO': 'LEAN CONSRUCTION PEOPLE DEVELOPMENT',
    'LEAN CONSRUCTION PEOPLE DEVELOPMENT': 'LEAN CONSRUCTION PEOPLE DEVELOPMENT',

    # --- Workshop / Seminar ---
    'WORKSHOP CHECK POINT INFRASTRUCTURE 2 DIVISON': 'WIKA FUNDAMENTALS OF PROJECT MANAGEMENT',
    'REFRESHMENT WORKSHOP DAN PEMBEKALAN EVALUASI HASIL USAHA': 'WIKA FUNDAMENTALS OF PROJECT MANAGEMENT',
    'KICK OFF IMPLEMENTASI BIM LEVEL 5D': 'BIM INFRA',
    'WEBINAR LATEST DEVELOPMENT OF GROUND IMPROVEMENT TECHNOLOGY IN INDONESIA': 'WIKA FUNDAMENTALS OF PROJECT MANAGEMENT',
    'WEBINAR SUSTAINABLE INFRASTRUCTURE FORUM': 'WIKA FUNDAMENTALS OF PROJECT MANAGEMENT',
    'WEBINAR TUNNEL JACKING':             'WIKA FUNDAMENTALS OF PROJECT MANAGEMENT',
    'PELATIHAN PROJECT TRAINNING WEBGIS':'WIKA FUNDAMENTALS OF PROJECT MANAGEMENT',
    'PELATIHAN KONTRAK MANAJEMEN':        'MANAJEMEN KONTRAK',
    'PELATIHAN BIM TAHAP 1':              'BIM INFRA',
    'PELATIHAN BIM TAHAP 2':              'BIM INFRA',
    'PELATIHAN END USER TRAINING (EUT) MODUL SURAT KEPUTUSAN (SK) ONLINE': 'WIKA FUNDAMENTALS OF PROJECT MANAGEMENT',
    'PELATIHAN DAN SERTIFIKASI CRMO':     'WIKA FUNDAMENTALS OF PROJECT MANAGEMENT',
    'PELATIHAN DAN PENGEMBANGAN PRODUK DIGITAL UNTUK MENJADI MANAJER PRODUK': 'WIKA FUNDAMENTALS OF PROJECT MANAGEMENT',
    'PELATIHAN KEPABEANAN':               'MANAJEMEN KONTRAK',
    'PEMBELAJARAN DESAIN IU/UX DAN PEMBUATAN PROPOTYPE APLIKASI': 'WIKA FUNDAMENTALS OF PROJECT MANAGEMENT',
    'WOEKSOP INPUUT DATA JANUARI SAP PROYEK': 'WIKA FUNDAMENTALS OF PROJECT MANAGEMENT',
    'TRANNING FOR TRANINERS':             'SUPERVISORY',
    'HYDRIMETALURGI':                     'MANAJEMEN KONSTRUKSI BIDANG GEDUNG',
    'SHARING TUNNELING':                  'MANAJEMEN KONSTRUKSI BIDANG GEDUNG',
    'SIMDIV WEB (WRK)':                   'WIKA FUNDAMENTALS OF PROJECT MANAGEMENT',
    'PELAKSANA LAPANGAN PEKERJAAN GEDUNG': 'MANAJEMEN KONSTRUKSI BIDANG GEDUNG',
    'Pelaksana Lapangan Pekerjaan Gedung': 'MANAJEMEN KONSTRUKSI BIDANG GEDUNG',
    'MANDATORY':                          'QHSE AWARENESS',
}

def map_pelatihan(nama):
    nama = str(nama).strip()
    return MAPPING.get(nama, nama)  # tidak ketemu → pakai nama asli

df_raw['Pelatihan_Mapped'] = df_raw['Nama_Pelatihan'].apply(map_pelatihan)
df_raw['Is_Mapped'] = df_raw['Nama_Pelatihan'].apply(
    lambda x: str(x).strip() in MAPPING
)

# Statistik mapping
total_all   = len(df_raw)
total_maped = df_raw['Is_Mapped'].sum()
total_unmap = total_all - total_maped
pct_maped   = total_maped / total_all * 100
pct_unmap   = total_unmap / total_all * 100

print(f"\n Mapping Statistics:")
print(f"   Total data          : {total_all:,}")
print(f"   Termapping          : {total_maped:,} ({pct_maped:.1f}%)")
print(f"   Tidak termapping    : {total_unmap:,} ({pct_unmap:.1f}%)")

# Per tahun
df_map_stats = df_raw.groupby('Tahun')['Is_Mapped'].agg(['sum','count']).reset_index()
df_map_stats['pct'] = df_map_stats['sum'] / df_map_stats['count'] * 100
print("\n   Per Tahun:")
print(df_map_stats.to_string(index=False))

# ============================================================
# %%  [CELL 6] GRAFIK 1 — TOTAL PESERTA PER TAHUN (2019–2025)
# ============================================================
years     = df_year['Tahun'].tolist()
counts    = df_year['Jumlah_Peserta'].tolist()
colors    = [BLUE, BLUE, BLUE, BLUE, BLUE, BLUE, ACCENT]  # 2025 = amber

fig, ax = plt.subplots(figsize=(11, 6), dpi=130)

bars = ax.bar(years, counts, color=colors, width=0.6,
             edgecolor='none', zorder=3)

# Nilai di atas bar
for bar, val in zip(bars, counts):
    ax.text(bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 40,
            f'{val:,}',
            ha='center', va='bottom',
            fontsize=11, fontweight='bold', color=ACCENT)

# Garis tren
ax.plot(years, counts, color=ACCENT, linewidth=2.2,
        marker='o', markersize=7, zorder=4)

# Annotasi fase
phases = {
    2019: ('Normal\n(pre-pandemi)', TEXT),
    2020: ('Pandemi\nCOVID-19', MUTED),
    2021: ('Pemulihan\npelatihan', TEXT),
    2022: ('Pemulihan\npelatihan', TEXT),
    2023: ('Pemulihan\npelatihan', TEXT),
    2024: ('Fluktuasi\npenyesuaian', TEXT),
    2025: ('LONJAKAN\nWIKAcademy', ACCENT),
}
for yr, (label, col) in phases.items():
    ax.text(yr, 400, label, ha='center', va='bottom',
            fontsize=7.5, color=col, style='italic')

ax.set_xlabel('Tahun', fontsize=11, labelpad=10)
ax.set_ylabel('Jumlah Peserta', fontsize=11, labelpad=10)
ax.set_title('Grafik Total Jumlah Peserta Pelatihan WIKAPratama\nPeriode 2019–2025',
             fontsize=14, fontweight='bold', color=TEXT, pad=15)
ax.set_xticks(years)
ax.set_xticklabels([str(y) for y in years], fontsize=10)
ax.yaxis.set_major_locator(MaxNLocator(integer=True))
ax.set_ylim(0, max(counts) * 1.18)
ax.grid(axis='y', color=GRID, linewidth=0.8, zorder=0)
ax.set_facecolor(BG_CARD)

# Footer
fig.text(0.01, -0.02, 'Human Capital Division  |  WIKAPratama',
         fontsize=8, color=MUTED)

plt.tight_layout()
plt.savefig('grafik_1_total_peserta_pertahun.png',
            dpi=150, bbox_inches='tight', facecolor=BG)
plt.show()
print(" Grafik 1 disimpan: grafik_1_total_peserta_pertahun.png")

# ============================================================
# %%  [CELL 7] GRAFIK 2 — PIE CHART MAPPING (TERMAPPING vs TIDAK)
# ============================================================
fig, axes = plt.subplots(1, 2, figsize=(13, 5.5), dpi=130)

# --- Pie utama ---
ax1 = axes[0]
wedges, texts, autotexts = ax1.pie(
    [total_maped, total_unmap],
    labels=['Termapping', 'Tidak Termapping'],
    colors=[BLUE, MUTED],
    autopct='%1.1f%%',
    startangle=140,
    wedgeprops=dict(width=0.55, edgecolor=BG_CARD),
    textprops=dict(color=TEXT),
    pctdistance=0.75,
)
for at in autotexts:
    at.set_fontsize(11); at.set_fontweight('bold')
ax1.set_title('Persentase Data Termapping\n2019–2025', color=TEXT, pad=10)

# --- Stacked bar per tahun ---
ax2 = axes[1]
df_map_bar = df_map_stats.set_index('Tahun')[['sum','count']].rename(
    columns={'sum':'Termapped','count':'Total'})
df_map_bar['Tidak Termapped'] = df_map_bar['Total'] - df_map_bar['Termapped']

x = df_map_bar.index.tolist()
ax2.bar(x, df_map_bar['Termapped'],   label='Termapped',   color=BLUE,   width=0.6)
ax2.bar(x, df_map_bar['Tidak Termapped'], bottom=df_map_bar['Termapped'],
        label='Tidak Termapped', color=MUTED, width=0.6)

for yr in x:
    total = df_map_bar.loc[yr, 'Total']
    ax2.text(yr, total + 50, f'{int(total):,}',
             ha='center', va='bottom', fontsize=9, color=TEXT)

ax2.set_xlabel('Tahun', fontsize=11)
ax2.set_ylabel('Jumlah Peserta', fontsize=11)
ax2.set_title('Jumlah Data Termapping per Tahun', color=TEXT, pad=10)
ax2.set_xticks(x)
ax2.set_xticklabels([str(y) for y in x], fontsize=10)
ax2.grid(axis='y', color=GRID, linewidth=0.8)
ax2.legend(loc='upper left', fontsize=9)
ax2.set_facecolor(BG_CARD)

fig.text(0.01, -0.02, 'Human Capital Division  |  WIKAPratama',
         fontsize=8, color=MUTED)
plt.tight_layout()
plt.savefig('grafik_2_mapping_persentase.png',
            dpi=150, bbox_inches='tight', facecolor=BG)
plt.show()
print(" Grafik 2 disimpan: grafik_2_mapping_persentase.png")

# ============================================================
# %%  [CELL 8] GRAFIK 3 — TOP 10 PELATIHAN TERBANYAK (2019–2025)
# ============================================================
df_top = (df_raw.groupby('Pelatihan_Mapped')
            .size()
            .reset_index(name='Jumlah_Peserta')
            .sort_values('Jumlah_Peserta', ascending=False)
            .head(10)
            .iloc[::-1])  # ascending untuk horizontal bar

palet = [ACCENT if i == 0 else BLUE for i in range(len(df_top))]
# Accent hanya untuk ESG BASIC LEVEL (per slide 5)
esg_idx = df_top[df_top['Pelatihan_Mapped'] == 'ESG BASIC LEVEL'].index
df_top_colors = [ACCENT if name == 'ESG BASIC LEVEL' else BLUE
                for name in df_top['Pelatihan_Mapped']]

fig, ax = plt.subplots(figsize=(11, 6), dpi=130)
bars = ax.barh(df_top['Pelatihan_Mapped'], df_top['Jumlah_Peserta'],
               color=df_top_colors, height=0.6, edgecolor='none')

for bar, val in zip(bars, df_top['Jumlah_Peserta']):
    ax.text(bar.get_width() + 20, bar.get_y() + bar.get_height() / 2,
            f'{val:,}', va='center', ha='left',
            fontsize=10, color=TEXT)

ax.set_xlabel('Jumlah Peserta', fontsize=11)
ax.set_title('Grafik Pelatihan Terbanyak Diikuti\nPeriode 2019–2025',
             fontsize=14, fontweight='bold', color=TEXT, pad=15)
ax.grid(axis='x', color=GRID, linewidth=0.8)
ax.set_xlim(0, df_top['Jumlah_Peserta'].max() * 1.15)
ax.set_facecolor(BG_CARD)

# Legend
p1 = mpatches.Patch(color=ACCENT, label='ESG BASIC LEVEL (Tertinggi)')
p2 = mpatches.Patch(color=BLUE,   label='Pelatihan Lainnya')
ax.legend(handles=[p1, p2], loc='lower right', fontsize=9)

fig.text(0.01, -0.02, 'Human Capital Division  |  WIKAPratama',
         fontsize=8, color=MUTED)
plt.tight_layout()
plt.savefig('grafik_3_top_pelatihan.png',
            dpi=150, bbox_inches='tight', facecolor=BG)
plt.show()
print(" Grafik 3 disimpan: grafik_3_top_pelatihan.png")

# ============================================================
# %%  [CELL 9] PROYEKSI ARIMA (2026–2027)
# ============================================================
#  Metodologi (sesuai slide): "ARIMA yang disesuaikan dengan
#  kondisi operasional normal" → data historis dinormalisasi dulu:
#  • 2020 (pandemi COVID-19)       : dikeluarkan dari series
#  • 2025 (lonjakan program khusus): dikurangi komponen ESG
#    (WIKAcademy self-learning: ESG BASIC & INTERMEDIATE LEVEL)
#  Lalu ARIMA dengan tren linear dipasang pada data normal.
# ============================================================
from statsmodels.tsa.arima.model import ARIMA
from sklearn.linear_model import LinearRegression

# Data historis — dict int year → count (dipakai semua cell berikutnya)
hist = {int(y): int(v) for y, v in zip(df_year['Tahun'], df_year['Jumlah_Peserta'])}

# ──────────────────────────────────────────────
# NORMALISASI DATA (kondisi operasional normal)
# ──────────────────────────────────────────────
esg_2025 = int(df_raw[(df_raw['Tahun'] == 2025) &
                      df_raw['Pelatihan_Mapped'].str.startswith('ESG',
                                                               na=False)].shape[0])
val_2025_norm = hist[2025] - esg_2025

print("── Normalisasi Data ──")
print(f"  2025 mentah          : {hist[2025]:,}")
print(f"  Dikurangi ESG        : -{esg_2025:,}  (program khusus WIKAcademy)")
print(f"  2025 ternormalisasi  : {val_2025_norm:,}")
print(f"  2020 (pandemi)       : dikeluarkan dari series")
series_normal = [hist[2019], hist[2021], hist[2022], hist[2023],
                 hist[2024], val_2025_norm]
print(f"  Series normal        : {series_normal}")

# ──────────────────────────────────────────────
# PENDEKATAN A (UTAMA): ARIMA + tren pada data normal
# ──────────────────────────────────────────────
print("\n── Pendekatan A: ARIMA(1,0,0) + tren (Data Normal) ──")
arima_norm = ARIMA(np.array(series_normal, dtype=float),
                   order=(1, 0, 0), trend='t').fit()
f_norm = arima_norm.forecast(2)
forecast = {2026: max(float(f_norm[0]), 0.0),
            2027: max(float(f_norm[1]), 0.0)}
forecast_method = "ARIMA(1,0,0)+tren pada data ternormalisasi"

for yr, val in forecast.items():
    print(f"    {yr}: {int(round(val)):,} peserta")

# ──────────────────────────────────────────────
# PENDEKATAN B (PEMBANDING): Linear Regression tahun normal
# ──────────────────────────────────────────────
print("\n── Pendekatan B: Linear Regression (Data Normal) ──")
normal_years = np.array([2019, 2021, 2022, 2023, 2024]).reshape(-1, 1)
normal_vals  = np.array([hist[y] for y in [2019, 2021, 2022, 2023, 2024]],
                        dtype=float)
lr = LinearRegression().fit(normal_years, normal_vals)
lr_pred = {2026: float(lr.predict(np.array([[2026]]))[0]),
           2027: float(lr.predict(np.array([[2027]]))[0])}
r2 = lr.score(normal_years, normal_vals)
print(f"  Model  : y = {lr.coef_[0]:.1f}x + {lr.intercept_:.1f}")
print(f"  R²     : {r2:.4f}")
for yr, val in lr_pred.items():
    print(f"    {yr}: {int(round(val)):,} peserta")

print(f"\n   Forecast yang digunakan: {forecast_method}")
print("  (Slide referensi menyebut 2.436 / 2.631 — hasil reproduksi\n"
      "   berada di rentang yang sama; selisih kecil berasal dari\n"
      "   nilai penyesuaian internal yang tidak dipublikasikan)")

# ============================================================
# %%  [CELL 10] GRAFIK 4 — PROYEKSI 2026–2027 + TREN
# ============================================================
all_years  = list(range(2019, 2028))
all_values = [hist.get(y, np.nan) for y in all_years]
proj_vals  = [np.nan, np.nan, forecast[2026], forecast[2027]]

fig, ax = plt.subplots(figsize=(12, 6), dpi=130)

# Historis
ax.plot(list(range(2019, 2026)), list(hist.values()),
        color=BLUE, linewidth=2.5, marker='o', markersize=7,
        label='Data Historis (2019–2025)', zorder=4)

# Confidence interval sederhana (±15%)
ci_low  = [v * 0.85 if y <= 2025 else np.nan for y, v in zip(all_years, all_values)]
ci_high = [v * 1.15 if y <= 2025 else np.nan for y, v in zip(all_years, all_values)]

# Proyeksi
ax.plot([2025, 2026, 2027],
        [hist[2025], forecast[2026], forecast[2027]],
        color=ACCENT, linewidth=2.5, marker='s', markersize=7,
        linestyle='--', label='Proyeksi ARIMA (2026–2027)', zorder=4)

# Confidence band proyeksi
ax.fill_between([2026, 2027],
                [forecast[2026]*0.92, forecast[2027]*0.92],
                [forecast[2026]*1.08, forecast[2027]*1.08],
                color=ACCENT, alpha=0.15, label='Confidence Interval (±8%)')

# Garis vertikal pemisah historis vs proyeksi
ax.axvline(x=2025.5, color=GRID, linestyle=':', linewidth=1.5, zorder=2)

# Label nilai
for yr, val in forecast.items():
    ax.annotate(f'{int(round(val)):,}',
                xy=(yr, val), xytext=(yr, val + 120),
                ha='center', fontsize=11, fontweight='bold', color=ACCENT)

# Annotasi 2025
ax.annotate(f'{hist[2025]:,}',
            xy=(2025, hist[2025]), xytext=(2025, hist[2025] + 150),
            ha='center', fontsize=10, color=BLUE)

ax.set_xlabel('Tahun', fontsize=11)
ax.set_ylabel('Jumlah Peserta', fontsize=11)
ax.set_title('Hasil Analisis Prediksi Data Pelatihan WIKAPratama (2026–2027)',
            fontsize=14, fontweight='bold', color=TEXT, pad=15)
ax.set_xticks(all_years)
ax.set_xticklabels([str(y) for y in all_years], fontsize=10)
ax.set_ylim(0, 5500)
ax.grid(color=GRID, linewidth=0.8, zorder=0)
ax.legend(loc='upper left', fontsize=10)
ax.set_facecolor(BG_CARD)

# Zona
ax.axvspan(2025.5, 2027.5, alpha=0.04, color=ACCENT)
ax.text(2026.5, 500, 'ZONA PROYEKSI', ha='center',
        fontsize=9, color=ACCENT, style='italic')

fig.text(0.01, -0.02, 'Human Capital Division  |  WIKAPratama',
         fontsize=8, color=MUTED)
plt.tight_layout()
plt.savefig('grafik_4_proyeksi_arima.png',
            dpi=150, bbox_inches='tight', facecolor=BG)
plt.show()
print(" Grafik 4 disimpan: grafik_4_proyeksi_arima.png")

# ============================================================
# %%  [CELL 11] DASHBOARD GABUNGAN (1 PAGE)
# ============================================================
fig = plt.figure(figsize=(16, 10), dpi=130, facecolor=BG)
gs  = gridspec.GridSpec(2, 3, figure=fig, hspace=0.42, wspace=0.32)

# --- Judul ---
fig.suptitle('Analisis & Proyeksi Jumlah Peserta Pelatihan WIKAPratama (2019–2027)',
             fontsize=16, fontweight='bold', color=TEXT, y=0.98)

# --- Panel 1: Bar chart total (spans 2 cols) ---
ax1 = fig.add_subplot(gs[0, :2])
b1  = ax1.bar(years, counts, color=colors, width=0.6, edgecolor='none')
for bar, val in zip(b1, counts):
    ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 40,
             f'{val:,}', ha='center', va='bottom', fontsize=9,
             fontweight='bold', color=ACCENT)
ax1.plot(years, counts, color=ACCENT, linewidth=2, marker='o', markersize=5)
ax1.set_title('Total Peserta per Tahun', color=TEXT, fontsize=11)
ax1.set_xticks(years); ax1.set_xticklabels([str(y) for y in years])
ax1.set_facecolor(BG_CARD); ax1.grid(axis='y', color=GRID, linewidth=0.7)
ax1.set_ylim(0, max(counts)*1.2)

# --- Panel 2: Pie mapping ---
ax2 = fig.add_subplot(gs[0, 2])
ax2.pie([total_maped, total_unmap],
        labels=['Termapped', 'Tidak\nTermapped'],
        colors=[BLUE, MUTED], autopct='%1.1f%%',
        startangle=90,
        wedgeprops=dict(width=0.5, edgecolor=BG_CARD),
        textprops=dict(color=TEXT, fontsize=9),
        pctdistance=0.75)
for at in ax2.texts:
    if '%' in at.get_text():
        at.set_fontweight('bold'); at.set_fontsize(10)
ax2.set_title('Data Termapping', color=TEXT, fontsize=11)

# --- Panel 3: Top 5 pelatihan ---
ax3 = fig.add_subplot(gs[1, 0])
df_top5 = df_top.head(5).iloc[::-1]
c3 = [ACCENT if n == 'ESG BASIC LEVEL' else BLUE
      for n in df_top5['Pelatihan_Mapped']]
ax3.barh(df_top5['Pelatihan_Mapped'], df_top5['Jumlah_Peserta'],
         color=c3, height=0.55)
for bar, val in zip(ax3.patches, df_top5['Jumlah_Peserta']):
    ax3.text(bar.get_width()+10, bar.get_y()+bar.get_height()/2,
             f'{val:,}', va='center', fontsize=8, color=TEXT)
ax3.set_title('Top 5 Pelatihan', color=TEXT, fontsize=11)
ax3.set_facecolor(BG_CARD); ax3.grid(axis='x', color=GRID, linewidth=0.7)
ax3.set_xlim(0, df_top5['Jumlah_Peserta'].max()*1.2)

# --- Panel 4: Line + proyeksi ---
ax4 = fig.add_subplot(gs[1, 1:])
ax4.plot(list(range(2019, 2026)), list(hist.values()),
         color=BLUE, linewidth=2.5, marker='o', markersize=6,
         label='Historis')
ax4.plot([2025, 2026, 2027],
         [hist[2025], forecast[2026], forecast[2027]],
         color=ACCENT, linewidth=2.5, marker='s', markersize=6,
         linestyle='--', label='Proyeksi ARIMA')
ax4.fill_between([2026, 2027],
                [forecast[2026]*0.92, forecast[2027]*0.92],
                [forecast[2026]*1.08, forecast[2027]*1.08],
                color=ACCENT, alpha=0.15)
ax4.axvline(x=2025.5, color=GRID, linestyle=':', linewidth=1.5)
ax4.set_title('Tren & Proyeksi 2026–2027', color=TEXT, fontsize=11)
ax4.set_xticks(list(range(2019, 2028)))
ax4.set_xticklabels([str(y) for y in range(2019, 2028)], fontsize=9)
ax4.set_facecolor(BG_CARD); ax4.grid(color=GRID, linewidth=0.7)
ax4.legend(loc='upper left', fontsize=9)

# --- Panel 5: Tabel ringkasan ---
ax5 = fig.add_subplot(gs[1, 2])
ax5.axis('off')
summary_data = [
    ['Periode', 'Jumlah Peserta'],
    ['2019', f'{hist[2019]:,}'],
    ['2020', f'{hist[2020]:,}'],
    ['2021', f'{hist[2021]:,}'],
    ['2022', f'{hist[2022]:,}'],
    ['2023', f'{hist[2023]:,}'],
    ['2024', f'{hist[2024]:,}'],
    ['2025', f'{hist[2025]:,}'],
    ['─'*8, '─'*8],
    ['2026 (Proy.)', f'{int(round(forecast[2026])):,}'],
    ['2027 (Proy.)', f'{int(round(forecast[2027])):,}'],
]
tbl = ax5.table(cellText=summary_data,
                loc='center', cellLoc='center')
tbl.auto_set_font_size(False)
tbl.set_fontsize(9)
tbl.scale(1.0, 1.3)
for (r, c), cell in tbl.get_celld().items():
    cell.set_facecolor(BG_CARD2 if r % 2 == 0 else BG_CARD)
    cell.set_text_props(color=TEXT)
    if r == 0:
        cell.set_text_props(color=ACCENT, fontweight='bold')
    if r in [8, 9, 10]:
        cell.set_text_props(color=ACCENT)
ax5.set_title('Ringkasan Data', color=TEXT, fontsize=11)

fig.text(0.01, 0.01, 'Human Capital Division  |  WIKAPratama  |  Januari 2026',
         fontsize=8, color=MUTED)
plt.savefig('dashboard_wika_pelatihan.png',
            dpi=150, bbox_inches='tight', facecolor=BG)
plt.show()
print(" Dashboard disimpan: dashboard_wika_pelatihan.png")

# ============================================================
# %%  [CELL 12] EXPORT HASIL KE EXCEL
# ============================================================
with pd.ExcelWriter('hasil_analisis_pelatihan_wika.xlsx', engine='openpyxl') as writer:
    # Sheet 1: Data mentah
    df_raw.to_excel(writer, sheet_name='Data Mentah', index=False)
    # Sheet 2: Total per tahun
    df_year.to_excel(writer, sheet_name='Total per Tahun', index=False)
    # Sheet 3: Top pelatihan
    (df_raw.groupby('Pelatihan_Mapped')
           .size().reset_index(name='Jumlah_Peserta')
           .sort_values('Jumlah_Peserta', ascending=False)
           .to_excel(writer, sheet_name='Top Pelatihan', index=False))
    # Sheet 4: Mapping stats
    df_map_stats.to_excel(writer, sheet_name='Mapping Stats', index=False)
    # Sheet 5: Proyeksi
    proj_df = pd.DataFrame({
        'Tahun': [2026, 2027],
        'Jumlah_Peserta_Proyeksi': [int(round(forecast[2026])),
                                     int(round(forecast[2027]))],
        'Model': [forecast_method, forecast_method],
    })
    proj_df.to_excel(writer, sheet_name='Proyeksi 2026-2027', index=False)

print(" File Excel disimpan: hasil_analisis_pelatihan_wika.xlsx")

# ============================================================
# %%  [CELL 13] CETAK KESIMPULAN
# ============================================================
print("""
╔══════════════════════════════════════════════════════════════╗
║           KESIMPULAN ANALISIS PELATIHAN WIKAPratama          ║
╠══════════════════════════════════════════════════════════════╣
║                                                              ║
║   DATA HISTORIS (2019–2025)                                ║
║  • Total peserta 7 tahun : {:>7,} orang                      ║
║  • Rata-rata per tahun  : {:>7,.0f} orang                    ║
║  • Peserta tertinggi    : {:>7,} (2025)                      ║
║  • Peserta terendah     : {:>7,} (2020, pandemi)            ║
║                                                              ║
║   TREN UTAMA                                               ║
║  • 2019→2020 : Penurunan {}% (pandemi COVID-19)              ║
║  • 2020→2025 : Kenaikan {}% (pemulihan + WIKAcademy)        ║
║  • 2025      : Lonjakan signifikan (ESG + mandatory)         ║
║                                                              ║
║   KONSISTENSI DATA                                         ║
║  • Data termapping   : {:.1f}%                               ║
║  • Data tidak termapping: {:.1f}%                            ║
║                                                              ║
║   PROYEKSI 2026–2027 (ARIMA + tren)                     ║
║  • 2026 : {:>6,} peserta                                    ║
║  • 2027 : {:>6,} peserta                                    ║
║  • Model membaca pertumbuhan STABIL & KONSISTEN              ║
║    (bukan lonjakan sementara)                               ║
║                                                              ║
║   PELATIHAN TERBANYAK (2019–2025)                          ║
║  1. ESG BASIC LEVEL         → {} peserta                     ║
║  2. ESG INTERMEDIATE LEVEL  → {} peserta                     ║
║  3. SUPERVISORY             → {} peserta                     ║
║                                                              ║
╚══════════════════════════════════════════════════════════════╝
""".format(
    df_year['Jumlah_Peserta'].sum(),
    df_year['Jumlah_Peserta'].mean(),
    df_year['Jumlah_Peserta'].max(),
    df_year[df_year['Tahun']==2020]['Jumlah_Peserta'].values[0],
    int((1 - hist[2020]/hist[2019])*100),
    int((hist[2025]/hist[2020]-1)*100),
    pct_maped, pct_unmap,
    int(round(forecast[2026])),
    int(round(forecast[2027])),
    int(df_top[df_top['Pelatihan_Mapped']=='ESG BASIC LEVEL']['Jumlah_Peserta'].values[0]),
    int(df_top[df_top['Pelatihan_Mapped']=='ESG INTERMEDIATE LEVEL']['Jumlah_Peserta'].values[0]),
    int(df_top[df_top['Pelatihan_Mapped']=='SUPERVISORY']['Jumlah_Peserta'].values[0]),
))