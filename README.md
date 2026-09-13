# Analisis Pelatihan WIKAPratama 2019–2027

Proyek ini berisi analisis jumlah peserta pelatihan WIKAPratama periode 2019–2025 dan proyeksi jumlah peserta tahun 2026–2027.

## Isi proyek

- `Analisis_Pelatihan_WIKAPratama_Colab.ipynb` — notebook Google Colab self-contained, data sudah tertanam di notebook.
- `wika_pelatihan_colab.py` — versi script Python dari notebook.
- `grafik_1_total_peserta_pertahun.png` — grafik total peserta per tahun.
- `grafik_2_mapping_persentase.png` — grafik persentase mapping nama pelatihan.
- `grafik_3_top_pelatihan.png` — grafik top pelatihan.
- `grafik_4_proyeksi_arima.png` — grafik proyeksi 2026–2027.
- `dashboard_wika_pelatihan.png` — dashboard ringkasan hasil analisis.
- `hasil_analisis_pelatihan_wika.xlsx` — output hasil analisis.

## Ringkasan hasil

Dataset berisi 14.705 baris peserta pelatihan periode 2019–2025.

| Tahun | Jumlah Peserta |
|---|---:|
| 2019 | 1.878 |
| 2020 | 904 |
| 2021 | 1.611 |
| 2022 | 1.622 |
| 2023 | 2.510 |
| 2024 | 1.645 |
| 2025 | 4.535 |

Hasil mapping nama pelatihan:

- 14.124 data termapping.
- 581 data tidak termapping.
- Tingkat mapping: 96,0%.

Proyeksi peserta:

| Tahun | Proyeksi Peserta |
|---|---:|
| 2026 | 2.384 |
| 2027 | 2.667 |

Metode proyeksi menggunakan ARIMA pada data tahunan yang dinormalisasi, dengan pengecualian tahun 2020 sebagai tahun pandemi dan penyesuaian lonjakan ESG tahun 2025.

## Cara menjalankan di Google Colab

1. Buka notebook `Analisis_Pelatihan_WIKAPratama_Colab.ipynb`.
2. Pilih `Runtime` → `Run all`.
3. Notebook tidak perlu upload Excel manual karena data sudah tertanam di dalam notebook.

## Catatan data

File Excel sumber tidak disertakan di repository untuk menghindari publikasi data internal mentah. Notebook memakai data yang sudah dikemas agar analisis tetap bisa direproduksi.
