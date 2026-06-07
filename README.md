# Aplikasi Perhitungan Keekonomian Lapangan Migas

Aplikasi ini dibuat dengan **Python + Streamlit** untuk menghitung tabel keekonomian migas dan menyediakan tombol download hasil dalam format **XLSX**.

## Fitur

- Input jangka waktu proyek
- Input investasi capital dan non-capital
- Input data produksi tahun 1 sampai 7
- Input decline produksi
- Input harga minyak rata-rata
- Input Opex awal dan eskalasi Opex
- Pilihan metode depresiasi:
  - Straight Line
  - Double Declining Balance
  - Sum of Years Digits
- Input pajak
- Output langsung berupa tabel di layar
- Download hasil sebagai file `.xlsx`

## Cara menjalankan

1. Masuk ke folder project:

```bash
cd migas_project
```

2. Install dependency:

```bash
pip install -r requirements.txt
```

3. Jalankan aplikasi:

```bash
streamlit run app/main.py
```

## Struktur Folder

```text
migas_project/
├── app/
│   └── main.py
├── core/
│   ├── calculator.py
│   └── depreciation.py
├── utils/
│   └── excel_exporter.py
├── assets/
├── requirements.txt
└── README.md
```

## Catatan Satuan

- Produksi menggunakan satuan **Mbbl**.
- Harga minyak menggunakan **USD/bbl**.
- Income dihitung dalam **$M** dengan rumus:

```text
Income = Produksi (Mbbl) x Harga Minyak ($/bbl)
```

Karena 1 Mbbl = 1.000 bbl dan $M = 1.000 USD, satuannya konsisten menjadi $M.
