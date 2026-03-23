# KPI Data Flattener

App untuk mengubah file KPI mingguan Excel menjadi format flatten.

## Fitur

- Proses file KPI weekly (`.xlsx`) memakai logic di `Scripts/flatten_kpi.py`
- UI sederhana via Streamlit (`streamlit_app.py`)
- Output file Excel flatten siap download

## Struktur Utama

- `Scripts/flatten_kpi.py`: logic utama transformasi
- `streamlit_app.py`: UI web untuk upload, proses, dan download hasil
- `requirements.txt`: dependency Python

## Jalankan Lokal

1. Install dependency:

```bash
pip install -r requirements.txt
```

2. Run Streamlit:

```bash
streamlit run streamlit_app.py
```

3. Buka URL lokal Streamlit di browser, upload file KPI mingguan, isi `Week Label`, lalu klik **Run Flatten**.

## Deploy Gratis (Streamlit Community Cloud)

1. Push repo ini ke GitHub.
2. Buka `https://share.streamlit.io` dan login dengan GitHub.
3. Klik **New app**.
4. Pilih repository ini, branch `main`, dan main file `streamlit_app.py`.
5. Klik **Deploy**.

## Catatan

- Opsi simpan ke folder `Exports` cocok untuk running lokal.
- Saat deploy cloud, gunakan download button untuk mengambil hasil file.
