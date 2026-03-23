# Deploy Gratis (Streamlit Community Cloud)

## 1. Push project ke GitHub

Project ini minimal harus berisi:

- `streamlit_app.py`
- `requirements.txt`
- `Scripts/flatten_kpi.py`

## 2. Login ke Streamlit Community Cloud

- Buka: `https://share.streamlit.io`
- Login dengan akun GitHub

## 3. Deploy app

- Klik **New app**
- Pilih repository GitHub Anda
- Branch: `main` (atau branch Anda)
- Main file path: `streamlit_app.py`
- Klik **Deploy**

## 4. Pakai aplikasinya

- Upload file `.xlsx`
- Isi `Week Label`
- Klik **Run Flatten**
- Download hasil

## Catatan

- Opsi **Simpan juga ke folder Exports lokal server** sebaiknya dimatikan saat deploy cloud.
- Jika butuh update logic flatten, cukup update `Scripts/flatten_kpi.py` lalu redeploy/restart app.
