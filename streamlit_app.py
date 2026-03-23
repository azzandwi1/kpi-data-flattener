from __future__ import annotations

from datetime import date
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory

import streamlit as st


APP_DIR = Path(__file__).resolve().parent
SCRIPT_PATH = APP_DIR / "Scripts" / "flatten_kpi.py"

if not SCRIPT_PATH.exists():
    st.error(f"File not found: {SCRIPT_PATH}")
    st.stop()

# Import flatten_pivot directly from local script.
import importlib.util

spec = importlib.util.spec_from_file_location("flatten_kpi", SCRIPT_PATH)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)
flatten_pivot = mod.flatten_pivot


def default_week_label(today: date) -> str:
    month_short = today.strftime("%b")
    week_no = ((today.day - 1) // 7) + 1
    return f"Week {week_no} {month_short}"


st.set_page_config(page_title="KPI Flattener", layout="wide")
st.title("KPI Flattener")
st.caption("Upload file KPI mingguan, lalu generate hasil flatten ke format siap pakai.")

uploaded_file = st.file_uploader(
    "Input Excel",
    type=["xlsx"],
    help="Contoh: W04_03_2026_KPI_WEEKLY.xlsx",
)

col1, col2 = st.columns(2)
with col1:
    week_label = st.text_input("Week Label", value=default_week_label(date.today()))
with col2:
    output_name = st.text_input(
        "Output Filename",
        value="Combined KPI Weekly.xlsx",
        help="Nama file hasil download",
    )

save_to_exports = st.checkbox(
    "Simpan juga ke folder Exports lokal server",
    value=False,
    help="Opsional. Berguna kalau app dijalankan di mesin lokal Anda.",
)

run_btn = st.button("Run Flatten", type="primary")

if run_btn:
    if uploaded_file is None:
        st.warning("Upload file Excel terlebih dulu.")
    else:
        with st.spinner("Processing..."):
            try:
                with TemporaryDirectory() as tmp_dir:
                    tmp_dir_path = Path(tmp_dir)
                    input_path = tmp_dir_path / uploaded_file.name
                    input_path.write_bytes(uploaded_file.getbuffer())

                    output_path = tmp_dir_path / output_name
                    flatten_pivot(
                        input_path=str(input_path),
                        output_path=str(output_path),
                        week_label=week_label.strip() or default_week_label(date.today()),
                    )

                    output_bytes = output_path.read_bytes()

                    if save_to_exports:
                        exports_dir = APP_DIR / "Exports"
                        exports_dir.mkdir(parents=True, exist_ok=True)
                        local_output_path = exports_dir / output_name
                        local_output_path.write_bytes(output_bytes)
                        st.success(f"Hasil juga disimpan ke: {local_output_path}")

                    st.success("Flatten selesai.")
                    st.download_button(
                        label="Download Hasil",
                        data=BytesIO(output_bytes),
                        file_name=output_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
            except Exception as exc:
                st.error(f"Gagal menjalankan flatten: {exc}")
