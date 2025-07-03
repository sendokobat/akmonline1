import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

def process_xls(file, month_name):
    # (kode process_xls tetap sama)
    # Disingkat untuk kejelasan.
    pass

def convert_to_xlsx(df):
    output = BytesIO()

    # Buat workbook baru
    wb = Workbook()
    ws = wb.active
    ws.title = "Rekapitulasi AKM"

    # Header
    headers = [
        "No", "ID Ref", "Nama Pelanggan", "GSize Meter Terpasang", "Qmin Meter Terpasang", "Qmax Meter Terpasang",
        "Flowmax 150% >= Qmax (Jam)", "Flowmax 120% >= Qmax (Jam)", "Flowmax 100% >= Qmax (Jam)", "Flowmin <= Qmin (Jam)",
        "Jumlah Jam Operasi", "Persentase Flowmax 150% >= Qmax", "Persentase Flowmax 120% >= Qmax",
        "Persentase Flowmax 100% >= Qmax", "Persentase Flowmin <= Qmin",
        "Kesimpulan Bulan Ini", "Tekanan Outlet", "Diameter Spool", "Kesimpulan Bulan Lalu", "Kesimpulan Bulan Lalunya Lagi",
        "Status Meter", "Tipe Penyesuaian", "Nilai Penyesuaian", "Keterangan"
    ]

    ws.append(headers)

    # Data rows
    for _, row in df.iterrows():
        ws.append(list(row))

    # Format header
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    wb.save(output)
    output.seek(0)
    return output

def main():
    st.title("Analisa Flow Meter (Upload File)")

    uploaded_file = st.file_uploader("Upload file XLS/XLSX", type=["xls", "xlsx"])
    if not uploaded_file:
        return

    month_name = uploaded_file.name.split(".")[0]

    with st.spinner("Memproses data..."):
        result_df = process_xls(uploaded_file, month_name)

    st.success("Analisa selesai!")
    st.dataframe(result_df)

    xlsx_file = convert_to_xlsx(result_df)
    st.download_button("Download Hasil XLSX", xlsx_file, file_name=f"Rekapitulasi_AKM_{month_name}.xlsx")

if __name__ == "__main__":
    main()
