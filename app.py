import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# Konfigurasi Qmin dan Qmax berdasarkan GSize
METER_CONFIG = {
    16: (0.5, 25),
    25: (0.8, 40),
    40: (8, 65),
    65: (10, 100),
    100: (16, 160),
    160: (13, 250),
    250: (20, 400),
    400: (32, 650),
    650: (50, 1000),
    1000: (80, 1600),
    1600: (125, 2500),
    2500: (200, 4000),
}

def process_xls(file, month_name):
    all_results = []
    xls = pd.ExcelFile(file)

    for sheet_name in xls.sheet_names:
        try:
            sheet_df = pd.read_excel(file, sheet_name=sheet_name, header=None, nrows=14, usecols="A:B")
            data_df = pd.read_excel(file, sheet_name=sheet_name, header=12)

            nama_pelanggan = str(sheet_df.iloc[5, 0]).replace("Place Id:", "").strip()
            id_ref = sheet_df.iloc[4, 1]
            gsize_raw = sheet_df.iloc[9, 1]
            gsize_numeric = int(str(gsize_raw).lower().replace("g", ""))

            qmin, qmax = METER_CONFIG.get(gsize_numeric, (None, None))

            flow_col = "Flow (m3/h)"
            total_jam = len(data_df)

            kondisi = {
                "Kondisi 1": len(data_df[data_df[flow_col] >= 1.5 * qmax]),
                "Kondisi 2": len(data_df[(data_df[flow_col] >= 1.2 * qmax) & (data_df[flow_col] < 1.5 * qmax)]),
                "Kondisi 3": len(data_df[(data_df[flow_col] >= 1.0 * qmax) & (data_df[flow_col] < 1.2 * qmax)]),
                "Kondisi 8": len(data_df[data_df[flow_col] <= qmin])
            }

            persen = {k: v / total_jam * 100 for k, v in kondisi.items()}

            status_kondisi = {
                "Status Kondisi 1": persen["Kondisi 1"] >= 1,
                "Status Kondisi 2": persen["Kondisi 2"] >= 10,
                "Status Kondisi 3": persen["Kondisi 3"] >= 15,
                "Status Kondisi 4": total_jam >= 50 and (persen["Kondisi 1"] >= 1 or persen["Kondisi 3"] >= 15),
                "Status Kondisi 5": total_jam >= 50 and (persen["Kondisi 2"] >= 10 or persen["Kondisi 3"] >= 15),
                "Status Kondisi 6": total_jam >= 50 and (persen["Kondisi 1"] >= 1 or persen["Kondisi 2"] >= 10),
                "Status Kondisi 7": total_jam >= 30 and (persen["Kondisi 1"] >= 1 and persen["Kondisi 2"] >= 10 and persen["Kondisi 3"] >= 15),
                "Status Kondisi 8": persen["Kondisi 8"] >= 10,
            }

            if persen["Kondisi 1"] > 1:
                kesimpulan = "Overrange"
            elif persen["Kondisi 8"] > 10:
                kesimpulan = "Underrange"
            else:
                kesimpulan = "Normal"

            all_results.append({
                "No": len(all_results) + 1,
                "ID Ref": id_ref,
                "Nama Pelanggan": nama_pelanggan,
                "GSize Meter Terpasang": gsize_numeric,
                "Qmin Meter Terpasang": qmin,
                "Qmax Meter Terpasang": qmax,
                **kondisi,
                "Jumlah Jam Operasi": total_jam,
                **{f"Persentase {k}": round(v, 2) for k, v in persen.items()},
                **status_kondisi,
                "Kesimpulan Bulan Ini": kesimpulan,
                "Tekanan Outlet": "-",
                "Diameter Spool": "-",
                "Kesimpulan Bulan Lalu": "-",
                "Kesimpulan Bulan Lalunya Lagi": "-",
                "Status Meter": "-",
                "Tipe Penyesuaian": "-",
                "Nilai Penyesuaian": "-",
                "Keterangan": "-"
            })

        except Exception as e:
            st.warning(f"Gagal memproses sheet {sheet_name}: {e}")

    return pd.DataFrame(all_results)

def convert_to_xlsx(df):
    output = BytesIO()

    wb = Workbook()
    ws = wb.active
    ws.title = "Rekapitulasi AKM"

    headers = list(df.columns)
    ws.append(headers)

    for _, row in df.iterrows():
        ws.append(list(row))

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
