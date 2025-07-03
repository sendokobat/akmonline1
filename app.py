import streamlit as st
import pandas as pd
from io import BytesIO
import tempfile

# === CONFIG ===
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
            data_df = pd.read_excel(file, sheet_name=sheet_name, header=12, usecols=[14, 15, 16])  # Hanya ambil kolom flow

            nama_pelanggan = str(sheet_df.iloc[5, 0]).replace("Place Id:", "").strip()
            id_ref = sheet_df.iloc[4, 1]
            merk_evc = sheet_df.iloc[7, 1]
            gsize_raw = sheet_df.iloc[9, 1]
            gsize_numeric = int(str(gsize_raw).lower().replace("g", ""))

            qmin, qmax = METER_CONFIG.get(gsize_numeric, (None, None))

            flow_col = "Flow (m3/h)"
            data_df.columns = [flow_col, "Min. Flow (m3/h)", "Max. Flow (m3/h)"]

            total_jam = len(data_df)
            over_150 = len(data_df[data_df[flow_col] >= 1.5 * qmax])
            over_120 = len(data_df[(data_df[flow_col] >= 1.2 * qmax) & (data_df[flow_col] < 1.5 * qmax)])
            over_100 = len(data_df[(data_df[flow_col] >= 1.0 * qmax) & (data_df[flow_col] < 1.2 * qmax)])
            under = len(data_df[data_df[flow_col] <= qmin])

            persen_150 = over_150 / total_jam * 100
            persen_120 = over_120 / total_jam * 100
            persen_100 = over_100 / total_jam * 100
            persen_under = under / total_jam * 100

            if persen_150 > 1:
                kesimpulan = "Overrange"
            elif persen_under > 10:
                kesimpulan = "Underrange"
            else:
                kesimpulan = "Normal"

            all_results.append({
                "No": len(all_results) + 1,
                "ID Ref": id_ref,
                "Nama Pelanggan": nama_pelanggan,
                "GSize": gsize_numeric,
                "Qmin Meter": qmin,
                "Qmax Meter": qmax,
                "Flowmax 150% >= Qmax": over_150,
                "Flowmax 120% >= Qmax": over_120,
                "Flowmax 100% >= Qmax": over_100,
                "Flowmin <= Qmin": under,
                "Jumlah Jam Operasi": total_jam,
                "Persentase Flowmax 150% >= Qmax": round(persen_150, 2),
                "Persentase Flowmax 120% >= Qmax": round(persen_120, 2),
                "Persentase Flowmax 100% >= Qmax": round(persen_100, 2),
                "Persentase Flowmin <= Qmin": round(persen_under, 2),
                f"Kesimpulan Bulan {month_name}": kesimpulan
            })

        except Exception as e:
            st.warning(f"Gagal memproses sheet {sheet_name}: {e}")

    return pd.DataFrame(all_results)

def convert_to_xlsx(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Hasil Analisa')
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
    st.download_button("Download Hasil XLSX", xlsx_file, file_name=f"Analisa_{month_name}.xlsx")

if __name__ == "__main__":
    main()
