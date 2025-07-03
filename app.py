import streamlit as st
import pandas as pd
import os

# Deteksi Kolom Berdasarkan Merk EVC
def detect_columns(evcm_name):
    evcm_name = evcm_name.lower()
    if "minielcor" in evcm_name or "elcor" in evcm_name:
        return {"flow": "Flow (m3/h)", "flow_min": "Min. Flow (m3/h)", "flow_max": "Max. Flow (m3/h)", "pressure": "Pressure (bar)"}
    elif "itron" in evcm_name:
        return {"flow": "Flow (m3/h)", "flow_min": "Flow min (m3/h)", "flow_max": "Flow max (m3/h)", "pressure": "Pressure (bar)"}
    else:
        raise Exception(f"Merk EVC '{evcm_name}' tidak dikenali")

# Proses Setiap Sheet
def process_sheet(sheet_name, sheet_df, month_name, uploaded_file):
    try:
        merk_evc = sheet_df.iloc[7, 1] if sheet_df.shape[0] >= 8 and sheet_df.shape[1] >= 2 else "Unknown"
        col_map = detect_columns(merk_evc)
        data_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=12)

        flow_col = col_map["flow"]
        flow_min_col = col_map["flow_min"]
        flow_max_col = col_map["flow_max"]

        if flow_col not in data_df.columns:
            raise Exception(f"Kolom '{flow_col}' tidak ditemukan")

        gsize = sheet_df.iloc[9, 1]
        id_ref = sheet_df.iloc[4, 1]
        nama_pelanggan = str(sheet_df.iloc[5, 0]).replace("Place Id:", "").strip()

        total_jam = len(data_df)
        over_150 = len(data_df[data_df[flow_col] >= 1.5 * data_df[flow_max_col]])
        over_120 = len(data_df[(data_df[flow_col] >= 1.2 * data_df[flow_max_col]) & (data_df[flow_col] < 1.5 * data_df[flow_max_col])])
        over_100 = len(data_df[(data_df[flow_col] >= 1.0 * data_df[flow_max_col]) & (data_df[flow_col] < 1.2 * data_df[flow_max_col])])
        under = len(data_df[data_df[flow_col] <= data_df[flow_min_col]])

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

        return {
            "Nama Pelanggan": nama_pelanggan,
            "ID Ref": id_ref,
            "Gsize": gsize,
            "Total Jam": total_jam,
            "Persen Flow >=150%": persen_150,
            "Persen Flow >=120%": persen_120,
            "Persen Flow >=100%": persen_100,
            "Persen Flow <=Qmin": persen_under,
            f"Kesimpulan Bulan {month_name}": kesimpulan,
            "Tekanan Outlet": data_df[col_map["pressure"]].mean() if col_map.get("pressure") in data_df.columns else None
        }

    except Exception as e:
        st.warning(f"Gagal memproses sheet {sheet_name}: {e}")
        return None

def main():
    st.title("Analisa Flow Meter Pelanggan")

    uploaded_file = st.file_uploader("Upload file Excel", type=["xls", "xlsx", "csv"])

    if uploaded_file:
        file_name = uploaded_file.name
        if "export" in file_name.lower():
            st.error("Silakan ganti nama file Excel sesuai nama bulan, contoh: Juli2025.xlsx")
            return

        month_name = os.path.splitext(file_name)[0]

        xls = pd.ExcelFile(uploaded_file)
        all_results = []

        for sheet_name in xls.sheet_names:
            sheet_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
            result = process_sheet(sheet_name, sheet_df, month_name, uploaded_file)
            if result:
                all_results.append(result)

        if all_results:
            result_df = pd.DataFrame(all_results)
            st.dataframe(result_df)

            # Download hasil
            csv = result_df.to_csv(index=False).encode('utf-8')
            excel_output = pd.ExcelWriter(f"Analisa_{month_name}.xlsx", engine='xlsxwriter')
            result_df.to_excel(excel_output, index=False)
            excel_output.close()

            st.download_button(
                label="Download Hasil CSV",
                data=csv,
                file_name=f"Analisa_{month_name}.csv",
                mime="text/csv"
            )

            with open(f"Analisa_{month_name}.xlsx", "rb") as f:
                st.download_button(
                    label="Download Hasil Excel",
                    data=f.read(),
                    file_name=f"Analisa_{month_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()
