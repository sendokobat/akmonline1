import streamlit as st
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from io import BytesIO
import time
import os

# === CONFIG ===
AVE_URL = "https://ave.pgncom.co.id/website/account/index.rails"
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

# === SCRAPER ===
def login_and_download(username, password):
    options = Options()
    options.add_argument('--headless')
    driver = webdriver.Chrome(options=options)

    driver.get(AVE_URL)

    # Login
    driver.find_element(By.NAME, "name").send_keys(username)
    driver.find_element(By.NAME, "password").send_keys(password)
    driver.find_element(By.XPATH, "//input[@type='submit']").click()

    # Wait for main page to load
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ext-gen92")))

    # Example: scrape table or trigger download here (sesuaikan dengan kebutuhanmu)

    # Close driver
    driver.quit()

    # Simulasi hasil download file
    return "downloaded_file.xls"

# === ANALYSIS ===
def process_xls(file_path, month_name):
    all_results = []
    xls = pd.ExcelFile(file_path)

    for sheet_name in xls.sheet_names:
        df_header = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=14, usecols="A:B")
        merk_evc = df_header.iloc[7, 1]
        gsize_raw = df_header.iloc[9, 1]
        gsize_numeric = int(str(gsize_raw).lower().replace("g", ""))
        qmin, qmax = METER_CONFIG.get(gsize_numeric, (None, None))

        df_data = pd.read_excel(file_path, sheet_name=sheet_name, header=12)

        flow_col = "Flow (m3/h)"

        total_jam = len(df_data)
        over_150 = len(df_data[df_data[flow_col] >= 1.5 * qmax])
        over_120 = len(df_data[(df_data[flow_col] >= 1.2 * qmax) & (df_data[flow_col] < 1.5 * qmax)])
        over_100 = len(df_data[(df_data[flow_col] >= 1.0 * qmax) & (df_data[flow_col] < 1.2 * qmax)])
        under = len(df_data[df_data[flow_col] <= qmin])

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
            "Sheet": sheet_name,
            "GSize": gsize_raw,
            "Qmin": qmin,
            "Qmax": qmax,
            "Total Jam": total_jam,
            "Persen Flow >=150%": persen_150,
            "Persen Flow >=120%": persen_120,
            "Persen Flow >=100%": persen_100,
            "Persen Flow <=Qmin": persen_under,
            "Kesimpulan Bulan {}".format(month_name): kesimpulan
        })

    return pd.DataFrame(all_results)

# === STREAMLIT UI ===
def main():
    st.title("AVE Flow Analysis")

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login & Download"):
        with st.spinner("Downloading data..."):
            file_path = login_and_download(username, password)

        st.success("Download complete: " + file_path)

        month_name = "Juli2025"
        result_df = process_xls(file_path, month_name)
        st.dataframe(result_df)

        st.download_button("Download Hasil CSV", data=result_df.to_csv(index=False), file_name="hasil.csv")

if __name__ == "__main__":
    main()
