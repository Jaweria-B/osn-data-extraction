import streamlit as st
import atexit
from selenium import webdriver
from openpyxl import load_workbook
import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.core.os_manager import ChromeType
import time

# Set Streamlit page configuration
st.set_page_config(
    page_title="OSN Data Extraction Tool",
    page_icon="📊",
    layout="centered",
)

# Function to initialize Chrome driver
# def initialize_driver():
#     options = uc.ChromeOptions()
#     options.add_argument("--no-sandbox")
#     options.add_argument("--disable-dev-shm-usage")
#     options.add_argument("--start-maximized")
#     service = Service(ChromeDriverManager().install())
#     driver = uc.Chrome(service=service, options=options)
#     driver.implicitly_wait(10)
#     atexit.register(driver.quit)  # Ensure driver quits on exit
#     return driver

# Initialize Chrome WebDriver
def initialize_driver():
    options = Options()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--headless")  # Use headless mode for deployment
    options.add_argument("--disable-gpu")
    
    service = Service(ChromeDriverManager(chrome_type=ChromeType.CHROMIUM).install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.implicitly_wait(10)
    return driver

# Safe element finder
def safe_find_element(driver, by, value, timeout=30):
    try:
        return WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, value))
        )
    except TimeoutException:
        return None

# Streamlit UI setup
st.markdown("<h1 style='text-align: center; color: #4CAF50;'>📊 OSN Customer Data Extraction Tool</h1>", unsafe_allow_html=True)
st.write(
    """
    Welcome to the **OSN Data Extraction Tool**! Upload your Excel file, log in, and let the app automatically extract and update customer data for you.
    """
)

# Upload Excel file
uploaded_file = st.file_uploader("📂 Upload your Excel file (.xlsx)", type=["xlsx"], help="Upload an Excel file containing Hardware Numbers.")

if uploaded_file is not None:
    # Load Excel file into DataFrame
    df = pd.read_excel(uploaded_file)
    st.success("File uploaded successfully! Preview below:")
    st.dataframe(df.head(), use_container_width=True)

    df = df.dropna(subset=["Hardware No"])
    df["Hardware No"] = df["Hardware No"].astype(int)  # Ensure numeric data


    # Initialize driver and navigate to login page
    if st.button("🚀 Initialize and Login"):
        driver = initialize_driver()
        st.session_state.driver = driver
        driver.get("https://unify.osn.com/")
        st.info("Navigated to login page. Please log in manually.")
        st.session_state.logged_in = True

    # Start extraction after login
    if "logged_in" in st.session_state and st.session_state.logged_in:
        if st.button("🔍 Start Data Extraction"):
            st.info("Starting data extraction...")
            progress = st.progress(0)
            results = []

            # Get the driver from session_state
            driver = st.session_state.driver

            # Loop through each row in the Excel file
            for index, row in df.iterrows():
                hardware_no = "000" + str(row["Hardware No"])  # Adjust column name to match your Excel sheet
                
                try:
                    # Select "Smartcard Number" in the dropdown
                    dropdown = Select(driver.find_element(By.ID, "SearchCustomDdl"))
                    dropdown.select_by_value("3|numeric")  # Select Smartcard Number
                    
                    # Enter the Hardware Number in the input field
                    search_box = driver.find_element(By.ID, "txtSearch")  # Adjust ID if necessary
                    search_box.clear()
                    search_box.send_keys(hardware_no)

                    # Click Search Button
                    search_button = driver.find_element(By.ID, "btnSearch2")
                    search_button.click()
                    time.sleep(2)  # Wait for results to load

                    # Extract Customer No
                    try:
                        customer_no_element = WebDriverWait(driver, 60).until(
                            EC.presence_of_element_located((By.XPATH, "//span/a[contains(@href, 'CustomerLandingNew.aspx')]"))
                        )
                        customer_no = customer_no_element.text.strip()
                    except TimeoutException:
                        st.error(f"⚠️ Timeout while waiting for Customer No for Hardware No: {hardware_no}")
                        continue

                    # Extract Customer Name
                    try:
                        customer_name_element = WebDriverWait(driver, 60).until(
                            EC.presence_of_element_located((By.ID, "MainContent_CustPersonalInfo_lblName"))
                        )
                        customer_name = customer_name_element.text.strip()
                    except TimeoutException:
                        st.error(f"⚠️ Timeout while waiting for Customer Name for Hardware No: {hardware_no}")
                        continue

                    # Extract Mobile
                    try:
                        mobile_element = WebDriverWait(driver, 60).until(
                            EC.presence_of_element_located((By.ID, "MainContent_CustPersonalInfo_hplMobile"))
                        )
                        mobile = mobile_element.text.strip()
                    except TimeoutException:
                        st.error(f"⚠️ Timeout while waiting for Mobile for Hardware No: {hardware_no}")
                        continue

                    # Extract Table Data
                    try:
                        table = WebDriverWait(driver, 60).until(
                            EC.presence_of_element_located((By.ID, "JColResizerGridProductList"))
                        )
                        tbody = table.find_element(By.TAG_NAME, "tbody")
                        rows = tbody.find_elements(By.TAG_NAME, "tr")

                        product_value = None
                        charge_until_date = None
                        contract_end_date = None

                        # Iterate over rows to find the one with "OSN SW Packages"
                        for row in rows:
                            cells = row.find_elements(By.TAG_NAME, "td")
                            if len(cells) >= 18:  # Ensure row has enough columns
                                product_category = cells[16].text.strip()  # Adjust index for Product Category
                                if product_category == "OSN SW Packages":
                                    product_value = cells[1].text.strip()  # Adjust index for Product
                                    charge_until_date = cells[8].text.strip()  # Adjust index for Charge Until Date
                                    contract_end_date = cells[10].text.strip()  # Adjust index for Contract End Date
                                    break

                    except TimeoutException:
                        st.error(f"⚠️ Timeout while waiting for table data for Hardware No: {hardware_no}")
                        continue


                    if product_value and charge_until_date and contract_end_date:
                        # Update Excel DataFrame
                        df.at[index, "Product"] = product_value
                        df.at[index, "Charge Until Date"] = charge_until_date
                        df.at[index, "Expire Date"] = contract_end_date

                        st.success(f"✅ Data extracted for Hardware No: {hardware_no}")
                    else:
                        st.warning(f"No matching row found for Hardware No: {hardware_no}")

                    # Update progress bar
                    progress.progress((index + 1) / len(df))

                except NoSuchElementException as e:
                    st.error(f"⚠️ Failed for Hardware No: {hardware_no}, Error: {e}")
                except Exception as e:
                    st.error(f"⚠️ Unexpected error for Hardware No: {hardware_no}, Error: {e}")

            # Save the updated DataFrame
            updated_file = "updated_data.xlsx"
            df.to_excel(updated_file, index=False)
            st.success("🎉 Data extraction complete!")

            # Download button for the updated file
            st.download_button(
                label="💾 Download Updated File",
                data=open(updated_file, "rb"),
                file_name="updated_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

# Footer
st.markdown(
    """
    <hr>
    <p style='text-align: center;'>
        Developed by <a href="https://hitech-experts.com/" target="_blank">Hitech Experts</a> ❤️
    </p>
    """,
    unsafe_allow_html=True,
)
