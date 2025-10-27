# app.py
# ==============================
# MSC Tracker – Paste Containers + Download Excel
# ==============================

import time
import random
import logging
import pandas as pd
from io import BytesIO
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import streamlit as st

# ---------------- CONFIG ----------------
HEADLESS = False
MAX_WAIT = 10
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36"
)
# ----------------------------------------

logging.basicConfig(format="%(asctime)s | %(levelname)s | %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)


def tiny_pause(a=0.06, b=0.18):
    time.sleep(random.uniform(a, b))


def create_driver(headless=HEADLESS):
    opts = Options()
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.add_argument("--window-size=1400,900")
    opts.add_argument(f"--user-agent={USER_AGENT}")
    if headless:
        opts.add_argument("--headless=new")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=opts)
    wait = WebDriverWait(driver, MAX_WAIT)
    return driver, wait


def close_cookie_popup_if_present(driver, wait):
    try:
        selectors = [
            (By.ID, "onetrust-accept-btn-handler"),
            (By.CSS_SELECTOR, "button#onetrust-accept-btn-handler"),
            (By.XPATH, "//button[contains(text(),'Accept') or contains(text(),'Accept All')]"),
        ]
        for by, sel in selectors:
            try:
                btn = wait.until(EC.element_to_be_clickable((by, sel)))
                driver.execute_script("arguments[0].scrollIntoView(true);", btn)
                tiny_pause(0.08, 0.18)
                try:
                    btn.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", btn)
                break
            except Exception:
                continue
    except Exception:
        pass


def get_results_snapshot(driver):
    try:
        el = driver.find_element(By.CSS_SELECTOR, ".msc-flow-tracking__data, .msc-flow-tracking__cell")
        return el.text.strip()[:200]
    except Exception:
        return ""


def extract_tracking_data(driver):
    data = {
        "ETA": None,
        "Port of Discharge": None,
        "Vessel/Voyage": None,
        "Equipment Handling Facility": None,
    }
    try:
        eta = driver.find_element(
            By.XPATH,
            "//span[contains(@class,'data-heading')][contains(text(),'POD ETA')]/following-sibling::span",
        ).text.strip()
        data["ETA"] = eta or None
    except Exception:
        pass

    try:
        pod = driver.find_element(
            By.XPATH, "//span[contains(text(),'Port of Discharge')]/following-sibling::span"
        ).text.strip()
        data["Port of Discharge"] = pod or None
    except Exception:
        pass

    try:
        vessel = driver.find_element(
            By.XPATH,
            "//div[contains(@class,'msc-flow-tracking__cell--five')]//span[contains(@class,'data-value') and normalize-space(text())!='N.A']",
        ).text.strip()
        if vessel:
            data["Vessel/Voyage"] = vessel
    except Exception:
        pass

    try:
        facility = driver.find_element(
            By.XPATH,
            "//div[contains(@class,'msc-flow-tracking__cell--six')]//span[contains(@class,'data-value') and normalize-space(text())!='N.A']",
        ).text.strip()
        if facility:
            data["Equipment Handling Facility"] = facility
    except Exception:
        pass

    return data


def wait_for_change(driver, prev_snapshot, timeout=6):
    end = time.time() + timeout
    while time.time() < end:
        cur = get_results_snapshot(driver)
        if cur and cur != prev_snapshot:
            return True
        time.sleep(0.12)
    return False


def submit_container_quick(driver, input_el, container):
    script = """
    const el = arguments[0];
    const val = arguments[1];
    el.focus();
    el.value = val;
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
    """
    driver.execute_script(script, input_el, container)
    tiny_pause(0.02, 0.08)
    input_el.send_keys(Keys.RETURN)


def track_containers(container_list):
    driver, wait = create_driver(HEADLESS)
    results = []
    try:
        driver.get("https://www.msc.com/en/track-a-shipment")
        tiny_pause(0.3, 0.6)
        close_cookie_popup_if_present(driver, wait)
        input_field = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input#trackingNumber")))
        prev_snapshot = get_results_snapshot(driver)

        for idx, container in enumerate(container_list, start=1):
            st.info(f"Tracking {idx}/{len(container_list)} → {container}")
            submit_container_quick(driver, input_field, container)
            changed = wait_for_change(driver, prev_snapshot, timeout=6)
            if changed:
                tiny_pause(0.12, 0.35)
            data = extract_tracking_data(driver)
            results.append({"Container Number": container, **data})
            prev_snapshot = get_results_snapshot(driver)
            tiny_pause(0.4, 0.9)
    finally:
        driver.quit()
    return pd.DataFrame(results)


# ---------------- STREAMLIT UI ----------------

st.set_page_config(page_title="MSC Container Tracker", layout="centered")

st.title("🚢 MSC Container Tracker")
st.write("Paste all your container numbers (one per line) below and click **Track**.")

container_input = st.text_area(
    "Container Numbers", placeholder="E.g.\nMSDU5837828\nCAAU8042212\nTLLU8783634", height=200
)

if st.button("Track Containers 🚀"):
    container_list = [c.strip() for c in container_input.splitlines() if c.strip()]
    if not container_list:
        st.warning("Please enter at least one container number.")
    else:
        with st.spinner("Scraping data from MSC... Please wait ⏳"):
            df_results = track_containers(container_list)
        st.success("✅ Tracking complete!")
        st.dataframe(df_results)

        # Convert to Excel and prepare for download
        output = BytesIO()
        df_results.to_excel(output, index=False)
        output.seek(0)
        st.download_button(
            label="📥 Download tracked_containers.xlsx",
            data=output,
            file_name="tracked_containers.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
