# main.py
# ===========================================
# MSC Container Tracking (Faster — single page, quicker waits)
# ===========================================

import time
import random
import logging
import sys
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ---------------- CONFIG ----------------
INPUT_FILE = "data.xlsx"
OUTPUT_FILE = "tracked_containers.xlsx"
HEADLESS = False      # visible browser
MAX_WAIT = 10         # shorter explicit wait
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36"
)
# ----------------------------------------

# logging
logging.basicConfig(format="%(asctime)s | %(levelname)s | %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)


def tiny_pause(a=0.06, b=0.18):
    time.sleep(random.uniform(a, b))


def create_driver(headless=HEADLESS):
    opts = Options()
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--disable-infobars")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.add_argument("--window-size=1400,900")
    opts.add_argument(f"--user-agent={USER_AGENT}")

    if headless:
        opts.add_argument("--headless=new")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=opts)

    # small navigator masks to avoid trivial detection
    try:
        driver.execute_cdp_cmd(
            "Page.addScriptToEvaluateOnNewDocument",
            {
                "source": """
                    Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                    Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3,4,5]});
                    Object.defineProperty(navigator, 'languages', {get: () => ['en-US','en']});
                """
            },
        )
    except Exception:
        pass

    wait = WebDriverWait(driver, MAX_WAIT)
    return driver, wait


def close_cookie_popup_if_present(driver, wait):
    """Accept cookie popup if present (OneTrust or similar)"""
    try:
        selectors = [
            (By.ID, "onetrust-accept-btn-handler"),
            (By.CSS_SELECTOR, "button#onetrust-accept-btn-handler"),
            (By.XPATH, "//button[contains(text(),'Accept') or contains(text(),'Accept All') or contains(text(),'I Agree')]")
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
                logger.info("Closed cookie popup via %s %s", by, sel)
                break
            except Exception:
                continue

        # wait briefly for overlay to disappear (non-fatal)
        try:
            wait.until_not(EC.presence_of_element_located((By.CLASS_NAME, "onetrust-pc-dark-filter")))
        except Exception:
            pass
    except Exception:
        pass


def get_results_snapshot(driver):
    """Return a short snapshot string of the results area used to detect changes quickly."""
    try:
        # prefer a data cell if present
        el = driver.find_element(By.CSS_SELECTOR, ".msc-flow-tracking__data, .msc-flow-tracking__cell")
        text = el.text.strip()
        # use only first ~100 chars to make compare cheap
        return text[:200]
    except Exception:
        return ""


def extract_tracking_data(driver):
    """
    Return dict with ETA, Port of Discharge, Vessel/Voyage, Equipment Handling Facility.
    """
    data = {"ETA": None, "Port of Discharge": None, "Vessel/Voyage": None, "Equipment Handling Facility": None}

    try:
        eta_el = driver.find_element(
            By.XPATH,
            "//span[contains(@class,'data-heading')][contains(text(),'POD ETA')]/following-sibling::span"
        )
        data["ETA"] = eta_el.text.strip() or None
    except Exception:
        pass

    try:
        pod_el = driver.find_element(
            By.XPATH,
            "//span[contains(text(),'Port of Discharge')]/following-sibling::span"
        )
        data["Port of Discharge"] = pod_el.text.strip() or None
    except Exception:
        pass

    try:
        vessel_el = driver.find_element(
            By.XPATH,
            "//div[contains(@class,'msc-flow-tracking__cell--five')]//span[contains(@class,'data-value') and normalize-space(text())!='N.A']"
        )
        txt = vessel_el.text.strip()
        if txt:
            data["Vessel/Voyage"] = txt
    except Exception:
        # fallback
        try:
            vessel_el2 = driver.find_element(
                By.XPATH,
                "//div[contains(@class,'msc-flow-tracking__cell') and .//span[contains(text(),'Vessel') or contains(text(),'Voyage')]]//span[contains(@class,'data-value')]"
            )
            txt2 = vessel_el2.text.strip()
            if txt2 and txt2.upper() != "N.A":
                data["Vessel/Voyage"] = txt2
        except Exception:
            pass

    try:
        facility_el = driver.find_element(
            By.XPATH,
            "//div[contains(@class,'msc-flow-tracking__cell--six')]//span[contains(@class,'data-value') and normalize-space(text())!='N.A']"
        )
        t = facility_el.text.strip()
        if t:
            data["Equipment Handling Facility"] = t
    except Exception:
        # fallback tooltip
        try:
            facility_el2 = driver.find_element(
                By.XPATH,
                "//div[contains(@class,'msc-flow-tracking__cell--six')]//div[contains(@class,'msc-flow-tracking__tooltip')]//span[contains(@class,'data-value')]"
            )
            t2 = facility_el2.text.strip()
            if t2 and t2.upper() != "N.A":
                data["Equipment Handling Facility"] = t2
        except Exception:
            pass

    return data


def submit_container_quick(driver, input_el, container_number):
    """
    Set the input value via JS and trigger input events, then press Enter.
    This is faster than clicking + typing each time.
    """
    # set value and dispatch input event (some frameworks rely on it)
    script = """
    const el = arguments[0];
    const val = arguments[1];
    el.focus();
    el.value = val;
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
    """
    driver.execute_script(script, input_el, container_number)
    tiny_pause(0.02, 0.08)
    input_el.send_keys(Keys.RETURN)


def wait_for_change(driver, prev_snapshot, timeout=6):
    """Wait until results snapshot changes or timeout (short). Returns True if changed."""
    end = time.time() + timeout
    while time.time() < end:
        cur = get_results_snapshot(driver)
        if cur and cur != prev_snapshot:
            return True
        # small sleep to avoid tight loop
        time.sleep(0.12)
    return False


def main():
    # input
    df = pd.read_excel(INPUT_FILE)
    if "Container Number" not in df.columns:
        raise ValueError("Input Excel must contain a 'Container Number' column.")

    driver, wait = create_driver(headless=HEADLESS)
    results = []
    total = len(df)
    logger.info("Opening MSC tracking page once and reusing it for all containers.")
    try:
        # open page once
        driver.get("https://www.msc.com/en/track-a-shipment")
        tiny_pause(0.2, 0.6)

        # accept cookies once
        close_cookie_popup_if_present(driver, wait)

        # locate the input once (reuse)
        input_field = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input#trackingNumber")))
        # ensure visible
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", input_field)
        tiny_pause(0.06, 0.12)

        prev_snapshot = get_results_snapshot(driver)

        for idx, row in df.iterrows():
            container = str(row["Container Number"]).strip()
            logger.info("[%d/%d] Tracking %s", idx + 1, total, container)

            # quick submit (JS + Enter)
            submit_container_quick(driver, input_field, container)

            # wait for small change in results (short timeout)
            changed = wait_for_change(driver, prev_snapshot, timeout=6)

            # small extra pause to let rendering settle if changed
            if changed:
                tiny_pause(0.12, 0.35)
            else:
                # if no change, give a tiny bit more time (rare)
                tiny_pause(0.4, 0.9)

            # extract
            data = extract_tracking_data(driver)
            results.append({**row.to_dict(), **data})

            # update snapshot (cheap)
            prev_snapshot = get_results_snapshot(driver)

            # VERY small randomized pause before next iteration
            tiny_pause(0.5, 1.1)

    finally:
        driver.quit()
        logger.info("Chrome driver closed")

    # save
    pd.DataFrame(results).to_excel(OUTPUT_FILE, index=False)
    logger.info("Saved results to %s", OUTPUT_FILE)


if __name__ == "__main__":
    main()
