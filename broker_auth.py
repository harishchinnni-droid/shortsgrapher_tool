import os
import json
import time
import pyotp
import traceback
import urllib.parse as urlparse
from datetime import datetime
from ist_clock import now_ist
from kiteconnect import KiteConnect
from SmartApi import SmartConnect
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementNotInteractableException, StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager
import file_mgmt

# Paths
# [CHANGED -- cloud/Colab portability] derives from file_mgmt.BASE_DIR
# (itself ALGO_BASE_DIR-overridable) instead of hardcoding this pipeline's
# root a second time -- see file_mgmt.py's BASE_DIR docstring.
JSON_DIR = os.path.join(file_mgmt.BASE_DIR, "01_JSON_Files")
ANGEL_CREDS_FILE  = os.path.join(JSON_DIR, "harish_angel_one.json")
ZERODHA_CREDS_FILE = os.path.join(JSON_DIR, "harish_zerodha.json")
ZERODHA_TOKEN_CACHE = os.path.join(JSON_DIR, "zerodha_access_token.json")
# [ADDED] Angel One previously had no same-day cache -- initialize_angel_one()
# did a full fresh login every single call. Zerodha's login is expensive
# (headless browser) so it needed a cache; Angel One's is a cheap direct API
# call, which is probably why this was skipped originally. But Harish's own
# spec for this pipeline was "save the token for using it again on same
# date" for BOTH brokers, and re-logging in on every call still means
# needlessly generating a brand new session (and burning a TOTP window) each
# time this is invoked mid-day -- so bringing it in line with Zerodha's
# pattern rather than leaving the asymmetry.
ANGEL_TOKEN_CACHE = os.path.join(JSON_DIR, "angel_one_access_token.json")

def load_json(filepath):
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"CRITICAL ERROR: Configuration missing at {filepath}")
    with open(filepath, 'r') as f:
        return json.load(f)

def initialize_angel_one():
    print("[SYSTEM] Initializing Angel One...")
    creds = load_json(ANGEL_CREDS_FILE)
    smart_api = SmartConnect(api_key=creds['api_key'])

    # Same-day cache reuse, mirroring initialize_zerodha() below. Angel One
    # tokens are session-based (not guaranteed to expire on a fixed IST
    # clock the way Kite's do), so if a cached jwt/refresh token stops being
    # accepted mid-day, the fix is to delete ANGEL_TOKEN_CACHE's file and
    # let this fall through to a fresh login again -- not to disable caching.
    today_str = now_ist().strftime("%Y-%m-%d")
    if os.path.exists(ANGEL_TOKEN_CACHE):
        with open(ANGEL_TOKEN_CACHE, 'r') as f:
            cache = json.load(f)
            if cache.get('date') == today_str and cache.get('jwt_token') and cache.get('refresh_token'):
                smart_api.setAccessToken(cache['jwt_token'])
                smart_api.setRefreshToken(cache['refresh_token'])
                if cache.get('feed_token'):
                    smart_api.setFeedToken(cache['feed_token'])
                print("[SYSTEM] Angel One: reused cached token for today.")
                return smart_api

    totp = pyotp.TOTP(creds['totp_secret']).now()
    data = smart_api.generateSession(creds['client_id'], creds['password'], totp)
    if not data['status']:
        raise Exception(f"[FAILURE] Angel One Login Failed: {data}")

    # generateSession() already calls setAccessToken/setRefreshToken/
    # setFeedToken internally -- read the raw (unprefixed) values straight
    # off smart_api rather than the response dict, whose data.jwtToken comes
    # back with a "Bearer " prefix baked in that would break auth headers
    # if it were ever fed back into setAccessToken() as-is.
    with open(ANGEL_TOKEN_CACHE, 'w') as f:
        json.dump({
            "date": today_str,
            "jwt_token": smart_api.access_token,
            "refresh_token": smart_api.refresh_token,
            "feed_token": smart_api.feed_token,
        }, f)

    return smart_api

def initialize_zerodha():
    print("[SYSTEM] Initializing Zerodha...")
    creds = load_json(ZERODHA_CREDS_FILE)
    kite = KiteConnect(api_key=creds['api_key'])
    
    # [FIX] was datetime.now() -- Kite access tokens expire on the IST
    # trading-day boundary, not the host machine's local midnight.
    today_str = now_ist().strftime("%Y-%m-%d")
    if os.path.exists(ZERODHA_TOKEN_CACHE):
        with open(ZERODHA_TOKEN_CACHE, 'r') as f:
            cache = json.load(f)
            if cache.get('date') == today_str and cache.get('access_token'):
                kite.set_access_token(cache['access_token'])
                return kite

    print("[SYSTEM] Initiating Automated Selenium Login for Zerodha...")
    options = webdriver.ChromeOptions()
    options.add_argument('--headless=new')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--start-maximized')
    options.add_argument('--window-size=1366,900')
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36')
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    options.add_experimental_option('useAutomationExtension', False)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    try:
        driver.get(f"https://kite.trade/connect/login?v=3&api_key={creds['api_key']}")
        wait = WebDriverWait(driver, 20)

        wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='text']"))).send_keys(creds['user_id'])
        driver.find_element(By.XPATH, "//input[@type='password']").send_keys(creds['password'])
        
        password_field = driver.find_element(By.XPATH, "//input[@type='password']")
        driver.find_element(By.XPATH, "//button[@type='submit']").click()
        wait.until(EC.staleness_of(password_field))

        totp_value = pyotp.TOTP(creds['totp_secret']).now()

        def _find_visible_input():
            for c in driver.find_elements(By.XPATH, "//input[@type='text' or @type='tel' or @type='number' or @type='password']"):
                try:
                    if c.is_displayed() and c.is_enabled():
                        return c
                except Exception:
                    continue
            return None

        wait.until(lambda d: _find_visible_input() is not None or d.find_elements(By.TAG_NAME, "iframe"))
        totp_input = _find_visible_input()
        in_frame = False

        if totp_input is None:
            frames = driver.find_elements(By.TAG_NAME, "iframe")
            for idx, frame in enumerate(frames):
                driver.switch_to.frame(frame)
                totp_input = _find_visible_input()
                if totp_input is not None:
                    in_frame = True
                    break
                driver.switch_to.default_content()

        if totp_input is None:
            driver.switch_to.default_content()
            raise Exception("Could not locate a visible TOTP input field.")

        totp_input.click()
        totp_input.send_keys(totp_value)

        # Kite's TOTP field auto-submits via its own JS listener as soon as
        # the 6th digit lands -- it doesn't wait for Enter. That means this
        # keypress is racing the page's own redirect: usually harmless, but
        # if the auto-submit already fired, the input can be mid-transition
        # (hidden/detached/disabled) by the time Selenium gets here, raising
        # ElementNotInteractableException or StaleElementReferenceException.
        # The login has very likely already gone through in that case, so
        # swallow it here (same pattern already used for the submit-button
        # fallback right below) instead of failing the whole run over it.
        try:
            totp_input.send_keys(Keys.RETURN)
        except (ElementNotInteractableException, StaleElementReferenceException):
            pass

        try:
            driver.find_element(By.XPATH, "//button[@type='submit' or contains(text(),'Continue')]").click()
        except Exception:
            pass

        if in_frame:
            driver.switch_to.default_content()

        wait.until(EC.url_contains("request_token"))
        request_token = urlparse.parse_qs(urlparse.urlparse(driver.current_url).query)['request_token'][0]
    except Exception as e:
        driver.quit()
        raise Exception(f"[FAILURE] Selenium Automation Failed: {traceback.format_exc()}")
    driver.quit()

    access_token = kite.generate_session(request_token, api_secret=creds['api_secret'])["access_token"]
    kite.set_access_token(access_token)

    with open(ZERODHA_TOKEN_CACHE, 'w') as f:
        json.dump({"date": today_str, "access_token": access_token}, f)
    return kite
