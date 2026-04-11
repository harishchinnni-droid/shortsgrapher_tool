"""
Broker Authentication Module
Handles login for Angel One and Zerodha trading platforms.
"""

import json
import pyotp
import time
import os
from datetime import datetime
from urllib.parse import urlparse, parse_qs
from SmartApi import SmartConnect
from kiteconnect import KiteConnect
import kiteconnect.exceptions as ex
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementNotInteractableException, StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager

from src.logging_config import LoggingManager

logger = LoggingManager.get_logger()


class BrokerAuthentication:
    """Handles authentication for multiple brokers."""
    
    def __init__(self, json_dir, cache_file):
        """
        Initialize broker authentication.
        
        Args:
            json_dir: Directory containing credential JSON files
            cache_file: Path to session cache file
        """
        self.json_dir = json_dir
        self.cache_file = cache_file
        self.logger = logger
    
    def load_credentials(self, filepath):
        """
        Load credentials from JSON file.
        
        Args:
            filepath: Path to credentials file
            
        Returns:
            Dictionary containing credentials
        """
        try:
            with open(filepath, 'r') as file:
                credentials = json.load(file)
                self.logger.info(f"Loaded credentials from {filepath}")
                return credentials
        except FileNotFoundError:
            error_msg = f"Credentials file not found at {filepath}"
            self.logger.error(error_msg)
            raise Exception(f"CRITICAL ERROR: {error_msg}")
        except json.JSONDecodeError:
            error_msg = f"Invalid JSON structure in {filepath}"
            self.logger.error(error_msg)
            raise Exception(f"CRITICAL ERROR: {error_msg}")
    
    def get_today_str(self):
        """Get today's date as string."""
        return datetime.now().strftime("%Y-%m-%d")
    
    def load_cache(self):
        """Load session cache if available and not stale."""
        if not os.path.exists(self.cache_file):
            self.logger.debug("Cache file does not exist")
            return {}
        
        try:
            with open(self.cache_file, 'r') as file:
                cache = json.load(file)
                if cache.get("date") != self.get_today_str():
                    self.logger.warning("Stale cache detected. Forcing new authentication.")
                    return {}
                self.logger.info("Cache loaded successfully")
                return cache
        except Exception as e:
            self.logger.warning(f"Error reading cache: {e}. Rebuilding.")
            return {}
    
    def update_cache(self, key, value):
        """Update session cache."""
        cache = self.load_cache()
        cache["date"] = self.get_today_str()
        cache[key] = value
        
        os.makedirs(os.path.dirname(self.cache_file), exist_ok=True)
        
        with open(self.cache_file, 'w') as file:
            json.dump(cache, file, indent=4)
        
        self.logger.info(f"Cache updated: {key}")
    
    def login_angel_one(self, creds):
        """
        Authenticate with Angel One broker.
        
        Args:
            creds: Credentials dictionary
            
        Returns:
            SmartConnect API instance or None
        """
        self.logger.info("\n[ANGEL ONE] Initiating Connection...")
        
        try:
            api_key = creds.get("api_key")
            client_id = creds.get("client_id") or creds.get("user_id")
            password = creds.get("password")
            totp_secret = creds.get("totp_secret")
            
            smartApi = SmartConnect(api_key=api_key)
            
            cache = self.load_cache()
            if "angel_token" in cache and "angel_feed_token" in cache:
                smartApi.access_token = cache["angel_token"]
                smartApi.feed_token = cache["angel_feed_token"]
                self.logger.info("[ANGEL ONE] SUCCESS. Session restored from cache.")
                return smartApi
            
            self.logger.info("[ANGEL ONE] Generating new session...")
            totp = pyotp.TOTP(totp_secret).now()
            data = smartApi.generateSession(client_id, password, totp)
            
            if data['status']:
                self.update_cache("angel_token", data['data']['jwtToken'])
                self.update_cache("angel_feed_token", data['data']['feedToken'])
                self.logger.info("[ANGEL ONE] SUCCESS. Access Token Generated and Cached.")
                return smartApi
            else:
                error_msg = f"Login Failed: {data['message']}"
                self.logger.error(f"[ANGEL ONE] {error_msg}")
                return None
        
        except Exception as e:
            self.logger.error(f"[ANGEL ONE] FAILURE: {e}", exc_info=True)
            return None
    
    def login_zerodha(self, creds):
        """
        Authenticate with Zerodha broker using Selenium automation.
        
        Args:
            creds: Credentials dictionary
            
        Returns:
            KiteConnect API instance or None
        """
        self.logger.info("\n[ZERODHA] Initiating Connection...")
        driver = None
        
        try:
            api_key = creds.get("api_key")
            api_secret = creds.get("api_secret")
            user_id = creds.get("user_id")
            password = creds.get("password")
            
            kite = KiteConnect(api_key=api_key)
            
            cache = self.load_cache()
            if "zerodha_access_token" in cache:
                cached_token = cache["zerodha_access_token"]
                kite.set_access_token(cached_token)
                
                try:
                    kite.profile()
                    self.logger.info("[ZERODHA] SUCCESS. Session restored from cache and verified.")
                    return kite
                except ex.TokenException:
                    self.logger.warning("[ZERODHA] WARNING: Cached token is invalid or expired. Forcing new login.")
                except Exception as e:
                    self.logger.warning(f"[ZERODHA] WARNING: API test failed ({e}). Forcing new login.")
            
            self.logger.info("[ZERODHA] Spinning up browser for automated credential injection...")
            login_url = kite.login_url()
            
            chrome_options = Options()
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--window-size=1920,1080")
            
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=chrome_options)
            
            driver.get(login_url)
            wait = WebDriverWait(driver, 15)
            
            self.logger.info("[ZERODHA-DEBUG] Injecting core credentials...")
            userid_field = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='text']")))
            userid_field.send_keys(user_id)
            
            password_field = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='password']")))
            password_field.send_keys(password)
            
            submit_btn = driver.find_element(By.XPATH, "//button[@type='submit']")
            submit_btn.click()
            
            self.logger.info("[ZERODHA-DEBUG] Observing browser state for TOTP or Redirect...")
            
            timeout = time.time() + 20
            redirect_url = None
            totp_injected = False
            
            while time.time() < timeout:
                try:
                    current_url = driver.current_url
                    
                    if "request_token=" in current_url:
                        redirect_url = current_url
                        break
                    
                    if not totp_injected:
                        inputs = driver.find_elements(By.XPATH, "//form//input[not(@type='hidden')]")
                        for field in inputs:
                            if field.is_displayed() and field.is_enabled():
                                self.logger.info("Manual TOTP input required (see console)")
                                manual_totp = input(">>> ACTION REQUIRED: Enter 6-digit Zerodha App Code / TOTP: ").strip()
                                
                                if not manual_totp or len(manual_totp) != 6:
                                    raise ValueError("Invalid TOTP length entered. Script aborted.")
                                
                                field.send_keys(manual_totp)
                                time.sleep(0.5)
                                field.send_keys(Keys.RETURN)
                                totp_injected = True
                                break
                
                except StaleElementReferenceException:
                    pass
                except ValueError as ve:
                    raise ve
                except Exception:
                    pass
                
                time.sleep(0.5)
            
            if not redirect_url:
                raise Exception("Timeout: Execution stalled. Failed to reach redirect URL within 20 seconds.")
            
            driver.quit()
            
            parsed_url = urlparse(redirect_url)
            request_token = parse_qs(parsed_url.query).get('request_token')
            
            if not request_token:
                raise Exception("Redirected, but 'request_token' parameter is missing.")
            
            data = kite.generate_session(request_token[0], api_secret=api_secret)
            access_token = data["access_token"]
            kite.set_access_token(access_token)
            
            self.update_cache("zerodha_access_token", access_token)
            
            self.logger.info("[ZERODHA] SUCCESS. Access Token Generated and Cached.")
            return kite
        
        except Exception as e:
            self.logger.error(f"[ZERODHA] FAILURE: {e}", exc_info=True)
            if driver:
                driver.quit()
            return None
