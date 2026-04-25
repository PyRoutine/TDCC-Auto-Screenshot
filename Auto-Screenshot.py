import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import sys
import os
import time
import io
import re
import datetime
import hashlib
import base64
import subprocess
import uuid
import json

from PIL import Image

# 強化加密所需模組
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.backends import default_backend

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import StaleElementReferenceException

# --- 自定義異常 ---
class SystemMaintenanceError(Exception):
    pass

class LoginTimeoutError(Exception): 
    pass

# --- 全域變數初始化 ---
driver = None
base_path = "./screenshots/"
user_accounts = [] 
shot_speed = 1
main_window_geom = ""
screenshot_mode = 1 
login_type = "券商網路下單憑證" 
disclaimer_agreed = False
join_draw = False 
session_results = {}
user_name_map = {}
execution_logs = []
name_source_mode = 1

os.environ["QTWEBENGINE_CHROMIUM_FLAGS"] = "--log-level=3"

# --- 硬體金鑰與加密 ---
def get_hw_key():
    try:
        cmd = 'wmic baseboard get serialnumber'
        output = subprocess.check_output(cmd, shell=True).decode(errors='ignore')
        hw_id = output.split('\n')[1].strip()
        
        if not hw_id or "To be filled" in hw_id or "None" in hw_id:
            cmd = 'wmic cpu get processorid'
            output = subprocess.check_output(cmd, shell=True).decode(errors='ignore')
            hw_id = output.split('\n')[1].strip()
            
        if not hw_id or "To be filled" in hw_id or "None" in hw_id:
             hw_id = str(uuid.getnode())
             
    except:
        try: hw_id = str(uuid.getnode())
        except: hw_id = "Default_Fallback_Seed_12345"
        
    salt = b'TDCC_AutoVote_Secure_Salt_V4'
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=100000,
        backend=default_backend()
    )
    return base64.urlsafe_b64encode(kdf.derive(hw_id.encode()))

cipher = Fernet(get_hw_key())

def encrypt_data(data_str):
    if not data_str: return ""
    return cipher.encrypt(data_str.encode()).decode()

def decrypt_data(encrypted_str):
    if not encrypted_str: return ""
    try:
        return cipher.decrypt(encrypted_str.encode()).decode()
    except:
        return "" 

def get_anonymous_dirname(user_id):
    if not user_id: return "unknown"
    return hashlib.sha256(user_id.encode()).hexdigest()

def get_app_data_path():
    app_data = os.getenv('APPDATA')
    if not app_data:
        app_data = os.path.expanduser("~")
    config_dir = os.path.join(app_data, "TDCC_AutoVote_Configs")
    if not os.path.exists(config_dir):
        os.makedirs(config_dir)
    return config_dir

CONFIG_DIR = get_app_data_path()

def log_msg(msg):
    global execution_logs
    now = datetime.datetime.now()
    timestamp = now.strftime("%H:%M:%S.") + f"{now.microsecond // 1000:03d}"
    log_line = f"[{timestamp}] {msg}"
    print(log_line)
    execution_logs.append(log_line)

def clean_filename(text):
    return re.sub(r'[\\/*?:"<>|]', "", text).strip()

def get_executable_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

os.chdir(get_executable_dir())

def force_quit_driver(driver_instance):
    try:
        subprocess.run("taskkill /F /IM msedge.exe /T", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        subprocess.run("taskkill /F /IM msedgedriver.exe /T", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except: pass

# --- 核心自動化函數 ---

def get_driver():
    edge_options = Options()
    edge_options.page_load_strategy = 'eager' 
    edge_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    edge_options.add_experimental_option('useAutomationExtension', False)
    edge_options.add_argument("--disable-infobars")
    edge_options.add_argument("--disable-notifications")
    edge_options.add_argument("--disable-gpu")   
    edge_options.add_argument("--force-device-scale-factor=1")
    edge_options.add_argument("--disable-features=BlockInsecurePrivateNetworkRequests,PrivateNetworkAccessSendPreflights,IsolateOrigins,site-per-process")
    edge_options.add_argument("--disable-web-security")
    edge_options.add_argument("--allow-running-insecure-content")
    edge_options.add_argument("--allow-insecure-localhost")
    edge_options.add_argument("--ignore-certificate-errors")
    edge_options.add_argument("--remote-allow-origins=*")
    prefs = {
        "profile.default_content_setting_values.notifications": 2,
        "profile.managed_default_content_settings.insecure_private_network_requests": 1,
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False
    }
    edge_options.add_experimental_option("prefs", prefs)
    try:
        driver = webdriver.Edge(options=edge_options)
        try:
            import ctypes
            screen_w = ctypes.windll.user32.GetSystemMetrics(0)
            screen_h = ctypes.windll.user32.GetSystemMetrics(1)
            
            if screen_w > 1600:
                driver.set_window_position(0, 0)
                driver.set_window_size(1550, 1000)
            else:
                target_w = int(screen_w * 0.9)
                target_h = int(screen_h * 0.9)
                pos_x = int((screen_w - target_w) / 6)
                pos_y = int((screen_h - target_h) / 5)
                driver.set_window_position(pos_x, pos_y)
                driver.set_window_size(target_w, target_h)
        except:
            driver.maximize_window()

        try:
            driver.execute_cdp_cmd("Emulation.setDeviceMetricsOverride", {
                "width": 1550,             
                "height": 1000,            
                "deviceScaleFactor": 1,    
                "mobile": False
            })
        except Exception as cdp_e:
            pass 

        return driver
    except Exception as e:
        log_msg(f"瀏覽器啟動遇到問題: {e}")
        raise e
    
def logout():
    global driver
    try:
        if driver:
            log_msg("登出中，請稍候...")
            driver.get("https://stockservices.tdcc.com.tw/evote/logout.html")
            time.sleep(0.1) 
    except: pass

def close_tdcc_upload_tab_and_back(driver_instance, original_window=None, timeout=5):
    try:
        if not original_window:
            original_window = driver_instance.current_window_handle
    except:
        original_window = None

    end_time = time.time() + timeout
    handled = False

    while time.time() < end_time:
        try: handles = driver_instance.window_handles[:]
        except: break

        for handle in handles:
            if handle == original_window: continue
            try: driver_instance.switch_to.window(handle)
            except: continue

            try: current_url = driver_instance.current_url or ""
            except: current_url = ""

            if "/TDCCWEB/upload/" in current_url:
                try: driver_instance.close()
                except: pass
                handled = True
                break

        if handled: break
        time.sleep(0.2)

    try: handles = driver_instance.window_handles[:]
    except: handles = []

    try:
        if original_window and original_window in handles:
            driver_instance.switch_to.window(original_window)
        elif handles:
            driver_instance.switch_to.window(handles[0])
    except: pass    

def pass_active_form():
    global driver, join_draw
    try:
        form_=driver.find_element(By.ID, "msgDialog")
        if "抽獎" in form_.text:
            if join_draw:
                log_msg("🎉 偵測到抽獎視窗！依設定停留 5 分鐘 (300秒) 讓您手動參加...")
                time.sleep(300)
                log_msg("⏳ 5 分鐘結束，關閉提示繼續自動流程...")
            try:
                driver.find_element(By.ID, "msgDialog_okBtn").click()
            except:
                pass
    except:
        pass
        
    try:
        form_btn=driver.find_element(By.ID, "comfirmDialog_skipBtn")
        if "抽獎" in form_btn.text:
            if join_draw:
                log_msg("🎉 偵測到抽獎活動！依設定停留 5 分鐘 (300秒) 讓您手動參加...")
                time.sleep(300)
                log_msg("⏳ 5 分鐘結束，點擊略過繼續自動流程...")
            try:
                form_btn.click()
            except:
                pass
    except:
        pass

def autoLogin(user_ID, current_login_type="券商網路下單憑證"):
    global driver, shot_speed
    log_msg(f"正在為您登入帳號: {user_ID} ({current_login_type})")
    
    base_wait = 0.1 * shot_speed
    try:
        driver.set_page_load_timeout(30) 
        driver.set_script_timeout(30)
        driver.implicitly_wait(0.2)
    except: pass

    try: driver.get("https://stockservices.tdcc.com.tw/evote/login/shareholder.html")
    except: pass

    input_timeout = max(2.0, 10.0 * shot_speed)
    start_wait = time.time()
    while time.time() - start_wait < input_timeout:
        try:
            msg_btns = driver.find_elements(By.ID, "msgDialog_okBtn")
            if msg_btns and msg_btns[0].is_displayed():
                log_msg("首頁偵測到系統對話框，嘗試關閉...")
                msg_btns[0].click()
                time.sleep(0.1)
        except: pass
        
        try:
            robot_close = driver.find_elements(By.CSS_SELECTOR, 'button[onclick="$.modal.close();return false;"]')
            if robot_close and robot_close[0].is_displayed():
                robot_close[0].click()
                time.sleep(0.1)
        except: pass

        try:
            driver.find_element(By.NAME,"pageIdNo").clear()
            driver.find_element(By.NAME,"pageIdNo").send_keys(user_ID)
            break
        except: time.sleep(base_wait)
    
    try: 
        log_msg(f"選擇登入方式: {current_login_type}")
        driver.find_element(By.NAME,"caType").send_keys(current_login_type)
    except: pass
    
    try: driver.find_element(By.ID, 'loginBtn').click()
    except: pass
    
    is_mobile_or_natural = False
    if current_login_type == "券商網路下單憑證":
        HARD_TIMEOUT_SECONDS = 20.0
        log_msg("等待券商憑證驗證 (限時20秒)...")
    else:
        is_mobile_or_natural = True
        HARD_TIMEOUT_SECONDS = 120.0
        log_msg(f"等待{current_login_type} (請注意手機/插卡，限時2分鐘)...")
    
    login_start_time = time.time()
    
    while True:
        if time.time() - login_start_time > HARD_TIMEOUT_SECONDS:
            raise LoginTimeoutError("Timeout")

        time.sleep(base_wait*0.5) 

        try:
            msg_btns = driver.find_elements(By.ID, "msgDialog_okBtn")
            if msg_btns and msg_btns[0].is_displayed():
                msg_btns[0].click()
                time.sleep(0.1)
                try: driver.find_element(By.ID, 'loginBtn').click()
                except: pass
        except: pass
        
        try:
            robot_close = driver.find_elements(By.CSS_SELECTOR, 'button[onclick="$.modal.close();return false;"]')
            if robot_close and robot_close[0].is_displayed():
                robot_close[0].click()
                time.sleep(0.1)
                try: driver.find_element(By.ID, 'loginBtn').click()
                except: pass
        except: pass
        
        try:
            from selenium.webdriver.support import expected_conditions as EC
            from selenium.webdriver.support.ui import WebDriverWait
            from selenium.common.exceptions import TimeoutException
            
            try:
                WebDriverWait(driver, 1).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                alert.accept() 
                raise LoginTimeoutError("Cert Error Alert")
            except TimeoutException:
                pass
        except LoginTimeoutError: raise
        except: pass

        try:
            current_url = driver.current_url
            login_success = False
            
            if is_mobile_or_natural:
                if "tc_estock_welshas" in current_url:
                    login_success = True
            else:
                if "login/shareholder" not in current_url:
                    login_success = True

            if login_success:
                if "tc_estock_welshas" not in current_url:
                     driver.get("https://stockservices.tdcc.com.tw/evote/shareholder/000/tc_estock_welshas.html")
                try: driver.execute_script("document.body.style.zoom = '100%'")
                except: pass
                break
        except: pass

        try:
            if "系統維護中" in driver.find_element(By.TAG_NAME,'body').text:
                raise SystemMaintenanceError("System Maintenance")
        except SystemMaintenanceError: raise
        except: pass

        for _ in range(3):
            try: 
                btn = driver.find_element(By.ID, "comfirmDialog_okBtn")
                if btn.is_displayed(): btn.click(); time.sleep(0.5)
            except: pass
            
            try:
                skip_btn = driver.find_element(By.ID, "comfirmDialog_skipBtn")
                if skip_btn.is_displayed(): skip_btn.click(); time.sleep(0.5)
            except: pass
            
            try:
                text_btns = driver.find_elements(By.XPATH, "//button[contains(text(),'略過')] | //button[contains(text(),'稍後')] | //button[contains(text(),'不參加')] | //a[contains(text(),'略過')] | //a[contains(text(),'稍後')]")
                for tb in text_btns:
                    if tb.is_displayed():
                        tb.click(); time.sleep(0.5); break
            except: pass

            try:
                agree_link = driver.find_element(By.CSS_SELECTOR, 'a[id="agreeLink"]')
                if agree_link.is_displayed():
                    original_window = driver.current_window_handle
                    old_handles = set(driver.window_handles)
                    agree_link.click()
                    for _ in range(20):
                        try:
                            if len(set(driver.window_handles) - old_handles) > 0: break
                        except: pass
                        time.sleep(0.2)
                    close_tdcc_upload_tab_and_back(driver, original_window=original_window, timeout=5)
                    time.sleep(0.5)
            except: pass

            try:
                agree_terms = driver.find_element(By.CSS_SELECTOR, 'input[id="agreeTerms"]')
                if agree_terms.is_displayed() and not agree_terms.is_selected():
                    agree_terms.click(); time.sleep(0.5)
            except: pass

            try:
                agree_btn = driver.find_element(By.CSS_SELECTOR, 'a[class="btnAgree btn-style btn-b btn-lg"]')
                if agree_btn.is_displayed(): agree_btn.click(); time.sleep(0.5)
            except: pass

            try:
                btn1 = driver.find_element(By.NAME, 'btn1')
                if btn1.is_displayed(): btn1.click(); time.sleep(0.5)
            except: pass

def screenshot(user_id, info, override_name=None):
    global base_path, screenshot_mode, driver, user_name_map, shot_speed 
    try:
        driver.execute_script("""
            document.body.style.width = '1600px';      
            document.body.style.minWidth = '1600px';   
            document.body.style.marginLeft = '0px';
            document.body.style.zoom = '1.1'; 
        """)
        time.sleep(0.2 * shot_speed) 

        x_start, y_start = 255, 170
        width_rect, height_rect = 1248, 339
        
        body_el = driver.find_element(By.TAG_NAME, "body")
        png_data = body_el.screenshot_as_png
        
        img = Image.open(io.BytesIO(png_data))
        
        left = x_start
        top = y_start
        right = x_start + width_rect
        bottom = y_start + height_rect
        
        cropped_img = img.crop((left, top, min(img.width, right), min(img.height, bottom)))
        
        display_name = override_name if override_name else user_name_map.get(user_id, str(user_id))
        stock_name = info[1].replace('*','')
        stock_name = clean_filename(stock_name)
        if len(stock_name) > 20: stock_name = stock_name[:20]

        date_prefix = datetime.datetime.now().strftime("%Y%m%d")

        if screenshot_mode == 1:
            save_dir = os.path.join(base_path, display_name)
            filename = f"{date_prefix}_{info[0]}_{stock_name}.png"
        else:
            safe_display_name = display_name if len(display_name) < 10 else display_name[:10]
            save_dir = base_path
            filename = f"{date_prefix}_{info[0]}_{stock_name}_{safe_display_name}.png"

        if not os.path.exists(save_dir): os.makedirs(save_dir)
        cropped_img.save(os.path.join(save_dir, filename))
        
        try: img.close()
        except: pass
        
        log_msg(f"截圖已保存: {filename}")
        return 0
    except Exception as e:
        log_msg(f"截圖失敗: {e}")
        return 1

def auto_screenshot(user_id, stock_id):
    global driver, session_results, user_name_map, shot_speed, name_source_mode
    try: driver.implicitly_wait(0.01)
    except: pass
    
    if user_id not in session_results:
        session_results[user_id] = {'fail_screenshot': [], 'success_screenshot': []}
    if 'success_screenshot' not in session_results[user_id]:
        session_results[user_id]['success_screenshot'] = []

    base_wait = 0.1 * shot_speed
    try:
        log_msg(f"搜尋股票: {stock_id}")
        pass_active_form()
        if "tc_estock_welshas.html" not in driver.current_url:
            driver.get("https://stockservices.tdcc.com.tw/evote/shareholder/000/tc_estock_welshas.html")
        
        try: driver.execute_script("document.body.style.zoom = '100%'")
        except: pass

        for _ in range(100):
            try:
                driver.find_element(By.NAME,'qryStockId')
                break
            except: time.sleep(base_wait)
            
        driver.find_element(By.NAME,'qryStockId').clear()
        driver.find_element(By.NAME,'qryStockId').send_keys(stock_id)

        search_btn = driver.find_element(By.CSS_SELECTOR,'a[onclick="qryByStockId();"]')
        try:
            search_btn.click() 
        except:
            driver.execute_script("arguments[0].click();", search_btn)

        found_result = False
        row_status_text = "" 
        for _ in range(100): 
            time.sleep(base_wait)
            try:
                rows = driver.find_elements(By.TAG_NAME,'tr')
                if len(rows) > 1 and str(stock_id) in rows[1].text:
                    found_result = True
                    row_status_text = rows[1].text 
                    break
            except: pass
            
        if not found_result:
            log_msg(f"找不到代號: {stock_id}")
            session_results[user_id]['fail_screenshot'].append(stock_id)
            return 2

        if "未投票" in row_status_text:
            log_msg(f"[{stock_id}] 狀態為「未投票」，跳過截圖任務。")
            session_results[user_id]['fail_screenshot'].append(f"{stock_id} (未投票)")
            return 2

        voteinfo = []
        try:
            row = driver.find_elements(By.TAG_NAME,'tr')[1]
            parts = row.text.split(" ")
            voteinfo.extend(parts[0:2])
            report_text = f"{voteinfo[0]} {voteinfo[1]}".strip() if len(voteinfo) > 1 else stock_id

            # ================= 【第一道防線：列表頁提前比對】 =================
            import glob
            stock_id_str = voteinfo[0]
            stock_name_str = clean_filename(voteinfo[1].replace('*',''))
            if len(stock_name_str) > 20: stock_name_str = stock_name_str[:20]

            # 抓取並清理當前 UI 帳號名稱
            current_user_name = clean_filename(user_name_map.get(user_id, str(user_id)))
            safe_display_name = current_user_name if len(current_user_name) < 10 else current_user_name[:10]

            pattern_mode1 = os.path.join(base_path, current_user_name, f"*_{stock_id_str}_{stock_name_str}.png")
            pattern_mode2 = os.path.join(base_path, f"*_{stock_id_str}_{stock_name_str}_{safe_display_name}.png")

            if glob.glob(pattern_mode1) or glob.glob(pattern_mode2):
                log_msg(f"[{stock_id_str}] [{current_user_name}] 圖片已存在，提前跳過截圖。")
                session_results[user_id]['success_screenshot'].append(f"{report_text} (已存在跳過)")
                return 0
            # ===========================================================================

            page_loaded = False
            for attempt in range(5):
                try:
                    if driver.find_elements(By.CSS_SELECTOR, 'button[onclick*="back"], input[onclick*="back"]'):
                        page_loaded = True
                        break
                    
                    msg_btns = driver.find_elements(By.ID, "msgDialog_okBtn")
                    if msg_btns and msg_btns[0].is_displayed():
                        msg_btns[0].click()
                        time.sleep(base_wait)
                except: pass

                try:
                    current_row = driver.find_elements(By.TAG_NAME,'tr')[1]
                    links = current_row.find_elements(By.TAG_NAME,'a')
                    target_link = None
                    for link in links:
                        if "查詢" in link.text:
                            target_link = link
                            break
                    if target_link:
                        if attempt > 0: log_msg(f"第 {attempt+1} 次嘗試進入頁面...")
                        target_link.click()
                        check_limit = 100 
                        for _ in range(check_limit):
                            if driver.find_elements(By.CSS_SELECTOR, 'button[onclick*="back"], input[onclick*="back"]'):
                                page_loaded = True
                                break
                            
                            msg_btns = driver.find_elements(By.ID, "msgDialog_okBtn")
                            if msg_btns and msg_btns[0].is_displayed():
                                try: msg_btns[0].click()
                                except: pass
                                
                            time.sleep(base_wait)
                        if page_loaded: break 
                    else:
                        time.sleep(base_wait * 2)
                        continue
                except: time.sleep(base_wait * 2)
            
            if not page_loaded:
                try:
                    nav_btns = driver.find_elements(By.XPATH, "//button | //a | //input[@type='button']")
                    for btn in nav_btns:
                        if not btn.is_displayed(): continue
                        txt = btn.text.strip()
                        val = btn.get_attribute("value")
                        check_str = (txt + str(val)).strip()
                        if "返回" in check_str or "上一頁" in check_str or "列表" in check_str:
                            page_loaded = True
                            break
                except: pass

            if not page_loaded:
                log_msg(f"[{stock_id}] 進入內頁失敗，跳過截圖")
                session_results[user_id]['fail_screenshot'].append(report_text)
                return 2

            try: driver.execute_script("document.body.style.zoom = '100%'")
            except: pass

            actual_display_name = user_name_map.get(user_id, str(user_id))

            if name_source_mode == 2: 
                detected_name = ""
                start_search = time.time()
                while time.time() - start_search < 3.0:
                    try:
                        btn = driver.find_element(By.ID, "msgDialog_okBtn")
                        if btn.is_displayed(): btn.click()
                    except: pass
                    try: driver.execute_script("document.body.style.zoom = '100%'")
                    except: pass
                    
                    try:
                        try:
                            name_td = driver.find_element(By.XPATH, "//th[contains(text(), '戶名')]/following-sibling::td")
                            if name_td and name_td.text.strip():
                                detected_name = name_td.text.strip()
                        except: pass

                        if not detected_name:
                            try:
                                trs = driver.find_elements(By.TAG_NAME, "tr")
                                for tr in trs:
                                    if "戶名" in tr.text:
                                        tds = tr.find_elements(By.TAG_NAME, "td")
                                        if tds:
                                            detected_name = tds[-1].text.strip()
                                            break
                            except: pass

                        if not detected_name:
                            body_text = driver.find_element(By.TAG_NAME, "body").text
                            match = re.search(r'戶名\s*[:：]?\s*([^\s　]+)', body_text)
                            if match:
                                potential_name = match.group(1)
                                if 2 <= len(potential_name) <= 20:
                                    detected_name = potential_name

                        if detected_name:
                            clean_name = clean_filename(detected_name)
                            actual_display_name = clean_name
                            log_msg(f"成功抓到實際戶名: {clean_name}")
                            break
                    except: pass
                    time.sleep(base_wait)
            else:
                for _ in range(10):
                    try:
                        btn = driver.find_element(By.ID, "msgDialog_okBtn")
                        if btn.is_displayed():
                            btn.click()
                            break 
                    except: pass
                    time.sleep(base_wait)

            # ================= 【第二道防線：抓到實際戶名後再次比對】 =================
            # 解決「網頁實際戶名」與「設定頁名稱」不同導致第一道防線漏掉的問題
            safe_actual_name = actual_display_name if len(actual_display_name) < 10 else actual_display_name[:10]
            pattern_mode1_real = os.path.join(base_path, actual_display_name, f"*_{stock_id_str}_{stock_name_str}.png")
            pattern_mode2_real = os.path.join(base_path, f"*_{stock_id_str}_{stock_name_str}_{safe_actual_name}.png")

            if glob.glob(pattern_mode1_real) or glob.glob(pattern_mode2_real):
                log_msg(f"[{stock_id_str}] [{actual_display_name}] 圖片已存在，跳過重複截圖。")
                session_results[user_id]['success_screenshot'].append(f"{report_text} (已存在跳過)")
                try: driver.execute_script("arguments[0].click();", driver.find_element(By.CSS_SELECTOR,'button[onclick="back(); return false;"]'))
                except: pass
                return 0
            # ===========================================================================
            voteinfo.append("unknown") 
            res = screenshot(user_id, voteinfo, override_name=actual_display_name) 
            if res != 0: session_results[user_id]['fail_screenshot'].append(report_text)
            else: session_results[user_id]['success_screenshot'].append(report_text)
            
            try: driver.execute_script("arguments[0].click();", driver.find_element(By.CSS_SELECTOR,'button[onclick="back(); return false;"]'))
            except: pass
            return res
        except Exception as e:
            log_msg(f"流程錯誤: {e}")
            fail_record = locals().get('report_text', stock_id)
            session_results[user_id]['fail_screenshot'].append(fail_record)
            return 1
    except: return 1

def generate_session_report(start_t=None, end_t=None, count=0):
    global execution_logs, shot_speed, login_type 
    try:
        log_dir = "Log"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
            
        file_name = f"{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}_單筆補圖任務報告.txt"
        report_name = os.path.join(log_dir, file_name)
        
        def chunks(lst, n):
            for i in range(0, len(lst), n):
                yield lst[i:i + n]

        with open(report_name, "w", encoding="utf-8") as f:
            f.write("=== 股東e票通 - 單筆補圖專用版任務報告 ===\n")
            f.write(f"時間: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            
            f.write(f"【執行環境設定】\n")
            f.write(f"  - 登入方式: 依各帳號設定獨立切換\n")
            f.write(f"  - 截圖等待速率: {shot_speed}\n")

            if start_t and end_t:
                total_seconds = end_t - start_t
                avg_seconds = total_seconds / count if count > 0 else 0
                f.write(f"總耗時: {total_seconds:.2f} 秒\n")
                f.write(f"處理家數: {count}\n")
                f.write(f"平均每間耗時: {avg_seconds:.2f} 秒\n")
            f.write("\n")
            
            for uid, res in session_results.items():
                d_name = user_name_map.get(uid, uid)
                f.write(f"【帳號: {d_name} ({uid})】\n")
                
                succ_shots = res.get('success_screenshot', [])
                f.write(f"  - 截圖成功: {len(succ_shots)}\n")
                if succ_shots:
                    for chunk in chunks(succ_shots, 5):
                        f.write(f"    內容: {', '.join(chunk)}\n")
                        
                fail_shots = res.get('fail_screenshot', [])
                f.write(f"  - 截圖失敗/跳過: {len(fail_shots)}\n")
                if fail_shots: 
                    for chunk in chunks(fail_shots, 5):
                        f.write(f"    內容: {', '.join(chunk)}\n")
                f.write("-" * 30 + "\n")
            
            f.write("\n\n==========================================\n")
            f.write("           詳細執行歷程          \n")
            f.write("==========================================\n")
            for line in execution_logs:
                f.write(line + "\n")
                
        execution_logs.clear()
        log_msg(f"報告與詳細歷程已產生: {report_name}")
    except Exception as e:
        log_msg(f"報告產生失敗: {e}")

class TextRedirector(object):
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag
    def write(self, str):
        self.widget.configure(state="normal")
        self.widget.insert("end", str, (self.tag,))
        self.widget.see("end")
        self.widget.configure(state="disabled")
    def flush(self): pass

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("股東e票通 - 獨立補圖專用小幫手")
        
        self.load_config()
        
        if main_window_geom:
            self.geometry(main_window_geom)
        else:
            self.geometry(f"900x600+{int(self.winfo_screenwidth() / 8)}+{int(self.winfo_screenheight() / 10)}")
        
        self.minsize(500, 500)
        self.resizable(True, True)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.style = ttk.Style(self)
        try: self.style.theme_use('clam') 
        except: pass
        
        main_font = ('Microsoft JhengHei', 10)
        bold_font = ('Microsoft JhengHei', 10, 'bold')
        app_bg = '#E8E8E8'
        text_color = '#111111'
        
        self.configure(bg=app_bg) 
        self.style.configure('.', font=main_font, background=app_bg, foreground=text_color)
        self.style.configure('TFrame', background=app_bg)
        self.style.configure('TLabelframe', background=app_bg, font=bold_font, foreground=text_color)
        self.style.configure('TLabelframe.Label', background=app_bg, font=bold_font, foreground=text_color)
        self.style.configure('Action.TButton', font=bold_font, foreground='white', background='#333333', borderwidth=0)
        self.style.map('Action.TButton', background=[('active', '#000000')])
        self.style.configure('Red.TButton', font=bold_font, foreground='white', background='#CC0000', borderwidth=0)
        self.style.map('Red.TButton', background=[('active', '#990000')]) 
        
        self.create_widgets()
        sys.stdout = TextRedirector(self.log_text, "stdout")
        sys.stderr = TextRedirector(self.log_text, "stderr")
        
        log_msg("=== 歡迎使用 單筆補圖專用小幫手 ===")
        log_msg("本工具僅保留「登入後針對指定代號進行截圖」之功能。")
        
    def on_closing(self):
        global main_window_geom
        main_window_geom = self.geometry()
        self.save_config()
        self.destroy()

    def load_config(self):
        global user_accounts, shot_speed, screenshot_mode, login_type, main_window_geom, join_draw, name_source_mode
        conf_path = os.path.join(CONFIG_DIR, 'program_setting.conf')
        
        if os.path.exists(conf_path):
            try:
                with open(conf_path, 'r', encoding='utf8') as f:
                    content = f.read()
                    for line in content.split('\n'):
                        if "screenshot_mode:::" in line: screenshot_mode = int(line.split(":::")[1])
                        if "name_source_mode:::" in line: name_source_mode = int(line.split(":::")[1])
                        if "shot_speed:::" in line: shot_speed = float(line.split(":::")[1])
                        if "login_type:::" in line: login_type = line.split(":::")[1]
                        if "main_window_geom:::" in line: main_window_geom = line.split(":::")[1].strip()
                        if "join_draw:::" in line: join_draw = (line.split(":::")[1].strip() == 'True')
                        if "user_accounts:::" in line:
                            try:
                                dec = decrypt_data(line.split(":::")[1])
                                if dec: user_accounts = json.loads(dec)
                            except: pass
                        if "shareholderIDs:::" in line and not user_accounts: 
                            try:
                                dec = decrypt_data(line.split(":::")[1])
                                if dec: 
                                    old_ids = dec.split("|/|")
                                    user_accounts = [{'name': f'帳號{i+1}', 'id': x.strip(), 'login_type': login_type} for i, x in enumerate(old_ids) if x.strip()]
                            except: pass
            except: pass

        self.shot_speed_var = tk.StringVar(value=str(shot_speed))
        self.screenshot_mode_var = tk.IntVar(value=screenshot_mode)
        self.name_source_mode_var = tk.IntVar(value=name_source_mode)
        self.join_draw_var = tk.BooleanVar(value=join_draw) 

    def save_config(self):
        global shot_speed, screenshot_mode, login_type, main_window_geom, join_draw, name_source_mode
        try: main_window_geom = self.geometry()
        except: pass

        try:
            try: s_val = float(self.shot_speed_var.get()); shot_speed = s_val 
            except ValueError:
                messagebox.showerror("設定錯誤", "速度倍率請輸入有效的數字")
                return
            screenshot_mode = self.screenshot_mode_var.get()
            name_source_mode = self.name_source_mode_var.get()
            join_draw = self.join_draw_var.get() 
                
            conf_path = os.path.join(CONFIG_DIR, 'program_setting.conf')
            
            with open(conf_path, 'w', encoding='utf8') as f:
                f.write(f"screenshot_mode:::{screenshot_mode}\n")
                f.write(f"name_source_mode:::{name_source_mode}\n")
                f.write(f"shot_speed:::{shot_speed}\n")
                f.write(f"login_type:::{login_type}\n")
                f.write(f"main_window_geom:::{main_window_geom}\n")
                f.write(f"join_draw:::{join_draw}\n")
                encrypted_acc = encrypt_data(json.dumps(user_accounts, ensure_ascii=False))
                f.write(f"user_accounts:::{encrypted_acc}\n")
                f.write("hash:::SECURE_ENCRYPTED_V4\n")
            log_msg("設定已儲存。")
        except Exception as e: log_msg(f"儲存設定失敗: {e}")

    def build_user_checklist(self, parent, vars_dict, select_all=True):
        container = ttk.Frame(parent)
        container.pack(fill="both", expand=True)

        select_all_var = tk.BooleanVar(value=select_all)
        def toggle_all():
            state = select_all_var.get()
            for v in vars_dict.values(): v.set(state)

        tk.Checkbutton(container, text="全選", variable=select_all_var, command=toggle_all, bg='#E8E8E8', activebackground='#E8E8E8', selectcolor='#FFFFFF').pack(anchor="w", pady=(0, 2))

        base_frame = tk.Frame(container, bg='#E8E8E8')
        base_frame.pack(fill="both", expand=True)

        scrollbar = ttk.Scrollbar(base_frame, orient="vertical")
        canvas = tk.Canvas(base_frame, height=30, bg='#E8E8E8', highlightthickness=0, yscrollcommand=scrollbar.set)
        scrollbar.config(command=canvas.yview)
        
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame._canvas = canvas
        scrollable_frame._scrollbar = scrollbar

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        canvas.bind('<Enter>', lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        canvas.bind('<Leave>', lambda e: canvas.unbind_all("<MouseWheel>"))

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack_forget()
        return scrollable_frame

    def refresh_user_lists(self):
        for w in self.single_scroll_frame.winfo_children(): w.destroy()
        self.check_vars_single.clear()

        max_cols = 5 
        col_idx = 0
        row_idx = 0

        for acc in user_accounts:
            uid = acc['id']
            disp_text = f"{acc['name']}"

            self.check_vars_single[uid] = tk.BooleanVar(value=True)

            tk.Checkbutton(self.single_scroll_frame, text=disp_text, variable=self.check_vars_single[uid], bg='#E8E8E8', activebackground='#E8E8E8', selectcolor='#FFFFFF').grid(row=row_idx, column=col_idx, padx=(0, 15), pady=1, sticky="w")
            
            col_idx += 1
            if col_idx >= max_cols:
                col_idx = 0
                row_idx += 1

        def update_scroll_visibility(frame):
            frame.update_idletasks() 
            canvas = frame._canvas
            scrollbar = frame._scrollbar
            req_h = frame.winfo_reqheight()
            
            if req_h <= 35:
                new_h = req_h if req_h > 10 else 32
                scrollbar.pack_forget()
            elif req_h <= 65:
                new_h = req_h
                scrollbar.pack_forget()
            else:
                new_h = 65
                scrollbar.pack(side="right", fill="y")
            canvas.config(height=new_h)

        update_scroll_visibility(self.single_scroll_frame)

        for item in self.user_tree.get_children(): self.user_tree.delete(item)
        for acc in user_accounts:
            self.user_tree.insert("", "end", values=(acc['name'], acc['id'], acc['login_type']))
            
    def add_or_update_user(self):
        name = self.entry_uname.get().strip()
        uid = self.entry_uid.get().strip().upper()
        login = self.combo_ulogin.get().strip()
        
        if not name or not uid:
            messagebox.showwarning("提示", "姓名與身分證不能空白！")
            return
        
        for acc in user_accounts:
            if acc['id'] != uid and acc['name'] == name:
                messagebox.showerror("錯誤", f"姓名「{name}」已存在於其他帳號中，請使用不同的名稱！")
                return

        target_idx = -1
        for i, acc in enumerate(user_accounts):
            if acc['id'] == uid:
                target_idx = i
                break
                
        if target_idx != -1:
            old_name = user_accounts[target_idx]['name']
            if not messagebox.askyesno("重複確認", f"身分證「{uid}」已經存在 (原姓名：{old_name})。\n\n請問是否要覆蓋更新此帳號的資料？"):
                return 
            user_accounts[target_idx]['name'] = name
            user_accounts[target_idx]['login_type'] = login
            log_msg(f"✅ 已更新帳號資料: {name} ({uid})")
        else:
            user_accounts.append({'name': name, 'id': uid, 'login_type': login})
            log_msg(f"✅ 已新增帳號: {name} ({uid})")
            
        self.save_config()
        self.refresh_user_lists()
        self.entry_uname.delete(0, tk.END)
        self.entry_uid.delete(0, tk.END)
        
    def delete_selected_user(self):
        selected = self.user_tree.selection()
        if not selected: return
        
        uid = str(self.user_tree.item(selected[0])['values'][1])
        global user_accounts
        user_accounts = [acc for acc in user_accounts if str(acc['id']) != uid]
        
        self.save_config()
        self.refresh_user_lists()
        log_msg(f"🗑️ 已刪除帳號: {uid}")

    def _adj_val(self, var, delta):
        try:
            curr_val = float(var.get()) if var.get() else 1.0
            new_val = round(curr_val + delta, 1)
            if new_val < 1.0: new_val = 0.5 # 修改這裡：將最小值限制為 1.0
            var.set(str(new_val))
        except ValueError: pass

    def create_widgets(self):
        main_container = ttk.Frame(self, padding="10")
        main_container.pack(fill="both", expand=True)
        
        self.pane = ttk.PanedWindow(main_container, orient=tk.HORIZONTAL)
        self.pane.pack(fill="both", expand=True)
        
        left_frame = ttk.Frame(self.pane)
        right_frame = ttk.Frame(self.pane)
        
        self.pane.add(left_frame, weight=1)
        self.pane.add(right_frame, weight=1)
               
        right_content = ttk.Frame(right_frame, padding=(10, 0, 0, 0))
        right_content.pack(fill="both", expand=True)

        tab_control = ttk.Notebook(left_frame)
        
        tab1 = ttk.Frame(tab_control, padding="10")
        tab2 = ttk.Frame(tab_control, padding="10") 
        tab3 = ttk.Frame(tab_control, padding="10")
        
        tab_control.add(tab1, text='  補圖任務  ')
        tab_control.add(tab2, text='  設定  ')
        tab_control.add(tab3, text='  系統資訊  ') 

        tab_control.pack(side="top", expand=True, fill="both")

        # === 右側狀態區 ===
        self.frame_log = ttk.LabelFrame(right_content, text=" 執行狀態 ")
        self.frame_log.pack(fill="both", expand=True)
        self.log_text = scrolledtext.ScrolledText(self.frame_log, width=50, state='disabled', font=('Consolas', 9), bg='#ffffff', fg='#333333')
        self.log_text.pack(expand=True, fill="both", padx=5, pady=5)
        
        # === Tab 1: 補圖任務 ===
        self.check_vars_single = {}

        frame_mode2 = ttk.LabelFrame(tab1, text=" 單筆/多筆補圖工具 ")
        frame_mode2.pack(fill="x", pady=5, ipady=5)
        
        ttk.Label(frame_mode2, text="勾選欲執行帳號:").pack(anchor="w", padx=10)
        list_container2 = ttk.Frame(frame_mode2)
        list_container2.pack(fill="x", padx=10, pady=2)
        self.single_scroll_frame = self.build_user_checklist(list_container2, self.check_vars_single)

        grid_frame = ttk.Frame(frame_mode2)
        grid_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(grid_frame, text="股票代號:\n(可貼上Excel欄位\n或用逗號區隔)").grid(row=0, column=0, padx=5, pady=5, sticky="nw")
        text_frame = ttk.Frame(grid_frame)
        text_frame.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        self.stock_list_entry = tk.Text(text_frame, width=20, height=8, font=('Microsoft JhengHei', 10))
        self.stock_list_entry.pack(side="left", fill="both", expand=True)
        scrollbar1 = ttk.Scrollbar(text_frame, orient="vertical", command=self.stock_list_entry.yview)
        self.stock_list_entry.configure(yscrollcommand=scrollbar1.set)
        
        def auto_hide_scrollbar1(event):
            if float(self.stock_list_entry.index('end-1c').split('.')[0]) > 8:
                scrollbar1.pack(side="right", fill="y")
            else:
                scrollbar1.pack_forget()
        self.stock_list_entry.bind('<KeyRelease>', auto_hide_scrollbar1)

        grid_frame.columnconfigure(1, weight=1)
        
        btn_mode2 = ttk.Button(frame_mode2, text="啟動補圖任務", style='Action.TButton', command=self.start_screenshot_task, cursor="hand2")
        btn_mode2.pack(fill="x", padx=15, pady=15)
        
        # === Tab 2: 設定 ===
        frame_setting = ttk.Frame(tab2)
        frame_setting.pack(fill="both", expand=True, padx=10, pady=5)
        
        user_frame = ttk.LabelFrame(frame_setting, text=" 帳號明細管理 ")
        user_frame.pack(fill="x", pady=2)

        add_frame = ttk.Frame(user_frame)
        add_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(add_frame, text="姓名:").grid(row=0, column=0, padx=2, pady=5, sticky="e")
        self.entry_uname = ttk.Entry(add_frame, width=12)
        self.entry_uname.grid(row=0, column=1, padx=2, pady=5, sticky="we")
        
        ttk.Label(add_frame, text="身分證:").grid(row=0, column=2, padx=2, pady=5, sticky="e")
        self.entry_uid = ttk.Entry(add_frame, width=15)
        self.entry_uid.grid(row=0, column=3, padx=2, pady=5, sticky="we")
        
        ttk.Label(add_frame, text="憑證:").grid(row=1, column=0, padx=2, pady=5, sticky="e")
        login_types = ["券商網路下單憑證", "自然人憑證", "行動自然人憑證"]
        self.combo_ulogin = ttk.Combobox(add_frame, values=login_types, state="readonly", width=18)
        self.combo_ulogin.set("券商網路下單憑證")
        self.combo_ulogin.grid(row=1, column=1, columnspan=2, padx=2, pady=5, sticky="we")
        
        ttk.Button(add_frame, text="儲存", command=self.add_or_update_user).grid(row=1, column=3, padx=5, pady=5, sticky="e")
        
        add_frame.columnconfigure(1, weight=1)
        add_frame.columnconfigure(3, weight=1)

        tree_frame = ttk.Frame(user_frame)
        tree_frame.pack(fill="x", padx=5, pady=2)
        
        cols = ("name", "id", "login")
        self.user_tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=4)
        self.user_tree.heading("name", text="姓名")
        self.user_tree.heading("id", text="身分證")
        self.user_tree.heading("login", text="登入方式")
        self.user_tree.column("name", width=60)
        self.user_tree.column("id", width=100)
        self.user_tree.column("login", width=120)
        
        tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.user_tree.yview)
        self.user_tree.configure(yscrollcommand=tree_scroll.set)
        self.user_tree.pack(side="left", fill="x", expand=True)
        
        original_refresh = self.refresh_user_lists
        def wrapped_refresh():
            original_refresh()
            if len(self.user_tree.get_children()) > 4: tree_scroll.pack(side="right", fill="y")
            else: tree_scroll.pack_forget()
        self.refresh_user_lists = wrapped_refresh

        ttk.Button(user_frame, text="刪除選取帳號", command=self.delete_selected_user, style='Red.TButton').pack(anchor="e", padx=5, pady=(0,5))

        spd_frame = ttk.Frame(frame_setting)
        spd_frame.pack(fill="x", pady=5)
        
        s_frame = ttk.Frame(spd_frame)
        s_frame.pack(fill="x", pady=2)
        ttk.Label(s_frame, text="截圖等待速度：", width=12).pack(side="left")
        ttk.Button(s_frame, text="-", width=2, command=lambda: self._adj_val(self.shot_speed_var, -0.1)).pack(side="left", fill="y", padx=2)
        ttk.Entry(s_frame, textvariable=self.shot_speed_var, width=5, justify='center').pack(side="left", fill="y", ipady=2, padx=2)
        ttk.Button(s_frame, text="+", width=2, command=lambda: self._adj_val(self.shot_speed_var, 0.1)).pack(side="left", fill="y", padx=2)

        file_frame = ttk.Frame(frame_setting)
        file_frame.pack(fill="x", pady=2)
        ttk.Label(file_frame, text="截圖存檔方式:").pack(anchor="w", pady=(0,2))
        ttk.Radiobutton(file_frame, text="各帳號獨立資料夾", variable=self.screenshot_mode_var, value=1).pack(side="left", padx=(10, 5))
        ttk.Radiobutton(file_frame, text="放一起(檔名含使用者名稱)", variable=self.screenshot_mode_var, value=2).pack(side="left")
        
        name_frame = ttk.Frame(frame_setting)
        name_frame.pack(fill="x", pady=2)
        ttk.Label(name_frame, text="使用者名稱來源:").pack(anchor="w", pady=(0,2))
        ttk.Radiobutton(name_frame, text="程式自訂名稱", variable=self.name_source_mode_var, value=1).pack(side="left", padx=(10, 5))
        ttk.Radiobutton(name_frame, text="網頁實際戶名", variable=self.name_source_mode_var, value=2).pack(side="left")

        draw_frame = ttk.Frame(frame_setting)
        draw_frame.pack(fill="x", pady=5)
        tk.Checkbutton(draw_frame, text="遇到抽獎頁面時，暫停 5 分鐘讓我手動參加抽獎", variable=self.join_draw_var, bg='#E8E8E8', activebackground='#E8E8E8', font=('Microsoft JhengHei', 10), wraplength=360).pack(anchor="w")
        
        ttk.Button(frame_setting, text="儲存全部設定", style='Red.TButton', command=self.save_config, cursor="hand2").pack(pady=10, ipady=1, fill='x')

        # === Tab 3: 系統資訊 ===
        frame_info = ttk.Frame(tab3)
        frame_info.pack(fill="both", expand=True, padx=10, pady=12)

        path_info_frame = ttk.LabelFrame(frame_info, text=" 📂 資料夾路徑 (可全選複製) ")
        path_info_frame.pack(fill="x", pady=(0, 15))

        ttk.Label(path_info_frame, text="設定檔記憶位置:", foreground="#555").pack(anchor="w", padx=10, pady=(5,0))
        path_entry_conf = ttk.Entry(path_info_frame)
        path_entry_conf.insert(0, CONFIG_DIR) 
        path_entry_conf.configure(state="readonly") 
        path_entry_conf.pack(fill="x", padx=10, pady=5)

        ttk.Label(path_info_frame, text="截圖存檔位置:", foreground="#555").pack(anchor="w", padx=10, pady=(5,0))
        path_entry_shot = ttk.Entry(path_info_frame)
        abs_shot_path = os.path.abspath(base_path) 
        path_entry_shot.insert(0, abs_shot_path)
        path_entry_shot.configure(state="readonly")
        path_entry_shot.pack(fill="x", padx=10, pady=(0,10))

        self.refresh_user_lists() 

    def _finish_task(self, task_start_time=None, total_items=0):
        task_end_time = time.time() if task_start_time else None
        generate_session_report(task_start_time, task_end_time, total_items)

        total_screenshots = sum(len(res.get('success_screenshot', [])) for res in session_results.values())
        if total_screenshots > 0:
            self.after(1000, self._open_folder_and_notify)
        else:
            log_msg("本次任務無截圖紀錄。")
            self._pop_topmost_message("任務結束！報告已經產生！\n\n(提示：本次無成功截圖紀錄)")
            
    def _open_folder_and_notify(self):
        abs_path = os.path.abspath(base_path)
        log_msg("正在為您開啟截圖資料夾...")
        try: subprocess.Popen(f'explorer "{abs_path}"')
        except Exception as e: log_msg(f"開啟截圖資料夾失敗: {e}")
        self.after(1200, self._resize_folder_and_notify, abs_path)

    def _resize_folder_and_notify(self, abs_path):
        import ctypes
        try:
            folder_name = os.path.basename(abs_path)
            hwnd = ctypes.windll.user32.FindWindowW("CabinetWClass", folder_name)
            if not hwnd: hwnd = ctypes.windll.user32.FindWindowW(None, folder_name)

            if hwnd:
                sw = ctypes.windll.user32.GetSystemMetrics(0)
                sh = ctypes.windll.user32.GetSystemMetrics(1)
                w, h = sw // 2, sh // 2
                x, y = 50, 50 
                ctypes.windll.user32.SetWindowPos(hwnd, -1, x, y, w, h, 0x0040)
                ctypes.windll.user32.SetForegroundWindow(hwnd)
                self.after(500, lambda: ctypes.windll.user32.SetWindowPos(hwnd, -2, x, y, w, h, 0x0040))
        except Exception as e:
            log_msg(f"資料夾視窗自動排版失敗: {e}")

        self.after(1000, self._pop_topmost_message, "任務搞定！報告已經產生！\n\n資料夾已為您開啟。")

    def _pop_topmost_message(self, msg):
        self.attributes('-topmost', True)
        self.update()
        self.attributes('-topmost', False)
        
        fail_count = sum(len(res.get('fail_screenshot', [])) for res in session_results.values())

        if fail_count > 0:
            msg = msg.replace("任務搞定！", "任務執行結束！")
            warning_msg = f"{msg}\n\n⚠️ 注意：偵測到 {fail_count} 筆截圖失敗(或略過)！\n👉 請查看 LOG 報告檔確認詳細原因。"
            messagebox.showwarning("任務結束 (有部分失敗)", warning_msg)
        else:
            messagebox.showinfo("完成", msg)

        self.on_closing()

    def start_screenshot_task(self):
        self.save_config()  
        target_accounts = [acc for acc in user_accounts if self.check_vars_single.get(acc['id'], tk.BooleanVar()).get()]
        if not target_accounts: return messagebox.showwarning("提示", "請勾選至少一個帳號！")
        stocks = self.stock_list_entry.get("1.0", tk.END).strip()
        if not stocks: return messagebox.showwarning("錯誤", "請輸入股票代號")
        
        self.log_text.configure(state='normal'); self.log_text.delete(1.0, tk.END); self.log_text.configure(state='disabled')
        threading.Thread(target=self.run_screenshot_task, args=(target_accounts, stocks), daemon=True).start()

    def run_screenshot_task(self, target_accounts, stocks_str):
        global driver, session_results, user_name_map
        session_results = {}; user_name_map = {} 
        start_time = time.time(); total_items = 0
        
        log_msg(f"=== 開始補圖，共 {len(target_accounts)} 個帳號 ===")
        stock_list = re.findall(r'\d+', stocks_str)
        if not stock_list:
            log_msg("沒有有效的股票代號")
            return
            
        maintenance_flag = False
        try:
            if driver is None: driver = get_driver()
        except: 
            log_msg("瀏覽器啟動失敗"); return

        for acc in target_accounts: user_name_map[acc['id']] = acc['name']

        try:
            for acc in target_accounts:
                target_id = acc['id']
                current_login_type = acc['login_type']

                if maintenance_flag: break
                try: driver.current_url
                except:
                    log_msg("瀏覽器意外關閉，重啟中...")
                    try: driver = get_driver()
                    except: continue

                try:
                    log_msg(f"--- 補圖帳號: {target_id} (使用: {current_login_type}) ---")
                    try: 
                        autoLogin(target_id, current_login_type)
                    except SystemMaintenanceError:
                        log_msg("系統維護中，停止任務")
                        maintenance_flag = True; break
                    except LoginTimeoutError: 
                        log_msg("登入逾時，將重啟瀏覽器...")
                        session_results.setdefault(target_id, {'fail_screenshot': [], 'success_screenshot': []})
                        session_results[target_id]['fail_screenshot'].append("登入失敗")
                        force_quit_driver(driver)
                        driver = None
                        continue

                    for stock_id in stock_list:
                        total_items += 1
                        auto_screenshot(target_id, stock_id)
                    
                    logout()
                except Exception as e:
                    log_msg(f"{target_id} 執行失敗: {e}")
                    force_quit_driver(driver)
                    driver = None
        finally:
            force_quit_driver(driver)
            driver = None

        log_msg("=== 補圖任務結束 ===")
        self.after(0, lambda: self._finish_task(start_time, total_items))

if __name__ == "__main__":
    app = App()
    app.mainloop()
