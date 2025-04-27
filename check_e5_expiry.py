#!/usr/bin/python3
# -*- coding: utf8 -*-
"""
说明: 
- 此脚本使用Selenium自动登录Microsoft账号。
- 登录成功后，导航到指定的OAuth URL以获取授权码。
- 将包含授权码的重定向URL通过 OneDriveUploader 上传到 OneDrive。
- 在GitHub Actions上运行时，浏览器和驱动程序由工作流安装。
- 环境变量 `MS_E5_ACCOUNTS` 从 GitHub Secrets 读取: email-password&email2-password2...
- (可选) `sendNotify.py` 用于发送通知，需要配置相应的 Secrets。
"""
import os
import time
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
from urllib.parse import urlparse, parse_qs
import subprocess

# --- Optional Notification Setup ---
# Ensure sendNotify.py is in your repository if you use this
try:
    from sendNotify import send
except ImportError:
    print("通知文件 sendNotify.py 未找到，将仅打印到控制台。")
    def send(title, content):
        print(f"--- {title} ---")
        print(content)
        print("--- End Notification ---")
# --- End Notification Setup ---

List = [] # To store output messages

# --- Configuration ---
LOGIN_URL = 'https://admin.microsoft.com/' # Use admin center for login context initially
OAUTH_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=78d4dc35-7e46-42c6-9023-2d39314433a5&response_type=code&redirect_uri=http://localhost/onedrive-login&response_mode=query&scope=offline_access%20User.Read%20Files.ReadWrite.All"
REDIRECT_URI_START = "http://localhost/onedrive-login" 
ONEDRIVE_UPLOADER = "/usr/local/bin/OneDriveUploader"  # Path to OneDriveUploader
ONEDRIVE_AUTH_CONFIG = "auth1106.json"  # OneDriveUploader auth config file

# --- Helper Function ---
def get_webdriver():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless") 
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage") 
    options.add_argument("--window-size=1920,1080") 
    options.add_argument("user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36") 
    
    # Specify the path to the chromium binary installed by apt
    options.binary_location = "/usr/bin/chromium-browser" 

    try:
       driver = webdriver.Chrome(options=options) 
       List.append("  - WebDriver 初始化成功 (使用 /usr/bin/chromium-browser)。") 
       return driver
    except WebDriverException as e:
       if "no chrome binary at /usr/bin/chromium-browser" in str(e) or "cannot find chrome binary" in str(e).lower():
            List.append(f"!! 错误：无法在指定位置 /usr/bin/chromium-browser 找到 Chromium。尝试 /usr/bin/chromium ...")
            options.binary_location = "/usr/bin/chromium" # Try alternative common path
            try:
                driver = webdriver.Chrome(options=options)
                List.append("  - WebDriver 初始化成功 (使用 /usr/bin/chromium)。") 
                return driver
            except WebDriverException as e2:
                 List.append(f"!! 错误：在 /usr/bin/chromium 也无法初始化 WebDriver: {e2}")
                 return None
            except Exception as e3:
                 List.append(f"!! 尝试 /usr/bin/chromium 时发生意外错误: {e3}")
                 return None
       List.append(f"!! 错误：无法初始化 WebDriver: {e}")
       return None
    except Exception as e:
       List.append(f"!! 错误：初始化WebDriver时发生意外错误: {e}")
       return None

# --- Function to perform login and get OAuth code ---
def get_oauth_code(username, password):
    """Logs into Microsoft account and attempts to capture OAuth redirect URL."""
    List.append(f"开始处理账号: {username}")
    driver = get_webdriver()
    if not driver:
        List.append(f"!! 处理失败: {username} (WebDriver 初始化失败)")
        return 

    try:
        # --- Login Steps (same as original script) ---
        driver.get(LOGIN_URL)
        wait = WebDriverWait(driver, 60) 

        # Step 1: Enter Email
        email_field = wait.until(EC.visibility_of_element_located((By.ID, "i0116")))
        email_field.send_keys(username)
        next_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
        driver.execute_script("arguments[0].click();", next_button)
        time.sleep(random.uniform(4, 6)) 

        # Step 2: Enter Password
        password_field = wait.until(EC.visibility_of_element_located((By.ID, "i0118")))
        time.sleep(0.7)
        password_field.send_keys(password)
        signin_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
        driver.execute_script("arguments[0].click();", signin_button)

        # Step 3: Handle "Stay signed in?" (KMSI)
        try:
            kmsi_button_no = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.ID, "idBtn_Back"))
            )
            driver.execute_script("arguments[0].click();", kmsi_button_no)
        except TimeoutException:
            pass

        # Step 4: Navigate to OAuth URL
        driver.get(OAUTH_URL)
        time.sleep(3)
        WebDriverWait(driver, 30).until(lambda d: REDIRECT_URI_START in d.current_url)
        redirected_url = driver.current_url

        # Step 5: Upload Redirect URL to OneDrive
        upload_to_onedrive(username, redirected_url)
    except Exception as e:
        List.append(f"!! 处理账号 {username} 时发生意外错误: {e}")
    finally:
        driver.quit()

def upload_to_onedrive(username, content):
    """Uploads the given content to OneDrive using OneDriveUploader."""
    try:
        temp_file = f"oauth_redirect_{username}.txt"
        with open(temp_file, 'w') as file:
            file.write(content)
        
        # Execute OneDriveUploader command
        upload_command = [ONEDRIVE_UPLOADER, '-c', ONEDRIVE_AUTH_CONFIG, '-s', temp_file]
        result = subprocess.run(upload_command, capture_output=True, text=True)
        
        if result.returncode == 0:
            List.append(f"  - 成功上传文件到 OneDrive: {temp_file}")
        else:
            List.append(f"!! 上传到 OneDrive 失败: {result.stderr}")
    except Exception as e:
        List.append(f"!! 上传到 OneDrive 时发生意外错误: {e}")
    finally:
        if os.path.exists(temp_file):
            os.remove(temp_file)

# --- Main Function ---
if __name__ == "__main__":
    accounts = os.getenv('MS_E5_ACCOUNTS', '').split('&')
    if not accounts or accounts == ['']:
        List.append("!! 错误: 未找到环境变量 MS_E5_ACCOUNTS。")
        send("MS OAuth 登录自动化", '\n'.join(List))
        exit(1)

    for account in accounts:
        try:
            username, password = account.split('-')
            get_oauth_code(username, password)
        except ValueError:
            List.append(f"!! 错误: 无效账号配置: {account} (应为 email-password 格式)")

    send("MS OAuth 登录自动化", '\n'.join(List))
