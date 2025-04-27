#!/usr/bin/python3
# -*- coding: utf8 -*-
"""
说明: 
- 此脚本使用Selenium自动登录Microsoft 365 Admin Center并检查E5订阅有效期。
- 如果找不到有效期信息，会打开指定的授权URL，并将返回值保存到文件中。
- 在GitHub Actions上运行时，浏览器和驱动程序由工作流安装。
- 环境变量 `MS_E5_ACCOUNTS` 从 GitHub Secrets 读取: email-password&email2-password2...
"""
import os
import time
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException

# --- Optional Notification Setup ---
try:
    from sendNotify import send
except ImportError:
    print("通知文件 sendNotify.py 未找到，将仅打印到控制台。")
    def send(title, content):
        print(f"--- {title} ---")
        print(content)
        print("--- End Notification ---")

# --- Configuration ---
LOGIN_URL = 'https://admin.microsoft.com/'
SUBSCRIPTIONS_URL = 'https://admin.microsoft.com/Adminportal/Home?source=applauncher#/subscriptions'
AUTH_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=78d4dc35-7e46-42c6-9023-2d39314433a5&response_type=code&redirect_uri=http://localhost/onedrive-login&response_mode=query&scope=offline_access%20User.Read%20Files.ReadWrite.All'
TARGET_SUBSCRIPTION_NAME = "Microsoft 365 E5"
OUTPUT_FILE = 'oauth_response.txt'
List = []  # To store output messages

# --- Helper Function ---
def get_webdriver():
    """
    Initialize the WebDriver with Chrome options.
    """
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36")
    
    try:
        driver = webdriver.Chrome(options=options)  # Selenium will auto-detect ChromeDriver
        List.append("  - WebDriver 初始化成功。")
        return driver
    except WebDriverException as e:
        List.append(f"!! 错误：无法初始化 WebDriver: {e}")
        List.append("!! 请检查工作流中的 ChromeDriver 和 Google Chrome 安装步骤。")
        return None
    except Exception as e:
        List.append(f"!! 错误：初始化 WebDriver 时发生意外错误: {e}")
        return None

def save_browser_response(driver, url, output_file):
    """
    Opens the provided URL and saves the browser's response to a file.
    """
    try:
        driver.get(url)
        time.sleep(5)  # Give some time for the page to load and redirect
        current_url = driver.current_url
        with open(output_file, 'w', encoding='utf-8') as file:
            file.write(current_url)
        List.append(f"  - 授权页面返回值已保存到 {output_file}")
    except Exception as e:
        List.append(f"!! 错误：保存授权页面返回值时发生错误: {e}")

def check_e5_expiry(username, password):
    """Logs into Microsoft Admin Center and checks E5 subscription expiry."""
    List.append(f"开始检查账号: {username}")
    driver = get_webdriver()
    if not driver:
        List.append(f"!! 检查失败: {username} (WebDriver 初始化失败)")
        return 

    try:
        driver.get(LOGIN_URL)
        wait = WebDriverWait(driver, 45)

        # --- Login Step 1: Enter Email ---
        email_field = wait.until(EC.visibility_of_element_located((By.ID, "i0116")))
        email_field.send_keys(username)
        next_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
        driver.execute_script("arguments[0].click();", next_button)
        List.append("  - 输入邮箱并点击下一步")
        time.sleep(random.uniform(3, 5))

        # --- Login Step 2: Enter Password ---
        password_field = wait.until(EC.visibility_of_element_located((By.ID, "i0118")))
        password_field.send_keys(password)
        signin_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
        driver.execute_script("arguments[0].click();", signin_button)
        List.append("  - 输入密码并点击登录")
        time.sleep(random.uniform(3, 5))

        # --- Handle "Stay signed in?" ---
        try:
            kmsi_button_no = wait.until(EC.element_to_be_clickable((By.ID, "idBtn_Back")))
            driver.execute_script("arguments[0].click();", kmsi_button_no)
            List.append("  - 处理 '保持登录状态?' -> 否")
        except TimeoutException:
            List.append("  - 未出现 '保持登录状态?' 继续...")

        # --- Navigate to Subscriptions Page ---
        driver.get(SUBSCRIPTIONS_URL)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[data-is-scrollable='true']")))
        List.append("  - 成功导航到订阅页面")
        time.sleep(random.uniform(2, 4))

        # --- Find E5 Subscription and Expiry Date ---
        subscription_cards = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[role='row'], div[data-automation-id^='DetailsCard']")))
        for card in subscription_cards:
            title_element = card.find_element(By.CSS_SELECTOR, "div[data-automation-id='ProductTitle']")
            if TARGET_SUBSCRIPTION_NAME in title_element.text:
                expiry_element = card.find_element(By.XPATH, ".//*[contains(text(), 'Expires') or contains(text(), '到期')]")
                expiry_text = expiry_element.text.strip()
                List.append(f"  - 找到订阅 '{TARGET_SUBSCRIPTION_NAME}' 有效期信息: {expiry_text}")
                break
        else:
            List.append(f"!! 未找到订阅 '{TARGET_SUBSCRIPTION_NAME}' 或有效期信息。打开授权页面...")
            save_browser_response(driver, AUTH_URL, OUTPUT_FILE)

    except Exception as e:
        List.append(f"!! 发生意外错误: {e}")
    finally:
        driver.quit()
        List.append(f"检查完成: {username}")

if __name__ == '__main__':
    account_env_var = 'MS_E5_ACCOUNTS'
    if account_env_var in os.environ:
        accounts_str = os.environ[account_env_var]
        users = accounts_str.split('&')
        List.append(f'检测到 {len(users)} 个账号配置。')

        for i, user_pair in enumerate(users):
            if '-' in user_pair:
                username, password = user_pair.split('-', 1)
                check_e5_expiry(username.strip(), password.strip())
                time.sleep(random.uniform(8, 15))  # Delay between accounts
            else:
                List.append(f"!! 错误：账号格式不正确，跳过: {user_pair[:10]}...")

        final_output = '\n'.join(List)
        print(final_output)
        send('Microsoft E5 订阅检查报告', final_output)
    else:
        List.append(f"!! 错误：未找到环境变量 {account_env_var}。")
        print('\n'.join(List))
