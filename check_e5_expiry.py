#!/usr/bin/python3
# -*- coding: utf8 -*-
"""
说明: 
- 此脚本使用Selenium自动登录Microsoft账号。
- 登录成功后，导航到指定的OAuth URL以获取授权码。
- 将包含授权码的重定向URL保存到文件 oauth_redirect_{username}.txt。
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

# --- Optional Notification Setup ---
# Ensure sendNotify.py is in your repository if you use this
try:
    from sendNotify import send
    # Add checks relevant to your sendNotify.py implementation if needed
    # Example: if not os.environ.get('PUSH_PLUS_TOKEN'): print("Warning: PUSH_PLUS_TOKEN secret not set.")
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
# The OAuth URL you want to navigate to *after* login
OAUTH_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=78d4dc35-7e46-42c6-9023-2d39314433a5&response_type=code&redirect_uri=http://localhost/onedrive-login&response_mode=query&scope=offline_access%20User.Read%20Files.ReadWrite.All"
REDIRECT_URI_START = "http://localhost/onedrive-login" 

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

    login_successful = False
    try:
        # --- Login Steps (same as before) ---
        driver.get(LOGIN_URL) # Start login via a standard MS page
        wait = WebDriverWait(driver, 60) 

        # Step 1: Enter Email
        try:
            email_field = wait.until(EC.visibility_of_element_located((By.ID, "i0116")))
            email_field.send_keys(username)
            next_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
            driver.execute_script("arguments[0].click();", next_button)
            List.append("  - 输入邮箱并点击下一步")
        except (NoSuchElementException, TimeoutException) as e:
            List.append(f"!! 登录错误 (Email): 找不到邮箱输入框或超时。 {e}")
            driver.save_screenshot(f"error_login_email_{username}.png") 
            return 

        time.sleep(random.uniform(4, 6)) 

        # Step 2: Enter Password
        try:
            password_field = wait.until(EC.visibility_of_element_located((By.ID, "i0118")))
            time.sleep(0.7)
            password_field.send_keys(password)
            signin_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
            driver.execute_script("arguments[0].click();", signin_button)
            List.append("  - 输入密码并点击登录")
        except (NoSuchElementException, TimeoutException) as e:
            try: # Check for common failure patterns
                if driver.find_element(By.ID, "i0118").is_displayed():
                   List.append("!! 登录警告: 似乎仍在密码页面，密码可能错误。")
                elif "login.microsoftonline.com/error" in driver.current_url:
                     List.append("!! 登录错误: 重定向到错误页面，检查凭据。")
                elif "MFA" in driver.page_source or "Proofup" in driver.page_source:
                     List.append("!! 登录错误: 检测到 MFA 或安全信息验证提示，无法自动处理。")
                else:
                     List.append(f"!! 登录错误 (Password): 找不到密码输入框或登录按钮。 {e}")
            except Exception:
                 List.append(f"!! 登录错误 (Password): 找不到密码输入框或登录按钮。 {e}")
            driver.save_screenshot(f"error_login_password_{username}.png")
            return

        # Step 3: Handle "Stay signed in?" (KMSI)
        try:
            kmsi_button_no = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.ID, "idBtn_Back")) # The "No" button
            )
            driver.execute_script("arguments[0].click();", kmsi_button_no)
            List.append("  - 处理 '保持登录状态?' -> 否")
            login_successful = True # Assume login is good if we passed KMSI
        except TimeoutException:
            List.append("  - 未出现 '保持登录状态?' 弹窗 (或已超时)，假定登录成功并继续...")
            # Check if we landed somewhere expected after login (e.g., admin portal)
            time.sleep(2) # allow redirects
            if "admin.microsoft.com" in driver.current_url or "portal.office.com" in driver.current_url:
                 List.append("  - 当前 URL 似乎是登录后页面，继续...")
                 login_successful = True
            else:
                 List.append("!! 登录警告: 未出现 KMSI 且 URL 未知，登录状态不确定。")
                 driver.save_screenshot(f"error_login_unknown_state_{username}.png")
                 # Proceed cautiously, maybe OAuth step will clarify
                 login_successful = True # Let OAuth step try anyway
        except NoSuchElementException as e:
            List.append(f"!! 登录错误 (KMSI): 无法找到 '保持登录状态?' 按钮。 {e}")
            driver.save_screenshot(f"error_login_kmsi_{username}.png")
            # Might still be logged in, let OAuth step try
            login_successful = True

        # --- NEW: OAuth Flow Step ---
        if login_successful:
            List.append(f"  - 登录成功，尝试导航到 OAuth URL...")
            captured_url = ""
            try:
                driver.get(OAUTH_URL)
                List.append(f"  - 已导航到: {OAUTH_URL[:80]}...") # Log start of OAuth URL

                # Wait for the URL to change to the REDIRECT_URI (containing the code)
                # This might happen very quickly if consent is already granted.
                # Need to handle potential consent screen here if it appears.
                
                consent_found = False
                try:
                     # Brief wait specifically for consent screen elements
                     consent_accept_button = WebDriverWait(driver, 15).until(
                          EC.element_to_be_clickable((By.XPATH, "//button[@id='consentAcceptButton'] | //button[contains(., 'Accept')] | //button[contains(., '接受')]"))
                     )
                     List.append("  - 检测到权限许可屏幕，尝试点击接受...")
                     driver.execute_script("arguments[0].click();", consent_accept_button)
                     consent_found = True
                except TimeoutException:
                     List.append("  - 未检测到权限许可屏幕 (或已超时)，等待重定向...")
                     pass # No consent screen found (or it timed out), proceed to wait for redirect
                except Exception as consent_err:
                     List.append(f"  - 尝试处理权限许可屏幕时出错: {consent_err}")
                     # Continue anyway, maybe redirect happens regardless

                # Now, wait specifically for the redirect URL
                List.append(f"  - 等待重定向到: {REDIRECT_URI_START}...")
                redirect_wait = WebDriverWait(driver, 45) # Wait up to 45s for redirect
                
                # Wait until the URL *starts with* the redirect URI
                redirect_wait.until(EC.url_starts_with(REDIRECT_URI_START)) 
                
                # Immediately capture the URL once the condition is met
                captured_url = driver.current_url
                List.append(f"  - 检测到重定向 URL!")

                # Validate if the captured URL looks correct and contains the code
                if captured_url.startswith(REDIRECT_URI_START) and "code=" in captured_url:
                    List.append(f"  - >> 成功捕获含授权码的 URL: {captured_url}")
                    
                    # Save the captured URL to a file
                    filename = f"oauth_redirect_{username}.txt"
                    try:
                        with open(filename, 'w') as f:
                            f.write(captured_url)
                        List.append(f"  - >> 已将 URL 保存到文件: {filename}")
                    except Exception as file_err:
                        List.append(f"  - !! 错误: 无法将 URL 保存到文件 {filename}: {file_err}")

                elif captured_url.startswith(REDIRECT_URI_START) and "error=" in captured_url:
                     List.append(f"  - !! 错误: 重定向 URL 包含错误信息: {captured_url}")
                     driver.save_screenshot(f"error_oauth_redirect_error_{username}.png")
                else:
                    List.append(f"  - !! 警告: 检测到重定向 URL，但格式不符合预期 (缺少 'code='?): {captured_url}")
                    driver.save_screenshot(f"error_oauth_redirect_unexpected_{username}.png")

            except TimeoutException:
                List.append(f"!! 错误: 等待重定向到 '{REDIRECT_URI_START}' 超时。")
                List.append(f"  - 当前 URL: {driver.current_url}")
                List.append(f"  - 可能原因: 登录失败、需要手动交互 (如 Captcha)、网络问题或 OAuth 配置错误。")
                driver.save_screenshot(f"error_oauth_redirect_timeout_{username}.png")
            except Exception as oauth_err:
                List.append(f"!! 导航到 OAuth URL 或处理重定向时发生意外错误: {oauth_err}")
                try:
                    List.append(f"  - 发生错误时 URL: {driver.current_url}")
                    driver.save_screenshot(f"error_oauth_unexpected_{username}.png")
                except Exception: pass # Ignore errors during error handling
        else:
             List.append("!! 跳过 OAuth 步骤，因为登录似乎未成功。")


        # --- E5 Check Removed ---
        # The original code for navigating to /subscriptions and checking expiry is removed/commented out.
        # --- End E5 Check Removal ---

    except Exception as e:
        List.append(f"!! 在处理账号 {username} 时发生意外的 Selenium 错误: {e}")
        try:
           driver.save_screenshot(f"error_unexpected_{username}.png") 
        except Exception as screen_err:
           List.append(f"!! (附加错误) 保存截图失败: {screen_err}")
    finally:
        if driver:
            driver.quit() 
        List.append(f"处理完成: {username}")


# --- Main Execution Block ---
if __name__ == '__main__':
    account_env_var = 'MS_E5_ACCOUNTS' 
    if account_env_var in os.environ:
        accounts_str = os.environ[account_env_var]
        if not accounts_str:
             List.append(f'!! 错误：环境变量 {account_env_var} 为空。请在 GitHub Secrets 中设置。')
        else:
            users = [acc.strip() for acc in accounts_str.split('&') if acc.strip()] 
            List.append(f'检测到 {len(users)} 个账号配置。')
            
            account_counter = 0
            for i, user_pair in enumerate(users):
                if '-' not in user_pair:
                    List.append(f'!! 错误：账号 {i+1} 格式不正确 (缺少 "-")，跳过: "{user_pair[:15]}..."')
                    continue
                
                name = "" 
                try:
                   name, pwd = user_pair.split('-', 1) 
                   name = name.strip()
                   pwd = pwd.strip()
                   if not name or not pwd:
                       List.append(f'!! 错误：账号 {i+1} 用户名或密码为空，跳过。')
                       continue
                       
                   account_counter += 1
                   List.append(f'\n======> [账号 {account_counter}: {name}] 开始 <======')
                   # Call the function to perform login and get OAuth code
                   get_oauth_code(name, pwd) 
                   List.append(f'======> [账号 {account_counter}: {name}] 结束 <======\n')
                   
                   sleep_time = random.uniform(8, 15) # Slightly shorter delay might be okay
                   List.append(f"  -- 账号间暂停 {sleep_time:.1f} 秒 --")
                   time.sleep(sleep_time) 

                except ValueError:
                   List.append(f'!! 错误：无法解析账号 {i+1}，格式应为 email-password。跳过: "{user_pair[:15]}..."')
                except Exception as e:
                   account_id_for_error = name if name else f"账号 {i+1} ({user_pair[:15]}...)"
                   List.append(f'!! 处理 {account_id_for_error} 时发生未知错误: {e}')
                   if not (List and f"账号 {account_counter}" in List[-1]): # Avoid double end markers
                        List.append(f'======> [账号 {account_counter}] 处理因错误中止 <======\n')


            # --- Final Output and Notification ---
            final_output = '\n'.join(List)
            print("--- Script Execution Summary ---")
            print(final_output)
            print("--- End Summary ---")
            
            try:
               send('Microsoft OAuth Code 获取尝试报告', final_output) # Changed report title
               print("通知已发送 (如果 sendNotify 配置正确)。")
            except NameError:
               print("通知跳过 (sendNotify 未导入或配置)。")
               pass 
            except Exception as notify_err:
               print(f"!! 发送通知时出错: {notify_err}")
            
    else:
        error_message = f'!! 严重错误：未找到环境变量 {account_env_var}。请在 GitHub Secrets 中配置账号信息。'
        print(error_message)
        List.append(error_message)
        try:
            send('Microsoft OAuth 脚本配置错误', error_message)
        except NameError:
            pass
        except Exception as notify_err:
            print(f"!! 发送配置错误通知时出错: {notify_err}")

