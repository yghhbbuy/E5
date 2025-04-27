#!/usr/bin/python3
# -*- coding: utf8 -*-
"""
说明: 
- 此脚本使用Selenium自动登录Microsoft 365 Admin Center并检查E5订阅有效期。
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

List = [] # To store output messages

# --- Configuration ---
LOGIN_URL = 'https://admin.microsoft.com/'
SUBSCRIPTIONS_URL = 'https://admin.microsoft.com/Adminportal/Home?source=applauncher#/subscriptions'
TARGET_SUBSCRIPTION_NAME = "Microsoft 365 E5" 

# --- Helper Function ---
def get_webdriver():
    options = webdriver.ChromeOptions()
    # Crucial options for GitHub Actions/headless environments
    options.add_argument("--headless") 
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage") 
    options.add_argument("--window-size=1920,1080")
    # Use a common user agent
    options.add_argument("user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36") 
    
    try:
       driver = webdriver.Chrome(options=options) 
       List.append("  - WebDriver 初始化成功。")
       return driver
    except WebDriverException as e:
       List.append(f"!! 错误：无法初始化WebDriver: {e}")
       List.append("!! 请检查工作流中的 ChromeDriver 安装步骤。")
       return None
    except Exception as e:
       List.append(f"!! 错误：初始化WebDriver时发生意外错误: {e}")
       return None


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
        try:
            email_field = wait.until(EC.visibility_of_element_located((By.ID, "i0116")))
            email_field.send_keys(username)
            next_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
            driver.execute_script("arguments[0].click();", next_button)
            List.append("  - 输入邮箱并点击下一步")
        except (NoSuchElementException, TimeoutException) as e:
            List.append(f"!! 错误：找不到邮箱输入框或超时。页面可能更改。 {e}")
            driver.save_screenshot(f"error_email_input_{username}.png") 
            return 

        time.sleep(random.uniform(3, 5)) 

        # --- Login Step 2: Enter Password ---
        try:
            password_field = wait.until(EC.visibility_of_element_located((By.ID, "i0118")))
            time.sleep(0.5)
            password_field.send_keys(password)
            signin_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
            driver.execute_script("arguments[0].click();", signin_button)
            List.append("  - 输入密码并点击登录")
        except (NoSuchElementException, TimeoutException) as e:
            try:
                if driver.find_element(By.ID, "i0118").is_displayed():
                   List.append("!! 警告: 似乎仍在密码页面，密码可能错误或登录流程异常。")
                else: raise NoSuchElementException 
            except NoSuchElementException:
                List.append(f"!! 错误：找不到密码输入框或登录按钮。密码错误或页面结构更改。 {e}")
            driver.save_screenshot(f"error_password_input_{username}.png")
            return

        # --- Login Step 3: Handle "Stay signed in?" (KMSI) ---
        try:
            kmsi_button_no = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.ID, "idBtn_Back")) 
            )
            driver.execute_script("arguments[0].click();", kmsi_button_no)
            List.append("  - 处理 '保持登录状态?' -> 否")
        except TimeoutException:
            List.append("  - 未出现 '保持登录状态?' 弹窗 (或已超时)，继续...")
            if "admin.microsoft.com" not in driver.current_url:
                 List.append("!! 警告: 未出现KMSI弹窗，且当前URL不是Admin Center。登录可能失败。")
                 driver.save_screenshot(f"error_post_login_url_{username}.png")

        # --- Navigate to Subscriptions Page ---
        List.append("  - 尝试导航到订阅页面...")
        time.sleep(random.uniform(4, 7)) 
        
        try:
            driver.get(SUBSCRIPTIONS_URL)
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[data-is-scrollable='true']")))
            List.append("  - 成功导航到订阅页面")
            time.sleep(random.uniform(2, 4)) 
        except TimeoutException:
            List.append("!! 错误：导航到订阅页面超时或找不到预期元素。登录失败或页面结构更改。")
            driver.save_screenshot(f"error_nav_subscriptions_{username}.png")
            return

        # --- Find E5 Subscription and Expiry Date ---
        try:
            List.append(f"  - 正在查找订阅: '{TARGET_SUBSCRIPTION_NAME}'")
            subscription_cards = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[role='row'], div[data-automation-id^='DetailsCard']"))) 

            found = False
            for card in subscription_cards:
                try:
                    title_element = card.find_element(By.CSS_SELECTOR, "div[data-automation-id='ProductTitle'], span[data-automation-id='ProductName']")
                    card_title = title_element.text
                    
                    if TARGET_SUBSCRIPTION_NAME in card_title:
                        List.append(f"  - 找到包含 '{TARGET_SUBSCRIPTION_NAME}' 的订阅卡片: '{card_title}'")
                        try:
                            expiry_element = card.find_element(By.XPATH, ".//*[contains(text(), 'Expires') or contains(text(), '到期')]")
                            expiry_text = expiry_element.text.strip()
                            List.append(f"  - >> 有效期信息: {expiry_text}")
                            found = True
                            break 
                        except NoSuchElementException:
                            List.append(f"  - !! 警告: 在 '{card_title}' 卡片中找到 E5，但未能定位 'Expires'/'到期' 文本。检查HTML结构。")

                except NoSuchElementException:
                    continue 

            if not found:
                List.append(f"!! 未找到与 '{TARGET_SUBSCRIPTION_NAME}' 匹配且包含可识别有效期信息的订阅。")
                driver.save_screenshot(f"error_sub_not_found_or_no_date_{username}.png")

        except TimeoutException:
            List.append("!! 错误：加载订阅列表超时。")
            driver.save_screenshot(f"error_loading_subs_{username}.png")

    except Exception as e:
        List.append(f"!! 发生意外的Selenium错误: {e}")
        try:
           driver.save_screenshot(f"error_unexpected_{username}.png") 
        except Exception as screen_err:
           List.append(f"!! (附加错误) 保存截图失败: {screen_err}")
    finally:
        if driver:
            driver.quit()
        List.append(f"检查完成: {username}")


if __name__ == '__main__':
    account_env_var = 'MS_E5_ACCOUNTS' 
    if account_env_var in os.environ:
        accounts_str = os.environ[account_env_var]
        if not accounts_str:
             List.append(f'!! 错误：环境变量 {account_env_var} 为空。请在 GitHub Secrets 中设置。')
        else:
            users = accounts_str.split('&')
            List.append(f'检测到 {len(users)} 个账号配置。')
            
            account_counter = 0
            for i, user_pair in enumerate(users):
                user_pair = user_pair.strip() 

                if '-' not in user_pair:
                    List.append(f'!! 错误：账号 {i+1} 格式不正确 (缺少 "-")，跳过: "{user_pair[:10]}..."')
                    continue
                    
                try:
                   name, pwd = user_pair.split('-', 1) 
                   name = name.strip()
                   pwd = pwd.strip()
                   if not name or not pwd:
                       List.append(f'!! 错误：账号 {i+1} 用户名或密码为空，跳过。')
                       continue
                       
                   account_counter += 1
                   List.append(f'\n======> [账号 {account_counter}: {name}] 开始 <======')
                   check_e5_expiry(name, pwd)
                   List.append(f'======> [账号 {account_counter}: {name}] 结束 <======\n')
                   
                   sleep_time = random.uniform(8, 15)
                   List.append(f"  -- 账号间暂停 {sleep_time:.1f} 秒 --")
                   time.sleep(sleep_time) 

                except ValueError:
                   List.append(f'!! 错误：无法解析账号 {i+1}，格式应为 email-password。跳过: "{user_pair[:10]}..."')
                except Exception as e:
                   List.append(f'!! 处理账号 {i+1} ({name}) 时发生未知错误: {e}')
                   List.append(f'======> [账号 {account_counter}] 检查因错误结束 <======\n')

            final_output = '\n'.join(List)
            print("--- Script Execution Summary ---")
            print(final_output)
            
            try:
               send('Microsoft E5 订阅检查报告', final_output)
            except NameError:
               pass 
            except Exception as notify_err:
               print(f"!! 发送通知时出错: {notify_err}")
            
    else:
        print(f'!! 错误：未找到环境变量 {account_env_var}。请在 GitHub Secrets 中配置。')
        List.append(f'错误：未找到环境变量 {account_env_var}。')
        try:
            send('Microsoft E5 订阅检查错误', f'错误：未配置 GitHub Secret {account_env_var}')
        except NameError:
            pass
        except Exception as notify_err:
            print(f"!! 发送配置错误通知时出错: {notify_err}")
