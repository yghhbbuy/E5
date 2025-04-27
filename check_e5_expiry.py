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
LOGIN_URL = 'https://admin.microsoft.com/'
SUBSCRIPTIONS_URL = 'https://admin.microsoft.com/Adminportal/Home?source=applauncher#/subscriptions'
# Check your subscription page for the exact name
TARGET_SUBSCRIPTION_NAME = "Microsoft 365 E5" 

# --- Helper Function ---
def get_webdriver():
    options = webdriver.ChromeOptions()
    # Crucial options for GitHub Actions/headless environments
    options.add_argument("--headless") 
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage") 
    # Some sites behave differently with smaller viewports
    options.add_argument("--window-size=1920,1080") 
    # Use a common user agent to avoid fingerprinting issues
    options.add_argument("user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36") 
    
    # --- FIX for "no chrome binary" error ---
    # Specify the path to the chromium binary installed by apt in the workflow
    options.binary_location = "/usr/bin/chromium-browser" 
    # --- END FIX ---

    # In GitHub Actions with apt install, chromedriver should be in PATH
    try:
       # Selenium will use the specified binary_location with the chromedriver found in PATH
       driver = webdriver.Chrome(options=options) 
       List.append("  - WebDriver 初始化成功 (使用 /usr/bin/chromium-browser)。") 
       return driver
    except WebDriverException as e:
       # Check if the specified binary location was actually wrong
       if "no chrome binary at /usr/bin/chromium-browser" in str(e) or "cannot find chrome binary" in str(e).lower():
            List.append(f"!! 错误：无法在指定位置 /usr/bin/chromium-browser 找到 Chromium。尝试 /usr/bin/chromium ...")
            options.binary_location = "/usr/bin/chromium" # Try alternative common path
            try:
                # Retry initialization with the alternative path
                driver = webdriver.Chrome(options=options)
                List.append("  - WebDriver 初始化成功 (使用 /usr/bin/chromium)。") 
                return driver
            except WebDriverException as e2:
                 List.append(f"!! 错误：在 /usr/bin/chromium 也无法初始化 WebDriver: {e2}")
                 List.append("!! 请检查工作流中的 chromium-browser 安装和实际路径。")
                 return None
            except Exception as e3:
                 List.append(f"!! 尝试 /usr/bin/chromium 时发生意外错误: {e3}")
                 return None


       # If error was different, report it
       List.append(f"!! 错误：无法初始化 WebDriver: {e}")
       List.append("!! 请检查工作流中的 ChromeDriver 安装步骤和版本兼容性。")
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
        # Increased wait time for potentially slow cloud environments
        wait = WebDriverWait(driver, 60) # Extend wait time further if needed

        # --- Login Step 1: Enter Email ---
        try:
            # Wait for email field to be visible and ready
            email_field = wait.until(EC.visibility_of_element_located((By.ID, "i0116")))
            email_field.send_keys(username)
            # Wait for the 'Next' button to be clickable
            next_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
            # Using JS click can be more reliable in automated environments
            driver.execute_script("arguments[0].click();", next_button)
            List.append("  - 输入邮箱并点击下一步")
        except (NoSuchElementException, TimeoutException) as e:
            List.append(f"!! 错误：找不到邮箱输入框或超时。页面可能更改。 {e}")
            driver.save_screenshot(f"error_email_input_{username}.png") 
            return # Stop check for this user

        # Add a slightly longer pause for transitions, especially after JS interaction
        time.sleep(random.uniform(4, 6)) 

        # --- Login Step 2: Enter Password ---
        try:
            # Wait for password field to be visible
            password_field = wait.until(EC.visibility_of_element_located((By.ID, "i0118")))
            # Brief pause before sending keys can sometimes help
            time.sleep(0.7)
            password_field.send_keys(password)
            # Wait for sign-in button to be clickable
            signin_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
            driver.execute_script("arguments[0].click();", signin_button)
            List.append("  - 输入密码并点击登录")
        except (NoSuchElementException, TimeoutException) as e:
            # Check common alternative scenarios (like wrong password page staying)
            try:
                if driver.find_element(By.ID, "i0118").is_displayed():
                   List.append("!! 警告: 似乎仍在密码页面，密码可能错误或登录流程异常。")
                else: raise NoSuchElementException # Re-raise if not the password field
            except NoSuchElementException:
                 # If it's not the password field, maybe it's MFA or other prompt
                 try:
                    if "aadcdn.msauth.net" in driver.current_url or "MFA" in driver.page_source:
                         List.append("!! 错误: 检测到 MFA 或身份验证提示页面，无法自动处理。")
                    else:
                         List.append(f"!! 错误：找不到密码输入框或登录按钮。密码错误或页面结构更改。 {e}")
                 except Exception: # Catch errors checking page source/URL
                     List.append(f"!! 错误：找不到密码输入框或登录按钮。密码错误或页面结构更改。 {e}")
            driver.save_screenshot(f"error_password_input_{username}.png")
            return

        # --- Login Step 3: Handle "Stay signed in?" (KMSI) ---
        try:
            # Wait a bit longer for this prompt, it can be delayed
            kmsi_button_no = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.ID, "idBtn_Back")) # The "No" button
            )
            driver.execute_script("arguments[0].click();", kmsi_button_no)
            List.append("  - 处理 '保持登录状态?' -> 否")
        except TimeoutException:
            List.append("  - 未出现 '保持登录状态?' 弹窗 (或已超时)，继续...")
            # Double-check if we are logged in by looking for expected elements or URL
            time.sleep(2) # Allow redirects to settle
            if "admin.microsoft.com" not in driver.current_url:
                 List.append("!! 警告: 未出现KMSI弹窗，且当前URL不是Admin Center。登录可能失败。")
                 driver.save_screenshot(f"error_post_login_url_{username}.png")
                 # Consider returning if strict login check is needed, but nav might fix it
        except NoSuchElementException as e:
            List.append(f"!! 错误：无法找到 '保持登录状态?' 按钮。页面结构可能更改。 {e}")
            driver.save_screenshot(f"error_kmsi_button_{username}.png")
            # Continue cautiously, maybe the flow changed

        # --- Navigate to Subscriptions Page ---
        List.append("  - 尝试导航到订阅页面...")
        # Allow generous time for dashboard loading/redirects before navigating
        time.sleep(random.uniform(6, 10)) 
        
        try:
            driver.get(SUBSCRIPTIONS_URL)
            # Wait for a reliable container element on the subscriptions page.
            # This selector looks for the main content area where product cards/rows appear.
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='grid'], div[data-is-scrollable='true']")))
            List.append("  - 成功导航到订阅页面")
            # Allow time for the dynamic content (subscription list) to load within the container
            time.sleep(random.uniform(4, 7)) 
        except TimeoutException:
            List.append("!! 错误：导航到订阅页面超时或找不到预期容器元素。登录失败或页面结构更改。")
            driver.save_screenshot(f"error_nav_subscriptions_{username}.png")
            return
        except Exception as e:
             List.append(f"!! 导航到订阅页面时发生意外错误: {e}")
             driver.save_screenshot(f"error_nav_subscriptions_{username}.png")
             return

        # --- Find E5 Subscription and Expiry Date ---
        try:
            List.append(f"  - 正在查找订阅: '{TARGET_SUBSCRIPTION_NAME}'")
            
            # Wait longer for subscription items/cards to be present as they load dynamically
            subscription_cards = WebDriverWait(driver, 45).until(EC.presence_of_all_elements_located(
                # Look for elements likely representing a subscription row or card
                (By.CSS_SELECTOR, "div[role='row'], div[data-automation-id^='DetailsCard'], div[class*='SubscriptionCard']"))) 

            found = False
            if not subscription_cards:
                 List.append("!! 警告: 未在页面上检测到任何订阅卡片/行元素。页面可能为空或结构更改。")
                 driver.save_screenshot(f"error_no_sub_cards_{username}.png")

            for card_index, card in enumerate(subscription_cards):
                try:
                    # Try to find the title within the card first using common patterns
                    # Use more specific selectors if possible by inspecting the page
                    title_element = card.find_element(By.CSS_SELECTOR, "div[data-automation-id='ProductTitle'], span[data-automation-id='ProductName'], h3[class*='CardTitle']")
                    card_title = title_element.text
                    
                    # Check if the target name is part of the found title
                    if TARGET_SUBSCRIPTION_NAME in card_title:
                        List.append(f"  - [卡片 {card_index+1}] 找到含 '{TARGET_SUBSCRIPTION_NAME}' 的订阅: '{card_title}'")
                        
                        # Now find the expiry date *within this specific card*. Be precise.
                        # Look for text patterns or specific data attributes.
                        try:
                            # Search within the current 'card' element using relative XPath
                            expiry_element = card.find_element(By.XPATH, ".//*[contains(text(), 'Expires') or contains(text(), '到期') or contains(text(), '到期日期') or contains(text(), 'Expiration date')]")
                            expiry_text = expiry_element.text.strip()
                            # Clean up common prefixes if needed
                            expiry_text = expiry_text.replace("Expires on", "").replace("Expires", "").replace("到期日期", "").replace("到期", "").replace("Expiration date", "").strip()
                            List.append(f"  - >> 有效期信息: {expiry_text}")
                            found = True
                            break # Stop after finding the first relevant E5 subscription

                        except NoSuchElementException:
                            # If text search fails, try finding by known data-automation-id (requires inspection)
                            try:
                                expiry_alt = card.find_element(By.CSS_SELECTOR, "[data-automation-id='SubscriptionEndDate'], [data-automationid='expirationdate']")
                                expiry_text = expiry_alt.text.strip()
                                if expiry_text:
                                   List.append(f"  - >> (备选定位) 有效期信息: {expiry_text}")
                                   found = True
                                   break
                                else:
                                   # Element found but empty? Log it.
                                   List.append(f"  - !! 警告: 在 '{card_title}' 卡片中找到备选有效期元素但无文本。")
                            except NoSuchElementException:
                                List.append(f"  - !! 警告: 在 '{card_title}' 卡片中未能定位有效期信息 (文本或备选ID)。需要检查HTML结构。")
                                driver.save_screenshot(f"error_find_expiry_detail_{username}_{card_index}.png")
                                # Don't break yet, maybe another card structure matches E5 better?

                except NoSuchElementException:
                    # This card might not be a subscription card or has a different structure
                    # List.append(f"  - Debug: Card {card_index+1} skipped (no standard title found).") # Optional debug
                    continue # Check the next potential card element

            if not found:
                List.append(f"!! 未找到与 '{TARGET_SUBSCRIPTION_NAME}' 匹配且包含可识别有效期信息的订阅。")
                driver.save_screenshot(f"error_sub_not_found_or_no_date_{username}.png")

        except TimeoutException:
            List.append("!! 错误：等待订阅列表加载超时。页面可能为空或加载缓慢。")
            driver.save_screenshot(f"error_loading_subs_{username}.png")
        except Exception as e:
            List.append(f"!! 查找订阅或有效期时出错: {e}")
            driver.save_screenshot(f"error_finding_subs_{username}.png")

    except Exception as e:
        List.append(f"!! 在处理账号 {username} 时发生意外的Selenium错误: {e}")
        try:
           # Attempt to save screenshot even on unexpected errors
           driver.save_screenshot(f"error_unexpected_{username}.png") 
        except Exception as screen_err:
           List.append(f"!! (附加错误) 保存截图失败: {screen_err}")
    finally:
        if driver:
            # Ensure browser is closed even if errors occurred
            driver.quit() 
        List.append(f"检查完成: {username}")


if __name__ == '__main__':
    # --- Environment Variable Processing ---
    account_env_var = 'MS_E5_ACCOUNTS' 
    if account_env_var in os.environ:
        accounts_str = os.environ[account_env_var]
        if not accounts_str:
             List.append(f'!! 错误：环境变量 {account_env_var} 为空。请在 GitHub Secrets 中设置。')
        else:
            # Split accounts, handling potential extra whitespace
            users = [acc.strip() for acc in accounts_str.split('&') if acc.strip()] 
            List.append(f'检测到 {len(users)} 个账号配置。')
            
            account_counter = 0
            for i, user_pair in enumerate(users):
                
                if '-' not in user_pair:
                    List.append(f'!! 错误：账号 {i+1} 格式不正确 (缺少 "-")，跳过: "{user_pair[:15]}..."')
                    continue
                    
                name = "" # Initialize name for potential error messages
                try:
                   # Split only on the first hyphen to handle passwords with hyphens
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
                   
                   # Add delay between accounts - crucial for not getting blocked/throttled
                   sleep_time = random.uniform(10, 20) 
                   List.append(f"  -- 账号间暂停 {sleep_time:.1f} 秒 --")
                   time.sleep(sleep_time) 

                except ValueError:
                   List.append(f'!! 错误：无法解析账号 {i+1}，格式应为 email-password。跳过: "{user_pair[:15]}..."')
                except Exception as e:
                   # Log error associated with the specific account if possible
                   account_id_for_error = name if name else f"账号 {i+1} ({user_pair[:15]}...)"
                   List.append(f'!! 处理 {account_id_for_error} 时发生未知错误: {e}')
                   # Ensure the end marker is added even if an error occurs within the loop iteration
                   if account_counter > 0 and f"账号 {account_counter}" in List[-1]: # Avoid double end markers
                       pass
                   else:
                       List.append(f'======> [账号 {account_counter}] 检查因错误中止 <======\n')


            # --- Final Output and Notification ---
            final_output = '\n'.join(List)
            print("--- Script Execution Summary ---")
            print(final_output)
            print("--- End Summary ---")
            
            # Send notification using sendNotify.py if configured
            try:
               send('Microsoft E5 订阅检查报告', final_output)
               print("通知已发送 (如果 sendNotify 配置正确)。")
            except NameError:
               # send function wasn't defined (ImportError occurred)
               print("通知跳过 (sendNotify 未导入或配置)。")
               pass 
            except Exception as notify_err:
               print(f"!! 发送通知时出错: {notify_err}")
            
    else:
        # Critical configuration error
        error_message = f'!! 严重错误：未找到环境变量 {account_env_var}。请在 GitHub Secrets 中配置账号信息。'
        print(error_message)
        List.append(error_message)
        # Attempt to send notification about the missing configuration
        try:
            send('Microsoft E5 脚本配置错误', error_message)
        except NameError:
            pass
        except Exception as notify_err:
            print(f"!! 发送配置错误通知时出错: {notify_err}")
