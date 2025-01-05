from DrissionPage import ChromiumOptions, Chromium
from DrissionPage.common import Keys
import re
import time
import random
from cursor_auth_manager import CursorAuthManager

def get_veri_code(tab):
    """获取验证码"""
    username = account.split('@')[0]
    try:
        while True: 
            if tab.ele('@id=pre_button'):
                tab.actions.click('@id=pre_button').type(Keys.CTRL_A).key_down(Keys.BACKSPACE).key_up(Keys.BACKSPACE).input(username).key_down(Keys.ENTER).key_up(Keys.ENTER)
                break
            time.sleep(1)

        while True:
            new_mail = tab.ele('@class=mail')
            if new_mail:
                if new_mail.text:
                    print('最新的邮件：', new_mail.text)
                    tab.actions.click('@class=mail')
                    break
                else:
                    print(new_mail)
                    break
            time.sleep(1)

        if tab.ele('@class=overflow-auto mb-20'):
            email_content = tab.ele('@class=overflow-auto mb-20').text
            verification_code = re.search(r'verification code is (\d{6})', email_content)
            if verification_code:
                code = verification_code.group(1)
                print('验证码：', code)
            else:
                print('未找到验证码')
       
        if tab.ele('@id=delete_mail'):
            tab.actions.click('@id=delete_mail')
            time.sleep(1)

        if tab.ele('@id=confirm_mail'):
            tab.actions.click('@id=confirm_mail')
            print("删除邮件")
        tab.close()
    except Exception as e:
        print(e)

    return code

def handle_turnstile(tab):
    """处理 Turnstile 验证"""
    print("准备处理验证")
    try:  
        while True:
            try:
                challengeCheck = (tab.ele('@id=cf-turnstile', timeout=2)
                                    .child()
                                    .shadow_root
                                    .ele("tag:iframe")
                                    .ele("tag:body")
                                    .sr("tag:input"))
                                    
                if challengeCheck:
                    print("验证框加载完成")
                    time.sleep(random.uniform(1, 3))
                    challengeCheck.click()
                    print("验证按钮已点击，等待验证完成...")
                    time.sleep(2)
                    return True
            except:
                pass

            if tab.ele('@name=password'):
                print("无需验证")   
                break            
            if tab.ele('@data-index=0'):
                print("无需验证")   
                break
            if tab.ele('Account Settings'):
                print("无需验证")   
                break       

            time.sleep(random.uniform(1,2))       
    except Exception as e:
        print(e)
        print('跳过验证')
        return False

def get_cursor_session_token(tab):
    """获取cursor session token"""
    cookies = tab.cookies()
    cursor_session_token = None
    for cookie in cookies:
        if cookie['name'] == 'WorkosCursorSessionToken':
            cursor_session_token = cookie['value'].split('%3A%3A')[1]
            break
    return cursor_session_token

def sign_up_account(browser, tab, account, password, first_name, last_name):
    """注册账户流程"""
    print("\n开始注册新账户...")
    tab.get(sign_up_url)

    try:
        if tab.ele('@name=first_name'):
            print("已打开注册页面")
            tab.actions.click('@name=first_name').input(first_name)
            time.sleep(random.uniform(1,3))

            tab.actions.click('@name=last_name').input(last_name)
            time.sleep(random.uniform(1,3))

            tab.actions.click('@name=email').input(account)
            print("输入邮箱" )
            time.sleep(random.uniform(1,3))

            tab.actions.click('@type=submit')
            print("点击注册按钮")

    except Exception as e:
        print("打开注册页面失败")
        return False

    handle_turnstile(tab)            

    try:
        if tab.ele('@name=password'):
            tab.ele('@name=password').input(password)
            print("输入密码")
            time.sleep(random.uniform(1,3))

            tab.ele('@type=submit').click()
            print("点击Continue按钮")

    except Exception as e:
        print("输入密码失败")
        return False

    time.sleep(random.uniform(1,3))
    if tab.ele('This email is not available.'):
        print('This email is not available.')
        return False

    handle_turnstile(tab)

    while True:
        try:
            if tab.ele('Account Settings'):
                break
            if tab.ele('@data-index=0'):
                tab_mail = browser.new_tab(mail_url)
                browser.activate_tab(tab_mail)
                print("打开邮箱页面")
                code = get_veri_code(tab_mail)

                if code:
                    print("获取验证码成功：", code)
                    browser.activate_tab(tab)
                else:
                    print("获取验证码失败，程序退出")
                    return False

                i = 0
                for digit in code:
                    tab.ele(f'@data-index={i}').input(digit)
                    time.sleep(random.uniform(0.1,0.3))
                    i += 1
                break
        except Exception as e:
            print(e)

    handle_turnstile(tab)
    
    time.sleep(random.uniform(1,3))
    print("进入设置页面")
    tab.get(settings_url)

    print("\n注册完成")
    print("Cursor 账号： " + account)
    print("       密码： " + password)
    return True

def main():
    # 预定义的注册信息
    global account, password, first_name, last_name
    account = "drfyapi@mailto.plus"        # 邮箱地址
    password = "GDjk@20241123"         # 密码
    first_name = "feng"                # 名字
    last_name = "yun"                  # 姓氏

    # 配置信息
    global sign_up_url, settings_url, mail_url
    sign_up_url = 'https://authenticator.cursor.sh/sign-up'
    settings_url = 'https://www.cursor.com/settings'
    mail_url = 'https://tempmail.plus'

    # 浏览器配置
    co = ChromiumOptions()
    co.add_extension("turnstilePatch")
    co.headless()                            #无头模式
    co.set_user_agent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.6723.92 Safari/537.36')
    co.set_pref('credentials_enable_service', False)
    co.set_argument('--hide-crash-restore-bubble') 
    co.auto_port()

    browser = Chromium(co)
    tab = browser.latest_tab
    tab.run_js("try { turnstile.reset() } catch(e) { }")
    
    print("\n开始执行注册流程")
    print(f"注册信息：")
    print(f"邮箱: {account}")
    print(f"密码: {password}")
    print(f"姓名: {first_name} {last_name}")
    
    # 执行注册流程
    if sign_up_account(browser, tab, account, password, first_name, last_name):
        token = get_cursor_session_token(tab)
        if token:
            print(f"\nCursorSessionToken: {token}")
        else:
            print("获取token失败")
    else:
        print("账户注册失败")

    print("\n脚本执行完毕")
    browser.quit()

if __name__ == "__main__":
    main() 