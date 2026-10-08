from selenium import webdriver
# from selenium.webdriver.chrome.service import Service
import time
# import pandas as pd
from selenium.webdriver.common.by import By
import os
from selenium.webdriver.edge.options import Options

import keyring


PASSWORD = keyring.get_password(
    "YNUUniversityAutoLogin_Password",
    "account"
)

if PASSWORD is None:
    raise RuntimeError(
        "Windows資格情報からパスワードを取得できませんでした"
    )

#print(PASSWORD)

USERNAME = keyring.get_password(
    "YNUUniversityAutoLogin_User",
    "account"
)

if USERNAME is None:
    raise RuntimeError(
        "Windows資格情報からユーザー名を取得できませんでした"
    )

#print(USERNAME)

USER=USERNAME
PASS=PASSWORD

# edge_user_data = os.path.expandvars(
#     r"%LOCALAPPDATA%\Microsoft\Edge\User Data"
# )

# options = Options()
# options.add_argument(f"--user-data-dir={edge_user_data}")
# options.add_argument("--profile-directory=Default")

# browser = webdriver.Edge(options=options)
browser = webdriver.Edge()
#browser = webdriver.Chrome()
browser.implicitly_wait(3)

#ログインしたいページへ転移

url_login = "https://lms.ynu.ac.jp/lms/lginLgir/"
browser.get(url_login)
browser_from = browser.find_element(By.NAME, 'loginButton')  # Updated method using By.NAME
time.sleep(1)
browser_from.click()
time.sleep(3)
print("ログインページにアクセス")




from selenium.webdriver.common.by import By
#テキストに入力

element = browser.find_element(By.ID, 'idToken1')
element.clear()
element.send_keys(USER)

element = browser.find_element(By.ID, 'idToken2')
element.clear()
element.send_keys(PASS)




#入力した値でログインを実行
browser_from = browser.find_element(By.NAME, 'callback_2')  # Updated method using By.NAME
time.sleep(3)
browser_from.click()
print("ログインの実行　学内の場合")




import win32com.client
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

inbox = namespace.GetDefaultFolder(6)  # 6 = 受信トレイ
mails = inbox.Items




mails = inbox.Items
mails.Sort("[ReceivedTime]", True)




time.sleep(30)
for i in range(1, 2):
    mail = mails.Item(i)
    print(mail.ReceivedTime, mail.Subject)

    print("本文:")
    print(mail.ReceivedTime, mail.Body)




import re
import pyperclip
# 受信トレイ
inbox = namespace.GetDefaultFolder(6)

# 新しい順
items = inbox.Items
items.Sort("[ReceivedTime]", True)
mail = items.Item(1)

is_auth_mail = "認証コード送信" in mail.Subject
auth_elements = browser.find_elements(By.ID, 'idToken1')

if auth_elements and auth_elements[0].is_displayed():
    print("認証コード入力欄が表示されています")
    if is_auth_mail:
        match = re.search(r"\d{8}", mail.Body)
        if match:
            code = match.group(0)

            # クリップボードへコピー
            pyperclip.copy(code)

            print("認証コード:", code)
            print("クリップボードへコピーしました")

            # テキストに入力
            element = browser.find_element(By.ID, 'idToken1')
            copied_text = pyperclip.paste()
            element.send_keys(copied_text)

            # 入力した値でログインを実行
            browser_from = browser.find_element(By.NAME, 'callback_1')  
            time.sleep(3)
            browser_from.click()
            print("ログインの実行")
        else:
            print("認証メールですが、本文から8桁のコードが見つかりませんでした")
    else:
        print(f"最新メール（件名: {mail.Subject}）は認証メールではありませんでした")
else:
    print("2段階認証画面は表示されていません（学内アクセス等によりログイン完了）")



# body = mail.Body
# match = re.search(r"\d{8}", body)

# is_auth_mail = "認証コード送信" in mail.Subject

# if match and is_auth_mail:
#     code = match.group(0)

#     # クリップボードへコピー
#     pyperclip.copy(code)

#     print("認証コード:", code)
#     print("クリップボードへコピーしました")
#     from selenium.webdriver.common.by import By
#     #テキストに入力

#     element = browser.find_element(By.ID, 'idToken1')
#     copied_text=pyperclip.paste()
#     element.send_keys(copied_text)








#     #入力した値でログインを実行
#     browser_from = browser.find_element(By.NAME, 'callback_1')  
#     time.sleep(3)
#     browser_from.click()
#     print("ログインの実行")
# else:
#     print("認証コードが見つかりません")




# from selenium.webdriver.common.by import By
# #テキストに入力

# element = browser.find_element(By.ID, 'idToken1')
# copied_text=pyperclip.paste()
# element.send_keys(copied_text)








# #入力した値でログインを実行
# browser_from = browser.find_element(By.NAME, 'callback_1')  
# time.sleep(3)
# browser_from.click()
# print("ログインの実行")




input("終了する場合はEnterキーを押してください...")