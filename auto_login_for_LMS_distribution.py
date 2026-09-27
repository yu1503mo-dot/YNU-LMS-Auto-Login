from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import time
import pandas as pd
from selenium.webdriver.common.by import By


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
        "Windows資格情報からパスワードを取得できませんでした"
    )

#print(USERNAME)

USER=USERNAME
PASS=PASSWORD

browser = webdriver.Chrome()
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

body = mail.Body
match = re.search(r"\d{8}", body)

if match:
    code = match.group(0)

    # クリップボードへコピー
    pyperclip.copy(code)

    print("認証コード:", code)
    print("クリップボードへコピーしました")
else:
    print("認証コードが見つかりません")




from selenium.webdriver.common.by import By
#テキストに入力

element = browser.find_element(By.ID, 'idToken1')
copied_text=pyperclip.paste()
element.send_keys(copied_text)








#入力した値でログインを実行
browser_from = browser.find_element(By.NAME, 'callback_1')  
time.sleep(3)
browser_from.click()
print("ログインの実行")




input("終了する場合はEnterキーを押してください...")