



# ============================================================
# set_up_for_autolog_LMS
#
# YNU LMS Auto Login 配布用初期設定プログラム
#
# このプログラムで行うこと
#
# ① ユーザーID・パスワードを入力
# ② Windows資格情報マネージャーへ保存
# ③ 保存確認
# ④ Pythonの場所を自動取得
# ⑤ auto_login_for_LMS_distribution.py の場所を自動取得
# ⑥ タスクスケジューラへ登録
# ⑦ タスク登録の確認
#
# ============================================================


import keyring
import os
# import shutil
import subprocess
import sys

# ============================================================
# 設定
# ============================================================

SERVICE_USER = "YNUUniversityAutoLogin_User"

SERVICE_PASSWORD = "YNUUniversityAutoLogin_Password"

ACCOUNT = "account"

TASK_NAME = "YNU LMS Auto Login"

SCRIPT_NAME = "auto_login_for_LMS_distribution.py"


# ============================================================
# ① ユーザーID・パスワード入力
# ============================================================

print("========================================")
print(" YNU LMS Auto Login 初期設定")
print("========================================")
print()

USER = input("ユーザーIDを入力してください: ").strip()

if not USER:
    print()
    print("エラー：ユーザーIDが入力されていません。")
    input("Enterキーを押して終了してください...")
    exit()


PASSWORD = input("パスワードを入力してください: ")

if not PASSWORD:
    print()
    print("エラー：パスワードが入力されていません。")
    input("Enterキーを押して終了してください...")
    exit()


print()
print("認証情報をWindows資格情報マネージャーに保存しています...")
print()


# ============================================================
# ② Credential Managerへ保存
# ============================================================

keyring.set_password(
    SERVICE_USER,
    ACCOUNT,
    USER
)

keyring.set_password(
    SERVICE_PASSWORD,
    ACCOUNT,
    PASSWORD
)


# ============================================================
# ③ 保存確認
# ============================================================

SAVED_USER = keyring.get_password(
    SERVICE_USER,
    ACCOUNT
)

SAVED_PASSWORD = keyring.get_password(
    SERVICE_PASSWORD,
    ACCOUNT
)


if SAVED_USER is None:

    print()
    print("========================================")
    print(" エラー")
    print("========================================")
    print()
    print("ユーザーIDの保存に失敗しました。")
    print()

    input("Enterキーを押して終了してください...")
    exit()


if SAVED_PASSWORD is None:

    print()
    print("========================================")
    print(" エラー")
    print("========================================")
    print()
    print("パスワードの保存に失敗しました。")
    print()

    input("Enterキーを押して終了してください...")
    exit()


print("========================================")
print(" Credential Managerへの登録完了")
print("========================================")
print()
print("✓ ユーザーIDを保存しました")
print("✓ パスワードを保存しました")
print()


# ============================================================
# ④ Pythonの場所を自動取得
# ============================================================

print("Pythonの場所を確認しています...")
print()


# PYTHON_PATH = shutil.which("python")


PYTHON_PATH = sys.executable

if PYTHON_PATH is None:

    print("========================================")
    print(" エラー")
    print("========================================")
    print()
    print("Pythonが見つかりません。")
    print()
    print("Pythonをインストールし、")
    print("pythonコマンドが使用できる状態にしてください。")
    print()

    input("Enterキーを押して終了してください...")
    exit()


print("使用するPython:")
print(PYTHON_PATH)
print()


# ============================================================
# ⑤ auto_login_for_LMS_distribution.py の場所を取得
# ============================================================

BASE_DIR = os.path.dirname(
    os.path.abspath(__file__)
)

SCRIPT_PATH = os.path.join(
    BASE_DIR,
    SCRIPT_NAME
)


print("自動ログインプログラムの場所を確認しています...")
print()


if not os.path.isfile(SCRIPT_PATH):

    print("========================================")
    print(" エラー")
    print("========================================")
    print()
    print(SCRIPT_NAME + " が見つかりません。")
    print()
    print("必要なファイル:")
    print(SCRIPT_PATH)
    print()

    input("Enterキーを押して終了してください...")
    exit()


print("自動ログインプログラム:")
print(SCRIPT_PATH)
print()


# ============================================================
# ⑥ 同じ名前のタスクが存在するか確認
# ============================================================

print("既存のタスクを確認しています...")
print()


result = subprocess.run(
    [
        "schtasks",
        "/Query",
        "/TN",
        TASK_NAME
    ],
    capture_output=True,
    text=True
)


if result.returncode == 0:

    print("========================================")
    print(" タスク登録に失敗しました")
    print("========================================")
    print()
    print("同じ名前のタスクがすでに存在します。")
    print()
    print("タスク名:")
    print(TASK_NAME)
    print()
    print("既存のタスクは削除していません。")
    print()
    print("必要であれば、既存のタスクを確認してから")
    print("手動で削除してください。")
    print()

    input("Enterキーを押して終了してください...")
    exit()


print("同じ名前のタスクは存在しません。")
print()


# ============================================================
# ⑦ タスクスケジューラへ登録
# ============================================================

print("タスクスケジューラに登録しています...")
print()


result = subprocess.run(
    [
        "schtasks",
        "/Create",
        "/TN",
        TASK_NAME,
        "/TR",
        f'"{PYTHON_PATH}" "{SCRIPT_PATH}"',
        "/SC",
        "ONLOGON",
        "/RL",
        "LIMITED"
    ],
    capture_output=True,
    text=True
)


if result.returncode != 0:

    print("========================================")
    print(" タスク登録に失敗しました")
    print("========================================")
    print()

    if result.stdout:
        print(result.stdout)

    if result.stderr:
        print(result.stderr)

    print()

    input("Enterキーを押して終了してください...")
    exit()


# ============================================================
# ⑧ タスクの電源条件を変更
# ============================================================

print("タスクの電源条件を設定しています...")
print()


import win32com.client

schedule_service = win32com.client.Dispatch("Schedule.Service")
schedule_service.Connect()

# 今回登録した「YNU LMS Auto Login」だけを取得
# 

root_folder = schedule_service.GetFolder("\\")
task = root_folder.GetTask(TASK_NAME)

task_definition = task.Definition

task_settings = task_definition.Settings

# バッテリー駆動でもタスクを開始する
task_settings.DisallowStartIfOnBatteries = False

# バッテリー駆動に切り替わってもタスクを停止しない
task_settings.StopIfGoingOnBatteries = False

# 設定を保存
root_folder.RegisterTaskDefinition(
    TASK_NAME,
    task_definition,
    4,
    "",
    "",
    task.Definition.Principal.LogonType
)

print("✓ 電源条件を設定しました")
print("✓ バッテリー駆動でもタスクを開始します")
print("✓ バッテリー駆動に切り替わっても停止しません")
print()



# ============================================================
# ⑧ タスク登録後の確認
# ============================================================

print("タスクが登録されたか確認しています...")
print()


result = subprocess.run(
    [
        "schtasks",
        "/Query",
        "/TN",
        TASK_NAME
    ],
    capture_output=True,
    text=True
)


if result.returncode != 0:

    print("========================================")
    print(" タスク登録の確認に失敗しました")
    print("========================================")
    print()
    print("タスクを登録しましたが、")
    print("登録後の確認ができませんでした。")
    print()

    input("Enterキーを押して終了してください...")
    exit()


# ============================================================
# ⑨ セットアップ完了
# ============================================================

print()
print("========================================")
print(" 初期設定が完了しました")
print("========================================")
print()

print("【Credential Manager】")
print("✓ ユーザーIDを保存しました")
print("✓ パスワードを保存しました")
print()

print("【Python】")
print(PYTHON_PATH)
print()

print("【自動ログインプログラム】")
print(SCRIPT_PATH)
print()

print("【タスクスケジューラ】")
print("✓ タスクを登録しました")
print()
print("タスク名:")
print(TASK_NAME)
print()
print("実行タイミング:")
print("Windowsへのログオン時")
print()

print("Windowsにログオンすると、")
print("自動ログインプログラムが実行されます。")
print()

input("Enterキーを押して終了してください...")


# # ============================================================
# # set_up_for_autolog_LMSの第一段階
# # ローカルの資格情報マネージャーにユーザーIDとパスワードを保存するパート
# #ユーザーにユーザーIDとパスワードをターミナルから入力してもらう
# #=========================================================

# import keyring
# from getpass import getpass


# # ============================================================
# # 設定
# # ============================================================

# SERVICE_USER = "YNUUniversityAutoLogin_User"
# SERVICE_PASSWORD = "YNUUniversityAutoLogin_Password"
# ACCOUNT = "account"


# # ============================================================
# # ① ユーザーID・パスワード入力
# # ============================================================

# print("========================================")
# print(" YNU LMS Auto Login 初期設定")
# print("========================================")
# print()

# USER = input("ユーザーIDを入力してください: ").strip()

# if not USER:
#     print("エラー：ユーザーIDが入力されていません。")
#     input("Enterキーを押して終了してください...")
#     exit()


# PASSWORD = input("パスワードを入力してください: ")

# if not PASSWORD:
#     print("エラー：パスワードが入力されていません。")
#     input("Enterキーを押して終了してください...")
#     exit()


# print()
# print("認証情報をWindows資格情報マネージャーに保存しています...")


# # ============================================================
# # ② Credential Managerへ保存
# # ============================================================

# keyring.set_password(
#     SERVICE_USER,
#     ACCOUNT,
#     USER
# )

# keyring.set_password(
#     SERVICE_PASSWORD,
#     ACCOUNT,
#     PASSWORD
# )


# # ============================================================
# # ③ 保存確認
# # ============================================================

# SAVED_USER = keyring.get_password(
#     SERVICE_USER,
#     ACCOUNT
# )

# SAVED_PASSWORD = keyring.get_password(
#     SERVICE_PASSWORD,
#     ACCOUNT
# )


# # print(SAVED_USER)
# # print(SAVED_PASSWORD)


# if SAVED_USER is None:
#     print()
#     print("エラー：ユーザーIDの保存に失敗しました。")
#     input("Enterキーを押して終了してください...")
#     exit()


# if SAVED_PASSWORD is None:
#     print()
#     print("エラー：パスワードの保存に失敗しました。")
#     input("Enterキーを押して終了してください...")
#     exit()


# print()
# print("========================================")
# print(" Credential Managerへの登録完了")
# print("========================================")
# print()
# print("✓ ユーザーIDを保存しました")
# print("✓ パスワードを保存しました")
# print()
# print("初期設定①～③は完了です。")
# print()

# input("Enterキーを押して終了してください...")

# # ============================================================
# # set_up_for_autolog_LMSの第一段階が完了
# # ローカルの資格情報マネージャーにユーザーIDとパスワードを保存するパートが完了しました。
# #=========================================================


