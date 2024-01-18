import smtplib
import ssl
import time
from email.mime.text import MIMEText
import gspread
import pandas as pd
import requests
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import datetime
import pytz
from password import password


def scraping():
    subtitle_list = []
    term_list = []
    target_list = []
    subject_list = []
    class_list = []
    contact_list = []

    
    # もっと見るをdisabledになるまでクリック
    while browser.find_element_by_id('moreBtn').is_enabled():
        browser.find_element_by_id('moreBtn').click()
    
    # 一つ一つ情報を取得
    js_list = [ onclick.get_attribute('onclick') for onclick in browser.find_element_by_class_name('box-link-contents').find_elements_by_tag_name('a')] 
    for js in js_list:
        # 内容に移動
        browser.execute_script(js)
        time.sleep(1)    
        
        # 情報を取得
        padding_top_no = browser.find_elements_by_class_name('padding-top-no')
        subtitle = padding_top_no[0].find_element_by_tag_name('h2')
        subject =padding_top_no[0].find_element_by_class_name('text-contents')

        subtitle_list.append(subtitle.text)
        subject_list.append(subject.text)


        text_contents = padding_top_no[1].find_elements_by_class_name('text-contents')

        term_list.append(f'掲示期間：{text_contents[0].text}')
        target_list.append(f'連絡対象：{text_contents[1].text}')
        class_list.append(f'お知らせ区分：{text_contents[2].text}')

        contact_list.append(f'問合せ先：{padding_top_no[2].text[5:]}')

        # 一覧へ戻るをクリック
        browser.find_element_by_name('cancel').click()
        time.sleep(1)
        
    # dataframeにまとめる
    dfs = pd.DataFrame(zip(subtitle_list, term_list, target_list, subject_list, class_list, contact_list),
                       columns=['サブタイトル', '掲示期間', '連絡対象', '連絡内容', 'お知らせ区分', '問合せ先'])

    return dfs

def send_line(nothing):
    nothing_number = len(nothing.index)

    for x in range(nothing_number):
        nothing_subtitle = nothing.iloc[x, 0]
        nothing_term = nothing.iloc[x, 1]
        nothing_target = nothing.iloc[x, 2]
        nothing_subject = nothing.iloc[x, 3]
        nothing_class = nothing.iloc[x, 4]
        nothing_contact = nothing.iloc[x, 5]

        # headers
        access_token = LINE_TOKEN  # 発行されたトークンへ置き換える
        headers = {"Authorization": f"Bearer {access_token}"}
        # send massage
        text = f'{nothing_term}\n{nothing_target}\n\n{nothing_subject}\n{nothing_class}\n{nothing_contact}'
        data = {"message": f"\n{text}"}
        requests.post("https://notify-api.line.me/api/notify", data=data, headers=headers)


def send_email(nothing):
    nothing_number = len(nothing.index)

    for x in range(nothing_number):
        nothing_subtitle = nothing.iloc[x, 0]
        nothing_term = nothing.iloc[x, 1]
        nothing_target = nothing.iloc[x, 2]
        nothing_subject = nothing.iloc[x, 3]
        nothing_class = nothing.iloc[x, 4]
        nothing_contact = nothing.iloc[x, 5]

        email_nothing_subject = nothing_subject.replace('\n', '<br>')

        email_text = f'{nothing_term}<br>{nothing_target}<br>{email_nothing_subject}<br>{nothing_class}<br>{nothing_contact}'
        # SMTP認証情報
        account = FROM_EMAIL
        password = SMTP_PASSWORD
        # 送受信先
        to_email = TO_EMAIL
        from_email = FROM_EMAIL

        # MIMEの作成
        subject = nothing_subtitle
        message = email_text
        msg = MIMEText(message, "html")
        msg["Subject"] = subject
        msg["To"] = to_email
        msg["From"] = from_email

        server = smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context())

        server.login(account, password)
        server.send_message(msg)


ID = password.ID
PASSWORD = password.PASSWORD
SPREADSHEET_KEY = password.SPREADSHEET_KEY
LINE_TOKEN = password.LINE_TOKEN
TO_EMAIL = password.TO_EMAIL
FROM_EMAIL = password.FROM_EMAIL
SMTP_PASSWORD = password.SMTP_PASSWORD

# print start time
dt_now = datetime.datetime.now(pytz.timezone('Asia/Tokyo'))
print(f'{dt_now} start')


options = webdriver.ChromeOptions()
options.add_argument("--headless")

browser = webdriver.Chrome(ChromeDriverManager().install(), options=options)

url = 'https://plas.soka.ac.jp/csp/plas/login.csp'
browser.get(url)

id = ID
password = PASSWORD

elem_id = browser.find_element_by_id('plas-username')
elem_pw = browser.find_element_by_id('plas-password')

elem_id.send_keys(id)
elem_pw.send_keys(password)

# ログインボタンを押す
time.sleep(1)
browser.find_elements_by_class_name('btn-area')[0].click()
time.sleep(1)

# 授業・学習・留学へ
browser.find_element_by_id('menu-header-news0111').click()
time.sleep(1)

dfs_class = scraping()

# 学生生活・その他へ
browser.find_element_by_id('menu-header-news0112').click()
time.sleep(1)

dfs_life = scraping()

# キャリア・教職・資格等へ
browser.find_element_by_id('menu-header-news0113').click()
time.sleep(1)

dfs_career = scraping()

time.sleep(2)

dfs = pd.concat([dfs_class, dfs_life, dfs_career], ignore_index=True)

# 以下google spreadsheetにアクセス
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

# 認証情報設定
# ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
credentials = ServiceAccountCredentials.from_json_keyfile_name('./portalsite-scraping-838d148686bf.json', scope)
# OAuth2の資格情報を使用してGoogle APIにログインします。
gc = gspread.authorize(credentials)

# 共有設定したスプレッドシートキーを変数[SPREADSHEET_KEY]に格納する。
SPREADSHEET_KEY = SPREADSHEET_KEY

# 共有設定したスプレッドシートのシート1を開く
worksheet = gc.open_by_key(SPREADSHEET_KEY).sheet1

# google spreadsheetをdataframeに追加
dfs_logs = pd.DataFrame(worksheet.get_all_values())
dfs_logs.columns = ['サブタイトル', '掲示期間', '連絡対象', '連絡内容', 'お知らせ区分', '問合せ先']
dfs_logs.drop(0, inplace=True)

nothing = dfs[~dfs['連絡内容'].isin(dfs_logs['連絡内容'])]

if not nothing.empty:
    try:
        send_line(nothing)
    except:
        pass
    try:
        send_email(nothing)
    except:
        pass
    try:
        num = len(dfs_logs.index) + 2

        col_lastnum = len(nothing.columns)  # DataFrameの列数
        row_lastnum = len(nothing.index)  # DataFrameの行数

        cell_list = worksheet.range('A' + str(num) + ':F' + str(row_lastnum + num - 1))
        for cell in cell_list:
            val = nothing.iloc[cell.row - num][cell.col - 1]
            cell.value = val

        worksheet.update_cells(cell_list)
    except:
        pass

else:
    pass

browser.close()

#print end time
dt_now = datetime.datetime.now(pytz.timezone('Asia/Tokyo'))
print(f'{dt_now} end')

