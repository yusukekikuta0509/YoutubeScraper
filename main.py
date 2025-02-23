import csv
import os
import re
import time
from urllib.parse import quote

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

##############################################################################
# .env から読み込むための設定
##############################################################################
from dotenv import load_dotenv
load_dotenv()

##############################################################################
# gspread 用のインポート
##############################################################################
import gspread
from oauth2client.service_account import ServiceAccountCredentials

##############################################################################
# 0. ユーザーから検索キーワードをカンマ区切りで入力してもらう関数
##############################################################################
def get_user_keywords():
    """
    ユーザーにカンマ区切りでキーワードを入力してもらい、
    ["切り抜き", "日常"] のようなリストにして返す。
    """
    raw_input_str = input("検索キーワードをカンマ区切りで入力してください（例: 切り抜き,日常）: ")
    # カンマで分割し、前後の空白をtrim
    keywords_list = [k.strip() for k in raw_input_str.split(",") if k.strip()]
    if not keywords_list:
        print("[WARN] キーワードが入力されませんでした。処理を終了します。")
        exit(0)
    return keywords_list


##############################################################################
# 1. ブラウザ関連
##############################################################################
def setup_browser():
    """ブラウザのセットアップ (ヘッドレス推奨)"""
    options = webdriver.ChromeOptions()
    # デバッグしたい時は以下をコメントアウトしてください
    # options.add_argument("--headless")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

def wait_with_message(seconds, message):
    """指定秒数待機しつつログを出す"""
    print(f"[DEBUG] {message} → {seconds}秒待機します。")
    time.sleep(seconds)

def switch_to_new_tab(driver):
    """最後に開いたタブへ切り替える"""
    try:
        tabs = driver.window_handles
        driver.switch_to.window(tabs[-1])
        print("[DEBUG] 新タブに切り替えました。")
    except Exception as e:
        print(f"[ERROR] 新タブ切り替え失敗: {e}")

def switch_to_first_tab(driver):
    """最初のタブへ戻る"""
    try:
        tabs = driver.window_handles
        driver.switch_to.window(tabs[0])
        print("[DEBUG] 最初のタブに戻りました。")
    except Exception as e:
        print(f"[ERROR] タブ切り替え(最初)失敗: {e}")


##############################################################################
# 2. CSV初期化・保存
##############################################################################
def initialize_csv(csv_file):
    """CSVファイルの初期化 (ヘッダー作成)"""
    if not os.path.exists(csv_file):
        with open(csv_file, mode='w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            # 今回は [チャンネル名, チャンネルID, キーワード, メールアドレス] の4カラム
            writer.writerow(["ChannelName", "ChannelID", "Keyword", "Email"])
        print(f"[DEBUG] CSVファイル {csv_file} を新規作成しました。")

def save_to_csv(csv_file, rows):
    """
    rows = [[channel_name, channel_id, keyword, email], ...]
    """
    with open(csv_file, mode='a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerows(rows)
        print(f"[DEBUG] CSVに {len(rows)} 行を書き込みました: {csv_file}")


##############################################################################
# 3. 検索一覧ページから全チャンネルIDを取得
##############################################################################
def get_all_channel_ids_on_page(driver):
    """
    検索一覧ページ中にある '@...' の pタグをすべて取り出してリストで返す。
    例: p要素で '@' を含むものを全部取得
    """
    channel_ids = []
    try:
        p_elems = driver.find_elements(By.XPATH, "//p[contains(text(),'@')]")
        for p in p_elems:
            raw_id = p.text.strip()
            if raw_id.startswith("@"):
                channel_ids.append(raw_id)
        print(f"[DEBUG] 一覧ページでチャンネルIDを合計 {len(channel_ids)} 個検出: {channel_ids}")
    except Exception as e:
        print(f"[ERROR] get_all_channel_ids_on_page 失敗: {e}")
    return channel_ids


##############################################################################
# 4. アナリティクスページURL生成 → 新タブで開く
##############################################################################
def open_analytics_page(driver, channel_id):
    """
    チャンネルID (@雑学博士 など) から
    https://www.viewstats.com/@(エンコードされた雑学博士)/channelytics
    を生成し、新タブで開く
    """
    try:
        if not channel_id:
            print("[WARN] channel_id が空です。")
            return False

        if not channel_id.startswith("@"):
            channel_id = "@" + channel_id

        encoded_id = quote(channel_id, encoding='utf-8')
        analytics_url = f"https://www.viewstats.com/{encoded_id}/channelytics"

        driver.execute_script(f"window.open('{analytics_url}','_blank');")
        print(f"[DEBUG] アナリティクスURL: {analytics_url} → 新タブでオープン")
        wait_with_message(3, "アナリティクスページを開くまで待機")
        return True

    except Exception as e:
        print(f"[ERROR] open_analytics_page 失敗: {e}")
        return False


##############################################################################
# 5. アナリティクスページで No data found チェック & YouTube遷移
##############################################################################
def check_no_data_found(driver):
    """
    アナリティクスページに 'No data found' が表示されていれば True
    """
    try:
        body_text = driver.find_element(By.TAG_NAME, "body").text
        if "No data found" in body_text:
            print("[DEBUG] No data found → スキップ")
            return True
        return False
    except Exception as e:
        print(f"[ERROR] check_no_data_found 失敗: {e}")
        return False

def open_youtube_tab(driver):
    """
    アナリティクスページにある YouTube へのリンクをクリックして新しいタブで遷移
    """
    try:
        link_xpath = "/html/body/main/div/div[2]/div[1]/div/div[1]/div[2]/div[2]/a"
        link_elem = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, link_xpath))
        )
        driver.execute_script("arguments[0].click();", link_elem)
        print("[DEBUG] YouTubeリンクをクリックしました。")
        wait_with_message(2, "YouTubeタブへ移行待ち")
        return True
    except Exception as e:
        print(f"[ERROR] open_youtube_tab失敗: {e}")
        return False


##############################################################################
# 6. YouTubeページで情報取得 (チャンネル名(H2), メールアドレスなど)
##############################################################################
def get_youtube_channel_name(driver):
    """
    YouTubeのチャンネル名(H2)を取得 (XPath は YouTubeのUI変更により要調整の可能性あり)
    """
    try:
        # 下記は例として h2タグを想定したXPATH (実際のYouTubeの変更によっては要調整)
        h2_xpath = (
            "/html/body/ytd-app/div[1]/ytd-page-manager/ytd-browse/div[3]/"
            "ytd-tabbed-page-header/tp-yt-app-header-layout/div/tp-yt-app-header/"
            "div[2]/div/div[2]/yt-page-header-renderer/yt-page-header-view-model/"
            "div/div[1]/div/yt-dynamic-text-view-model/h2"
        )
        h2_elem = WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.XPATH, h2_xpath))
        )
        channel_name = h2_elem.text.strip()
        print(f"[DEBUG] YouTubeチャンネル名(H2): {channel_name}")
        return channel_name
    except Exception as e:
        print(f"[WARN] get_youtube_channel_name失敗: {e}")
        return "YouTubeチャンネル不明"

def click_youtube_show_more(driver):
    """
    YouTubeの「さらに表示」ボタンをクリック
    """
    try:
        show_btn_xpath = (
            "/html/body/ytd-app/div[1]/ytd-page-manager/ytd-browse/div[3]/"
            "ytd-tabbed-page-header/tp-yt-app-header-layout/div/tp-yt-app-header/"
            "div[2]/div/div[2]/yt-page-header-renderer/yt-page-header-view-model/"
            "div/div[1]/div/yt-description-preview-view-model/truncated-text/button"
        )
        show_btn = WebDriverWait(driver, 8).until(
            EC.element_to_be_clickable((By.XPATH, show_btn_xpath))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", show_btn)
        driver.execute_script("arguments[0].click();", show_btn)
        print("[DEBUG] YouTube さらに表示ボタンをクリック")
        wait_with_message(3, "popup表示待ち")
        return True
    except Exception as e:
        print(f"[ERROR] click_youtube_show_more失敗: {e}")
        return False

def get_youtube_about_text(driver):
    """
    YouTubeチャンネルの概要テキストを取得
    """
    try:
        about_xpath = (
            "/html/body/ytd-app/ytd-popup-container/tp-yt-paper-dialog/"
            "ytd-engagement-panel-section-list-renderer/div[2]/ytd-section-list-renderer/"
            "div[2]/ytd-item-section-renderer/div[3]/ytd-about-channel-renderer/div/"
            "yt-attributed-string/span"
        )
        about_elem = WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.XPATH, about_xpath))
        )
        raw_text = about_elem.text.strip()
        print(f"[DEBUG] YouTube概要テキスト: {raw_text}")
        return raw_text
    except Exception as e:
        print(f"[WARN] get_youtube_about_text失敗: {e}")
        return ""

def parse_email_from_text(raw_text):
    """
    raw_text からメールアドレスを抽出 (先頭1件のみ)
    """
    pattern = r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}"
    matches = re.findall(pattern, raw_text)
    if matches:
        return matches[0]
    return ""


##############################################################################
# 7. Googleスプレッドシートにアップロードする関数
##############################################################################
def upload_csv_to_google_spreadsheet(
    csv_file, 
    spreadsheet_key,
    sheet_name,
    credentials_json="credentials.json"
):
    """
    指定されたcsvファイルをGoogleスプレッドシートにアップロードする。
    既に同名シートが存在する場合は削除し、新規シートを作成して全データを書き込む。
    
    :param csv_file: アップロードするCSVファイルパス
    :param spreadsheet_key: GoogleスプレッドシートのID
    :param sheet_name: アップロード先シート名
    :param credentials_json: サービスアカウントの認証JSONファイル名
    """
    # OAuth用のスコープ
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive",
    ]
    # 認証
    creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_json, scope)
    client = gspread.authorize(creds)

    # スプレッドシートを開く
    sh = client.open_by_key(spreadsheet_key)

    # 同名のシートがあれば削除する
    try:
        worksheet = sh.worksheet(sheet_name)
        sh.del_worksheet(worksheet)
        print(f"[DEBUG] 既存シート '{sheet_name}' を削除しました。")
    except gspread.exceptions.WorksheetNotFound:
        print(f"[INFO] シート '{sheet_name}' は存在しません。新規作成に進みます。")

    # 新しいワークシートを作成
    worksheet = sh.add_worksheet(title=sheet_name, rows="1000", cols="20")
    print(f"[DEBUG] 新規シート '{sheet_name}' を作成しました。")

    # CSVファイルを読み込み
    with open(csv_file, 'r', encoding='utf-8') as f:
        csv_data = list(csv.reader(f))

    # A1セルから書き込み
    worksheet.update('A1', csv_data)
    print(f"[DEBUG] スプレッドシートに '{csv_file}' の内容をアップロードしました。")


##############################################################################
# 8. メイン処理 (スクレイピング)
##############################################################################
def scrape_viewstats():
    # (0) CSV初期化と、ユーザー入力からキーワードリストを取得
    csv_file = "viewstats_data.csv"
    initialize_csv(csv_file)
    keywords = get_user_keywords()

    # ブラウザセットアップ
    driver = setup_browser()

    # キーワードごとに処理
    for keyword in keywords:
        print(f"\n[INFO] 現在のキーワード: '{keyword}'")

        # ページ1～3まで巡回 (必要に応じて変更可能)
        no_result_for_this_keyword = False
        for page in range(1, 4):
            try:
                print(f"[INFO] キーワード '{keyword}' ページ {page} を処理")
                url = f"https://www.viewstats.com/?page={page}&q={keyword}"
                driver.get(url)

                # ページが表示されるまで待機 (ボタン or 本文を待つ)
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, "body"))
                )
                wait_with_message(2, "検索結果ページを読み込み完了")

                # (A) 「No result found」 チェック
                body_text = driver.find_element(By.TAG_NAME, "body").text
                if "No result found" in body_text:
                    print("[WARN] 'No result found' が表示されました。→ 次のキーワードへスキップ")
                    no_result_for_this_keyword = True
                    break  # すぐに次のキーワードへ

                # (B) 通常の処理: チャンネルIDを取得
                channel_ids = get_all_channel_ids_on_page(driver)
                if not channel_ids:
                    print(f"[WARN] ページ {page} でチャンネルIDを取得できませんでした。")
                    continue

                print(f"[INFO] ページ {page} で {len(channel_ids)} 個のID: {channel_ids}")

                # (C) 各チャンネルIDに対してループ
                for idx, channel_id in enumerate(channel_ids):
                    print(f"[INFO] ★★★ {keyword}: ページ {page}, アカウント {idx+1}/{len(channel_ids)} ★★★")

                    # 1) アナリティクスページを新タブで開く
                    success_analytics = open_analytics_page(driver, channel_id)
                    if not success_analytics:
                        print("[WARN] アナリティクスページを開けませんでした。スキップ")
                        continue

                    try:
                        WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)
                        switch_to_new_tab(driver)
                        wait_with_message(2, "アナリティクスページへ切り替え完了")

                        # 2) 'No data found' チェック
                        if check_no_data_found(driver):
                            driver.close()
                            switch_to_first_tab(driver)
                            wait_with_message(2, "No data found → スキップ")
                            continue

                        # 3) YouTubeページへ遷移
                        success_yt = open_youtube_tab(driver)
                        if not success_yt:
                            print("[WARN] YouTubeリンクを開けませんでした。スキップ")
                            driver.close()
                            switch_to_first_tab(driver)
                            continue

                        WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 2)
                        switch_to_new_tab(driver)
                        wait_with_message(2, "YouTubeタブに切り替え完了")

                        # 4) チャンネル名(H2) & メールアドレス取得
                        channel_name = get_youtube_channel_name(driver)  # H2
                        click_youtube_show_more(driver)
                        about_text = get_youtube_about_text(driver)
                        email = parse_email_from_text(about_text)

                        # CSVに書き込み & 即時スプレッドシートへアップロード
                        # (※ 頻回に呼び出すとAPIコールが多くなるので注意)
                        row = [[channel_name, channel_id, keyword, email]]
                        save_to_csv(csv_file, row)
                        save_to_csv_and_update_sheet(csv_file, row)

                        # YouTubeタブを閉じ → アナリティクスページに戻る
                        driver.close()
                        switch_to_new_tab(driver)
                        wait_with_message(2, "YouTubeタブを閉じてアナリティクスに戻りました")

                        # アナリティクスページを閉じ → 検索一覧へ戻る
                        driver.close()
                        switch_to_first_tab(driver)
                        wait_with_message(2, "アナリティクスページを閉じて検索一覧へ戻りました")

                    except Exception as e:
                        print(f"[ERROR] アカウント処理例外: {e}")
                        # タブが残っていれば全て閉じて検索一覧(最初のタブ)に戻す
                        while len(driver.window_handles) > 1:
                            driver.switch_to.window(driver.window_handles[-1])
                            driver.close()
                        switch_to_first_tab(driver)

                        # CSVにデフォルト行を記録 (チャンネル取得失敗)
                        row = [[f"チャンネル取得失敗", channel_id, keyword, ""]]
                        save_to_csv(csv_file, row)

            except Exception as e:
                print(f"[ERROR] ページ {page} の処理でエラー: {e}")
                continue

        # キーワードごとのページ巡回を終えて、結果が無ければ次のキーワードへ
        if no_result_for_this_keyword:
            continue

    # 全キーワード処理終了
    driver.quit()
    print("[DONE] スクレイピング完了です。")

    # スクレイピング後にCSVをスプレッドシートへアップロードする
    # (必要に応じて .env で SPREADSHEET_KEY, SHEET_NAME, CREDENTIALS_JSON をセット)
    spreadsheet_key = os.getenv("SPREADSHEET_KEY")
    sheet_name = os.getenv("SHEET_NAME", "viewstats_result")  # デフォルト名
    credentials_json = os.getenv("CREDENTIALS_JSON", "credentials.json")
    upload_csv_to_google_spreadsheet(csv_file, spreadsheet_key, sheet_name, credentials_json)


##############################################################################
# 追加: CSVに1行追加して即座にシートを更新する関数の例
##############################################################################
def save_to_csv_and_update_sheet(csv_file, row):
    """
    CSVに新しい行を追加し、Googleスプレッドシートを更新するユーティリティ関数例。
    """
    # row は [[col1, col2, ...], [col1, col2, ...]] の2次元リストを想定
    with open(csv_file, mode='a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerows(row)
        print(f"[DEBUG] CSVに {len(row)} 行を書き込みました: {row}")

    # Googleスプレッドシートを更新
    spreadsheet_key = os.getenv("SPREADSHEET_KEY")
    sheet_name = os.getenv("SHEET_NAME", "viewstats_result")
    credentials_json = os.getenv("CREDENTIALS_JSON", "credentials.json")

    time.sleep(5)  # 速すぎるとAPI制限に引っかかるかもしれないので適宜調整してください

    upload_csv_to_google_spreadsheet(csv_file, spreadsheet_key, sheet_name, credentials_json)


##############################################################################
# エントリーポイント
##############################################################################
if __name__ == "__main__":
    scrape_viewstats()
