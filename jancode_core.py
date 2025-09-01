# jancode_core.py
# jancode.xyz のスクレイピングとGoogle Sheets操作のコアロジック

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import requests
from bs4 import BeautifulSoup
import time
import socket
from urllib.parse import urljoin

class JanCodeScraper:
    """
    jancode.xyz からJANコード情報をスクレイピングし、Google Sheetsを更新するクラス。
    """
    BASE_URL = "https://www.jancode.xyz/"
    SEARCH_URL = "https://www.jancode.xyz/code/"
    
    def __init__(self, config, logger_callback=print):
        """
        コンストラクタ
        """
        self.config = config
        self.logger = logger_callback
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        })
        self._setup_gsheets()

    def _setup_gsheets(self):
        """Google Sheetsサービスを初期化する"""
        self.logger("🔧 Google Sheets サービスを初期化中...")
        try:
            scopes = ["https://www.googleapis.com/auth/spreadsheets"]
            creds = Credentials.from_service_account_file(self.config["json_path"], scopes=scopes)
            self.sheets_service = build("sheets", "v4", credentials=creds, cache_discovery=False)
            gc = gspread.authorize(creds)
            spreadsheet = gc.open_by_key(self.config["spreadsheet_id"])
            self.sheet = spreadsheet.worksheet(self.config["sheet_name"])
            self.logger("✅ Google Sheets サービス初期化完了")
        except gspread.exceptions.WorksheetNotFound:
            self.logger(f"❌ ワークシートが見つかりません: '{self.config['sheet_name']}'")
            raise
        except Exception as e:
            self.logger(f"❌ Google Sheetsの認証に失敗しました: {e}")
            raise

    def _column_letter_to_number(self, column_letter):
        """アルファベットの列文字を数値に変換"""
        num = 0
        for char in column_letter.upper():
            num = num * 26 + (ord(char) - ord('A') + 1)
        return num

    def _get_detail_page_urls(self, jan_codes):
        """一括検索を行い、詳細ページのURLリストを取得する"""
        self.logger(f"🔍 {len(jan_codes)}件のJANコードを一括検索中...")
        payload = {
            "q": "\n".join(jan_codes),
            "action": "search",
            "process": "code_multi"
        }
        try:
            response = self.session.post(self.SEARCH_URL, data=payload, timeout=60)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')
            
            urls = []
            for a_tag in soup.select(".result-box-out > a"):
                href = a_tag.get('href')
                if href:
                    full_url = urljoin(self.BASE_URL, href)
                    urls.append(full_url)
            
            self.logger(f"✅ 詳細ページURLを{len(urls)}件取得しました。")
            return urls
        except requests.exceptions.RequestException as e:
            self.logger(f"❌ 一括検索リクエスト中にエラーが発生しました: {e}")
            return []

    def _scrape_detail_page(self, url):
        """詳細ページから全情報を抽出する"""
        try:
            response = self.session.get(url, timeout=60)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')
            
            info = {}
            table = soup.find('table', class_='table-block')
            if not table:
                return None

            for row in table.find_all('tr'):
                th = row.find('th')
                td = row.find('td')
                if th and td:
                    key = th.text.strip()
                    # --- ▼キー名をヘッダーと統一▼ ---
                    if key == "商品イメージ":
                        img_tag = td.find('img')
                        info["商品イメージURL"] = urljoin(self.BASE_URL, img_tag['src']) if img_tag and 'src' in img_tag.attrs else ''
                    elif key == "価格調査":
                        links = {a.find('img')['src'].split('/')[-1].split('.')[0]: a['href'] for a in td.select('a') if a.find('img')}
                        info["楽天URL"] = links.get('rakuten', '')
                        info["YahooURL"] = links.get('yahoo', '')
                        info["AmazonURL"] = links.get('amazon', '')
                    elif key == "JANシンボル":
                        img_tag = td.find('img')
                        info["JANシンボル画像URL"] = urljoin(self.BASE_URL, img_tag['src']) if img_tag and 'src' in img_tag.attrs else ''
                    # --- ▲ここまで修正▲ ---
                    elif key == "商品ジャンル":
                        info[key] = " > ".join([a.text.strip() for a in td.find_all('a')])
                    else:
                        info[key] = td.text.strip()
            
            info["詳細ページURL"] = url
            return info
        except requests.exceptions.RequestException as e:
            self.logger(f"❌ 詳細ページ取得エラー ({url}): {e}")
            return None
        except Exception as e:
            self.logger(f"❌ 詳細ページ解析エラー ({url}): {e}")
            return None

    def _check_and_create_headers(self):
        """出力列のヘッダーを確認し、なければ作成する"""
        self.logger("🔍 ヘッダーの確認...")
        try:
            # --- ▼ここから修正▼ ---
            # GUIで指定された出力開始列を取得
            start_col_letter = self.config['output_start_col_letter'].upper()
            start_col_num = self._column_letter_to_number(start_col_letter)

            # ヘッダーとして期待される値のリスト（出力列のみ）
            expected_headers = [
                "商品名", "会社名", "会社名カナ", "商品ジャンル", "コードタイプ",
                "商品イメージURL", "JANシンボル画像URL", "楽天URL", "YahooURL",
                "AmazonURL", "詳細ページURL"
            ]
            
            # スプレッドシートの1行目全体の値を取得
            current_headers_all = self.sheet.row_values(1)
            
            # 必要な長さに足りない場合は空文字で埋める
            if len(current_headers_all) < start_col_num + len(expected_headers):
                padding_needed = (start_col_num + len(expected_headers)) - len(current_headers_all)
                current_headers_all.extend([""] * padding_needed)
            
            # 実際のヘッダー部分をスライスで取得
            actual_headers_slice = current_headers_all[start_col_num - 1 : start_col_num - 1 + len(expected_headers)]

            # ヘッダーが期待通りか比較
            if actual_headers_slice != expected_headers:
                self.logger("ℹ️ ヘッダーを作成または更新します...")
                
                # 更新範囲を正しく指定
                range_to_update = f"{start_col_letter}1"
                self.sheet.update(range_to_update, [expected_headers], value_input_option='USER_ENTERED')
                
                self.logger("✅ ヘッダーの作成/更新が完了しました。")
            else:
                self.logger("✅ ヘッダーは既に存在します。")
            # --- ▲ここまで修正▲ ---

        except Exception as e:
            self.logger(f"⚠️ ヘッダーの確認・作成中にエラーが発生しました: {e}")
            raise

    def run_process(self):
        """メインの処理ループを実行する"""
        try:
            self.logger("\n🚀 jancode.xyz スクレイピング処理開始 🚀")
            self._check_and_create_headers()
        except Exception as e:
            self.logger(f"CRITICAL: ヘッダーの準備に失敗したため、処理を中止します。: {e}")
            return

        current_row = self.config["start_row"]
        batch_size = self.config.get("batch_size", 100)
        
        while True:
            self.logger(f"\n--- {current_row}行目からのバッチ処理を開始 ---")
            try:
                jan_col_num = self._column_letter_to_number(self.config["jan_col_letter"])
                jan_codes_raw = self.sheet.col_values(jan_col_num)[current_row - 1 : current_row - 1 + batch_size]
                jan_codes_to_process = [code for code in jan_codes_raw if code.strip()]
            except Exception as e:
                self.logger(f"❌ スプレッドシートからのデータ取得に失敗しました: {e}")
                break

            if not jan_codes_to_process:
                self.logger("ℹ️ 処理対象のJANコードが見つかりませんでした。処理を終了します。")
                break

            detail_urls = self._get_detail_page_urls(jan_codes_to_process)
            time.sleep(self.config.get("delay", 3))

            if not detail_urls:
                current_row += len(jan_codes_raw) if jan_codes_raw else batch_size
                continue

            all_scraped_data = {}
            for url in detail_urls:
                self.logger(f"📄 詳細ページをスクレイピング中: {url}")
                scraped_info = self._scrape_detail_page(url)
                if scraped_info and "コード番号" in scraped_info:
                    all_scraped_data[scraped_info["コード番号"]] = scraped_info
                time.sleep(self.config.get("delay", 3))

            if all_scraped_data:
                self.logger(f"📝 {len(all_scraped_data)}件のデータをスプレッドシートに書き込みます...")
                update_requests = []
                
                for i, jan_in_sheet in enumerate(jan_codes_raw):
                    if jan_in_sheet in all_scraped_data:
                        data = all_scraped_data[jan_in_sheet]
                        row_to_write = current_row + i
                        # --- ▼書き込む値のキーをヘッダーと統一▼ ---
                        values = [
                            data.get("商品名", ""), data.get("会社名", ""), data.get("会社名カナ", ""),
                            data.get("商品ジャンル", ""), data.get("コードタイプ", ""),
                            data.get("商品イメージURL", ""), data.get("JANシンボル画像URL", ""),
                            data.get("楽天URL", ""), data.get("YahooURL", ""), data.get("AmazonURL", ""),
                            data.get("詳細ページURL", "")
                        ]
                        # --- ▲ここまで修正▲ ---
                        range_str = f"{self.config['output_start_col_letter']}{row_to_write}"
                        update_requests.append({"range": range_str, "values": [values]})
                
                if update_requests:
                    self.sheet.batch_update(update_requests, value_input_option='USER_ENTERED')
                    self.logger("✅ 書き込み完了。")

            current_row += len(jan_codes_raw) if jan_codes_raw else batch_size
        
        self.logger("\n🎉 全処理完了！ 🎉")
