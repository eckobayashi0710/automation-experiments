# rakutenJANcore.py
# 楽天ブックス/商品APIとGoogle Sheets操作のコアロジック

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import requests
import time
import socket

class RakutenBooksFinder:
    """
    JAN/ISBNコードを元に楽天ブックスで書籍を検索し、
    見つからない場合は楽天商品市場を検索してGoogle Sheetsを更新するクラス。
    """
    # APIエンドポイントを両方定義
    RAKUTEN_BOOKS_API_URL = "https://app.rakuten.co.jp/services/api/BooksBook/Search/20170404"
    RAKUTEN_PRODUCT_API_URL = "https://app.rakuten.co.jp/services/api/IchibaItem/Search/20220601"
    
    def __init__(self, config, logger_callback=print):
        """
        コンストラクタ
        """
        self.config = config
        self.logger = logger_callback
        self._setup_services()

    def _setup_services(self):
        """APIキーや認証情報を用いて各種サービスを初期化する"""
        self.logger("🔧 サービスを初期化中...")
        socket.setdefaulttimeout(self.config.get("timeout", 120))
        try:
            scopes = ["https://www.googleapis.com/auth/spreadsheets"]
            creds = Credentials.from_service_account_file(
                self.config["json_path"], scopes=scopes
            )
            self.sheets_service = build("sheets", "v4", credentials=creds, cache_discovery=False)
            gc = gspread.authorize(creds)
            spreadsheet = gc.open_by_key(self.config["spreadsheet_id"])
            self.sheet = spreadsheet.worksheet(self.config["sheet_name"])
            self.logger("✅ Google Sheets サービス初期化完了")
        except gspread.exceptions.WorksheetNotFound:
            self.logger(f"❌ ワークシートが見つかりません: '{self.config['sheet_name']}'。")
            raise
        except Exception as e:
            self.logger(f"❌ Google Sheetsの認証または接続に失敗しました: {e}")
            raise

    def _column_letter_to_number(self, column_letter):
        """アルファベットの列文字を数値に変換 (A=1, B=2, ...)"""
        num = 0
        for char in column_letter.upper():
            num = num * 26 + (ord(char) - ord('A') + 1)
        return num

    def _call_rakuten_books_api(self, jan_code):
        """楽天ブックス書籍検索APIを呼び出す"""
        self.logger(f"  📚 楽天ブックスで検索中...")
        params = {
            "applicationId": self.config["rakuten_app_id"],
            "affiliateId": self.config.get("rakuten_affiliate_id", ""),
            "isbnJan": jan_code,
            "formatVersion": 2,
            "hits": 1
        }
        try:
            response = requests.get(self.RAKUTEN_BOOKS_API_URL, params=params, timeout=30)
            response.raise_for_status()
            data = response.json()

            if data.get("count", 0) == 0 or not data.get("Items"):
                return None

            item = data["Items"][0]
            returned_isbn = item.get("isbn", "")
            normalized_jan_code = str(jan_code).replace("-", "")
            normalized_returned_isbn = str(returned_isbn).replace("-", "")

            if normalized_returned_isbn != normalized_jan_code:
                self.logger(f"  ⚠️ ブックスAPIの検索結果コードが不一致 (返却: {returned_isbn})")
                return None
            
            # データを統一的な形式で返す
            return {
                "type": "書籍",
                "name": item.get("title", "情報なし"),
                "price": item.get("itemPrice", "情報なし"),
                "url": item.get("itemUrl", "情報なし"),
                "detail": item.get("author", "情報なし"), # 著者
                "caption": item.get("itemCaption", "情報なし"),
                "review_avg": item.get("reviewAverage", "情報なし"),
                "image_url": item.get("largeImageUrl", "情報なし").replace('?_ex=200x200', '')
            }
        except requests.exceptions.RequestException as e:
            self.logger(f"  ❌ ブックスAPIリクエストエラー: {e}")
            return None
        except Exception as e:
            self.logger(f"  ❌ ブックスAPI処理中に予期せぬエラー: {e}")
            return None

    def _call_rakuten_product_api(self, jan_code):
        """楽天商品検索APIを呼び出す"""
        self.logger(f"  🛒 楽天商品市場で検索中...")
        params = {
            "applicationId": self.config["rakuten_app_id"],
            "affiliateId": self.config.get("rakuten_affiliate_id", ""),
            "keyword": jan_code, # 商品検索ではkeywordを使う
            "formatVersion": 2,
            "hits": 1
        }
        try:
            response = requests.get(self.RAKUTEN_PRODUCT_API_URL, params=params, timeout=30)
            response.raise_for_status()
            data = response.json()

            if data.get("count", 0) == 0 or not data.get("Items"):
                return None

            item = data["Items"][0]

            # --- エラー修正箇所 ---
            # 画像URLの取得ロジックを安定化
            image_url = "情報なし"
            image_urls_list = item.get("mediumImageUrls", [])
            if image_urls_list:
                first_image_data = image_urls_list[0]
                if isinstance(first_image_data, dict):
                    image_url = first_image_data.get("imageUrl", "情報なし")
                elif isinstance(first_image_data, str):
                    image_url = first_image_data
            
            image_url = image_url.replace('?_ex=128x128', '')
            # --- 修正ここまで ---

            # データを統一的な形式で返す
            return {
                "type": "商品",
                "name": item.get("itemName", "情報なし"),
                "price": item.get("itemPrice", "情報なし"),
                "url": item.get("itemUrl", "情報なし"),
                "detail": item.get("shopName", "情報なし"), # 店舗名
                "caption": item.get("itemCaption", "情報なし"),
                "review_avg": item.get("reviewAverage", "情報なし"),
                "image_url": image_url
            }
        except requests.exceptions.RequestException as e:
            self.logger(f"  ❌ 商品APIリクエストエラー: {e}")
            return None
        except Exception as e:
            self.logger(f"  ❌ 商品API処理中に予期せぬエラー: {e}")
            return None

    def _check_and_create_headers(self):
        """出力列のヘッダーを確認し、なければ作成する"""
        try:
            self.logger("🔍 ヘッダーの確認...")
            header_row = 1 
            start_col = self.config['output_start_col_letter']
            
            first_header_cell_value = self.sheet.cell(header_row, self._column_letter_to_number(start_col)).value
            
            # ヘッダーを汎用的なものに変更
            expected_headers = [
                "種別", "名称", "価格", "URL", "詳細(著者/店舗)", 
                "商品説明", "レビュー平均", "画像URL"
            ]

            if first_header_cell_value != expected_headers[0]:
                self.logger("ℹ️ ヘッダーを作成または更新します...")
                range_to_update = f"{start_col}{header_row}"
                self.sheet.update(range_to_update, [expected_headers], value_input_option='USER_ENTERED')
                self.logger("✅ ヘッダーの作成/更新が完了しました。")
            else:
                self.logger("✅ ヘッダーは既に存在します。")
        except gspread.exceptions.APIError as e:
            self.logger(f"❌ Google Sheets APIエラー: ヘッダーの確認中に問題が発生しました。: {e}")
            raise
        except Exception as e:
            self.logger(f"⚠️ ヘッダーの確認・作成中に予期せぬエラーが発生しました: {e}")

    def _get_batch_data(self, start_row, batch_size):
        """Google Sheetsから処理対象のJAN/ISBNコードを取得する"""
        jan_col_letter = self.config["jan_col_letter"]
        output_col_letter = self.config["output_start_col_letter"]
        end_row = start_row + batch_size - 1
        
        try:
            self.logger(f"📊 データ取得中: {start_row}行目から{batch_size}行")
            range_to_get = f"{self.sheet.title}!{jan_col_letter}{start_row}:{output_col_letter}{end_row}"
            response = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=self.config["spreadsheet_id"],
                range=range_to_get
            ).execute()
            
            values = response.get('values', [])
            batch_data = []
            jan_col_index = 0
            output_col_index = self._column_letter_to_number(output_col_letter) - self._column_letter_to_number(jan_col_letter)

            for i, row_values in enumerate(values):
                jan_code = row_values[jan_col_index] if len(row_values) > jan_col_index else ""
                is_processed = len(row_values) > output_col_index and row_values[output_col_index]

                if jan_code.strip() and not is_processed:
                    batch_data.append({'row': start_row + i, 'jan': jan_code.strip()})
            
            self.logger(f"✅ データ取得完了。処理対象は{len(batch_data)}件です。")
            return batch_data
        except Exception as e:
            self.logger(f"❌ データ取得エラー: {e}")
            return []

    def _batch_update_sheets(self, update_data):
        """取得した情報をGoogle Sheetsに一括更新する"""
        if not update_data:
            return

        try:
            self.logger(f"📝 {len(update_data)}件の情報を一括更新中...")
            output_col = self.config['output_start_col_letter']
            data_to_write = []
            for item in update_data:
                # 書き込むデータを汎用的な形式に合わせる
                row_values = [
                    item['product_info']['type'],
                    item['product_info']['name'],
                    item['product_info']['price'],
                    item['product_info']['url'],
                    item['product_info']['detail'],
                    item['product_info']['caption'],
                    item['product_info']['review_avg'],
                    item['product_info']['image_url']
                ]
                data_to_write.append({
                    'range': f"{self.sheet.title}!{output_col}{item['row']}",
                    'values': [row_values]
                })
            
            body = {'valueInputOption': 'USER_ENTERED', 'data': data_to_write}
            result = self.sheets_service.spreadsheets().values().batchUpdate(
                spreadsheetId=self.config["spreadsheet_id"], body=body
            ).execute()
            
            self.logger(f"✅ 更新成功: {result.get('totalUpdatedCells', 0)}セル")
        except Exception as e:
            self.logger(f"❌ 一括更新エラー: {e}")

    def run_process(self):
        """メインの処理ループを実行する"""
        try:
            self.logger("\n🚀 楽天情報取得処理開始 🚀")
            self._check_and_create_headers()
        except Exception:
            self.logger(f"CRITICAL: ヘッダーの準備に失敗したため、処理を中止します。")
            return

        current_row = self.config["start_row"]
        consecutive_empty_batches = 0
        
        while True:
            batch_data = self._get_batch_data(current_row, self.config["batch_size"])
            
            if not batch_data:
                consecutive_empty_batches += 1
                self.logger(f"ℹ️ 行 {current_row}-{current_row + self.config['batch_size']-1} に処理対象データなし ({consecutive_empty_batches}/3)")
                if consecutive_empty_batches >= 3:
                    self.logger("🛑 3回連続で処理対象がなかったため、処理を終了します。")
                    break
                current_row += self.config["batch_size"]
                continue
            
            consecutive_empty_batches = 0
            update_data = []
            for item in batch_data:
                self.logger(f"🔍 JAN/ISBN [{item['jan']}] を検索中...")
                
                # まずブックスAPIで検索
                product_info = self._call_rakuten_books_api(item['jan'])
                
                # 見つからなければ商品APIで検索
                if not product_info:
                    product_info = self._call_rakuten_product_api(item['jan'])

                if product_info:
                    update_data.append({'row': item['row'], 'product_info': product_info})
                    self.logger(f"  => 取得成功: {str(product_info['name'])[:30]}...")
                else:
                    self.logger(f"  => 最終的に情報は見つかりませんでした。")

                time.sleep(1)

            if update_data:
                self._batch_update_sheets(update_data)
            
            current_row += self.config["batch_size"]
            self.logger(f"⏳ 次のバッチまで{self.config['api_delay']}秒待機...")
            time.sleep(self.config["api_delay"])
            
        self.logger("\n🎉 全処理完了！ 🎉")
