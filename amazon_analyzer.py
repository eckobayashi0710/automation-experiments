# amazon_analyzer.py
# Amazon商品ページのHTML構造を解析し、データ項目をJSONで出力するツール

import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog
import threading
import requests
from bs4 import BeautifulSoup
import json
import re
import time # <--- この行を追加しました

class AmazonAnalyzerApp(tk.Tk):
    """
    Amazon商品ページのHTML構造を解析するためのGUIアプリケーション。
    """
    def __init__(self):
        super().__init__()
        self.title("Amazon商品ページ HTML解析ツール")
        self.geometry("800x600")

        self.analysis_results = []

        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- URL入力エリア ---
        url_frame = ttk.LabelFrame(main_frame, text="1. 解析したいAmazon商品ページのURLを入力 (1行に1URL)", padding="10")
        url_frame.pack(fill=tk.X, pady=5)
        self.url_text = scrolledtext.ScrolledText(url_frame, wrap=tk.WORD, height=10)
        self.url_text.pack(fill=tk.X, expand=True)

        # --- 操作ボタン ---
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=5)
        self.analyze_button = ttk.Button(button_frame, text="2. 解析実行", command=self.start_analysis)
        self.analyze_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.save_button = ttk.Button(button_frame, text="3. 解析結果をJSONで保存", state=tk.DISABLED, command=self.save_results)
        self.save_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # --- ログ表示エリア ---
        log_frame = ttk.LabelFrame(main_frame, text="ログ・解析結果プレビュー", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=15)
        self.log_area.pack(fill=tk.BOTH, expand=True)

    def log(self, message):
        """ログエリアにメッセージを追記する"""
        self.log_area.insert(tk.END, message + '\n')
        self.log_area.see(tk.END)
        self.update_idletasks()

    def start_analysis(self):
        """解析処理を別スレッドで開始する"""
        urls = self.url_text.get('1.0', tk.END).strip().split('\n')
        urls = [url.strip() for url in urls if url.strip()]
        
        if not urls:
            self.log("エラー: URLが入力されていません。")
            return

        self.analyze_button.config(state=tk.DISABLED, text="解析中...")
        self.save_button.config(state=tk.DISABLED)
        self.log_area.delete('1.0', tk.END)
        self.analysis_results = []

        thread = threading.Thread(target=self.run_analysis_thread, args=(urls,))
        thread.daemon = True
        thread.start()

    def run_analysis_thread(self, urls):
        """バックグラウンドでスクレイピングと解析を実行する"""
        self.log(f"解析を開始します... 対象URL: {len(urls)}件")
        
        session = requests.Session()
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 1.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
            "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
        }
        
        for i, url in enumerate(urls):
            self.log(f"\n--- [{i+1}/{len(urls)}] URLを解析中: {url[:50]}...")
            try:
                response = session.get(url, headers=headers, timeout=20)
                response.raise_for_status()
                
                soup = BeautifulSoup(response.text, 'html.parser')
                product_data = self.analyze_html(soup)
                
                if product_data:
                    self.analysis_results.append({"url": url, "data": product_data})
                    self.log("✅ 解析成功。")
                else:
                    self.log("⚠️ ページ構造が異なるため、主要な情報を取得できませんでした。")

            except requests.exceptions.HTTPError as e:
                self.log(f"❌ HTTPエラー: {e.response.status_code} - ページが存在しないか、アクセスがブロックされた可能性があります。")
            except requests.exceptions.RequestException as e:
                self.log(f"❌ リクエストエラー: {e}")
            except Exception as e:
                self.log(f"❌ 不明なエラー: {e}")
            
            time.sleep(2) # サーバー負荷軽減

        self.log("\n--- 全てのURLの解析が完了しました ---")
        self.log("プレビュー:")
        self.log(json.dumps(self.analysis_results, indent=2, ensure_ascii=False))
        
        self.analyze_button.config(state=tk.NORMAL, text="2. 解析実行")
        if self.analysis_results:
            self.save_button.config(state=tk.NORMAL)

    def analyze_html(self, soup):
        """BeautifulSoupオブジェクトから指定された領域の情報を抽出する"""
        data = {}

        # --- 領域1: div#ppd (価格、タイトル、箇条書きなど) ---
        ppd_div = soup.find('div', id='ppd')
        if ppd_div:
            self.log("  - div#ppd を解析中...")
            
            # 商品タイトル
            title_tag = ppd_div.find('span', id='productTitle')
            data['product_title'] = title_tag.text.strip() if title_tag else 'N/A'

            # ブランド
            byline_tag = ppd_div.find('div', id='bylineInfo_feature_div')
            data['brand'] = byline_tag.text.strip() if byline_tag else 'N/A'
            
            # 価格
            price_whole = ppd_div.select_one('.a-price-whole')
            price_fraction = ppd_div.select_one('.a-price-fraction')
            data['price'] = (price_whole.text.strip() + price_fraction.text.strip()) if price_whole and price_fraction else 'N/A'
            
            # 箇条書き (商品の特徴)
            feature_bullets_ul = ppd_div.find('ul', class_='a-unordered-list a-vertical a-spacing-mini')
            if feature_bullets_ul:
                features = [li.text.strip() for li in feature_bullets_ul.find_all('li')]
                data['feature_bullets'] = features

        # --- 領域2: div.a-column.a-span12.a-span-last (詳細情報テーブル) ---
        # このクラス名は一般的すぎるため、より具体的なIDやクラスで絞り込む
        details_div = soup.find('div', id='detailBullets_feature_div') # 新しいレイアウト
        if not details_div:
            details_div = soup.find('div', id='productDetails_feature_div') # 古いレイアウト
        
        if details_div:
            self.log("  - 詳細情報セクションを解析中...")
            details = {}
            # 新しいレイアウト (箇条書き形式)
            for li in details_div.select('li'):
                key_tag = li.select_one('span.a-text-bold')
                if key_tag:
                    key = key_tag.text.replace(':', '').replace('\n', '').strip()
                    value = key_tag.find_next_sibling('span').text.strip() if key_tag.find_next_sibling('span') else ''
                    details[key] = value
            data['product_details_list'] = details
        
        # 古いレイアウト (テーブル形式)
        tech_spec_table = soup.find('table', id='productDetails_techSpec_section_1')
        if tech_spec_table:
            self.log("  - 技術仕様テーブルを解析中...")
            tech_specs = {}
            for tr in tech_spec_table.find_all('tr'):
                th = tr.find('th')
                td = tr.find('td')
                if th and td:
                    key = th.text.strip()
                    value = td.text.strip()
                    tech_specs[key] = re.sub(r'\s+', ' ', value)
            data['product_details_table'] = tech_specs
            
        return data

    def save_results(self):
        """解析結果をJSONファイルとして保存する"""
        if not self.analysis_results:
            self.log("エラー: 保存するデータがありません。")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            title="解析結果を保存",
            initialfile="amazon_analysis_results.json"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(self.analysis_results, f, indent=4, ensure_ascii=False)
                self.log(f"✅ 解析結果を {file_path} に保存しました。")
            except Exception as e:
                self.log(f"❌ ファイルの保存中にエラーが発生しました: {e}")

if __name__ == "__main__":
    app = AmazonAnalyzerApp()
    app.mainloop()
