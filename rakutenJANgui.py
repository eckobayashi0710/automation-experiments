# rakutenJANgui.py
# 設定用GUIと処理実行のトリガー

import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import threading
import os
import json

# rakutenJANcore.pyのRakutenBooksFinderクラスをインポート
try:
    # 呼び出すファイル名をrakutenJANcoreに変更
    from rakutenJANcore import RakutenBooksFinder
except ImportError:
    # エラーメッセージのファイル名を変更
    print("エラー: rakutenJANcore.py が同じディレクトリに見つかりません。")
    exit()

class App(tk.Tk):
    CONFIG_FILE = "config.json"

    def __init__(self):
        super().__init__()
        self.title("楽天ブックス JAN/ISBN検索ツール")
        self.geometry("650x600")

        # 設定値を保持するTkinter変数
        self.config_vars = {
            "rakuten_app_id": tk.StringVar(),
            "rakuten_affiliate_id": tk.StringVar(),
            "json_path": tk.StringVar(),
            "spreadsheet_id": tk.StringVar(),
            "sheet_name": tk.StringVar(),
            "jan_col_letter": tk.StringVar(value="A"),
            "output_start_col_letter": tk.StringVar(value="B"),
            "start_row": tk.IntVar(value=2),
            "batch_size": tk.IntVar(value=50),
            "api_delay": tk.DoubleVar(value=3.0),
        }

        self.create_widgets()
        self.load_config()

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        """GUIのウィジェットを作成・配置する"""
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- 設定フレーム ---
        settings_frame = ttk.LabelFrame(main_frame, text="設定", padding="10")
        settings_frame.pack(fill=tk.X, expand=True)

        # 楽天API設定
        ttk.Label(settings_frame, text="楽天Application ID:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(settings_frame, textvariable=self.config_vars["rakuten_app_id"], width=50).grid(row=0, column=1, sticky=tk.EW)
        
        ttk.Label(settings_frame, text="楽天Affiliate ID (任意):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(settings_frame, textvariable=self.config_vars["rakuten_affiliate_id"], width=50).grid(row=1, column=1, sticky=tk.EW)

        # Google API設定
        ttk.Label(settings_frame, text="Google認証JSON:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        json_frame = ttk.Frame(settings_frame)
        json_frame.grid(row=2, column=1, sticky=tk.EW)
        ttk.Entry(json_frame, textvariable=self.config_vars["json_path"], width=40).pack(side=tk.LEFT, expand=True, fill=tk.X)
        ttk.Button(json_frame, text="参照...", command=self.browse_json).pack(side=tk.LEFT, padx=(5,0))

        # スプレッドシート設定
        ttk.Label(settings_frame, text="スプレッドシートID:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(settings_frame, textvariable=self.config_vars["spreadsheet_id"]).grid(row=3, column=1, sticky=tk.EW)

        ttk.Label(settings_frame, text="シート名:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(settings_frame, textvariable=self.config_vars["sheet_name"]).grid(row=4, column=1, sticky=tk.EW)

        # 列設定
        col_frame = ttk.Frame(settings_frame)
        col_frame.grid(row=5, column=1, sticky=tk.EW)
        ttk.Label(settings_frame, text="列 (JAN/ISBNコード/出力開始):").grid(row=5, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(col_frame, textvariable=self.config_vars["jan_col_letter"], width=5).pack(side=tk.LEFT)
        ttk.Label(col_frame, text="/").pack(side=tk.LEFT, padx=5)
        ttk.Entry(col_frame, textvariable=self.config_vars["output_start_col_letter"], width=5).pack(side=tk.LEFT)
        ttk.Label(col_frame, text="(タイトル,価格,URL等の順に出力)").pack(side=tk.LEFT, padx=10)


        # 行・バッチ設定
        row_frame = ttk.Frame(settings_frame)
        row_frame.grid(row=6, column=1, sticky=tk.EW)
        ttk.Label(settings_frame, text="開始行/バッチサイズ:").grid(row=6, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(row_frame, textvariable=self.config_vars["start_row"], width=5).pack(side=tk.LEFT)
        ttk.Label(row_frame, text="/").pack(side=tk.LEFT, padx=5)
        ttk.Entry(row_frame, textvariable=self.config_vars["batch_size"], width=5).pack(side=tk.LEFT)

        settings_frame.columnconfigure(1, weight=1)

        # --- 実行ボタン ---
        self.run_button = ttk.Button(main_frame, text="処理実行", command=self.start_process)
        self.run_button.pack(pady=10, fill=tk.X)

        # --- ログ表示エリア ---
        log_frame = ttk.LabelFrame(main_frame, text="ログ", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=15)
        self.log_area.pack(fill=tk.BOTH, expand=True)

    def browse_json(self):
        """「参照」ボタンでJSONファイルを選択するダイアログを開く"""
        file_path = filedialog.askopenfilename(
            title="Google Service Account JSONを選択",
            filetypes=[("JSON files", "*.json")]
        )
        if file_path:
            self.config_vars["json_path"].set(file_path)

    def log(self, message):
        """ログエリアにメッセージを追記する"""
        if self.log_area:
            self.log_area.insert(tk.END, message + '\n')
            self.log_area.see(tk.END)
            self.update_idletasks()

    def load_config(self):
        """config.jsonから設定を読み込む"""
        if os.path.exists(self.CONFIG_FILE):
            try:
                with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
                
                for key, value in config_data.items():
                    if key in self.config_vars:
                        self.config_vars[key].set(value)
                self.log("ℹ️ 保存された設定を読み込みました。")
            except (json.JSONDecodeError, IOError) as e:
                self.log(f"⚠️ 設定ファイルの読み込みに失敗しました: {e}")
        else:
            self.log("ℹ️ 設定ファイルが見つかりません。")

    def save_config(self):
        """現在の設定をconfig.jsonに保存する"""
        config_data = {key: var.get() for key, var in self.config_vars.items()}
        try:
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=4, ensure_ascii=False)
            self.log("💾 設定を config.json に保存しました。")
        except IOError as e:
            self.log(f"⚠️ 設定の保存に失敗しました: {e}")

    def on_closing(self):
        """ウィンドウを閉じる際に設定を保存する"""
        self.save_config()
        self.destroy()

    def start_process(self):
        """処理開始ボタンが押されたときの処理"""
        config = {key: var.get() for key, var in self.config_vars.items()}

        # 必須項目チェック
        if not all([config["rakuten_app_id"], config["json_path"], config["spreadsheet_id"]]):
            self.log("❌ エラー: 楽天AppID、認証JSON、スプレッドシートIDは必須です。")
            return
        
        if not os.path.exists(config["json_path"]):
            self.log(f"❌ エラー: 認証JSONファイルが見つかりません: {config['json_path']}")
            return

        self.save_config()

        self.run_button.config(state=tk.DISABLED, text="処理中...")
        self.log_area.delete('1.0', tk.END)

        # 処理を別スレッドで実行してGUIが固まるのを防ぐ
        thread = threading.Thread(target=self.run_in_thread, args=(config,))
        thread.daemon = True
        thread.start()

    def run_in_thread(self, config):
        """バックグラウンドでコアロジックを実行する"""
        try:
            finder = RakutenBooksFinder(config, logger_callback=self.log)
            finder.run_process()
        except Exception as e:
            self.log(f"\nCRITICAL ERROR: 予期せぬエラーが発生しました。\n{e}")
        finally:
            self.run_button.config(state=tk.NORMAL, text="処理実行")

if __name__ == "__main__":
    app = App()
    app.mainloop()
