import pandas as pd
import re
from openpyxl import load_workbook, Workbook
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading

class EmailRecordExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("メールアドレス含有レコード抽出ツール")
        self.root.geometry("1000x700")
        
        # メールアドレスの正規表現パターン
        self.email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        self.extracted_records = []  # 抽出したレコード全体を保存
        self.df = None  # 元のDataFrame
        
        self.setup_ui()
    
    def setup_ui(self):
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # ファイル選択セクション
        file_frame = ttk.LabelFrame(main_frame, text="ファイル選択", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(file_frame, text="入力Excelファイル:").grid(row=0, column=0, sticky=tk.W)
        self.input_file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.input_file_var, width=60).grid(row=0, column=1, padx=(10, 5))
        ttk.Button(file_frame, text="参照", command=self.browse_input_file).grid(row=0, column=2)
        
        ttk.Label(file_frame, text="出力Excelファイル:").grid(row=1, column=0, sticky=tk.W, pady=(10, 0))
        self.output_file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.output_file_var, width=60).grid(row=1, column=1, padx=(10, 5), pady=(10, 0))
        ttk.Button(file_frame, text="参照", command=self.browse_output_file).grid(row=1, column=2, pady=(10, 0))
        
        # 抽出オプション
        option_frame = ttk.LabelFrame(main_frame, text="抽出オプション", padding="10")
        option_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(option_frame, text="レコード抽出", command=self.extract_records).grid(row=0, column=0, padx=(0, 10))
        
        # 重複処理オプション
        self.duplicate_var = tk.StringVar(value="keep_first")
        ttk.Label(option_frame, text="重複処理:").grid(row=0, column=1, padx=(20, 5))
        duplicate_combo = ttk.Combobox(option_frame, textvariable=self.duplicate_var, width=15)
        duplicate_combo['values'] = ('keep_first', 'keep_all', 'remove_all')
        duplicate_combo.grid(row=0, column=2)
        
        # 検索フィルター
        ttk.Label(option_frame, text="ドメインフィルター:").grid(row=1, column=0, sticky=tk.W, pady=(10, 0))
        self.domain_filter_var = tk.StringVar()
        ttk.Entry(option_frame, textvariable=self.domain_filter_var, width=20).grid(row=1, column=1, padx=(10, 0), pady=(10, 0))
        ttk.Button(option_frame, text="フィルター適用", command=self.apply_filter).grid(row=1, column=2, padx=(10, 0), pady=(10, 0))
        
        # 結果表示セクション
        result_frame = ttk.LabelFrame(main_frame, text="抽出結果", padding="10")
        result_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Treeviewを作成（動的に列を設定）
        self.tree = ttk.Treeview(result_frame, show='tree headings', height=15)
        
        # スクロールバー
        v_scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(result_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # 統計情報
        self.stats_var = tk.StringVar(value="抽出済み: 0レコード")
        ttk.Label(result_frame, textvariable=self.stats_var).grid(row=2, column=0, sticky=tk.W, pady=(10, 0))
        
        # アクションセクション
        action_frame = ttk.LabelFrame(main_frame, text="アクション", padding="10")
        action_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
        # 出力設定
        ttk.Label(action_frame, text="出力シート名:").grid(row=0, column=0, sticky=tk.W)
        self.sheet_name_var = tk.StringVar(value="メールアドレス含有レコード")
        ttk.Entry(action_frame, textvariable=self.sheet_name_var, width=25).grid(row=0, column=1, padx=(10, 20))
        
        # ボタン
        ttk.Button(action_frame, text="選択レコードを出力", command=self.save_selected_records).grid(row=0, column=2, padx=(0, 10))
        ttk.Button(action_frame, text="全レコードを出力", command=self.save_all_records).grid(row=0, column=3, padx=(0, 10))
        ttk.Button(action_frame, text="ファイル状態確認", command=self.check_file_status).grid(row=1, column=2, padx=(0, 10), pady=(5, 0))
        ttk.Button(action_frame, text="結果をクリア", command=self.clear_results).grid(row=0, column=4)
        
        # プログレスバー
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # グリッドの重みを設定
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
    
    def browse_input_file(self):
        filename = filedialog.askopenfilename(
            title="入力Excelファイルを選択",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.input_file_var.set(filename)
    
    def browse_output_file(self):
        filename = filedialog.asksaveasfilename(
            title="出力Excelファイルを選択",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.output_file_var.set(filename)
    
    def extract_records(self):
        input_file = self.input_file_var.get()
        if not input_file:
            messagebox.showerror("エラー", "入力ファイルを選択してください")
            return
        
        self.progress.start(10)
        threading.Thread(target=self._extract_records_worker, args=(input_file,), daemon=True).start()
    
    def _extract_records_worker(self, input_file):
        try:
            # Excelファイルを読み込み
            self.df = pd.read_excel(input_file)
            
            # メールアドレスを含む行を抽出
            email_rows = []
            for index, row in self.df.iterrows():
                has_email = False
                email_info = []
                
                # 各列をチェック
                for col in self.df.columns:
                    cell_value = str(row[col])
                    if cell_value != 'nan':
                        found_emails = re.findall(self.email_pattern, cell_value)
                        if found_emails:
                            has_email = True
                            email_info.extend(found_emails)
                
                if has_email:
                    # 行全体の情報を保存
                    record_info = {
                        'index': index,
                        'excel_row': index + 2,  # Excelの行番号
                        'emails': list(set(email_info)),  # 重複を除去
                        'data': row.to_dict()  # 行全体のデータ
                    }
                    email_rows.append(record_info)
            
            # 重複処理
            processed_rows = self._process_duplicates(email_rows)
            
            self.root.after(0, self._update_record_list, processed_rows)
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("エラー", f"ファイル読み込みエラー: {str(e)}"))
        finally:
            self.root.after(0, self.progress.stop)
    
    def _process_duplicates(self, email_rows):
        """重複処理のロジック"""
        duplicate_option = self.duplicate_var.get()
        
        if duplicate_option == "keep_all":
            return email_rows
        
        # メールアドレスベースで重複をチェック
        seen_emails = set()
        processed_rows = []
        
        for row in email_rows:
            row_emails = set(row['emails'])
            
            if duplicate_option == "keep_first":
                # 初回出現のみ保持
                if not any(email in seen_emails for email in row_emails):
                    processed_rows.append(row)
                    seen_emails.update(row_emails)
            
            elif duplicate_option == "remove_all":
                # 重複があるレコードは除外
                if not any(email in seen_emails for email in row_emails):
                    # 後でこのメールアドレスが他の行にもあるかチェック
                    is_unique = True
                    for other_row in email_rows:
                        if other_row['index'] != row['index']:
                            if any(email in other_row['emails'] for email in row_emails):
                                is_unique = False
                                break
                    
                    if is_unique:
                        processed_rows.append(row)
                    
                    seen_emails.update(row_emails)
        
        return processed_rows
    
    def _update_record_list(self, records):
        self.extracted_records = records
        self._refresh_tree_view()
    
    def _refresh_tree_view(self):
        # Treeviewをクリア
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        if not self.extracted_records:
            return
        
        # 列を動的に設定
        if self.df is not None:
            columns = ['Row'] + ['Emails'] + list(self.df.columns)
            self.tree['columns'] = columns
            self.tree['show'] = 'tree headings'
            
            # ヘッダーを設定
            self.tree.heading('#0', text='No.')
            self.tree.column('#0', width=50)
            
            for col in columns:
                self.tree.heading(col, text=col)
                if col == 'Emails':
                    self.tree.column(col, width=200)
                elif col == 'Row':
                    self.tree.column(col, width=50)
                else:
                    self.tree.column(col, width=120)
        
        # ドメインフィルターを適用
        domain_filter = self.domain_filter_var.get().lower().strip()
        filtered_records = []
        
        for i, record in enumerate(self.extracted_records):
            if not domain_filter or any(domain_filter in email.lower() for email in record['emails']):
                filtered_records.append(record)
                
                # Treeviewに行を追加
                emails_str = ', '.join(record['emails'])
                row_values = [record['excel_row'], emails_str]
                
                # データの各列を追加
                for col in self.df.columns:
                    value = record['data'].get(col, '')
                    if pd.isna(value):
                        value = ''
                    row_values.append(str(value))
                
                self.tree.insert('', 'end', iid=i, text=str(i+1), values=row_values)
        
        # 統計を更新
        total = len(self.extracted_records)
        filtered = len(filtered_records)
        if domain_filter:
            self.stats_var.set(f"抽出済み: {total}レコード / 表示中: {filtered}レコード")
        else:
            self.stats_var.set(f"抽出済み: {total}レコード")
    
    def apply_filter(self):
        self._refresh_tree_view()
    
    def save_selected_records(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "保存するレコードを選択してください")
            return
        
        records_to_save = []
        for item in selected_items:
            index = int(self.tree.item(item)['text']) - 1
            if 0 <= index < len(self.extracted_records):
                records_to_save.append(self.extracted_records[index])
        
        self._save_records_to_excel(records_to_save)
    
    def save_all_records(self):
        if not self.extracted_records:
            messagebox.showwarning("警告", "保存するレコードがありません")
            return
        
        # フィルタリングされたレコードを保存
        domain_filter = self.domain_filter_var.get().lower().strip()
        records_to_save = []
        
        for record in self.extracted_records:
            if not domain_filter or any(domain_filter in email.lower() for email in record['emails']):
                records_to_save.append(record)
        
        self._save_records_to_excel(records_to_save)
    
    def _save_records_to_excel(self, records):
        output_file = self.output_file_var.get()
        if not output_file:
            messagebox.showerror("エラー", "出力ファイルを指定してください")
            return
        
        try:
            # DataFrameを作成
            save_data = []
            for record in records:
                row_data = record['data'].copy()
                row_data['抽出されたメールアドレス'] = ', '.join(record['emails'])
                row_data['元の行番号'] = record['excel_row']
                save_data.append(row_data)
            
            df_save = pd.DataFrame(save_data)
            
            # Excelファイルに保存
            sheet_name = self.sheet_name_var.get() or "メールアドレス含有レコード"
            
            # ファイルが開かれているかチェック
            import os
            if os.path.exists(output_file):
                try:
                    # ファイルが使用中かテスト
                    with open(output_file, 'r+b'):
                        pass
                except IOError:
                    # ファイルが開かれている場合の対処
                    response = messagebox.askyesno(
                        "ファイルが使用中です", 
                        f"ファイル '{os.path.basename(output_file)}' が他のアプリケーションで開かれている可能性があります。\n"
                        "ファイルを閉じてから「はい」をクリックするか、\n"
                        "「いいえ」をクリックして別名で保存してください。"
                    )
                    if not response:
                        # 別名保存のダイアログを表示
                        new_file = filedialog.asksaveasfilename(
                            title="別名で保存",
                            defaultextension=".xlsx",
                            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                            initialdir=os.path.dirname(output_file),
                            initialfile=f"{os.path.splitext(os.path.basename(output_file))[0]}_new.xlsx"
                        )
                        if new_file:
                            output_file = new_file
                            self.output_file_var.set(output_file)
                        else:
                            return
            
            # 保存実行
            with pd.ExcelWriter(output_file, engine='openpyxl', mode='w') as writer:
                df_save.to_excel(writer, sheet_name=sheet_name, index=False)
            
            messagebox.showinfo("成功", f"{len(records)}レコードを {os.path.basename(output_file)} に保存しました")
            
        except PermissionError:
            messagebox.showerror(
                "権限エラー", 
                f"ファイルの保存に失敗しました。\n\n"
                f"考えられる原因:\n"
                f"• ファイルが他のアプリケーション（Excel等）で開かれている\n"
                f"• フォルダへの書き込み権限がない\n"
                f"• ファイルが読み取り専用になっている\n\n"
                f"対処法:\n"
                f"• Excelファイルを閉じてから再実行\n"
                f"• 管理者権限で実行\n"
                f"• 別の保存場所を選択"
            )
        except FileNotFoundError:
            messagebox.showerror(
                "ファイルエラー",
                f"指定されたフォルダが見つかりません。\n"
                f"保存先のフォルダが存在することを確認してください。"
            )
        except Exception as e:
            error_msg = str(e)
            if "Permission denied" in error_msg:
                messagebox.showerror(
                    "アクセス拒否エラー",
                    f"ファイルへのアクセスが拒否されました。\n\n"
                    f"• ファイルを閉じてから再実行してください\n"
                    f"• 別の保存場所を試してください\n"
                    f"• 管理者権限で実行してください\n\n"
                    f"詳細: {error_msg}"
                )
            else:
                messagebox.showerror("エラー", f"保存エラー: {error_msg}")
    
    def check_file_status(self):
        """出力ファイルの状態を確認"""
        output_file = self.output_file_var.get()
        if not output_file:
            messagebox.showwarning("警告", "出力ファイルを指定してください")
            return
        
        import os
        
        # ファイルの存在確認
        if not os.path.exists(output_file):
            messagebox.showinfo("ファイル状態", f"ファイルは存在しません。新規作成されます。\n\nパス: {output_file}")
            return
        
        # ファイルのアクセス確認
        try:
            with open(output_file, 'r+b'):
                file_size = os.path.getsize(output_file)
                file_size_mb = file_size / (1024 * 1024)
                messagebox.showinfo(
                    "ファイル状態", 
                    f"ファイルは使用可能です。\n\n"
                    f"パス: {os.path.basename(output_file)}\n"
                    f"サイズ: {file_size_mb:.2f} MB\n"
                    f"状態: 書き込み可能"
                )
        except IOError:
            messagebox.showwarning(
                "ファイル状態", 
                f"ファイルが他のアプリケーションで使用中です。\n\n"
                f"パス: {os.path.basename(output_file)}\n"
                f"状態: 使用中（書き込み不可）\n\n"
                f"対処法:\n"
                f"• Excelファイルを閉じる\n"
                f"• 別名で保存する"
            )
        except Exception as e:
            messagebox.showerror("エラー", f"ファイル状態の確認中にエラーが発生しました:\n{str(e)}")

    def clear_results(self):
        self.extracted_records = []
        self.df = None
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.stats_var.set("抽出済み: 0レコード")
        self.domain_filter_var.set("")

def main():
    root = tk.Tk()
    app = EmailRecordExtractor(root)
    root.mainloop()

if __name__ == "__main__":
    main()
