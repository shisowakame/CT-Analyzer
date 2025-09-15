import os
import datetime
import pandas as pd

class Exporter:
    def export_history_to_excel(self, history_data):
        try:
            # DataFrameの準備
            columns = ['No.']
            for i in range(self.series_count):
                columns.extend([f'Folder{i+1} Mean', f'Folder{i+1} Std Dev', f'Folder{i+1} Info'])
            
            data_rows = []
            for r_idx, row_data in enumerate(history_data):
                flat_row = [r_idx + 1]
                for series_stat in row_data:
                    flat_row.extend([series_stat.get('mean', ''), series_stat.get('std', ''), series_stat.get('info', '')])
                data_rows.append(flat_row)
            
            df = pd.DataFrame(data_rows, columns=columns)
            # info列以外をfloat型に変換
            for col in df.columns:
                if col.startswith('Folder') and not col.endswith('Info'):
                    df[col] = pd.to_numeric(df[col], errors='coerce')

            # --- 平均行の追加 ---
            # 2行空白行を挿入
            empty_row = [''] * len(columns)
            df_blank = pd.DataFrame([empty_row, empty_row], columns=columns)
            # 平均タイトル行（info列を除外）
            mean_title_no_info = [col for col in columns if not col.endswith('Info')]
            df_mean_title = pd.DataFrame([mean_title_no_info], columns=mean_title_no_info)
            # 平均値行（info列を除外）
            mean_values_no_info = ['']
            for i in range(self.series_count):
                mean_col = f'Folder{i+1} Mean'
                std_col = f'Folder{i+1} Std Dev'
                mean_val = df[mean_col].mean() if not df[mean_col].isnull().all() else ''
                std_val = df[std_col].mean() if not df[std_col].isnull().all() else ''
                mean_values_no_info.extend([
                    round(mean_val, 4) if mean_val != '' else '',
                    round(std_val, 4) if std_val != '' else ''
                ])
            df_mean_values = pd.DataFrame([mean_values_no_info], columns=mean_title_no_info)
            # 結合
            df_final = pd.concat([df, df_blank, df_mean_title, df_mean_values], ignore_index=True)
            
            # 現在の日時を取得し、ファイル名を生成
            now = datetime.datetime.now()
            # YYYYMMDD_HHMMSS の形式でフォーマット
            timestamp = now.strftime("%Y%m%d_%H%M%S") 
            file_name = f"ROI-history-{timestamp}.xlsx"
            
            # スクリプトが実行されているディレクトリに保存
            #script_dir = os.path.dirname(os.path.abspath(__file__))
            #file_path = os.path.join(script_dir, file_name)
            
            # ダウンロードフォルダに保存
            downloads_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
            file_path = os.path.join(downloads_dir, file_name)

            df_final.to_excel(file_path, index=False, engine='openpyxl')
            
            return {'success': True, 'filePath': file_path}
        except Exception as e:
            # エラーの詳細をメッセージに含める
            return {'success': False, 'message': f"ファイルの保存中にエラーが発生しました: {str(e)}"}
