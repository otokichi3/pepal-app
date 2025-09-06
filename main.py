import csv
import logging
from datetime import datetime
import os
import json
import uuid
import io
import sys
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Google Sheets APIの設定
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# ログメッセージをキャプチャするためのリスト
captured_logs = []

def load_config(config_file='config.json'):
    """設定ファイルを読み込む"""
    try:
        if not os.path.exists(config_file):
            logging.error(f"設定ファイル '{config_file}' が見つかりません")
            return None
        
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        logging.info(f"設定ファイル '{config_file}' を読み込みました")
        return config
        
    except json.JSONDecodeError as e:
        logging.error(f"設定ファイルのJSON形式が正しくありません: {str(e)}")
        return None
    except Exception as e:
        logging.error(f"設定ファイルの読み込み中にエラーが発生しました: {str(e)}")
        return None

# ログの設定
def setup_logging():
    """ログの設定を行う"""
    global captured_logs
    captured_logs = []  # ログメッセージをリセット
    
    # logディレクトリの作成（存在しない場合）
    log_dir = "log"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    # ログファイル名を現在の日時で作成
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_filename = os.path.join(log_dir, f"csv_log_{timestamp}.log")
    
    # カスタムハンドラーを作成
    class LogCaptureHandler(logging.Handler):
        def emit(self, record):
            log_message = self.format(record)
            captured_logs.append(log_message)
    
    # ログの設定
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filename, encoding='utf-8'),
            logging.StreamHandler(),  # コンソールにも出力
            LogCaptureHandler()  # ログメッセージをキャプチャ
        ]
    )
    return log_filename

def get_or_create_folder(service, parent_folder_id, folder_name):
    """指定された親フォルダ内にフォルダを作成または取得する"""
    try:
        # 既存のフォルダを検索
        query = f"'{parent_folder_id}' in parents and name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
        results = service.files().list(q=query).execute()
        files = results.get('files', [])
        
        if files:
            # 既存のフォルダが見つかった場合
            folder_id = files[0]['id']
            logging.info(f"既存のフォルダ '{folder_name}' を使用します (ID: {folder_id})")
            return folder_id
        else:
            # 新しいフォルダを作成
            folder_metadata = {
                'name': folder_name,
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [parent_folder_id]
            }
            folder = service.files().create(body=folder_metadata, fields='id').execute()
            folder_id = folder.get('id')
            logging.info(f"新しいフォルダ '{folder_name}' を作成しました (ID: {folder_id})")
            return folder_id
            
    except Exception as e:
        logging.error(f"フォルダの作成/取得中にエラーが発生しました: {str(e)}")
        return None

def upload_file_to_drive(service, file_path, folder_id, file_name=None):
    """ファイルをGoogle Driveにアップロードする"""
    try:
        if not os.path.exists(file_path):
            logging.error(f"アップロードするファイル '{file_path}' が見つかりません")
            return None
        
        # ファイル名が指定されていない場合は、元のファイル名を使用
        if file_name is None:
            file_name = os.path.basename(file_path)
        
        # ファイルのメディアタイプを決定
        file_extension = os.path.splitext(file_path)[1].lower()
        mime_type = 'text/plain'  # デフォルト
        if file_extension == '.csv':
            mime_type = 'text/csv'
        elif file_extension == '.log':
            mime_type = 'text/plain'
        
        # ファイルのメタデータ
        file_metadata = {
            'name': file_name,
            'parents': [folder_id]
        }
        
        # ファイルのアップロード
        media = MediaFileUpload(file_path, mimetype=mime_type, resumable=True)
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id,name'
        ).execute()
        
        file_id = file.get('id')
        file_name = file.get('name')
        logging.info(f"ファイル '{file_name}' をアップロードしました (ID: {file_id})")
        return file_id
        
    except Exception as e:
        logging.error(f"ファイルのアップロード中にエラーが発生しました: {str(e)}")
        return None

def upload_files_to_drive(creds, config):
    """ログファイルとCSVファイルをGoogle Driveにアップロードする"""
    drive_details = ""
    try:
        logging.info("Google Driveへのファイルアップロードを開始します")
        drive_details += "Google Driveへのファイルアップロードを開始\n"
        
        # Google Drive APIサービスを構築
        service = build('drive', 'v3', credentials=creds)
        
        # 設定から値を取得
        drive_folder_id = config.get('drive_folder_id')
        log_folder_name = config.get('log_folder_name', 'log')
        csv_folder_name = config.get('csv_folder_name', 'csv')
        csv_file_path = config.get('csv_file_path')
        
        drive_details += f"DriveフォルダID: {drive_folder_id}\n"
        drive_details += f"CSVフォルダ名: {csv_folder_name}\n"
        drive_details += f"CSVファイルパス: {csv_file_path}\n"
        
        # csvフォルダを取得または作成
        csv_folder_id = get_or_create_folder(service, drive_folder_id, csv_folder_name)
        if not csv_folder_id:
            error_msg = "csvフォルダの作成に失敗しました"
            logging.warning(error_msg)
            drive_details += error_msg + "\n"
            return False, drive_details
        
        drive_details += f"CSVフォルダID: {csv_folder_id}\n"
        
        # アップロード結果を追跡
        upload_success = True
        
        # CSVファイルをアップロード
        if csv_file_path and os.path.exists(csv_file_path):
            # タイムスタンプ付きのファイル名でアップロード
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            csv_file_name = f"test_{timestamp}.csv"
            drive_details += f"アップロードファイル名: {csv_file_name}\n"
            
            csv_upload_result = upload_file_to_drive(service, csv_file_path, csv_folder_id, csv_file_name)
            if not csv_upload_result:
                upload_success = False
                error_msg = "CSVファイルのアップロードに失敗しました"
                logging.warning(error_msg)
                drive_details += error_msg + "\n"
            else:
                drive_details += "CSVファイルのアップロードが成功\n"
        else:
            drive_details += "CSVファイルが存在しないため、アップロードをスキップ\n"
        
        if upload_success:
            logging.info("Google Driveへのファイルアップロードが完了しました")
            drive_details += "Google Driveへのファイルアップロードが完了\n"
        else:
            logging.warning("Google Driveへのファイルアップロードに一部または全部失敗しました")
            drive_details += "Google Driveへのファイルアップロードに一部または全部失敗\n"
        
        return upload_success, drive_details
        
    except Exception as e:
        error_msg = f"Google Driveへのファイルアップロード中にエラーが発生しました: {str(e)}"
        logging.warning(error_msg)
        import traceback
        logging.warning(f"詳細なエラー情報: {traceback.format_exc()}")
        drive_details += error_msg + "\n"
        drive_details += f"詳細なエラー情報: {traceback.format_exc()}\n"
        return False, drive_details

def read_csv_data(csv_filename):
    """CSVファイルを読み取ってデータを返す"""
    csv_details = ""
    try:
        # CSVファイルの存在確認
        if not os.path.exists(csv_filename):
            logging.error(f"CSVファイル '{csv_filename}' が見つかりません")
            return None, None, "CSVファイルが見つかりません"
        
        logging.info(f"CSVファイル '{csv_filename}' の読み取りを開始します")
        csv_details += f"ファイルパス: {csv_filename}\n"
        
        with open(csv_filename, 'r', encoding='utf-8') as file:
            csv_reader = csv.reader(file)
            
            # ヘッダー行を読み取り
            header = next(csv_reader)
            logging.info("=== CSVヘッダー ===")
            logging.info(f"ヘッダー: {header}")
            logging.info(f"カラム数: {len(header)}")
            
            csv_details += f"ヘッダー: {header[:5]}...\n"  # 最初の5列のみ表示
            csv_details += f"カラム数: {len(header)}\n"
            
            # データ行を読み取り
            logging.info("=== CSVデータ ===")
            data_rows = []
            row_count = 0
            for row_num, row in enumerate(csv_reader, start=2):  # 2行目から開始
                row_count += 1
                data_rows.append(row)
                logging.info(f"行 {row_num}: {row[:3]}...")  # 最初の3列のみ表示
                
                csv_details += f"行 {row_num}: {row[:3]}...\n"
            
            logging.info(f"=== 読み取り完了 ===")
            logging.info(f"総行数: {row_count + 1} (ヘッダー含む)")
            logging.info(f"データ行数: {row_count}")
            
            csv_details += f"総行数: {row_count + 1} (ヘッダー含む)\n"
            csv_details += f"データ行数: {row_count}\n"
            
            return header, data_rows, csv_details
            
    except UnicodeDecodeError:
        error_msg = "ファイルのエンコーディングエラーが発生しました。UTF-8でエンコードされているか確認してください。"
        logging.error(error_msg)
        return None, None, error_msg
    except Exception as e:
        error_msg = f"CSVファイルの読み取り中にエラーが発生しました: {str(e)}"
        logging.error(error_msg)
        return None, None, error_msg

def write_to_google_sheets_and_drive(header, data_rows, config):
    """Googleスプレッドシートにデータを書き込み、ファイルをDriveにアップロードする"""
    try:
        logging.info("Googleスプレッドシートへの書き込みを開始します")
        sheet_details = "Googleスプレッドシートへの書き込みを開始\n"
        
        # 設定から値を取得
        creds_file = config.get('creds_file', 'creds.json')
        spreadsheet_id = config.get('spreadsheet_id')
        sheet_name = config.get('sheet_name', 'sheet1')
        
        sheet_details += f"スプレッドシートID: {spreadsheet_id}\n"
        sheet_details += f"シート名: {sheet_name}\n"
        
        # 認証情報ファイルの確認
        if not os.path.exists(creds_file):
            error_msg = f"認証情報ファイル '{creds_file}' が見つかりません"
            logging.error(error_msg)
            return False, sheet_details, error_msg
        
        # 認証情報の読み込み
        creds = Credentials.from_service_account_file(creds_file, scopes=SCOPES)
        client = gspread.authorize(creds)
        
        # スプレッドシートを開く
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.worksheet(sheet_name)
        
        logging.info(f"スプレッドシート '{spreadsheet.title}' のシート '{sheet_name}' にアクセスしました")
        sheet_details += f"スプレッドシート名: {spreadsheet.title}\n"
        sheet_details += f"アクセスしたシート: {sheet_name}\n"
        
        # 既存のデータをクリア
        worksheet.clear()
        logging.info("既存のデータをクリアしました")
        sheet_details += "既存のデータをクリア\n"
        
        # ヘッダー行を書き込み
        worksheet.append_row(header)
        logging.info("ヘッダー行を書き込みました")
        sheet_details += f"ヘッダー行を書き込み: {header}\n"
        
        # データ行を書き込み
        for i, row in enumerate(data_rows, start=2):
            worksheet.append_row(row)
            logging.info(f"データ行 {i} を書き込みました: {row[:3]}...")  # 最初の3列のみ表示
            sheet_details += f"データ行 {i}: {row[:3]}...\n"
        
        logging.info(f"合計 {len(data_rows)} 行のデータを書き込みました")
        logging.info("Googleスプレッドシートへの書き込みが完了しました")
        sheet_details += f"合計 {len(data_rows)} 行のデータを書き込み完了\n"
        
        # Google Driveにファイルをアップロード
        drive_success, drive_details = upload_files_to_drive(creds, config)
        if drive_success:
            logging.info("Google Driveへのファイルアップロードが完了しました")
            drive_details += "Google Driveへのファイルアップロードが完了\n"
        else:
            logging.warning("Google Driveへのファイルアップロードに失敗しました")
            drive_details += "Google Driveへのファイルアップロードに失敗\n"
        
        return drive_success, sheet_details, drive_details
        
    except FileNotFoundError:
        error_msg = "認証情報ファイルが見つかりません"
        logging.error(error_msg)
        return False, sheet_details, error_msg
    except Exception as e:
        error_msg = f"Googleスプレッドシートへの書き込み中にエラーが発生しました: {str(e)}"
        logging.error(error_msg)
        import traceback
        logging.error(f"詳細なエラー情報: {traceback.format_exc()}")
        return False, sheet_details, error_msg

def log_to_spreadsheet(config, execution_id, status, message="", row_count="", heading_link=""):
    """実行ログをスプレッドシートに記録する"""
    try:
        # 設定から値を取得
        creds_file = config.get('creds_file', 'creds.json')
        spreadsheet_id = config.get('spreadsheet_id')
        log_sheet_name = config.get('log_sheet_name', '実行履歴')
        
        # 認証情報ファイルの確認
        if not os.path.exists(creds_file):
            logging.warning(f"認証情報ファイル '{creds_file}' が見つかりません")
            return False
        
        # 認証情報の読み込み
        creds = Credentials.from_service_account_file(creds_file, scopes=SCOPES)
        client = gspread.authorize(creds)
        
        # スプレッドシートを開く
        spreadsheet = client.open_by_key(spreadsheet_id)
        
        try:
            # ログシートを取得
            log_worksheet = spreadsheet.worksheet(log_sheet_name)
        except gspread.WorksheetNotFound:
            # ログシートが存在しない場合は作成
            log_worksheet = spreadsheet.add_worksheet(title=log_sheet_name, rows=1000, cols=10)
            # ヘッダー行を追加
            header = ["実行ID", "実行日時", "ステータス", "メッセージ", "CSVファイルパス", "処理行数", "Google Docsリンク"]
            log_worksheet.append_row(header)
            logging.info(f"ログシート '{log_sheet_name}' を作成しました")
        
        # 現在の日時を取得
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # ログデータを準備
        csv_file_path = config.get('csv_file_path', '')
        log_data = [execution_id, current_time, status, message, csv_file_path, row_count, heading_link]
        
        # 最終行に追加
        log_worksheet.append_row(log_data)
        logging.info(f"実行ログをスプレッドシートに記録しました: {status} (実行ID: {execution_id})")
        
        return True
        
    except Exception as e:
        logging.warning(f"スプレッドシートへのログ記録中にエラーが発生しました: {str(e)}")
        return False

def log_to_google_docs(config, execution_id, status, message="", row_count="", csv_details="", sheet_details="", drive_details=""):
    """実行ログをGoogle Docsに記録する"""
    try:
        # 設定から値を取得
        creds_file = config.get('creds_file', 'creds.json')
        log_doc_id = config.get('log_doc_id')
        
        if not log_doc_id:
            logging.warning("設定ファイルに 'log_doc_id' が設定されていません")
            return False
        
        # 認証情報ファイルの確認
        if not os.path.exists(creds_file):
            logging.warning(f"認証情報ファイル '{creds_file}' が見つかりません")
            return False
        
        # 認証情報の読み込み
        creds = Credentials.from_service_account_file(creds_file, scopes=SCOPES)
        
        # Google Docs APIサービスを構築
        docs_service = build('docs', 'v1', credentials=creds)
        
        # 現在の日時を取得
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # 実行IDをH1見出しとして挿入
        heading_text = f"実行ID: {execution_id}"
        
        # まず見出しテキストを挿入
        requests = [
            {
                'insertText': {
                    'location': {
                        'index': 1
                    },
                    'text': heading_text + '\n'
                }
            }
        ]
        
        docs_service.documents().batchUpdate(
            documentId=log_doc_id,
            body={'requests': requests}
        ).execute()
        
        # 見出しの位置を取得（挿入後のインデックス）
        heading_end_index = len(heading_text) + 1
        
        # 見出し部分だけをH1スタイルに設定
        requests = [
            {
                'updateParagraphStyle': {
                    'range': {
                        'startIndex': 1,
                        'endIndex': heading_end_index
                    },
                    'paragraphStyle': {
                        'namedStyleType': 'HEADING_1'
                    },
                    'fields': 'namedStyleType'
                }
            }
        ]
        
        docs_service.documents().batchUpdate(
            documentId=log_doc_id,
            body={'requests': requests}
        ).execute()
        
        # 詳細なログエントリを作成
        log_entry = f"\n実行日時: {current_time}\n"
        log_entry += f"ステータス: {status}\n"
        
        if message:
            log_entry += f"メッセージ: {message}\n"
        
        if row_count:
            log_entry += f"処理行数: {row_count}\n"
        
        # キャプチャされたログメッセージを追加
        if captured_logs:
            log_entry += f"実行ログ:\n"
            for log_msg in captured_logs:
                log_entry += f"{log_msg}\n"
        
        if csv_details:
            log_entry += f"CSV詳細:\n{csv_details}\n"
        
        if sheet_details:
            log_entry += f"スプレッドシート詳細:\n{sheet_details}\n"
        
        if drive_details:
            log_entry += f"Drive詳細:\n{drive_details}\n"
        
        log_entry += f"{'='*60}\n"
        log_entry += f"【実行ID: {execution_id} 終了】\n\n"
        
        # 見出しの後にログ内容を挿入（通常のテキストとして）
        requests = [
            {
                'insertText': {
                    'location': {
                        'index': heading_end_index
                    },
                    'text': log_entry
                }
            }
        ]
        
        docs_service.documents().batchUpdate(
            documentId=log_doc_id,
            body={'requests': requests}
        ).execute()
        
        # ログ内容部分だけをNORMAL_TEXTスタイルに設定（実行ID行は除外）
        log_start_index = heading_end_index + 1  # 改行の後から開始
        log_end_index = heading_end_index + len(log_entry)
        
        requests = [
            {
                'updateParagraphStyle': {
                    'range': {
                        'startIndex': log_start_index,
                        'endIndex': log_end_index
                    },
                    'paragraphStyle': {
                        'namedStyleType': 'NORMAL_TEXT'
                    },
                    'fields': 'namedStyleType'
                }
            }
        ]
        
        docs_service.documents().batchUpdate(
            documentId=log_doc_id,
            body={'requests': requests}
        ).execute()
        
        # 見出しへのリンクを生成
        doc_url = f"https://docs.google.com/document/d/{log_doc_id}/edit"
        heading_link = f"{doc_url}#heading=h.{execution_id}"
        
        logging.info(f"実行ログをGoogle Docsに記録しました: {status} (実行ID: {execution_id})")
        logging.info(f"見出しリンク: {heading_link}")
        return True, heading_link
        
    except Exception as e:
        logging.error(f"Google Docsへのログ記録中にエラーが発生しました: {str(e)}")
        return False, ""

def main():
    """メイン処理"""
    # 実行IDを生成
    execution_id = str(uuid.uuid4())[:8]  # 8文字の短縮UUID
    
    # ログの設定
    log_filename = setup_logging()
    logging.info(f"CSVファイル読み取り・Googleスプレッドシート書き込みプログラムを開始します (実行ID: {execution_id})")
    
    # 設定ファイルを読み込み
    config = load_config()
    if not config:
        logging.error("設定ファイルの読み込みに失敗しました。プログラムを終了します。")
        # エラーログをGoogle Docsに記録
        try:
            log_to_google_docs(config, execution_id, "エラー", "設定ファイルの読み込みに失敗")
        except:
            pass
        return
    
    # 設定からCSVファイルパスを取得
    csv_filename = config.get('csv_file_path')
    if not csv_filename:
        logging.error("設定ファイルに 'csv_file_path' が設定されていません")
        # エラーログをGoogle Docsに記録
        log_to_google_docs(config, execution_id, "エラー", "CSVファイルパスが設定されていません")
        return
    
    # CSVファイルを読み取り
    header, data_rows, csv_details = read_csv_data(csv_filename)
    
    if header and data_rows:
        # Googleスプレッドシートに書き込み、Driveにファイルをアップロード
        success, sheet_details, drive_details = write_to_google_sheets_and_drive(header, data_rows, config)
        
        if success:
            logging.info("すべての処理が正常に完了しました")
            # 成功ログをGoogle Docsに記録
            docs_success, heading_link = log_to_google_docs(config, execution_id, "成功", f"CSVデータを正常に処理しました（{len(data_rows)}行）", str(len(data_rows)), csv_details, sheet_details, drive_details)
            # スプレッドシートにもログを記録
            log_to_spreadsheet(config, execution_id, "成功", f"CSVデータを正常に処理しました（{len(data_rows)}行）", str(len(data_rows)), heading_link)
        else:
            logging.error("Googleスプレッドシートへの書き込みまたはDriveへのアップロードに失敗しました")
            # エラーログをGoogle Docsに記録
            docs_success, heading_link = log_to_google_docs(config, execution_id, "エラー", "スプレッドシートへの書き込みまたはDriveへのアップロードに失敗", "", csv_details, sheet_details, drive_details)
            # スプレッドシートにもログを記録
            log_to_spreadsheet(config, execution_id, "エラー", "スプレッドシートへの書き込みまたはDriveへのアップロードに失敗", "", heading_link)
    else:
        logging.error("CSVファイルの読み取りに失敗しました")
        # エラーログをGoogle Docsに記録
        docs_success, heading_link = log_to_google_docs(config, execution_id, "エラー", "CSVファイルの読み取りに失敗", "", csv_details, "", "")
        # スプレッドシートにもログを記録
        log_to_spreadsheet(config, execution_id, "エラー", "CSVファイルの読み取りに失敗", "", heading_link)
    
    logging.info(f"処理が完了しました。ログファイル: {log_filename} (実行ID: {execution_id})")

if __name__ == "__main__":
    main()