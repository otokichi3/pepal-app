import csv
import logging
from datetime import datetime
import os
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Google Sheets APIの設定
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# スプレッドシートのID（URLから抽出）
SPREADSHEET_ID = '1KDEaN6_9r4UslrBKm42mso02sHXa_6c46E1l5-d5hZM'
SHEET_NAME = 'sheet1'

# Google DriveフォルダのID
DRIVE_FOLDER_ID = '1p03-tKNaZnLCANbXKFe1hJPmr4C1AGp8'
LOG_FOLDER_NAME = 'log'
CSV_FOLDER_NAME = 'csv'

# ログの設定
def setup_logging():
    """ログの設定を行う"""
    # logディレクトリの作成（存在しない場合）
    log_dir = "log"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    # ログファイル名を現在の日時で作成
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_filename = os.path.join(log_dir, f"csv_log_{timestamp}.log")
    
    # ログの設定
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filename, encoding='utf-8'),
            logging.StreamHandler()  # コンソールにも出力
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

def upload_files_to_drive(creds):
    """ログファイルとCSVファイルをGoogle Driveにアップロードする"""
    try:
        logging.info("Google Driveへのファイルアップロードを開始します")
        
        # Google Drive APIサービスを構築
        service = build('drive', 'v3', credentials=creds)
        
        # logフォルダを取得または作成
        log_folder_id = get_or_create_folder(service, DRIVE_FOLDER_ID, LOG_FOLDER_NAME)
        if not log_folder_id:
            logging.error("logフォルダの作成に失敗しました")
            return False
        
        # csvフォルダを取得または作成
        csv_folder_id = get_or_create_folder(service, DRIVE_FOLDER_ID, CSV_FOLDER_NAME)
        if not csv_folder_id:
            logging.error("csvフォルダの作成に失敗しました")
            return False
        
        # 最新のログファイルをアップロード
        log_dir = "log"
        if os.path.exists(log_dir):
            log_files = [f for f in os.listdir(log_dir) if f.endswith('.log')]
            if log_files:
                # 最新のログファイルを取得
                latest_log = max(log_files, key=lambda x: os.path.getctime(os.path.join(log_dir, x)))
                log_file_path = os.path.join(log_dir, latest_log)
                upload_file_to_drive(service, log_file_path, log_folder_id)
        
        # CSVファイルをアップロード
        csv_file_path = "test.csv"
        if os.path.exists(csv_file_path):
            # タイムスタンプ付きのファイル名でアップロード
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            csv_file_name = f"test_{timestamp}.csv"
            upload_file_to_drive(service, csv_file_path, csv_folder_id, csv_file_name)
        
        logging.info("Google Driveへのファイルアップロードが完了しました")
        return True
        
    except Exception as e:
        logging.error(f"Google Driveへのファイルアップロード中にエラーが発生しました: {str(e)}")
        import traceback
        logging.error(f"詳細なエラー情報: {traceback.format_exc()}")
        return False

def read_csv_data(csv_filename):
    """CSVファイルを読み取ってデータを返す"""
    try:
        # CSVファイルの存在確認
        if not os.path.exists(csv_filename):
            logging.error(f"CSVファイル '{csv_filename}' が見つかりません")
            return None, None
        
        logging.info(f"CSVファイル '{csv_filename}' の読み取りを開始します")
        
        with open(csv_filename, 'r', encoding='utf-8') as file:
            csv_reader = csv.reader(file)
            
            # ヘッダー行を読み取り
            header = next(csv_reader)
            logging.info("=== CSVヘッダー ===")
            logging.info(f"ヘッダー: {header}")
            logging.info(f"カラム数: {len(header)}")
            
            # データ行を読み取り
            logging.info("=== CSVデータ ===")
            data_rows = []
            row_count = 0
            for row_num, row in enumerate(csv_reader, start=2):  # 2行目から開始
                row_count += 1
                data_rows.append(row)
                logging.info(f"行 {row_num}: {row}")
                
                # 各カラムの内容も詳細にログ出力
                for col_num, (header_name, value) in enumerate(zip(header, row)):
                    if value.strip():  # 空でない値のみ出力
                        logging.info(f"  カラム {col_num+1} ({header_name}): {value}")
            
            logging.info(f"=== 読み取り完了 ===")
            logging.info(f"総行数: {row_count + 1} (ヘッダー含む)")
            logging.info(f"データ行数: {row_count}")
            
            return header, data_rows
            
    except UnicodeDecodeError:
        logging.error("ファイルのエンコーディングエラーが発生しました。UTF-8でエンコードされているか確認してください。")
        return None, None
    except Exception as e:
        logging.error(f"CSVファイルの読み取り中にエラーが発生しました: {str(e)}")
        return None, None

def write_to_google_sheets_and_drive(header, data_rows):
    """Googleスプレッドシートにデータを書き込み、ファイルをDriveにアップロードする"""
    try:
        logging.info("Googleスプレッドシートへの書き込みを開始します")
        
        # 認証情報ファイルの確認
        creds_file = "creds.json"
        if not os.path.exists(creds_file):
            logging.error(f"認証情報ファイル '{creds_file}' が見つかりません")
            logging.error("Google Cloud Consoleからサービスアカウントキーをダウンロードしてください")
            logging.error("1. Google Cloud Consoleにアクセス")
            logging.error("2. プロジェクトを作成または選択")
            logging.error("3. APIとサービス > 認証情報")
            logging.error("4. サービスアカウントを作成")
            logging.error("5. キーを作成（JSON形式）")
            logging.error("6. ダウンロードしたJSONファイルを 'creds.json' として保存")
            logging.error("7. スプレッドシートにサービスアカウントのメールアドレスを共有設定で追加")
            return False
        
        # 認証情報の読み込み
        creds = Credentials.from_service_account_file(creds_file, scopes=SCOPES)
        client = gspread.authorize(creds)
        
        # スプレッドシートを開く
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        worksheet = spreadsheet.worksheet(SHEET_NAME)
        
        logging.info(f"スプレッドシート '{spreadsheet.title}' のシート '{SHEET_NAME}' にアクセスしました")
        
        # 既存のデータをクリア
        worksheet.clear()
        logging.info("既存のデータをクリアしました")
        
        # ヘッダー行を書き込み
        worksheet.append_row(header)
        logging.info("ヘッダー行を書き込みました")
        
        # データ行を書き込み
        for i, row in enumerate(data_rows, start=2):
            worksheet.append_row(row)
            logging.info(f"データ行 {i} を書き込みました: {row[:3]}...")  # 最初の3列のみ表示
        
        logging.info(f"合計 {len(data_rows)} 行のデータを書き込みました")
        logging.info("Googleスプレッドシートへの書き込みが完了しました")
        
        # Google Driveにファイルをアップロード
        drive_success = upload_files_to_drive(creds)
        if drive_success:
            logging.info("Google Driveへのファイルアップロードが完了しました")
        else:
            logging.error("Google Driveへのファイルアップロードに失敗しました")
        
        return True
        
    except FileNotFoundError:
        logging.error("認証情報ファイルが見つかりません")
        return False
    except Exception as e:
        logging.error(f"Googleスプレッドシートへの書き込み中にエラーが発生しました: {str(e)}")
        import traceback
        logging.error(f"詳細なエラー情報: {traceback.format_exc()}")
        return False

def main():
    """メイン処理"""
    # ログの設定
    log_filename = setup_logging()
    logging.info("CSVファイル読み取り・Googleスプレッドシート書き込みプログラムを開始します")
    
    # CSVファイル名
    csv_filename = "test.csv"
    
    # CSVファイルを読み取り
    header, data_rows = read_csv_data(csv_filename)
    
    if header and data_rows:
        # Googleスプレッドシートに書き込み、Driveにファイルをアップロード
        success = write_to_google_sheets_and_drive(header, data_rows)
        
        if success:
            logging.info("すべての処理が正常に完了しました")
        else:
            logging.error("Googleスプレッドシートへの書き込みまたはDriveへのアップロードに失敗しました")
    else:
        logging.error("CSVファイルの読み取りに失敗しました")
    
    logging.info(f"処理が完了しました。ログファイル: {log_filename}")

if __name__ == "__main__":
    main()