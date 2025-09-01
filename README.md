# CSV to Google Sheets with Logging and Drive Integration

CSVファイルを読み取り、Googleスプレッドシートに書き込み、ログファイルを生成するPythonアプリケーションです。

## 機能

- 📊 CSVファイルの読み取りと解析
- 📝 詳細なログファイルの生成（ローカルとGoogle Drive）
- 📋 Googleスプレッドシートへのデータ書き込み
- 🔐 Google API認証（サービスアカウント）
- 📁 Google Driveへのファイルバックアップ

## セットアップ

### 1. リポジトリのクローン

```bash
git clone https://github.com/otokichi3/pepal-app.git
cd pepal-app
```

### 2. 依存関係のインストール

```bash
pip install -r requirements.txt
```

### 3. Google Cloud Consoleでの設定

#### 3.1 プロジェクトの作成
1. [Google Cloud Console](https://console.cloud.google.com/)にアクセス
2. 新しいプロジェクトを作成または既存のプロジェクトを選択

#### 3.2 APIの有効化
- Google Sheets API
- Google Drive API

#### 3.3 サービスアカウントの作成
1. 「APIとサービス」→「認証情報」
2. 「認証情報を作成」→「サービスアカウント」
3. サービスアカウント名を入力（例：`csv-to-sheets`）
4. 「キーを作成」→「JSON」でダウンロード

#### 3.4 認証情報ファイルの配置
ダウンロードしたJSONファイルを`creds.json`としてプロジェクトルートに配置

### 4. スプレッドシートの共有設定

1. 対象のGoogleスプレッドシートを開く
2. 「共有」ボタンをクリック
3. サービスアカウントのメールアドレス（`creds.json`内の`client_email`）を追加
4. 「編集者」権限を付与

## 使用方法

### 基本的な使用方法

```bash
python main.py
```

### 設定のカスタマイズ

`main.py`内の以下の定数を変更して、対象のスプレッドシートとDriveフォルダを指定：

```python
# スプレッドシートのID（URLから抽出）
SPREADSHEET_ID = 'your-spreadsheet-id'
SHEET_NAME = 'sheet1'

# Google DriveフォルダのID
DRIVE_FOLDER_ID = 'your-drive-folder-id'
```

## ファイル構成

```
pepal-app/
├── main.py              # メインスクリプト
├── requirements.txt     # 依存関係
├── test.csv            # サンプルCSVファイル
├── creds.json          # 認証情報（.gitignoreで除外）
├── .gitignore          # Git除外設定
├── README.md           # このファイル
└── log/                # ログファイル（.gitignoreで除外）
    └── csv_log_*.log   # 実行ログ
```

## CSVファイル形式

以下のカラムを含むCSVファイルに対応：

- 必要項目
- 受付No
- 伝票No
- 枝番
- 納入年月日
- 得意先名
- 得意先住所
- 配送地域CD
- 漢字届け先名
- 漢字届け先住所
- 請求備考
- 入出庫備考
- 配送備考
- 断切備考
- 時間指定
- 請求商品CD
- 請求商品名
- 売上数量
- 売上重量
- 売上赤黒抹消ＳＢＬ
- 請求寸法
- 請求KG連量（or請求表示連量）
- 断裁方法
- 請求数量(単位ごとに再計算)
- 請求単位
- 請求包数
- オペレーターNO
- 得意先CD
- 届け先CD（or得意先CD）

## ログ機能

### ローカルログ
- `log/`ディレクトリにタイムスタンプ付きで保存
- CSVファイルの読み取り詳細
- Googleスプレッドシートへの書き込み状況
- エラー情報とデバッグ情報

### Google Driveログ
- 指定されたDriveフォルダの`log/`サブフォルダに保存
- CSVファイルのバックアップも`csv/`サブフォルダに保存

## トラブルシューティング

### よくある問題

#### 1. 認証エラー
```
PermissionError: The caller does not have permission
```
**解決方法**: スプレッドシートにサービスアカウントのメールアドレスを共有設定で追加

#### 2. ストレージクォータエラー
```
storageQuotaExceeded
```
**解決方法**: サービスアカウントにはストレージクォータがないため、共有ドライブを使用するか、OAuth委任を使用

#### 3. ファイルが見つからない
```
CSVファイル 'test.csv' が見つかりません
```
**解決方法**: CSVファイルが正しい場所に配置されているか確認

## 開発者向け情報

### 依存関係

- `gspread>=6.2.0` - Google Sheets API
- `google-api-python-client>=2.179.0` - Google Drive API
- `google-auth>=2.40.0` - Google認証

### 環境変数

必要に応じて以下の環境変数を設定：

```bash
export GOOGLE_APPLICATION_CREDENTIALS="path/to/creds.json"
```

## ライセンス

このプロジェクトはMITライセンスの下で公開されています。

## 貢献

プルリクエストやイシューの報告を歓迎します。

## 更新履歴

- **v1.0.0** - 初期リリース
  - CSVファイル読み取り機能
  - Googleスプレッドシート書き込み機能
  - ログ機能
  - Google Drive統合
