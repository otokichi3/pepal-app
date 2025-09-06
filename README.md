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

`config.json`ファイルで設定を変更：

```json
{
    "spreadsheet_id": "your-spreadsheet-id",
    "sheet_name": "sheet1",
    "log_sheet_name": "実行履歴",
    "log_doc_id": "your-google-docs-id",
    "drive_folder_id": "your-drive-folder-id",
    "csv_file_path": "Z:\\test.csv",
    "creds_file": "creds.json",
    "log_folder_name": "log",
    "csv_folder_name": "csv"
}
```

## ファイル構成

```
pepal-app/
├── main.py              # メインスクリプト
├── config.json          # 設定ファイル
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

## リファクタリング提案

現在のコードの読みやすさを向上させるための包括的なリファクタリング案：

### 1. アーキテクチャの改善

#### 1.1 クラスベース設計への移行
- **現状**: 関数ベースの設計で、グローバル変数（`captured_logs`）を使用
- **改善案**: 
  - `CSVProcessor`クラス: CSV読み取り処理
  - `GoogleSheetsManager`クラス: スプレッドシート操作
  - `GoogleDriveManager`クラス: Drive操作
  - `Logger`クラス: ログ管理
  - `ConfigManager`クラス: 設定管理

#### 1.2 依存性注入（DI）の導入
- **現状**: 各関数内で認証情報を直接読み込み
- **改善案**: 認証情報を外部から注入し、テストしやすくする

#### 1.3 エラーハンドリングの統一
- **現状**: 各関数で異なるエラーハンドリング方式
- **改善案**: カスタム例外クラスと統一されたエラーハンドリング

### 2. 関数の分割と単一責任原則

#### 2.1 長い関数の分割
- **`main()`関数** (66行): 複数の責任を持つ
  - 設定読み込み
  - CSV処理
  - ログ記録
  - エラーハンドリング
- **改善案**: 各責任を別々の関数に分割

#### 2.2 `log_to_google_docs()`関数の分割
- **現状**: 147行の巨大な関数
- **改善案**:
  - `_insert_heading()`: 見出し挿入
  - `_apply_heading_style()`: 見出しスタイル適用
  - `_insert_log_content()`: ログ内容挿入
  - `_apply_content_style()`: 内容スタイル適用
  - `_generate_heading_link()`: リンク生成

#### 2.3 `upload_files_to_drive()`関数の分割
- **現状**: 67行で複数の責任
- **改善案**:
  - `_setup_drive_service()`: サービス初期化
  - `_create_or_get_folder()`: フォルダ操作
  - `_upload_csv_file()`: CSVファイルアップロード
  - `_upload_log_file()`: ログファイルアップロード

### 3. 設定管理の改善

#### 3.1 設定バリデーション
- **現状**: 設定値の存在チェックのみ
- **改善案**: 
  - 必須項目の検証
  - 値の形式チェック（URL、ID形式など）
  - デフォルト値の設定

#### 3.2 環境別設定
- **現状**: 単一の`config.json`
- **改善案**: 
  - `config.dev.json`
  - `config.prod.json`
  - 環境変数での上書き

### 4. ログ管理の改善

#### 4.1 ログレベルの体系化
- **現状**: ログレベルが一貫していない
- **改善案**:
  - `DEBUG`: 詳細なデバッグ情報
  - `INFO`: 一般的な処理情報
  - `WARNING`: 警告（Driveアップロード失敗など）
  - `ERROR`: エラー（処理失敗）
  - `CRITICAL`: 致命的エラー

#### 4.2 構造化ログ
- **現状**: プレーンテキストログ
- **改善案**: JSON形式の構造化ログ

#### 4.3 ログローテーション
- **現状**: ファイルサイズ制限なし
- **改善案**: サイズベースまたは日付ベースのローテーション

### 5. データモデルの導入

#### 5.1 データクラスの使用
```python
from dataclasses import dataclass
from typing import List, Optional

@dataclass
class CSVData:
    header: List[str]
    rows: List[List[str]]
    file_path: str
    row_count: int

@dataclass
class ExecutionResult:
    execution_id: str
    status: str
    message: str
    csv_details: str
    sheet_details: str
    drive_details: str
    warning: Optional[str] = None
```

#### 5.2 型ヒントの追加
- **現状**: 型ヒントが不十分
- **改善案**: 全ての関数に型ヒントを追加

### 6. テスト可能性の向上

#### 6.1 モック化の容易さ
- **現状**: Google API呼び出しが直接組み込まれている
- **改善案**: インターフェースを定義し、モック実装を提供

#### 6.2 単体テストの追加
- **現状**: テストコードなし
- **改善案**: 
  - `test_csv_processor.py`
  - `test_google_sheets_manager.py`
  - `test_logger.py`

### 7. パフォーマンスの改善

#### 7.1 非同期処理
- **現状**: 同期処理のみ
- **改善案**: Google API呼び出しを非同期化

#### 7.2 バッチ処理
- **現状**: 行ごとのスプレッドシート書き込み
- **改善案**: バッチでの一括書き込み

#### 7.3 キャッシュ機能
- **現状**: 認証情報を毎回読み込み
- **改善案**: 認証情報のキャッシュ

### 8. セキュリティの向上

#### 8.1 機密情報の管理
- **現状**: 認証情報がファイルに保存
- **改善案**: 
  - 環境変数での管理
  - 暗号化された設定ファイル
  - シークレット管理サービスとの統合

#### 8.2 入力検証
- **現状**: CSVファイルの内容検証なし
- **改善案**: 
  - ファイルサイズ制限
  - 列数制限
  - データ型検証

### 9. ユーザビリティの向上

#### 9.1 コマンドライン引数
- **現状**: 設定ファイルのみ
- **改善案**: 
  - `--config` オプション
  - `--csv-file` オプション
  - `--dry-run` オプション

#### 9.2 プログレスバー
- **現状**: 進捗表示なし
- **改善案**: 大容量ファイル処理時の進捗表示

#### 9.3 設定ウィザード
- **現状**: 手動設定
- **改善案**: 初回実行時の設定ウィザード

### 10. ドキュメントの改善

#### 10.1 APIドキュメント
- **現状**: 関数のdocstringが不十分
- **改善案**: 
  - 詳細なdocstring
  - 型情報の記載
  - 使用例の追加

#### 10.2 アーキテクチャドキュメント
- **現状**: アーキテクチャの説明なし
- **改善案**: 
  - システム構成図
  - データフロー図
  - クラス図

### 11. エラーハンドリングの改善

#### 11.1 カスタム例外
```python
class CSVProcessingError(Exception):
    """CSV処理エラー"""
    pass

class GoogleAPIError(Exception):
    """Google APIエラー"""
    pass

class ConfigurationError(Exception):
    """設定エラー"""
    pass
```

#### 11.2 リトライ機能
- **現状**: エラー時のリトライなし
- **改善案**: 指数バックオフでのリトライ

#### 11.3 エラー通知
- **現状**: ログのみ
- **改善案**: 
  - メール通知
  - Slack通知
  - 監視システムとの統合

### 12. 設定の柔軟性

#### 12.1 プラグインシステム
- **現状**: 固定の処理フロー
- **改善案**: カスタム処理のプラグイン化

#### 12.2 テンプレート機能
- **現状**: 固定のスプレッドシート形式
- **改善案**: テンプレートベースの出力

### 13. 監視とメトリクス

#### 13.1 メトリクス収集
- **現状**: メトリクスなし
- **改善案**: 
  - 処理時間の計測
  - エラー率の追跡
  - データ量の監視

#### 13.2 ヘルスチェック
- **現状**: ヘルスチェック機能なし
- **改善案**: API接続確認エンドポイント

### 14. 国際化対応

#### 14.1 多言語対応
- **現状**: 日本語のみ
- **改善案**: 
  - 英語対応
  - 設定ファイルでの言語選択

#### 14.2 タイムゾーン対応
- **現状**: ローカル時間のみ
- **改善案**: 設定可能なタイムゾーン

### 15. デプロイメントの改善

#### 15.1 Docker化
- **現状**: ローカル実行のみ
- **改善案**: 
  - Dockerfile
  - docker-compose.yml
  - 本番環境での実行

#### 15.2 CI/CD
- **現状**: 自動化なし
- **改善案**: 
  - GitHub Actions
  - 自動テスト
  - 自動デプロイ

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