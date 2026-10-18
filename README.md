# InternTrack — 28卒ポートフォリオ（公開デモ版）

就職活動の企業情報・選考履歴・予定を一元管理する Flask 製 Web アプリです。実際の個人データやメールを公開しないため、**採用担当者向けデモを本番運用版とは分離**しています。

## デモ版で見られること

- ホーム画面・応募予定・進捗
- 企業の一覧・詳細・追加・編集・削除
- 選考履歴やカレンダーの表示・編集
- インターン・本選考の状況管理

閲覧者ごとに仮のデータベースを生成します。表示される企業・日程は架空のサンプルです。保存した変更は**一時的**であり、一定時間経過・サーバー再起動・再デプロイ時などに消える場合があります。

メール連携・OAuth・アカウント登録・パスワード変更は公開デモでは無効にしています。**本番のメール連携の実装はコード上にありますが、デモから実行できません。**

## Renderで公開する手順

1. このフォルダの中身 **だけ** を、新しいGitHubリポジトリ（例: `interntrack-portfolio`）にアップロードします。以前の `.git` 履歴や元の `interntrack.7z` はアップロードしないでください。
2. GitHubで公開されたファイルの中に `.env`、`.db`、`google_oauth_client.json`、実際のメール情報・OAuthトークンが**ないことを必ず目視で確認**します。
3. [Render](https://dashboard.render.com/) でログインし、**New → Blueprint** を選びます。
4. 作成したGitHubリポジトリを接続し、リポジトリ内の `render.yaml` を使ってデプロイします。
5. Renderが `INTERNTRACK_SECRET_KEY` をランダム生成し、`INTERNTRACK_DEMO_MODE=1` に設定していることを確認します。
6. デプロイ完了後、Renderの `https://...onrender.com/` にアクセスします。ホーム・企業一覧・カレンダー・追加・編集が表示されるか確認します。
7. サイボウズの成果物欄に、**GitHubのリポジトリURL**と**Renderで発行されたデモURL**を記載します。

Render無料プランでは起動まで時間がかかる場合があり、ファイルシステムは永続化されません。これはポートフォリオ用デモとしての想定動作です。実運用・一般公開の利用者データ保存用途には使わないでください。

## Windowsでローカルテストする場合

Python 3.12を用意し、PowerShellでこのフォルダに移動します。

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
$env:INTERNTRACK_DEMO_MODE = "1"
$env:INTERNTRACK_DEMO_LOCAL = "1"
$env:INTERNTRACK_SECRET_KEY = python -c "import secrets; print(secrets.token_hex(32))"
python app.py
```

`http://127.0.0.1:5000/` をブラウザで開きます。

## 使用技術

Python / Flask / SQLite / HTML / CSS / JavaScript / Google Gmail API / Microsoft Graph API（デモ版では外部連携を無効化）

## セキュリティ・運用上の注意

- このコードは**ポートフォリオデモ専用**です。一般ユーザーが実データを預けるための運用版ではありません。
- `.gitignore` は機密ファイルの混入防止を補助するだけです。アップロード前に各ファイルを確認してください。
- GitHubの元プロジェクトにシークレットが過去コミットされている場合、ファイルを消しても履歴に残ります。**新規リポジトリに本デモ版だけを登録**してください。必要なら関連キーをローテーションしてください。
- ブラウザ上で実在する企業・氏名・メールアドレス等の個人情報を入力しないでください。
- サーバー上のデータベースは各訪問者で分離していますが、無料デモとしての負荷対策や業務用の監査機能は実装していません。
- セッション識別子・COOKIE設定・デモ用秘密鍵は本番時に必ずRenderの環境変数から生成・設定します。
- 今後、実際の他ユーザーへの一般公開には、本番用DB（PostgreSQL等）、CSRF対策、厳密な入力検証、認証基盤、バックアップ、監視等の見直しが必要です。

## ファイル構成

- `app.py`: 既存Flaskアプリ + 独立したポートフォリオデモモード
- `templates/`, `static/`: オリジナルUIを維持
- `render.yaml`: Renderのデプロイ設定（GitHubで接続）
- `requirements.txt`: Python依存ライブラリ
- `.gitignore`: 秘密情報・データベースの除外ルール
- `PORTFOLIO_TEXT.md`: 企業提出用の紹介文
