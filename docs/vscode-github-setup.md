# GitHubリポジトリをVSCodeで編集するための準備と作業フロー

## 1. 事前準備（初回のみ）

### 1.1 必須ツールのインストール
| ツール | 用途 | 確認コマンド |
|---|---|---|
| Git | バージョン管理本体 | `git --version` |
| VSCode | エディタ | - |
| GitHubアカウント | リポジトリのホスティング | - |

### 1.2 Gitの初期設定（マシンごとに1回）
```bash
git config --global user.name "あなたの名前"
git config --global user.email "GitHubに登録したメールアドレス"
```

### 1.3 認証方法の設定（いずれか一方）
- **HTTPS + Personal Access Token（PAT）**：push時にユーザー名とPATの入力を求められる。GitHub側で「Settings → Developer settings → Personal access tokens」から発行。
- **SSH鍵**：`ssh-keygen`で鍵を生成し、公開鍵をGitHubの「Settings → SSH and GPG keys」に登録。以降は認証入力不要になり、こちらが一般的に推奨される。

### 1.4 VSCode拡張機能（推奨）
- **GitHub Pull Requests and Issues**：VSCode上でPRやIssueを直接操作できる
- **GitLens**：コミット履歴・変更箇所の可視化
- （Git連携自体はVSCode標準機能で完結するため必須ではない）

## 2. リポジトリを手元に持ってくる（クローン）

### 方法A：VSCodeから行う
1. `Ctrl+Shift+P` → 「Git: Clone」を選択
2. リポジトリのURLを入力（GitHubの「Code」ボタンからコピー）
3. 保存先フォルダを選択 → クローン完了後「Open」を選ぶ

### 方法B：ターミナルから行う
```bash
git clone https://github.com/ユーザー名/リポジトリ名.git
cd リポジトリ名
code .
```

これでローカルにリポジトリが複製され、VSCodeで開かれた状態になります。

## 3. 日常の作業フロー

```
① git pull（最新化）
     ↓
② 作業用ブランチを作成
     ↓
③ コードを編集
     ↓
④ git add / git commit（変更を記録）
     ↓
⑤ git push（リモートへ反映）
     ↓
⑥ Pull Request（PR）を作成
     ↓
⑦ レビュー・マージ
     ↓
①に戻る（pullして最新化）
```

### 各ステップの実操作

**① 最新化**
```bash
git pull origin main
```

**② ブランチ作成**
```bash
git checkout -b feature/作業内容がわかる名前
```
（VSCodeなら左下のブランチ名表示をクリック → 「Create new branch」でも可）

**③ 編集**
VSCode上で通常どおりファイルを編集。変更したファイルはソース管理タブ（`Ctrl+Shift+G`）に一覧表示される。

**④ ステージ＋コミット**
```bash
git add .
git commit -m "変更内容の要約"
```
VSCodeなら、ソース管理タブで変更ファイル横の「+」をクリックしてステージし、メッセージ入力欄に書いて「✓ Commit」でも同じことができる。

**⑤ プッシュ**
```bash
git push -u origin feature/作業内容がわかる名前
```
（初回のみ`-u`でリモートブランチと紐付け。2回目以降は`git push`だけでよい）

**⑥ PR作成**
GitHub上、またはVSCode拡張「GitHub Pull Requests and Issues」からPRを作成し、レビュー依頼を出す。

**⑦ マージ後**
マージが完了したら、ローカルの`main`に戻って最新化する：
```bash
git checkout main
git pull origin main
```

## 4. つまずきやすいポイント

- **pull忘れ**：他人の変更が入っているのに古い状態で作業すると、pushやマージ時にコンフリクトが起きやすい。作業開始前は必ずpull。
- **mainブランチで直接編集しない**：必ず作業用ブランチを切ってから編集する。
- **認証エラー**：HTTPS方式でPATの有効期限が切れるとpushが失敗する。SSH方式なら再発行の手間がない。
