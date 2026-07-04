# GitHub Copilot Chat（VS Code）の使い方

## 前提
- 2025年12月の統合により、旧「GitHub Copilot」（インライン補完）と「GitHub Copilot Chat」は**「GitHub Copilot Chat」1本**に統合済み
- VS Code拡張機能タブで「GitHub Copilot Chat」（発行元：GitHub）をインストールし、GitHubアカウントでサインインすれば利用可能
- 無料プラン：月50回のChat利用、2,000回のコード補完まで無料

## チャットパネルを開く
- サイドバーの吹き出しアイコンをクリック
- または `Ctrl+Alt+I`（Windows）

## 基本の質問
- チャット欄に日本語でそのまま質問できる（例：「このコードの問題点を指摘して」）
- コードを選択した状態で質問すると、その範囲を踏まえて回答してくれる

## 便利なスラッシュコマンド
| コマンド | 内容 |
|---|---|
| `/fix` | 選択したコードのエラーを説明しながら修正案を提示 |
| `/tests` | 選択した関数のユニットテストのひな形を自動生成 |
| `/explain` | 選択したコードの解説 |

## インラインチャット（エディタ上で直接編集）
- コード上で `Ctrl+I` を押すと、その場に小さな入力欄が出る
- 指示した内容をカーソル位置に直接反映してくれる

## エージェントモード
- チャットパネル内でモード切り替え（Ask / Edit / Agent 等）が可能
- 「Agent」を選ぶと、複数ファイルの一括編集・ターミナルコマンド実行・テスト実行・PR作成まで自律的にこなす
- Claude Codeのエージェント機能とかなり近い動き

## 参考リンク
- [Asking GitHub Copilot questions in your IDE - GitHub Docs](https://docs.github.com/en/copilot/how-tos/chat-with-copilot/chat-in-ide)
- [GitHub Copilot Chat - Visual Studio Marketplace](https://marketplace.visualstudio.com/items?itemName=GitHub.copilot-chat)
