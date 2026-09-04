---
date: 2026-09-03
type: doc
project: claude_playground
title: GitHub DesktopでもPRを必須化する方針
status: done
decision: サーバー側のbranch protection rule（Require a pull request before merging）を主策とし、Desktop上の運用ルール変更は補助とする
tags: [github, git, workflow, branch-protection]
---

## 依頼

「今GitHubデスクトップで作業していて、プルリクエストとかしなくてもどんどんマージできちゃってる
んだけど、これってmainを直接編集してるからだよね？あるべき形にするにはどうすればいい？
作業場所はclaude_playground配下で、主力はマークダウンファイルで、作業フローなどはマーメイドで
作成すること」

## 前提・制約

- 対象は `claude_playground` リポジトリ（GitHub: `Ponthai-project/claude_playground`）。
- 操作はGitHub Desktop中心。
- 信玄側に `gh` CLI未導入・GitHub API認証なしのため、リモート側のbranch protection設定を
  API経由で代行できない（殿ご自身のブラウザ操作が必要）。
- `docs/vscode-github-setup.md` に既存の日常フロー（PR作成含む）の記述があり、重複させない。

## 検討した選択肢

| 案 | 内容 | 却下/採用理由 |
|---|---|---|
| A. ドキュメントで注意喚起するのみ | 「mainを直接編集しない」という運用ルールを文書化するだけ | 既存の `vscode-github-setup.md` に既にこの注意書きがあるにもかかわらず、直近コミット履歴では実際に直接pushが多発していた実例がある。運用者の意識だけに頼る策は今回まさに機能しなかったため却下 |
| B. GitHub側のbranch protection ruleでmainへの直接pushを拒否 | Require a pull request before merging を有効化 | サーバー側で技術的に強制でき、Desktopの「ローカルmerge→push」を確実に弾ける。個人開発でも設定可能で、承認者数は0でよい。採用 |
| C. ローカルのpre-push hookで直接push/mergeを止める | `.git/hooks/pre-push` にチェックスクリプトを仕込む | クローンごとに個別設定が必要で、フックはリポジトリのバージョン管理外（`.git/`配下）にあるため他端末・再クローン時に引き継がれず、強制力が弱い。保険的な併用は可だが主策にはしない |

## 決定と決め手

主策としてBの「GitHub側branch protection rule（Require a pull request before merging）」を
採用した。決め手は、サーバー側で強制されるためDesktop・コマンドライン・他端末を問わず
mainへの直接push自体を技術的に拒否できる点を最優先したため。Aは今回の再発そのものであり
不採用、Cは強制力の一貫性に欠けるため主策からは外した。

## やったこと

- `git log --oneline` と `git branch -a` で現状を確認し、直近15コミット中PR経由マージは
  1件のみ（残りは直接push）であることを実証した。
- `gh` CLI・`curl`ともに使用不可のため、branch protectionの現在値はAPIで確認できず、
  git履歴からの間接証拠のみで診断した（未検証の推測ではなく実コミット履歴に基づく）。
- `docs/github-desktop-pr-workflow.md` を新規作成し、現状フローと推奨フローをMermaidの
  flowchartで対比、branch protectionの設定手順（Settings → Branches）とGitHub Desktop側の
  操作対応表を記載した。

## 見直し条件

殿がGitHub.com上でbranch protection ruleを設定した後、`git push origin main` を直接叩いて
実際に拒否されるか殿ご自身に実機確認いただく（信玄側は認証情報を持たず再現不可のため）。

## 未解決・次の一手

- branch protectionの有効化自体は殿の手動操作待ち。
- リモートに残る未整理ブランチ（`claude/asset-tokenization-blockchain-3bz4b7` 等3件）の
  マージ済み判定・削除は別件として次回対応。

## memory昇格候補

なし（本件はclaude_playground固有のリポジトリ設定であり、他プロジェクトへ横展開する
恒久ルールではないため）。
