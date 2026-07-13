# まさおAIじっくり解説ch — AI駆動開発・ハーネスエンジニアリング関連動画まとめ

YouTubeチャンネル「[まさおAIじっくり解説ch](https://www.youtube.com/@ai_masaou)」（運営：AIサービス開発エンジニア まさお）の動画のうち、**AI駆動開発**および特に**ハーネスエンジニアリング**（AIエージェントを自律的かつ高精度に動かすための「実行環境・制約・フィードバック設計」）について解説している動画をまとめました。

> ⚠️ 注記: この一覧はWeb検索によって収集したものです。YouTube側のアクセス制限により各動画ページを直接開いて内容を検証することができなかったため、タイトル・公開時期は検索結果のスニペットに基づいています。視聴前に念のためチャンネル上で存在を確認してください。

---

## 1. ハーネスエンジニアリング特集動画（本題）

このチャンネルは「ハーネスエンジニアリング」を継続的なテーマとして複数回にわたり深掘りしています。時系列順に並べると、概念紹介 → 実践設計 → 総集編という流れになっています。

| # | タイトル | 公開時期 | 内容の概要 | URL |
|---|---|---|---|---|
| 1 | AIエージェントを使う人・作る人のための「ハーネスエンジニアリング」詳解 | 2026/2/28 | ハーネスエンジニアリングの概念を初めて詳しく解説した回。エージェントを「使う人」「作る人」双方の視点から整理 | [視聴](https://www.youtube.com/watch?v=VGL45KZGOkM) |
| 2 | 【全員必見】ハーネスエンジニアリングとは？AIエージェントの必須知識を徹底解説 | 2026/4/5 | コンテキスト管理・ツール制御・ガードレールなど、AIエージェントの実行基盤設計に必要な基礎知識を整理 | [視聴](https://www.youtube.com/watch?v=JvCIgFPgOlk) |
| 3 | 【AI駆動開発】AI自走環境構築・運用スペシャル #1 〜ハーネスエンジニアリングへの入口〜 | 2026/4/6 | AIが自走（自律的に開発を進める）できる環境構築をテーマにしたシリーズの第1回。ハーネスエンジニアリングへの導入 | [視聴](https://www.youtube.com/watch?v=PSJex2e4JWI) |
| 4 | 【AI駆動開発】AI自走環境整備・運用スペシャル #3 | — | 上記シリーズの続編（#3）。AI自走環境の整備・運用について継続解説 | [視聴](https://www.youtube.com/watch?v=4yPTAn1_okU) |
| 5 | Claude Codeハーネスエンジニアリング まず抑えるべき基礎知識 | — | Claude Codeに特化したハーネス設計の基礎知識回 | [視聴](https://www.youtube.com/watch?v=hxCEABf0Nfc) |
| 6 | 【ClaudeCode】思い通りに動かす設計力「ハーネス・エンジニアリング」入門書 | 2026/4/26 | Claude Codeを「思い通りに」動かすための設計力としてハーネスエンジニアリングを入門書形式で解説 | [視聴](https://www.youtube.com/watch?v=2kTyq4IQEXs) |
| 7 | 【自律開発】Claude Codeでハーネス設計すると開発が自動で進む！【Claude Codeハーネスエンジニアリング】 | — | ハーネス設計によって開発が自動的に進行する自律開発の実例を紹介 | [視聴](https://www.youtube.com/watch?v=Wfz-gdWcItM) |
| 8 | 【Claude Code ハーネスエンジニアリング完全版】同じAIでも「使いこなせる人」と「使えない人」の差｜概念・メリット・設計12パターン | 2026/5/5 | 約70分の講座形式。概念・メリットに加え、実践的な**設計12パターン**を体系的に解説する集大成的な回 | [視聴](https://www.youtube.com/watch?v=lbNVqcBNyH4) |
| 9 | 【保存版】Claude Code最強アプデ "Dynamic Workflows"｜1,000体のAIを並列で自律運用する方法。プロンプトもコンテキストも超える"ハーネス"を初心者向けに完全解説 | — | Claude Codeのアップデート機能"Dynamic Workflows"を題材に、大規模並列エージェント運用とハーネス設計を解説 | [視聴](https://www.youtube.com/watch?v=vVaOblwZ_k8) |
| 10 | 【神回】LINE Harnessの新機能 Claude Codeで自動化した方法を徹底解説 | — | LINEが公開した"Harness"関連新機能とClaude Codeを組み合わせた自動化事例 | [視聴](https://www.youtube.com/watch?v=3Ov00YZI9ts) |
| 11 | 【ハーネスエンジニアリング完全解説】ここまでを時系列で整理（Hashicorp, Langchain, OpenAI, マーティン・ファウラー etc.）／作り手と使い手で違う"言葉の定義" | 2026/5/11 | Hashicorp・Langchain・OpenAI・マーティン・ファウラーなど各所の言説を時系列に整理し、「ハーネス」という言葉の定義の違いを総括する回 | [視聴](https://www.youtube.com/watch?v=qUxjJywT1aw) |
| 12 | ハーネスエンジニアリングの起源（Shorts） | — | ハーネスエンジニアリングという言葉・概念の起源を短くまとめたショート動画 | [視聴](https://www.youtube.com/shorts/DMswbZq7bPg) |

### 全体の流れ
1. **導入編**（#1, #2, #3）: ハーネスエンジニアリングとは何か、なぜ必要かという概念解説
2. **Claude Code特化編**（#5, #6, #7）: Claude Codeでの具体的なハーネス設計・自律開発の実践
3. **総集編・応用編**（#8, #9, #10, #11, #12）: 12の設計パターン、大規模並列運用、他社動向を含む総括・応用

---

## 2. AI駆動開発 全般に関する動画（ハーネス以外）

ハーネス以外にも、Claude Code / Codexを用いたAI駆動開発の実践的な使い方を扱った動画が多数あります。

| タイトル | 概要 | URL |
|---|---|---|
| 個人開発を志す人必見！本当は教えたくない全身全霊のまさお式AI駆動サービス開発の全てを1分で解説 | AI駆動によるサービス開発の全体像を1分で凝縮して解説 | [視聴](https://www.youtube.com/watch?v=VHOmt5Kag8k) |
| 【超有料級】ClaudeCodeを本気の徹底解説！全ての基礎から知るべき全てを網羅した完全解説 | Claude Codeの基礎から実践までを網羅する総合解説回 | [視聴](https://www.youtube.com/watch?v=WRbG_22RfeI) |
| 【神回！】『Claude Code』完全入門マニュアル！【初心者でも絶対に理解できる！】 | 初心者向けのClaude Code完全入門マニュアル | [視聴](https://www.youtube.com/watch?v=DkClEbyXyq4) |
| ClaudeCodeやCodexのサブエージェントの使い方！Googleの論文の結果や普段使っている方法を解説してみた | サブエージェント活用法。Google論文の知見も紹介 | [視聴](https://www.youtube.com/watch?v=3g34nnMBa0U) |
| 【前編】Claude Codeサブエージェント完全ガイド｜全体像・基本機能・メリットを徹底解説 | サブエージェント機能の全体像と基本メリットの解説（前編） | [視聴](https://www.youtube.com/watch?v=IAqgR-mklks) |
| 【Claude Code サブエージェント 完全版】AI時代はサブエージェントをいかにうまく使うかで決まる | サブエージェント活用の完全版・フルコース | [視聴](https://www.youtube.com/watch?v=vWNC47fLwgw) |
| 【実験】4体のAIエージェントでYouTube運営させてみた！【Claude Code】 | 4体のAIエージェントによるYouTubeチャンネル運営実験 | [視聴](https://www.youtube.com/watch?v=MxYK2l_A43c) |

---

## まとめ

まさおAIじっくり解説chでは、2026年2月頃から「ハーネスエンジニアリング」を継続テーマとして扱っており、特に**Claude Codeを使ったAIエージェントの実行環境設計**（コンテキスト管理・ツール制御・ガードレール・設計パターン）に焦点を当てています。中でも以下の2本は特に重要です。

- **入門に最適**: [AIエージェントを使う人・作る人のための「ハーネスエンジニアリング」詳解](https://www.youtube.com/watch?v=VGL45KZGOkM)（初出・概念解説）
- **総まとめとして最適**: [【Claude Code ハーネスエンジニアリング完全版】](https://www.youtube.com/watch?v=lbNVqcBNyH4)（設計12パターンを含む70分の集大成）／[【ハーネスエンジニアリング完全解説】ここまでを時系列で整理](https://www.youtube.com/watch?v=qUxjJywT1aw)（業界全体の議論の整理）
