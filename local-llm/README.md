# ローカルLLM チャットアプリ

Java + LangChain4j + Ollama で動くローカルLLMチャットアプリです。

## 必要なもの

- Java 17+
- Maven 3.6+
- [Ollama](https://ollama.com/) （ローカルLLMサーバー）

## セットアップ手順

### 1. Ollamaをインストール

https://ollama.com/download から Windows版をダウンロードしてインストール。

### 2. モデルをダウンロード

ターミナルで以下を実行（3Bモデル、約2GB）:

```bash
ollama pull llama3.2:3b
```

CPUのみの場合はより小さいモデルが快適:

```bash
ollama pull phi3.5        # 3.8B、軽量で高品質（英語向き）
ollama pull phi4-mini     # 3.8B、最新版
```

### 3. Ollamaを起動

インストール後は自動起動します。手動起動は:

```bash
ollama serve
```

動作確認: http://localhost:11434 にアクセスして `Ollama is running` と表示されればOK。

### 4. アプリをビルドして実行

```bash
cd local-llm
mvn package -q
java -jar target/local-llm-1.0-SNAPSHOT.jar
```

または Maven で直接実行:

```bash
mvn compile exec:java -Dexec.mainClass="com.playground.llm.ChatApp"
```

## モデルの変更

`ChatApp.java` の `MODEL_NAME` を変更するだけ:

```java
private static final String MODEL_NAME = "phi3.5";  // 例
```

## トラブルシューティング

| エラー | 原因 | 対処 |
|--------|------|------|
| Ollamaに接続できません | Ollamaが未起動 | `ollama serve` を実行 |
| モデルが見つからない | モデル未取得 | `ollama pull <モデル名>` を実行 |
| 応答が遅い | CPUのみで大きいモデルを実行中 | 小さいモデル（3B以下）に変更 |
