package com.playground.llm;

import dev.langchain4j.model.ollama.OllamaChatModel;
import dev.langchain4j.model.chat.ChatLanguageModel;
import dev.langchain4j.data.message.AiMessage;
import dev.langchain4j.data.message.UserMessage;
import dev.langchain4j.model.output.Response;

import java.util.Scanner;

public class ChatApp {

    // Ollamaのデフォルトエンドポイント
    private static final String OLLAMA_URL = "http://localhost:11434";

    // 使用するモデル（後で変更可能）
    private static final String MODEL_NAME = "llama3.2:3b";

    public static void main(String[] args) {
        System.out.println("=== ローカルLLM チャットアプリ ===");
        System.out.println("モデル: " + MODEL_NAME);
        System.out.println("Ollama URL: " + OLLAMA_URL);
        System.out.println("終了するには 'exit' または 'quit' と入力してください");
        System.out.println("==========================================\n");

        ChatLanguageModel model = OllamaChatModel.builder()
                .baseUrl(OLLAMA_URL)
                .modelName(MODEL_NAME)
                .temperature(0.7)
                .build();

        Scanner scanner = new Scanner(System.in);

        while (true) {
            System.out.print("あなた: ");
            String input = scanner.nextLine().trim();

            if (input.isEmpty()) continue;
            if (input.equalsIgnoreCase("exit") || input.equalsIgnoreCase("quit")) {
                System.out.println("終了します。");
                break;
            }

            System.out.print("AI: ");
            try {
                Response<AiMessage> response = model.generate(UserMessage.from(input));
                System.out.println(response.content().text());
            } catch (Exception e) {
                System.out.println("[エラー] Ollamaに接続できません。Ollamaが起動しているか確認してください。");
                System.out.println("詳細: " + e.getMessage());
            }
            System.out.println();
        }

        scanner.close();
    }
}
