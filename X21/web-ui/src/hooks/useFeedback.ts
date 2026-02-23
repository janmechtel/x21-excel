import { useState } from "react";

import type { ChatMessage } from "@/types/chat";
import { webSocketChatService } from "@/services/webSocketChatService";

interface UseFeedbackParams {
  chatHistory: ChatMessage[];
  setChatHistory: React.Dispatch<React.SetStateAction<ChatMessage[]>>;
  scrollToBottom?: () => void;
}

export function useFeedback({
  chatHistory,
  setChatHistory,
  scrollToBottom,
}: UseFeedbackParams) {
  const [showInlineComment, setShowInlineComment] = useState<string | null>(
    null,
  );
  const [commentText, setCommentText] = useState("");

  const submitScore = (messageId: string, score: "up" | "down") => {
    const message = chatHistory.find((msg) => msg.id === messageId);
    if (!message) {
      console.error("Message not found for scoring");
      return;
    }

    setChatHistory((prev) =>
      prev.map((msg) =>
        msg.id === messageId
          ? { ...msg, score: msg.score === score ? null : score }
          : msg,
      ),
    );

    try {
      const numericScore = score === "up" ? 1 : score === "down" ? -1 : 0;
      webSocketChatService
        .sendScoreOnly(numericScore)
        .then((success) => {
          if (success) {
            console.log("Score sent successfully via WebSocket to score:score");
          } else {
            console.error("Failed to send score - WebSocket not connected");
          }
        })
        .catch((error) => {
          console.error("Error sending score via WebSocket:", error);
        });
    } catch (error) {
      console.error("Error sending score via WebSocket:", error);
    }
  };

  const handleScoreMessage = (messageId: string, score: "up" | "down") => {
    const message = chatHistory.find((msg) => msg.id === messageId);
    if (!message) {
      console.error("Message not found for scoring");
      return;
    }

    submitScore(messageId, score);

    if (score === "down") {
      setShowInlineComment(messageId);
      setCommentText("");
      setTimeout(() => {
        scrollToBottom?.();
      }, 50);
    }
  };

  const handleCommentSubmit = (messageId: string) => {
    const message = chatHistory.find((msg) => msg.id === messageId);
    if (!message) {
      console.error("Message not found for feedback");
      return;
    }

    try {
      webSocketChatService
        .sendFeedback(commentText)
        .then((success) => {
          if (success) {
            console.log(
              "Feedback sent successfully via WebSocket to score:feedback",
            );
          } else {
            console.error("Failed to send feedback - WebSocket not connected");
          }
        })
        .catch((error) => {
          console.error("Error sending feedback via WebSocket:", error);
        });
    } catch (error) {
      console.error("Error sending feedback via WebSocket:", error);
    }

    setShowInlineComment(null);
    setCommentText("");
  };

  const handleCommentCancel = () => {
    setShowInlineComment(null);
    setCommentText("");
  };

  return {
    showInlineComment,
    commentText,
    setCommentText,
    handleScoreMessage,
    handleCommentSubmit,
    handleCommentCancel,
  };
}
