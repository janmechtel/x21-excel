import { ChatMessage } from "@/types/chat";
import { ClaudeContentTypes } from "@/types/chat";

export const formatThinkingDuration = (
  startTime: number,
  endTime?: number,
): string => {
  const end = endTime || Date.now();
  const duration = Math.round((end - startTime) / 1000);
  return `${duration}s`;
};

export const shouldShowCollapseIndicator = (content: string): boolean => {
  const lines = content.split("\n");
  const lineCount = lines.length;
  const charCount = content.length;

  return (
    lineCount > 3 || charCount > 300 || lines.some((line) => line.length > 100)
  );
};

export const groupMessagesIntoConversations = (chatHistory: ChatMessage[]) => {
  const conversations: Array<{
    userMessage: ChatMessage;
    assistantMessages: ChatMessage[];
  }> = [];

  let currentConversation: {
    userMessage: ChatMessage;
    assistantMessages: ChatMessage[];
  } | null = null;

  for (const message of chatHistory) {
    if (message.role === "user") {
      if (currentConversation) {
        conversations.push(currentConversation);
      }
      currentConversation = {
        userMessage: message,
        assistantMessages: [],
      };
    } else if (message.role === "assistant" && currentConversation) {
      currentConversation.assistantMessages.push(message);
    }
  }

  if (currentConversation) {
    conversations.push(currentConversation);
  }

  return conversations;
};

export const findFirstTool = (message: ChatMessage) => {
  if (!message.contentBlocks) return null;

  for (const block of message.contentBlocks) {
    if (
      block.type === ClaudeContentTypes.TOOL_USE &&
      block.isComplete &&
      block.toolUseId &&
      block.toolName
    ) {
      return {
        toolUseId: block.toolUseId,
        toolName: block.toolName,
        toolNumber: block.toolNumber || 0,
      };
    }
  }
  return null;
};

export const findFirstToolFromMessage = (
  chatHistory: ChatMessage[],
  userMessage: ChatMessage,
) => {
  const userMessageIndex = chatHistory.findIndex(
    (msg) => msg.id === userMessage.id,
  );
  if (userMessageIndex === -1 || userMessageIndex >= chatHistory.length - 1) {
    return null;
  }

  const assistantMessage = chatHistory[userMessageIndex + 1];
  if (assistantMessage.role !== "assistant") {
    return null;
  }

  return findFirstTool(assistantMessage);
};

export const isConversationEffectivelyReverted = (
  chatHistory: ChatMessage[],
  revertedConversations: Set<string>,
  userMessage: ChatMessage,
) => {
  const conversations = groupMessagesIntoConversations(chatHistory);
  const currentConversationIndex = conversations.findIndex(
    (conv) => conv.userMessage.id === userMessage.id,
  );

  if (currentConversationIndex === -1) return false;

  if (revertedConversations.has(userMessage.id)) {
    return true;
  }

  for (let i = 0; i < currentConversationIndex; i++) {
    if (revertedConversations.has(conversations[i].userMessage.id)) {
      return true;
    }
  }

  return false;
};
