export function getCompactConversationSystemMessage(): string {
  return `You are an AI assistant that compacts conversations about Microsoft Excel work. Summarize the conversation concisely while preserving key information.

- Omit courtesy phrases like "Thanks for your request!" or "Perfect! Let me do it!"
- Provide the summary directly with no introduction`;
}
