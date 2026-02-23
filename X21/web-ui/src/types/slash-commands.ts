export interface SlashCommandDefinition {
  id: string;
  name: string;
  title: string;
  description: string;
  prompt: string;
  icon?: string;
  category?: string;
  keywords?: string[];
  requiresInput?: boolean;
  inputPlaceholder?: string;
  defaultInput?: string;
}

export type SlashCommandMap = Record<string, SlashCommandDefinition>;
