import {
  BarChart3,
  BookOpen,
  Calculator,
  ClipboardCheck,
  Command,
  FileImage,
  type LucideIcon,
  Palette,
  Slash,
  Sparkles,
  Tags,
  Wand2,
} from "lucide-react";
import * as LucideIcons from "lucide-react";

// Pre-defined map for common icons (optimal performance)
const ICON_MAP: Record<string, LucideIcon> = {
  Slash,
  Command,
  BookOpen,
  Calculator,
  Palette,
  ClipboardCheck,
  FileImage,
  Tags,
  Sparkles,
  BarChart3,
  Wand2,
};

const DEFAULT_ICON = Slash;

/**
 * Resolves an icon name to a Lucide React component.
 *
 * Strategy:
 * 1. Check pre-defined ICON_MAP first (fastest)
 * 2. Attempt dynamic lookup from lucide-react exports
 * 3. Fallback to default Slash icon
 *
 * @param iconName - Case-sensitive Lucide icon name (e.g., "BookOpen", "Calculator")
 * @returns Lucide icon component
 */
export function resolveIcon(iconName?: string | null): LucideIcon {
  if (!iconName || typeof iconName !== "string") {
    return DEFAULT_ICON;
  }

  const normalizedName = iconName.trim();
  if (!normalizedName) return DEFAULT_ICON;

  // Check pre-defined map first (O(1) lookup)
  if (ICON_MAP[normalizedName]) {
    return ICON_MAP[normalizedName];
  }

  // Attempt dynamic lookup from lucide-react exports
  const LucideIconsTyped = LucideIcons as Record<string, unknown>;
  const dynamicIcon = LucideIconsTyped[normalizedName];

  if (dynamicIcon && typeof dynamicIcon === "function") {
    return dynamicIcon as LucideIcon;
  }

  // Fallback: log warning in dev mode
  if (import.meta.env.DEV) {
    console.warn(
      `[IconResolver] Unknown icon: "${iconName}". Falling back to Slash.`,
    );
  }

  return DEFAULT_ICON;
}
