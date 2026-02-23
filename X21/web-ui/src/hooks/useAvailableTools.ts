import { useEffect, useState } from "react";
import { getApiBase } from "@/services/apiBase";
import { AVAILABLE_TOOLS } from "@/constants/tools";

interface Tool {
  id: string;
  name: string;
  description: string;
}

/**
 * Hook to fetch available tools from backend
 * Falls back to all AVAILABLE_TOOLS if fetch fails
 */
export function useAvailableTools() {
  const [availableTools, setAvailableTools] = useState(AVAILABLE_TOOLS);
  const [isLoading, setIsLoading] = useState(true);

  useEffect(() => {
    const fetchAvailableTools = async () => {
      try {
        const base = await getApiBase();
        const url = `${base}/api/tools`;
        const response = await fetch(url);
        if (!response.ok) {
          console.warn("Failed to fetch available tools, using defaults");
          setAvailableTools(AVAILABLE_TOOLS);
          return;
        }

        const data = await response.json();
        const backendToolIds = new Set(data.tools.map((tool: Tool) => tool.id));

        // Filter AVAILABLE_TOOLS to only include tools that backend supports
        const filteredTools = AVAILABLE_TOOLS.filter((tool) =>
          backendToolIds.has(tool.id),
        );

        setAvailableTools(filteredTools);
      } catch (error) {
        console.warn("Error fetching available tools:", error);
        // Fallback to all tools on error
        setAvailableTools(AVAILABLE_TOOLS);
      } finally {
        setIsLoading(false);
      }
    };

    fetchAvailableTools();
  }, []);

  return { availableTools, isLoading };
}
