import { getApiBase } from "./apiBase";

type WorkbookSummary = {
  id: string;
  summaryText: string;
  createdAt: number;
  sheetsAffected?: number;
  comparisonType?: "self" | "external";
  comparisonFilePath?: string | null;
  comparisonFileModifiedAt?: number | null;
};

export async function fetchChangelogSummaries(
  workbookKey: string,
  limit = 50,
): Promise<WorkbookSummary[]> {
  try {
    const baseUrl = await getApiBase();
    const response = await fetch(
      `${baseUrl}/api/workbook-summaries?workbookKey=${encodeURIComponent(
        workbookKey,
      )}&limit=${limit}`,
    );

    if (!response.ok) {
      console.error("Failed to load saved summaries:", response.statusText);
      return [];
    }

    const data = await response.json();
    return data.summaries || [];
  } catch (error) {
    console.error("Error loading saved summaries:", error);
    return [];
  }
}

export async function updateChangelogEntry(
  id: string,
  summaryText: string,
): Promise<boolean> {
  try {
    const baseUrl = await getApiBase();
    const response = await fetch(`${baseUrl}/api/workbook-summaries`, {
      method: "PATCH",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ id, summaryText }),
    });

    if (!response.ok) {
      console.error("Failed to update workbook summary:", response.statusText);
      return false;
    }

    return true;
  } catch (error) {
    console.error("Error updating workbook summary:", error);
    return false;
  }
}

export async function deleteChangelogEntry(id: string): Promise<boolean> {
  try {
    const baseUrl = await getApiBase();
    const response = await fetch(`${baseUrl}/api/workbook-summaries`, {
      method: "DELETE",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ id }),
    });

    if (!response.ok) {
      console.error("Failed to delete workbook summary:", response.statusText);
      return false;
    }

    return true;
  } catch (error) {
    console.error("Error deleting workbook summary:", error);
    return false;
  }
}
