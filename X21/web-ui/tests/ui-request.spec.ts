import { expect, test } from "@playwright/test";
import { bypassAuthentication, waitForAppReady } from "./helpers/test-actions";
import { WebSocketMessageTypes } from "../src/types/chat";

test.describe("UI request in chat", () => {
  test.beforeEach(async ({ page }) => {
    await bypassAuthentication(page);
    await page.goto("/");
    await waitForAppReady(page);

    await page.evaluate(() => {
      const svc = (window as any).__x21WebSocketService;
      if (svc) {
        svc.sendToolResult = async (_toolUseId: string, output: any) => {
          (window as any).__uiRequestResult = output;
          return true;
        };
      }
    });
  });

  test("renders and submits a ui_request form card", async ({ page }) => {
    await page.evaluate(() => {
      const send = (window as any).__x21DebugSocketMessage;
      send?.({
        type: WebSocketMessageTypes.UI_REQUEST,
        toolUseId: "tool-ui-test",
        payload: {
          title: "Confirm details",
          description: "Tell us how to proceed",
          mode: "blocking",
          controls: [
            {
              id: "confirm",
              kind: "boolean",
              label: "Proceed?",
              required: true,
            },
            {
              id: "cadence",
              kind: "segmented",
              label: "Cadence",
              options: [
                { id: "monthly", label: "Monthly" },
                { id: "yearly", label: "Yearly" },
              ],
            },
          ],
        },
      });
    });

    const card = page.getByTestId("ui-request-card");
    await expect(card).toBeVisible();

    const continueButton = card.getByRole("button", { name: /continue/i });
    await expect(continueButton).toBeDisabled();

    await card.getByRole("button", { name: "Yes" }).click();
    await expect(continueButton).toBeEnabled();

    await continueButton.click();

    await expect(card.getByText(/Proceed\?: Yes/)).toBeVisible({
      timeout: 5000,
    });

    const result = await page.evaluate(() => (window as any).__uiRequestResult);
    expect(result).toEqual({ confirm: { value: true } });
  });
});
