import { webViewBridge } from "./webViewBridge";

export function wsUrlToHttpBase(wsUrl: string): string {
  const http = wsUrl.startsWith("wss://") ? "https://" : "http://";
  const noProtocol = wsUrl.replace(/^wss?:\/\//, "");
  const withoutWsPath = noProtocol.replace(/\/ws\/?$/, "");
  return `${http}${withoutWsPath}`;
}

export async function getApiBase(): Promise<string> {
  const wsUrl = await webViewBridge.getWebSocketUrl();
  return wsUrlToHttpBase(wsUrl);
}
