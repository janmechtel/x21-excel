export type WSHandlers = {
  onOpen?: () => void;
  onMessage?: (data: any) => void;
  onClose?: (event: CloseEvent) => void;
  onError?: (event: Event) => void;
};

class WebSocketClient {
  private socket: WebSocket | null = null;
  private url: string;

  constructor(url: string) {
    this.url = url;
  }

  connect(handlers: WSHandlers = {}) {
    if (
      this.socket &&
      (this.socket.readyState === WebSocket.OPEN ||
        this.socket.readyState === WebSocket.CONNECTING)
    ) {
      return;
    }

    try {
      console.log("Attempting to connect to:", this.url);
      this.socket = new WebSocket(this.url);

      this.socket.onopen = () => {
        console.log("WebSocket connected successfully");
        handlers.onOpen?.();
      };

      this.socket.onmessage = (event: MessageEvent) => {
        try {
          const data =
            typeof event.data === "string"
              ? JSON.parse(event.data)
              : event.data;
          handlers.onMessage?.(data);
        } catch {
          handlers.onMessage?.(event.data);
        }
      };

      this.socket.onclose = (event: CloseEvent) => {
        handlers.onClose?.(event);
      };

      this.socket.onerror = (event: Event) => {
        console.error("WebSocket connection failed:", event);
        handlers.onError?.(event);
      };
    } catch (error) {
      console.error("Failed to create WebSocket:", error);
      handlers.onError?.(error as Event);
    }
  }

  send(message: any) {
    if (!this.socket || this.socket.readyState !== WebSocket.OPEN) return false;
    const payload =
      typeof message === "string" ? message : JSON.stringify(message);
    this.socket.send(payload);
    return true;
  }

  close(code?: number, reason?: string) {
    if (this.socket) {
      this.socket.close(code, reason);
      this.socket = null;
    }
  }
}

export const createWsClient = (url: string) => new WebSocketClient(url);
