// Custom error class for tool execution failures
export class ToolExecutionError extends Error {
  constructor(
    message: string,
    public readonly toolId: string,
    public readonly toolName: string,
    public readonly originalError: Error,
  ) {
    super(message);
    this.name = "ToolExecutionError";
  }
}
