export class UserCancellationError extends Error {
  public readonly requestId: string;

  constructor(message: string, requestId: string) {
    super(message);
    this.name = "UserCancellationError";
    this.requestId = requestId;

    // Maintains proper stack trace for where our error was thrown (only available on V8)
    if (Error.captureStackTrace) {
      Error.captureStackTrace(this, UserCancellationError);
    }
  }
}
