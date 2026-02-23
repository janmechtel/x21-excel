import { createLogger } from "../utils/logger.ts";

const logger = createLogger("ErrorHandler");

export interface ApiError extends Error {
  status?: number;
  errorType?: string;
}

export class ErrorHandler {
  static createResponse(error: unknown, defaultStatus: number = 500): Response {
    const errorMessage = error instanceof Error
      ? error.message
      : "Unknown error";
    const errorName = error instanceof Error ? error.name : "UnknownError";
    const status = (error as ApiError)?.status || defaultStatus;
    const errorType = (error as ApiError)?.errorType || errorName;

    logger.error("API Error:", {
      error: errorMessage,
      errorName: errorName,
      status: status,
      errorType: errorType,
      stack: error instanceof Error ? error.stack : undefined,
    });

    return new Response(
      JSON.stringify({
        success: false,
        error: errorMessage,
        errorType: errorType,
      }),
      {
        status: status,
        headers: {
          "Content-Type": "application/json",
          "Access-Control-Allow-Origin": "*",
        },
      },
    );
  }

  static createSuccessResponse(data: any, status: number = 200): Response {
    return new Response(
      JSON.stringify({
        success: true,
        ...data,
      }),
      {
        status: status,
        headers: {
          "Content-Type": "application/json",
          "Access-Control-Allow-Origin": "*",
        },
      },
    );
  }

  static createValidationError(message: string): ApiError {
    const error = new Error(message) as ApiError;
    error.status = 400;
    error.errorType = "ValidationError";
    return error;
  }

  static createNotFoundError(message: string): ApiError {
    const error = new Error(message) as ApiError;
    error.status = 404;
    error.errorType = "NotFoundError";
    return error;
  }
}
