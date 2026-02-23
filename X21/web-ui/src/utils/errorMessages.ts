export function getErrorDisplayMessage(errorType: string): {
  title: string;
  description: string;
} {
  switch (errorType) {
    case "OVERLOADED":
      return {
        title: "Service Overloaded",
        description:
          "The AI service is currently experiencing high traffic. Please wait a moment and try again.",
      };
    case "RATE_LIMIT_ERROR":
      return {
        title: "Rate Limit Exceeded",
        description:
          "You have sent too many requests in a short time. Please wait a moment before trying again.",
      };
    case "INVALID_REQUEST_ERROR":
      return {
        title: "Invalid Request",
        description: `
          Please Restart the Chat.

          Pausible Causes:
          Either the request size is too big.
          The Attached PDF > 100 pages.
          Or we are in a inconsistent state.
          `,
      };
    case "REQUEST_TOO_LARGE":
      return {
        title: "Request Size Too Large",
        description:
          "The size of your request exceeds the maximum allowed. Please reduce the content or file size and/or restart the chat.",
      };
    case "CANCELLED":
      return {
        title: "Request Cancelled",
        description: "The request was cancelled.",
      };
    default:
      return {
        title: "Unexpected Error",
        description:
          "An unexpected error occurred while processing your request. Try again, and if it persists, restart the chat.",
      };
  }
}
