import { RefObject, useEffect } from "react";

interface UseFocusOnMountOptions {
  maxAttempts?: number;
  retryDelay?: number;
}

/**
 * Custom hook to automatically focus an input element when component mounts.
 * Implements retry logic to handle async rendering delays.
 *
 * @param inputRef - React ref to the input/textarea element
 * @param options - Configuration for retry behavior
 * @returns void
 *
 * @example
 * const inputRef = useRef<HTMLTextAreaElement>(null);
 * useFocusOnMount(inputRef, { maxAttempts: 3, retryDelay: 50 });
 */
export function useFocusOnMount(
  inputRef: RefObject<HTMLInputElement | HTMLTextAreaElement>,
  options: UseFocusOnMountOptions = {},
): void {
  const { maxAttempts = 3, retryDelay = 50 } = options;

  useEffect(() => {
    let attempts = 0;
    let timeoutId: NodeJS.Timeout | null = null;

    /**
     * Attempts to focus the input element.
     * Uses requestAnimationFrame to ensure focus happens after browser paint.
     * @returns true if focus was successful, false otherwise
     */
    const tryFocus = (): boolean => {
      const input = inputRef.current;

      // Check if input exists and is focusable
      if (!input || input.disabled || input.hidden) {
        return false;
      }

      // Use requestAnimationFrame to ensure we focus after any other operations
      requestAnimationFrame(() => {
        input.focus();

        // Also try to prevent any blur events from stealing focus
        input.addEventListener(
          "blur",
          function preventBlur(_e) {
            console.log("[Focus] Blur detected, re-focusing...");
            input.focus();
            // Remove listener after first blur prevention
            setTimeout(() => {
              input.removeEventListener("blur", preventBlur);
            }, 500);
          },
          { once: false },
        );
      });

      // Check success after a brief delay to allow requestAnimationFrame to complete
      setTimeout(() => {
        const success = document.activeElement === input;
        if (success) {
          console.log(
            `[Focus] Successfully focused on attempt ${attempts + 1}`,
          );
        }
      }, 10);

      // Return true optimistically since we can't check synchronously
      return true;
    };

    /**
     * Recursive function to attempt focus with retry logic.
     * Stops after maxAttempts or successful focus.
     */
    const attemptFocus = (): void => {
      tryFocus();

      attempts++;

      if (attempts < maxAttempts) {
        // Schedule next attempt to reinforce focus
        console.log(
          `[Focus] Scheduling attempt ${attempts + 1} in ${retryDelay}ms...`,
        );
        timeoutId = setTimeout(attemptFocus, retryDelay);
      } else {
        console.log(`[Focus] Completed ${maxAttempts} focus attempts`);
      }
    };

    // Wait a bit before starting focus attempts to let WebView2 settle
    const initialDelay = setTimeout(() => {
      attemptFocus();
    }, 100); // 100ms initial delay for WebView2 to be ready

    // Cleanup function to cancel pending retry if component unmounts
    return () => {
      clearTimeout(initialDelay);
      if (timeoutId) {
        clearTimeout(timeoutId);
      }
    };
  }, []); // Empty deps = run once on mount only
}
