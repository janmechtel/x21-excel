import { useCallback, useEffect, useRef, useState } from "react";

import type { ToolGroupDecisions, ToolGroups } from "@/types/chat";

interface Params {
  chatHistory: unknown[];
  isStreaming: boolean;
  getPendingToolCount: () => number;
  toolGroups: ToolGroups;
  toolGroupDecisions: ToolGroupDecisions;
  chatContainerRef: React.RefObject<HTMLDivElement>;
  responseEndRef: React.RefObject<HTMLDivElement>;
}

export function useChatViewport({
  chatHistory,
  isStreaming,
  getPendingToolCount,
  toolGroups,
  toolGroupDecisions,
  chatContainerRef,
  responseEndRef,
}: Params) {
  const [isNearBottom, setIsNearBottom] = useState(true);
  const [activeStickyMessageId, setActiveStickyMessageId] = useState<
    string | null
  >(null);
  const isNearBottomRef = useRef(true);
  const scrollTimeoutRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  const scrollToBottom = useCallback(
    (behavior: ScrollBehavior = "smooth") => {
      responseEndRef.current?.scrollIntoView({ behavior });
      isNearBottomRef.current = true;
      setIsNearBottom(true);
    },
    [responseEndRef],
  );

  const computeIsNearBottom = useCallback(() => {
    if (!chatContainerRef.current) return true;
    const { scrollTop, scrollHeight, clientHeight } = chatContainerRef.current;
    const threshold = 100; // Allow 100px tolerance for "near bottom" detection.
    return scrollHeight - scrollTop - clientHeight < threshold;
  }, [chatContainerRef]);

  const handleScroll = useCallback(() => {
    const nearBottom = computeIsNearBottom();
    isNearBottomRef.current = nearBottom;
    if (scrollTimeoutRef.current) {
      clearTimeout(scrollTimeoutRef.current);
    }
    scrollTimeoutRef.current = setTimeout(() => {
      setIsNearBottom(nearBottom);
      scrollTimeoutRef.current = null;
    }, 50);
  }, [computeIsNearBottom]);

  // eslint-disable-next-line react-hooks/exhaustive-deps
  // isNearBottomRef is a ref and doesn't need to be in dependencies.
  useEffect(() => {
    // Only scroll if there are messages (don't scroll on empty chat)
    if (chatHistory.length > 0 && isNearBottomRef.current) {
      scrollToBottom("auto");
    }
  }, [chatHistory, scrollToBottom]);

  useEffect(() => {
    if (getPendingToolCount() > 0 && isNearBottomRef.current) {
      scrollToBottom("auto");
    }
  }, [toolGroups, toolGroupDecisions, scrollToBottom, getPendingToolCount]);

  useEffect(() => {
    if (isStreaming) {
      // Ensure we scroll when streaming starts (if there are messages)
      if (chatHistory.length > 0 && isNearBottomRef.current) {
        scrollToBottom("auto");
      }
    }
  }, [isStreaming, chatHistory.length, scrollToBottom]);

  useEffect(() => {
    const observerOptions = {
      root: chatContainerRef.current,
      threshold: 0.1,
      rootMargin: "0px 0px -90% 0px",
    };

    const observer = new IntersectionObserver((entries) => {
      entries.forEach((entry) => {
        const messageId = entry.target.getAttribute("data-message-id");
        if (entry.isIntersecting && messageId) {
          setActiveStickyMessageId(messageId);
        }
      });
    }, observerOptions);

    const conversationGroups = document.querySelectorAll(".conversation-group");
    conversationGroups.forEach((group) => observer.observe(group));

    return () => observer.disconnect();
  }, [chatHistory, chatContainerRef]);

  useEffect(() => {
    return () => {
      if (scrollTimeoutRef.current) {
        clearTimeout(scrollTimeoutRef.current);
        scrollTimeoutRef.current = null;
      }
    };
  }, []);

  return {
    isNearBottom,
    activeStickyMessageId,
    handleScroll,
    scrollToBottom,
  };
}
