import { useCallback, useRef, useState } from "react";

import type { AttachedFile } from "@/types/chat";
import { webViewBridge } from "@/services/webViewBridge";

const SUPPORTED_FILE_TYPES = [
  "application/pdf",
  "image/gif",
  "image/png",
  "image/jpeg",
  "image/jpg",
  "image/webm",
];

export function useAttachments() {
  const [attachedFiles, setAttachedFiles] = useState<AttachedFile[]>([]);
  const [isConvertingFile, setIsConvertingFile] = useState(false);
  const [isDragOver, setIsDragOver] = useState(false);
  const [_, setDragCounter] = useState(0);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const convertFileToBase64 = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const result = reader.result as string;
        const base64 = result.split(",")[1];
        resolve(base64);
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  };

  const handleFileUpload = async (files: FileList | File[]) => {
    const fileArray = Array.from(files);
    const supportedFiles = fileArray.filter((file) =>
      SUPPORTED_FILE_TYPES.includes(file.type.toLowerCase()),
    );

    if (supportedFiles.length === 0) {
      return;
    }

    if (supportedFiles.length !== fileArray.length) {
      alert(
        "Some files were skipped because they are not in supported formats. Supported formats: PDF, GIF, PNG, JPEG, WEBM.",
      );
    }

    setIsConvertingFile(true);

    try {
      const convertedFiles = await Promise.all(
        supportedFiles.map(async (file) => ({
          name: file.name,
          type: file.type,
          size: file.size,
          base64: await convertFileToBase64(file),
        })),
      );

      setAttachedFiles((prev) => [...prev, ...convertedFiles]);
    } catch (error) {
      console.error("Error converting files to base64:", error);
      alert("Error processing files. Please try again.");
    } finally {
      setIsConvertingFile(false);
    }
  };

  const handlePaste = async (e: React.ClipboardEvent) => {
    const items = Array.from(e.clipboardData.items);
    const hasText = items.some(
      (item) => item.type === "text/plain" || item.type === "text/html",
    );

    const imageItems = items.filter((item) =>
      SUPPORTED_FILE_TYPES.includes(item.type.toLowerCase()),
    );
    if (hasText && imageItems.length > 0) {
      return;
    }
    if (imageItems.length === 0) {
      return;
    }

    e.preventDefault();
    setIsConvertingFile(true);

    try {
      const convertedFiles = await Promise.all(
        imageItems.map(async (item, index) => {
          const file = item.getAsFile();
          if (!file) return null;

          const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
          const extension = file.type.split("/")[1] || "png";
          const fileName = `pasted-image-${timestamp}-${
            index + 1
          }.${extension}`;

          return {
            name: fileName,
            type: file.type,
            size: file.size,
            base64: await convertFileToBase64(file),
          };
        }),
      );

      const validFiles = convertedFiles.filter(
        (file) => file !== null,
      ) as AttachedFile[];
      if (validFiles.length > 0) {
        setAttachedFiles((prev) => [...prev, ...validFiles]);
      }
    } catch (error) {
      console.error("Error processing pasted images:", error);
      alert("Error processing pasted images. Please try again.");
    } finally {
      setIsConvertingFile(false);
    }
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.dataTransfer.types.includes("Files")) {
      setIsDragOver(true);
    }
  };

  const handleDragEnter = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.dataTransfer.types.includes("Files")) {
      setDragCounter((prev) => prev + 1);
      setIsDragOver(true);
    }
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragCounter((prev) => {
      const newCounter = prev - 1;
      if (newCounter <= 0) {
        setIsDragOver(false);
        return 0;
      }
      return newCounter;
    });
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(false);
    setDragCounter(0);

    const files = e.dataTransfer.files;
    if (files.length > 0) {
      void handleFileUpload(files);
    }
  };

  const handleFileInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      void handleFileUpload(files);
    }
    e.target.value = "";
  };

  const removeAttachedFile = (index: number) => {
    setAttachedFiles((prev) => prev.filter((_, i) => i !== index));
  };

  const clearAllAttachedFiles = () => {
    setAttachedFiles([]);
  };

  const formatFileSize = (bytes: number): string => {
    if (bytes === 0) return "0 Bytes";
    const k = 1024;
    const sizes = ["Bytes", "KB", "MB", "GB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
  };

  const openFileDialog = useCallback(() => {
    // Notify C# that file dialog is opening
    webViewBridge.send("fileDialogOpening", {});

    // Open the file dialog
    fileInputRef.current?.click();

    // C# side will detect when focus returns and automatically restore focus to WebView
  }, []);

  return {
    attachedFiles,
    setAttachedFiles,
    isConvertingFile,
    isDragOver,
    fileInputRef,
    handlePaste,
    handleDragOver,
    handleDragEnter,
    handleDragLeave,
    handleDrop,
    handleFileInputChange,
    removeAttachedFile,
    clearAllAttachedFiles,
    formatFileSize,
    openFileDialog,
  };
}
