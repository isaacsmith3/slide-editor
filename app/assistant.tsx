"use client";

import { useState, useRef, useEffect } from "react";
import { AssistantRuntimeProvider, useThread } from "@assistant-ui/react";
import {
  useChatRuntime,
  AssistantChatTransport,
} from "@assistant-ui/react-ai-sdk";
import { Thread } from "@/components/assistant-ui/thread";
import {
  SidebarInset,
  SidebarProvider,
  SidebarTrigger,
} from "@/components/ui/sidebar";
import { ThreadListSidebar } from "@/components/assistant-ui/threadlist-sidebar";
import { Separator } from "@/components/ui/separator";
import {
  Breadcrumb,
  BreadcrumbItem,
  BreadcrumbLink,
  BreadcrumbList,
  BreadcrumbPage,
  BreadcrumbSeparator,
} from "@/components/ui/breadcrumb";
import { Button } from "@/components/ui/button";
import { PptxViewer } from "@/components/PptxViewer";
import axios from "axios";

function ViewerRefreshHandler({
  fileUrl,
  setFileUrl,
}: {
  fileUrl: string;
  setFileUrl: (url: string) => void;
}) {
  const thread = useThread();
  const messages = thread.messages;
  const lastProcessedRef = useRef<string>("");
  const isInitialMount = useRef(true);

  useEffect(() => {
    // Skip on initial mount to avoid triggering refresh on first load
    if (isInitialMount.current) {
      isInitialMount.current = false;
      return;
    }

    if (messages.length > 0) {
      const lastMessage = messages[messages.length - 1];
      // Check if the last message is from assistant and contains success message
      if (lastMessage.role === "assistant" && lastMessage.content) {
        const content =
          typeof lastMessage.content === "string"
            ? lastMessage.content
            : Array.isArray(lastMessage.content)
              ? lastMessage.content.find(
                  (part: { type: string; text: string }) =>
                    part.type === "text",
                )?.text || ""
              : "";

        const messageId = lastMessage.id || "";

        // If message indicates success and we haven't processed it yet, refresh the viewer
        if (
          messageId !== lastProcessedRef.current &&
          (content.includes("Changes applied successfully") ||
            content.includes("Command executed") ||
            content.includes("successfully"))
        ) {
          lastProcessedRef.current = messageId;
          // Backend now waits for image conversion, so shorter delay is sufficient
          // Still add a small delay to ensure file system operations are fully complete
          setTimeout(() => {
            setFileUrl(`http://localhost:3001/api/pptx?v=${Date.now()}`);
          }, 1000); // Reduced since backend waits for conversion
        }
      }
    }
  }, [messages, setFileUrl]);

  return null;
}

export const Assistant = () => {
  const [isPptxUploaded, setIsPptxUploaded] = useState(false);
  const [fileUrl, setFileUrl] = useState<string>("");
  const [uploading, setUploading] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileUpload = async (
    event: React.ChangeEvent<HTMLInputElement>,
  ) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setUploading(true);
    const formData = new FormData();
    formData.append("pptx", file);

    try {
      const response = await axios.post(
        "http://localhost:3001/api/upload",
        formData,
        {
          headers: {
            "Content-Type": "multipart/form-data",
          },
        },
      );

      if (response.data.success) {
        setIsPptxUploaded(true);
        setFileUrl(`http://localhost:3001/api/pptx?v=${Date.now()}`);
      }
    } catch (error) {
      console.error("Upload error:", error);
      alert("Failed to upload file. Please try again.");
    } finally {
      setUploading(false);
    }
  };

  const runtime = useChatRuntime({
    transport: new AssistantChatTransport({
      api: "/api/chat",
    }),
  });

  return (
    <AssistantRuntimeProvider runtime={runtime}>
      {isPptxUploaded && fileUrl && (
        <ViewerRefreshHandler fileUrl={fileUrl} setFileUrl={setFileUrl} />
      )}
      <SidebarProvider>
        <div className="flex h-dvh w-full pr-0.5">
          <ThreadListSidebar />
          <SidebarInset>
            <header className="flex h-16 shrink-0 items-center gap-2 border-b px-4">
              <SidebarTrigger />
              <Separator orientation="vertical" className="mr-2 h-4" />
              <Breadcrumb>
                <BreadcrumbList>
                  <BreadcrumbItem className="hidden md:block">
                    <BreadcrumbLink
                      href="https://www.assistant-ui.com/docs/getting-started"
                      target="_blank"
                      rel="noopener noreferrer"
                    >
                      PPTX Editor
                    </BreadcrumbLink>
                  </BreadcrumbItem>
                  <BreadcrumbSeparator className="hidden md:block" />
                  <BreadcrumbItem>
                    <BreadcrumbPage>AI-Powered Editor</BreadcrumbPage>
                  </BreadcrumbItem>
                </BreadcrumbList>
              </Breadcrumb>
              <div className="ml-auto flex items-center gap-2">
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".pptx"
                  onChange={handleFileUpload}
                  className="hidden"
                  id="pptx-upload"
                />
                <Button
                  onClick={() => fileInputRef.current?.click()}
                  disabled={uploading}
                  variant="outline"
                  size="sm"
                >
                  {uploading
                    ? "Uploading..."
                    : isPptxUploaded
                      ? "Replace PPTX"
                      : "Upload PPTX"}
                </Button>
              </div>
            </header>
            <div className="flex flex-1 overflow-hidden">
              {isPptxUploaded && fileUrl ? (
                <div className="flex-1 overflow-hidden border-r">
                  <PptxViewer fileUrl={fileUrl} />
                </div>
              ) : null}
              <div className={isPptxUploaded ? "w-1/2 border-l" : "w-full"}>
                <Thread />
              </div>
            </div>
          </SidebarInset>
        </div>
      </SidebarProvider>
    </AssistantRuntimeProvider>
  );
};
