import { UIMessage } from "ai";
import { openai } from "@ai-sdk/openai";
import { streamText, convertToModelMessages } from "ai";

export async function POST(req: Request) {
  // Log immediately when route is hit
  console.log("🚀 POST /api/chat called");

  // Declare variables outside try block so they're available in catch
  let messages: UIMessage[] | undefined;
  let body: any;

  try {
    // Log the raw request
    console.log("=== Chat API Route Called ===");
    console.log("Request URL:", req.url);
    console.log("Request method:", req.method);

    // Get request headers
    const headers: Record<string, string> = {};
    req.headers.forEach((value, key) => {
      headers[key] = value;
    });
    console.log("Request headers:", headers);

    // Check content type
    const contentType = req.headers.get("content-type");
    console.log("Content-Type:", contentType);

    if (!contentType || !contentType.includes("application/json")) {
      console.warn("⚠️ Content-Type is not application/json:", contentType);
      // Still try to parse, but log the issue
    }

    try {
      // Read the request body as text first to log it, then parse as JSON
      const bodyText = await req.text();

      if (!bodyText || bodyText.trim() === "") {
        console.error("❌ Empty request body");
        return new Response(JSON.stringify({ error: "Empty request body" }), {
          status: 400,
          headers: { "Content-Type": "application/json" },
        });
      }

      console.log(
        "Raw request body (first 1000 chars):",
        bodyText.substring(0, 1000),
      );

      try {
        body = JSON.parse(bodyText);
        console.log("✅ Successfully parsed request body");
        console.log("Body keys:", Object.keys(body));
        console.log("Full body:", JSON.stringify(body, null, 2));
      } catch (jsonParseError) {
        console.error("❌ Failed to parse body as JSON:", jsonParseError);
        console.error("Body text that failed to parse:", bodyText);
        return new Response(
          JSON.stringify({
            error: "Invalid JSON in request body",
            details:
              jsonParseError instanceof Error
                ? jsonParseError.message
                : "Unknown error",
          }),
          { status: 400, headers: { "Content-Type": "application/json" } },
        );
      }
    } catch (readError) {
      console.error("❌ Failed to read request body:", readError);
      return new Response(
        JSON.stringify({
          error: "Failed to read request body",
          details:
            readError instanceof Error ? readError.message : "Unknown error",
        }),
        { status: 400, headers: { "Content-Type": "application/json" } },
      );
    }

    // Extract messages from body
    messages = body.messages;

    console.log("Chat API route called:", {
      messagesLength: messages?.length,
      lastMessage: messages?.[messages.length - 1],
      bodyKeys: Object.keys(body),
    });

    if (!messages || !Array.isArray(messages) || messages.length === 0) {
      console.error("No messages in request:", {
        messages,
        type: typeof messages,
      });
      return new Response(
        JSON.stringify({
          error: "No messages provided",
          received: body,
        }),
        { status: 400, headers: { "Content-Type": "application/json" } },
      );
    }

    // Extract the last user message
    const lastMessage = messages[messages.length - 1];
    console.log("Last message details:", {
      role: lastMessage?.role,
      hasContent: !!lastMessage?.content,
      hasParts: !!lastMessage?.parts,
      contentType: typeof lastMessage?.content,
      content: lastMessage?.content,
      parts: lastMessage?.parts,
      fullMessage: lastMessage,
    });

    if (!lastMessage) {
      console.error("No last message found");
      return new Response(
        JSON.stringify({
          error: "No last message found",
          messagesLength: messages.length,
        }),
        { status: 400, headers: { "Content-Type": "application/json" } },
      );
    }

    if (lastMessage.role !== "user") {
      console.error("Last message is not from user:", {
        role: lastMessage.role,
        message: lastMessage,
      });
      return new Response(
        JSON.stringify({
          error: "Last message must be from user",
          role: lastMessage.role,
        }),
        { status: 400, headers: { "Content-Type": "application/json" } },
      );
    }

    // Extract text content from message
    // assistant-ui uses 'parts' array, AI SDK uses 'content'
    let userPrompt = "";

    // Try content first (AI SDK format)
    if (typeof lastMessage.content === "string") {
      userPrompt = lastMessage.content;
    } else if (Array.isArray(lastMessage.content)) {
      // Find text parts in the content array
      const textPart = lastMessage.content.find((part: any) => {
        return part.type === "text" || typeof part === "string";
      });
      if (textPart) {
        userPrompt =
          typeof textPart === "string"
            ? textPart
            : textPart.text || textPart.content || "";
      }
    } else if (lastMessage.content && typeof lastMessage.content === "object") {
      // Try to extract text from object
      userPrompt =
        (lastMessage.content as any).text ||
        (lastMessage.content as any).content ||
        String(lastMessage.content);
    }

    // If no content found, try parts (assistant-ui format)
    if (!userPrompt && lastMessage.parts && Array.isArray(lastMessage.parts)) {
      console.log(
        "Extracting from parts array. Full parts:",
        JSON.stringify(lastMessage.parts, null, 2),
      );

      for (let i = 0; i < lastMessage.parts.length; i++) {
        const part = lastMessage.parts[i];
        console.log(`Processing part ${i}:`, {
          part,
          type: typeof part,
          isArray: Array.isArray(part),
          keys:
            part && typeof part === "object" && !Array.isArray(part)
              ? Object.keys(part)
              : [],
        });

        if (typeof part === "string") {
          userPrompt = part;
          console.log("✅ Found string part:", userPrompt);
          break;
        } else if (part && typeof part === "object" && !Array.isArray(part)) {
          // Try all possible text properties
          const textProperties = [
            "text",
            "content",
            "value",
            "message",
            "data",
          ];
          for (const prop of textProperties) {
            if (part[prop] && typeof part[prop] === "string") {
              userPrompt = part[prop];
              console.log(`✅ Found part.${prop}:`, userPrompt);
              break;
            }
          }

          if (userPrompt) break;

          // If part has type "text", try to extract from it
          if (part.type === "text" || part.type === "message") {
            // Try all properties of the part object
            for (const key in part) {
              if (
                key !== "type" &&
                typeof part[key] === "string" &&
                part[key].trim()
              ) {
                userPrompt = part[key];
                console.log(
                  `✅ Found part.${key} (type=${part.type}):`,
                  userPrompt,
                );
                break;
              }
            }
          }

          if (userPrompt) break;
        } else if (Array.isArray(part)) {
          // If part is an array, recursively search
          for (const subPart of part) {
            if (typeof subPart === "string") {
              userPrompt = subPart;
              console.log("✅ Found string in nested array:", userPrompt);
              break;
            } else if (subPart && typeof subPart === "object" && subPart.text) {
              userPrompt = subPart.text;
              console.log("✅ Found nested part.text:", userPrompt);
              break;
            }
          }
          if (userPrompt) break;
        }
      }
    }

    // Last resort: try to stringify the whole message
    if (!userPrompt) {
      console.log(
        "No text found in content or parts, trying to extract from message object",
      );
      const messageStr = JSON.stringify(lastMessage);
      // This is a fallback - shouldn't normally be needed
      userPrompt = messageStr;
    }

    console.log("Extracted user prompt:", {
      prompt: userPrompt,
      promptLength: userPrompt.length,
      isEmpty: !userPrompt || userPrompt.trim() === "",
    });

    if (!userPrompt || userPrompt.trim() === "") {
      console.error("No text content in message:", {
        content: lastMessage.content,
        contentType: typeof lastMessage.content,
        contentIsArray: Array.isArray(lastMessage.content),
      });
      return new Response(
        JSON.stringify({
          error: "No text content in message",
          content: lastMessage.content,
          contentType: typeof lastMessage.content,
        }),
        { status: 400, headers: { "Content-Type": "application/json" } },
      );
    }

    // Call the backend API
    console.log("Calling backend API with prompt:", userPrompt);
    const backendResponse = await fetch("http://localhost:3001/api/chat", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ prompt: userPrompt }),
    });

    console.log("Backend response status:", backendResponse.status);

    if (!backendResponse.ok) {
      const error = await backendResponse.json().catch(() => ({
        error: `Backend error: ${backendResponse.status} ${backendResponse.statusText}`,
      }));
      console.error("Backend error:", error);
      throw new Error(
        error.error ||
          error.message ||
          `Backend error: ${backendResponse.status}`,
      );
    }

    const data = await backendResponse.json();

    // Format the response text
    const responseText =
      data.message ||
      `Command executed: ${JSON.stringify(data.command, null, 2)}`;

    // Create a streaming response using AI SDK that returns our custom text
    // We'll use streamText but with a system message that forces our response
    const result = streamText({
      model: openai("gpt-4o-mini"),
      system: `You are a PPTX editor assistant. Respond only with the following message exactly as provided: ${responseText}`,
      messages: convertToModelMessages(messages),
      maxTokens: 500,
    });

    // Return the streaming response using toUIMessageStreamResponse() for assistant-ui compatibility
    return result.toUIMessageStreamResponse();
  } catch (error) {
    console.error("Error in chat route:", error);
    console.error(
      "Error stack:",
      error instanceof Error ? error.stack : "No stack",
    );
    console.error("Error details:", {
      name: error instanceof Error ? error.name : "Unknown",
      message: error instanceof Error ? error.message : String(error),
    });

    const errorMessage =
      error instanceof Error ? error.message : "Unknown error";

    // If we don't have messages, return a simple error response
    if (!messages || messages.length === 0) {
      return new Response(
        JSON.stringify({
          error: errorMessage,
          type: "chat_error",
        }),
        {
          status: 500,
          headers: { "Content-Type": "application/json" },
        },
      );
    }

    // Return error as a streaming response
    try {
      const result = streamText({
        model: openai("gpt-4o-mini"),
        system: `You are a PPTX editor assistant. Respond with a clear error message: ${errorMessage}. If the error mentions "No PPTX file loaded", tell the user to upload a PPTX file first using the upload button.`,
        messages: convertToModelMessages(messages || []),
        maxTokens: 200,
      });

      return result.toUIMessageStreamResponse();
    } catch (streamError) {
      console.error("Error creating stream response:", streamError);
      return new Response(
        JSON.stringify({
          error: errorMessage,
          type: "chat_error",
          streamError:
            streamError instanceof Error
              ? streamError.message
              : String(streamError),
        }),
        {
          status: 500,
          headers: { "Content-Type": "application/json" },
        },
      );
    }
  }
}
