import OpenAI from "openai";
import { SlideContext } from "./parser";

// Lazy initialization to ensure environment variables are loaded
let openaiClient: OpenAI | null = null;

function getOpenAIClient(): OpenAI {
  if (!openaiClient) {
    const apiKey = process.env.OPENAI_API_KEY;
    if (!apiKey) {
      throw new Error(
        "OPENAI_API_KEY environment variable is not set. Please create a .env file in the backend directory with your OpenAI API key.",
      );
    }
    openaiClient = new OpenAI({
      apiKey: apiKey,
    });
  }
  return openaiClient;
}

export interface Command {
  action: "update_text" | "change_bg";
  slide: number;
  oldText?: string;
  newText?: string;
  color?: string;
}

/**
 * Translates a user prompt into a structured command using OpenAI
 */
export async function translatePromptToCommand(
  userPrompt: string,
  slideContext: SlideContext[] | null,
): Promise<Command> {
  const systemPrompt = `You are a PPTX editing assistant. The user will provide a command to edit a presentation.
The current presentation structure is:
${JSON.stringify(slideContext || [], null, 2)}

Based on the user's command, return a single JSON object representing the action to take.
Valid actions are:
- {"action": "update_text", "slide": <number>, "oldText": "<string>", "newText": "<string>"}
- {"action": "change_bg", "slide": <number>, "color": "<hex_color>"}

Only return the JSON object, no additional text or explanation.`;

  try {
    console.log("Translating prompt to command:", {
      prompt: userPrompt,
      slideCount: slideContext?.length || 0,
    });

    const openai = getOpenAIClient();
    const response = await openai.chat.completions.create({
      model: "gpt-4o-mini",
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: userPrompt },
      ],
      temperature: 0.3,
      response_format: { type: "json_object" },
    });

    const content = response.choices[0]?.message?.content;
    if (!content) {
      throw new Error("No response from OpenAI");
    }

    console.log("OpenAI response:", content);

    // Parse the JSON response
    let command: Command;
    try {
      command = JSON.parse(content) as Command;
    } catch (parseError) {
      console.error("Failed to parse OpenAI response as JSON:", content);
      throw new Error(
        `Invalid JSON response from AI: ${parseError instanceof Error ? parseError.message : "Unknown error"}`,
      );
    }

    // Validate command structure
    if (!command.action || !command.slide) {
      console.error("Invalid command structure:", command);
      throw new Error(
        `Invalid command structure from AI. Missing action or slide. Received: ${JSON.stringify(command)}`,
      );
    }

    // Validate action type
    if (command.action !== "update_text" && command.action !== "change_bg") {
      throw new Error(
        `Invalid action type: ${command.action}. Must be "update_text" or "change_bg"`,
      );
    }

    // Validate slide number
    if (command.slide < 1) {
      throw new Error(
        `Invalid slide number: ${command.slide}. Slide numbers must be >= 1`,
      );
    }

    if (slideContext && command.slide > slideContext.length) {
      throw new Error(
        `Slide ${command.slide} does not exist. Presentation only has ${slideContext.length} slides.`,
      );
    }

    console.log("Successfully parsed command:", command);
    return command;
  } catch (error) {
    console.error("Error translating prompt to command:", error);
    if (error instanceof Error) {
      // Re-throw with more context
      throw new Error(`Failed to translate command: ${error.message}`);
    }
    throw error;
  }
}
