import fs from "fs";
import path from "path";
import unzipper from "unzipper";
import { parseString } from "xml2js";

export interface SlideContext {
  slide: number;
  text: string[];
}

/**
 * Parses a PPTX file and extracts text content from slides for AI context
 */
export async function parsePptxForContext(
  filePath: string,
): Promise<SlideContext[]> {
  const slides: SlideContext[] = [];
  const tempDir = path.join(path.dirname(filePath), "temp_" + Date.now());

  try {
    // Create temp directory
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }

    // Unzip the PPTX file
    await new Promise<void>((resolve, reject) => {
      fs.createReadStream(filePath)
        .pipe(unzipper.Extract({ path: tempDir }))
        .on("close", () => resolve())
        .on("error", (err) => reject(err));
    });

    // Read slide files from ppt/slides/
    const slidesDir = path.join(tempDir, "ppt", "slides");

    if (!fs.existsSync(slidesDir)) {
      // Cleanup and return empty array
      cleanupTempDir(tempDir);
      return [];
    }

    const slideFiles = fs
      .readdirSync(slidesDir)
      .filter((file) => file.startsWith("slide") && file.endsWith(".xml"))
      .sort((a, b) => {
        // Extract slide numbers for proper sorting
        const numA = parseInt(a.match(/slide(\d+)\.xml/)?.[1] || "0");
        const numB = parseInt(b.match(/slide(\d+)\.xml/)?.[1] || "0");
        return numA - numB;
      });

    // Parse each slide
    for (let i = 0; i < slideFiles.length; i++) {
      const slideFile = slideFiles[i];
      const slidePath = path.join(slidesDir, slideFile);
      const slideNumber = i + 1;

      try {
        const slideXml = fs.readFileSync(slidePath, "utf-8");
        const textElements = await extractTextFromSlideXml(slideXml);

        slides.push({
          slide: slideNumber,
          text: textElements,
        });
      } catch (err) {
        console.error(`Error parsing slide ${slideNumber}:`, err);
        // Continue with other slides even if one fails
        slides.push({
          slide: slideNumber,
          text: [],
        });
      }
    }

    // Cleanup temp directory
    cleanupTempDir(tempDir);

    return slides;
  } catch (err) {
    cleanupTempDir(tempDir);
    throw err;
  }
}

/**
 * Extracts text from slide XML by finding all <a:t> elements
 */
async function extractTextFromSlideXml(xml: string): Promise<string[]> {
  return new Promise((resolve, reject) => {
    parseString(xml, (err, result) => {
      if (err) {
        reject(err);
        return;
      }

      const textElements: string[] = [];

      // Recursively find all text elements
      function findTextElements(obj: any): void {
        if (typeof obj !== "object" || obj === null) {
          return;
        }

        // Check if this is a text element (a:t)
        if (obj["a:t"] && Array.isArray(obj["a:t"])) {
          obj["a:t"].forEach((textObj: any) => {
            if (typeof textObj === "string") {
              if (textObj.trim()) {
                textElements.push(textObj.trim());
              }
            } else if (textObj._) {
              // Sometimes text is wrapped in _ property
              if (textObj._.trim()) {
                textElements.push(textObj._.trim());
              }
            } else if (Array.isArray(textObj)) {
              textObj.forEach((item: any) => {
                if (typeof item === "string" && item.trim()) {
                  textElements.push(item.trim());
                } else if (
                  item &&
                  typeof item === "object" &&
                  item._ &&
                  item._.trim()
                ) {
                  textElements.push(item._.trim());
                }
              });
            }
          });
        }

        // Recursively search in all properties
        for (const key in obj) {
          if (obj.hasOwnProperty(key)) {
            findTextElements(obj[key]);
          }
        }
      }

      findTextElements(result);
      resolve(textElements);
    });
  });
}

/**
 * Cleanup temporary directory
 */
function cleanupTempDir(tempDir: string): void {
  try {
    if (fs.existsSync(tempDir)) {
      fs.rmSync(tempDir, { recursive: true, force: true });
    }
  } catch (err) {
    console.error("Error cleaning up temp directory:", err);
  }
}
