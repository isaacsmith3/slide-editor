import fs from "fs";
import JSZip from "jszip";
import { parseString } from "xml2js";

export interface SlideInfo {
  slideNumber: number;
  title: string;
  text: string[];
}

/**
 * Gets information about all slides in a PPTX file
 */
export async function getSlidesInfo(filePath: string): Promise<SlideInfo[]> {
  const slides: SlideInfo[] = [];

  try {
    // Read the PPTX file (it's a ZIP)
    const data = fs.readFileSync(filePath);
    const zip = await JSZip.loadAsync(data);

    // Get all slide files
    const slideFiles: string[] = [];
    zip.forEach((relativePath) => {
      if (
        relativePath.startsWith("ppt/slides/slide") &&
        relativePath.endsWith(".xml")
      ) {
        slideFiles.push(relativePath);
      }
    });

    // Sort slide files by number
    slideFiles.sort((a, b) => {
      const numA = parseInt(a.match(/slide(\d+)\.xml/)?.[1] || "0");
      const numB = parseInt(b.match(/slide(\d+)\.xml/)?.[1] || "0");
      return numA - numB;
    });

    // Parse each slide
    for (let i = 0; i < slideFiles.length; i++) {
      const slideFile = slideFiles[i];
      const slideXml = await zip.file(slideFile)?.async("string");

      if (!slideXml) continue;

      const slideNumber = i + 1;
      const textElements: string[] = [];

      // Parse XML to extract text
      const result = await new Promise<Record<string, unknown>>(
        (resolve, reject) => {
          parseString(slideXml, (err, parsed) => {
            if (err) reject(err);
            else resolve(parsed);
          });
        },
      );

      // Extract text from the slide
      function extractText(obj: unknown): void {
        if (typeof obj !== "object" || obj === null) {
          return;
        }

        const objRecord = obj as Record<string, unknown>;
        if (objRecord["a:t"] && Array.isArray(objRecord["a:t"])) {
          objRecord["a:t"].forEach((textObj: unknown) => {
            if (typeof textObj === "string" && textObj.trim()) {
              textElements.push(textObj.trim());
            } else if (textObj && typeof textObj === "object") {
              const textObjRecord = textObj as Record<string, unknown>;
              if (
                textObjRecord._ &&
                typeof textObjRecord._ === "string" &&
                textObjRecord._.trim()
              ) {
                textElements.push(textObjRecord._.trim());
              }
            } else if (Array.isArray(textObj)) {
              textObj.forEach((item: unknown) => {
                if (typeof item === "string" && item.trim()) {
                  textElements.push(item.trim());
                }
              });
            }
          });
        }

        for (const key in objRecord) {
          if (objRecord.hasOwnProperty(key)) {
            extractText(objRecord[key]);
          }
        }
      }

      extractText(result);

      slides.push({
        slideNumber,
        title: textElements[0] || `Slide ${slideNumber}`,
        text: textElements,
      });
    }
  } catch (error) {
    console.error("Error getting slides info:", error);
  }

  return slides;
}
