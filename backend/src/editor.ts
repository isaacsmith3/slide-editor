import fs from "fs";
import path from "path";
import JSZip from "jszip";
import { parseString, Builder } from "xml2js";
import { Command } from "./ai";

/**
 * Applies a command to modify a PPTX file
 */
export async function applyPptxChange(
  filePath: string,
  command: Command,
): Promise<void> {
  try {
    // Read the PPTX file (it's a ZIP)
    const data = fs.readFileSync(filePath);
    const zip = await JSZip.loadAsync(data);

    switch (command.action) {
      case "update_text":
        await updateTextInSlide(
          zip,
          command.slide,
          command.oldText || "",
          command.newText || "",
        );
        break;
      case "change_bg":
        await changeSlideBackground(
          zip,
          command.slide,
          command.color || "FFFFFF",
        );
        break;
      default:
        throw new Error(`Unknown action: ${(command as any).action}`);
    }

    // Save the modified ZIP back to file
    const modifiedData = await zip.generateAsync({ type: "nodebuffer" });
    fs.writeFileSync(filePath, modifiedData);
  } catch (error) {
    console.error("Error applying PPTX change:", error);
    throw error;
  }
}

/**
 * Updates text in a specific slide
 */
async function updateTextInSlide(
  zip: JSZip,
  slideNumber: number,
  oldText: string,
  newText: string,
): Promise<void> {
  const slidePath = `ppt/slides/slide${slideNumber}.xml`;

  const slideFile = zip.file(slidePath);
  if (!slideFile) {
    throw new Error(`Slide ${slideNumber} not found`);
  }

  const slideXml = await slideFile.async("string");

  // Parse XML
  const result = await new Promise<any>((resolve, reject) => {
    parseString(slideXml, (err, parsed) => {
      if (err) reject(err);
      else resolve(parsed);
    });
  });

  // Recursively find and replace text in <a:t> elements
  function replaceTextInObject(obj: any): boolean {
    if (typeof obj !== "object" || obj === null) {
      return false;
    }

    let replaced = false;

    // Check if this is a text element
    if (obj["a:t"]) {
      if (Array.isArray(obj["a:t"])) {
        obj["a:t"].forEach((textItem: any, index: number) => {
          if (typeof textItem === "string") {
            if (textItem.includes(oldText)) {
              obj["a:t"][index] = textItem.replace(
                new RegExp(oldText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g"),
                newText,
              );
              replaced = true;
            }
          } else if (textItem && typeof textItem === "object") {
            // Text might be in _ property or as an object
            if (textItem._ && typeof textItem._ === "string") {
              if (textItem._.includes(oldText)) {
                obj["a:t"][index]._ = textItem._.replace(
                  new RegExp(
                    oldText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
                    "g",
                  ),
                  newText,
                );
                replaced = true;
              }
            } else if (Array.isArray(textItem)) {
              // Handle array case
              textItem.forEach((item: any, itemIndex: number) => {
                if (typeof item === "string" && item.includes(oldText)) {
                  obj["a:t"][index][itemIndex] = item.replace(
                    new RegExp(
                      oldText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
                      "g",
                    ),
                    newText,
                  );
                  replaced = true;
                }
              });
            }
          }
        });
      } else if (typeof obj["a:t"] === "string") {
        if (obj["a:t"].includes(oldText)) {
          obj["a:t"] = obj["a:t"].replace(
            new RegExp(oldText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g"),
            newText,
          );
          replaced = true;
        }
      }
    }

    // Recursively search in all properties
    for (const key in obj) {
      if (obj.hasOwnProperty(key)) {
        if (replaceTextInObject(obj[key])) {
          replaced = true;
        }
      }
    }

    return replaced;
  }

  const textReplaced = replaceTextInObject(result);

  if (!textReplaced && oldText) {
    console.warn(`Text "${oldText}" not found in slide ${slideNumber}`);
  }

  // Convert back to XML
  const builder = new Builder({
    xmldec: { version: "1.0", encoding: "UTF-8", standalone: true },
  });
  const modifiedXml = builder.buildObject(result);

  // Update the file in the ZIP
  zip.file(slidePath, modifiedXml);
}

/**
 * Changes the background color of a slide
 */
async function changeSlideBackground(
  zip: JSZip,
  slideNumber: number,
  color: string,
): Promise<void> {
  const slidePath = `ppt/slides/slide${slideNumber}.xml`;

  const slideFile = zip.file(slidePath);
  if (!slideFile) {
    throw new Error(`Slide ${slideNumber} not found`);
  }

  const slideXml = await slideFile.async("string");

  // Parse XML
  const result = await new Promise<any>((resolve, reject) => {
    parseString(slideXml, (err, parsed) => {
      if (err) reject(err);
      else resolve(parsed);
    });
  });

  // Find or create the p:cSld element
  if (!result["p:sld"]) {
    throw new Error("Invalid slide structure");
  }

  const slide = result["p:sld"][0];

  // Find or create cSld
  if (!slide["p:cSld"]) {
    slide["p:cSld"] = [{}];
  }

  const cSld = slide["p:cSld"][0];

  // Find or create bg
  if (!cSld["p:bg"]) {
    cSld["p:bg"] = [{}];
  }

  const bg = cSld["p:bg"][0];

  // Find or create bgPr
  if (!bg["a:bgPr"]) {
    bg["a:bgPr"] = [{}];
  }

  const bgPr = bg["a:bgPr"][0];

  // Find or create solidFill
  if (!bgPr["a:solidFill"]) {
    bgPr["a:solidFill"] = [{}];
  }

  const solidFill = bgPr["a:solidFill"][0];

  // Set the color (remove # if present, ensure it's hex)
  const hexColor = color.replace("#", "").toUpperCase();

  // Find or create srgbClr
  if (!solidFill["a:srgbClr"]) {
    solidFill["a:srgbClr"] = [{}];
  }

  const srgbClr = solidFill["a:srgbClr"][0];
  srgbClr["$"] = { val: hexColor };

  // Convert back to XML
  const builder = new Builder({
    xmldec: { version: "1.0", encoding: "UTF-8", standalone: true },
  });
  const modifiedXml = builder.buildObject(result);

  // Update the file in the ZIP
  zip.file(slidePath, modifiedXml);
}
