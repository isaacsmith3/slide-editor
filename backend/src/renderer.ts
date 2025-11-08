import fs from "fs";
import path from "path";
import JSZip from "jszip";
import { parseString } from "xml2js";

export interface SlideImage {
  slideNumber: number;
  imagePath: string;
  imageData: Buffer;
}

export interface SlideLayout {
  slideNumber: number;
  width: number;
  height: number;
  background?: string;
  shapes: Array<{
    type: string;
    x?: number;
    y?: number;
    width?: number;
    height?: number;
    text?: string;
    fill?: string;
    imageRef?: string;
  }>;
}

/**
 * Extracts images from a PPTX file
 */
export async function extractSlideImages(
  filePath: string,
): Promise<Map<string, Buffer>> {
  const images = new Map<string, Buffer>();

  try {
    const data = fs.readFileSync(filePath);
    const zip = await JSZip.loadAsync(data);

    // Extract all images from the media folder
    zip.forEach(async (relativePath, file) => {
      if (
        relativePath.startsWith("ppt/media/") &&
        (relativePath.endsWith(".png") ||
          relativePath.endsWith(".jpg") ||
          relativePath.endsWith(".jpeg") ||
          relativePath.endsWith(".gif"))
      ) {
        const imageData = await file.async("nodebuffer");
        const imageName = path.basename(relativePath);
        images.set(imageName, imageData);
      }
    });

    // Also check for images in slide folders
    zip.forEach(async (relativePath, file) => {
      if (
        relativePath.includes("/media/") &&
        (relativePath.endsWith(".png") ||
          relativePath.endsWith(".jpg") ||
          relativePath.endsWith(".jpeg"))
      ) {
        const imageData = await file.async("nodebuffer");
        const imageName = path.basename(relativePath);
        if (!images.has(imageName)) {
          images.set(imageName, imageData);
        }
      }
    });
  } catch (error) {
    console.error("Error extracting images:", error);
  }

  return images;
}

/**
 * Gets slide layout information including images and shapes
 */
export async function getSlideLayouts(
  filePath: string,
): Promise<SlideLayout[]> {
  const layouts: SlideLayout[] = [];
  const images = await extractSlideImages(filePath);

  try {
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

    // Sort slide files
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
      const layout = await parseSlideLayout(slideXml, slideNumber, images);
      layouts.push(layout);
    }
  } catch (error) {
    console.error("Error getting slide layouts:", error);
  }

  return layouts;
}

/**
 * Parses a slide XML to extract layout information
 */
async function parseSlideLayout(
  slideXml: string,
  slideNumber: number,
  images: Map<string, Buffer>,
): Promise<SlideLayout> {
  const layout: SlideLayout = {
    slideNumber,
    width: 9144000, // Default PowerPoint width in EMUs
    height: 6858000, // Default PowerPoint height in EMUs
    shapes: [],
  };

  try {
    const result = await new Promise<Record<string, unknown>>(
      (resolve, reject) => {
        parseString(slideXml, (err, parsed) => {
          if (err) reject(err);
          else resolve(parsed);
        });
      },
    );

    // Extract dimensions from p:sldSz if available
    const slide = result["p:sld"]?.[0] as Record<string, unknown> | undefined;
    if (slide?.["p:cSld"]?.[0]) {
      const cSld = (slide["p:cSld"] as unknown[])[0] as Record<string, unknown>;

      // Extract background
      if (cSld["p:bg"]) {
        const bg = (cSld["p:bg"] as unknown[])[0] as Record<string, unknown>;
        if (bg["a:bgPr"]?.[0]) {
          const bgPr = (bg["a:bgPr"] as unknown[])[0] as Record<
            string,
            unknown
          >;
          if (bgPr["a:solidFill"]?.[0]) {
            const solidFill = (bgPr["a:solidFill"] as unknown[])[0] as Record<
              string,
              unknown
            >;
            if (solidFill["a:srgbClr"]?.[0]) {
              const srgbClr = (
                solidFill["a:srgbClr"] as unknown[]
              )[0] as Record<string, unknown>;
              const attrs = srgbClr["$"] as Record<string, string> | undefined;
              if (attrs?.val) {
                layout.background = `#${attrs.val}`;
              }
            }
          }
        }
      }

      // Extract shapes
      if (cSld["p:spTree"]?.[0]) {
        const spTree = (cSld["p:spTree"] as unknown[])[0] as Record<
          string,
          unknown
        >;
        extractShapes(spTree, layout, images);
      }
    }
  } catch (error) {
    console.error(`Error parsing slide ${slideNumber} layout:`, error);
  }

  return layout;
}

/**
 * Recursively extracts shapes from the XML tree
 */
function extractShapes(
  node: Record<string, unknown>,
  layout: SlideLayout,
  images: Map<string, Buffer>,
): void {
  // Extract text shapes
  if (node["p:sp"]) {
    const shapes = Array.isArray(node["p:sp"]) ? node["p:sp"] : [node["p:sp"]];
    shapes.forEach((shape: unknown) => {
      const shapeRecord = shape as Record<string, unknown>;
      const shapeInfo: SlideLayout["shapes"][0] = {
        type: "text",
      };

      // Extract position and size
      if (shapeRecord["p:spPr"]) {
        const spPrArray = shapeRecord["p:spPr"] as unknown[];
        const spPr = spPrArray[0] as Record<string, unknown>;
        if (spPr["a:xfrm"]) {
          const xfrmArray = spPr["a:xfrm"] as unknown[];
          const xfrm = xfrmArray[0] as Record<string, unknown>;
          if (xfrm["a:off"]) {
            const offArray = xfrm["a:off"] as unknown[];
            const off = offArray[0] as Record<string, unknown>;
            const attrs = off["$"] as Record<string, string> | undefined;
            if (attrs) {
              shapeInfo.x = parseInt(attrs.x || "0");
              shapeInfo.y = parseInt(attrs.y || "0");
            }
          }
          if (xfrm["a:ext"]) {
            const extArray = xfrm["a:ext"] as unknown[];
            const ext = extArray[0] as Record<string, unknown>;
            const attrs = ext["$"] as Record<string, string> | undefined;
            if (attrs) {
              shapeInfo.width = parseInt(attrs.cx || "0");
              shapeInfo.height = parseInt(attrs.cy || "0");
            }
          }
        }
      }

      // Extract text
      if (shapeRecord["p:txBody"]) {
        const txBodyArray = shapeRecord["p:txBody"] as unknown[];
        const txBody = txBodyArray[0] as Record<string, unknown>;
        const text = extractTextFromBody(txBody);
        if (text) {
          shapeInfo.text = text;
        }
      }

      layout.shapes.push(shapeInfo);
    });
  }

  // Extract picture shapes
  if (node["p:pic"]) {
    const pics = Array.isArray(node["p:pic"]) ? node["p:pic"] : [node["p:pic"]];
    pics.forEach((pic: unknown) => {
      const picRecord = pic as Record<string, unknown>;
      const shapeInfo: SlideLayout["shapes"][0] = {
        type: "image",
      };

      // Extract image reference
      if (picRecord["p:blipFill"]?.[0]?.["a:blip"]?.[0]?.["$"]?.["r:embed"]) {
        const embedId = (
          (picRecord["p:blipFill"] as unknown[])[0] as Record<string, unknown>
        )?.["a:blip"]?.[0] as Record<string, unknown> | undefined;
        const attrs = embedId?.["$"] as Record<string, string> | undefined;
        if (attrs?.["r:embed"]) {
          shapeInfo.imageRef = attrs["r:embed"];
        }
      }

      // Extract position
      if (picRecord["p:spPr"]?.[0]) {
        const spPr = (picRecord["p:spPr"] as unknown[])[0] as Record<
          string,
          unknown
        >;
        if (spPr["a:xfrm"]?.[0]) {
          const xfrm = (spPr["a:xfrm"] as unknown[])[0] as Record<
            string,
            unknown
          >;
          if (xfrm["a:off"]?.[0]) {
            const off = (xfrm["a:off"] as unknown[])[0] as Record<
              string,
              unknown
            >;
            const attrs = off["$"] as Record<string, string> | undefined;
            if (attrs) {
              shapeInfo.x = parseInt(attrs.x || "0");
              shapeInfo.y = parseInt(attrs.y || "0");
            }
          }
          if (xfrm["a:ext"]?.[0]) {
            const ext = (xfrm["a:ext"] as unknown[])[0] as Record<
              string,
              unknown
            >;
            const attrs = ext["$"] as Record<string, string> | undefined;
            if (attrs) {
              shapeInfo.width = parseInt(attrs.cx || "0");
              shapeInfo.height = parseInt(attrs.cy || "0");
            }
          }
        }
      }

      layout.shapes.push(shapeInfo);
    });
  }

  // Recursively process nested shapes
  if (node["p:spTree"]) {
    const spTrees = Array.isArray(node["p:spTree"])
      ? node["p:spTree"]
      : [node["p:spTree"]];
    spTrees.forEach((tree: unknown) => {
      extractShapes(tree as Record<string, unknown>, layout, images);
    });
  }
}

/**
 * Extracts text from a text body element
 */
function extractTextFromBody(txBody: Record<string, unknown>): string {
  const textParts: string[] = [];

  function extractFromNode(node: unknown): void {
    if (typeof node === "string") {
      if (node.trim()) {
        textParts.push(node.trim());
      }
    } else if (node && typeof node === "object") {
      const nodeRecord = node as Record<string, unknown>;
      if (nodeRecord["a:t"]) {
        const texts = Array.isArray(nodeRecord["a:t"])
          ? nodeRecord["a:t"]
          : [nodeRecord["a:t"]];
        texts.forEach((text: unknown) => {
          if (typeof text === "string") {
            if (text.trim()) {
              textParts.push(text.trim());
            }
          } else if (text && typeof text === "object") {
            const textRecord = text as Record<string, unknown>;
            if (textRecord._ && typeof textRecord._ === "string") {
              if (textRecord._.trim()) {
                textParts.push(textRecord._.trim());
              }
            }
          }
        });
      }

      // Recursively process children
      for (const key in nodeRecord) {
        if (nodeRecord.hasOwnProperty(key)) {
          extractFromNode(nodeRecord[key]);
        }
      }
    }
  }

  extractFromNode(txBody);
  return textParts.join(" ");
}
