// Load environment variables first, before any other imports
import dotenv from "dotenv";

// Load .env file from the backend directory
// process.cwd() will be the backend directory when running npm run dev
dotenv.config();

import express from "express";
import cors from "cors";
import multer from "multer";
import path from "path";
import fs from "fs";
import { parsePptxForContext, SlideContext } from "./parser";
import { translatePromptToCommand } from "./ai";
import { applyPptxChange } from "./editor";
import { getSlidesInfo } from "./slides";
import { extractSlideImages, getSlideLayouts } from "./renderer";
import { convertSlidesToImages } from "./imageConverter";

const app = express();
const PORT = 3001;

// Store slide context in memory
let slideContext: SlideContext[] | null = null;

// Ensure uploads directory exists
const uploadsDir = path.join(__dirname, "..", "uploads");
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
}

// Multer configuration
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadsDir);
  },
  filename: (req, file, cb) => {
    // Always save as current.pptx
    cb(null, "current.pptx");
  },
});

const upload = multer({
  storage: storage,
  fileFilter: (req, file, cb) => {
    if (
      file.mimetype ===
        "application/vnd.openxmlformats-officedocument.presentationml.presentation" ||
      file.originalname.endsWith(".pptx")
    ) {
      cb(null, true);
    } else {
      cb(new Error("Only .pptx files are allowed"));
    }
  },
  limits: {
    fileSize: 50 * 1024 * 1024, // 50MB limit
  },
});

// Middleware
app.use(cors());
app.use(express.json());

// Simple route for testing
app.get("/", (req, res) => {
  res.json({ message: "PPTX Editor Backend API" });
});

// Health check endpoint
app.get("/api/health", (req, res) => {
  res.json({
    status: "ok",
    hasFile: fs.existsSync(path.join(uploadsDir, "current.pptx")),
    hasSlideContext: !!slideContext,
    slideCount: slideContext?.length || 0,
  });
});

// Upload route
app.post("/api/upload", upload.single("pptx"), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: "No file uploaded" });
  }

  try {
    // Parse the PPTX file to extract slide context
    const filePath = path.join(uploadsDir, "current.pptx");
    slideContext = await parsePptxForContext(filePath);

    console.log(`Parsed ${slideContext.length} slides`);

    // Convert slides to images in the background (non-blocking)
    const tempDir = path.join(uploadsDir, "slide_images");
    convertSlidesToImages(filePath, tempDir).catch((err) => {
      console.error("Background slide conversion failed:", err);
    });

    res.json({
      success: true,
      message: "File uploaded successfully",
      filename: req.file.filename,
      slideCount: slideContext.length,
    });
  } catch (error) {
    console.error("Error parsing PPTX:", error);
    res.status(500).json({
      error: "Failed to parse PPTX file",
      message: error instanceof Error ? error.message : "Unknown error",
    });
  }
});

// Serve PPTX file
app.get("/api/pptx", (req, res) => {
  const filePath = path.join(uploadsDir, "current.pptx");

  if (!fs.existsSync(filePath)) {
    return res
      .status(404)
      .json({ error: "No PPTX file found. Please upload a file first." });
  }

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  );
  res.setHeader("Content-Disposition", "inline; filename=current.pptx");
  res.sendFile(filePath);
});

// Get slides info for preview
app.get("/api/slides", async (req, res) => {
  try {
    const filePath = path.join(uploadsDir, "current.pptx");

    if (!fs.existsSync(filePath)) {
      return res
        .status(404)
        .json({ error: "No PPTX file found. Please upload a file first." });
    }

    const slidesInfo = await getSlidesInfo(filePath);

    // Set no-cache headers to ensure fresh slide data is always served
    res.setHeader(
      "Cache-Control",
      "no-store, no-cache, must-revalidate, proxy-revalidate",
    );
    res.setHeader("Pragma", "no-cache");
    res.setHeader("Expires", "0");

    res.json({ slides: slidesInfo });
  } catch (error) {
    console.error("Error getting slides info:", error);
    res.status(500).json({
      error: "Failed to get slides info",
      message: error instanceof Error ? error.message : "Unknown error",
    });
  }
});

// Get slide images
app.get("/api/slide-images", async (req, res) => {
  try {
    const filePath = path.join(uploadsDir, "current.pptx");

    if (!fs.existsSync(filePath)) {
      return res
        .status(404)
        .json({ error: "No PPTX file found. Please upload a file first." });
    }

    const images = await extractSlideImages(filePath);
    const imageMap: Record<string, string> = {};

    // Convert images to base64 data URLs
    images.forEach((imageData, imageName) => {
      const base64 = imageData.toString("base64");
      const ext = path.extname(imageName).toLowerCase();
      const mimeType =
        ext === ".png"
          ? "image/png"
          : ext === ".jpg" || ext === ".jpeg"
            ? "image/jpeg"
            : ext === ".gif"
              ? "image/gif"
              : "image/png";
      imageMap[imageName] = `data:${mimeType};base64,${base64}`;
    });

    res.json({ images: imageMap });
  } catch (error) {
    console.error("Error getting slide images:", error);
    res.status(500).json({
      error: "Failed to get slide images",
      message: error instanceof Error ? error.message : "Unknown error",
    });
  }
});

// Get slide layouts with visual information
app.get("/api/slide-layouts", async (req, res) => {
  try {
    const filePath = path.join(uploadsDir, "current.pptx");

    if (!fs.existsSync(filePath)) {
      return res
        .status(404)
        .json({ error: "No PPTX file found. Please upload a file first." });
    }

    const layouts = await getSlideLayouts(filePath);
    const images = await extractSlideImages(filePath);

    // Convert images to base64 and attach to layouts
    const imageMap: Record<string, string> = {};
    images.forEach((imageData, imageName) => {
      const base64 = imageData.toString("base64");
      const ext = path.extname(imageName).toLowerCase();
      const mimeType =
        ext === ".png"
          ? "image/png"
          : ext === ".jpg" || ext === ".jpeg"
            ? "image/jpeg"
            : ext === ".gif"
              ? "image/gif"
              : "image/png";
      imageMap[imageName] = `data:${mimeType};base64,${base64}`;
    });

    res.json({ layouts, images: imageMap });
  } catch (error) {
    console.error("Error getting slide layouts:", error);
    res.status(500).json({
      error: "Failed to get slide layouts",
      message: error instanceof Error ? error.message : "Unknown error",
    });
  }
});

// Get slide images (converted from PPTX)
app.get("/api/slide-images/:slideNumber", async (req, res) => {
  try {
    const filePath = path.join(uploadsDir, "current.pptx");
    const slideNumber = parseInt(req.params.slideNumber);

    if (!fs.existsSync(filePath)) {
      return res
        .status(404)
        .json({ error: "No PPTX file found. Please upload a file first." });
    }

    // Try to convert slides to images using LibreOffice
    const tempDir = path.join(uploadsDir, "slide_images");
    const slideImages = await convertSlidesToImages(filePath, tempDir);

    if (
      slideImages.length > 0 &&
      slideNumber > 0 &&
      slideNumber <= slideImages.length
    ) {
      const imagePath = slideImages[slideNumber - 1];
      if (fs.existsSync(imagePath)) {
        res.setHeader("Content-Type", "image/png");
        res.sendFile(imagePath);
        return;
      }
    }

    // Fallback: return 404
    res.status(404).json({ error: "Slide image not found" });
  } catch (error) {
    console.error("Error getting slide image:", error);
    res.status(500).json({
      error: "Failed to get slide image",
      message: error instanceof Error ? error.message : "Unknown error",
    });
  }
});

// Get all slide images as base64
app.get("/api/slide-images-all", async (req, res) => {
  try {
    const filePath = path.join(uploadsDir, "current.pptx");

    if (!fs.existsSync(filePath)) {
      return res
        .status(404)
        .json({ error: "No PPTX file found. Please upload a file first." });
    }

    // Try to convert slides to images using LibreOffice
    const tempDir = path.join(uploadsDir, "slide_images");
    let slideImages = await convertSlidesToImages(filePath, tempDir);

    // If conversion just happened, wait a moment for files to be fully written
    if (slideImages.length === 0) {
      // Wait a bit more for conversion to complete
      await new Promise((resolve) => setTimeout(resolve, 3000));

      // Check if images already exist from a previous conversion
      if (fs.existsSync(tempDir)) {
        const files = fs.readdirSync(tempDir);
        // Only get slide_*.png files (our standardized format)
        const pngFiles = files
          .filter((file) => {
            const lower = file.toLowerCase();
            return lower.endsWith(".png") && /^slide_\d+\.png$/.test(lower);
          })
          .sort((a, b) => {
            const numA = parseInt(a.match(/\d+/)?.[0] || "0");
            const numB = parseInt(b.match(/\d+/)?.[0] || "0");
            return numA - numB;
          })
          .map((file) => path.join(tempDir, file));

        if (pngFiles.length > 0) {
          console.log(`Found ${pngFiles.length} existing slide images`);
          slideImages = pngFiles;
        }
      }
    }

    console.log(`Total slide images found: ${slideImages.length}`);

    const slideImageData: Record<number, string> = {};

    if (slideImages.length > 0) {
      // Read all slide images and convert to base64
      console.log(`Reading ${slideImages.length} slide images from disk...`);
      for (let i = 0; i < slideImages.length; i++) {
        const imagePath = slideImages[i];
        if (fs.existsSync(imagePath)) {
          try {
            // Get file stats to verify it was recently modified
            const stats = fs.statSync(imagePath);
            console.log(
              `Reading slide ${i + 1} image: ${path.basename(imagePath)} (modified: ${stats.mtime.toISOString()})`,
            );

            const imageData = fs.readFileSync(imagePath);
            const base64 = imageData.toString("base64");
            // Map slide number (1-indexed) to base64 image
            slideImageData[i + 1] = `data:image/png;base64,${base64}`;
            console.log(
              `✅ Loaded slide ${i + 1} image (${Math.round(imageData.length / 1024)}KB, base64 preview: ${base64.substring(0, 30)}...)`,
            );
          } catch (err) {
            console.error(`Error reading slide image ${i + 1}:`, err);
          }
        } else {
          console.warn(`⚠️ Slide image ${i + 1} not found: ${imagePath}`);
        }
      }
    }

    console.log(
      `Returning ${Object.keys(slideImageData).length} slide images to frontend`,
    );

    // Set no-cache headers to ensure fresh images are always served
    res.setHeader(
      "Cache-Control",
      "no-store, no-cache, must-revalidate, proxy-revalidate",
    );
    res.setHeader("Pragma", "no-cache");
    res.setHeader("Expires", "0");

    res.json({
      slides: slideImageData,
      count: Object.keys(slideImageData).length,
      hasImages: Object.keys(slideImageData).length > 0,
    });
  } catch (error) {
    console.error("Error getting slide images:", error);
    res.status(500).json({
      error: "Failed to get slide images",
      message: error instanceof Error ? error.message : "Unknown error",
    });
  }
});

// Chat endpoint
app.post("/api/chat", async (req, res) => {
  try {
    console.log("Chat request received:", {
      body: req.body,
      hasSlideContext: !!slideContext,
      slideContextLength: slideContext?.length || 0,
    });

    const { prompt } = req.body;

    if (!prompt || typeof prompt !== "string") {
      console.error("Invalid prompt:", { prompt, type: typeof prompt });
      return res.status(400).json({
        error: "Prompt is required",
        received: req.body,
      });
    }

    // Check if slide context exists, if not but file exists, try to load it
    const filePath = path.join(uploadsDir, "current.pptx");
    if (
      (!slideContext || slideContext.length === 0) &&
      fs.existsSync(filePath)
    ) {
      console.log(
        "Slide context is empty but file exists. Attempting to reload...",
      );
      try {
        slideContext = await parsePptxForContext(filePath);
        console.log(
          `Reloaded ${slideContext.length} slides from existing file`,
        );
      } catch (error) {
        console.error("Error reloading slide context:", error);
        return res.status(400).json({
          error: "Failed to load PPTX file. Please re-upload the file.",
          hint: "The file exists but could not be parsed. Try uploading again.",
        });
      }
    }

    if (!slideContext || slideContext.length === 0) {
      console.error("No slide context available and no file to load from.");
      return res.status(400).json({
        error: "No PPTX file loaded. Please upload a file first.",
        hint: "Make sure you've uploaded a PPTX file using the upload button.",
      });
    }

    // Translate user prompt to command
    const command = await translatePromptToCommand(prompt, slideContext);

    console.log("Generated command:", command);

    // Apply the command to the PPTX file (filePath already defined above)
    await applyPptxChange(filePath, command);

    // Re-parse the PPTX to update the context
    slideContext = await parsePptxForContext(filePath);
    console.log(`Updated context: ${slideContext.length} slides`);

    // Re-convert slides to images and WAIT for completion
    // This prevents race conditions where frontend refreshes before images are ready
    const tempDir = path.join(uploadsDir, "slide_images");
    console.log("Starting image conversion after edit...");
    try {
      await convertSlidesToImages(filePath, tempDir);
      console.log("Image conversion completed successfully");
    } catch (err) {
      console.error("Image conversion after edit failed:", err);
      // Don't fail the request, but log the error
      // Images might still be available from previous conversion
    }

    res.json({
      success: true,
      command: command,
      message: "Changes applied successfully.",
    });
  } catch (error) {
    console.error("Error in /api/chat:", error);
    res.status(500).json({
      error: "Failed to process chat request",
      message: error instanceof Error ? error.message : "Unknown error",
    });
  }
});

app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
  if (!process.env.OPENAI_API_KEY) {
    console.warn("Warning: OPENAI_API_KEY environment variable is not set");
  }
});
