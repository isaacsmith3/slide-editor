import { exec } from "child_process";
import { promisify } from "util";
import fs from "fs";
import path from "path";

const execAsync = promisify(exec);

// Helper to check if a command exists
async function commandExists(command: string): Promise<boolean> {
  try {
    if (process.platform === "win32") {
      await execAsync(`where ${command}`, { timeout: 5000 });
    } else {
      await execAsync(`which ${command}`, { timeout: 5000 });
    }
    return true;
  } catch {
    // Check if it's a direct path that exists
    if (command.includes("/") || command.includes("\\")) {
      return fs.existsSync(command);
    }
    return false;
  }
}

/**
 * Converts PPTX slides to images using LibreOffice (if available)
 * Uses LibreOffice Impress to export each slide as a separate PNG image
 */
export async function convertSlidesToImages(
  filePath: string,
  outputDir: string,
): Promise<string[]> {
  // Create output directory
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  // DON'T delete old images yet - we'll do it atomically after new ones are ready
  // This prevents race conditions where frontend tries to load images during conversion
  const tempDir = path.join(outputDir, ".temp_conversion");
  if (fs.existsSync(tempDir)) {
    // Clean up any previous temp conversion
    try {
      fs.rmSync(tempDir, { recursive: true, force: true });
    } catch {}
  }
  fs.mkdirSync(tempDir, { recursive: true });

  // Try different LibreOffice commands (varies by OS and installation)
  const libreOfficePaths = [
    "libreoffice",
    "soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/libreoffice",
    "/usr/local/bin/libreoffice",
  ];

  for (const libreOfficePath of libreOfficePaths) {
    try {
      // Check if LibreOffice exists
      if (!(await commandExists(libreOfficePath))) {
        continue;
      }

      console.log(`Attempting conversion with: ${libreOfficePath}`);

      // Strategy: Convert PPTX to PDF first (guarantees all slides), then PDF to images
      // This is more reliable than direct PNG conversion which might only export first slide

      const inputBasename = path.basename(filePath, path.extname(filePath));
      const pdfFile = path.join(tempDir, `${inputBasename}.pdf`);

      // Step 1: Convert PPTX to PDF (includes all slides)
      console.log(`Step 1: Converting PPTX to PDF (all slides)...`);
      console.log(`Input file: ${filePath}`);
      console.log(`Temp directory: ${tempDir}`);
      console.log(`Expected PDF: ${pdfFile}`);

      const pdfCommand = `"${libreOfficePath}" --headless --invisible --nodefault --convert-to pdf --outdir "${tempDir}" "${filePath}"`;
      console.log(`Running command: ${pdfCommand}`);

      try {
        const result = await execAsync(pdfCommand, {
          timeout: 60000,
          maxBuffer: 10 * 1024 * 1024,
        });

        if (result.stdout) {
          console.log(`LibreOffice stdout: ${result.stdout}`);
        }
        if (result.stderr) {
          console.log(`LibreOffice stderr: ${result.stderr}`);
        }

        // Wait for PDF to be created and check multiple times
        let pdfCreated = false;
        for (let attempt = 0; attempt < 10; attempt++) {
          await new Promise((resolve) => setTimeout(resolve, 1000));

          // Check if PDF exists
          if (fs.existsSync(pdfFile)) {
            console.log(`PDF created successfully: ${pdfFile}`);
            pdfCreated = true;
            break;
          }

          // Also check for PDF files in temp directory (LibreOffice might name it differently)
          if (fs.existsSync(tempDir)) {
            const files = fs.readdirSync(tempDir);
            const pdfFiles = files.filter((f) =>
              f.toLowerCase().endsWith(".pdf"),
            );
            if (pdfFiles.length > 0) {
              const actualPdf = path.join(tempDir, pdfFiles[0]);
              console.log(
                `Found PDF file: ${actualPdf} (may be different name)`,
              );
              // Update pdfFile to the actual file
              if (actualPdf !== pdfFile) {
                try {
                  fs.renameSync(actualPdf, pdfFile);
                  console.log(`Renamed PDF to expected name`);
                } catch {
                  console.log(
                    `Could not rename, using actual file: ${actualPdf}`,
                  );
                }
              }
              pdfCreated = true;
              break;
            }
          }
        }

        if (!pdfCreated) {
          console.log(
            `PDF file not found after conversion attempt. Checking temp directory...`,
          );
          if (fs.existsSync(tempDir)) {
            const files = fs.readdirSync(tempDir);
            console.log(`Files in temp directory: ${files.join(", ")}`);
          }
          console.log(`Trying direct PNG conversion as fallback...`);
          // Try direct PNG conversion as fallback
          const pngCommand = `"${libreOfficePath}" --headless --invisible --nodefault --convert-to png --outdir "${tempDir}" "${filePath}"`;
          console.log(`Running PNG command: ${pngCommand}`);
          try {
            const pngResult = await execAsync(pngCommand, { timeout: 60000 });
            if (pngResult.stdout)
              console.log(`PNG conversion stdout: ${pngResult.stdout}`);
            if (pngResult.stderr)
              console.log(`PNG conversion stderr: ${pngResult.stderr}`);
            await new Promise((resolve) => setTimeout(resolve, 3000));
          } catch (pngErr) {
            console.error(`PNG conversion error:`, pngErr);
          }
        }
      } catch (err) {
        console.error(`PDF conversion error:`, err);
        console.log(`Trying direct PNG conversion...`);
        // Fallback to direct PNG
        const pngCommand = `"${libreOfficePath}" --headless --invisible --nodefault --convert-to png --outdir "${tempDir}" "${filePath}"`;
        try {
          const pngResult = await execAsync(pngCommand, { timeout: 60000 });
          if (pngResult.stdout)
            console.log(`PNG conversion stdout: ${pngResult.stdout}`);
          if (pngResult.stderr)
            console.log(`PNG conversion stderr: ${pngResult.stderr}`);
          await new Promise((resolve) => setTimeout(resolve, 3000));
        } catch (pngErr) {
          console.error(`PNG conversion also failed:`, pngErr);
        }
      }

      // Step 2: Convert PDF to PNG images using pdftoppm (best quality, guaranteed all pages)
      if (fs.existsSync(pdfFile)) {
        console.log(`Step 2: Converting PDF to PNG images...`);

        // Try pdftoppm first (from poppler-utils - best tool for this)
        const pdftoppmCommand = `pdftoppm -png -r 200 "${pdfFile}" "${path.join(tempDir, "slide")}"`;
        try {
          await execAsync(pdftoppmCommand, { timeout: 90000 });
          await new Promise((resolve) => setTimeout(resolve, 2000));

          // pdftoppm creates slide-01.png, slide-02.png, etc.
          const allFiles = fs.readdirSync(tempDir);
          const slideFiles = allFiles
            .filter((f) => {
              const lower = f.toLowerCase();
              return lower.endsWith(".png") && /slide-\d+\.png/.test(lower);
            })
            .sort((a, b) => {
              const numA = parseInt(a.match(/\d+/)?.[0] || "0");
              const numB = parseInt(b.match(/\d+/)?.[0] || "0");
              return numA - numB;
            })
            .map((f) => path.join(tempDir, f));

          if (slideFiles.length > 0) {
            // Rename to slide_001.png format for consistency (in temp dir)
            const renamedFiles: string[] = [];

            for (let i = 0; i < slideFiles.length; i++) {
              const oldPath = slideFiles[i];
              const newName = `slide_${String(i + 1).padStart(3, "0")}.png`;
              const newPath = path.join(tempDir, newName);

              try {
                if (oldPath !== newPath) {
                  // Remove new file if it already exists
                  if (fs.existsSync(newPath)) {
                    fs.unlinkSync(newPath);
                  }
                  fs.renameSync(oldPath, newPath);
                }
                renamedFiles.push(newPath);
              } catch (renameErr) {
                console.error(
                  `Error renaming ${oldPath} to ${newPath}:`,
                  renameErr,
                );
                // If rename fails but file exists, use it
                if (fs.existsSync(newPath)) {
                  renamedFiles.push(newPath);
                  // Try to remove old file
                  try {
                    fs.unlinkSync(oldPath);
                  } catch {}
                }
              }
            }

            // Clean up PDF
            try {
              fs.unlinkSync(pdfFile);
            } catch {}

            // Clean up any remaining slide-*.png files (original pdftoppm output)
            allFiles.forEach((f) => {
              if (/^slide-\d+\.png$/i.test(f)) {
                try {
                  fs.unlinkSync(path.join(tempDir, f));
                  console.log(`Cleaned up pdftoppm output: ${f}`);
                } catch {}
              }
            });

            console.log(
              `Successfully converted ${renamedFiles.length} slides using pdftoppm in temp directory`,
            );

            // Atomically move files from tempDir to outputDir
            // First, delete old images in outputDir
            if (fs.existsSync(outputDir)) {
              const existingFiles = fs.readdirSync(outputDir);
              existingFiles.forEach((file) => {
                if (/^slide_\d+\.png$/i.test(file)) {
                  try {
                    fs.unlinkSync(path.join(outputDir, file));
                  } catch {}
                }
              });
            }

            // Move new images from tempDir to outputDir
            for (const tempFile of renamedFiles) {
              const fileName = path.basename(tempFile);
              const finalPath = path.join(outputDir, fileName);
              try {
                fs.renameSync(tempFile, finalPath);
              } catch (err) {
                console.error(`Error moving ${tempFile} to ${finalPath}:`, err);
                // Fallback: copy instead of move
                try {
                  fs.copyFileSync(tempFile, finalPath);
                } catch (copyErr) {
                  console.error(
                    `Error copying ${tempFile} to ${finalPath}:`,
                    copyErr,
                  );
                }
              }
            }

            // Clean up temp directory
            try {
              fs.rmSync(tempDir, { recursive: true, force: true });
            } catch {}

            // Return final paths
            return renamedFiles.map((f) =>
              path.join(outputDir, path.basename(f)),
            );
          }
        } catch {
          console.log(
            `pdftoppm not available, using LibreOffice for PDF to PNG conversion...`,
          );

          // Fallback: Use LibreOffice to convert PDF to PNG
          // This should create one PNG per page (PDF pages = slides)
          const pdfToPngCommand = `"${libreOfficePath}" --headless --invisible --nodefault --convert-to png --outdir "${tempDir}" "${pdfFile}"`;
          try {
            await execAsync(pdfToPngCommand, { timeout: 90000 });
            await new Promise((resolve) => setTimeout(resolve, 3000));
          } catch {
            // Continue to file check
          }
        }
      }

      // Wait for files to be written and collect all PNGs from tempDir
      // Check multiple times as conversion can be slow
      let allPngs: string[] = [];
      for (let attempt = 0; attempt < 15; attempt++) {
        await new Promise((resolve) => setTimeout(resolve, 1000));

        if (fs.existsSync(tempDir)) {
          const files = fs.readdirSync(tempDir);
          const pngFiles = files
            .filter((file) => {
              const lower = file.toLowerCase();
              const basename = path.basename(file, ".png").toLowerCase();
              const pdfBasename = path.basename(pdfFile, ".pdf").toLowerCase();

              // Include PNG files, but exclude:
              // - temp files
              // - the single PNG created from PDF (if LibreOffice only exported one)
              // We want to include: slide_*.png, slide-*.png, current_*.png, etc.

              if (!lower.endsWith(".png") || lower.includes("temp")) {
                return false;
              }

              // If this is the PDF filename converted to PNG (single file), exclude it
              // UNLESS it's part of a numbered series (current_001.png, etc.)
              if (basename === pdfBasename && !basename.match(/\d+$/)) {
                return false;
              }

              return true;
            })
            .map((file) => path.join(tempDir, file));

          // If we found PNGs, collect them all
          if (pngFiles.length > 0) {
            allPngs = pngFiles;
            // Wait a bit more to ensure all files are written
            if (attempt < 5) {
              continue;
            }
            break;
          }
        }
      }

      if (allPngs.length === 0 || !fs.existsSync(tempDir)) {
        console.log(`No PNG files found in ${tempDir}`);
        continue;
      }

      console.log(`Found ${allPngs.length} PNG file(s) to process`);

      if (allPngs.length > 0) {
        // Sort PNGs intelligently
        allPngs.sort((a, b) => {
          const fileA = path.basename(a);
          const fileB = path.basename(b);

          // Extract numbers from filenames (handle various formats: slide_001, slide-01, current_001, etc.)
          const numA = parseInt(fileA.match(/(\d+)/)?.[1] || "0");
          const numB = parseInt(fileB.match(/(\d+)/)?.[1] || "0");

          if (numA !== 0 && numB !== 0 && numA !== numB) {
            return numA - numB;
          }

          // If numbers are the same or missing, sort by filename
          // Also sort by modification time as fallback
          try {
            const statsA = fs.statSync(a);
            const statsB = fs.statSync(b);
            if (statsA.mtimeMs !== statsB.mtimeMs) {
              return statsA.mtimeMs - statsB.mtimeMs;
            }
          } catch {
            // Ignore stat errors
          }

          return fileA.localeCompare(fileB);
        });

        console.log(
          `Processing ${allPngs.length} PNG file(s): ${allPngs.map((f) => path.basename(f)).join(", ")}`,
        );

        // Warn if we only got one image (likely only first slide was exported)
        if (allPngs.length === 1) {
          console.warn(
            "WARNING: Only 1 PNG image was created. This usually means only the first slide was exported.\n" +
              "For best results with all slides, install pdftoppm:\n" +
              "  macOS: brew install poppler\n" +
              "  Linux: sudo apt-get install poppler-utils\n" +
              "The system will use pdftoppm automatically if available, which exports all slides reliably.",
          );
        }

        // Filter out duplicates - only keep slide_*.png files, remove slide-*.png
        const uniquePngs = allPngs.filter((pngPath) => {
          const filename = path.basename(pngPath);
          // Keep slide_*.png files, exclude slide-*.png (pdftoppm format)
          if (/^slide_\d+\.png$/i.test(filename)) {
            return true;
          }
          // For other formats, check if there's a corresponding slide_*.png in tempDir
          const match = filename.match(/\d+/);
          if (match) {
            const slideNum = match[0];
            const correspondingFile = path.join(
              tempDir,
              `slide_${slideNum.padStart(3, "0")}.png`,
            );
            // If corresponding file exists, exclude this one
            if (fs.existsSync(correspondingFile)) {
              console.log(
                `Excluding duplicate: ${filename} (${path.basename(correspondingFile)} exists)`,
              );
              try {
                fs.unlinkSync(pngPath);
              } catch {}
              return false;
            }
          }
          return true;
        });

        console.log(
          `After deduplication: ${uniquePngs.length} unique PNG files`,
        );

        // Rename to consistent format: slide_001.png, slide_002.png, etc. (in tempDir)
        const renamedFiles: string[] = [];
        for (let i = 0; i < uniquePngs.length; i++) {
          const oldPath = uniquePngs[i];
          const filename = path.basename(oldPath);
          const newFileName = `slide_${String(i + 1).padStart(3, "0")}.png`;
          const newPath = path.join(tempDir, newFileName);

          // If it's already in the correct format, just use it
          if (filename === newFileName) {
            renamedFiles.push(newPath);
            continue;
          }

          try {
            // Remove existing file if it exists
            if (fs.existsSync(newPath)) {
              fs.unlinkSync(newPath);
            }
            fs.renameSync(oldPath, newPath);
            renamedFiles.push(newPath);
          } catch (err) {
            console.error(`Error renaming ${oldPath} to ${newPath}:`, err);
            // If rename fails but new file exists, use it
            if (fs.existsSync(newPath)) {
              renamedFiles.push(newPath);
              // Try to remove old file
              try {
                fs.unlinkSync(oldPath);
              } catch {}
            } else {
              // Use original path as fallback
              renamedFiles.push(oldPath);
            }
          }
        }

        if (renamedFiles.length > 0) {
          console.log(
            `Successfully converted ${renamedFiles.length} slides to images using ${libreOfficePath} in temp directory`,
          );

          // Atomically move files from tempDir to outputDir
          // First, delete old images in outputDir
          if (fs.existsSync(outputDir)) {
            const existingFiles = fs.readdirSync(outputDir);
            existingFiles.forEach((file) => {
              if (/^slide_\d+\.png$/i.test(file)) {
                try {
                  fs.unlinkSync(path.join(outputDir, file));
                } catch {}
              }
            });
          }

          // Move new images from tempDir to outputDir
          const finalFiles: string[] = [];
          for (const tempFile of renamedFiles) {
            const fileName = path.basename(tempFile);
            const finalPath = path.join(outputDir, fileName);
            try {
              fs.renameSync(tempFile, finalPath);
              finalFiles.push(finalPath);
            } catch (err) {
              console.error(`Error moving ${tempFile} to ${finalPath}:`, err);
              // Fallback: copy instead of move
              try {
                fs.copyFileSync(tempFile, finalPath);
                finalFiles.push(finalPath);
              } catch (copyErr) {
                console.error(
                  `Error copying ${tempFile} to ${finalPath}:`,
                  copyErr,
                );
              }
            }
          }

          // Clean up temp directory
          try {
            fs.rmSync(tempDir, { recursive: true, force: true });
          } catch {}

          return finalFiles;
        }
      }
    } catch (error) {
      const errorMessage =
        error instanceof Error ? error.message : String(error);
      console.log(`Failed with ${libreOfficePath}: ${errorMessage}`);
      // Try next command
      continue;
    }
  }

  console.warn(
    "LibreOffice not available or conversion failed. To enable visual slide rendering, please install LibreOffice:\n" +
      "  macOS: brew install --cask libreoffice\n" +
      "  Linux: sudo apt-get install libreoffice\n" +
      "  Windows: Download from https://www.libreoffice.org/download/\n" +
      "Slides will be displayed as text only until LibreOffice is installed.",
  );
  return [];
}
