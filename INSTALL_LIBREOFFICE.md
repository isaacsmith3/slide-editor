# Installing LibreOffice for Visual Slide Rendering

To see your PowerPoint slides with **colors, images, and formatting** (not just text), you need to install LibreOffice.

## Why LibreOffice?

LibreOffice is used to convert PPTX files to PNG images, allowing you to see exactly how your slides look with all their visual elements preserved.

## Installation Instructions

### macOS (Recommended: Homebrew)

```bash
brew install --cask libreoffice
```

After installation, verify it works:

```bash
/Applications/LibreOffice.app/Contents/MacOS/soffice --version
```

### macOS (Manual Install)

1. Download LibreOffice from https://www.libreoffice.org/download/
2. Install the .dmg file
3. LibreOffice will be available at `/Applications/LibreOffice.app`

### Linux (Ubuntu/Debian)

```bash
sudo apt-get update
sudo apt-get install libreoffice
```

### Linux (Red Hat/CentOS)

```bash
sudo yum install libreoffice
```

### Windows

1. Download LibreOffice from https://www.libreoffice.org/download/
2. Run the installer
3. Make sure to add LibreOffice to your PATH during installation

## Optional: Install pdftoppm for Better Performance

For the best slide conversion quality and reliability, install `pdftoppm` (from poppler-utils):

### macOS

```bash
brew install poppler
```

### Linux (Ubuntu/Debian)

```bash
sudo apt-get install poppler-utils
```

### Linux (Red Hat/CentOS)

```bash
sudo yum install poppler-utils
```

**Why pdftoppm?**

- Converts all slides reliably (LibreOffice PDF-to-PNG can sometimes only export the first slide)
- Better image quality
- Faster conversion
- More reliable page-by-page export

The system will automatically use `pdftoppm` if available, otherwise it falls back to LibreOffice.

## Verify Installation

After installing LibreOffice (and optionally pdftoppm), restart your backend server and upload a PPTX file. The viewer should automatically convert and display all slides as images.

## Troubleshooting

If slides still show as text-only after installing LibreOffice:

1. **Check if LibreOffice is in your PATH:**

   ```bash
   which libreoffice
   # or
   which soffice
   ```

2. **Try the full path (macOS):**

   ```bash
   /Applications/LibreOffice.app/Contents/MacOS/soffice --version
   ```

3. **Check backend logs** for conversion errors

4. **Restart the backend server** after installing LibreOffice

## What You'll See

**Without LibreOffice:**

- Text-only slide content
- No colors, images, or formatting
- Functional but limited visual representation

**With LibreOffice:**

- Full visual slides with all formatting
- Colors, images, and layout preserved
- Exact representation of your PowerPoint slides
- Thumbnail navigation with visual previews
