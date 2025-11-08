# AI-Powered PPTX Editor

A web application that allows users to upload PPTX files, view them in the browser, and use AI-powered chat commands to edit presentations.

## Architecture

- **Frontend**: Next.js app (port 3000) with assistant-ui chat interface
- **Backend**: Express server (port 3001) for file operations, PPTX parsing, AI integration, and file editing

## Setup

### Prerequisites

- Node.js (v18 or higher)
- npm or yarn
- OpenAI API key

### Backend Setup

1. Navigate to the backend directory:

```bash
cd backend
```

2. Install dependencies:

```bash
npm install
```

3. Create a `.env` file in the `backend` directory:

```
OPENAI_API_KEY=your_openai_api_key_here
```

4. Start the backend server:

```bash
npm run dev
```

The backend server will run on http://localhost:3001

### Frontend Setup

1. In the root directory, create a `.env.local` file:

```
OPENAI_API_KEY=your_openai_api_key_here
```

2. Install dependencies (if not already installed):

```bash
npm install
```

3. Start the development server:

```bash
npm run dev
```

4. Open [http://localhost:3000](http://localhost:3000) in your browser

## Usage

1. **Upload a PPTX file**: Click the "Upload PPTX" button in the header and select a .pptx file
2. **View the presentation**: The uploaded presentation will be displayed in the viewer (left side)
3. **Edit via chat**: Use the chat interface (right side) to give AI-powered commands, such as:
   - "Change the title on slide 2 to 'Introduction'"
   - "Change the background color of slide 1 to blue"
   - "Update the text on slide 3 from 'Old Text' to 'New Text'"

The presentation will automatically refresh after successful edits.

## Features

- Upload and view PPTX files
- AI-powered text editing
- Background color changes
- Real-time preview updates
- Chat-based interface for natural language commands

## Technical Stack

### Frontend

- Next.js 15
- React 19
- TypeScript
- assistant-ui
- Tailwind CSS
- axios

### Backend

- Express
- TypeScript
- OpenAI API
- unzipper (for PPTX parsing)
- xml2js (for XML manipulation)
- jszip (for PPTX editing)
- multer (for file uploads)

## Project Structure

```
pavo/
├── app/                    # Next.js frontend
│   ├── api/chat/          # Chat API route (proxies to backend)
│   ├── assistant.tsx      # Main assistant component
│   └── page.tsx           # Home page
├── components/
│   ├── PptxViewer.tsx     # PPTX viewer component
│   └── ...
└── backend/                # Express backend
    ├── src/
    │   ├── index.ts       # Express server
    │   ├── parser.ts      # PPTX parsing
    │   ├── editor.ts      # PPTX editing
    │   └── ai.ts          # AI integration
    └── uploads/           # Uploaded PPTX files
```

## Development

### Running Both Servers

You need to run both the frontend and backend servers simultaneously:

1. Terminal 1 - Backend:

```bash
cd backend
npm run dev
```

2. Terminal 2 - Frontend:

```bash
npm run dev
```

## Visual Slide Rendering

To see slides with images, colors, and formatting (not just text), you need to install **LibreOffice**:

### macOS

```bash
brew install --cask libreoffice
```

### Linux

```bash
sudo apt-get install libreoffice
# or
sudo yum install libreoffice
```

### Windows

Download and install from [LibreOffice website](https://www.libreoffice.org/download/)

After installing LibreOffice, restart the backend server. The viewer will automatically convert PPTX slides to images and display them visually with all formatting, colors, and images preserved.

**Note:** If LibreOffice is not installed, the viewer will display slides as text-only content. This is functional but doesn't show the visual appearance of the slides.

## Notes

- Visual slide rendering requires LibreOffice to be installed (see above)
- All uploaded files are stored in `backend/uploads/current.pptx`
- Converted slide images are cached in `backend/uploads/slide_images/`
- The backend maintains slide context in memory (resets on server restart)
- PPTX editing is done by manipulating the ZIP/XML structure directly
