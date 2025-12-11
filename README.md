# Markdown Previewer & DOCX Converter

This is a simple Node.js application that provides a real-time Markdown preview and allows users to convert Markdown files to DOCX format with styling that matches the preview.

## Features

- **Real-time Preview:** View rendered Markdown as HTML.
- **Syntax Highlighting:** Supports code block syntax highlighting.
- **DOCX Conversion:** Convert Markdown files to styled Microsoft Word documents (.docx).
- **Custom Styling:** The generated DOCX files are styled to match the web preview's theme (Github-like light theme).
- **File Management:** Browse and preview Markdown files from the server or upload local files.

## Installation

1.  Clone the repository:
    ```bash
    git clone <your-repo-url>
    cd md
    ```

2.  Install dependencies:
    ```bash
    npm install
    ```

## Usage

1.  Start the server:
    ```bash
    node server.js
    ```

2.  Open your browser and navigate to `http://localhost:3004`.

3.  **Preview Files:**
    - Select a file from the list on the left (server files) or use "Select Local File" to preview a file from your computer.

4.  **Download as DOCX:**
    - Once a file is loaded, click the "Download as .docx" button in the top right corner to download the converted Word document.

## Project Structure

- `server.js`: The Express.js server handling API requests, file serving, and DOCX conversion.
- `public/`: Contains static assets (HTML, CSS, JS) for the frontend.
  - `index.html`: Main application page.
  - `style.css`: Application styles.
  - `script.js`: Frontend logic for file handling and API interaction.

## Dependencies

- **express**: Web framework for Node.js.
- **marked**: Markdown parser.
- **highlight.js**: Syntax highlighting for code blocks.
- **markdown-docx**: Library for converting Markdown to DOCX.
- **multer**: Middleware for handling `multipart/form-data` (file uploads).

## License

MIT
