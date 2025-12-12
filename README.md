# Markdown Previewer & DOCX Converter

A feature-rich Node.js application that provides real-time Markdown preview and converts Markdown files to customizable styled DOCX format.

## ✨ Features

- **Real-time Preview** - Instantly view rendered Markdown as HTML
- **Syntax Highlighting** - Code block highlighting with GitHub Dark theme
- **DOCX Export** - Convert Markdown to formatted Microsoft Word documents (.docx)
- **Theme Customization** - Ribbon-style toolbar for DOCX export styling:
  - 🎨 **Colors** - Headings, links, code, quotes, tables, and more
  - 📏 **Sizes** - Font sizes, spacing, border widths
  - ⚙️ **Options** - Link underlines, table header styles, alternating rows
- **Local File Support** - Preview Markdown files from your computer
- **Server File Browser** - Browse and preview Markdown files from server directory
- **Persistent Settings** - Theme configurations saved to localStorage

## 📦 Installation

1. Clone the repository:
   ```bash
   git clone <your-repo-url>
   cd md
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

## 🚀 Usage

1. Start the server:
   ```bash
   node server.js
   ```

2. Open your browser and navigate to `http://localhost:3004`

3. **Preview Files:**
   - Select a file from the left sidebar (server files)
   - Or click "Select Local File" to upload local Markdown files

4. **Configure DOCX Styles:**
   - Use the Ribbon toolbar at the top to adjust export styles
   - Switch between "Colors", "Sizes", and "Options" tabs
   - Click "Reset" to restore default settings

5. **Export to DOCX:**
   - Click the "Export DOCX" button to download the converted Word document

## 📁 Project Structure

```
md/
├── server.js          # Express server (API, file serving, DOCX conversion)
├── package.json       # Project configuration and dependencies
└── public/            # Frontend static assets
    ├── index.html     # Main page
    ├── style.css      # Stylesheets
    └── script.js      # Frontend logic
```

## 🔧 API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/api/files` | GET | List server Markdown files |
| `/api/files/:filename` | GET | Get file content |
| `/api/docx-config` | GET | Get DOCX theme configuration options |
| `/api/convert` | POST | Convert Markdown to DOCX |

## 📚 Dependencies

| Package | Purpose |
|---------|---------|
| express | Node.js web framework |
| marked | Markdown parser |
| highlight.js | Code syntax highlighting |
| markdown-docx | Markdown to DOCX conversion |
| multer | File upload middleware |
| jszip | DOCX post-processing (table styles) |
| get-port | Auto-detect available port |

## 📄 License

MIT
