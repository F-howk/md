const express = require("express");
const fs = require("fs");
const path = require("path");
const marked = require("marked");
const hljs = require("highlight.js");
const multer = require("multer");
const { markdownDocx, Packer } = require("markdown-docx"); // Correctly import markdown-docx

// Define the highlight function with JSON pretty-printing
const highlightFn = function (code, lang) {
  if (lang === "json") {
    try {
      const jsonObj = JSON.parse(code);
      code = JSON.stringify(jsonObj, null, 2);
    } catch (e) {
      // Not valid JSON, fall back
    }
  }
  const language = hljs.getLanguage(lang) ? lang : "plaintext";
  return hljs.highlight(code, { language }).value;
};

// --- Marked.js setup for browser preview ---
const renderer = new marked.Renderer();

marked.setOptions({
  gfm: true,
  breaks: false,
  renderer: renderer,
  highlight: highlightFn, // Restore global highlight option
});
// --- End of Marked.js setup ---

const app = express();
const port = 3004; // Fixed port

const storage = multer.memoryStorage();
const upload = multer({
  storage: storage,
  limits: { fileSize: 100 * 1024 * 1024 },
});

app.use(express.static("public"));

app.get("/api/files", (req, res) => {
  fs.readdir(__dirname, (err, files) => {
    if (err) {
      return res.status(500).send("Unable to scan directory: " + err);
    }
    const markdownFiles = files.filter(
      (file) => path.extname(file).toLowerCase() === ".md"
    );
    res.json(markdownFiles);
  });
});

// --- Refactored DOCX conversion endpoint using markdown-docx ---
app.post("/api/convert/docx/:filename", upload.single("mdfile"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send("No file uploaded.");
    }

    let markdown = req.file.buffer.toString("utf8");
    let filename = req.params.filename;

    if (!markdown) {
      return res.status(400).send("Markdown content is empty.");
    }

    const doc = await markdownDocx(markdown, {
      theme: {
        heading1: "333333",
        heading2: "333333",
        heading3: "333333",
        heading4: "333333",
        heading5: "333333",
        heading6: "333333",
        code: "333333",
        codeBackground: "F8F8F8",
        codespan: "333333",
        blockquote: "666666",
        blockquoteBackground: "FFFFFF",
        link: "007BFF",
        tableHeaderBackground: "F2F2F2",
        border: "DDDDDD",
        hr: "EEEEEE",
        del: "666666",
        html: "333333",
      },
    });
    const docxBuffer = await Packer.toBuffer(doc);

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    const encodedFilename = encodeURIComponent(filename);
    res.setHeader(
      "Content-Disposition",
      `attachment; filename*=UTF-8''${encodedFilename}`
    );
    res.send(docxBuffer); // Send buffer directly
  } catch (error) {
    console.error("Error converting to DOCX:", error);
    res.status(500).send("Error converting file.");
  }
});
// --- End of refactored endpoint ---

app.get("/api/files/:filename", (req, res) => {
  const { filename } = req.params;
  if (filename.includes("..")) {
    return res.status(400).send("Invalid filename.");
  }
  const filePath = path.join(__dirname, filename);

  fs.readFile(filePath, "utf8", (err, data) => {
    if (err) {
      return res.status(404).send("File not found or could not be read.");
    }
    res.type("text/plain").send(data);
  });
});

app.use(express.text({ type: "text/plain", limit: "100mb" }));

app.post("/api/render", (req, res) => {
  try {
    const markdown = req.body;
    const htmlContent = marked.parse(markdown);
    res.type("text/html").send(htmlContent);
  } catch (error) {
    console.error("Error rendering markdown:", error);
    res.status(500).send("Error rendering markdown.");
  }
});

app.listen(port, () => {
  console.log(`Markdown preview server running at http://localhost:${port}`);
});
