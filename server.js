const express = require("express");
const fs = require("fs");
const path = require("path");
const marked = require("marked");
const hljs = require("highlight.js");
const multer = require("multer");
const { markdownDocx, Packer } = require("markdown-docx");
const JSZip = require("jszip");

// ======= DOCX 主题配置定义 =======
// 定义 markdown-docx 支持的所有主题配置选项 (基于 IMarkdownTheme 接口)
// 默认值来自 markdown-docx/src/styles/themes.ts
const DOCX_THEME_CONFIG = {
  // 颜色配置 (十六进制值，不带 #)
  colors: {
    heading1: { label: "标题1颜色", default: "2F5597", type: "color" },
    heading2: { label: "标题2颜色", default: "5B9BD5", type: "color" },
    heading3: { label: "标题3颜色", default: "44546A", type: "color" },
    heading4: { label: "标题4颜色", default: "44546A", type: "color" },
    heading5: { label: "标题5颜色", default: "44546A", type: "color" },
    heading6: { label: "标题6颜色", default: "44546A", type: "color" },
    link: { label: "链接颜色", default: "0563C1", type: "color" },
    code: { label: "代码块文本颜色", default: "032F62", type: "color" },
    codespan: { label: "行内代码颜色", default: "70AD47", type: "color" },
    codeBackground: { label: "代码背景颜色", default: "F6F6F7", type: "color" },
    blockquote: { label: "引用文本颜色", default: "666666", type: "color" },
    blockquoteBackground: { label: "引用背景颜色", default: "F9F9F9", type: "color" },
    del: { label: "删除线颜色", default: "FF0000", type: "color" },
    tag: { label: "HTML标签颜色", default: "ED7D31", type: "color" },
    html: { label: "HTML内容颜色", default: "4472C4", type: "color" },
    tableHeaderBackground: { label: "表头背景颜色", default: "F2F2F2", type: "color" },
    tableHeaderTextColor: { label: "表头文字颜色", default: "000000", type: "color" },
    tableBodyBackground: { label: "表格内容背景", default: "FFFFFF", type: "color" },
    tableAltRowBackground: { label: "交替行背景色", default: "F9F9F9", type: "color" },
    border: { label: "边框颜色", default: "A5A5A5", type: "color" },
    hr: { label: "分隔线颜色", default: "D9D9D9", type: "color" },
  },
  // 尺寸配置 (半点单位，如 36 = 18pt)
  sizes: {
    heading1Size: { label: "标题1字号", default: 36, type: "number", min: 12, max: 144 },
    heading2Size: { label: "标题2字号", default: 32, type: "number", min: 12, max: 120 },
    heading3Size: { label: "标题3字号", default: 28, type: "number", min: 12, max: 96 },
    heading4Size: { label: "标题4字号", default: 26, type: "number", min: 12, max: 72 },
    heading5Size: { label: "标题5字号", default: 24, type: "number", min: 12, max: 60 },
    heading6Size: { label: "标题6字号", default: 24, type: "number", min: 12, max: 48 },
    spaceSize: { label: "段落间距", default: 12, type: "number", min: 6, max: 48 },
    codeSize: { label: "代码字号", default: 22, type: "number", min: 8, max: 48 },
    tableSize: { label: "表格字号", default: 21, type: "number", min: 14, max: 36 },
    tableBorderWidth: { label: "表格边框宽度", default: 4, type: "number", min: 1, max: 24 },
    tableCellPadding: { label: "单元格内边距", default: 80, type: "number", min: 20, max: 300 },
  },
  // 其他配置
  options: {
    linkUnderline: { label: "链接下划线", default: true, type: "boolean" },
    tableHeaderBold: { label: "表头文字加粗", default: true, type: "boolean" },
    tableHeaderCenter: { label: "表头文字居中", default: true, type: "boolean" },
    tableAltRowEnabled: { label: "启用交替行颜色", default: false, type: "boolean" },
    tableCenterAlign: { label: "表格居中", default: true, type: "boolean" },
  }
};

// 生成默认主题配置
function getDefaultTheme() {
  const theme = {};
  for (const category of Object.values(DOCX_THEME_CONFIG)) {
    for (const [key, config] of Object.entries(category)) {
      theme[key] = config.default;
    }
  }
  return theme;
}

// 验证并清理主题配置 (只允许配置中定义的选项)
function validateTheme(inputTheme) {
  const validTheme = {};
  const defaultTheme = getDefaultTheme();
  
  for (const category of Object.values(DOCX_THEME_CONFIG)) {
    for (const [key, config] of Object.entries(category)) {
      if (inputTheme && inputTheme[key] !== undefined) {
        // 根据类型验证和转换值
        if (config.type === "color") {
          // 颜色值：确保是有效的6位十六进制
          let colorVal = String(inputTheme[key]).replace(/^#/, "").toUpperCase();
          if (/^[0-9A-F]{6}$/.test(colorVal)) {
            validTheme[key] = colorVal;
          } else {
            validTheme[key] = config.default;
          }
        } else if (config.type === "number") {
          // 数字值：确保在范围内
          let numVal = parseInt(inputTheme[key], 10);
          if (!isNaN(numVal) && numVal >= config.min && numVal <= config.max) {
            validTheme[key] = numVal;
          } else {
            validTheme[key] = config.default;
          }
        } else if (config.type === "boolean") {
          // 布尔值
          validTheme[key] = Boolean(inputTheme[key]);
        } else {
          validTheme[key] = config.default;
        }
      } else {
        validTheme[key] = defaultTheme[key];
      }
    }
  }
  
  return validTheme;
}
// ======= DOCX 主题配置定义结束 =======

// ======= 表格样式后处理函数 =======
/**
 * 处理 DOCX 文件中的表格样式
 * 通过解压 DOCX，修改 document.xml 中的表格边框和表头背景，然后重新打包
 */
async function applyTableStyles(docxBuffer, themeOptions) {
  // 提取所有表格相关配置
  const borderColor = themeOptions.border || "A5A5A5";
  const borderWidth = themeOptions.tableBorderWidth || 4;
  const tableHeaderBg = themeOptions.tableHeaderBackground || "F2F2F2";
  const tableHeaderTextColor = themeOptions.tableHeaderTextColor || "000000";
  const tableBodyBg = themeOptions.tableBodyBackground || "FFFFFF";
  const tableAltRowBg = themeOptions.tableAltRowBackground || "F9F9F9";
  const tableSize = themeOptions.tableSize || 21;
  const cellPadding = themeOptions.tableCellPadding || 80;
  const tableHeaderBold = themeOptions.tableHeaderBold !== false;
  const tableHeaderCenter = themeOptions.tableHeaderCenter !== false;
  const tableAltRowEnabled = themeOptions.tableAltRowEnabled === true;
  const tableCenterAlign = themeOptions.tableCenterAlign !== false;
  
  try {
    // 加载 DOCX (实际上是 ZIP 文件)
    const zip = await JSZip.loadAsync(docxBuffer);
    
    // 读取 document.xml
    const documentXmlPath = "word/document.xml";
    let documentXml = await zip.file(documentXmlPath).async("string");
    
    // 辅助函数：为表格生成边框 XML
    const generateBordersXml = () => `<w:tblBorders>
      <w:top w:val="single" w:sz="${borderWidth}" w:space="0" w:color="${borderColor}"/>
      <w:left w:val="single" w:sz="${borderWidth}" w:space="0" w:color="${borderColor}"/>
      <w:bottom w:val="single" w:sz="${borderWidth}" w:space="0" w:color="${borderColor}"/>
      <w:right w:val="single" w:sz="${borderWidth}" w:space="0" w:color="${borderColor}"/>
      <w:insideH w:val="single" w:sz="${borderWidth}" w:space="0" w:color="${borderColor}"/>
      <w:insideV w:val="single" w:sz="${borderWidth}" w:space="0" w:color="${borderColor}"/>
    </w:tblBorders>`;
    
    // 辅助函数：生成单元格边距 XML
    const generateCellMarginXml = () => `<w:tblCellMar>
      <w:top w:w="${cellPadding}" w:type="dxa"/>
      <w:left w:w="${cellPadding}" w:type="dxa"/>
      <w:bottom w:w="${cellPadding}" w:type="dxa"/>
      <w:right w:w="${cellPadding}" w:type="dxa"/>
    </w:tblCellMar>`;
    
    // 1. 替换现有表格边框或添加边框
    documentXml = documentXml.replace(
      /<w:tblBorders>[\s\S]*?<\/w:tblBorders>/g,
      generateBordersXml()
    );
    
    // 2. 处理 tblPr：添加边框、单元格边距、居中对齐
    documentXml = documentXml.replace(
      /<w:tblPr>([\s\S]*?)<\/w:tblPr>/g,
      (match, content) => {
        let newContent = content;
        
        // 移除现有的 tblBorders 和 tblCellMar
        newContent = newContent.replace(/<w:tblBorders>[\s\S]*?<\/w:tblBorders>/g, '');
        newContent = newContent.replace(/<w:tblCellMar>[\s\S]*?<\/w:tblCellMar>/g, '');
        newContent = newContent.replace(/<w:jc[^>]*\/>/g, '');
        
        // 添加边框
        newContent += generateBordersXml();
        
        // 添加单元格边距
        newContent += generateCellMarginXml();
        
        // 添加表格居中
        if (tableCenterAlign) {
          newContent += '<w:jc w:val="center"/>';
        }
        
        return `<w:tblPr>${newContent}</w:tblPr>`;
      }
    );
    
    // 3. 处理所有表格，识别行并应用样式
    documentXml = documentXml.replace(
      /<w:tbl>([\s\S]*?)<\/w:tbl>/g,
      (tableMatch, tableContent) => {
        let rowIndex = 0;
        
        // 处理每一行
        const updatedTableContent = tableContent.replace(
          /<w:tr([^>]*)>([\s\S]*?)<\/w:tr>/g,
          (rowMatch, rowAttrs, rowContent) => {
            const isHeaderRow = rowContent.includes('<w:tblHeader') || rowIndex === 0;
            const isAltRow = !isHeaderRow && tableAltRowEnabled && (rowIndex % 2 === 0);
            rowIndex++;
            
            // 处理行中的每个单元格
            const updatedRowContent = rowContent.replace(
              /<w:tc>([\s\S]*?)<\/w:tc>/g,
              (cellMatch, cellContent) => {
                // 处理 tcPr
                let updatedCellContent;
                
                if (cellContent.includes('<w:tcPr>')) {
                  updatedCellContent = cellContent.replace(
                    /<w:tcPr>([\s\S]*?)<\/w:tcPr>/,
                    (tcPrMatch, tcPrContent) => {
                      let newTcPrContent = tcPrContent;
                      
                      // 移除现有的 shading
                      newTcPrContent = newTcPrContent.replace(/<w:shd[^>]*\/>/g, '');
                      
                      // 添加背景色
                      if (isHeaderRow) {
                        newTcPrContent += `<w:shd w:val="clear" w:color="auto" w:fill="${tableHeaderBg}"/>`;
                      } else if (isAltRow) {
                        newTcPrContent += `<w:shd w:val="clear" w:color="auto" w:fill="${tableAltRowBg}"/>`;
                      } else if (tableBodyBg !== "FFFFFF") {
                        newTcPrContent += `<w:shd w:val="clear" w:color="auto" w:fill="${tableBodyBg}"/>`;
                      }
                      
                      return `<w:tcPr>${newTcPrContent}</w:tcPr>`;
                    }
                  );
                } else {
                  // 如果没有 tcPr，创建一个
                  let tcPrContent = '';
                  if (isHeaderRow) {
                    tcPrContent = `<w:shd w:val="clear" w:color="auto" w:fill="${tableHeaderBg}"/>`;
                  } else if (isAltRow) {
                    tcPrContent = `<w:shd w:val="clear" w:color="auto" w:fill="${tableAltRowBg}"/>`;
                  } else if (tableBodyBg !== "FFFFFF") {
                    tcPrContent = `<w:shd w:val="clear" w:color="auto" w:fill="${tableBodyBg}"/>`;
                  }
                  if (tcPrContent) {
                    updatedCellContent = `<w:tcPr>${tcPrContent}</w:tcPr>` + cellContent;
                  } else {
                    updatedCellContent = cellContent;
                  }
                }
                
                // 处理单元格中的段落文字样式
                if (isHeaderRow) {
                  updatedCellContent = updatedCellContent.replace(
                    /<w:p([^>]*)>([\s\S]*?)<\/w:p>/g,
                    (pMatch, pAttrs, pContent) => {
                      // 处理段落属性 - 添加居中对齐
                      let pPrContent = '';
                      const pPrMatch = pContent.match(/<w:pPr>([\s\S]*?)<\/w:pPr>/);
                      if (pPrMatch) {
                        pPrContent = pPrMatch[1];
                        pContent = pContent.replace(/<w:pPr>[\s\S]*?<\/w:pPr>/, '');
                      }
                      
                      // 移除现有的对齐设置
                      pPrContent = pPrContent.replace(/<w:jc[^>]*\/>/g, '');
                      
                      // 添加居中对齐
                      if (tableHeaderCenter) {
                        pPrContent += '<w:jc w:val="center"/>';
                      }
                      
                      // 处理 run 属性
                      let updatedPContent = pContent.replace(
                        /<w:r>([\s\S]*?)<\/w:r>/g,
                        (rMatch, rContent) => {
                          let rPrContent = '';
                          
                          // 提取现有的 rPr 内容
                          const rPrMatch = rContent.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/);
                          if (rPrMatch) {
                            rPrContent = rPrMatch[1];
                            rContent = rContent.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/, '');
                          }
                          
                          // 移除现有的颜色、字号、加粗设置
                          rPrContent = rPrContent.replace(/<w:color[^>]*\/>/g, '');
                          rPrContent = rPrContent.replace(/<w:sz[^>]*\/>/g, '');
                          rPrContent = rPrContent.replace(/<w:szCs[^>]*\/>/g, '');
                          rPrContent = rPrContent.replace(/<w:b[^>]*\/>/g, '');
                          
                          // 添加表头文字颜色
                          rPrContent += `<w:color w:val="${tableHeaderTextColor}"/>`;
                          // 添加字号
                          rPrContent += `<w:sz w:val="${tableSize}"/><w:szCs w:val="${tableSize}"/>`;
                          // 添加加粗
                          if (tableHeaderBold) {
                            rPrContent += '<w:b/><w:bCs/>';
                          }
                          
                          return `<w:r><w:rPr>${rPrContent}</w:rPr>${rContent}</w:r>`;
                        }
                      );
                      
                      // 构建最终段落
                      const finalPPr = pPrContent ? `<w:pPr>${pPrContent}</w:pPr>` : '';
                      return `<w:p${pAttrs}>${finalPPr}${updatedPContent}</w:p>`;
                    }
                  );
                } else {
                  // 普通行：只设置字号
                  updatedCellContent = updatedCellContent.replace(
                    /<w:r>([\s\S]*?)<\/w:r>/g,
                    (rMatch, rContent) => {
                      let rPrContent = '';
                      
                      const rPrMatch = rContent.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/);
                      if (rPrMatch) {
                        rPrContent = rPrMatch[1];
                        rContent = rContent.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/, '');
                      }
                      
                      // 移除现有字号设置
                      rPrContent = rPrContent.replace(/<w:sz[^>]*\/>/g, '');
                      rPrContent = rPrContent.replace(/<w:szCs[^>]*\/>/g, '');
                      
                      // 添加字号
                      rPrContent += `<w:sz w:val="${tableSize}"/><w:szCs w:val="${tableSize}"/>`;
                      
                      return `<w:r><w:rPr>${rPrContent}</w:rPr>${rContent}</w:r>`;
                    }
                  );
                }
                
                return `<w:tc>${updatedCellContent}</w:tc>`;
              }
            );
            
            return `<w:tr${rowAttrs}>${updatedRowContent}</w:tr>`;
          }
        );
        
        return `<w:tbl>${updatedTableContent}</w:tbl>`;
      }
    );
    
    // 更新 ZIP 中的 document.xml
    zip.file(documentXmlPath, documentXml);
    
    // 重新生成 DOCX buffer
    const newBuffer = await zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 9 }
    });
    
    return newBuffer;
  } catch (err) {
    console.error("Error applying table styles:", err);
    // 如果处理失败，返回原始 buffer
    return docxBuffer;
  }
}
// ======= 表格样式后处理函数结束 =======

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

// 获取 DOCX 主题配置选项 API
app.get("/api/docx-config", (req, res) => {
  res.json({
    config: DOCX_THEME_CONFIG,
    defaults: getDefaultTheme()
  });
});

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

    // 使用验证后的主题配置
    let themeOptions = getDefaultTheme();

    if (req.body.theme) {
      try {
        const customTheme = JSON.parse(req.body.theme);
        // 验证并只应用允许的配置选项
        themeOptions = validateTheme(customTheme);
      } catch (e) {
        console.error("Error parsing custom theme:", e);
      }
    }

    const doc = await markdownDocx(markdown, {
      theme: themeOptions,
    });

    // 先生成原始 DOCX buffer
    let docxBuffer = await Packer.toBuffer(doc);
    
    // 使用后处理函数应用表格样式
    docxBuffer = await applyTableStyles(docxBuffer, themeOptions);

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
