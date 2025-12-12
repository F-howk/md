document.addEventListener("DOMContentLoaded", () => {
  const fileList = document.getElementById("file-list");
  const content = document.getElementById("content");
  const contentToolbar = document.getElementById("content-toolbar");
  const downloadBtn = document.getElementById("download-docx-btn");
  const localFileInput = document.getElementById("local-file-input");
  const configContainer = document.getElementById("config-container");
  const resetConfigBtn = document.getElementById("reset-config-btn");
  const styleToolbar = document.getElementById("style-settings");
  const toggleRibbonBtn = document.getElementById("toggle-ribbon-btn");
  
  let activeFileElement = null;
  let currentMarkdown = "";
  let currentFileName = "document";
  let localFileObjects = {}; // 存储本地文件对象
  
  // DOCX 主题配置状态
  let docxConfig = null;      // 服务器返回的配置定义
  let currentTheme = {};      // 当前用户设置的主题值
  let defaultTheme = {};      // 默认主题值
  let activeTab = "colors";   // 当前激活的选项卡

  // ======= 工具函数 =======
  
  // 防抖函数 - 减少频繁更新导致的卡顿
  function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
      const later = () => {
        clearTimeout(timeout);
        func(...args);
      };
      clearTimeout(timeout);
      timeout = setTimeout(later, wait);
    };
  }
  
  // 节流函数 - 用于颜色选择器
  function throttle(func, limit) {
    let inThrottle;
    return function(...args) {
      if (!inThrottle) {
        func(...args);
        inThrottle = true;
        setTimeout(() => inThrottle = false, limit);
      }
    };
  }

  // ======= 配置管理 =======
  
  // 从服务器加载配置选项
  async function loadDocxConfig() {
    try {
      const response = await fetch("/api/docx-config");
      if (!response.ok) throw new Error("Failed to load config");
      
      const data = await response.json();
      docxConfig = data.config;
      defaultTheme = data.defaults;
      
      // 从 localStorage 加载保存的配置，如果没有则使用默认值
      const savedTheme = localStorage.getItem("docxTheme");
      if (savedTheme) {
        try {
          currentTheme = JSON.parse(savedTheme);
          // 确保所有默认键都存在
          currentTheme = { ...defaultTheme, ...currentTheme };
        } catch (e) {
          currentTheme = { ...defaultTheme };
        }
      } else {
        currentTheme = { ...defaultTheme };
      }
      
      renderConfigUI();
    } catch (error) {
      console.error("Error loading DOCX config:", error);
      configContainer.innerHTML = '<p style="color: red; font-size: 12px;">加载配置失败</p>';
    }
  }

  // 渲染配置界面 - Word Ribbon 风格
  function renderConfigUI() {
    if (!docxConfig) return;
    
    configContainer.innerHTML = "";
    
    // 创建防抖版本的更新函数
    const debouncedUpdate = debounce(updatePreviewStyles, 50);
    const throttledUpdate = throttle(updatePreviewStyles, 30);
    const debouncedSave = debounce(saveTheme, 300);
    
    const sections = [
      { key: "colors", title: "颜色" },
      { key: "sizes", title: "尺寸" },
      { key: "options", title: "选项" }
    ];
    
    // 为每个分类创建面板
    sections.forEach(section => {
      const sectionConfig = docxConfig[section.key];
      if (!sectionConfig || Object.keys(sectionConfig).length === 0) return;
      
      const panel = document.createElement("div");
      panel.className = "ribbon-panel" + (section.key === activeTab ? " active" : "");
      panel.dataset.tab = section.key;
      
      // 对配置项进行分组（例如标题相关、表格相关等）
      const groups = groupConfigItems(section.key, sectionConfig);
      
      groups.forEach(group => {
        const groupDiv = document.createElement("div");
        groupDiv.className = "ribbon-group";
        
        const itemsDiv = document.createElement("div");
        itemsDiv.className = "ribbon-group-items";
        
        group.items.forEach(({ key, config }) => {
          const itemDiv = document.createElement("div");
          itemDiv.className = "ribbon-item";
          
          let input;
          
          if (config.type === "color") {
            input = document.createElement("input");
            input.type = "color";
            input.id = `config-${key}`;
            input.value = "#" + (currentTheme[key] || config.default);
            input.title = config.label;
            input.addEventListener("input", (e) => {
              currentTheme[key] = e.target.value.replace("#", "").toUpperCase();
              throttledUpdate();
              debouncedSave();
            });
            itemDiv.appendChild(input);
          } else if (config.type === "number") {
            input = document.createElement("input");
            input.type = "number";
            input.id = `config-${key}`;
            input.min = config.min;
            input.max = config.max;
            input.value = currentTheme[key] !== undefined ? currentTheme[key] : config.default;
            input.title = config.label;
            input.addEventListener("input", (e) => {
              let val = parseInt(e.target.value, 10);
              if (!isNaN(val)) {
                if (val < config.min) val = config.min;
                if (val > config.max) val = config.max;
                currentTheme[key] = val;
                debouncedUpdate();
                debouncedSave();
              }
            });
            input.addEventListener("blur", (e) => {
              let val = parseInt(e.target.value, 10);
              if (isNaN(val) || val < config.min) val = config.min;
              if (val > config.max) val = config.max;
              e.target.value = val;
              currentTheme[key] = val;
              updatePreviewStyles();
              saveTheme();
            });
            itemDiv.appendChild(input);
          } else if (config.type === "boolean") {
            const wrapper = document.createElement("div");
            wrapper.className = "checkbox-wrapper";
            
            input = document.createElement("input");
            input.type = "checkbox";
            input.id = `config-${key}`;
            input.checked = currentTheme[key] !== undefined ? currentTheme[key] : config.default;
            input.title = config.label;
            
            const checkLabel = document.createElement("span");
            checkLabel.textContent = input.checked ? "开" : "关";
            
            input.addEventListener("change", (e) => {
              currentTheme[key] = e.target.checked;
              checkLabel.textContent = e.target.checked ? "开" : "关";
              updatePreviewStyles();
              saveTheme();
            });
            
            wrapper.appendChild(input);
            wrapper.appendChild(checkLabel);
            itemDiv.appendChild(wrapper);
          }
          
          const label = document.createElement("label");
          label.setAttribute("for", `config-${key}`);
          label.textContent = config.label;
          itemDiv.appendChild(label);
          
          itemsDiv.appendChild(itemDiv);
        });
        
        groupDiv.appendChild(itemsDiv);
        
        const titleDiv = document.createElement("div");
        titleDiv.className = "ribbon-group-title";
        titleDiv.textContent = group.title;
        groupDiv.appendChild(titleDiv);
        
        panel.appendChild(groupDiv);
      });
      
      configContainer.appendChild(panel);
    });
    
    // 绑定选项卡切换事件
    document.querySelectorAll(".ribbon-tab").forEach(tab => {
      tab.addEventListener("click", () => {
        const tabKey = tab.dataset.tab;
        activeTab = tabKey;
        
        // 更新选项卡状态
        document.querySelectorAll(".ribbon-tab").forEach(t => t.classList.remove("active"));
        tab.classList.add("active");
        
        // 更新面板显示
        document.querySelectorAll(".ribbon-panel").forEach(p => p.classList.remove("active"));
        const panel = document.querySelector(`.ribbon-panel[data-tab="${tabKey}"]`);
        if (panel) panel.classList.add("active");
      });
    });
  }
  
  // 将配置项按功能分组
  function groupConfigItems(sectionKey, sectionConfig) {
    const items = Object.entries(sectionConfig).map(([key, config]) => ({ key, config }));
    
    if (sectionKey === "colors") {
      return [
        { title: "标题", items: items.filter(i => i.key.startsWith("heading")) },
        { title: "文本", items: items.filter(i => ["link", "blockquote", "del"].includes(i.key)) },
        { title: "代码", items: items.filter(i => i.key.includes("code") || i.key.includes("Code")) },
        { title: "表格", items: items.filter(i => i.key.includes("table") || i.key.includes("Table")) },
        { title: "其他", items: items.filter(i => ["border", "hr", "tag", "html", "blockquoteBackground"].includes(i.key)) },
      ].filter(g => g.items.length > 0);
    } else if (sectionKey === "sizes") {
      return [
        { title: "标题字号", items: items.filter(i => i.key.includes("heading") || i.key.includes("Heading")) },
        { title: "其他", items: items.filter(i => !i.key.includes("heading") && !i.key.includes("Heading")) },
      ].filter(g => g.items.length > 0);
    } else {
      return [{ title: "设置", items }];
    }
  }

  // 保存主题到 localStorage
  function saveTheme() {
    localStorage.setItem("docxTheme", JSON.stringify(currentTheme));
  }

  // 重置为默认值
  function resetToDefaults() {
    currentTheme = { ...defaultTheme };
    localStorage.removeItem("docxTheme");
    renderConfigUI();
    updatePreviewStyles();
  }

  // 获取当前主题（用于导出）
  function getTheme() {
    return { ...currentTheme };
  }

  // 更新预览样式 - 使用 CSS 变量优化性能
  function updatePreviewStyles() {
    // 首次调用时创建静态样式表
    const staticStyleId = 'static-preview-styles';
    if (!document.getElementById(staticStyleId)) {
      const staticStyle = document.createElement('style');
      staticStyle.id = staticStyleId;
      staticStyle.textContent = `
        .markdown-body h1 { color: var(--h1-color) !important; font-size: var(--h1-size) !important; }
        .markdown-body h2 { color: var(--h2-color) !important; font-size: var(--h2-size) !important; }
        .markdown-body h3 { color: var(--h3-color) !important; font-size: var(--h3-size) !important; }
        .markdown-body h4 { color: var(--h4-color) !important; font-size: var(--h4-size) !important; }
        .markdown-body h5 { color: var(--h5-color) !important; font-size: var(--h5-size) !important; }
        .markdown-body h6 { color: var(--h6-color) !important; font-size: var(--h6-size) !important; }
        .markdown-body a { color: var(--link-color) !important; text-decoration: var(--link-decoration) !important; }
        .markdown-body pre code { color: var(--code-color) !important; font-size: var(--code-size) !important; }
        .markdown-body code { color: var(--codespan-color) !important; }
        .markdown-body pre { background-color: var(--code-bg) !important; border-color: var(--border-color) !important; }
        .markdown-body blockquote { color: var(--blockquote-color) !important; background-color: var(--blockquote-bg) !important; border-left-color: var(--blockquote-color) !important; }
        .markdown-body table { 
          border-color: var(--border-color) !important; 
          border-width: var(--table-border-width) !important;
          margin-left: var(--table-margin-left) !important;
          margin-right: var(--table-margin-right) !important;
        }
        .markdown-body table th { 
          background-color: var(--table-header-bg) !important; 
          color: var(--table-header-text) !important;
          font-size: var(--table-size) !important;
          font-weight: var(--table-header-weight) !important;
          padding: var(--table-cell-padding) !important;
          border-color: var(--border-color) !important;
          border-width: var(--table-border-width) !important;
        }
        .markdown-body table td { 
          background-color: var(--table-body-bg) !important;
          font-size: var(--table-size) !important;
          padding: var(--table-cell-padding) !important;
          border-color: var(--border-color) !important;
          border-width: var(--table-border-width) !important;
        }
        .markdown-body table tbody tr:nth-child(even) td { 
          background-color: var(--table-alt-row-bg) !important; 
        }
        .markdown-body hr { background-color: var(--hr-color) !important; border-color: var(--hr-color) !important; }
        .markdown-body del { color: var(--del-color) !important; }
      `;
      document.head.appendChild(staticStyle);
    }

    // 更新 CSS 变量 - 这比重写整个样式表更高效
    const root = document.documentElement;
    
    // 颜色变量
    root.style.setProperty('--h1-color', '#' + (currentTheme.heading1 || '2F5597'));
    root.style.setProperty('--h2-color', '#' + (currentTheme.heading2 || '5B9BD5'));
    root.style.setProperty('--h3-color', '#' + (currentTheme.heading3 || '44546A'));
    root.style.setProperty('--h4-color', '#' + (currentTheme.heading4 || '44546A'));
    root.style.setProperty('--h5-color', '#' + (currentTheme.heading5 || '44546A'));
    root.style.setProperty('--h6-color', '#' + (currentTheme.heading6 || '44546A'));
    root.style.setProperty('--link-color', '#' + (currentTheme.link || '0563C1'));
    root.style.setProperty('--link-decoration', currentTheme.linkUnderline === false ? 'none' : 'underline');
    root.style.setProperty('--code-color', '#' + (currentTheme.code || '032F62'));
    root.style.setProperty('--codespan-color', '#' + (currentTheme.codespan || '70AD47'));
    root.style.setProperty('--code-bg', '#' + (currentTheme.codeBackground || 'F6F6F7'));
    root.style.setProperty('--blockquote-color', '#' + (currentTheme.blockquote || '666666'));
    root.style.setProperty('--blockquote-bg', '#' + (currentTheme.blockquoteBackground || 'F9F9F9'));
    root.style.setProperty('--border-color', '#' + (currentTheme.border || 'A5A5A5'));
    root.style.setProperty('--hr-color', '#' + (currentTheme.hr || 'D9D9D9'));
    root.style.setProperty('--del-color', '#' + (currentTheme.del || 'FF0000'));
    
    // 表格颜色
    root.style.setProperty('--table-header-bg', '#' + (currentTheme.tableHeaderBackground || 'F2F2F2'));
    root.style.setProperty('--table-header-text', '#' + (currentTheme.tableHeaderTextColor || '000000'));
    root.style.setProperty('--table-body-bg', '#' + (currentTheme.tableBodyBackground || 'FFFFFF'));
    root.style.setProperty('--table-alt-row-bg', currentTheme.tableAltRowEnabled ? '#' + (currentTheme.tableAltRowBackground || 'F9F9F9') : 'var(--table-body-bg)');
    
    // 尺寸变量
    root.style.setProperty('--h1-size', ((currentTheme.heading1Size || 36) / 2) + 'pt');
    root.style.setProperty('--h2-size', ((currentTheme.heading2Size || 32) / 2) + 'pt');
    root.style.setProperty('--h3-size', ((currentTheme.heading3Size || 28) / 2) + 'pt');
    root.style.setProperty('--h4-size', ((currentTheme.heading4Size || 26) / 2) + 'pt');
    root.style.setProperty('--h5-size', ((currentTheme.heading5Size || 24) / 2) + 'pt');
    root.style.setProperty('--h6-size', ((currentTheme.heading6Size || 24) / 2) + 'pt');
    root.style.setProperty('--code-size', ((currentTheme.codeSize || 22) / 2) + 'pt');
    
    // 表格尺寸
    root.style.setProperty('--table-size', ((currentTheme.tableSize || 21) / 2) + 'pt');
    root.style.setProperty('--table-border-width', ((currentTheme.tableBorderWidth || 4) / 2) + 'px');
    root.style.setProperty('--table-cell-padding', ((currentTheme.tableCellPadding || 80) / 20) + 'px');
    root.style.setProperty('--table-header-weight', currentTheme.tableHeaderBold !== false ? 'bold' : 'normal');
    
    // 表格居中
    const centerAlign = currentTheme.tableCenterAlign !== false;
    root.style.setProperty('--table-margin-left', centerAlign ? 'auto' : '0');
    root.style.setProperty('--table-margin-right', centerAlign ? 'auto' : '0');
  }

  // 重置按钮事件
  resetConfigBtn.addEventListener("click", () => {
    if (confirm("确定要将所有样式设置重置为默认值吗？")) {
      resetToDefaults();
    }
  });

  // 收起/展开 Ribbon 工具栏
  toggleRibbonBtn.addEventListener("click", () => {
    styleToolbar.classList.toggle("collapsed");
    // 保存折叠状态
    localStorage.setItem("ribbonCollapsed", styleToolbar.classList.contains("collapsed"));
  });

  // 恢复 Ribbon 折叠状态
  if (localStorage.getItem("ribbonCollapsed") === "true") {
    styleToolbar.classList.add("collapsed");
  }

  // ======= 文件管理 =======
  
  // 用于通过服务器渲染 Markdown 并显示它的函数
  function renderAndDisplay(filename, markdown) {
    currentMarkdown = markdown;
    currentFileName = filename;
    content.innerHTML = "Rendering...";

    fetch("/api/render", {
      method: "POST",
      headers: {
        "Content-Type": "text/plain",
      },
      body: markdown,
    })
      .then((response) => {
        if (!response.ok) {
          throw new Error("Server rendering failed");
        }
        return response.text();
      })
      .then((html) => {
        content.innerHTML = html;
        contentToolbar.style.display = "flex";
        updatePreviewStyles(); // Apply styles after rendering
      })
      .catch((error) => {
        console.error("Error rendering content:", error);
        content.innerHTML = `<p style="color: red;">Error rendering content.</p>`;
      });
  }

  // 处理文件列表项的激活状态
  function setActiveFile(item) {
    if (activeFileElement) {
      activeFileElement.classList.remove("active");
    }
    activeFileElement = item;
    activeFileElement.classList.add("active");
  }

  // 从服务器端 Markdown 文件加载并显示内容
  function loadServerFile(filename) {
    content.innerHTML = "Loading...";
    contentToolbar.style.display = "none";
    fetch(`/api/files/${encodeURIComponent(filename)}`)
      .then((response) => {
        if (!response.ok) {
          throw new Error("Network response was not ok");
        }
        return response.text();
      })
      .then((markdown) => {
        renderAndDisplay(filename, markdown);
      })
      .catch((error) => {
        console.error("Error fetching file content from server:", error);
        content.innerHTML = `<p style="color: red;">Error loading file from server: ${filename}</p>`;
        currentMarkdown = "";
        currentFileName = "document";
      });
  }

  // 处理本地文件选择和显示
  localFileInput.addEventListener("change", (event) => {
    const files = event.target.files;
    if (files.length > 0) {
      Array.from(files).forEach((file) => {
        localFileObjects[file.name] = file; // Store the File object
        const reader = new FileReader();
        reader.onload = (e) => {
          const markdown = e.target.result;
          const item = document.createElement("li");
          item.textContent = file.name;
          item.dataset.filename = file.name;
          item.dataset.type = "local";
          item.addEventListener("click", () => {
            setActiveFile(item);
            renderAndDisplay(file.name, markdown);
          });
          fileList.appendChild(item);
        };
        reader.readAsText(file);
      });
      event.target.value = "";
    }
  });

  // 从服务器获取 Markdown 文件列表
  fetch("/api/files")
    .then((response) => response.json())
    .then((files) => {
      if (files.length === 0 && fileList.children.length === 0) {
        const item = document.createElement("li");
        item.textContent = "No .md files found on server.";
        fileList.appendChild(item);
        return;
      }
      files.forEach((file) => {
        const item = document.createElement("li");
        item.textContent = file;
        item.dataset.filename = file;
        item.dataset.type = "server";
        item.addEventListener("click", () => {
          setActiveFile(item);
          loadServerFile(file);
        });
        fileList.appendChild(item);
      });
    })
    .catch((error) => {
      console.error("Error fetching file list from server:", error);
      if (fileList.children.length === 0) {
        const item = document.createElement("li");
        item.textContent = "Error loading server files.";
        fileList.appendChild(item);
      }
    });

  // 处理下载按钮点击事件 (使用 FormData)
  downloadBtn.addEventListener("click", () => {
    // 修正检查: 允许空文件，但如果从未加载过任何文件则不允许。
    if (!activeFileElement) {
      alert("Please select a file to download.");
      return;
    }

    const originalText = downloadBtn.textContent;
    downloadBtn.textContent = "Converting...";
    downloadBtn.disabled = true;

    const formData = new FormData();
    const fileType = activeFileElement.dataset.type;

    // Add theme data - 只包含服务器允许的配置项
    const theme = getTheme();
    formData.append('theme', JSON.stringify(theme));

    // 更健壮的逻辑来构建表单数据
    if (fileType === "local") {
      const file = localFileObjects[currentFileName];
      if (file) {
        formData.append("mdfile", file);
      } else {
        alert(`Error: Could not find the local file object for ${currentFileName}`);
        downloadBtn.disabled = false;
        return;
      }
    } else if (fileType === "server") {
      const blob = new Blob([currentMarkdown], { type: "text/markdown" });
      formData.append("mdfile", blob, currentFileName);
    } else {
      alert(`Error: Unknown file type to download.`);
      downloadBtn.disabled = false;
      return;
    }

    const outputFilename = currentFileName.replace(/\.md$/i, "") + ".docx";

    fetch(`/api/convert/docx/${encodeURIComponent(outputFilename)}`, {
      method: "POST",
      body: formData,
    })
      .then(async (response) => {
        if (!response.ok) {
          const errorText = await response.text();
          throw new Error(errorText || "Conversion failed on the server.");
        }
        return response
          .blob()
          .then((blob) => ({ blob, filename: outputFilename }));
      })
      .then(({ blob, filename }) => {
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.style.display = "none";
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
      })
      .catch((error) => {
        console.error("Download error:", error);
        alert(`Could not download the file: ${error.message}`);
      })
      .finally(() => {
        downloadBtn.textContent = originalText;
        downloadBtn.disabled = false;
      });
  });

  // 初始化加载配置
  loadDocxConfig();
});
