document.addEventListener("DOMContentLoaded", () => {
  // --- Constants & Globals ---
  const DOM = {
    fileList: document.getElementById("file-list"),
    content: document.getElementById("content"),
    contentToolbar: document.getElementById("content-toolbar"),
    downloadBtn: document.getElementById("download-docx-btn"),
    localFileInput: document.getElementById("local-file-input"),
    configContainer: document.getElementById("config-container"),
    resetConfigBtn: document.getElementById("reset-config-btn"),
    styleToolbar: document.getElementById("style-settings"),
    toggleRibbonBtn: document.getElementById("toggle-ribbon-btn"),
    manualInputBtn: document.getElementById("manual-input-btn"),
    manualEditor: document.getElementById("manual-editor"),
    contentWrapper: document.querySelector(".content-wrapper"),
  };

  let state = {
    activeFileElement: null,
    currentMarkdown: "",
    currentFileName: "document",
    localFileObjects: {},
    isManualMode: false,
    docxConfig: null,
    currentTheme: {},
    defaultTheme: {},
    activeTab: "colors",
  };

  // --- Helper Functions ---
  const debounce = (func, wait) => {
    let timeout;
    return (...args) => {
      clearTimeout(timeout);
      timeout = setTimeout(() => func(...args), wait);
    };
  };

  const throttle = (func, limit) => {
    let inThrottle;
    return (...args) => {
      if (!inThrottle) {
        func(...args);
        inThrottle = true;
        setTimeout(() => (inThrottle = false), limit);
      }
    };
  };

  const createElement = (
    tag,
    className = "",
    attributes = {},
    children = []
  ) => {
    const el = document.createElement(tag);
    if (className) el.className = className;
    Object.entries(attributes).forEach(([key, value]) => {
      if (key === "textContent") el.textContent = value;
      else if (key.startsWith("on") && typeof value === "function") {
        el.addEventListener(key.substring(2).toLowerCase(), value);
      } else el.setAttribute(key, value);
    });
    children.forEach((child) => {
      if (typeof child === "string")
        el.appendChild(document.createTextNode(child));
      else if (child) el.appendChild(child);
    });
    return el;
  };

  // --- Configuration Management ---
  async function loadDocxConfig() {
    try {
      const response = await fetch("/api/docx-config");
      if (!response.ok) throw new Error("Failed to load config");

      const data = await response.json();
      state.docxConfig = data.config;
      state.defaultTheme = data.defaults;

      const savedTheme = localStorage.getItem("docxTheme");
      state.currentTheme = savedTheme
        ? { ...state.defaultTheme, ...JSON.parse(savedTheme) }
        : { ...state.defaultTheme };

      renderConfigUI();
    } catch (error) {
      console.error("Error loading DOCX config:", error);
      DOM.configContainer.innerHTML =
        '<p style="color: red; font-size: 12px;">加载配置失败</p>';
    }
  }

  function renderConfigUI() {
    if (!state.docxConfig) return;
    DOM.configContainer.innerHTML = "";

    const debouncedUpdate = debounce(updatePreviewStyles, 50);
    const throttledUpdate = throttle(updatePreviewStyles, 30);
    const debouncedSave = debounce(saveTheme, 300);

    const sections = [
      { key: "colors", title: "颜色" },
      { key: "sizes", title: "尺寸" },
      { key: "options", title: "选项" },
    ];

    sections.forEach((section) => {
      const sectionConfig = state.docxConfig[section.key];
      if (!sectionConfig || Object.keys(sectionConfig).length === 0) return;

      const panel = createElement(
        "div",
        "ribbon-panel" + (section.key === state.activeTab ? " active" : ""),
        { "data-tab": section.key }
      );

      groupConfigItems(section.key, sectionConfig).forEach((group) => {
        const itemsDiv = createElement(
          "div",
          "ribbon-group-items",
          {},
          group.items.map(({ key, config }) =>
            createConfigItem(
              key,
              config,
              throttledUpdate,
              debouncedUpdate,
              debouncedSave
            )
          )
        );

        const groupDiv = createElement("div", "ribbon-group", {}, [
          itemsDiv,
          createElement("div", "ribbon-group-title", {
            textContent: group.title,
          }),
        ]);

        panel.appendChild(groupDiv);
      });

      DOM.configContainer.appendChild(panel);
    });

    setupTabListeners();
  }

  function createConfigItem(
    key,
    config,
    throttledUpdate,
    debouncedUpdate,
    debouncedSave
  ) {
    const itemDiv = createElement("div", "ribbon-item");
    let input;

    const commonProps = {
      id: `config-${key}`,
      title: config.label,
    };

    if (config.type === "color") {
      input = createElement("input", "", {
        ...commonProps,
        type: "color",
        value: "#" + (state.currentTheme[key] || config.default),
        onInput: (e) => {
          state.currentTheme[key] = e.target.value
            .replace("#", "")
            .toUpperCase();
          throttledUpdate();
          debouncedSave();
        },
      });
      itemDiv.appendChild(input);
    } else if (config.type === "number") {
      input = createElement("input", "", {
        ...commonProps,
        type: "number",
        min: config.min,
        max: config.max,
        value:
          state.currentTheme[key] !== undefined
            ? state.currentTheme[key]
            : config.default,
        onInput: (e) => {
          let val = parseInt(e.target.value, 10);
          if (!isNaN(val)) {
            if (val < config.min) val = config.min;
            if (val > config.max) val = config.max;
            state.currentTheme[key] = val;
            debouncedUpdate();
            debouncedSave();
          }
        },
        onBlur: (e) => {
          let val = parseInt(e.target.value, 10);
          if (isNaN(val) || val < config.min) val = config.min;
          if (val > config.max) val = config.max;
          e.target.value = val;
          state.currentTheme[key] = val;
          updatePreviewStyles();
          saveTheme();
        },
      });
      itemDiv.appendChild(input);
    } else if (config.type === "boolean") {
      const isChecked =
        state.currentTheme[key] !== undefined
          ? state.currentTheme[key]
          : config.default;
      const checkLabel = createElement("span", "", {
        textContent: isChecked ? "开" : "关",
      });

      input = createElement("input", "", {
        ...commonProps,
        type: "checkbox",
        checked: isChecked ? "true" : undefined, // checked attribute presence
        onChange: (e) => {
          state.currentTheme[key] = e.target.checked;
          checkLabel.textContent = e.target.checked ? "开" : "关";
          // Fix: manually set checked property for sync if needed, though simple DOM handles it
          updatePreviewStyles();
          saveTheme();
        },
      });
      // Manually set checked property because setAttribute doesn't set property for boolean attributes reliably in all contexts
      input.checked = isChecked;

      const wrapper = createElement("div", "checkbox-wrapper", {}, [
        input,
        checkLabel,
      ]);
      itemDiv.appendChild(wrapper);
    }

    itemDiv.appendChild(
      createElement("label", "", {
        for: `config-${key}`,
        textContent: config.label,
      })
    );
    return itemDiv;
  }

  function setupTabListeners() {
    document.querySelectorAll(".ribbon-tab").forEach((tab) => {
      tab.addEventListener("click", () => {
        const tabKey = tab.dataset.tab;
        state.activeTab = tabKey;
        document
          .querySelectorAll(".ribbon-tab")
          .forEach((t) => t.classList.remove("active"));
        tab.classList.add("active");
        document
          .querySelectorAll(".ribbon-panel")
          .forEach((p) => p.classList.remove("active"));
        const panel = document.querySelector(
          `.ribbon-panel[data-tab="${tabKey}"]`
        );
        if (panel) panel.classList.add("active");
      });
    });
  }

  function groupConfigItems(sectionKey, sectionConfig) {
    const items = Object.entries(sectionConfig).map(([key, config]) => ({
      key,
      config,
    }));

    if (sectionKey === "colors") {
      const groups = [
        { prefix: "heading", title: "标题" },
        {
          filter: (k) => ["link", "blockquote", "del"].includes(k),
          title: "文本",
        },
        { filter: (k) => k.toLowerCase().includes("code"), title: "代码" },
        { filter: (k) => k.toLowerCase().includes("table"), title: "表格" },
        {
          filter: (k) =>
            ["border", "hr", "tag", "html", "blockquoteBackground"].includes(k),
          title: "其他",
        },
      ];

      return groups
        .map((g) => ({
          title: g.title,
          items: items.filter((i) =>
            g.prefix ? i.key.startsWith(g.prefix) : g.filter(i.key)
          ),
        }))
        .filter((g) => g.items.length > 0);
    } else if (sectionKey === "sizes") {
      return [
        {
          title: "标题字号",
          items: items.filter((i) => i.key.toLowerCase().includes("heading")),
        },
        {
          title: "其他",
          items: items.filter((i) => !i.key.toLowerCase().includes("heading")),
        },
      ].filter((g) => g.items.length > 0);
    }
    return [{ title: "设置", items }];
  }

  function saveTheme() {
    localStorage.setItem("docxTheme", JSON.stringify(state.currentTheme));
  }

  function resetToDefaults() {
    state.currentTheme = { ...state.defaultTheme };
    localStorage.removeItem("docxTheme");
    renderConfigUI();
    updatePreviewStyles();
  }

  function getTheme() {
    return { ...state.currentTheme };
  }

  function updatePreviewStyles() {
    // 静态样式只注入一次
    const staticStyleId = "static-preview-styles";
    if (!document.getElementById(staticStyleId)) {
      const staticStyle = document.createElement("style");
      staticStyle.id = staticStyleId;
      staticStyle.textContent = `
        .markdown-body h1 { color: var(--h1-color) !important; font-size: var(--h1-size) !important; }
        .markdown-body h2 { color: var(--h2-color) !important; font-size: var(--h2-size) !important; }
        .markdown-body h3 { color: var(--h3-color) !important; font-size: var(--h3-size) !important; }
        .markdown-body h4 { color: var(--h4-color) !important; font-size: var(--h4-size) !important; }
        .markdown-body h5 { color: var(--h5-color) !important; font-size: var(--h5-size) !important; }
        .markdown-body h6 { color: var(--h6-color) !important; font-size: var(--h6-size) !important; }
        .markdown-body a { color: var(--link-color) !important; text-decoration: var(--link-decoration) !important; }
        .markdown-body pre code { color: var(--code-color) !important; font-size: var(--code-size) !important; background-color: transparent !important; }
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

    const root = document.documentElement;
    const { currentTheme } = state;

    // Helper to set CSS variable
    const setVar = (name, value) => root.style.setProperty(name, value);
    const hex = (key, def) => "#" + (currentTheme[key] || def);
    const pt = (key, def) => (currentTheme[key] || def) / 2 + "pt";

    // Color mappings [cssVar, themeKey, defaultVal]
    const colors = [
      ["--h1-color", "heading1", "2F5597"],
      ["--h2-color", "heading2", "5B9BD5"],
      ["--h3-color", "heading3", "44546A"],
      ["--h4-color", "heading4", "44546A"],
      ["--h5-color", "heading5", "44546A"],
      ["--h6-color", "heading6", "44546A"],
      ["--link-color", "link", "0563C1"],
      ["--code-color", "code", "032F62"],
      ["--codespan-color", "codespan", "70AD47"],
      ["--code-bg", "codeBackground", "F6F6F7"],
      ["--blockquote-color", "blockquote", "666666"],
      ["--blockquote-bg", "blockquoteBackground", "F9F9F9"],
      ["--border-color", "border", "A5A5A5"],
      ["--hr-color", "hr", "D9D9D9"],
      ["--del-color", "del", "FF0000"],
      ["--table-header-bg", "tableHeaderBackground", "F2F2F2"],
      ["--table-header-text", "tableHeaderTextColor", "000000"],
      ["--table-body-bg", "tableBodyBackground", "FFFFFF"],
    ];

    colors.forEach(([v, k, d]) => setVar(v, hex(k, d)));

    // Special logic values
    setVar(
      "--link-decoration",
      currentTheme.linkUnderline === false ? "none" : "underline"
    );
    setVar(
      "--table-alt-row-bg",
      currentTheme.tableAltRowEnabled
        ? hex("tableAltRowBackground", "F9F9F9")
        : "var(--table-body-bg)"
    );

    // Size mappings
    const sizes = [
      ["--h1-size", "heading1Size", 36],
      ["--h2-size", "heading2Size", 32],
      ["--h3-size", "heading3Size", 28],
      ["--h4-size", "heading4Size", 26],
      ["--h5-size", "heading5Size", 24],
      ["--h6-size", "heading6Size", 24],
      ["--code-size", "codeSize", 22],
      ["--table-size", "tableSize", 21],
    ];

    sizes.forEach(([v, k, d]) => setVar(v, pt(k, d)));

    // Other sizes
    setVar(
      "--table-border-width",
      (currentTheme.tableBorderWidth || 4) / 2 + "px"
    );
    setVar(
      "--table-cell-padding",
      (currentTheme.tableCellPadding || 80) / 20 + "px"
    );
    setVar(
      "--table-header-weight",
      currentTheme.tableHeaderBold !== false ? "bold" : "normal"
    );

    const centerAlign = currentTheme.tableCenterAlign !== false;
    setVar("--table-margin-left", centerAlign ? "auto" : "0");
    setVar("--table-margin-right", centerAlign ? "auto" : "0");
  }

  // ======= Manual Mode & UI Interaction =======

  function toggleManualMode(active) {
    state.isManualMode = active;
    if (active) {
      if (state.activeFileElement) {
        state.activeFileElement.classList.remove("active");
        state.activeFileElement = null;
      }
      DOM.manualInputBtn.classList.add("active");
      DOM.manualEditor.style.display = "block";
      DOM.contentWrapper.classList.add("split-view");
      state.currentFileName = "manual-input";

      if (!DOM.manualEditor.value) {
        DOM.manualEditor.value = "# Manual Input\n\nStart typing here...";
      }
      state.currentMarkdown = DOM.manualEditor.value;
      renderAndDisplay("manual-input", state.currentMarkdown);
    } else {
      DOM.manualInputBtn.classList.remove("active");
      DOM.manualEditor.style.display = "none";
      DOM.contentWrapper.classList.remove("split-view");
    }
  }

  DOM.manualInputBtn.addEventListener(
    "click",
    () => !state.isManualMode && toggleManualMode(true)
  );

  DOM.manualEditor.addEventListener(
    "input",
    debounce((e) => {
      state.currentMarkdown = e.target.value;
      renderAndDisplay("manual-input", state.currentMarkdown);
    }, 300)
  );

  DOM.resetConfigBtn.addEventListener("click", () => {
    if (confirm("确定要将所有样式设置重置为默认值吗？")) resetToDefaults();
  });

  DOM.toggleRibbonBtn.addEventListener("click", () => {
    DOM.styleToolbar.classList.toggle("collapsed");
    localStorage.setItem(
      "ribbonCollapsed",
      DOM.styleToolbar.classList.contains("collapsed")
    );
  });

  if (localStorage.getItem("ribbonCollapsed") === "true") {
    DOM.styleToolbar.classList.add("collapsed");
  }

  // ======= File Management =======

  function renderAndDisplay(filename, markdown) {
    state.currentMarkdown = markdown;
    state.currentFileName = filename;
    DOM.content.innerHTML = "Rendering...";

    fetch("/api/render", {
      method: "POST",
      headers: { "Content-Type": "text/plain" },
      body: markdown,
    })
      .then((res) => {
        if (!res.ok) throw new Error("Server rendering failed");
        return res.text();
      })
      .then((html) => {
        DOM.content.innerHTML = html;
        DOM.contentToolbar.style.display = "flex";
        updatePreviewStyles();
      })
      .catch((error) => {
        console.error("Error rendering content:", error);
        DOM.content.innerHTML = `<p style="color: red;">Error rendering content.</p>`;
      });
  }

  function setActiveFile(item) {
    if (state.isManualMode) toggleManualMode(false);
    if (state.activeFileElement)
      state.activeFileElement.classList.remove("active");
    state.activeFileElement = item;
    state.activeFileElement.classList.add("active");
  }

  function loadServerFile(filename) {
    DOM.content.innerHTML = "Loading...";
    DOM.contentToolbar.style.display = "none";
    fetch(`/api/files/${encodeURIComponent(filename)}`)
      .then((res) => {
        if (!res.ok) throw new Error("Network response was not ok");
        return res.text();
      })
      .then((md) => renderAndDisplay(filename, md))
      .catch((error) => {
        console.error("Error fetching file:", error);
        DOM.content.innerHTML = `<p style="color: red;">Error loading file: ${filename}</p>`;
        state.currentMarkdown = "";
        state.currentFileName = "document";
      });
  }

  DOM.localFileInput.addEventListener("change", (event) => {
    const files = event.target.files;
    if (files.length > 0) {
      Array.from(files).forEach((file) => {
        state.localFileObjects[file.name] = file;
        const reader = new FileReader();
        reader.onload = (e) => {
          const item = createElement("li", "", {
            textContent: file.name,
            title: file.name,
            "data-filename": file.name,
            "data-type": "local",
            onClick: () => {
              setActiveFile(item);
              renderAndDisplay(file.name, e.target.result);
            },
          });
          DOM.fileList.appendChild(item);
        };
        reader.readAsText(file);
      });
      event.target.value = "";
    }
  });

  // Load server files
  fetch("/api/files")
    .then((res) => res.json())
    .then((files) => {
      if (files.length === 0 && DOM.fileList.children.length === 0) {
        DOM.fileList.appendChild(
          createElement("li", "", {
            textContent: "No .md files found on server.",
          })
        );
        return;
      }
      files.forEach((file) => {
        const item = createElement("li", "", {
          textContent: file,
          title: file, // Added title attribute
          "data-filename": file,
          "data-type": "server",
          onClick: () => {
            setActiveFile(item);
            loadServerFile(file);
          },
        });
        DOM.fileList.appendChild(item);
      });
    })
    .catch((error) => {
      console.error("Error fetching file list:", error);
      if (DOM.fileList.children.length === 0) {
        DOM.fileList.appendChild(
          createElement("li", "", {
            textContent: "Error loading server files.",
          })
        );
      }
    });

  DOM.downloadBtn.addEventListener("click", () => {
    if (!state.activeFileElement && !state.isManualMode) {
      alert("Please select a file to download.");
      return;
    }

    const originalText = DOM.downloadBtn.textContent;
    DOM.downloadBtn.textContent = "Converting...";
    DOM.downloadBtn.disabled = true;

    const formData = new FormData();
    const fileType = state.isManualMode
      ? "manual"
      : state.activeFileElement.dataset.type;

    formData.append("theme", JSON.stringify(getTheme()));

    if (fileType === "local") {
      const file = state.localFileObjects[state.currentFileName];
      if (file) formData.append("mdfile", file);
      else {
        alert(
          `Error: Could not find the local file object for ${state.currentFileName}`
        );
        DOM.downloadBtn.disabled = false;
        return;
      }
    } else {
      let filenameToSend = state.currentFileName;
      if (!filenameToSend.toLowerCase().endsWith(".md"))
        filenameToSend += ".md";
      const blob = new Blob([state.currentMarkdown], { type: "text/markdown" });
      formData.append("mdfile", blob, filenameToSend);
    }

    let outputFilename = state.currentFileName;
    if (outputFilename.toLowerCase().endsWith(".md")) {
      outputFilename = outputFilename.replace(/\.md$/i, "") + ".docx";
    } else {
      outputFilename += ".docx";
    }

    fetch(`/api/convert/docx/${encodeURIComponent(outputFilename)}`, {
      method: "POST",
      body: formData,
    })
      .then(async (res) => {
        if (!res.ok) throw new Error((await res.text()) || "Conversion failed");
        return res.blob().then((blob) => ({ blob, filename: outputFilename }));
      })
      .then(({ blob, filename }) => {
        const url = window.URL.createObjectURL(blob);
        const a = createElement("a", "", {
          href: url,
          download: filename,
          style: "display:none",
        });
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
        DOM.downloadBtn.textContent = originalText;
        DOM.downloadBtn.disabled = false;
      });
  });

  loadDocxConfig();

  const logoImg = document.getElementById("logoImg");
  const styleElem = document.createElement("style");
  styleElem.innerHTML = `
    #logo-background {
      stroke-dasharray: 360;
      stroke-dashoffset: 360;
      animation: draw-main 0.8s 0.1s ease-out forwards;
    }

    .code-symbol {
      opacity: 0;
      animation: fade-in 0.8s 0.1s ease-out forwards;
    }

    .line-anim {
      transform: scaleX(0);
      transform-origin: center;
      animation: extend 0.8s ease-out forwards;
    }

    .line-anim:nth-child(1) { animation-delay: 0.1s; }
    .line-anim:nth-child(2) { animation-delay: 0.1.5s; }
    .line-anim:nth-child(3) { animation-delay: 0.2s; }
    .line-anim:nth-child(4) { animation-delay: 0.2.5s; }
    .line-anim:nth-child(5) { animation-delay: 0.3s; }

    @keyframes draw-main { to { stroke-dashoffset: 0; } }
    @keyframes fade-in { to { opacity: 1; } }
    @keyframes extend { to { transform: scaleX(1); } }
  `;

  fetch("./icon.svg")
    .then((res) => res.text())
    .then((svgText) => {
      const parser = new DOMParser();
      const svgDoc = parser.parseFromString(svgText, "image/svg+xml");
      const svgElem = svgDoc.documentElement;
      svgElem.appendChild(styleElem);
      const serializer = new XMLSerializer();
      const modifiedSvgText = serializer.serializeToString(svgElem);
      logoImg.src = URL.createObjectURL(
        new Blob([modifiedSvgText], { type: "image/svg+xml" })
      );
    });
});
