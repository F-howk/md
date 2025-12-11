
document.addEventListener('DOMContentLoaded', () => {
    const fileList = document.getElementById('file-list');
    const content = document.getElementById('content');
    const contentToolbar = document.getElementById('content-toolbar');
    const downloadBtn = document.getElementById('download-docx-btn');
    const localFileInput = document.getElementById('local-file-input');
    let activeFileElement = null;
    let currentMarkdown = '';
    let currentFileName = 'document';
    let localFileObjects = {}; // 新增: 存储本地文件对象

    // 用于通过服务器渲染 Markdown 并显示它的函数
    function renderAndDisplay(filename, markdown) {
        currentMarkdown = markdown;
        currentFileName = filename;
        content.innerHTML = 'Rendering...';

        fetch('/api/render', {
            method: 'POST',
            headers: {
                'Content-Type': 'text/plain',
            },
            body: markdown,
        })
        .then(response => {
            if (!response.ok) {
                throw new Error('Server rendering failed');
            }
            return response.text();
        })
        .then(html => {
            content.innerHTML = html;
            contentToolbar.style.display = 'flex';
        })
        .catch(error => {
            console.error('Error rendering content:', error);
            content.innerHTML = `<p style="color: red;">Error rendering content.</p>`;
        });
    }

    // 处理文件列表项的激活状态
    function setActiveFile(item) {
        if (activeFileElement) {
            activeFileElement.classList.remove('active');
        }
        activeFileElement = item;
        activeFileElement.classList.add('active');
    }

    // 从服务器端 Markdown 文件加载并显示内容
    function loadServerFile(filename) {
        content.innerHTML = 'Loading...';
        contentToolbar.style.display = 'none';
        fetch(`/api/files/${encodeURIComponent(filename)}`)
            .then(response => {
                if (!response.ok) {
                    throw new Error('Network response was not ok');
                }
                return response.text();
            })
            .then(markdown => {
                renderAndDisplay(filename, markdown);
            })
            .catch(error => {
                console.error('Error fetching file content from server:', error);
                content.innerHTML = `<p style="color: red;">Error loading file from server: ${filename}</p>`;
                currentMarkdown = '';
                currentFileName = 'document';
            });
    }

    // 处理本地文件选择和显示
    localFileInput.addEventListener('change', (event) => {
        const files = event.target.files;
        if (files.length > 0) {
            Array.from(files).forEach(file => {
                localFileObjects[file.name] = file; // Store the File object
                const reader = new FileReader();
                reader.onload = (e) => {
                    const markdown = e.target.result;
                    const item = document.createElement('li');
                    item.textContent = file.name;
                    item.dataset.filename = file.name;
                    item.dataset.type = 'local';
                    item.addEventListener('click', () => {
                        setActiveFile(item);
                        renderAndDisplay(file.name, markdown);
                    });
                    fileList.appendChild(item);
                };
                reader.readAsText(file);
            });
            event.target.value = '';
        }
    });

    // 从服务器获取 Markdown 文件列表
    fetch('/api/files')
        .then(response => response.json())
        .then(files => {
            if (files.length === 0 && fileList.children.length === 0) {
                const item = document.createElement('li');
                item.textContent = 'No .md files found on server.';
                fileList.appendChild(item);
                return;
            }
            files.forEach(file => {
                const item = document.createElement('li');
                item.textContent = file;
                item.dataset.filename = file;
                item.dataset.type = 'server';
                item.addEventListener('click', () => {
                    setActiveFile(item);
                    loadServerFile(file);
                });
                fileList.appendChild(item);
            });
        })
        .catch(error => {
            console.error('Error fetching file list from server:', error);
            if (fileList.children.length === 0) {
                const item = document.createElement('li');
                item.textContent = 'Error loading server files.';
                fileList.appendChild(item);
            }
        });

    // 处理下载按钮点击事件 (使用 FormData)
    downloadBtn.addEventListener('click', () => {
        // 修正检查: 允许空文件，但如果从未加载过任何文件则不允许。
        if (!activeFileElement) {
            alert('Please select a file to download.');
            return;
        }

        downloadBtn.textContent = 'Converting...';
        downloadBtn.disabled = true;

        const formData = new FormData();
        const fileType = activeFileElement.dataset.type;

        // 更健壮的逻辑来构建表单数据
        if (fileType === 'local') {
            const file = localFileObjects[currentFileName];
            if (file) {
                formData.append('mdfile', file);
            } else {
                alert(`Error: Could not find the local file object for ${currentFileName}`);
                downloadBtn.textContent = 'Download as .docx';
                downloadBtn.disabled = false;
                return;
            }
        } else if (fileType === 'server') {
            const blob = new Blob([currentMarkdown], { type: 'text/markdown' });
            formData.append('mdfile', blob, currentFileName);
        } else {
            alert(`Error: Unknown file type to download.`);
            downloadBtn.textContent = 'Download as .docx';
            downloadBtn.disabled = false;
            return;
        }

        const outputFilename = currentFileName.replace(/\.md$/i, '') + '.docx';

        fetch(`/api/convert/docx/${encodeURIComponent(outputFilename)}`, {
            method: 'POST',
            body: formData, // 浏览器将正确设置 Content-Type 头部
        })
        .then(async response => {
            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(errorText || 'Conversion failed on the server.');
            }
            return response.blob().then(blob => ({ blob, filename: outputFilename }));
        })
        .then(({ blob, filename }) => {
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
        })
        .catch(error => {
            console.error('Download error:', error);
            alert(`Could not download the file: ${error.message}`);
        })
        .finally(() => {
            downloadBtn.textContent = 'Download as .docx';
            downloadBtn.disabled = false;
        });
    });
});
