/**
 * 文档转换工具 Pro v2.0
 * 支持 Word、PDF、PPT 格式互转
 * 纯前端实现，保护隐私
 */

class DocConverter {
    constructor() {
        this.currentFile = null;
        this.currentFormat = null;
        this.currentDownloadUrl = null;
        this.init();
    }

    init() {
        this.bindEvents();
        this.configurePDFjs();
    }

    // 配置 PDF.js
    configurePDFjs() {
        if (typeof pdfjsLib !== 'undefined') {
            pdfjsLib.GlobalWorkerOptions.workerSrc = 
                'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
            console.log('✅ PDF.js 配置成功');
        } else {
            console.warn('⚠️ PDF.js 库加载失败');
        }
    }

    // 绑定事件
    bindEvents() {
        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');
        const convertBtn = document.getElementById('convertBtn');
        const removeFileBtn = document.getElementById('removeFile');

        // 点击上传 - 支持桌面和移动端
        const handleClick = (e) => {
            e.preventDefault();
            e.stopPropagation();
            fileInput.click();
        };
        
        dropZone.addEventListener('click', handleClick);
        dropZone.addEventListener('touchend', handleClick);
        
        // 防止移动端双击缩放
        let lastTouchEnd = 0;
        dropZone.addEventListener('touchend', (e) => {
            const now = Date.now();
            if (now - lastTouchEnd <= 300) {
                e.preventDefault();
            }
            lastTouchEnd = now;
        }, false);

        // 文件选择
        fileInput.addEventListener('change', (e) => {
            if (e.target.files[0]) {
                this.handleFileSelect(e.target.files[0]);
            }
        });

        // 拖拽上传
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
            });
        });

        ['dragenter', 'dragover'].forEach(eventName => {
            dropZone.addEventListener(eventName, () => {
                dropZone.classList.add('dragover');
            });
        });

        ['dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, () => {
                dropZone.classList.remove('dragover');
            });
        });

        dropZone.addEventListener('drop', (e) => {
            const file = e.dataTransfer.files[0];
            if (file) this.handleFileSelect(file);
        });

        // 转换按钮
        convertBtn.addEventListener('click', () => this.startConversion());

        // 删除文件
        removeFileBtn.addEventListener('click', () => this.clearFile());
    }

    // 处理文件选择
    handleFileSelect(file) {
        // 验证文件
        const validation = this.validateFile(file);
        if (!validation.valid) {
            this.showError(validation.message);
            return;
        }

        this.currentFile = file;
        this.showFileInfo(file);
        this.showFormatOptions(file.name);
        this.hideError();
    }

    // 验证文件
    validateFile(file) {
        const maxSize = 50 * 1024 * 1024; // 50MB
        const validTypes = ['.docx', '.pdf', '.pptx'];
        const ext = '.' + file.name.split('.').pop().toLowerCase();

        if (!file) {
            return { valid: false, message: '请选择文件' };
        }

        if (file.size === 0) {
            return { valid: false, message: '文件为空，请选择其他文件' };
        }

        if (file.size > maxSize) {
            return { valid: false, message: '文件太大！最大支持 50MB' };
        }

        if (!validTypes.includes(ext)) {
            return { valid: false, message: '不支持的格式！请上传 .docx, .pdf 或 .pptx 文件' };
        }

        return { valid: true };
    }

    // 显示文件信息
    showFileInfo(file) {
        const fileInfo = document.getElementById('fileInfo');
        const fileName = document.getElementById('fileName');
        const fileMeta = document.getElementById('fileMeta');

        fileName.textContent = file.name;
        fileMeta.textContent = `${this.formatSize(file.size)} · ${file.type || '未知类型'}`;
        fileInfo.classList.add('show');
    }

    // 格式化文件大小
    formatSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    // 显示格式选项
    showFormatOptions(filename) {
        const formatSection = document.getElementById('formatSection');
        const formatOptions = document.getElementById('formatOptions');
        const convertBtn = document.getElementById('convertBtn');
        const ext = '.' + filename.split('.').pop().toLowerCase();

        const formats = {
            '.docx': [
                { value: 'pdf', label: 'PDF', icon: '📄', desc: '保留格式' },
                { value: 'html', label: 'HTML', icon: '🌐', desc: '网页格式' },
                { value: 'txt', label: '纯文本', icon: '📝', desc: '提取文字' }
            ],
            '.pdf': [
                { value: 'docx', label: 'Word', icon: '📘', desc: '可编辑' },
                { value: 'html', label: 'HTML', icon: '🌐', desc: '网页格式' },
                { value: 'txt', label: '纯文本', icon: '📝', desc: '提取文字' }
            ],
            '.pptx': [
                { value: 'pdf', label: 'PDF', icon: '📄', desc: '幻灯片PDF' },
                { value: 'images', label: '图片集', icon: '🖼️', desc: '提取图片' }
            ]
        };

        const options = formats[ext] || [];
        formatOptions.innerHTML = options.map(opt => `
            <button class="format-btn" data-format="${opt.value}">
                <span>${opt.icon}</span>
                <span>${opt.label}</span>
            </button>
        `).join('');

        // 绑定选择事件
        formatOptions.querySelectorAll('.format-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                formatOptions.querySelectorAll('.format-btn').forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                this.currentFormat = btn.dataset.format;
                convertBtn.disabled = false;
                convertBtn.textContent = `🚀 转换为 ${btn.querySelector('span:last-child').textContent}`;
            });
        });

        formatSection.classList.add('show');
        convertBtn.disabled = true;
        convertBtn.textContent = '🚀 开始转换';
    }

    // 清空文件
    clearFile() {
        this.currentFile = null;
        this.currentFormat = null;
        
        document.getElementById('fileInfo').classList.remove('show');
        document.getElementById('formatSection').classList.remove('show');
        document.getElementById('convertBtn').disabled = true;
        document.getElementById('convertBtn').textContent = '🚀 开始转换';
        document.getElementById('fileInput').value = '';
        
        this.hideResult();
        this.hideError();
    }

    // 开始转换
    async startConversion() {
        if (!this.currentFile || !this.currentFormat) {
            this.showError('请先选择文件和目标格式');
            return;
        }

        const progressContainer = document.getElementById('progressContainer');
        const convertBtn = document.getElementById('convertBtn');
        const resultSection = document.getElementById('resultSection');

        // 隐藏之前的结果
        resultSection.classList.remove('show');
        this.hideError();
        
        // 显示进度
        progressContainer.classList.add('show');
        convertBtn.disabled = true;

        try {
            this.updateProgress(10, '正在读取文件...');
            await this.sleep(300);
            
            this.updateProgress(30, '正在解析内容...');
            const arrayBuffer = await this.fileToArrayBuffer(this.currentFile);
            
            this.updateProgress(50, '正在转换格式...');
            const result = await this.convertFile(arrayBuffer);
            
            this.updateProgress(90, '正在生成文件...');
            await this.sleep(300);
            
            this.updateProgress(100, '完成！');
            this.showResult(result);
            
        } catch (error) {
            console.error('转换失败:', error);
            this.showError(this.getErrorMessage(error));
        } finally {
            progressContainer.classList.remove('show');
            convertBtn.disabled = false;
            convertBtn.textContent = '🚀 开始转换';
            this.updateProgress(0, '准备中...');
        }
    }

    // 更新进度
    updateProgress(percent, text) {
        document.getElementById('progressFill').style.width = percent + '%';
        document.getElementById('progressText').textContent = text;
        document.getElementById('progressPercent').textContent = percent + '%';
    }

    // 转换文件
    async convertFile(arrayBuffer) {
        const ext = '.' + this.currentFile.name.split('.').pop().toLowerCase();

        switch (ext) {
            case '.docx':
                return await this.convertDocx(arrayBuffer, this.currentFormat);
            case '.pdf':
                return await this.convertPdf(arrayBuffer, this.currentFormat);
            case '.pptx':
                return await this.convertPptx(arrayBuffer, this.currentFormat);
            default:
                throw new Error('不支持的文件格式');
        }
    }

    // Word 转换
    async convertDocx(arrayBuffer, targetFormat) {
        switch (targetFormat) {
            case 'pdf':
                return await this.docxToPdf(arrayBuffer);
            case 'html':
                return await this.docxToHtml(arrayBuffer);
            case 'txt':
                return await this.docxToTxt(arrayBuffer);
            default:
                throw new Error('不支持的转换格式');
        }
    }

    // PDF 转换
    async convertPdf(arrayBuffer, targetFormat) {
        switch (targetFormat) {
            case 'docx':
                return await this.pdfToDocx(arrayBuffer);
            case 'html':
                return await this.pdfToHtml(arrayBuffer);
            case 'txt':
                return await this.pdfToTxt(arrayBuffer);
            default:
                throw new Error('不支持的转换格式');
        }
    }

    // PPT 转换
    async convertPptx(arrayBuffer, targetFormat) {
        switch (targetFormat) {
            case 'pdf':
                return await this.pptxToPdf(arrayBuffer);
            case 'images':
                return await this.pptxToImages(arrayBuffer);
            default:
                throw new Error('不支持的转换格式');
        }
    }

    // Word → PDF
    async docxToPdf(arrayBuffer) {
        // 检查 mammoth 是否可用
        if (typeof mammoth === 'undefined') {
            throw new Error('Mammoth.js 库未加载，请检查网络连接');
        }

        try {
            const result = await mammoth.convertToHtml({ arrayBuffer });

            // 检查转换结果
            if (result.messages && result.messages.length > 0) {
                console.log('Mammoth 转换消息:', result.messages);
            }

            const html = this.createHtmlDocument(result.value, 'Word转PDF');

            const element = document.createElement('div');
            element.innerHTML = html;
            element.style.position = 'absolute';
            element.style.left = '-9999px';
            element.style.width = '210mm';
            document.body.appendChild(element);

            const opt = {
                margin: [15, 15, 15, 15],
                filename: this.getOutputFilename('pdf'),
                image: { type: 'jpeg', quality: 0.98 },
                html2canvas: {
                    scale: 2,
                    useCORS: true,
                    logging: false,
                    letterRendering: true
                },
                jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
            };

            // 使用 Promise 包装 html2pdf
            return new Promise((resolve, reject) => {
                html2pdf()
                    .set(opt)
                    .from(element)
                    .toPdf()
                    .get('pdf')
                    .then((pdf) => {
                        document.body.removeChild(element);
                        const blob = pdf.output('blob');
                        resolve({ blob, filename: opt.filename, type: 'application/pdf' });
                    })
                    .catch((err) => {
                        if (document.body.contains(element)) {
                            document.body.removeChild(element);
                        }
                        reject(new Error('PDF生成失败: ' + err.message));
                    });
            });
        } catch (error) {
            throw new Error('Word 转换失败: ' + error.message);
        }
    }

    // Word → HTML
    async docxToHtml(arrayBuffer) {
        if (typeof mammoth === 'undefined') {
            throw new Error('Mammoth.js 库未加载，请检查网络连接');
        }

        try {
            const result = await mammoth.convertToHtml({ arrayBuffer });
            const html = this.createHtmlDocument(result.value, 'Word转HTML', true);
            const blob = new Blob([html], { type: 'text/html; charset=utf-8' });
            return { blob, filename: this.getOutputFilename('html'), type: 'text/html' };
        } catch (error) {
            throw new Error('HTML 转换失败: ' + error.message);
        }
    }

    // Word → TXT
    async docxToTxt(arrayBuffer) {
        if (typeof mammoth === 'undefined') {
            throw new Error('Mammoth.js 库未加载，请检查网络连接');
        }

        try {
            const result = await mammoth.extractRawText({ arrayBuffer });
            // 清理文本：移除多余空行
            const cleanText = result.value
                .replace(/\n{3,}/g, '\n\n')
                .trim();
            const blob = new Blob([cleanText], { type: 'text/plain; charset=utf-8' });
            return { blob, filename: this.getOutputFilename('txt'), type: 'text/plain' };
        } catch (error) {
            throw new Error('文本提取失败: ' + error.message);
        }
    }

    // PDF → TXT
    async pdfToTxt(arrayBuffer) {
        const { text, numPages } = await this.extractPdfText(arrayBuffer);
        const header = `PDF文档转换结果\n================\n页数: ${numPages}\n\n`;
        const blob = new Blob([header + text], { type: 'text/plain; charset=utf-8' });
        return { blob, filename: this.getOutputFilename('txt'), type: 'text/plain' };
    }

    // PDF → HTML
    async pdfToHtml(arrayBuffer) {
        const { text, numPages } = await this.extractPdfText(arrayBuffer);
        const html = this.createPdfHtmlDocument(text, numPages);
        const blob = new Blob([html], { type: 'text/html; charset=utf-8' });
        return { blob, filename: this.getOutputFilename('html'), type: 'text/html' };
    }

    // PDF → Word
    async pdfToDocx(arrayBuffer) {
        const { text, numPages } = await this.extractPdfText(arrayBuffer);
        const html = this.createWordDocument(text, numPages);
        const blob = new Blob([html], { type: 'application/msword; charset=utf-8' });
        return { blob, filename: this.getOutputFilename('doc'), type: 'application/msword' };
    }

    // PPT → PDF
    async pptxToPdf(arrayBuffer) {
        const pptxInfo = await this.extractPptxContent(arrayBuffer);
        const html = this.createPptxHtmlDocument(pptxInfo);

        const element = document.createElement('div');
        element.innerHTML = html;
        element.style.position = 'absolute';
        element.style.left = '-9999px';
        element.style.width = '297mm'; // A4 横向宽度
        document.body.appendChild(element);

        const opt = {
            margin: [10, 10, 10, 10],
            filename: this.getOutputFilename('pdf'),
            image: { type: 'jpeg', quality: 0.95 },
            html2canvas: {
                scale: 2,
                useCORS: true,
                logging: false,
                letterRendering: true
            },
            jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape' }
        };

        return new Promise((resolve, reject) => {
            html2pdf()
                .set(opt)
                .from(element)
                .toPdf()
                .get('pdf')
                .then((pdf) => {
                    document.body.removeChild(element);
                    const blob = pdf.output('blob');
                    resolve({ blob, filename: opt.filename, type: 'application/pdf' });
                })
                .catch((err) => {
                    if (document.body.contains(element)) {
                        document.body.removeChild(element);
                    }
                    reject(new Error('PPT转PDF失败: ' + err.message));
                });
        });
    }

    // PPT → Images
    async pptxToImages(arrayBuffer) {
        const pptxInfo = await this.extractPptxContent(arrayBuffer);
        const html = this.createPptxImagesDocument(pptxInfo);
        const blob = new Blob([html], { type: 'text/html; charset=utf-8' });
        return { blob, filename: this.getOutputFilename('html'), type: 'text/html' };
    }

    // 提取 PDF 文本
    async extractPdfText(arrayBuffer) {
        if (typeof pdfjsLib === 'undefined') {
            throw new Error('PDF.js 库未加载，请检查网络连接');
        }

        let pdf = null;

        try {
            const loadingTask = pdfjsLib.getDocument({
                data: arrayBuffer,
                useSystemFonts: true,
                cMapUrl: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/cmaps/',
                cMapPacked: true
            });

            pdf = await Promise.race([
                loadingTask.promise,
                new Promise((_, reject) =>
                    setTimeout(() => reject(new Error('PDF 加载超时，文件可能太大')), 30000)
                )
            ]);

            let fullText = '';
            const maxPages = Math.min(pdf.numPages, 100);

            for (let i = 1; i <= maxPages; i++) {
                try {
                    const page = await pdf.getPage(i);
                    const textContent = await page.getTextContent();

                    let pageText = '';
                    let lastY = null;
                    let lastX = null;

                    for (const item of textContent.items) {
                        if (item.str && item.str.trim()) {
                            // 检测换行（Y坐标变化较大）
                            if (lastY !== null && Math.abs(item.transform[5] - lastY) > 5) {
                                pageText += '\n';
                            }
                            // 检测空格（X坐标变化较大）
                            else if (lastX !== null && item.transform[4] - lastX > 10) {
                                pageText += ' ';
                            }

                            pageText += item.str;
                            lastY = item.transform[5];
                            lastX = item.transform[4] + (item.width || 0);
                        }
                    }

                    if (pageText.trim()) {
                        fullText += `\n--- 第 ${i} 页 ---\n${pageText.trim()}\n`;
                    }

                    // 清理页面资源
                    page.cleanup();

                    // 每10页更新一次进度
                    if (i % 10 === 0) {
                        this.updateProgress(30 + Math.floor((i / maxPages) * 20), `正在解析第 ${i}/${maxPages} 页...`);
                    }
                } catch (pageError) {
                    console.warn(`第 ${i} 页解析失败:`, pageError);
                    fullText += `\n--- 第 ${i} 页 ---\n(解析失败)\n`;
                }
            }

            // 销毁 PDF 文档
            if (pdf && pdf.destroy) {
                pdf.destroy();
            }

            return { text: fullText || '(无文本内容)', numPages: pdf.numPages };
        } catch (error) {
            // 清理资源
            if (pdf && pdf.destroy) {
                try { pdf.destroy(); } catch (e) {}
            }

            if (error.message.includes('Invalid')) {
                throw new Error('PDF 文件已损坏或格式无效');
            } else if (error.message.includes('password')) {
                throw new Error('PDF 文件已加密，无法转换');
            }
            throw error;
        }
    }

    // 提取 PPTX 内容
    async extractPptxContent(arrayBuffer) {
        if (typeof JSZip === 'undefined') {
            throw new Error('JSZip库未加载');
        }

        try {
            const zip = await JSZip.loadAsync(arrayBuffer);
            const slides = Object.keys(zip.files).filter(name =>
                name.startsWith('ppt/slides/slide') && name.endsWith('.xml')
            ).sort((a, b) => {
                // 按幻灯片编号排序
                const numA = parseInt(a.match(/slide(\d+)\.xml$/)?.[1] || 0);
                const numB = parseInt(b.match(/slide(\d+)\.xml$/)?.[1] || 0);
                return numA - numB;
            });

            if (slides.length === 0) {
                throw new Error('无法找到幻灯片内容');
            }

            let content = [];
            const maxSlides = Math.min(slides.length, 50); // 最多处理50张幻灯片

            for (let i = 0; i < maxSlides; i++) {
                try {
                    const slideFile = zip.file(slides[i]);
                    if (!slideFile) continue;

                    const slideContent = await slideFile.async('text');
                    // 提取所有文本节点
                    const texts = slideContent.match(/<a:t>([^<]*)<\/a:t>/g) || [];
                    const slideText = texts
                        .map(t => t.replace(/<\/?a:t>/g, ''))
                        .filter(t => t.trim())
                        .join(' ');

                    content.push({
                        slide: i + 1,
                        text: slideText.substring(0, 500) + (slideText.length > 500 ? '...' : '') || '(无文本内容)'
                    });
                } catch (slideError) {
                    console.warn(`幻灯片 ${i + 1} 解析失败:`, slideError);
                    content.push({
                        slide: i + 1,
                        text: '(解析失败)'
                    });
                }
            }

            return { slideCount: slides.length, content };
        } catch (error) {
            throw new Error('PPT 解析失败: ' + error.message);
        }
    }

    // 创建 HTML 文档
    createHtmlDocument(content, title, isWeb = false) {
        const fontFamily = '"Noto Sans SC", "Microsoft YaHei", "PingFang SC", Arial, sans-serif';

        // 清理 mammoth 生成的 HTML 中的潜在问题
        const cleanContent = content
            // 移除空的 style 属性
            .replace(/style=""/g, '')
            // 确保图片有 alt 属性
            .replace(/<img([^>]*)>/g, (match, attrs) => {
                if (!attrs.includes('alt=')) {
                    return `<img${attrs} alt="">`;
                }
                return match;
            })
            // 修复可能的自闭合标签问题
            .replace(/<br>/g, '<br/>')
            .replace(/<hr>/g, '<hr/>');

        return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${title}</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@400;700&display=swap');
        body {
            font-family: ${fontFamily};
            padding: ${isWeb ? '40px 20px' : '40px'};
            line-height: 1.8;
            color: #333;
            font-size: 11pt;
            ${isWeb ? 'max-width: 800px; margin: 0 auto;' : ''}
            word-wrap: break-word;
        }
        h1, h2, h3, h4, h5, h6 {
            color: #222;
            margin-top: 20px;
            margin-bottom: 10px;
            page-break-after: avoid;
        }
        p { margin-bottom: 12px; text-align: justify; }
        table { border-collapse: collapse; width: 100%; margin: 15px 0; page-break-inside: avoid; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f5f5f5; font-weight: bold; }
        ul, ol { margin: 10px 0; padding-left: 30px; }
        img { max-width: 100%; height: auto; }
        pre, code { background: #f5f5f5; padding: 2px 5px; border-radius: 3px; }
        pre { padding: 10px; overflow-x: auto; }
    </style>
</head>
<body>${cleanContent}</body>
</html>`;
    }

    // 创建 PDF HTML 文档
    createPdfHtmlDocument(text, numPages) {
        // 转义 HTML 特殊字符
        const escapeHtml = (str) => str
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');

        // 处理文本：保留换行，转义 HTML
        const processedText = escapeHtml(text)
            .replace(/--- 第 (\d+) 页 ---/g, '</div><div class="page"><div class="page-number">第 $1 页</div>')
            .replace(/\n/g, '<br>');

        return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PDF转换结果</title>
    <style>
        body { 
            font-family: "Noto Sans SC", "Microsoft YaHei", Arial, sans-serif; 
            max-width: 800px; 
            margin: 0 auto; 
            padding: 40px; 
            line-height: 1.8;
            word-wrap: break-word;
        }
        h1 { color: #222; }
        .page { 
            border-bottom: 2px solid #eee; 
            padding: 20px 0; 
            margin-bottom: 20px; 
        }
        .page-number { 
            color: #667eea; 
            font-size: 14px; 
            margin-bottom: 15px;
            font-weight: bold;
        }
    </style>
</head>
<body>
    <h1>📄 PDF 转换结果</h1>
    <p>总页数: ${numPages}</p>
    <hr>
    <div>${processedText}</div>
</body>
</html>`;
    }

    // 创建 Word 文档
    createWordDocument(text, numPages) {
        // 完全转义文本中的 HTML 标签
        const escapedText = text
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');

        return `<html xmlns:o='urn:schemas-microsoft-com:office:office' 
              xmlns:w='urn:schemas-microsoft-com:office:word' 
              xmlns='http://www.w3.org/TR/REC-html40'>
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <title>PDF转换结果</title>
    <style>
        body { 
            font-family: "Microsoft YaHei", Arial, sans-serif; 
            padding: 40px; 
            line-height: 1.8;
            font-size: 12pt;
        }
        pre { white-space: pre-wrap; word-wrap: break-word; font-family: "Microsoft YaHei", Arial, sans-serif; }
    </style>
</head>
<body>
    <h1>PDF 转换结果</h1>
    <p>原始PDF页数: ${numPages}</p>
    <hr>
    <pre>${escapedText}</pre>
</body>
</html>`;
    }

    // 创建 PPTX HTML 文档
    createPptxHtmlDocument(pptxInfo) {
        const slidesHtml = pptxInfo.content.map(s => `
            <div class="slide">
                <div class="slide-number">幻灯片 ${s.slide}</div>
                <div class="slide-content">${s.text || '(无文本内容)'}</div>
            </div>
        `).join('');

        return `<html>
<head>
    <meta charset="UTF-8">
    <style>
        body { font-family: "Microsoft YaHei", Arial, sans-serif; padding: 40px; }
        .slide { 
            background: white;
            border: 2px solid #ddd; 
            margin-bottom: 30px; 
            padding: 40px; 
            min-height: 400px; 
            page-break-after: always;
        }
        .slide-number { color: #667eea; font-size: 14px; margin-bottom: 20px; font-weight: bold; }
        .info { background: #fff3cd; border: 1px solid #ffeaa7; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
    </style>
</head>
<body>
    <div class="info">
        <strong>📊 PPT转换结果</strong><br>
        原始文件: ${this.currentFile.name}<br>
        总幻灯片数: ${pptxInfo.slideCount}
    </div>
    ${slidesHtml || '<div class="slide"><div class="slide-content">无法提取幻灯片内容</div></div>'}
</body>
</html>`;
    }

    // 创建 PPTX 图片文档
    createPptxImagesDocument(pptxInfo) {
        return `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>PPT 内容提取</title>
    <style>
        body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 40px; }
        .info { background: #f0f0f0; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
        .slide-info { background: white; border: 1px solid #ddd; padding: 15px; margin-bottom: 10px; border-radius: 5px; }
        .slide-num { color: #667eea; font-weight: bold; }
    </style>
</head>
<body>
    <h1>🖼️ PPT 内容提取</h1>
    <div class="info">
        <p><strong>原始文件:</strong> ${this.currentFile.name}</p>
        <p><strong>幻灯片数:</strong> ${pptxInfo.slideCount}</p>
        <p><strong>说明:</strong> 纯前端环境下，PPT转图片需要服务器支持。已提取文本内容供参考。</p>
    </div>
    ${pptxInfo.content.map(s => `
        <div class="slide-info">
            <div class="slide-num">幻灯片 ${s.slide}</div>
            <div>${s.text || '(无文本)'}</div>
        </div>
    `).join('')}
</body>
</html>`;
    }

    // 获取输出文件名
    getOutputFilename(ext) {
        const baseName = this.currentFile.name.replace(/\.[^/.]+$/, '');
        return `${baseName}_converted.${ext}`;
    }

    // 显示结果
    showResult(result) {
        const resultSection = document.getElementById('resultSection');
        const downloadBtn = document.getElementById('downloadBtn');
        const resultInfo = document.getElementById('resultInfo');

        // 验证结果
        if (!result || !result.blob) {
            this.showError('转换结果无效');
            return;
        }

        // 清理之前的URL
        if (this.currentDownloadUrl) {
            URL.revokeObjectURL(this.currentDownloadUrl);
        }

        try {
            // 创建新URL
            this.currentDownloadUrl = URL.createObjectURL(result.blob);
            downloadBtn.href = this.currentDownloadUrl;
            downloadBtn.download = result.filename || 'converted_file';

            resultInfo.textContent = `${result.filename || 'unknown'} · ${this.formatSize(result.blob.size)}`;
            resultSection.classList.add('show');
            resultSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        } catch (error) {
            console.error('显示结果失败:', error);
            this.showError('无法生成下载链接，请重试');
        }
    }

    // 隐藏结果
    hideResult() {
        document.getElementById('resultSection').classList.remove('show');
        if (this.currentDownloadUrl) {
            URL.revokeObjectURL(this.currentDownloadUrl);
            this.currentDownloadUrl = null;
        }
    }

    // 显示错误
    showError(message) {
        const errorElement = document.getElementById('errorMessage');
        if (!errorElement) return;

        // 确保消息是字符串
        const errorMsg = typeof message === 'string' ? message : '发生未知错误';
        errorElement.textContent = '❌ ' + errorMsg;
        errorElement.classList.add('show');

        // 自动隐藏
        if (this.errorTimeout) {
            clearTimeout(this.errorTimeout);
        }
        this.errorTimeout = setTimeout(() => this.hideError(), 8000);
    }

    // 隐藏错误
    hideError() {
        document.getElementById('errorMessage').classList.remove('show');
    }

    // 获取错误信息
    getErrorMessage(error) {
        if (!error || !error.message) {
            return '转换失败，请重试';
        }

        const msg = error.message.toLowerCase();

        // 网络错误
        if (msg.includes('network') || msg.includes('fetch') || msg.includes('load') || msg.includes('undefined')) {
            return '网络连接问题或库加载失败，请检查网络后刷新页面重试';
        }

        // 加密文件
        if (msg.includes('password') || msg.includes('encrypted')) {
            return '文件已加密，无法转换';
        }

        // 文件损坏
        if (msg.includes('corrupt') || msg.includes('invalid') || msg.includes('parse')) {
            return '文件已损坏或格式无效，请检查文件';
        }

        // 超时
        if (msg.includes('timeout')) {
            return '转换超时，请尝试较小的文件或检查网络';
        }

        // 内存不足
        if (msg.includes('memory') || msg.includes('quota')) {
            return '文件太大，内存不足，请尝试更小的文件';
        }

        // 空文件
        if (msg.includes('empty')) {
            return '文件为空，请选择其他文件';
        }

        // 返回原始错误信息（限制长度）
        const originalMsg = error.message;
        if (originalMsg.length > 100) {
            return originalMsg.substring(0, 100) + '...';
        }
        return originalMsg;
    }

    // 文件转 ArrayBuffer
    fileToArrayBuffer(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = () => reject(new Error('文件读取失败'));
            reader.readAsArrayBuffer(file);
        });
    }

    // 延迟
    sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
}

// 初始化
document.addEventListener('DOMContentLoaded', () => {
    try {
        window.docConverter = new DocConverter();
        console.log('✅ 文档转换工具 Pro v2.0 已加载');
        console.log('📚 支持的转换:');
        console.log('   Word (.docx) → PDF, HTML, TXT');
        console.log('   PDF (.pdf) → Word, HTML, TXT');
        console.log('   PPT (.pptx) → PDF, 图片集');
    } catch (error) {
        console.error('❌ 初始化失败:', error);
        alert('工具加载失败，请刷新页面重试');
    }
});
