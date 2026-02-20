// 文档转换工具 - 纯前端实现
class DocConverter {
    constructor() {
        this.currentFile = null;
        this.currentFormat = null;
        this.currentDownloadUrl = null;
        this.init();
    }

    init() {
        this.bindEvents();
    }

    bindEvents() {
        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');
        const convertBtn = document.getElementById('convertBtn');

        // 点击上传
        dropZone.addEventListener('click', () => fileInput.click());

        // 文件选择
        fileInput.addEventListener('change', (e) => this.handleFileSelect(e.target.files[0]));

        // 拖拽上传
        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.classList.add('dragover');
        });

        dropZone.addEventListener('dragleave', () => {
            dropZone.classList.remove('dragover');
        });

        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
            const file = e.dataTransfer.files[0];
            if (file) this.handleFileSelect(file);
        });

        // 转换按钮
        convertBtn.addEventListener('click', () => this.startConversion());
    }

    handleFileSelect(file) {
        if (!file) return;

        // 检查文件大小 (最大50MB)
        const maxSize = 50 * 1024 * 1024; // 50MB
        if (file.size > maxSize) {
            alert('文件太大啦！请上传小于50MB的文件 💕');
            return;
        }

        // 检查文件是否为空
        if (file.size === 0) {
            alert('文件是空的哦，请重新选择 😅');
            return;
        }

        const validTypes = ['.docx', '.pdf', '.pptx'];
        const ext = '.' + file.name.split('.').pop().toLowerCase();

        if (!validTypes.includes(ext)) {
            alert('请选择 .docx, .pdf 或 .pptx 格式的文件 📄');
            return;
        }

        this.currentFile = file;
        this.showFileInfo(file);
        this.showFormatOptions(ext);
    }

    showFileInfo(file) {
        const fileInfo = document.getElementById('fileInfo');
        const fileName = document.getElementById('fileName');
        const fileSize = document.getElementById('fileSize');

        fileName.textContent = file.name;
        fileSize.textContent = `(${(file.size / 1024 / 1024).toFixed(2)} MB)`;
        fileInfo.classList.add('show');
    }

    showFormatOptions(currentExt) {
        const formatSection = document.getElementById('formatSection');
        const formatOptions = document.getElementById('formatOptions');
        const convertBtn = document.getElementById('convertBtn');

        // 根据当前格式显示可转换的选项
        const formats = {
            '.docx': [
                { value: 'pdf', label: 'PDF', icon: '📄' },
                { value: 'html', label: 'HTML', icon: '🌐' },
                { value: 'txt', label: '纯文本', icon: '📝' }
            ],
            '.pdf': [
                { value: 'docx', label: 'Word', icon: '📘' },
                { value: 'txt', label: '纯文本', icon: '📝' },
                { value: 'html', label: 'HTML', icon: '🌐' }
            ],
            '.pptx': [
                { value: 'pdf', label: 'PDF', icon: '📄' },
                { value: 'images', label: '图片集', icon: '🖼️' }
            ]
        };

        const options = formats[currentExt] || [];
        formatOptions.innerHTML = options.map(opt => `
            <button class="format-btn" data-format="${opt.value}">
                ${opt.icon} ${opt.label}
            </button>
        `).join('');

        // 绑定格式选择事件
        formatOptions.querySelectorAll('.format-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                formatOptions.querySelectorAll('.format-btn').forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                this.currentFormat = btn.dataset.format;
                convertBtn.disabled = false;
            });
        });

        formatSection.style.display = 'block';
        convertBtn.disabled = true;
    }

    async startConversion() {
        if (!this.currentFile || !this.currentFormat) {
            alert('请先选择文件和目标格式哦 😊');
            return;
        }

        const progressBar = document.getElementById('progressBar');
        const progressFill = document.getElementById('progressFill');
        const convertBtn = document.getElementById('convertBtn');
        const resultSection = document.getElementById('resultSection');

        // 隐藏之前的结果
        resultSection.classList.remove('show');
        
        progressBar.classList.add('show');
        convertBtn.disabled = true;
        convertBtn.textContent = '正在读取文件...';

        try {
            // 步骤1: 读取文件
            this.updateProgress(10);
            convertBtn.textContent = '正在解析...';
            await this.sleep(200);
            
            // 步骤2: 转换
            this.updateProgress(30);
            convertBtn.textContent = '正在转换...';
            
            const result = await this.convertFile();
            
            // 步骤3: 生成文件
            this.updateProgress(80);
            convertBtn.textContent = '正在生成...';
            await this.sleep(300);
            
            this.updateProgress(100);
            this.showResult(result);
            
        } catch (error) {
            console.error('转换失败:', error);
            let errorMsg = '转换失败 😢\n\n';
            
            if (error.message.includes('network') || error.message.includes('fetch')) {
                errorMsg += '网络连接问题，请检查网络后重试';
            } else if (error.message.includes('password') || error.message.includes('encrypted')) {
                errorMsg += '文件可能被加密，无法转换';
            } else if (error.message.includes('corrupt') || error.message.includes('invalid')) {
                errorMsg += '文件可能已损坏，请检查文件';
            } else {
                errorMsg += error.message;
            }
            
            alert(errorMsg);
        } finally {
            progressBar.classList.remove('show');
            convertBtn.disabled = false;
            convertBtn.textContent = '开始转换';
            this.updateProgress(0);
        }
    }

    updateProgress(percent) {
        document.getElementById('progressFill').style.width = percent + '%';
    }

    sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    async convertFile() {
        const ext = '.' + this.currentFile.name.split('.').pop().toLowerCase();
        const arrayBuffer = await this.fileToArrayBuffer(this.currentFile);

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

    fileToArrayBuffer(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
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

    async docxToPdf(arrayBuffer) {
        // 使用 mammoth 转换为HTML，保留格式
        const result = await mammoth.convertToHtml({ arrayBuffer });
        const htmlContent = result.value;

        // 创建完整的HTML文档
        const html = `
            <!DOCTYPE html>
            <html lang="zh-CN">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <style>
                    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@400;700&display=swap');
                    
                    body { 
                        font-family: "Noto Sans SC", "Microsoft YaHei", "SimHei", "PingFang SC", Arial, sans-serif; 
                        padding: 40px; 
                        line-height: 1.8;
                        color: #333;
                        font-size: 11pt;
                    }
                    h1, h2, h3, h4, h5, h6 { 
                        color: #222; 
                        margin-top: 20px;
                        margin-bottom: 10px;
                        font-family: "Noto Sans SC", "Microsoft YaHei", "SimHei", sans-serif;
                    }
                    h1 { font-size: 24pt; }
                    h2 { font-size: 20pt; }
                    h3 { font-size: 16pt; }
                    p { 
                        margin-bottom: 12px; 
                        text-align: justify;
                    }
                    table { 
                        border-collapse: collapse; 
                        width: 100%; 
                        margin: 15px 0;
                    }
                    th, td { 
                        border: 1px solid #ddd; 
                        padding: 8px; 
                        text-align: left;
                    }
                    th { 
                        background-color: #f5f5f5; 
                        font-weight: bold;
                    }
                    ul, ol { 
                        margin: 10px 0; 
                        padding-left: 30px;
                    }
                    li { 
                        margin-bottom: 5px;
                    }
                    strong { font-weight: bold; }
                    em { font-style: italic; }
                </style>
            </head>
            <body>
                ${htmlContent}
            </body>
            </html>
        `;

        // 使用 html2pdf
        const element = document.createElement('div');
        element.innerHTML = html;
        document.body.appendChild(element);

        const opt = {
            margin: [15, 15, 15, 15],
            filename: 'converted.pdf',
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { 
                scale: 2,
                useCORS: true,
                letterRendering: true
            },
            jsPDF: { 
                unit: 'mm', 
                format: 'a4', 
                orientation: 'portrait'
            }
        };

        try {
            const pdf = await html2pdf().set(opt).from(element).output('blob');
            document.body.removeChild(element);
            return { blob: pdf, filename: 'converted.pdf', type: 'application/pdf' };
        } catch (error) {
            document.body.removeChild(element);
            throw new Error('PDF生成失败: ' + error.message);
        }
    }

    async docxToHtml(arrayBuffer) {
        const result = await mammoth.convertToHtml({ arrayBuffer });
        const html = `
            <!DOCTYPE html>
            <html lang="zh-CN">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Converted Document</title>
                <style>
                    body { 
                        font-family: "Microsoft YaHei", "SimHei", "PingFang SC", "Hiragino Sans GB", Arial, sans-serif; 
                        max-width: 800px; 
                        margin: 0 auto; 
                        padding: 40px; 
                        line-height: 1.8;
                        color: #333;
                    }
                    h1, h2, h3, h4, h5, h6 { 
                        color: #222; 
                        margin-top: 20px;
                        margin-bottom: 10px;
                        font-family: "Microsoft YaHei", "SimHei", sans-serif;
                    }
                    p { 
                        margin-bottom: 12px;
                        text-align: justify;
                    }
                    table { 
                        border-collapse: collapse; 
                        width: 100%; 
                        margin: 15px 0;
                    }
                    th, td { 
                        border: 1px solid #ddd; 
                        padding: 8px; 
                        text-align: left;
                    }
                    th { 
                        background-color: #f5f5f5; 
                        font-weight: bold;
                    }
                    ul, ol { 
                        margin: 10px 0; 
                        padding-left: 30px;
                    }
                </style>
            </head>
            <body>
                ${result.value}
            </body>
            </html>
        `;

        const blob = new Blob([html], { type: 'text/html; charset=utf-8' });
        return { blob, filename: 'converted.html', type: 'text/html' };
    }

    async docxToTxt(arrayBuffer) {
        const result = await mammoth.extractRawText({ arrayBuffer });
        const blob = new Blob([result.value], { type: 'text/plain' });
        return { blob, filename: 'converted.txt', type: 'text/plain' };
    }

    // PDF 转换
    async convertPdf(arrayBuffer, targetFormat) {
        switch (targetFormat) {
            case 'txt':
                return await this.pdfToTxt(arrayBuffer);
            case 'html':
                return await this.pdfToHtml(arrayBuffer);
            case 'docx':
                return await this.pdfToDocx(arrayBuffer);
            default:
                throw new Error('不支持的转换格式');
        }
    }

    async extractPdfText(arrayBuffer) {
        try {
            // 使用PDF.js提取文本，添加超时
            // cMap用于支持中文等CJK字符
            const loadingTask = pdfjsLib.getDocument({ 
                data: arrayBuffer,
                useSystemFonts: true,
                cMapUrl: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/cmaps/',
                cMapPacked: true
            });
            
            const pdf = await Promise.race([
                loadingTask.promise,
                new Promise((_, reject) => 
                    setTimeout(() => reject(new Error('PDF加载超时')), 30000)
                )
            ]);
            
            let fullText = '';
            const maxPages = Math.min(pdf.numPages, 100); // 最多处理100页
            
            for (let i = 1; i <= maxPages; i++) {
                try {
                    const page = await pdf.getPage(i);
                    const textContent = await page.getTextContent();
                    const pageText = textContent.items
                        .map(item => item.str)
                        .join(' ')
                        .replace(/\s+/g, ' ')
                        .trim();
                    
                    if (pageText) {
                        fullText += `\n--- 第 ${i} 页 ---\n${pageText}\n`;
                    }
                    
                    // 释放页面资源
                    page.cleanup();
                } catch (pageError) {
                    console.warn(`第${i}页提取失败:`, pageError);
                    fullText += `\n--- 第 ${i} 页 ---\n(无法提取内容)\n`;
                }
            }
            
            return { 
                text: fullText || '(未能提取到文本内容)', 
                numPages: pdf.numPages,
                processedPages: maxPages
            };
        } catch (error) {
            console.error('PDF文本提取错误:', error);
            throw new Error('PDF解析失败: ' + (error.message || '未知错误'));
        }
    }

    async pdfToTxt(arrayBuffer) {
        try {
            const { text, numPages } = await this.extractPdfText(arrayBuffer);
            const header = `PDF 文档转换结果\n页数: ${numPages}\n\n`;
            const blob = new Blob([header + text], { type: 'text/plain' });
            return { blob, filename: 'converted.txt', type: 'text/plain' };
        } catch (error) {
            throw new Error('PDF文本提取失败: ' + error.message);
        }
    }

    async pdfToHtml(arrayBuffer) {
        try {
            const { text, numPages } = await this.extractPdfText(arrayBuffer);
            // 处理文本，保留换行但改善格式
            const processedText = text
                .replace(/--- 第 (\d+) 页 ---/g, '</div><div class="page"><div class="page-number">第 $1 页</div>')
                .replace(/\n/g, '<br>');
            
            const html = `
                <!DOCTYPE html>
                <html lang="zh-CN">
                <head>
                    <meta charset="UTF-8">
                    <meta name="viewport" content="width=device-width, initial-scale=1.0">
                    <title>Converted PDF</title>
                    <style>
                        body { 
                            font-family: "Microsoft YaHei", "SimHei", "PingFang SC", Arial, sans-serif; 
                            max-width: 800px; 
                            margin: 0 auto; 
                            padding: 40px; 
                            line-height: 1.8;
                            color: #333;
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
                        .page-content {
                            white-space: pre-wrap;
                            word-wrap: break-word;
                        }
                    </style>
                </head>
                <body>
                    <h1>📄 PDF 转换结果</h1>
                    <p>总页数: ${numPages}</p>
                    <hr>
                    <div class="page-content">${processedText}</div>
                </body>
                </html>
            `;
            const blob = new Blob([html], { type: 'text/html; charset=utf-8' });
            return { blob, filename: 'converted.html', type: 'text/html' };
        } catch (error) {
            throw new Error('PDF转HTML失败: ' + error.message);
        }
    }

    async pdfToDocx(arrayBuffer) {
        try {
            const { text, numPages } = await this.extractPdfText(arrayBuffer);
            
            // 创建HTML格式的Word文档（可以被Word打开）
            const html = `
                <html xmlns:o='urn:schemas-microsoft-com:office:office' 
                      xmlns:w='urn:schemas-microsoft-com:office:word' 
                      xmlns='http://www.w3.org/TR/REC-html40'>
                <head>
                    <meta charset="UTF-8">
                    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
                    <title>Converted PDF</title>
                    <style>
                        body { 
                            font-family: "Microsoft YaHei", "SimSun", Arial, sans-serif; 
                            padding: 40px; 
                            line-height: 1.8;
                            font-size: 12pt;
                        }
                        h1 { 
                            color: #333; 
                            font-family: "Microsoft YaHei", "SimHei", sans-serif;
                        }
                        .page-break { page-break-before: always; }
                        pre { 
                            white-space: pre-wrap; 
                            font-family: "Microsoft YaHei", "SimSun", sans-serif;
                            font-size: 11pt;
                            line-height: 1.6;
                        }
                    </style>
                </head>
                <body>
                    <h1>PDF 转换结果</h1>
                    <p>原始PDF页数: ${numPages}</p>
                    <hr>
                    <pre>${text.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</pre>
                </body>
                </html>
            `;
            
            const blob = new Blob([html], { type: 'application/msword; charset=utf-8' });
            return { blob, filename: 'converted.doc', type: 'application/msword' };
        } catch (error) {
            throw new Error('PDF转Word失败: ' + error.message);
        }
    }

    // PPT 转换 - 使用JSZip解析PPTX文件
    async convertPptx(arrayBuffer, targetFormat) {
        // 检查是否支持JSZip
        if (typeof JSZip === 'undefined') {
            // 动态加载JSZip
            await this.loadScript('https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js');
        }
        
        switch (targetFormat) {
            case 'pdf':
                return await this.pptxToPdf(arrayBuffer);
            case 'images':
                return await this.pptxToImages(arrayBuffer);
            default:
                throw new Error('不支持的转换格式');
        }
    }

    async loadScript(src) {
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = src;
            script.onload = resolve;
            script.onerror = reject;
            document.head.appendChild(script);
        });
    }

    async extractPptxContent(arrayBuffer) {
        try {
            const zip = await JSZip.loadAsync(arrayBuffer);
            
            // 读取幻灯片数量
            const slides = Object.keys(zip.files).filter(name => 
                name.startsWith('ppt/slides/slide') && name.endsWith('.xml')
            );
            
            // 尝试读取内容（简化版）
            let content = [];
            for (let i = 0; i < Math.min(slides.length, 5); i++) {
                const slideContent = await zip.file(slides[i]).async('text');
                // 提取文本内容（简单正则）
                const texts = slideContent.match(/<a:t>([^<]+)<\/a:t>/g) || [];
                const slideText = texts.map(t => t.replace(/<\/?a:t>/g, '')).join(' ');
                content.push({
                    slide: i + 1,
                    text: slideText.substring(0, 200) + (slideText.length > 200 ? '...' : '')
                });
            }
            
            return {
                slideCount: slides.length,
                content: content
            };
        } catch (error) {
            console.error('PPT解析失败:', error);
            return {
                slideCount: 0,
                content: [],
                error: error.message
            };
        }
    }

    async pptxToPdf(arrayBuffer) {
        const pptxInfo = await this.extractPptxContent(arrayBuffer);
        
        // 生成幻灯片HTML
        const slidesHtml = pptxInfo.content.map((slide, index) => `
            <div class="slide">
                <div class="slide-number">幻灯片 ${slide.slide}</div>
                <div class="slide-content">${slide.text || '(无文本内容)'}</div>
            </div>
        `).join('');
        
        const html = `
            <html>
            <head>
                <meta charset="UTF-8">
                <style>
                    body { font-family: "Microsoft YaHei", Arial, sans-serif; padding: 40px; background: #f5f5f5; }
                    .slide { 
                        background: white;
                        border: 2px solid #ddd; 
                        margin-bottom: 30px; 
                        padding: 40px; 
                        min-height: 400px; 
                        page-break-after: always;
                        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
                    }
                    .slide-number { 
                        color: #667eea; 
                        font-size: 14px; 
                        margin-bottom: 20px;
                        font-weight: bold;
                    }
                    .slide-content { 
                        font-size: 16px; 
                        line-height: 1.6;
                        color: #333;
                    }
                    .info {
                        background: #fff3cd;
                        border: 1px solid #ffeaa7;
                        padding: 15px;
                        border-radius: 5px;
                        margin-bottom: 20px;
                    }
                </style>
            </head>
            <body>
                <div class="info">
                    <strong>📊 PPT转换结果</strong><br>
                    原始文件: ${this.currentFile.name}<br>
                    总幻灯片数: ${pptxInfo.slideCount}<br>
                    <small>注：纯前端PPT解析有限制，仅提取文本内容</small>
                </div>
                ${slidesHtml || '<div class="slide"><div class="slide-content">无法提取幻灯片内容</div></div>'}
            </body>
            </html>
        `;

        const element = document.createElement('div');
        element.innerHTML = html;
        document.body.appendChild(element);

        const opt = {
            margin: [10, 10, 10, 10],
            filename: 'presentation.pdf',
            image: { type: 'jpeg', quality: 0.95 },
            html2canvas: { scale: 2, useCORS: true },
            jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape' }
        };

        try {
            const pdf = await html2pdf().set(opt).from(element).output('blob');
            document.body.removeChild(element);
            return { blob: pdf, filename: 'presentation.pdf', type: 'application/pdf' };
        } catch (error) {
            document.body.removeChild(element);
            throw new Error('PDF生成失败: ' + error.message);
        }
    }

    async pptxToImages(arrayBuffer) {
        const pptxInfo = await this.extractPptxContent(arrayBuffer);
        
        const html = `
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="UTF-8">
                <title>PPT 转换结果</title>
                <style>
                    body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 40px; }
                    .info { background: #f0f0f0; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
                    .slide-info { 
                        background: white; 
                        border: 1px solid #ddd; 
                        padding: 15px; 
                        margin-bottom: 10px;
                        border-radius: 5px;
                    }
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
            </html>
        `;
        const blob = new Blob([html], { type: 'text/html' });
        return { blob, filename: 'pptx-content.html', type: 'text/html' };
    }

    showResult(result) {
        const resultSection = document.getElementById('resultSection');
        const downloadBtn = document.getElementById('downloadBtn');

        // 清理之前的URL对象，防止内存泄漏
        if (this.currentDownloadUrl) {
            URL.revokeObjectURL(this.currentDownloadUrl);
        }

        // 创建新的下载链接
        this.currentDownloadUrl = URL.createObjectURL(result.blob);
        downloadBtn.href = this.currentDownloadUrl;
        downloadBtn.download = result.filename;
        downloadBtn.textContent = `下载 ${result.filename}`;

        resultSection.classList.add('show');
        resultSection.scrollIntoView({ behavior: 'smooth' });
    }
}

// 初始化
document.addEventListener('DOMContentLoaded', () => {
    new DocConverter();
});
