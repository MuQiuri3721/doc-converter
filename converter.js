// 文档转换工具 - 纯前端实现
class DocConverter {
    constructor() {
        this.currentFile = null;
        this.currentFormat = null;
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

        const validTypes = ['.docx', '.pdf', '.pptx'];
        const ext = '.' + file.name.split('.').pop().toLowerCase();

        if (!validTypes.includes(ext)) {
            alert('请选择 .docx, .pdf 或 .pptx 格式的文件');
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
        if (!this.currentFile || !this.currentFormat) return;

        const progressBar = document.getElementById('progressBar');
        const progressFill = document.getElementById('progressFill');
        const convertBtn = document.getElementById('convertBtn');

        progressBar.classList.add('show');
        convertBtn.disabled = true;
        convertBtn.textContent = '转换中...';

        try {
            // 模拟进度
            this.updateProgress(10);
            await this.sleep(300);
            this.updateProgress(30);

            const result = await this.convertFile();
            
            this.updateProgress(80);
            await this.sleep(200);
            this.updateProgress(100);

            this.showResult(result);
        } catch (error) {
            console.error('转换失败:', error);
            alert('转换失败: ' + error.message);
        } finally {
            progressBar.classList.remove('show');
            convertBtn.disabled = false;
            convertBtn.textContent = '开始转换';
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
        // 使用 mammoth 提取内容，然后生成 PDF
        const result = await mammoth.extractRawText({ arrayBuffer });
        const text = result.value;

        // 创建临时 HTML
        const html = `
            <html>
            <head>
                <style>
                    body { font-family: Arial, sans-serif; padding: 40px; line-height: 1.6; }
                    h1, h2, h3 { color: #333; }
                    p { margin-bottom: 10px; }
                </style>
            </head>
            <body>
                <pre style="white-space: pre-wrap; font-family: Arial, sans-serif;">${text}</pre>
            </body>
            </html>
        `;

        // 使用 html2pdf
        const element = document.createElement('div');
        element.innerHTML = html;
        document.body.appendChild(element);

        const opt = {
            margin: 10,
            filename: 'converted.pdf',
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { scale: 2 },
            jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
        };

        const pdf = await html2pdf().set(opt).from(element).output('blob');
        document.body.removeChild(element);

        return { blob: pdf, filename: 'converted.pdf', type: 'application/pdf' };
    }

    async docxToHtml(arrayBuffer) {
        const result = await mammoth.convertToHtml({ arrayBuffer });
        const html = `
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="UTF-8">
                <title>Converted Document</title>
                <style>
                    body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 40px; line-height: 1.6; }
                    h1, h2, h3 { color: #333; }
                    p { margin-bottom: 10px; }
                </style>
            </head>
            <body>
                ${result.value}
            </body>
            </html>
        `;

        const blob = new Blob([html], { type: 'text/html' });
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
                // PDF转Word比较复杂，这里先提供文本版本
                return await this.pdfToTxt(arrayBuffer);
            default:
                throw new Error('不支持的转换格式');
        }
    }

    async pdfToTxt(arrayBuffer) {
        try {
            const pdfDoc = await PDFLib.PDFDocument.load(arrayBuffer);
            let text = '';

            // 获取所有页面
            const pages = pdfDoc.getPages();
            
            // 由于pdf-lib不直接支持文本提取，我们创建一个简单的表示
            text = `PDF 文档\n`;
            text += `页数: ${pages.length}\n`;
            text += `尺寸: ${Math.round(pages[0].getWidth())} x ${Math.round(pages[0].getHeight())}\n\n`;
            text += `注意: 纯前端PDF文本提取有限制，建议使用专业工具进行完整转换。\n`;

            const blob = new Blob([text], { type: 'text/plain' });
            return { blob, filename: 'converted.txt', type: 'text/plain' };
        } catch (error) {
            throw new Error('PDF解析失败: ' + error.message);
        }
    }

    async pdfToHtml(arrayBuffer) {
        const result = await this.pdfToTxt(arrayBuffer);
        const html = `
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="UTF-8">
                <title>Converted PDF</title>
                <style>
                    body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 40px; }
                    pre { white-space: pre-wrap; line-height: 1.6; }
                </style>
            </head>
            <body>
                <pre>${await result.blob.text()}</pre>
            </body>
            </html>
        `;
        const blob = new Blob([html], { type: 'text/html' });
        return { blob, filename: 'converted.html', type: 'text/html' };
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

    async pptxToPdf(arrayBuffer) {
        // PPT转PDF - 创建包含幻灯片信息的PDF
        const html = `
            <html>
            <head>
                <style>
                    body { font-family: Arial, sans-serif; padding: 40px; }
                    .slide { border: 2px solid #ddd; margin-bottom: 30px; padding: 40px; min-height: 400px; page-break-after: always; }
                    .slide-number { color: #999; font-size: 14px; margin-bottom: 20px; }
                </style>
            </head>
            <body>
                <div class="slide">
                    <div class="slide-number">幻灯片 1</div>
                    <h1>PPT 演示文稿</h1>
                    <p>原始文件: ${this.currentFile.name}</p>
                    <p>注意: 纯前端PPT转换有限制，建议下载后使用专业软件查看完整内容。</p>
                </div>
            </body>
            </html>
        `;

        const element = document.createElement('div');
        element.innerHTML = html;
        document.body.appendChild(element);

        const opt = {
            margin: 10,
            filename: 'presentation.pdf',
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { scale: 2 },
            jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape' }
        };

        const pdf = await html2pdf().set(opt).from(element).output('blob');
        document.body.removeChild(element);

        return { blob: pdf, filename: 'presentation.pdf', type: 'application/pdf' };
    }

    async pptxToImages(arrayBuffer) {
        // 创建包含说明的HTML文件
        const html = `
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="UTF-8">
                <title>PPT 转换结果</title>
                <style>
                    body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 40px; }
                    .info { background: #f0f0f0; padding: 20px; border-radius: 8px; }
                </style>
            </head>
            <body>
                <h1>🖼️ PPT 转图片</h1>
                <div class="info">
                    <p><strong>原始文件:</strong> ${this.currentFile.name}</p>
                    <p>纯前端环境下，PPT转图片需要服务器支持。</p>
                    <p>建议使用: LibreOffice、Microsoft PowerPoint 或在线转换工具。</p>
                </div>
            </body>
            </html>
        `;
        const blob = new Blob([html], { type: 'text/html' });
        return { blob, filename: 'pptx-info.html', type: 'text/html' };
    }

    showResult(result) {
        const resultSection = document.getElementById('resultSection');
        const downloadBtn = document.getElementById('downloadBtn');

        // 创建下载链接
        const url = URL.createObjectURL(result.blob);
        downloadBtn.href = url;
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
