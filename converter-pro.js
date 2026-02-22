/**
 * 文档转换工具 Pro v2.1
 * 基于GitHub优秀项目整合优化
 * 
 * 技术栈整合 (总Stars: 117,400+):
 * - mammoth.js (3.2k ⭐) - Word解析
 * - pdf.js (47k ⭐) - PDF解析  
 * - pdf-lib (8.3k ⭐) - PDF创建/修改
 * - pdfmake (12.2k ⭐) - PDF生成
 * - docx.js (3.4k ⭐) - Word生成
 * - SheetJS (33k ⭐) - Excel解析
 * - JSZip (8.5k ⭐) - ZIP处理
 * - html2pdf.js (1.8k ⭐) - 降级方案
 */

class DocConverterPro {
    constructor() {
        this.currentFile = null;
        this.currentFormat = null;
        this.currentDownloadUrl = null;
        this.loadedLibraries = {};
        this.maxFileSize = 50 * 1024 * 1024; // 50MB
        this.maxPdfPages = 50; // PDF最大处理页数
        this.maxPptSlides = 50; // PPT最大处理幻灯片数
        this.init();
    }

    init() {
        this.bindEvents();
        this.configurePDFjs();
        this.preloadLibraries();
    }

    // 配置 PDF.js
    configurePDFjs() {
        if (typeof pdfjsLib !== 'undefined') {
            pdfjsLib.GlobalWorkerOptions.workerSrc = 
                'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
            console.log('✅ PDF.js 配置成功');
        }
    }

    // 预加载关键库
    preloadLibraries() {
        // 动态加载 docx.js 用于更好的Word生成
        this.loadScript('https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.min.js', 'docx');
    }

    // 动态加载脚本
    loadScript(src, name) {
        return new Promise((resolve, reject) => {
            if (this.loadedLibraries[name]) {
                resolve();
                return;
            }
            
            const script = document.createElement('script');
            script.src = src;
            script.onload = () => {
                this.loadedLibraries[name] = true;
                console.log(`✅ ${name} 加载成功`);
                resolve();
            };
            script.onerror = reject;
            document.head.appendChild(script);
        });
    }

    // 绑定事件
    bindEvents() {
        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');

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
        document.getElementById('convertBtn').addEventListener('click', () => this.startConversion());
        document.getElementById('removeFile').addEventListener('click', () => this.clearFile());
    }

    // 处理文件选择
    handleFileSelect(file) {
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
        const validTypes = ['.docx', '.pdf', '.pptx', '.xlsx', '.xls', '.png', '.jpg', '.jpeg'];

        if (!file) {
            return { valid: false, message: '请选择文件' };
        }

        if (file.size === 0) {
            return { valid: false, message: '文件为空，请选择其他文件' };
        }

        if (file.size > this.maxFileSize) {
            return { valid: false, message: `文件太大！最大支持 ${this.formatSize(this.maxFileSize)}` };
        }

        const ext = '.' + file.name.split('.').pop().toLowerCase();
        if (!validTypes.includes(ext)) {
            return { valid: false, message: '不支持的格式！请上传 .docx, .pdf, .pptx, .xlsx, .png, .jpg 文件' };
        }

        return { valid: true };
    }

    // 显示文件信息
    showFileInfo(file) {
        document.getElementById('fileName').textContent = file.name;
        document.getElementById('fileSize').textContent = this.formatSize(file.size);
        document.getElementById('fileInfo').classList.add('show');
    }

    // 显示格式选项
    showFormatOptions(filename) {
        const ext = '.' + filename.split('.').pop().toLowerCase();
        const formats = this.getFormatOptions(ext);
        
        const container = document.getElementById('formatOptions');
        container.innerHTML = '';
        
        formats.forEach(format => {
            const btn = document.createElement('button');
            btn.className = 'format-btn';
            btn.textContent = format.label;
            btn.dataset.format = format.value;
            btn.addEventListener('click', () => this.selectFormat(format.value, btn));
            container.appendChild(btn);
        });
        
        container.classList.add('show');
    }

    // 获取格式选项
    getFormatOptions(ext) {
        const formats = {
            '.docx': [
                { value: 'pdf', label: 'PDF' },
                { value: 'html', label: 'HTML' },
                { value: 'txt', label: 'TXT' }
            ],
            '.pdf': [
                { value: 'docx', label: 'Word' },
                { value: 'html', label: 'HTML' },
                { value: 'txt', label: 'TXT' },
                { value: 'images', label: '图片' }
            ],
            '.pptx': [
                { value: 'pdf', label: 'PDF' },
                { value: 'images', label: '图片集' }
            ],
            '.xlsx': [
                { value: 'pdf', label: 'PDF' },
                { value: 'csv', label: 'CSV' },
                { value: 'json', label: 'JSON' }
            ],
            '.xls': [
                { value: 'pdf', label: 'PDF' },
                { value: 'csv', label: 'CSV' }
            ],
            '.png': [
                { value: 'pdf', label: 'PDF' },
                { value: 'jpg', label: 'JPG' }
            ],
            '.jpg': [
                { value: 'pdf', label: 'PDF' },
                { value: 'png', label: 'PNG' }
            ],
            '.jpeg': [
                { value: 'pdf', label: 'PDF' },
                { value: 'png', label: 'PNG' }
            ]
        };
        return formats[ext] || [];
    }

    // 选择格式
    selectFormat(format, btn) {
        this.currentFormat = format;
        document.querySelectorAll('.format-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        document.getElementById('convertBtn').disabled = false;
    }

    // 开始转换
    async startConversion() {
        if (!this.currentFile || !this.currentFormat) {
            this.showError('请选择文件和目标格式');
            return;
        }

        this.showLoading(true);
        
        try {
            const ext = '.' + this.currentFile.name.split('.').pop().toLowerCase();
            let result;

            // 根据文件类型和转换目标选择方法
            switch (ext + '->' + this.currentFormat) {
                // Word 转换
                case '.docx->pdf':
                    result = await this.convertDocxToPdf(this.currentFile);
                    break;
                case '.docx->html':
                    result = await this.convertDocxToHtml(this.currentFile);
                    break;
                case '.docx->txt':
                    result = await this.convertDocxToTxt(this.currentFile);
                    break;
                
                // PDF 转换
                case '.pdf->docx':
                    result = await this.convertPdfToDocx(this.currentFile);
                    break;
                case '.pdf->html':
                    result = await this.convertPdfToHtml(this.currentFile);
                    break;
                case '.pdf->txt':
                    result = await this.convertPdfToTxt(this.currentFile);
                    break;
                case '.pdf->images':
                    result = await this.convertPdfToImages(this.currentFile);
                    break;
                
                // PPT 转换
                case '.pptx->pdf':
                    result = await this.convertPptxToPdf(this.currentFile);
                    break;
                case '.pptx->images':
                    result = await this.convertPptxToImages(this.currentFile);
                    break;
                
                // Excel 转换
                case '.xlsx->pdf':
                case '.xls->pdf':
                    result = await this.convertExcelToPdf(this.currentFile);
                    break;
                case '.xlsx->csv':
                case '.xls->csv':
                    result = await this.convertExcelToCsv(this.currentFile);
                    break;
                case '.xlsx->json':
                    result = await this.convertExcelToJson(this.currentFile);
                    break;
                
                // 图片转换
                case '.png->pdf':
                case '.jpg->pdf':
                case '.jpeg->pdf':
                    result = await this.convertImageToPdf(this.currentFile);
                    break;
                case '.png->jpg':
                    result = await this.convertImageFormat(this.currentFile, 'jpeg');
                    break;
                case '.jpg->png':
                case '.jpeg->png':
                    result = await this.convertImageFormat(this.currentFile, 'png');
                    break;
                
                default:
                    throw new Error('不支持的转换类型');
            }

            this.showResult(result);
        } catch (error) {
            console.error('转换失败:', error);
            this.showError(this.getErrorMessage(error));
        } finally {
            this.showLoading(false);
        }
    }

    // ============ Word 转换方法 ============

    async convertDocxToPdf(file) {
        const arrayBuffer = await this.fileToArrayBuffer(file);
        const result = await mammoth.convertToHtml({ arrayBuffer });
        
        const html = this.createPdfHtmlDocument(result.value, file.name);
        
        const element = document.createElement('div');
        element.innerHTML = html;
        element.style.position = 'absolute';
        element.style.left = '-9999px';
        document.body.appendChild(element);
        
        try {
            const opt = {
                margin: 10,
                filename: this.getOutputFilename(file.name, 'pdf'),
                image: { type: 'jpeg', quality: 0.98 },
                html2canvas: { 
                    scale: 2,
                    useCORS: true,
                    logging: false
                },
                jsPDF: { 
                    unit: 'mm', 
                    format: 'a4', 
                    orientation: 'portrait',
                    compress: true
                }
            };
            
            const blob = await html2pdf().set(opt).from(element).output('blob');
            
            return {
                blob: blob,
                filename: this.getOutputFilename(file.name, 'pdf'),
                type: 'application/pdf'
            };
        } finally {
            document.body.removeChild(element);
        }
    }

    async convertDocxToHtml(file) {
        const arrayBuffer = await this.fileToArrayBuffer(file);
        const result = await mammoth.convertToHtml({ arrayBuffer });
        
        const html = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>${this.escapeHtml(file.name)}</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@400;700&display=swap');
        body { 
            font-family: "Noto Sans SC", Arial, sans-serif; 
            max-width: 800px; 
            margin: 0 auto; 
            padding: 40px;
            line-height: 1.6;
        }
        h1, h2, h3 { color: #333; }
        p { margin: 10px 0; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid #ddd; padding: 8px; }
        th { background: #f5f5f5; }
    </style>
</head>
<body>${result.value}</body>
</html>`;
        
        return {
            blob: new Blob([html], { type: 'text/html' }),
            filename: this.getOutputFilename(file.name, 'html'),
            type: 'text/html'
        };
    }

    async convertDocxToTxt(file) {
        const arrayBuffer = await this.fileToArrayBuffer(file);
        const result = await mammoth.extractRawText({ arrayBuffer });
        
        return {
            blob: new Blob([result.value], { type: 'text/plain' }),
            filename: this.getOutputFilename(file.name, 'txt'),
            type: 'text/plain'
        };
    }

    // ============ PDF 转换方法 ============

    async convertPdfToDocx(file) {
        // 检查 PDF.js 是否加载
        if (typeof pdfjsLib === 'undefined') {
            throw new Error('PDF.js 库未加载，请刷新页面重试');
        }
        
        const arrayBuffer = await this.fileToArrayBuffer(file);
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        
        let text = '';
        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const content = await page.getTextContent();
            text += content.items.map(item => item.str).join(' ') + '\n\n';
            page.cleanup();
        }
        pdf.destroy();
        
        // 使用 docx.js 创建更好的Word文档
        if (typeof docx !== 'undefined') {
            const { Document, Paragraph, Packer } = docx;
            
            const paragraphs = text.split('\n\n').map(p => 
                new Paragraph({ text: p.trim() })
            );
            
            const doc = new Document({
                sections: [{
                    properties: {},
                    children: paragraphs
                }]
            });
            
            const blob = await Packer.toBlob(doc);
            
            return {
                blob: blob,
                filename: this.getOutputFilename(file.name, 'docx'),
                type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            };
        } else {
            // 降级方案：使用HTML格式
            const html = `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
<head><meta charset='utf-8'><title>Document</title></head>
<body>${this.escapeHtml(text).replace(/\n/g, '<br>')}</body>
</html>`;
            
            return {
                blob: new Blob([html], { type: 'application/msword' }),
                filename: this.getOutputFilename(file.name, 'doc'),
                type: 'application/msword'
            };
        }
    }

    async convertPdfToHtml(file) {
        // 检查 PDF.js 是否加载
        if (typeof pdfjsLib === 'undefined') {
            throw new Error('PDF.js 库未加载，请刷新页面重试');
        }
        
        const arrayBuffer = await this.fileToArrayBuffer(file);
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        
        let html = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>${this.escapeHtml(file.name)}</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+SC&display=swap');
        body { 
            font-family: "Noto Sans SC", Arial, sans-serif; 
            max-width: 800px; 
            margin: 0 auto; 
            padding: 40px;
            line-height: 1.6;
        }
        .page { 
            border: 1px solid #ddd; 
            padding: 40px; 
            margin: 20px 0;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        .page-number {
            text-align: center;
            color: #999;
            margin-top: 10px;
            font-size: 12px;
        }
    </style>
</head>
<body>`;
        
        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const content = await page.getTextContent();
            const text = content.items.map(item => this.escapeHtml(item.str)).join(' ');
            
            html += `
        <div class="page">
            ${text.replace(/\n/g, '<br>')}
        </div>
        <div class="page-number">- 第 ${i} 页 -</div>`;
            
            page.cleanup();
        }
        
        html += '\n</body>\n</html>';
        pdf.destroy();
        
        return {
            blob: new Blob([html], { type: 'text/html' }),
            filename: this.getOutputFilename(file.name, 'html'),
            type: 'text/html'
        };
    }

    async convertPdfToTxt(file) {
        // 检查 PDF.js 是否加载
        if (typeof pdfjsLib === 'undefined') {
            throw new Error('PDF.js 库未加载，请刷新页面重试');
        }
        
        const arrayBuffer = await this.fileToArrayBuffer(file);
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        
        let text = '';
        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const content = await page.getTextContent();
            text += `=== 第 ${i} 页 ===\n`;
            text += content.items.map(item => item.str).join(' ') + '\n\n';
            page.cleanup();
        }
        pdf.destroy();
        
        return {
            blob: new Blob([text], { type: 'text/plain' }),
            filename: this.getOutputFilename(file.name, 'txt'),
            type: 'text/plain'
        };
    }

    async convertPdfToImages(file) {
        // 检查 PDF.js 是否加载
        if (typeof pdfjsLib === 'undefined') {
            throw new Error('PDF.js 库未加载，请刷新页面重试');
        }
        
        const arrayBuffer = await this.fileToArrayBuffer(file);
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        
        // 检查页数限制
        if (pdf.numPages > this.maxPdfPages) {
            pdf.destroy();
            throw new Error(`PDF页数过多！最多支持 ${this.maxPdfPages} 页，当前 ${pdf.numPages} 页`);
        }
        
        const images = [];
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        
        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const viewport = page.getViewport({ scale: 2 });
            
            canvas.width = viewport.width;
            canvas.height = viewport.height;
            
            await page.render({
                canvasContext: ctx,
                viewport: viewport
            }).promise;
            
            const blob = await new Promise(resolve => {
                canvas.toBlob(resolve, 'image/png');
            });
            
            images.push(blob);
            page.cleanup();
            
            // 每处理5页强制垃圾回收
            if (i % 5 === 0) {
                await new Promise(resolve => setTimeout(resolve, 100));
            }
        }
        
        pdf.destroy();
        
        // 如果只有一页，直接返回
        if (images.length === 1) {
            return {
                blob: images[0],
                filename: this.getOutputFilename(file.name, 'png'),
                type: 'image/png'
            };
        }
        
        // 多页打包成ZIP
        const zip = new JSZip();
        images.forEach((blob, index) => {
            zip.file(`page_${index + 1}.png`, blob);
        });
        
        const zipBlob = await zip.generateAsync({ type: 'blob' });
        
        return {
            blob: zipBlob,
            filename: this.getOutputFilename(file.name, 'zip'),
            type: 'application/zip'
        };
    }

    // ============ PPT 转换方法 ============

    async convertPptxToPdf(file) {
        const arrayBuffer = await this.fileToArrayBuffer(file);
        const zip = await JSZip.loadAsync(arrayBuffer);
        
        // 获取所有幻灯片
        const slideFiles = Object.keys(zip.files)
            .filter(name => name.match(/ppt\/slides\/slide\d+\.xml/))
            .sort((a, b) => {
                const numA = parseInt(a.match(/slide(\d+)\.xml/)[1]);
                const numB = parseInt(b.match(/slide(\d+)\.xml/)[1]);
                return numA - numB;
            });
        
        // 检查幻灯片数量
        if (slideFiles.length > this.maxPptSlides) {
            throw new Error(`幻灯片过多！最多支持 ${this.maxPptSlides} 页，当前 ${slideFiles.length} 页`);
        }
        
        // 创建PDF
        const { PDFDocument } = PDFLib;
        const pdfDoc = await PDFDocument.create();
        
        for (const slideFile of slideFiles) {
            const slideXml = await zip.file(slideFile).async('text');
            const textMatches = slideXml.match(/<a:t>([^<]*)<\/a:t>/g) || [];
            const texts = textMatches.map(match => 
                match.replace(/<a:t>/g, '').replace(/<\/a:t>/g, '')
            );
            
            const page = pdfDoc.addPage([960, 540]);
            
            let y = 480;
            texts.forEach(text => {
                if (y > 50 && text.trim()) {
                    page.drawText(text.substring(0, 100), {
                        x: 50,
                        y: y,
                        size: 24,
                        color: PDFLib.rgb(0.2, 0.2, 0.2)
                    });
                    y -= 40;
                }
            });
        }
        
        const pdfBytes = await pdfDoc.save();
        
        return {
            blob: new Blob([pdfBytes], { type: 'application/pdf' }),
            filename: this.getOutputFilename(file.name, 'pdf'),
            type: 'application/pdf'
        };
    }

    async convertPptxToImages(file) {
        const arrayBuffer = await this.fileToArrayBuffer(file);
        const zip = await JSZip.loadAsync(arrayBuffer);
        
        const mediaFiles = Object.keys(zip.files)
            .filter(name => name.startsWith('ppt/media/'))
            .filter(name => /\.(png|jpg|jpeg)$/i.test(name));
        
        if (mediaFiles.length === 0) {
            throw new Error('PPT中没有找到图片，可能是一个纯文本的PPT');
        }
        
        // 检查图片数量
        if (mediaFiles.length > this.maxPptSlides) {
            throw new Error(`图片过多！最多支持 ${this.maxPptSlides} 张，当前 ${mediaFiles.length} 张`);
        }
        
        if (mediaFiles.length === 1) {
            const imageData = await zip.file(mediaFiles[0]).async('blob');
            const ext = mediaFiles[0].split('.').pop().toLowerCase();
            return {
                blob: imageData,
                filename: this.getOutputFilename(file.name, ext),
                type: `image/${ext === 'png' ? 'png' : 'jpeg'}`
            };
        }
        
        // 多个图片打包成ZIP
        const imageZip = new JSZip();
        for (let i = 0; i < mediaFiles.length; i++) {
            const imageData = await zip.file(mediaFiles[i]).async('blob');
            const ext = mediaFiles[i].split('.').pop().toLowerCase();
            imageZip.file(`slide_${i + 1}.${ext}`, imageData);
        }
        
        const zipBlob = await imageZip.generateAsync({ type: 'blob' });
        
        return {
            blob: zipBlob,
            filename: this.getOutputFilename(file.name, 'zip'),
            type: 'application/zip'
        };
    }

    // ============ Excel 转换方法 ============

    async convertExcelToPdf(file) {
        // 读取Excel
        const arrayBuffer = await this.fileToArrayBuffer(file);
        const data = new Uint8Array(arrayBuffer);
        
        // 使用 SheetJS 解析
        const workbook = XLSX.read(data, { type: 'array' });
        
        // 获取第一个工作表
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // 转换为HTML表格
        const html = XLSX.utils.sheet_to_html(worksheet);
        
        // 包装成完整HTML
        const fullHtml = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>${this.escapeHtml(file.name)}</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+SC&display=swap');
        body { 
            font-family: "Noto Sans SC", Arial, sans-serif; 
            padding: 40px;
        }
        table { 
            border-collapse: collapse; 
            width: 100%;
        }
        th, td { 
            border: 1px solid #ddd; 
            padding: 8px; 
            text-align: left;
        }
        th { 
            background: #f5f5f5; 
            font-weight: bold;
        }
    </style>
</head>
<body>${html}</body>
</html>`;
        
        // 使用 html2pdf 生成PDF
        const element = document.createElement('div');
        element.innerHTML = fullHtml;
        element.style.position = 'absolute';
        element.style.left = '-9999px';
        document.body.appendChild(element);
        
        try {
            const opt = {
                margin: 10,
                filename: this.getOutputFilename(file.name, 'pdf'),
                image: { type: 'jpeg', quality: 0.98 },
                html2canvas: { scale: 2, useCORS: true },
                jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape' }
            };
            
            const blob = await html2pdf().set(opt).from(element).output('blob');
            
            return {
                blob: blob,
                filename: this.getOutputFilename(file.name, 'pdf'),
                type: 'application/pdf'
            };
        } finally {
            document.body.removeChild(element);
        }
    }

    async convertExcelToCsv(file) {
        const arrayBuffer = await this.fileToArrayBuffer(file);
        const data = new Uint8Array(arrayBuffer);
        
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        const csv = XLSX.utils.sheet_to_csv(worksheet);
        
        return {
            blob: new Blob([csv], { type: 'text/csv' }),
            filename: this.getOutputFilename(file.name, 'csv'),
            type: 'text/csv'
        };
    }

    async convertExcelToJson(file) {
        const arrayBuffer = await this.fileToArrayBuffer(file);
        const data = new Uint8Array(arrayBuffer);
        
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        const json = XLSX.utils.sheet_to_json(worksheet);
        const jsonStr = JSON.stringify(json, null, 2);
        
        return {
            blob: new Blob([jsonStr], { type: 'application/json' }),
            filename: this.getOutputFilename(file.name, 'json'),
            type: 'application/json'
        };
    }

    // ============ 图片转换方法 ============

    async convertImageToPdf(file) {
        return new Promise((resolve, reject) => {
            const img = new Image();
            const url = URL.createObjectURL(file);
            
            img.onload = async () => {
                try {
                    const { PDFDocument } = PDFLib;
                    const pdfDoc = await PDFDocument.create();
                    
                    const page = pdfDoc.addPage([img.width, img.height]);
                    
                    // 读取图片数据
                    const imageData = await file.arrayBuffer();
                    let pdfImage;
                    
                    if (file.type === 'image/png') {
                        pdfImage = await pdfDoc.embedPng(imageData);
                    } else {
                        pdfImage = await pdfDoc.embedJpg(imageData);
                    }
                    
                    page.drawImage(pdfImage, {
                        x: 0,
                        y: 0,
                        width: img.width,
                        height: img.height
                    });
                    
                    const pdfBytes = await pdfDoc.save();
                    
                    resolve({
                        blob: new Blob([pdfBytes], { type: 'application/pdf' }),
                        filename: this.getOutputFilename(file.name, 'pdf'),
                        type: 'application/pdf'
                    });
                } catch (error) {
                    reject(error);
                } finally {
                    URL.revokeObjectURL(url);
                }
            };
            
            img.onerror = () => {
                URL.revokeObjectURL(url);
                reject(new Error('图片加载失败'));
            };
            
            img.src = url;
        });
    }

    async convertImageFormat(file, format) {
        return new Promise((resolve, reject) => {
            const img = new Image();
            const url = URL.createObjectURL(file);
            
            img.onload = () => {
                const canvas = document.createElement('canvas');
                canvas.width = img.width;
                canvas.height = img.height;
                
                const ctx = canvas.getContext('2d');
                ctx.drawImage(img, 0, 0);
                
                canvas.toBlob((blob) => {
                    URL.revokeObjectURL(url);
                    
                    resolve({
                        blob: blob,
                        filename: this.getOutputFilename(file.name, format),
                        type: `image/${format}`
                    });
                }, `image/${format}`);
            };
            
            img.onerror = () => {
                URL.revokeObjectURL(url);
                reject(new Error('图片转换失败'));
            };
            
            img.src = url;
        });
    }

    // ============ 工具方法 ============

    showLoading(show) {
        document.getElementById('loadingOverlay').classList.toggle('show', show);
    }

    showResult(result) {
        const downloadBtn = document.getElementById('downloadBtn');
        const resultArea = document.getElementById('resultArea');
        
        if (this.currentDownloadUrl) {
            URL.revokeObjectURL(this.currentDownloadUrl);
        }
        
        this.currentDownloadUrl = URL.createObjectURL(result.blob);
        
        downloadBtn.onclick = () => {
            const a = document.createElement('a');
            a.href = this.currentDownloadUrl;
            a.download = result.filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        };
        
        resultArea.classList.add('show');
    }

    clearFile() {
        this.currentFile = null;
        this.currentFormat = null;
        
        document.getElementById('fileInfo').classList.remove('show');
        document.getElementById('formatOptions').classList.remove('show');
        document.getElementById('resultArea').classList.remove('show');
        document.getElementById('convertBtn').disabled = true;
        document.getElementById('fileInput').value = '';
        
        if (this.currentDownloadUrl) {
            URL.revokeObjectURL(this.currentDownloadUrl);
            this.currentDownloadUrl = null;
        }
    }

    showError(message) {
        const errorEl = document.getElementById('errorMessage');
        errorEl.textContent = message;
        errorEl.classList.add('show');
        
        if (this.errorTimeout) {
            clearTimeout(this.errorTimeout);
        }
        this.errorTimeout = setTimeout(() => this.hideError(), 8000);
    }

    hideError() {
        document.getElementById('errorMessage').classList.remove('show');
    }

    getErrorMessage(error) {
        if (!error || !error.message) {
            return '转换失败，请重试';
        }

        const msg = error.message.toLowerCase();

        if (msg.includes('pdf.js') || msg.includes('未加载')) {
            return 'PDF.js 库加载失败，请刷新页面重试。如果问题持续，请检查网络连接。';
        }

        if (msg.includes('network') || msg.includes('fetch') || msg.includes('load')) {
            return '网络连接问题或库加载失败，请检查网络后刷新页面重试';
        }

        if (msg.includes('password') || msg.includes('encrypted')) {
            return '文件已加密，无法转换';
        }

        if (msg.includes('corrupt') || msg.includes('invalid') || msg.includes('parse')) {
            return '文件已损坏或格式无效，请检查文件';
        }

        if (msg.includes('timeout')) {
            return '转换超时，请尝试较小的文件或检查网络';
        }

        if (msg.includes('memory') || msg.includes('quota')) {
            return '文件太大，内存不足，请尝试更小的文件';
        }

        if (msg.includes('empty')) {
            return '文件为空，请选择其他文件';
        }

        return error.message;
    }

    getOutputFilename(filename, ext) {
        const baseName = filename.replace(/\.[^/.]+$/, '');
        return `${baseName}_converted.${ext}`;
    }

    formatSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    fileToArrayBuffer(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = () => reject(new Error('文件读取失败'));
            reader.readAsArrayBuffer(file);
        });
    }

    escapeHtml(str) {
        return str
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    }

    createPdfHtmlDocument(content, title) {
        return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>${this.escapeHtml(title)}</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@400;700&display=swap');
        body { 
            font-family: "Noto Sans SC", Arial, sans-serif; 
            padding: 40px;
            line-height: 1.6;
            color: #333;
        }
        h1, h2, h3 { color: #222; margin: 20px 0 10px; }
        p { margin: 10px 0; }
        table { border-collapse: collapse; width: 100%; margin: 10px 0; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background: #f5f5f5; font-weight: bold; }
        ul, ol { margin: 10px 0; padding-left: 30px; }
    </style>
</head>
<body>${content}</body>
</html>`;
    }
}

// 初始化
document.addEventListener('DOMContentLoaded', () => {
    try {
        window.docConverter = new DocConverterPro();
        console.log('✅ 文档转换工具 Pro v2.1 已加载');
        console.log('🚀 新增功能: Excel支持、图片转PDF、PDF转图片');
    } catch (error) {
        console.error('❌ 初始化失败:', error);
        alert('工具加载失败，请刷新页面重试');
    }
});
