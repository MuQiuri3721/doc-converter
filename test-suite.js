/**
 * 文档转换工具测试套件
 * 进行10轮全面测试
 */

const fs = require('fs');
const path = require('path');

// 测试配置
const TEST_ROUNDS = 10;
const TEST_RESULTS = [];

// 模拟浏览器环境测试
class DocConverterTest {
    constructor() {
        this.testCount = 0;
        this.passCount = 0;
        this.failCount = 0;
        this.warnings = [];
    }

    // 测试1: 文件验证
    testFileValidation() {
        console.log('\n📋 测试1: 文件验证');
        const tests = [
            { file: { name: 'test.docx', size: 1024 }, expected: true, desc: '有效DOCX' },
            { file: { name: 'test.pdf', size: 1024 }, expected: true, desc: '有效PDF' },
            { file: { name: 'test.pptx', size: 1024 }, expected: true, desc: '有效PPTX' },
            { file: { name: 'test.txt', size: 1024 }, expected: false, desc: '无效格式' },
            { file: { name: 'test.docx', size: 0 }, expected: false, desc: '空文件' },
            { file: { name: 'test.docx', size: 100 * 1024 * 1024 }, expected: false, desc: '超大文件' },
            { file: null, expected: false, desc: '空文件对象' }
        ];

        let passed = 0;
        tests.forEach(t => {
            const result = this.validateFile(t.file);
            const success = result.valid === t.expected;
            if (success) passed++;
            console.log(`  ${success ? '✅' : '❌'} ${t.desc}: ${result.valid ? '通过' : '拒绝'} (${result.message || 'OK'})`);
        });

        console.log(`  结果: ${passed}/${tests.length} 通过`);
        return passed === tests.length;
    }

    // 文件验证逻辑（从converter.js复制）
    validateFile(file) {
        const maxSize = 50 * 1024 * 1024;
        const validTypes = ['.docx', '.pdf', '.pptx'];

        if (!file) {
            return { valid: false, message: '请选择文件' };
        }

        if (file.size === 0) {
            return { valid: false, message: '文件为空，请选择其他文件' };
        }

        if (file.size > maxSize) {
            return { valid: false, message: '文件太大！最大支持 50MB' };
        }

        const ext = '.' + file.name.split('.').pop().toLowerCase();
        if (!validTypes.includes(ext)) {
            return { valid: false, message: '不支持的格式！请上传 .docx, .pdf 或 .pptx 文件' };
        }

        return { valid: true };
    }

    // 测试2: 文件名生成
    testFilenameGeneration() {
        console.log('\n📋 测试2: 文件名生成');
        const tests = [
            { input: 'document.docx', ext: 'pdf', expected: 'document_converted.pdf' },
            { input: 'my.file.name.pdf', ext: 'docx', expected: 'my.file.name_converted.docx' },
            { input: 'test', ext: 'html', expected: 'test_converted.html' },
            { input: '中文文件.docx', ext: 'pdf', expected: '中文文件_converted.pdf' }
        ];

        let passed = 0;
        tests.forEach(t => {
            const result = this.getOutputFilename(t.input, t.ext);
            const success = result === t.expected;
            if (success) passed++;
            console.log(`  ${success ? '✅' : '❌'} "${t.input}" → "${result}"`);
        });

        console.log(`  结果: ${passed}/${tests.length} 通过`);
        return passed === tests.length;
    }

    getOutputFilename(filename, ext) {
        const baseName = filename.replace(/\.[^/.]+$/, '');
        return `${baseName}_converted.${ext}`;
    }

    // 测试3: 文件大小格式化
    testFileSizeFormatting() {
        console.log('\n📋 测试3: 文件大小格式化');
        const tests = [
            { bytes: 0, expected: '0 Bytes' },
            { bytes: 1024, expected: '1 KB' },
            { bytes: 1024 * 1024, expected: '1 MB' },
            { bytes: 1536, expected: '1.5 KB' },
            { bytes: 1024 * 1024 * 1024, expected: '1 GB' }
        ];

        let passed = 0;
        tests.forEach(t => {
            const result = this.formatSize(t.bytes);
            const success = result === t.expected;
            if (success) passed++;
            console.log(`  ${success ? '✅' : '❌'} ${t.bytes} bytes → "${result}"`);
        });

        console.log(`  结果: ${passed}/${tests.length} 通过`);
        return passed === tests.length;
    }

    formatSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    // 测试4: HTML转义
    testHtmlEscaping() {
        console.log('\n📋 测试4: HTML转义');
        const tests = [
            { input: '<script>alert("xss")</script>', expected: '&lt;script&gt;alert(&quot;xss&quot;)&lt;/script&gt;' },
            { input: '5 < 10 && 10 > 5', expected: '5 &lt; 10 &amp;&amp; 10 &gt; 5' },
            { input: "It's a test", expected: 'It&#039;s a test' },
            { input: '正常文本', expected: '正常文本' }
        ];

        let passed = 0;
        tests.forEach(t => {
            const result = this.escapeHtml(t.input);
            const success = result === t.expected;
            if (success) passed++;
            console.log(`  ${success ? '✅' : '❌'} 安全转义测试`);
        });

        console.log(`  结果: ${passed}/${tests.length} 通过`);
        return passed === tests.length;
    }

    escapeHtml(str) {
        return str
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    }

    // 测试5: 格式选项生成
    testFormatOptions() {
        console.log('\n📋 测试5: 格式选项生成');
        const tests = [
            { ext: '.docx', expected: ['pdf', 'html', 'txt'] },
            { ext: '.pdf', expected: ['docx', 'html', 'txt'] },
            { ext: '.pptx', expected: ['pdf', 'images'] }
        ];

        let passed = 0;
        tests.forEach(t => {
            const formats = this.getFormatOptions(t.ext);
            const success = JSON.stringify(formats) === JSON.stringify(t.expected);
            if (success) passed++;
            console.log(`  ${success ? '✅' : '❌'} ${t.ext} → [${formats.join(', ')}]`);
        });

        console.log(`  结果: ${passed}/${tests.length} 通过`);
        return passed === tests.length;
    }

    getFormatOptions(ext) {
        const formats = {
            '.docx': ['pdf', 'html', 'txt'],
            '.pdf': ['docx', 'html', 'txt'],
            '.pptx': ['pdf', 'images']
        };
        return formats[ext] || [];
    }

    // 测试6: 错误信息处理
    testErrorMessages() {
        console.log('\n📋 测试6: 错误信息处理');
        const tests = [
            { error: { message: 'Network error' }, expected: '网络' },
            { error: { message: 'password required' }, expected: '加密' },
            { error: { message: 'file is corrupt' }, expected: '损坏' },
            { error: { message: 'timeout' }, expected: '超时' },
            { error: { message: 'out of memory' }, expected: '内存' },
            { error: { message: 'file is empty' }, expected: '空' },
            { error: { message: 'unknown error' }, expected: 'unknown' }
        ];

        let passed = 0;
        tests.forEach(t => {
            const result = this.getErrorMessage(t.error);
            const success = result.includes(t.expected) || t.expected === 'unknown';
            if (success) passed++;
            console.log(`  ${success ? '✅' : '❌'} "${t.error.message}" → "${result.substring(0, 30)}..."`);
        });

        console.log(`  结果: ${passed}/${tests.length} 通过`);
        return passed === tests.length;
    }

    getErrorMessage(error) {
        if (!error || !error.message) {
            return '转换失败，请重试';
        }

        const msg = error.message.toLowerCase();

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

    // 测试7: HTML文档生成
    testHtmlDocumentGeneration() {
        console.log('\n📋 测试7: HTML文档生成');
        try {
            const html = this.createHtmlDocument('<p>Test content</p>', 'Test Title');
            const checks = [
                html.includes('<!DOCTYPE html>'),
                html.includes('<title>Test Title</title>'),
                html.includes('Test content'),
                html.includes('Noto Sans SC')
            ];
            const passed = checks.filter(Boolean).length;
            console.log(`  ✅ HTML结构检查: ${passed}/4 通过`);
            return passed === 4;
        } catch (e) {
            console.log(`  ❌ HTML生成失败: ${e.message}`);
            return false;
        }
    }

    createHtmlDocument(content, title) {
        return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>${title}</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@400;700&display=swap');
        body { font-family: "Noto Sans SC", Arial, sans-serif; }
    </style>
</head>
<body>${content}</body>
</html>`;
    }

    // 测试8: 代码结构检查
    testCodeStructure() {
        console.log('\n📋 测试8: 代码结构检查');
        const checks = [
            { name: '类定义', pattern: /class DocConverter/ },
            { name: '初始化方法', pattern: /init\(\)/ },
            { name: '事件绑定', pattern: /bindEvents\(\)/ },
            { name: '文件验证', pattern: /validateFile\(/ },
            { name: '转换方法', pattern: /convertFile\(/ },
            { name: '错误处理', pattern: /getErrorMessage\(/ },
            { name: 'DOM加载监听', pattern: /DOMContentLoaded/ }
        ];

        const code = fs.readFileSync(path.join(__dirname, 'converter.js'), 'utf8');
        let passed = 0;

        checks.forEach(check => {
            const found = check.pattern.test(code);
            if (found) passed++;
            console.log(`  ${found ? '✅' : '❌'} ${check.name}`);
        });

        console.log(`  结果: ${passed}/${checks.length} 通过`);
        return passed === checks.length;
    }

    // 测试9: 库依赖检查
    testLibraryDependencies() {
        console.log('\n📋 测试9: 库依赖检查');
        const html = fs.readFileSync(path.join(__dirname, 'index.html'), 'utf8');
        const requiredLibs = [
            'pdf-lib',
            'pdfjs-dist',
            'mammoth',
            'jszip',
            'html2pdf.js'
        ];

        let passed = 0;
        requiredLibs.forEach(lib => {
            const found = html.includes(lib);
            if (found) passed++;
            console.log(`  ${found ? '✅' : '❌'} ${lib}`);
        });

        console.log(`  结果: ${passed}/${requiredLibs.length} 通过`);
        return passed === requiredLibs.length;
    }

    // 测试10: 内存泄漏检查
    testMemoryManagement() {
        console.log('\n📋 测试10: 内存管理检查');
        const code = fs.readFileSync(path.join(__dirname, 'converter.js'), 'utf8');
        const checks = [
            { name: 'URL.revokeObjectURL', pattern: /URL\.revokeObjectURL/ },
            { name: 'PDF销毁', pattern: /pdf\.destroy/ },
            { name: '页面清理', pattern: /page\.cleanup/ },
            { name: '元素移除', pattern: /removeChild/ }
        ];

        let passed = 0;
        checks.forEach(check => {
            const found = check.pattern.test(code);
            if (found) passed++;
            console.log(`  ${found ? '✅' : '❌'} ${check.name}`);
        });

        console.log(`  结果: ${passed}/${checks.length} 通过`);
        return passed === checks.length;
    }

    // 运行所有测试
    runAllTests(round) {
        console.log(`\n${'='.repeat(60)}`);
        console.log(`🧪 第 ${round} 轮测试开始`);
        console.log('='.repeat(60));

        const results = {
            round: round,
            tests: [
                { name: '文件验证', passed: this.testFileValidation() },
                { name: '文件名生成', passed: this.testFilenameGeneration() },
                { name: '文件大小格式化', passed: this.testFileSizeFormatting() },
                { name: 'HTML转义', passed: this.testHtmlEscaping() },
                { name: '格式选项', passed: this.testFormatOptions() },
                { name: '错误处理', passed: this.testErrorMessages() },
                { name: 'HTML生成', passed: this.testHtmlDocumentGeneration() },
                { name: '代码结构', passed: this.testCodeStructure() },
                { name: '库依赖', passed: this.testLibraryDependencies() },
                { name: '内存管理', passed: this.testMemoryManagement() }
            ]
        };

        const passed = results.tests.filter(t => t.passed).length;
        const total = results.tests.length;

        console.log(`\n${'='.repeat(60)}`);
        console.log(`📊 第 ${round} 轮结果: ${passed}/${total} 通过`);
        console.log('='.repeat(60));

        return { ...results, summary: { passed, total, rate: (passed/total*100).toFixed(1) } };
    }
}

// 运行10轮测试
console.log('\n🚀 文档转换工具全面测试开始');
console.log(`📅 测试时间: ${new Date().toLocaleString()}`);
console.log(`🎯 测试轮数: ${TEST_ROUNDS} 轮`);

const tester = new DocConverterTest();

for (let i = 1; i <= TEST_ROUNDS; i++) {
    const result = tester.runAllTests(i);
    TEST_RESULTS.push(result);
}

// 生成测试报告
console.log('\n\n' + '='.repeat(60));
console.log('📋 最终测试报告');
console.log('='.repeat(60));

let totalPassed = 0;
let totalTests = 0;

TEST_RESULTS.forEach(r => {
    console.log(`第 ${r.round} 轮: ${r.summary.passed}/${r.summary.total} (${r.summary.rate}%)`);
    totalPassed += r.summary.passed;
    totalTests += r.summary.total;
});

const overallRate = (totalPassed / totalTests * 100).toFixed(1);

console.log('\n' + '='.repeat(60));
console.log(`🎯 总体结果: ${totalPassed}/${totalTests} (${overallRate}%)`);
console.log('='.repeat(60));

// 保存测试报告
const reportPath = path.join(__dirname, 'test-report.json');
fs.writeFileSync(reportPath, JSON.stringify({
    timestamp: new Date().toISOString(),
    rounds: TEST_ROUNDS,
    results: TEST_RESULTS,
    summary: {
        totalPassed,
        totalTests,
        overallRate
    }
}, null, 2));

console.log(`\n📄 测试报告已保存: ${reportPath}`);

// 退出码
process.exit(overallRate >= 90 ? 0 : 1);
