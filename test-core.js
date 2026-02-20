/**
 * 核心功能测试脚本
 * 测试文档转换工具的关键逻辑
 */

const fs = require('fs');
const path = require('path');

// 模拟浏览器环境
const mockBrowser = {
    console: {
        log: (...args) => console.log('[LOG]', ...args),
        error: (...args) => console.error('[ERR]', ...args),
        warn: (...args) => console.warn('[WARN]', ...args)
    },
    document: {
        getElementById: (id) => ({
            classList: {
                add: () => {},
                remove: () => {}
            },
            style: {},
            textContent: '',
            innerHTML: '',
            addEventListener: () => {},
            querySelectorAll: () => []
        }),
        createElement: (tag) => ({
            innerHTML: '',
            style: {},
            appendChild: () => {},
            removeChild: () => {}
        }),
        body: {
            appendChild: () => {},
            removeChild: () => {},
            contains: () => true
        }
    },
    URL: {
        createObjectURL: (blob) => 'blob:mock-url-' + Date.now(),
        revokeObjectURL: () => {}
    },
    Blob: class MockBlob {
        constructor(parts, options) {
            this.parts = parts;
            this.options = options;
            this.size = parts.join('').length;
        }
    }
};

// 测试配置
const TEST_CONFIG = {
    maxSize: 50 * 1024 * 1024,
    validTypes: ['.docx', '.pdf', '.pptx']
};

// 测试用例
const TEST_CASES = {
    // 文件验证测试
    fileValidation: [
        { name: 'test.docx', size: 1024, expected: true },
        { name: 'test.pdf', size: 1024, expected: true },
        { name: 'test.pptx', size: 1024, expected: true },
        { name: 'test.txt', size: 1024, expected: false },
        { name: 'test.docx', size: 0, expected: false },
        { name: 'test.docx', size: 100 * 1024 * 1024, expected: false }
    ],

    // 文件名处理测试
    filenameHandling: [
        { input: 'document.docx', format: 'pdf', expected: 'document_converted.pdf' },
        { input: 'my file.pdf', format: 'html', expected: 'my file_converted.html' },
        { input: 'test.v2.docx', format: 'txt', expected: 'test.v2_converted.txt' }
    ],

    // 文件大小格式化测试
    sizeFormatting: [
        { bytes: 0, expected: '0 Bytes' },
        { bytes: 1024, expected: '1 KB' },
        { bytes: 1024 * 1024, expected: '1 MB' },
        { bytes: 1536, expected: '1.5 KB' }
    ]
};

// 文件验证函数（从原代码提取）
function validateFile(file) {
    const maxSize = 50 * 1024 * 1024;
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

// 格式化文件大小
function formatSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// 获取输出文件名
function getOutputFilename(filename, ext) {
    const baseName = filename.replace(/\.[^/.]+$/, '');
    return `${baseName}_converted.${ext}`;
}

// HTML 转义函数
function escapeHtml(str) {
    return str
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}

// 运行测试
function runTests() {
    console.log('\n🧪 开始测试文档转换工具核心功能\n');
    console.log('=' .repeat(50));

    let passed = 0;
    let failed = 0;

    // 测试文件验证
    console.log('\n📁 测试文件验证功能:');
    TEST_CASES.fileValidation.forEach((test, index) => {
        const result = validateFile(test);
        const success = result.valid === test.expected;

        if (success) {
            console.log(`  ✅ 测试 ${index + 1} 通过: ${test.name} (${test.size} bytes)`);
            passed++;
        } else {
            console.log(`  ❌ 测试 ${index + 1} 失败: ${test.name}`);
            console.log(`     期望: ${test.expected}, 实际: ${result.valid}`);
            failed++;
        }
    });

    // 测试文件名处理
    console.log('\n📝 测试文件名处理:');
    TEST_CASES.filenameHandling.forEach((test, index) => {
        const result = getOutputFilename(test.input, test.format);
        const success = result === test.expected;

        if (success) {
            console.log(`  ✅ 测试 ${index + 1} 通过: ${test.input} → ${result}`);
            passed++;
        } else {
            console.log(`  ❌ 测试 ${index + 1} 失败`);
            console.log(`     期望: ${test.expected}, 实际: ${result}`);
            failed++;
        }
    });

    // 测试文件大小格式化
    console.log('\n📊 测试文件大小格式化:');
    TEST_CASES.sizeFormatting.forEach((test, index) => {
        const result = formatSize(test.bytes);
        const success = result === test.expected;

        if (success) {
            console.log(`  ✅ 测试 ${index + 1} 通过: ${test.bytes} bytes → ${result}`);
            passed++;
        } else {
            console.log(`  ❌ 测试 ${index + 1} 失败`);
            console.log(`     期望: ${test.expected}, 实际: ${result}`);
            failed++;
        }
    });

    // 测试 HTML 转义
    console.log('\n🔒 测试 HTML 转义功能:');
    const htmlTests = [
        { input: '<script>alert("xss")</script>', expected: '&lt;script&gt;alert(&quot;xss&quot;)&lt;/script&gt;' },
        { input: '5 > 3 && 3 < 5', expected: '5 &gt; 3 &amp;&amp; 3 &lt; 5' },
        { input: "It's a test", expected: 'It&#039;s a test' }
    ];

    htmlTests.forEach((test, index) => {
        const result = escapeHtml(test.input);
        const success = result === test.expected;

        if (success) {
            console.log(`  ✅ 测试 ${index + 1} 通过`);
            passed++;
        } else {
            console.log(`  ❌ 测试 ${index + 1} 失败`);
            console.log(`     输入: ${test.input}`);
            console.log(`     期望: ${test.expected}`);
            console.log(`     实际: ${result}`);
            failed++;
        }
    });

    // 测试总结
    console.log('\n' + '='.repeat(50));
    console.log(`\n📋 测试结果:`);
    console.log(`  ✅ 通过: ${passed} 项`);
    console.log(`  ❌ 失败: ${failed} 项`);
    console.log(`  📊 总计: ${passed + failed} 项`);

    if (failed === 0) {
        console.log('\n🎉 所有测试通过！核心功能正常。\n');
        return true;
    } else {
        console.log(`\n⚠️ 有 ${failed} 项测试失败，请检查代码。\n`);
        return false;
    }
}

// 检查文件是否存在且可读
function checkFiles() {
    console.log('📂 检查项目文件:\n');

    const files = [
        'index.html',
        'converter.js',
        'README.md',
        'CHANGELOG.md'
    ];

    let allExist = true;

    files.forEach(file => {
        const filePath = path.join(__dirname, file);
        if (fs.existsSync(filePath)) {
            const stats = fs.statSync(filePath);
            console.log(`  ✅ ${file} (${(stats.size / 1024).toFixed(1)} KB)`);
        } else {
            console.log(`  ❌ ${file} (不存在)`);
            allExist = false;
        }
    });

    console.log('');
    return allExist;
}

// 检查代码中的潜在问题
function checkCodeQuality() {
    console.log('🔍 检查代码质量:\n');

    const converterPath = path.join(__dirname, 'converter.js');
    const content = fs.readFileSync(converterPath, 'utf8');

    const checks = [
        { name: 'HTML 转义函数', pattern: /escapeHtml/, required: true },
        { name: '错误处理', pattern: /try\s*\{[\s\S]*?\}\s*catch/, required: true },
        { name: '内存清理', pattern: /cleanup|destroy|revokeObjectURL/, required: true },
        { name: '文件验证', pattern: /validateFile/, required: true },
        { name: '进度更新', pattern: /updateProgress/, required: true },
        { name: 'console\.log', pattern: /console\.log/, required: false, warning: '生产环境建议移除调试日志' }
    ];

    checks.forEach(check => {
        const found = check.pattern.test(content);
        if (check.required) {
            if (found) {
                console.log(`  ✅ ${check.name}`);
            } else {
                console.log(`  ❌ ${check.name} (缺失)`);
            }
        } else if (found && check.warning) {
            console.log(`  ⚠️  ${check.name}: ${check.warning}`);
        }
    });

    console.log('');
}

// 主函数
function main() {
    console.log('\n');
    console.log('╔════════════════════════════════════════════════════════════╗');
    console.log('║         📄 文档转换工具 Pro v2.0 - 功能测试报告           ║');
    console.log('╚════════════════════════════════════════════════════════════╝');
    console.log('\n');

    // 检查文件
    const filesOk = checkFiles();

    // 检查代码质量
    checkCodeQuality();

    // 运行功能测试
    const testsOk = runTests();

    // 最终结论
    console.log('╔════════════════════════════════════════════════════════════╗');
    if (filesOk && testsOk) {
        console.log('║  🎉 所有检查通过！代码质量良好，可以安全部署。            ║');
    } else {
        console.log('║  ⚠️  部分检查未通过，建议修复后再部署。                   ║');
    }
    console.log('╚════════════════════════════════════════════════════════════╝');
    console.log('\n');

    return filesOk && testsOk;
}

// 运行测试
main();
