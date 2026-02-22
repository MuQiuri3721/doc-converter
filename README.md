# 📄 文档转换工具 Pro v2.1

一个纯前端实现的文档格式转换工具，支持 Word、PDF、PPT 格式互转。

## ✨ 特性

- 🔒 **安全可靠** - 本地转换，文件不上传服务器
- ⚡ **快速转换** - 浏览器本地处理，秒级完成
- 🆓 **完全免费** - 无需注册，无使用限制
- 📱 **响应式设计** - 支持桌面和移动设备
- 🌐 **中文优化** - 完整的中文支持，无乱码

## 🚀 支持格式

| 源格式 | 目标格式 |
|--------|----------|
| Word (.docx) | PDF, HTML, TXT |
| PDF (.pdf) | Word, HTML, TXT |
| PPT (.pptx) | PDF, 图片集 |

## 📦 文件结构

```
doc-converter/
├── index.html      # 主页面
├── converter.js    # 转换逻辑
└── README.md       # 说明文档
```

## 🛠️ 技术栈

- **PDF 处理**: PDF.js, pdf-lib
- **Word 处理**: Mammoth.js
- **PPT 处理**: JSZip
- **PDF 生成**: html2pdf.js
- **样式**: 纯 CSS3

## 🚀 部署到 GitHub Pages

### 方法 1：直接上传

1. 下载本项目文件
2. 进入你的 GitHub 仓库
3. 上传 `index.html` 和 `converter.js`
4. 启用 GitHub Pages

### 方法 2：Git 命令行

```bash
# 克隆仓库
git clone https://github.com/MuQiuri3721/doc-converter.git
cd doc-converter

# 替换文件
# 将新的 index.html 和 converter.js 复制到此目录

# 提交更改
git add .
git commit -m "🎉 更新到 v2.0 Pro 版本"
git push origin main
```

## 📝 更新日志

### v2.1 Pro
- ✨ 新增 Excel 支持（XLSX/XLS 转 PDF/CSV/JSON）
- 🖼️ 新增图片转 PDF 功能
- 📄 新增 PDF 转图片功能
- 🚀 优化转换性能
- 🐛 修复中文乱码问题
- 📱 改进移动端适配
- 🎨 添加动画效果
- 🛡️ 增强错误处理

### v2.0 Pro
- ✨ 全新 UI 设计
- 🚀 优化转换性能
- 🐛 修复中文乱码问题
- 📱 改进移动端适配
- 🎨 添加动画效果
- 🛡️ 增强错误处理

### v1.0
- 🎉 初始版本发布
- 支持基本格式转换

## 💕 致谢

Made with love by 02

## 📄 许可证

MIT License
