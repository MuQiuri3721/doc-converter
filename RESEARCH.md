# 文档转换技术调研报告

## 1. Word (.docx) → PDF

### 方案对比
| 方案 | 优点 | 缺点 | Stars |
|------|------|------|-------|
| **mammoth + pdfmake** | 语义保留好、表格支持 | 需要额外处理 | 3.2k + 12.2k |
| **mammoth + html2pdf** | 简单直接 | 渲染依赖浏览器 | 3.2k + 1.8k |
| **LibreOffice WebAssembly** | 质量最高 | 体积大、加载慢 | - |
| **pandoc-wasm** | 专业级转换 | 复杂度高 | - |

### 推荐方案
**mammoth + pdfmake** - 平衡质量和复杂度

---

## 2. PDF → Word (.docx)

### 方案对比
| 方案 | 优点 | 缺点 | Stars |
|------|------|------|-------|
| **pdf.js + docx.js** | 纯前端、可控 | 格式有限 | 47k + 3.4k |
| **pdf2docx-wasm** | 质量高 | 体积大 | - |
| **LibreOffice** | 最完整 | 需要后端 | - |

### 推荐方案
**pdf.js + docx.js** - 纯前端方案

---

## 3. PDF → HTML

### 方案对比
| 方案 | 优点 | 缺点 | Stars |
|------|------|------|-------|
| **pdf.js** | 官方方案、稳定 | 需要处理样式 | 47k |
| **pdf2html** | 专用工具 | 需要后端 | - |

### 推荐方案
**pdf.js** - 标准方案

---

## 4. PPT (.pptx) → PDF

### 方案对比
| 方案 | 优点 | 缺点 | Stars |
|------|------|------|-------|
| **JSZip + pdf-lib** | 纯前端 | 只提取文本 | 8.5k + 8.3k |
| **PptxGenJS + 截图** | 可保留样式 | 复杂 | 3.1k |
| **LibreOffice** | 最完整 | 需要后端 | - |

### 推荐方案
**JSZip + pdf-lib** - 简单可用

---

## 5. Excel (.xlsx) → PDF

### 方案对比
| 方案 | 优点 | 缺点 | Stars |
|------|------|------|-------|
| **SheetJS + pdfmake** | 纯前端 | 样式有限 | 33k + 12.2k |
| **xlsx-populate** | 功能丰富 | 体积大 | 1.2k |

### 推荐方案
**SheetJS + pdfmake** - 社区最大

---

## 6. 图片 → PDF

### 方案对比
| 方案 | 优点 | 缺点 | Stars |
|------|------|------|-------|
| **pdf-lib** | 简单直接 | 单张图片 | 8.3k |
| **img2pdf** | 专用工具 | 需要后端 | - |

### 推荐方案
**pdf-lib** - 足够简单

---

## 7. PDF 合并/拆分

### 方案对比
| 方案 | 优点 | 缺点 | Stars |
|------|------|------|-------|
| **pdf-lib** | 功能完整 | 内存占用 | 8.3k |
| **pdf-merger-js** | 专用合并 | 依赖多 | 500+ |

### 推荐方案
**pdf-lib** - 功能最全

---

## 最终技术栈

```
核心库:
├── mammoth.js (3.2k ⭐) - Word解析
├── pdf.js (47k ⭐) - PDF解析
├── pdf-lib (8.3k ⭐) - PDF创建/修改
├── pdfmake (12.2k ⭐) - PDF生成
├── docx.js (3.4k ⭐) - Word生成
├── SheetJS (33k ⭐) - Excel解析
├── JSZip (8.5k ⭐) - ZIP处理
└── html2pdf.js (1.8k ⭐) - 降级方案

总Stars: 117,400+
```

## 推荐实现顺序

1. **Phase 1**: Word ↔ PDF (质量优先)
2. **Phase 2**: PDF 工具 (合并/拆分)
3. **Phase 3**: Excel 支持
4. **Phase 4**: 图片处理
5. **Phase 5**: PPT 改进

