#!/bin/bash
# 文档转换工具 Pro v2.0 - 部署脚本
# 使用方法: ./deploy.sh

set -e

echo "🚀 开始部署文档转换工具 Pro v2.0..."
echo ""

# 颜色定义
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# 检查 Git
if ! command -v git &> /dev/null; then
    echo "${RED}❌ 错误: 未安装 Git${NC}"
    echo "请先安装 Git: https://git-scm.com/downloads"
    exit 1
fi

# 配置信息
REPO_URL="https://github.com/MuQiuri3721/doc-converter.git"
REPO_NAME="doc-converter"

# 检查是否已有仓库
echo "${BLUE}📁 检查仓库...${NC}"
if [ -d "$REPO_NAME" ]; then
    echo "${YELLOW}⚠️  发现已有仓库，更新中...${NC}"
    cd "$REPO_NAME"
    git pull origin main || echo "${YELLOW}⚠️  拉取更新失败，继续部署...${NC}"
    cd ..
else
    echo "${BLUE}📥 克隆仓库...${NC}"
    git clone "$REPO_URL" || {
        echo "${RED}❌ 克隆失败，请检查仓库地址和权限${NC}"
        exit 1
    }
fi

cd "$REPO_NAME"

# 备份旧版本
echo "${BLUE}💾 备份旧版本...${NC}"
if [ -f "index.html" ]; then
    cp index.html "index.html.backup.$(date +%Y%m%d_%H%M%S)"
    echo "${GREEN}✅ 已备份旧版本${NC}"
fi

# 复制新文件
echo "${BLUE}📋 复制新文件...${NC}"
cp ../index.html . 2>/dev/null || echo "${YELLOW}⚠️  请手动复制 index.html${NC}"
cp ../converter.js . 2>/dev/null || echo "${YELLOW}⚠️  请手动复制 converter.js${NC}"

# Git 配置
echo "${BLUE}⚙️  配置 Git...${NC}"
git config user.name "MuQiuri3721" 2>/dev/null || true
git config user.email "1572206315@qq.com" 2>/dev/null || true

# 提交更改
echo "${BLUE}📤 提交更改...${NC}"
git add .
git commit -m "🎉 更新到 v2.0 Pro 版本

- 全新 UI 设计
- 优化转换性能
- 修复中文乱码问题
- 改进移动端适配
- 添加动画效果
- 增强错误处理" || echo "${YELLOW}⚠️  没有更改需要提交${NC}"

# 推送
echo "${BLUE}🚀 推送到 GitHub...${NC}"
echo "${YELLOW}💡 提示: 如果要求输入密码，请输入你的 GitHub Token${NC}"
git push origin main

echo ""
echo "${GREEN}✅ 部署完成！${NC}"
echo ""
echo "${BLUE}🌐 访问地址:${NC} https://muqiuri3721.github.io/doc-converter/"
echo ""
echo "${YELLOW}📌 注意事项:${NC}"
echo "   - 首次部署后需要 1-2 分钟生效"
echo "   - 如果页面未更新，请清除浏览器缓存"
echo "   - 确保 GitHub Pages 已启用"
echo ""
