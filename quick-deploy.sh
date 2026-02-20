#!/bin/bash
# ============================================
# 文档转换工具 - 快速部署脚本 (简化版)
# 使用方法: ./quick-deploy.sh <你的GitHub Token>
# ============================================

# 颜色
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

# 检查参数
if [ -z "$1" ]; then
    echo -e "${RED}❌ 错误: 需要提供 GitHub Token${NC}"
    echo ""
    echo "使用方法:"
    echo "  ./quick-deploy.sh ghp_xxxxxxxxxxxx"
    echo ""
    echo "获取 Token: https://github.com/settings/tokens"
    exit 1
fi

TOKEN="$1"
REPO="https://${TOKEN}@github.com/MuQiuri3721/doc-converter.git"

echo -e "${YELLOW}🚀 开始快速部署...${NC}"

# 步骤 1: 克隆仓库
echo "📥 克隆仓库..."
if [ -d "doc-converter" ]; then
    rm -rf doc-converter
fi
git clone "$REPO" doc-converter 2>/dev/null || {
    echo -e "${RED}❌ 克隆失败，检查 Token 是否正确${NC}"
    exit 1
}

# 步骤 2: 复制文件
echo "📋 复制新文件..."
cp index.html converter.js doc-converter/

# 步骤 3: 提交并推送
echo "📤 提交更改..."
cd doc-converter
git add .
git commit -m "🎉 Update to v2.0.1 - Bug fixes and improvements" || echo "没有更改需要提交"
git push origin main

# 完成
echo -e "${GREEN}✅ 部署完成!${NC}"
echo ""
echo "🌐 访问地址: https://muqiuri3721.github.io/doc-converter/"
echo ""
echo "💡 提示: 首次部署后需要 1-2 分钟生效"
