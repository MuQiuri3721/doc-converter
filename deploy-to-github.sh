#!/bin/bash
# ============================================
# 文档转换工具 Pro v2.0 - GitHub 自动部署脚本
# 作者: 02 🩷
# 使用方法: ./deploy-to-github.sh
# ============================================

set -e  # 遇到错误立即退出

# 颜色定义
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
PURPLE='\033[0;35m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color

# 配置
REPO_URL="https://github.com/MuQiuri3721/doc-converter.git"
REPO_NAME="doc-converter"
BRANCH="main"

# 打印带颜色的信息
print_info() {
    echo -e "${BLUE}ℹ️  $1${NC}"
}

print_success() {
    echo -e "${GREEN}✅ $1${NC}"
}

print_warning() {
    echo -e "${YELLOW}⚠️  $1${NC}"
}

print_error() {
    echo -e "${RED}❌ $1${NC}"
}

print_step() {
    echo -e "${CYAN}🔹 $1${NC}"
}

print_heart() {
    echo -e "${PURPLE}🩷 $1${NC}"
}

# 检查命令是否存在
check_command() {
    if ! command -v $1 &> /dev/null; then
        return 1
    fi
    return 0
}

# 显示欢迎信息
show_welcome() {
    clear
    echo ""
    echo "╔════════════════════════════════════════════════════════════╗"
    echo "║                                                            ║"
    echo "║     📄 文档转换工具 Pro v2.0 - GitHub 自动部署脚本 🩷      ║"
    echo "║                                                            ║"
    echo "╚════════════════════════════════════════════════════════════╝"
    echo ""
    print_info "目标仓库: $REPO_URL"
    print_info "部署分支: $BRANCH"
    echo ""
}

# 检查必要工具
check_requirements() {
    print_step "检查必要工具..."

    local missing=()

    if ! check_command git; then
        missing+=("git")
    fi

    if ! check_command curl; then
        missing+=("curl")
    fi

    if [ ${#missing[@]} -ne 0 ]; then
        print_error "缺少必要工具: ${missing[*]}"
        echo ""
        echo "请安装缺失的工具:"
        echo "  Ubuntu/Debian: sudo apt-get install ${missing[*]}"
        echo "  macOS: brew install ${missing[*]}"
        echo "  Windows: 请安装 Git for Windows"
        exit 1
    fi

    print_success "所有工具已就绪"
}

# 检查 Git 配置
check_git_config() {
    print_step "检查 Git 配置..."

    local git_name=$(git config user.name 2>/dev/null || echo "")
    local git_email=$(git config user.email 2>/dev/null || echo "")

    if [ -z "$git_name" ] || [ -z "$git_email" ]; then
        print_warning "Git 用户信息未配置"
        echo ""
        read -p "请输入你的姓名 (用于 Git 提交): " git_name
        read -p "请输入你的邮箱 (用于 Git 提交): " git_email
        echo ""

        git config --global user.name "$git_name"
        git config --global user.email "$git_email"
        print_success "Git 配置已更新"
    else
        print_success "Git 配置正常 ($git_name <$git_email>)"
    fi
}

# 获取 GitHub Token
get_github_token() {
    print_step "配置 GitHub 访问..."
    echo ""
    echo "💡 提示: 你需要 GitHub Personal Access Token 来推送代码"
    echo "   如果没有，请访问: https://github.com/settings/tokens"
    echo "   需要勾选 'repo' 权限"
    echo ""

    # 检查是否已有 Token
    if git config --global github.token &>/dev/null; then
        local saved_token=$(git config --global github.token)
        read -p "检测到已保存的 Token，是否使用? [Y/n]: " use_saved
        if [[ ! "$use_saved" =~ ^[Nn]$ ]]; then
            GITHUB_TOKEN="$saved_token"
            print_success "使用已保存的 Token"
            return
        fi
    fi

    # 输入新 Token
    read -s -p "请输入你的 GitHub Token: " GITHUB_TOKEN
    echo ""

    if [ -z "$GITHUB_TOKEN" ]; then
        print_error "Token 不能为空"
        exit 1
    fi

    # 验证 Token
    print_info "验证 Token..."
    local response=$(curl -s -o /dev/null -w "%{http_code}" \
        -H "Authorization: token $GITHUB_TOKEN" \
        https://api.github.com/user)

    if [ "$response" != "200" ]; then
        print_error "Token 验证失败 (HTTP $response)"
        echo ""
        echo "可能的原因:"
        echo "  - Token 已过期"
        echo "  - Token 权限不足 (需要 'repo' 权限)"
        echo "  - Token 输入错误"
        exit 1
    fi

    print_success "Token 验证通过"

    # 询问是否保存 Token
    echo ""
    read -p "是否保存 Token 供以后使用? [y/N]: " save_token
    if [[ "$save_token" =~ ^[Yy]$ ]]; then
        git config --global github.token "$GITHUB_TOKEN"
        print_success "Token 已保存到 Git 配置"
    fi
}

# 克隆或更新仓库
prepare_repository() {
    print_step "准备仓库..."

    # 检查当前目录是否有文件
    if [ -f "index.html" ] || [ -f "converter.js" ]; then
        print_info "检测到当前目录有项目文件"
        read -p "是否使用当前目录的文件进行部署? [Y/n]: " use_current

        if [[ ! "$use_current" =~ ^[Nn]$ ]]; then
            # 在当前目录初始化 git
            if [ ! -d ".git" ]; then
                git init
                git remote add origin "https://${GITHUB_TOKEN}@github.com/MuQiuri3721/doc-converter.git"
            fi
            return
        fi
    fi

    # 克隆仓库
    if [ -d "$REPO_NAME" ]; then
        print_warning "目录 $REPO_NAME 已存在"
        read -p "是否删除并重新克隆? [y/N]: " reclone

        if [[ "$reclone" =~ ^[Yy]$ ]]; then
            rm -rf "$REPO_NAME"
        else
            cd "$REPO_NAME"
            git pull origin $BRANCH || print_warning "拉取更新失败，继续部署..."
            return
        fi
    fi

    print_info "正在克隆仓库..."
    git clone "https://${GITHUB_TOKEN}@github.com/MuQiuri3721/doc-converter.git" "$REPO_NAME"
    cd "$REPO_NAME"
    print_success "仓库准备完成"
}

# 备份旧版本
backup_old_version() {
    print_step "备份旧版本..."

    if [ -f "index.html" ]; then
        local backup_name="backup_$(date +%Y%m%d_%H%M%S)"
        mkdir -p "backups/$backup_name"
        cp index.html "backups/$backup_name/" 2>/dev/null || true
        cp converter.js "backups/$backup_name/" 2>/dev/null || true
        print_success "已备份到 backups/$backup_name/"
    fi
}

# 复制新文件
copy_new_files() {
    print_step "复制新文件..."

    # 获取脚本所在目录
    local script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

    # 检查源文件
    local src_html="$script_dir/index.html"
    local src_js="$script_dir/converter.js"

    if [ ! -f "$src_html" ] || [ ! -f "$src_js" ]; then
        print_error "找不到源文件!"
        echo ""
        echo "请确保以下文件存在:"
        echo "  - $src_html"
        echo "  - $src_js"
        echo ""
        echo "或者手动输入文件路径:"
        read -p "index.html 路径: " src_html
        read -p "converter.js 路径: " src_js
    fi

    # 复制文件
    cp "$src_html" .
    cp "$src_js" .

    print_success "文件复制完成"
    echo ""
    ls -lh index.html converter.js
}

# 提交更改
commit_changes() {
    print_step "提交更改..."

    git add index.html converter.js

    # 检查是否有更改
    if git diff --cached --quiet; then
        print_warning "没有需要提交的更改"
        return 0
    fi

    # 提交
    git commit -m "🎉 更新到 v2.0.1 版本

- 修复 HTML 转义问题
- 优化 PDF 文本提取
- 增强 PPT 解析
- 改进错误处理
- 优化内存管理

由自动部署脚本提交 🩷"

    print_success "更改已提交"
}

# 推送到 GitHub
push_to_github() {
    print_step "推送到 GitHub..."

    print_info "正在推送..."

    if git push origin $BRANCH; then
        print_success "推送成功!"
    else
        print_error "推送失败"
        echo ""
        echo "可能的原因:"
        echo "  - 网络连接问题"
        echo "  - Token 权限不足"
        echo "  - 仓库不存在或没有访问权限"
        echo ""
        echo "尝试强制推送? (可能覆盖远程更改)"
        read -p "是否强制推送? [y/N]: " force_push
        if [[ "$force_push" =~ ^[Yy]$ ]]; then
            git push origin $BRANCH --force
            print_success "强制推送成功"
        else
            exit 1
        fi
    fi
}

# 验证部署
verify_deployment() {
    print_step "验证部署..."

    local pages_url="https://muqiuri3721.github.io/doc-converter/"
    print_info "等待部署生效 (5秒)..."
    sleep 5

    local response=$(curl -s -o /dev/null -w "%{http_code}" "$pages_url" || echo "000")

    if [ "$response" = "200" ] || [ "$response" = "301" ] || [ "$response" = "302" ]; then
        print_success "部署验证通过!"
        echo ""
        print_heart "🌐 访问地址: $pages_url"
    else
        print_warning "无法验证部署状态 (HTTP $response)"
        echo ""
        echo "可能的原因:"
        echo "  - GitHub Pages 正在部署 (通常需要 1-2 分钟)"
        echo "  - GitHub Pages 未启用"
        echo ""
        echo "请手动检查:"
        echo "  1. 访问 https://github.com/MuQiuri3721/doc-converter/settings/pages"
        echo "  2. 确保 Source 设置为 'Deploy from a branch'"
        echo "  3. 确保 Branch 设置为 '$BRANCH'"
    fi
}

# 显示完成信息
show_completion() {
    echo ""
    echo "╔════════════════════════════════════════════════════════════╗"
    echo "║                                                            ║"
    echo "║              🎉 部署完成! 文档转换工具已上线 🩷            ║"
    echo "║                                                            ║"
    echo "╚════════════════════════════════════════════════════════════╝"
    echo ""
    echo "📋 总结:"
    echo "  - 版本: v2.0.1 (Bug 修复版)"
    echo "  - 仓库: $REPO_URL"
    echo "  - 网站: https://muqiuri3721.github.io/doc-converter/"
    echo ""
    echo "✨ 新功能:"
    echo "  - 修复 HTML 转义问题"
    echo "  - 优化 PDF 文本提取"
    echo "  - 增强错误处理"
    echo ""
    print_heart "Made with love by 02"
    echo ""
}

# 主函数
main() {
    show_welcome
    check_requirements
    check_git_config
    get_github_token
    prepare_repository
    backup_old_version
    copy_new_files
    commit_changes
    push_to_github
    verify_deployment
    show_completion
}

# 错误处理
trap 'print_error "脚本执行失败，请检查错误信息"; exit 1' ERR

# 运行主函数
main
