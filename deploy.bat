@echo off
chcp 65001 >nul
title 文档转换工具部署脚本

echo.
echo ╔════════════════════════════════════════════════════════════╗
echo ║                                                            ║
echo ║     📄 文档转换工具 Pro v2.0 - Windows 部署脚本           ║
echo ║                                                            ║
echo ╚════════════════════════════════════════════════════════════╝
echo.

:: 检查 Git
where git >nul 2>nul
if %errorlevel% neq 0 (
    echo [❌] 错误: 未找到 Git
    echo.
    echo 请先安装 Git: https://git-scm.com/download/win
    pause
    exit /b 1
)

echo [✓] Git 已安装

:: 输入 Token
echo.
echo 💡 需要 GitHub Personal Access Token
set /p TOKEN="请输入你的 GitHub Token: "

if "%TOKEN%"=="" (
    echo [❌] Token 不能为空
    pause
    exit /b 1
)

set REPO=https://%TOKEN%@github.com/MuQiuri3721/doc-converter.git

echo.
echo [🚀] 开始部署...

:: 克隆仓库
echo [📥] 克隆仓库...
if exist "doc-converter" rmdir /s /q "doc-converter"
git clone "%REPO%" doc-converter 2>nul
if %errorlevel% neq 0 (
    echo [❌] 克隆失败，请检查 Token
    pause
    exit /b 1
)

:: 复制文件
echo [📋] 复制文件...
copy /y index.html doc-converter\
copy /y converter.js doc-converter\

:: 提交并推送
echo [📤] 提交更改...
cd doc-converter
git add .
git commit -m "🎉 Update to v2.0.1"
git push origin main

if %errorlevel% neq 0 (
    echo [❌] 推送失败
    pause
    exit /b 1
)

echo.
echo ╔════════════════════════════════════════════════════════════╗
echo ║                    🎉 部署成功! 🩷                         ║
echo ╚════════════════════════════════════════════════════════════╝
echo.
echo 🌐 访问地址: https://muqiuri3721.github.io/doc-converter/
echo.
echo 💡 提示: 首次部署后需要 1-2 分钟生效
echo.
pause
