@echo off
chcp 65001 >nul
title Excel合并工具 - Docker开发环境

echo 🚀 Excel合并工具 - Docker开发环境
echo ============================

docker-compose -f docker-compose.dev.yml up --build

echo.
echo 按任意键退出...
pause >nul