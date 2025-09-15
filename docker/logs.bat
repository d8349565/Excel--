@echo off
chcp 65001 >nul
title Excel合并工具 - 查看日志

echo 📊 Excel合并工具 - 实时日志
echo ========================
echo 按 Ctrl+C 退出日志查看
echo.

docker-compose logs -f