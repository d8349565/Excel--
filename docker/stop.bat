@echo off
chcp 65001 >nul
title 停止Excel合并工具

echo 🛑 停止Excel合并工具服务
echo ========================

docker-compose down

echo ✅ 服务已停止
echo 📁 数据已保存在Docker卷中，下次启动时会自动恢复
echo.
pause