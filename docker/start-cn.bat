@echo off
chcp 65001 >nul
title Excel合并工具 - Docker部署（中国网络优化）

echo 🐳 Excel合并工具 - Docker部署（中国网络优化）
echo ========================================

:: 检查Docker是否安装
docker --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Docker未安装，请先安装Docker Desktop
    echo 📥 下载地址：https://www.docker.com/products/docker-desktop
    pause
    exit /b 1
)

echo ✅ Docker环境检查通过
echo 🚀 使用中国镜像源加速构建...

:: 使用中国优化版配置构建
echo 🔨 构建Docker镜像（使用国内镜像源）...
docker-compose -f docker-compose.cn.yml build

echo 🚀 启动服务...
docker-compose -f docker-compose.cn.yml up -d

:: 等待服务启动
echo ⏳ 等待服务启动...
timeout /t 10 /nobreak >nul

:: 检查服务状态
docker-compose -f docker-compose.cn.yml ps | findstr "Up" >nul
if errorlevel 1 (
    echo ❌ 服务启动失败，请查看日志：
    docker-compose -f docker-compose.cn.yml logs
    pause
    exit /b 1
)

echo ✅ 服务启动成功！
echo.
echo 🌐 访问地址：
echo    http://localhost:8899
echo.
echo 📊 管理命令：
echo    查看日志: docker-compose -f docker-compose.cn.yml logs -f
echo    停止服务: docker-compose -f docker-compose.cn.yml down
echo    重启服务: docker-compose -f docker-compose.cn.yml restart
echo.
echo 按任意键打开浏览器...
pause >nul
start http://localhost:8899