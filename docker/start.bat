@echo off
chcp 65001 >nul
title Excel合并工具 - Docker部署

echo 🐳 Excel合并工具 - Docker部署
echo ================================

:: 检查Docker是否安装
docker --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Docker未安装，请先安装Docker Desktop
    echo 📥 下载地址：https://www.docker.com/products/docker-desktop
    pause
    exit /b 1
)

:: 检查docker-compose是否安装
docker-compose --version >nul 2>&1
if errorlevel 1 (
    echo ❌ docker-compose未安装，请先安装docker-compose
    pause
    exit /b 1
)

echo ✅ Docker环境检查通过

:: 构建并启动服务
echo 🔨 构建Docker镜像...
docker-compose build

echo 🚀 启动服务...
docker-compose up -d

:: 等待服务启动
echo ⏳ 等待服务启动...
timeout /t 10 /nobreak >nul

:: 检查服务状态
docker-compose ps | findstr "Up" >nul
if errorlevel 1 (
    echo ❌ 服务启动失败，请查看日志：
    docker-compose logs
    pause
    exit /b 1
)

echo ✅ 服务启动成功！
echo.
echo 🌐 访问地址：
echo    http://localhost:5000
echo.
echo 📊 管理命令：
echo    查看日志: docker-compose logs -f
echo    停止服务: docker-compose down
echo    重启服务: docker-compose restart
echo.
echo 按任意键打开浏览器...
pause >nul
start http://localhost:5000