#!/bin/bash
# Docker 部署启动脚本

echo "🐳 Excel合并工具 - Docker部署"
echo "================================"

# 检查Docker是否安装
if ! command -v docker &> /dev/null; then
    echo "❌ Docker未安装，请先安装Docker"
    echo "📥 下载地址：https://www.docker.com/products/docker-desktop"
    exit 1
fi

# 检查docker-compose是否安装
if ! command -v docker-compose &> /dev/null; then
    echo "❌ docker-compose未安装，请先安装docker-compose"
    exit 1
fi

echo "✅ Docker环境检查通过"

# 构建并启动服务
echo "🔨 构建Docker镜像..."
docker-compose build

echo "🚀 启动服务..."
docker-compose up -d

# 等待服务启动
echo "⏳ 等待服务启动..."
sleep 10

# 检查服务状态
if docker-compose ps | grep -q "Up"; then
    echo "✅ 服务启动成功！"
    echo ""
    echo "🌐 访问地址："
    echo "   http://localhost:5000"
    echo ""
    echo "📊 管理命令："
    echo "   查看日志: docker-compose logs -f"
    echo "   停止服务: docker-compose down"
    echo "   重启服务: docker-compose restart"
else
    echo "❌ 服务启动失败，请查看日志："
    docker-compose logs
fi