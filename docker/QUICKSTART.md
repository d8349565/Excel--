# 🐳 Docker 快速指南

## 📦 一键启动（推荐）

### Windows用户
```cmd
# 生产环境
start.bat

# 开发环境（代码实时更新）
dev.bat
```

### Linux/Mac用户
```bash
# 生产环境
./start.sh

# 开发环境
docker-compose -f docker-compose.dev.yml up --build
```

## 🎯 访问地址
http://localhost:5000

## 🛠 管理命令

| 操作 | Windows | Linux/Mac |
|------|---------|-----------|
| 启动服务 | `start.bat` | `./start.sh` |
| 停止服务 | `stop.bat` | `docker-compose down` |
| 查看日志 | `logs.bat` | `docker-compose logs -f` |
| 开发模式 | `dev.bat` | `docker-compose -f docker-compose.dev.yml up` |

## 📊 特性对比

| 特性 | 本地部署 | Docker部署 |
|------|----------|------------|
| 环境依赖 | 需要Python | 只需Docker |
| 启动速度 | 快 | 首次较慢 |
| 环境隔离 | 无 | 完全隔离 |
| 数据持久化 | 本地文件 | Docker卷 |
| 跨平台 | 需配置 | 开箱即用 |

## ⚡ 快速排错

1. **端口占用**: 修改`docker-compose.yml`中的端口映射
2. **权限问题**: Linux用户加入docker组 `sudo usermod -aG docker $USER`
3. **构建失败**: 运行 `docker-compose build --no-cache`
4. **数据丢失**: 检查Docker卷 `docker volume ls`