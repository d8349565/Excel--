# Docker 部署说明 🐳

## 快速开始

### 1️⃣ 安装Docker
- **Windows/Mac**: 下载安装 [Docker Desktop](https://www.docker.com/products/docker-desktop)
- **Linux**: 安装Docker和docker-compose
  ```bash
  # Ubuntu/Debian
  sudo apt-get update
  sudo apt-get install docker.io docker-compose
  
  # CentOS/RHEL
  sudo yum install docker docker-compose
  ```

### 2️⃣ 启动服务

**Windows用户：**
```cmd
# 双击运行
start.bat
```

**Linux/Mac用户：**
```bash
# 给脚本执行权限
chmod +x start.sh

# 运行启动脚本
./start.sh

# 或手动启动
docker-compose up -d
```

### 3️⃣ 访问应用
浏览器打开：http://localhost:5000

## 管理命令

### 查看服务状态
```bash
docker-compose ps
```

### 查看日志
```bash
# Windows
logs.bat

# Linux/Mac
docker-compose logs -f
```

### 停止服务
```bash
# Windows
stop.bat

# Linux/Mac
docker-compose down
```

### 重启服务
```bash
docker-compose restart
```

### 更新应用
```bash
# 重新构建镜像
docker-compose build

# 重启服务
docker-compose up -d
```

## 数据持久化

Docker部署会自动创建三个数据卷：
- `excel_uploads` - 上传文件
- `excel_results` - 处理结果
- `excel_logs` - 应用日志

数据在容器重启后会自动保留。

## 配置说明

### 端口配置
默认端口：5000
如需修改，编辑 `docker-compose.yml`：
```yaml
ports:
  - "8080:5000"  # 改为8080端口
```

### 资源限制
如需限制内存使用，在 `docker-compose.yml` 中添加：
```yaml
services:
  excel-tool:
    # ... 其他配置
    deploy:
      resources:
        limits:
          memory: 1G
        reservations:
          memory: 512M
```

## 故障排除

### 1. 端口被占用
```bash
# 查看端口占用
netstat -ano | findstr :5000

# 修改docker-compose.yml中的端口映射
ports:
  - "5001:5000"
```

### 2. 构建失败
```bash
# 清理缓存重新构建
docker-compose build --no-cache
```

### 3. 权限问题（Linux）
```bash
# 添加用户到docker组
sudo usermod -aG docker $USER

# 重新登录后生效
```

### 4. 数据丢失
数据存储在Docker卷中，可以通过以下命令查看：
```bash
# 查看所有卷
docker volume ls

# 查看卷详情
docker volume inspect docker_excel_uploads
```

## 卸载

### 停止并删除容器
```bash
docker-compose down
```

### 删除镜像
```bash
docker rmi excel-merge-tool
```

### 删除数据卷（注意：会丢失所有数据）
```bash
docker volume rm docker_excel_uploads docker_excel_results docker_excel_logs
```

## 优势

✅ **一键部署** - 无需配置Python环境  
✅ **环境隔离** - 不影响系统其他应用  
✅ **数据持久化** - 重启后数据不丢失  
✅ **易于管理** - 简单的启停操作  
✅ **跨平台** - Windows/Linux/Mac 统一体验