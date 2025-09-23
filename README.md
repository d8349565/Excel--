# Excel/CSV 汇总工具 v1.2.0

一个功能强大的基于 Flask 的 Web 应用程序，专为 Excel/CSV 文件处理而设计。提供完整的文件上传、预览、管理和数据处理解决方案，支持批量数据合并、单元格提取、列排序等高级功能。通过直观的 Web 界面，用户可以轻松处理大量 Excel/CSV 数据，无需复杂的编程知识。

![1758644736074](images/README/1758644736074.png)

## 主要特性

- **🔐 安全认证**: 双层密码保护系统，支持普通用户和管理员权限
- **📁 智能文件管理**: 支持 Excel (.xlsx/.xls) 和 CSV 文件上传，自动文件验证和清理
- **👀 实时预览**: 文件内容实时预览，支持分页显示和数据筛选
- **🔄 批量数据合并**: 多文件智能合并，支持自定义合并规则和冲突处理
- **📊 异步任务处理**: 后台任务队列，确保大文件处理不阻塞界面
- **📝 完整日志记录**: 用户操作详细日志，支持管理员审计
- **🗂️ 管理员面板**: 专属管理界面，查看系统状态和用户活动
- **📋 单元格提取**: 灵活的单元格数据提取功能，支持自定义规则
- **🎯 列管理优化**: 拖拽排序、列筛选、数据类型识别
- **🐳 容器化部署**: Docker 一键部署，支持生产环境快速上线
- **📦 便携打包**: PyInstaller 生成独立可执行文件，无需安装 Python 环境

## 技术栈

- **后端**: Flask (Python)
- **前端**: HTML/CSS/JavaScript, Bootstrap
- **数据处理**: Pandas, OpenPyXL
- **任务管理**: 自定义异步任务系统
- **容器化**: Docker & Docker Compose

## 快速开始

### 本地运行

1. **克隆项目**

   ```bash
   git clone <repository-url>
   cd Excel--
   ```
2. **安装依赖**

   ```bash
   cd excel_tool
   pip install -r requirements.txt
   ```
3. **运行应用**

   ```bash
   python app.py
   ```

   或使用启动脚本：

   ```bash
   start.bat
   ```
4. **访问应用**

   打开浏览器访问 `http://localhost:5000`

### Docker 部署

1. **构建并运行**

   ```bash
   cd docker
   docker-compose up -d
   ```
2. **访问应用**

   打开浏览器访问 `http://localhost:9999`

## 配置说明

### 环境变量

- `SECRET_KEY`: Flask 应用密钥
- `ACCESS_PASSWORD`: 用户访问密码（默认: 123456）
- `ADMIN_PASSWORD`: 管理员密码（默认: admin2025）

### 文件限制

- 单个文件最大: 100MB
- 总上传限制: 500MB
- 支持格式: .xlsx, .xls, .csv

## 使用指南

1. **登录**: 使用访问密码登录系统
2. **上传文件**: 选择并上传 Excel/CSV 文件
3. **预览文件**: 点击文件查看内容预览
4. **合并数据**: 选择多个文件进行数据合并
5. **下载结果**: 下载处理后的合并文件
6. **管理员功能**: 使用管理员密码查看用户操作日志

## 项目结构

```
Excel--
├── excel_tool/              # 主应用目录
│   ├── app.py              # Flask 应用主文件
│   ├── config.py           # 配置文件
│   ├── data_processor.py   # 数据处理模块
│   ├── file_manager.py     # 文件管理模块
│   ├── task_manager.py     # 任务管理模块
│   ├── user_logger.py      # 用户日志模块
│   ├── templates/          # HTML 模板
│   ├── static/             # 静态文件 (CSS/JS)
│   ├── uploads/            # 上传文件目录
│   ├── results/            # 处理结果目录
│   ├── logs/               # 日志文件目录
│   └── requirements.txt    # Python 依赖
├── docker/                 # Docker 部署配置
│   ├── Dockerfile
│   ├── docker-compose.yml
│   └── README.md
├── exe打包命令             # PyInstaller 打包命令
└── README.md              # 项目说明
```

## 开发说明

### 打包为可执行文件

使用 PyInstaller 打包：

```bash
cd excel_tool
pyinstaller --onefile --add-data "templates;templates" --add-data "static;static" app.py
```

### 任务管理

应用使用自定义的任务管理系统支持异步处理：

- 最大并发任务数: 5
- 任务超时时间: 1小时
- 支持任务状态跟踪和结果获取

## 安全特性

- 用户会话管理
- 文件上传验证
- 密码保护访问
- 管理员权限控制
- 自动文件清理（1天保留期）

## 许可证

本项目允许个人使用和商业使用，但需要注明源代码作者和项目地址。

**作者**: d8349565
**项目地址**: https://github.com/d8349565/Excel--

## 赞赏支持

如果这个项目对你有帮助，欢迎请作者喝一杯咖啡！☕

### 微信赞赏

<img src="images/README/1758644628077.png" alt="1758644628077" style="zoom:33%;" />

## 贡献

欢迎提交 Issue 和 Pull Request 来改进项目。

## 更新日志

### v1.2.0 (2025-09-24)

- ✨ 新增单元格提取功能
- 🔧 优化配置管理
- 🎨 优化系统管理员界面
- 📋 优化列管理逻辑，新增拖拽排序
- 🐛 修复部分bug

### v1.0.0

- 初始版本发布
- 支持文件上传、预览、合并
- 用户认证和管理员功能
- Docker 容器化支持
