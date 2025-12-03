# Excel拆分合并工具

一个基于Flask的Excel文件拆分和合并工具，提供Web界面操作。

## 功能特性

### 1. Excel拆分
- 按指定列拆分Excel文件
- 支持多sheet拆分
- 保留原始格式
- 自动打包下载

### 2. Excel合并
- 合并多个Excel文件
- 相同sheet自动合并
- 智能去重标题行
- 列自动对齐

## 快速开始

### 本地运行

1. 安装依赖：
```bash
pip install -r requirements.txt
```

2. 启动服务：
```bash
python app.py
```

3. 访问地址：
- 拆分功能: http://localhost:5001
- 合并功能: http://localhost:5001/merger

### 云端部署

本项目支持一键部署到以下平台：

#### Render.com (推荐)
[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com)

1. 注册 [Render](https://render.com) 账号
2. 创建新的 Web Service
3. 连接此GitHub仓库
4. Render会自动检测并部署

#### Railway
[![Deploy on Railway](https://railway.app/button.svg)](https://railway.app)

1. 注册 [Railway](https://railway.app) 账号
2. 点击上方按钮一键部署
3. 等待部署完成

## 技术栈

- Python 3.11
- Flask 3.0
- pandas 2.1
- openpyxl 3.1

## 配置文件说明

- `Procfile`: 部署启动命令
- `runtime.txt`: Python版本
- `requirements.txt`: 依赖包列表

## 使用限制

- 最大文件大小: 50MB
- 支持格式: .xlsx, .xls
