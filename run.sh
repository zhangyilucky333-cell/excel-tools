#!/bin/bash

# 安装依赖
echo "安装Python依赖..."
pip3 install -r requirements.txt

# 创建必要的目录
mkdir -p uploads outputs

# 启动应用
echo "启动Excel拆分/合并工具..."
python3 app.py
