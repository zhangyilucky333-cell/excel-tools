#!/bin/bash

# Excel 拆分/合并工具启动脚本 (Linux)

# 获取脚本所在目录
cd "$(dirname "$0")"

echo "================================================"
echo "  Excel 拆分/合并工具"
echo "================================================"
echo ""

# 检查 Python 是否安装
if ! command -v python3 &> /dev/null; then
    echo "❌ 错误: 未检测到 Python 3"
    echo "请先安装 Python 3.8 或更高版本"
    echo "Ubuntu/Debian: sudo apt install python3 python3-pip"
    echo "CentOS/RHEL: sudo yum install python3 python3-pip"
    echo ""
    read -p "按回车键退出..."
    exit 1
fi

echo "✓ Python 版本: $(python3 --version)"
echo ""

# 检查依赖是否安装
echo "正在检查依赖..."
if ! python3 -c "import pandas" 2>/dev/null; then
    echo "⚠ 首次运行，正在安装依赖包..."
    echo "这可能需要几分钟时间，请耐心等待..."
    echo ""
    pip3 install -r requirements.txt
    if [ $? -ne 0 ]; then
        echo ""
        echo "❌ 依赖安装失败，尝试使用国内镜像源..."
        pip3 install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
    fi
    echo ""
fi

echo "✓ 依赖检查完成"
echo ""

# 启动服务
echo "================================================"
echo "  正在启动服务..."
echo "================================================"
echo ""

# 在后台启动服务
python3 app.py &
APP_PID=$!

# 等待服务启动
sleep 3

# 尝试打开浏览器
echo "正在打开浏览器..."
if command -v xdg-open &> /dev/null; then
    xdg-open "http://127.0.0.1:5001" &
elif command -v gnome-open &> /dev/null; then
    gnome-open "http://127.0.0.1:5001" &
else
    echo "无法自动打开浏览器，请手动访问: http://127.0.0.1:5001"
fi

echo ""
echo "================================================"
echo "  服务已启动！"
echo "================================================"
echo ""
echo "拆分功能: http://127.0.0.1:5001"
echo "合并功能: http://127.0.0.1:5001/merger"
echo ""
echo "按 Ctrl+C 停止服务"
echo ""

# 等待进程
wait $APP_PID
