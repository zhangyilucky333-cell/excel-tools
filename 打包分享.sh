#!/bin/bash

# Excel 拆分/合并工具 - 打包分享脚本

cd "$(dirname "$0")"

PACKAGE_NAME="Excel拆分合并工具"
PACKAGE_DIR="${PACKAGE_NAME}"
ZIP_NAME="${PACKAGE_NAME}_$(date +%Y%m%d).zip"

echo "================================================"
echo "  Excel 拆分/合并工具 - 打包分享"
echo "================================================"
echo ""

# 创建临时打包目录
echo "📦 正在准备打包..."
rm -rf "${PACKAGE_DIR}"
mkdir -p "${PACKAGE_DIR}"

# 复制必要文件
echo "📋 复制文件..."
cp app.py "${PACKAGE_DIR}/"
cp excel_splitter.py "${PACKAGE_DIR}/"
cp excel_merger.py "${PACKAGE_DIR}/"
cp requirements.txt "${PACKAGE_DIR}/"
cp 使用说明.md "${PACKAGE_DIR}/"
cp 启动工具.command "${PACKAGE_DIR}/"
cp 启动工具.bat "${PACKAGE_DIR}/"
cp 启动工具.sh "${PACKAGE_DIR}/"
cp .gitignore "${PACKAGE_DIR}/"

# 复制模板文件夹
cp -r templates "${PACKAGE_DIR}/"

# 设置脚本执行权限
chmod +x "${PACKAGE_DIR}/启动工具.command"
chmod +x "${PACKAGE_DIR}/启动工具.sh"

# 创建空的文件夹（让用户知道这些会自动创建）
mkdir -p "${PACKAGE_DIR}/uploads"
mkdir -p "${PACKAGE_DIR}/outputs"

# 创建自述文件
cat > "${PACKAGE_DIR}/开始使用.txt" << 'EOF'
============================================
  Excel 拆分/合并工具 - 快速开始
============================================

📌 使用方法：

Windows 用户：
  双击 "启动工具.bat" 

Mac 用户：
  双击 "启动工具.command"
  (首次使用可能需要右键 -> 打开)

Linux 用户：
  运行 "启动工具.sh"

📖 详细说明：
  请查看 "使用说明.md" 文件

⚠️ 注意事项：
  1. 需要先安装 Python 3.8 或更高版本
  2. 首次运行会自动安装依赖包
  3. 所有数据处理都在本地进行，安全可靠

📞 遇到问题？
  请查看"使用说明.md"中的常见问题部分

============================================
版本: 1.0.0
更新日期: 2025-12-03
============================================
EOF

echo "✓ 文件复制完成"
echo ""

# 创建 ZIP 包
echo "🗜️  正在压缩..."
if command -v zip &> /dev/null; then
    rm -f "${ZIP_NAME}"
    zip -r "${ZIP_NAME}" "${PACKAGE_DIR}" -x "*.DS_Store" "*.pyc" "__pycache__/*"
    echo "✓ 压缩完成"
else
    echo "⚠ 未找到 zip 命令，跳过压缩"
    echo "打包目录: ${PACKAGE_DIR}"
fi

echo ""
echo "================================================"
echo "  ✅ 打包完成！"
echo "================================================"
echo ""
echo "📦 打包文件: ${ZIP_NAME}"
echo "📁 打包目录: ${PACKAGE_DIR}"
echo ""
echo "🎁 分享方式："
echo "   1. 将 ${ZIP_NAME} 发送给其他人"
echo "   2. 或直接分享 ${PACKAGE_DIR} 文件夹"
echo ""
echo "💡 接收者使用方法："
echo "   1. 解压 ZIP 文件（如果是压缩包）"
echo "   2. 根据操作系统双击相应的启动脚本"
echo "   3. 首次运行会自动安装依赖"
echo ""
