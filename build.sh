#!/usr/bin/env bash
# macOS 打包脚本：生成 .app 应用程序（需在 Mac 上执行）
set -euo pipefail

cd "$(dirname "$0")"

echo "========================================"
echo "  客户产品 Excel 汇总工具 - macOS 打包"
echo "========================================"
echo

if [[ "$(uname)" != "Darwin" ]]; then
  echo "错误：此脚本只能在 macOS 上运行。"
  echo "Windows 请使用 build.bat。"
  exit 1
fi

PYTHON="${PYTHON:-python3}"

echo "[1/3] 安装打包依赖..."
"$PYTHON" -m pip install -r requirements-build.txt -q

echo "[2/3] 开始打包（约 1–3 分钟）..."
"$PYTHON" -m PyInstaller build-mac.spec --noconfirm --clean

echo
echo "[3/3] 打包完成！"
echo
echo "应用程序位置："
echo "  dist/客户产品Excel汇总工具.app"
echo
echo "发给同事时，可将 .app 压缩为 zip 发送。"
echo "首次打开若被拦截，请在「系统设置 → 隐私与安全性」中允许运行。"
echo
