#!/bin/bash
# ================================================
#  Academic PPT Skill - 安裝腳本
# ================================================
set -e

# 嘗試找出專案目錄（支援 curl|bash 與本地執行兩種模式）
if [ -f "$(dirname "$0")/SKILL.md" ]; then
    SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
elif [ -f "./SKILL.md" ]; then
    SCRIPT_DIR="$(pwd)"
else
    echo "ERROR: 無法找到 SKILL.md，請確認在專案目錄下執行，或手動指定路徑"
    exit 1
fi

TARGET_DIR="$HOME/.hermes/skills/productivity/academic-ppt"

echo "=========================================="
echo "  學術簡報生成器 - 安裝程式"
echo "=========================================="

FORCE_OVERWRITE=false
if [ "$1" = "-y" ] || [ "$1" = "--force" ]; then
    FORCE_OVERWRITE=true
fi

# 檢查目錄
if [ -d "$TARGET_DIR" ]; then
    if [ "$FORCE_OVERWRITE" = false ]; then
        echo "⚠️  偵測到已安裝，是否覆寫？ (y/N)"
        read -r confirm
        if [ "$confirm" != "y" ] && [ "$confirm" != "Y" ]; then
            echo "安裝中止"
            exit 0
        fi
    fi
    rm -rf "$TARGET_DIR"
fi

# 建立目錄並複製
echo "📁 建立目錄結構..."
mkdir -p "$TARGET_DIR/scripts"

# 複製檔案
cp "$SCRIPT_DIR/SKILL.md" "$TARGET_DIR/"
cp "$SCRIPT_DIR/generate_ppt.py" "$TARGET_DIR/scripts/"

# 檢查 python-pptx
echo "🔍 檢查依賴..."
if python3 -c "import pptx" 2>/dev/null; then
    echo "✓ python-pptx 已安裝"
else
    echo "📦 安裝 python-pptx..."
    pip3 install python-pptx
fi

echo ""
echo "=========================================="
echo "  安裝完成！"
echo "=========================================="
echo ""
echo "執行方式："
echo "  python3 ~/.hermes/skills/productivity/academic-ppt/scripts/generate_ppt.py"
echo ""
echo "詳細說明請參考 SKILL.md"
