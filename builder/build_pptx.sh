#!/usr/bin/env bash
# ─────────────────────────────────────────────────────────────────────────────
# build_pptx.sh — 端到端流水线:slides/*.svg → output/*.pptx(原生可编辑)
#
#   1) 克隆 ppt-master(https://github.com/hugohe3/ppt-master,62,295★,MIT)
#   2) 安装依赖
#   3) 生成幻灯片源(SVG)
#   4) 初始化 ppt-master 项目并导入 SVG
#   5) 跑质量门禁(svg_quality_checker)
#   6) 编译成原生 PPTX(svg_to_pptx)
#   7) 渲染预览图(可选)
#
# 依赖:bash、git、python3(3.9+)、pip
# 用法:bash builder/build_pptx.sh
# ─────────────────────────────────────────────────────────────────────────────
set -euo pipefail

REPO_SLUG="hugohe3/ppt-master"
PROJECT_NAME="opensource-pptx-pipeline"
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
WORK="${PPT_MASTER_HOME:-${TMPDIR:-/tmp}/ppt-master}"
SKILL="$WORK/skills/ppt-master"
OUT_PPTX="$ROOT/output/svg-to-native-pptx-deck.pptx"

cd "$ROOT"

echo "==> 1/7 取工具链:$REPO_SLUG"
if [ ! -d "$SKILL" ]; then
  git clone --depth 1 "https://github.com/$REPO_SLUG.git" "$WORK"
fi
echo "    ppt-master 就位:$SKILL"

echo "==> 2/7 安装依赖"
pip install -q "python-pptx>=0.6.21" XlsxWriter skia-pathops uharfbuzz PyYAML

echo "==> 3/7 生成幻灯片源"
python3 builder/make_slides.py

echo "==> 4/7 初始化 ppt-master 项目"
rm -rf "$WORK"/projects/${PROJECT_NAME}_*
python3 "$SKILL/scripts/project_manager.py" init "$PROJECT_NAME" --quick-generate >/dev/null
PROJ="$(ls -dt "$WORK"/projects/${PROJECT_NAME}_* | head -1)"
rm -f "$PROJ"/svg_output/*.svg
cp slides/*.svg "$PROJ"/svg_output/
echo "    项目:$PROJ"

echo "==> 5/7 质量门禁(SVG 契约)"
python3 "$SKILL/scripts/svg_quality_checker.py" "$PROJ" \
  --quick-generate --canonical-authoring --stage final --json | grep -E "Fully passed|With errors|blocking" || true

echo "==> 6/7 编译:SVG → 原生 DrawingML PPTX"
mkdir -p "$(dirname "$OUT_PPTX")"
python3 "$SKILL/scripts/svg_to_pptx.py" "$PROJ" --quick-generate -o "$OUT_PPTX" | tail -6

echo "==> 7/7 预览图"
python3 builder/render_preview.py || echo "    (跳过预览)"

# 把质检与产物报告留档,方便复查
mkdir -p "$ROOT/output/reports"
cp -f "$PROJ"/validation/svg_quality_report.json "$ROOT/output/reports/" 2>/dev/null || true
cp -f "$PROJ"/validation/*report.json "$ROOT/output/reports/" 2>/dev/null || true

echo
echo "✅ 完成:$OUT_PPTX"
ls -lh "$OUT_PPTX" | awk '{print "   大小:", $5}'
