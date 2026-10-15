# pptx-deck-builder

用开源工具链**程序化生成原生 PowerPoint**:幻灯片先写成 SVG(可 diff、可评审),
再由编译器物化成 PowerPoint **原生 DrawingML 形状与文本框** —— 不是图片,也不是
填好的模板,双击任意元素都能接着编辑。

- 成品(8 页,16:9):[`output/svg-to-native-pptx-deck.pptx`](output/svg-to-native-pptx-deck.pptx)
- 每页预览:[`output/preview/`](output/preview)
- 幻灯片源(SVG):[`slides/`](slides) · 生成器:[`builder/make_slides.py`](builder/make_slides.py)

## 用到的开源项目

| 项目 | 说明 |
|---|---|
| [hugohe3/ppt-master](https://github.com/hugohe3/ppt-master) | 62,295 ★ · MIT · Python。工作流 + 质检 + **SVG→DrawingML 编译器** |
| [scanny/python-pptx](https://github.com/scanny/python-pptx) | 3,556 ★ · MIT · Python。PPTX 读写底座(OOXML 组装、母版/版式绑定) |

star 数截至 2026-10-15。

## 流水线(六步)

```bash
# 0) 取工具链
git clone --depth 1 https://github.com/hugohe3/ppt-master.git /tmp/ppt-master
SKILL=/tmp/ppt-master/skills/ppt-master
pip install python-pptx XlsxWriter skia-pathops uharfbuzz PyYAML

# 1) 生成幻灯片源(SVG 设计系统 + 8 页)
python3 builder/make_slides.py

# 2) 初始化 ppt-master 项目
python3 $SKILL/scripts/project_manager.py init opensource-pptx-pipeline --quick-generate
PROJ=$(ls -dt /tmp/ppt-master/projects/opensource-pptx-pipeline_* | head -1)
cp slides/*.svg $PROJ/svg_output/

# 3) 质量门禁:校验 SVG 契约(8/8 fully passed)
python3 $SKILL/scripts/svg_quality_checker.py $PROJ --quick-generate \
  --canonical-authoring --stage final --json

# 4) 编译成原生 PPTX
python3 $SKILL/scripts/svg_to_pptx.py $PROJ --quick-generate \
  -o output/svg-to-native-pptx-deck.pptx

# 5) 预览图 + 6) 交付
python3 builder/render_preview.py
```

一条命令跑完:`bash builder/build_pptx.sh`(需要 bash / git / python3 / pip)。

## 产物事实

| 项 | 值 |
|---|---|
| 页数 / 画布 | 8 页 · 13.33 × 7.5 inch(1280 × 720 的 16:9 画布) |
| 形状构成 | AUTO_SHAPE + TEXT_BOX + FREEFORM,**全部为原生对象** |
| 图片 / SmartArt | 0 张 / 0 个 |
| 质检结果 | 8/8 页 fully passed,0 error(`output/reports/`) |
| 外部密钥 | 0(无模型、无图片 API) |

## 诚实说明

- 预览图是从 `slides/*.svg` 渲染的(与编译进 PPTX 的同一份几何),不是 PPTX 截图;
  想把 PPTX 逐页截图需要额外装 LibreOffice。
- 版式里指定了 `Microsoft YaHei / PingFang SC / Noto Sans SC` 字体族,请在装有中文字体的
  机器上打开;字体缺失时 PowerPoint 会回退,行宽可能与预览略有出入。
- `builder/make_slides.py` 里的文字宽度估算只用于声明每组的 `data-pptx-bounds`,
  不参与定位。
- 本仓库的幻灯片内容(star 数、命令、对比结论)由 AI 在 GitHub Codespace 中生成,
  事实性数字取自 GitHub API,时间为 2026-10-15。

## 许可

本仓库代码与幻灯片采用 MIT。`ppt-master` 由 Hugo He 开发,同为 MIT 许可,未内联其源码。
