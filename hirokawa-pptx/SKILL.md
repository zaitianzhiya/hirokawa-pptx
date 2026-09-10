---
name: hirokawa-pptx
description: Create any corporate presentation in the HIROKAWA (上海广川科技) company template style — business trip reports (出差报告), R&D project progress tracking (研发项目进度), weekly/monthly reports (周报/月报), technical proposals, meeting minutes, client visit reports, and other company-standard decks. Use whenever the user asks to create a company PPT/presentation with the branded navy template, cover page, revision history, table of contents, content slides (text/table/form), and closing slide. Output is PPTX with built-in format and content validation.
---

# 公司通用 PPT 技能（HIROKAWA 企业模板）

适用于公司内部**所有**标准 PPT：出差报告、项目进度、周报月报、技术方案、会议纪要、客户拜访报告等。
版式统一为：封面 → 改版履历 → 目录 → 内容页×N → 封底。

## 快速上手（固定三步，最省 token）

```bash
# 1. 把内容写成 data.json（schema 见 scripts/build_deck.py 顶部注释，
#    示例见 references/examples/）
# 2. 生成
python3 scripts/build_deck.py data.json -o 报告.pptx
# 3. 成品检查（必做）：自动校验 + 渲染每页 PNG 到 报告_check/
python3 scripts/check_deck.py 报告.pptx
```

检查器退出码非 0（存在 FAIL）时，修正后重新生成并复检；
全部 PASS 后，逐页查看渲染 PNG 目视复核一次再交付。

**不要重新绘制版式，也不要手写 XML**——模板 `assets/template.pptx` 已 100% 还原公司规范。

## 内容页六种类型

| type | 适用场景 | 关键字段 |
|---|---|---|
| `text` | 正文叙述、章节展开 | `blocks`: [["h","小节标题"],["p","段落"],["b","◆要点"]] |
| `table` | 数据表、进度表、对比表 | `columns` + `rows`（自动套用藏青表头/蓝色正文），可选 `widths`、`blocks` 附注 |
| `form` | 概要信息、键值表单（出差概要等） | `fields`（每行 标签/值/标签/值，2 元行自动合并）+ `sections` 小节 |
| `image_text` | 图文页（设备介绍、方案说明） | `image` 图片路径，`image_side` left/right，`caption` 图注，`blocks` 文字 |
| `gallery` | 多图页（现场照片、截图集） | `images`: [{"path","caption"}] 1-4 张自动网格 + 藏青边框 |
| `fixed_image` | 固定版式图文页（对外正式报告） | `image` + `blocks`；图框位置尺寸由模板锁定 |

图片处理：所有插图自动按目标区域**中心裁剪适配**（不变形）；正式报告建议用实拍图。

## 无图片素材时：AI 生图工作流

JSON 中给 `image_prompt`（英文描述）代替 `image` 路径时：
1. 先用 image_generation 插件生成图片到本地（参考图需先 image-to-url 转公网 URL）
2. 把生成的本地路径填入 `image` 字段，再正常构建
3. 正式对外报告应将 AI 图替换为实拍照片（JSON 改路径重新生成即可）

报告类型只是 slides 组合不同。`references/examples/` 内置示例：
- `出差报告.json`：form(出差概要) + text(详细内容)
- `项目进度报告.json`：form + table + text
- `产品介绍.json`：image_text + gallery（图文演示）
- `通用汇报.json`：最简骨架

## 公司规范要点（模板已内置，无需手动处理）

- 封面标题字面 "**　" 前缀；文书编号；批准栏 APPR./CHECK./CHECK./DRAW.（不带括号）
- 改版履历 5 列：版本/改版时间/改版理由/作成/承认
- 表格：表头 #28166F 白字、正文统一 #5B9BD5；字体：封面微软雅黑、标题黑体、表格正文仿宋、西文 Times New Roman、目录 Arial
- 全部页：CONFIDENTIAL 徽标、页脚公司名+自动页码、左下版本号
- 装饰图形红屋顶朝上，不可倒置

## 成品检查项（check_deck.py 自动覆盖）

- 内容：占位符残留（（标题）/（正文）/\*\*\*/2024-\*-\* 等）、封面 "**" 前缀、封底 THANKS、履历表数据行
- 格式：元素越界、文本框重叠、表格填充色、字体合规、文本溢出估算
- 渲染 PNG 目视复核（依赖 soffice + pdftoppm；无渲染环境用 `--no-render` 跳过渲染，其余检查仍必做）

## 定制

- 公司名/部门/文书编号/版本号/签名人：均通过 JSON 字段覆盖，不改模板
- 内容页数量不限，`slides` 数组每项一页，顺序即页序
- 缺字体时替换：微软雅黑/黑体→Noto Sans CJK SC；仿宋→Noto Serif CJK SC；Times New Roman→Liberation Serif；Arial→Liberation Sans
- 需要全新版式（如甘特图）时，先阅读 `references/style_contract.md` 与 `references/structure_contract.md`，在 text 页基础上用 blocks 近似，或明确告知用户该版式需定制开发

## 文件清单

| 文件 | 用途 |
|---|---|
| `assets/template.pptx` | 标准模板（封面/履历/内容/封底 4 页版式） |
| `assets/*.png` | 备用图片素材，仅自定义版式时使用 |
| `scripts/build_deck.py` | JSON → PPTX 生成器 |
| `scripts/check_deck.py` | 成品检查器（生成后必跑） |
| `references/examples/*.json` | 各报告类型的示例数据 |
