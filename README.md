# hirokawa-pptx — 公司通用 PPT 技能（HIROKAWA 企业模板）

按上海广川科技（HIROKAWA）企业模板一键生成公司标准 PPT 的 Claude 技能：
出差报告、项目进度、周报月报、技术方案、会议纪要、客户拜访、产品介绍等。
版式统一为：封面 → 改版履历 → 目录 → 内容页×N → 封底。

## 特性

- 模板填充式生成，100% 还原公司品牌规范（藏青 #28166F / 表格镶边行 #CDCCD5·#E8E7EB 交替）
- 内容页六种类型：`text` / `table` / `form` / `image_text` / `gallery` / `fixed_image`
- 图片自动中心裁剪适配（不变形），多图自动网格 + 藏青边框
- 重点凸显：正文/表格支持 `[blue]`/`[red]`/`[green]`/`[orange]` 行内强调色（商务三色）
- 内置成品检查器：占位符残留、元素越界、文本框重叠、表格填充色、字体颜色白名单、字体合规、文本溢出
- 支持 `image_prompt` AI 生图工作流（无实拍图时先用 AI 生成再填 `image`）

## 目录结构

```
hirokawa-pptx/
├── SKILL.md                       # 技能说明
├── assets/
│   ├── template.pptx              # 标准模板（5 页版式，含固定图文基页）
│   └── *.png                      # 备用图片素材（仅自定义版式时使用）
├── references/
│   ├── examples/                  # 示例 data.json：出差报告 / 项目进度 / 通用汇报 / 产品介绍
│   ├── structure_contract.md      # 结构契约
│   └── style_contract.md          # 视觉规范
└── scripts/
    ├── build_deck.py              # JSON → PPTX 生成器
    └── check_deck.py              # 成品检查器（生成后必跑）
```

## 快速上手

```bash
pip install python-pptx pillow

# 1. 写 data.json（示例见 references/examples/）
# 2. 生成
python3 scripts/build_deck.py data.json -o 报告.pptx
# 3. 检查（必做）
python3 scripts/check_deck.py 报告.pptx        # 渲染 PNG 目视复核（依赖 soffice + pdftoppm）
python3 scripts/check_deck.py 报告.pptx --no-render   # 无渲染环境仅自动校验
```

## data.json 结构（节选）

```json
{
  "title": "出差报告",
  "doc_no": "（广川）技术A202609-01",
  "company": "上海广川科技有限公司",
  "dept": "技术本部",
  "version": "V1.0",
  "revision": [["00", "2026-09-04", "初版", "张三", "-"]],
  "toc": ["一、出差概要", "二、详细内容展开"],
  "slides": [
    {"type": "form", "title": "一、出差概要",
     "fields": [["出差时间", "…", "客户名称", "…"], ["出席人员", "…"]],
     "sections": [["◆ 内容", "…"]]},
    {"type": "text", "title": "二、详细内容展开",
     "blocks": [["h", "2.1 技术交流"], ["p", "正文…"], ["b", "要点…"]]}
  ],
  "thanks": "THANKS"
}
```

完整 schema 见 `hirokawa-pptx/SKILL.md` 或 `scripts/build_deck.py` 顶部注释。

## 安装为 Claude 个人技能

桌面 app 的「Save skill」目前只支持单文件 SKILL.md。可将技能打包为单文件 `.skill` 后安装，
或将 `hirokawa-pptx/` 目录拷入团队能力库的 `skills/` 目录按目录使用。
