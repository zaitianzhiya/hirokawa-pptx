# Style Contract

Reference source type: Uploaded PPTX artifact  
Reference artifact type: PPTX  
Reference File Type: PPTX  

## Color Palette

| Token | Hex | Usage |
|---|---|---|
| Primary Navy | #28166F | Cover/thanks backgrounds, section title text, footer bar, table header backgrounds, footer text |
| Secondary Blue | #5B9BD5 | Table body cell fill (uniform for all data rows on both revision history and trip summary tables) |
| Red Accent | #DA2424 | Decorative roof element on graphic |
| Red Pure | #FF0000 | "CONFIDENTIAL" badge text and border, department-usage label "（技术部内部用）" |
| White | #FFFFFF | Text on dark backgrounds, content-slide backgrounds, table header text, decorative white bars |
| Black/Dark | #000000 | Body text on light backgrounds |

## Typography System

Extracted directly from the reference PPTX:

- **Cover title**: 微软雅黑 (Microsoft YaHei), 44pt, bold, white on dark navy background
- **Cover prefix stars**: 微软雅黑, 44pt, bold, white — literal text "**" (two asterisks) before the title
- **Content slide titles**: 黑体 (SimHei), 24pt, bold, black
- **Section subtitle (TOC entries)**: Arial, 18pt, black
- **Table header**: 黑体, 18pt, bold, white on navy background
- **Table body / form content**: 仿宋 (FangSong) for CJK, Times New Roman for Latin/numeric, 18pt, black on #5B9BD5 fill
- **Body text**: Times New Roman, 18pt, black
- **Small labels**: 微软雅黑, 10-16pt depending on context
- **Closing "THANKS"**: Times New Roman, 66pt, bold, white on navy background
- **Footer**: Company name text in white, slide number in white

**CJK strategy**: The reference uses distinct fonts per context — 微软雅黑 for cover/branding, 黑体 for titles/headers, 仿宋 for table body CJK text, Arial for TOC entries. When generating output, use 微软雅黑 for cover elements, 黑体 for content slide titles and table headers, and 仿宋 for table body CJK content. If these specific fonts are unavailable, substitute with the closest stylistic equivalent (e.g., 宋体 or Noto Sans CJK for 黑体, Noto Serif CJK for 仿宋).

## Page/Screen Composition

- **Aspect ratio**: 16:9 widescreen
- **Slide dimensions**: 13.33" x 7.5"
- **Cover slide**: Full-bleed dark navy background (#28166F)
- **Content slides**: White background with content area
- **Closing slide**: Full-bleed dark navy background matching cover

## Header / Top Bar (Content Slides)

- **Top-right corner**: HIROKAWA corporate logo (color version: `assets/logo-color.png`) positioned at ~right edge, vertically near top
- **CONFIDENTIAL badge**: Red-bordered rectangular box with red "CONFIDENTIAL" text, positioned at top-right
- **Department usage note**: "（技术部内部用）" in red text, positioned below the CONFIDENTIAL badge, slightly left-aligned with it

## Header (Cover & Closing Slides)

- **Top-left corner**: HIROKAWA corporate logo (white version: `assets/logo-white.png`)
- **Top-right corner**: Same CONFIDENTIAL badge + "（技术部内部用）" red text as content slides

## Footer Bar

- Full-width dark navy (#28166F) horizontal bar at bottom (~0.47" height)
- Left: "上海广川科技有限公司 Shanghai HIROKAWA Technology Co.,Ltd." in white text
- Right: Slide/page number in white text
- The footer bar is an image asset: `assets/footer-bar.png`

## Title Underline (Content Slides)

- Thin horizontal navy line under the section title
- Position: directly below the title text box, spanning approximately the title width
- Color: #28166F

## Table Styling（2026-09 更新：镶边行规范）

- **Header row**: Background #28166F, text white, bold, 黑体, 18pt
- **Body rows（镶边行）**: 双色交替填充 —— 奇数行 #CDCCD5，偶数行 #E8E7EB，黑字
- **Borders**: Thin white grid lines between cells
- **Cell alignment**: Center for short labels, left for long content
- **Form-style tables** (trip summary): Labels in bold, values in regular weight
- 旧模板的 #5B9BD5 统一蓝色已废弃，生成时全部重着色为镶边行

## Emphasis Scheme（重点凸显统一方案：商务三色）

强调 = **加粗 + 变色**同时使用；单页强调色不超过 2 种；强调文字占比 ≤15%；
不单独用颜色传递信息（需配合文字语义）。

| 颜色 | 色值 | 应用场景 |
|---|---|---|
| 藏青（蓝） | #1F4E79 | 关键结论、核心数据、标题级强调 |
| 红 | #C00000 | 问题、风险、延期、不合格 |
| 绿 | #2E7D32 | 完成、达成、合格、正向结果 |
| 橙 | #ED7D31 | 进行中、待关注事项 |

行内标记语法（生成器自动解析）：`**加粗**`、`[red]…[/red]`、
`[blue]…[/blue]`、`[green]…[/green]`、`[orange]…[/orange]`（带色即加粗）。

## Decorative Elements

- **Cover/closing graphic**: Abstract geometric shape — red chevron roof pointing UPWARD (`assets/decorative-graphic.png`) + three white vertical bars beneath it, positioned right-of-center on cover and closing slides. The graphic must NOT be inverted or rotated 180° — the red roof must point upward like a house roof.
- **Diamond bullets**: Black diamond (◆) used as subsection markers before headings
- **Title underline**: Thin horizontal navy line under section titles on content slides

## Visual Density and Rhythm

- Content slides have generous whitespace
- Tables dominate content area for structured data
- Form-style tables use multi-column key/value layout with bold label cells
- Subsections use indented diamond bullet markers
- Trip summary table spans most of the slide width
