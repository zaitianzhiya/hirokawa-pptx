# -*- coding: utf-8 -*-
"""公司通用PPT成品检查器：格式 + 内容 自动校验，并渲染每页PNG供目视复核。

用法:
    python3 check_deck.py 报告.pptx              # 自动检查 + 渲染PNG
    python3 check_deck.py 报告.pptx --no-render  # 仅自动检查

检查项:
  [内容] 占位符残留（（标题）/（正文）/***/2024-*-* 等）
  [内容] 封面标题是否含 "**" 前缀；封底是否有 THANKS
  [内容] 改版履历表是否有数据行
  [格式] 文本框越出页面边界
  [格式] 文本形状之间的重叠（如 CONFIDENTIAL 与其他元素交叠）
  [格式] 表格单元格填充色（表头 #28166F / 正文 #5B9BD5）
  [格式] 字体合规（微软雅黑/黑体/仿宋/Times New Roman/Arial）
  [格式] 文本溢出估算（无 autofit 的文本框）
输出: 每项 PASS/WARN/FAIL；存在 FAIL 时退出码为 1。
渲染目录: <pptx同名目录>_check/page-N.png
"""
import os
import re
import subprocess
import sys

from pptx import Presentation
from pptx.util import Emu

ALLOWED_FONTS = {'微软雅黑', '黑体', '仿宋', 'Times New Roman', 'Arial',
                 'Microsoft YaHei', 'SimHei', 'FangSong',
                 'Noto Sans CJK SC', 'Noto Serif CJK SC', '宋体',
                 'Liberation Serif', 'Liberation Sans', 'Georgia', 'Helvetica',
                 'Meiryo UI', 'メイリオ'}  # 后两个为原模板页脚原生字体
PLACEHOLDER_PATTERNS = [r'（标题）', r'（正文）', r'正文123', r'\*\*\*',
                        r'2024-\*-\*', r'单击此处', r'编辑母版']
NAVY = '28166F'
BLUE = '5B9BD5'  # 旧规范色（履历模板历史遗留），新表不应再出现
BAND = {'CDCCD5', 'E8E7EB'}  # 镶边行规范色
# 强调色白名单（商务三色方案）：蓝=结论 红=问题 绿=完成 橙=进行中
EMPHASIS_OK = {'1F4E79', 'C00000', '2E7D32', 'ED7D31',
               'FF0000',        # CONFIDENTIAL 徽标
               '000000', 'FFFFFF', '808080'}


class Report:
    def __init__(self):
        self.items = []

    def add(self, level, msg):
        self.items.append((level, msg))

    @property
    def failed(self):
        return any(lv == 'FAIL' for lv, _ in self.items)

    def dump(self):
        icons = {'PASS': '✓', 'WARN': '⚠', 'FAIL': '✗'}
        for lv, msg in self.items:
            print(f'[{lv:4}] {icons[lv]} {msg}')
        n_fail = sum(1 for lv, _ in self.items if lv == 'FAIL')
        n_warn = sum(1 for lv, _ in self.items if lv == 'WARN')
        print(f'\n结果: {"FAIL" if n_fail else "PASS"}'
              f'（{n_fail} 个失败项, {n_warn} 个警告）')


def emu_in(v):
    return Emu(v).inches


def shape_text(sh):
    return sh.text_frame.text if sh.has_text_frame else ''


def check_bounds(prs, rep):
    W, H = emu_in(prs.slide_width), emu_in(prs.slide_height)
    bad = []
    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            is_pic = str(sh.shape_type).startswith('PICTURE')
            tol = 0.1 if is_pic else 0.03  # 页脚图片允许全出血
            l, t = emu_in(sh.left), emu_in(sh.top)
            r, b = l + emu_in(sh.width), t + emu_in(sh.height)
            if l < -tol or t < -tol or r > W + tol or b > H + tol:
                bad.append(f'第{i}页 {sh.shape_id}({shape_text(sh)[:12]!r}) '
                           f'({l:.2f},{t:.2f})-({r:.2f},{b:.2f})')
    if bad:
        for m in bad:
            rep.add('FAIL', '元素越出页面: ' + m)
    else:
        rep.add('PASS', '所有元素均在页面边界内')


def check_overlap(prs, rep):
    bad = []
    for i, s in enumerate(prs.slides, 1):
        boxes = []
        for sh in s.shapes:
            if not shape_text(sh).strip():
                continue
            boxes.append((sh, (emu_in(sh.left), emu_in(sh.top),
                               emu_in(sh.left) + emu_in(sh.width),
                               emu_in(sh.top) + emu_in(sh.height))))
        for a in range(len(boxes)):
            for b in range(a + 1, len(boxes)):
                (sa, ba), (sb, bb) = boxes[a], boxes[b]
                ix = min(ba[2], bb[2]) - max(ba[0], bb[0])
                iy = min(ba[3], bb[3]) - max(ba[1], bb[1])
                if ix > 0.05 and iy > 0.15:  # 忽略轻微贴边
                    area_a = (ba[2] - ba[0]) * (ba[3] - ba[1])
                    area_b = (bb[2] - bb[0]) * (bb[3] - bb[1])
                    if ix * iy > 0.35 * min(area_a, area_b):
                        bad.append(f'第{i}页 {shape_text(sa)[:10]!r} × '
                                   f'{shape_text(sb)[:10]!r}')
    if bad:
        for m in bad:
            rep.add('FAIL', '文本元素重叠: ' + m)
    else:
        rep.add('PASS', '文本元素无重叠')


def check_placeholders(prs, rep):
    hits = []
    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            txt = shape_text(sh)
            for pat in PLACEHOLDER_PATTERNS:
                if re.search(pat, txt):
                    hits.append(f'第{i}页 含占位符 /{pat}/: {txt[:20]!r}')
    if hits:
        for m in hits:
            rep.add('FAIL', '占位符残留: ' + m)
    else:
        rep.add('PASS', '无占位符残留')


def check_structure(prs, rep):
    slides = list(prs.slides)
    n = len(slides)
    if n < 4:
        rep.add('FAIL', f'页数不足: {n} 页（至少 封面/履历/目录/封底 4页）')
        return
    rep.add('PASS', f'共 {n} 页')
    cover_txt = ''.join(shape_text(sh) for sh in slides[0].shapes)
    if '**' in cover_txt:
        rep.add('PASS', '封面标题含 "**" 前缀')
    else:
        rep.add('WARN', '封面标题缺少 "**" 前缀（模板规范要求）')
    last_txt = ''.join(shape_text(sh) for sh in slides[-1].shapes)
    if 'THANK' in last_txt.upper():
        rep.add('PASS', '封底 THANKS 正常')
    else:
        rep.add('WARN', '封底缺少 THANKS 字样')
    # 改版履历数据
    tbl_ok = False
    for sh in slides[1].shapes:
        if sh.has_table:
            tbl = sh.table
            for row in list(tbl.rows)[1:]:
                if any(c.text.strip() for c in row.cells):
                    tbl_ok = True
    rep.add('PASS' if tbl_ok else 'WARN',
            '改版履历表含数据行' if tbl_ok else '改版履历表无数据行')


def _cell_fill_hex(cell):
    try:
        if cell.fill.type is None:
            return None
        return str(cell.fill.fore_color.rgb)
    except Exception:
        return None


def check_tables(prs, rep):
    checked = 0
    bad = []
    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            if not sh.has_table:
                continue
            tbl = sh.table
            is_header_tbl = len(tbl.rows) >= 5  # 履历表有表头；概要表无表头
            prev_band = {}
            for ri, row in enumerate(tbl.rows):
                for ci, cell in enumerate(row.cells):
                    hexf = _cell_fill_hex(cell)
                    if hexf is None:
                        continue
                    if is_header_tbl and ri == 0:
                        if hexf != NAVY:
                            bad.append(f'第{i}页表头行 颜色#{hexf}≠{NAVY}')
                    elif hexf == BLUE:
                        bad.append(f'第{i}页 仍为旧规范色#5B9BD5，'
                                   '应为镶边行 CDCCD5/E8E7EB')
                    elif hexf not in BAND and hexf != NAVY:
                        bad.append(f'第{i}页 单元格颜色#{hexf}非规范色')
                    elif hexf in BAND:
                        if ci in prev_band and prev_band[ci] == hexf:
                            rep.add('WARN', f'第{i}页表格第{ri+1}行与上行'
                                            '镶边色相同，应交替')
                        prev_band[ci] = hexf
            checked += 1
    if bad:
        for m in sorted(set(bad)):
            rep.add('FAIL', '表格填充色不符: ' + m)
    elif checked:
        rep.add('PASS', f'{checked} 个表格填充色合规')
    else:
        rep.add('PASS', '无表格需检查')


def check_fonts(prs, rep):
    bad = {}

    def scan_tf(tf, where):
        for p in tf.paragraphs:
            for r in p.runs:
                for tag in ('latin', 'ea'):
                    el = r._r.rPr.find(
                        '{http://schemas.openxmlformats.org/drawingml/2006/main}'
                        + tag) if r._r.rPr is not None else None
                    if el is not None:
                        fname = el.get('typeface')
                        if fname and fname not in ALLOWED_FONTS \
                                and not fname.startswith('+'):
                            bad.setdefault(fname, where)

    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            if sh.has_text_frame:
                scan_tf(sh.text_frame, f'第{i}页')
            if sh.has_table:
                for row in sh.table.rows:
                    for c in row.cells:
                        scan_tf(c.text_frame, f'第{i}页表格')
    if bad:
        for f, w in bad.items():
            rep.add('WARN', f'非规范字体 {f!r}（{w}）')
    else:
        rep.add('PASS', '字体全部合规')


def check_text_colors(prs, rep):
    """正文字体颜色白名单：黑/白 + 商务三色强调 + 徽标红。"""
    A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    bad = {}

    def scan_tf(tf, where):
        for p in tf.paragraphs:
            for r in p.runs:
                rPr = r._r.find(A + 'rPr')
                if rPr is None:
                    continue
                fill = rPr.find(A + 'solidFill')
                if fill is None:
                    continue
                srgb = fill.find(A + 'srgbClr')
                prst = fill.find(A + 'prstClr')
                if srgb is not None:
                    c = srgb.get('val', '').upper()
                elif prst is not None:
                    c = {'black': '000000', 'white': 'FFFFFF'}.get(
                        prst.get('val'), None)
                else:
                    continue  # schemeClr 等继承色不校验
                if c and c not in EMPHASIS_OK:
                    bad.setdefault('#' + c, where)

    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            if sh.has_text_frame:
                scan_tf(sh.text_frame, f'第{i}页')
            if sh.has_table:
                for row in sh.table.rows:
                    for c in row.cells:
                        scan_tf(c.text_frame, f'第{i}页表格')
    if bad:
        for c, w in bad.items():
            rep.add('WARN', f'非白名单字体颜色 {c}（{w}），'
                            '强调色应为 蓝1F4E79/红C00000/绿2E7D32/橙ED7D31')
    else:
        rep.add('PASS', '字体颜色全部在白名单内（含强调色规范）')


def check_overflow(prs, rep):
    """无 autofit 文本框的溢出估算（CJK≈1.0em, 西文≈0.55em, 行距1.2）。"""
    bad = []

    def est(text, size_pt, width_in):
        # 每行可容纳的 CJK 字符数
        per_line = max(1, int((width_in - 0.2) * 72 / size_pt))
        lines = 0
        for seg in text.split('\n'):
            w = sum(1.0 if ord(c) > 0x2E7F else 0.55 for c in seg)
            lines += max(1, -(-int(w) // per_line))
        return lines * size_pt * 1.25 / 72 + 0.1

    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            if not sh.has_text_frame:
                continue
            tf = sh.text_frame
            bodyPr = tf._txBody.find(
                '{http://schemas.openxmlformats.org/drawingml/2006/main}bodyPr')
            has_autofit = bodyPr is not None and (
                bodyPr.find('{http://schemas.openxmlformats.org/drawingml/'
                            '2006/main}normAutofit') is not None)
            if has_autofit:
                continue
            text = tf.text.strip()
            if not text:
                continue
            sizes = [r.font.size.pt for p in tf.paragraphs
                     for r in p.runs if r.font.size]
            size = max(sizes) if sizes else 12  # 未显式设字号按12pt估
            need = est(text, size, emu_in(sh.width))
            if need > emu_in(sh.height) * 1.15 + 0.1:
                bad.append(f'第{i}页 {text[:14]!r} 估算需{need:.2f}in '
                           f'> 框高{emu_in(sh.height):.2f}in')
    if bad:
        for m in bad:
            rep.add('WARN', '文本可能溢出: ' + m)
    else:
        rep.add('PASS', '无明显文本溢出风险')


def render_pngs(pptx_path, rep):
    outdir = os.path.splitext(pptx_path)[0] + '_check'
    os.makedirs(outdir, exist_ok=True)
    try:
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf',
                        '--outdir', outdir, pptx_path],
                       check=True, capture_output=True, timeout=180)
        pdf = os.path.join(outdir, os.path.splitext(
            os.path.basename(pptx_path))[0] + '.pdf')
        subprocess.run(['pdftoppm', '-png', '-r', '80', pdf,
                        os.path.join(outdir, 'page')],
                       check=True, capture_output=True, timeout=120)
        pngs = sorted(f for f in os.listdir(outdir) if f.endswith('.png'))
        rep.add('PASS', f'已渲染 {len(pngs)} 页预览 -> {outdir}/'
                        '（请逐页目视复核）')
    except Exception as e:
        rep.add('WARN', f'渲染预览失败: {e}')
    return outdir


def main():
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    if not args:
        print(__doc__)
        sys.exit(1)
    path = args[0]
    prs = Presentation(path)
    rep = Report()
    check_structure(prs, rep)
    check_placeholders(prs, rep)
    check_bounds(prs, rep)
    check_overlap(prs, rep)
    check_tables(prs, rep)
    check_fonts(prs, rep)
    check_text_colors(prs, rep)
    check_overflow(prs, rep)
    if '--no-render' not in sys.argv:
        render_pngs(path, rep)
    print(f'=== 检查报告: {os.path.basename(path)} ===')
    rep.dump()
    sys.exit(1 if rep.failed else 0)


if __name__ == '__main__':
    main()
