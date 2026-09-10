# -*- coding: utf-8 -*-
"""公司通用PPT生成器（HIROKAWA企业模板，模板填充式，无需重新绘制版式）

用法:
    python3 build_deck.py data.json -o 输出.pptx

data.json 结构（除 slides 外均可选，缺省保留模板占位/默认）:

{
  "title":     "出差报告",                  # 封面标题（自动加 "**　" 前缀样式）
  "doc_no":    "（广川）技术A202609-01",     # 文书编号
  "company":   "上海广川科技有限公司",       # 封面公司名
  "dept":      "技术本部",                  # 封面部门
  "draw_name": "张三", "draw_date": "26/09",# 封面批准栏作成签名与日期
  "version":   "V1.0",                     # 替换所有页左下角版本号
  "footer_company": "…",                   # 替换页脚公司名（一般不用动）
  "revision":  [["00","2026-09-04","初版","张三","-"]],  # 改版履历5列
  "toc":       ["一、出差概要","二、详细内容展开"],       # 目录条目
  "thanks":    "THANKS",                   # 封底文字
  "slides": [                              # 内容页列表，每页三种类型选一:
    {"type":"text", "title":"二、详细内容展开",
     "blocks":[["h","2.1 技术交流"],["p","正文…"],["b","要点…"]]},
    {"type":"table", "title":"三、测试数据",
     "columns":["项目","结果"], "rows":[["耐压","合格"]],
     "widths":[4.0,8.0],                    # 可选，列宽(英寸)
     "blocks":[["p","备注…"]]},             # 可选，表格下方附注
    {"type":"form", "title":"一、出差概要",
     "fields":[["出差时间","…","客户名称","…"],     # 每行4元: 标签/值/标签/值
               ["出差人员","…","出差目的","…"],
               ["出席人员","…"]],                  # 2元: 值自动合并整行
     "sections":[["◆ 内容","…"],["◆ 感想","…"]]},  # 可选，表单下方小节
    {"type":"image_text", "title":"二、设备介绍",   # 图文页
     "image":"/路径/图片.png",                      # 本地图片路径
     "image_prompt":"English prompt…",              # 可选：无图时AI生图(由agent先生成再填image)
     "image_side":"right",                          # right(默认)/left
     "caption":"▲ 图注", "blocks":[["p","…"]]},
    {"type":"gallery", "title":"三、现场照片",      # 多图页(1-4张网格)
     "images":[{"path":"…","caption":"图注"}, …]},
    {"type":"fixed_image", "title":"四、固定图文",  # 固定版式图文页(图框位置锁定)
     "image":"…", "blocks":[["p","…"]]}
  ]
}

依赖: pip install python-pptx pillow

blocks 条目类型: h=小节标题(黑体20粗) / p=正文段落(18) / b=◆要点(18粗)
"""
import copy
import json
import os
import sys
import tempfile
from io import BytesIO

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn

NAVY = RGBColor(0x28, 0x16, 0x6F)
BLUE = RGBColor(0x5B, 0x9B, 0xD5)
BLACK = RGBColor(0x00, 0x00, 0x00)

TEMPLATE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        '..', 'assets', 'template.pptx')


# ---------- 基础工具 ----------

def _set_para_text(p_el, text):
    """段落级替换：保留首个run的全部rPr，删除其余run/br/fld。"""
    runs = p_el.findall(qn('a:r'))
    for tag in ('a:br', 'a:fld'):
        for el in p_el.findall(qn(tag)):
            p_el.remove(el)
    if not runs:
        r = p_el.makeelement(qn('a:r'), {})
        t = p_el.makeelement(qn('a:t'), {})
        t.text = text
        r.append(t)
        p_el.append(r)
        return
    for r in runs[1:]:
        p_el.remove(r)
    t = runs[0].find(qn('a:t'))
    if t is None:
        t = runs[0].makeelement(qn('a:t'), {})
        runs[0].append(t)
    t.text = text


def replace_text(shape, new_text):
    """整框替换文本（XML级）：完整保留首段首run格式；\n分段克隆首段。"""
    txBody = shape.text_frame._txBody
    paras = txBody.findall(qn('a:p'))
    first = paras[0]
    for p in paras[1:]:
        txBody.remove(p)
    lines = new_text.split('\n')
    _set_para_text(first, lines[0])
    for ln in lines[1:]:
        newp = copy.deepcopy(first)
        _set_para_text(newp, ln)
        txBody.append(newp)


def find_shape(slide, contains=None, shape_type=None, ph_type=None):
    for sh in slide.shapes:
        if ph_type is not None:
            try:
                if not sh.is_placeholder or \
                        str(sh.placeholder_format.type).split(' ')[0] != ph_type:
                    continue
            except Exception:
                continue
        if shape_type is not None and str(sh.shape_type).split(' ')[0] != shape_type:
            continue
        if contains is not None:
            if not sh.has_text_frame or contains not in sh.text_frame.text:
                continue
        return sh
    return None


def set_cjk(run, ea_font, latin_font=None):
    rPr = run._r.get_or_add_rPr()
    for tag, face in (('a:latin', latin_font or ea_font), ('a:ea', ea_font)):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set('typeface', face)


def duplicate_slide(prs, src_slide):
    blank = prs.slide_layouts[6]
    dst = prs.slides.add_slide(blank)
    for sh in list(dst.shapes):
        sh._element.getparent().remove(sh._element)
    dst_cSld = dst.element.find(qn('p:cSld'))
    src_cSld = src_slide.element.find(qn('p:cSld'))
    bg = src_cSld.find(qn('p:bg'))
    if bg is not None:
        old = dst_cSld.find(qn('p:bg'))
        if old is not None:
            dst_cSld.remove(old)
        dst_cSld.insert(0, copy.deepcopy(bg))
    dst_tree = dst_cSld.find(qn('p:spTree'))
    for sh in src_slide.shapes:
        el = copy.deepcopy(sh._element)
        for blip in el.iter(qn('a:blip')):
            rid = blip.get(qn('r:embed'))
            if rid:
                img_part = src_slide.part.rels[rid].target_part
                _, new_rid = dst.part.get_or_add_image_part(
                    BytesIO(img_part.blob))
                blip.set(qn('r:embed'), new_rid)
        dst_tree.append(el)
    return dst


def reorder_slides(prs, ordered_slides):
    sldIdLst = prs.slides._sldIdLst
    pairs = [(slide.slide_id, sldId)
             for sldId, slide in zip(list(sldIdLst), prs.slides)]
    key = {sid: el for sid, el in pairs}
    ordered_ids = [s.slide_id for s in ordered_slides]
    for sid, el in pairs:
        sldIdLst.remove(el)
    for sid in ordered_ids:
        sldIdLst.append(key[sid])


def prep_content_slide(slide, title):
    """设置内容页标题，返回正文占位框（已扩到内容区）。"""
    t = find_shape(slide, contains='（标题）')
    if t is not None:
        replace_text(t, title)
    b = find_shape(slide, contains='（正文）')
    if b is not None:
        b.left, b.top = Inches(0.36), Inches(1.0)
        b.width, b.height = Inches(12.6), Inches(5.6)
        b.text_frame.word_wrap = True
    return b


def write_blocks(body_shape, blocks, clear=True):
    """向正文框写入 blocks: [[kind, text], ...]，kind ∈ h/p/b。"""
    tf = body_shape.text_frame
    for p in list(tf.paragraphs[1:]):
        p._p.getparent().remove(p._p)
    p0 = tf.paragraphs[0]
    for r in list(p0.runs):
        r._r.getparent().remove(r._r)

    def add(para, kind, text):
        r = para.add_run()
        if kind == 'h':
            r.text = text
            r.font.size = Pt(20)
            r.font.bold = True
            set_cjk(r, '黑体', 'Times New Roman')
        elif kind == 'b':
            r.text = '◆ ' + text
            r.font.size = Pt(18)
            r.font.bold = True
            set_cjk(r, '仿宋', 'Times New Roman')
        else:
            r.text = text
            r.font.size = Pt(18)
            set_cjk(r, '仿宋', 'Times New Roman')
        r.font.color.rgb = BLACK
    if blocks:
        add(p0, blocks[0][0], blocks[0][1])
        for kind, text in blocks[1:]:
            add(tf.add_paragraph(), kind, text)
    for p in tf.paragraphs:
        p.space_after = Pt(8)


def style_cell(cell, text, bold=False, header=False):
    cell.fill.solid()
    cell.fill.fore_color.rgb = NAVY if header else BLUE
    cell.text = text
    for p in cell.text_frame.paragraphs:
        if header:
            p.alignment = PP_ALIGN.CENTER
        for r in p.runs:
            r.font.size = Pt(18)
            r.font.bold = bold or header
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF) if header else BLACK
            set_cjk(r, '黑体' if header else '仿宋', 'Times New Roman')


# ---------- 封面/履历/目录/封底 ----------

def fill_cover(slide, d):
    if d.get('title'):
        sh = find_shape(slide, contains='**') or find_shape(
            slide, ph_type='CENTER_TITLE')
        if sh is not None:
            title = d['title']
            replace_text(sh, title if title.startswith('**')
                         else '**　' + title)
    if d.get('doc_no'):
        sh = find_shape(slide, contains='文書番号')
        if sh is not None:
            replace_text(sh, '文書番号：' + d['doc_no'])
    if d.get('company'):
        sh = find_shape(slide, contains='技术本部') or find_shape(
            slide, contains='上海广川科技有限公司')
        if sh is not None and Emu(sh.top).inches > 4:
            replace_text(sh, d['company'] + '\n' + d.get('dept', ''))
    sig = find_shape(slide, contains='***')
    if sig is not None and (d.get('draw_name') or d.get('draw_date')):
        replace_text(sig, (d.get('draw_name') or '') + '\n'
                     + (d.get('draw_date') or ''))


def fill_revision(slide, rows):
    if not rows:
        return
    tbl_shape = find_shape(slide, shape_type='TABLE')
    if tbl_shape is None:
        return
    tbl = tbl_shape.table
    for c in range(len(tbl.columns)):
        tbl.cell(1, c).text = ''
    for i, row in enumerate(rows[:len(tbl.rows) - 1]):
        for j, val in enumerate(row[:len(tbl.columns)]):
            cell = tbl.cell(1 + i, j)
            cell.text = str(val)
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER
                for r in p.runs:
                    r.font.size = Pt(18)
                    set_cjk(r, '仿宋', 'Times New Roman')


def fill_toc(slide, items):
    b = prep_content_slide(slide, '目  录')
    if b is None:
        return
    b.left, b.top = Inches(0.8), Inches(1.6)
    b.width, b.height = Inches(11.5), Inches(4.5)
    tf = b.text_frame
    for p in list(tf.paragraphs[1:]):
        p._p.getparent().remove(p._p)
    p0 = tf.paragraphs[0]
    for r in list(p0.runs):
        r._r.getparent().remove(r._r)

    def add(para, text):
        r = para.add_run()
        r.text = text
        r.font.size = Pt(24)
        set_cjk(r, 'Arial')
        r.font.color.rgb = BLACK
    add(p0, items[0] if items else '')
    for it in items[1:]:
        add(tf.add_paragraph(), it)
    for p in tf.paragraphs:
        p.space_after = Pt(18)


def fill_closing(slide, d):
    if d.get('thanks') and d['thanks'] != 'THANKS':
        sh = find_shape(slide, contains='THANKS')
        if sh is not None:
            replace_text(sh, d['thanks'])


def global_replace(prs, d):
    for slide in prs.slides:
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            txt = sh.text_frame.text
            if d.get('version') and txt.strip().startswith('V') \
                    and len(txt.strip()) <= 6:
                replace_text(sh, d['version'])
            elif d.get('footer_company') and 'HIROKAWA' in txt \
                    and Emu(sh.top).inches > 6.5:
                replace_text(sh, d['footer_company'])


# ---------- 三种内容页 ----------

def build_text_slide(slide, spec):
    b = prep_content_slide(slide, spec.get('title', ''))
    if b is not None:
        write_blocks(b, spec.get('blocks') or [])


def build_table_slide(slide, spec):
    b = prep_content_slide(slide, spec.get('title', ''))
    cols = spec.get('columns') or []
    rows = spec.get('rows') or []
    n_rows, n_cols = 1 + len(rows), max(1, len(cols))
    widths = spec.get('widths') or [12.6 / n_cols] * n_cols
    blocks = spec.get('blocks') or []
    tbl_h = min(0.55 * n_rows, 4.6 if not blocks else 3.6)
    gf = slide.shapes.add_table(n_rows, n_cols, Inches(0.36), Inches(1.1),
                                Inches(sum(widths)), Inches(tbl_h))
    tbl = gf.table
    for j, w in enumerate(widths):
        tbl.columns[j].width = Inches(w)
    for j, c in enumerate(cols):
        style_cell(tbl.cell(0, j), str(c), header=True)
    for i, row in enumerate(rows):
        for j in range(n_cols):
            val = row[j] if j < len(row) else ''
            style_cell(tbl.cell(1 + i, j), str(val))
    if blocks and b is not None:
        b.top = Inches(1.3 + tbl_h)
        b.height = Inches(6.6 - 1.3 - tbl_h)
        write_blocks(b, blocks)
    elif b is not None:
        b._element.getparent().remove(b._element)


def build_form_slide(slide, spec):
    b = prep_content_slide(slide, spec.get('title', ''))
    fields = spec.get('fields') or []
    sections = spec.get('sections') or []
    n_rows = max(1, len(fields))
    tbl_h = min(0.8 * n_rows, 3.2)
    gf = slide.shapes.add_table(n_rows, 4, Inches(0.36), Inches(1.1),
                                Inches(12.6), Inches(tbl_h))
    tbl = gf.table
    tbl.columns[0].width = Inches(1.6)
    tbl.columns[1].width = Inches(4.7)
    tbl.columns[2].width = Inches(1.6)
    tbl.columns[3].width = Inches(4.7)
    for i, row in enumerate(fields):
        row = list(row) + [''] * (4 - len(row))
        for j in range(4):
            style_cell(tbl.cell(i, j), str(row[j]), bold=(j % 2 == 0))
        if len(fields[i]) <= 2:  # 仅 标签+值：值合并整行
            tbl.cell(i, 1).merge(tbl.cell(i, 3))
    if b is not None:
        if sections:
            b.top = Inches(1.3 + tbl_h)
            b.height = Inches(6.7 - 1.3 - tbl_h)
            blocks = []
            for head, text in sections:
                blocks.append(['b', head.replace('◆', '').strip()])
                if text:
                    blocks.append(['p', text])
                blocks.append(['p', ''])
            write_blocks(b, blocks[:-1] if blocks else [])
        else:
            b._element.getparent().remove(b._element)


SLIDE_BUILDERS = {
    'text': build_text_slide,
    'table': build_table_slide,
    'form': build_form_slide,
}

# ---------- 方案A/B：插图页型 ----------

def _fit_crop(src_path, target_w_in, target_h_in, out_path):
    """按目标宽高比中心裁剪图片，避免变形。"""
    from PIL import Image
    img = Image.open(src_path).convert('RGB')
    tw, th = target_w_in, target_h_in
    w, h = img.size
    src_ratio, dst_ratio = w / h, tw / th
    if src_ratio > dst_ratio:
        nw = int(h * dst_ratio)
        x = (w - nw) // 2
        img = img.crop((x, 0, x + nw, h))
    else:
        nh = int(w / dst_ratio)
        y = (h - nh) // 2
        img = img.crop((0, y, w, y + nh))
    img.save(out_path, quality=92)
    return out_path


def add_picture_fit(slide, path, left, top, width, height, border=False):
    """裁剪适配后插入图片，可选藏青边框。"""
    import uuid
    tmp = os.path.join(tempfile.gettempdir(),
                       f'_crop_{uuid.uuid4().hex[:8]}.jpg')
    try:
        _fit_crop(path, Emu(width).inches, Emu(height).inches, tmp)
        pic = slide.shapes.add_picture(tmp, left, top, width, height)
        if border:
            pic.line.color.rgb = NAVY
            pic.line.width = Pt(1.5)
        return pic
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


def build_image_text_slide(slide, spec):
    """方案A 图文页：图片居左/右约45%宽，另一侧为文字blocks。"""
    b = prep_content_slide(slide, spec.get('title', ''))
    side = spec.get('image_side', 'right')
    img_path = spec.get('image')
    cap = spec.get('caption')
    if side == 'right':
        img_pos = (Inches(7.3), Inches(1.3), Inches(5.5), Inches(4.2))
        txt_pos = (Inches(0.36), Inches(1.1), Inches(6.6), Inches(5.4))
    else:
        img_pos = (Inches(0.36), Inches(1.3), Inches(5.5), Inches(4.2))
        txt_pos = (Inches(6.3), Inches(1.1), Inches(6.6), Inches(5.4))
    if img_path and os.path.exists(img_path):
        add_picture_fit(slide, img_path, *img_pos)
        if cap:
            tb = slide.shapes.add_textbox(img_pos[0],
                                          Inches(5.55), img_pos[2], Inches(0.4))
            r = tb.text_frame.paragraphs[0].add_run()
            r.text = cap
            r.font.size = Pt(12)
            set_cjk(r, '仿宋', 'Times New Roman')
            tb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    if b is not None:
        b.left, b.top, b.width, b.height = txt_pos
        write_blocks(b, spec.get('blocks') or [])


def build_gallery_slide(slide, spec):
    """方案A 多图页：1-4 张图片网格 + 图注。"""
    b = prep_content_slide(slide, spec.get('title', ''))
    if b is not None:
        b._element.getparent().remove(b._element)
    images = (spec.get('images') or [])[:4]
    n = len(images)
    if n == 0:
        return
    cols = 1 if n == 1 else 2
    rows = 1 if n <= 2 else 2
    gap = 0.3
    area_l, area_t, area_w, area_h = 0.36, 1.2, 12.6, 5.3
    cw = (area_w - gap * (cols - 1)) / cols
    ch = (area_h - gap * (rows - 1)) / rows - 0.4
    for i, item in enumerate(images):
        r_i, c_i = divmod(i, cols)
        l = Inches(area_l + c_i * (cw + gap))
        t = Inches(area_t + r_i * (ch + 0.4 + gap))
        path = item.get('path', '')
        if os.path.exists(path):
            add_picture_fit(slide, path, l, t, Inches(cw), Inches(ch),
                            border=True)
        tb = slide.shapes.add_textbox(l, Inches(
            area_t + r_i * (ch + 0.4 + gap) + ch + 0.02), Inches(cw),
            Inches(0.35))
        run = tb.text_frame.paragraphs[0].add_run()
        run.text = item.get('caption', '')
        run.font.size = Pt(14)
        run.font.bold = True
        set_cjk(run, '黑体', 'Times New Roman')
        tb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER


def build_fixed_image_slide(slide, spec):
    """方案B 固定版式图文页：替换模板中的【图片占位】框与正文框。"""
    t = find_shape(slide, contains='（标题）')
    if t is not None:
        replace_text(t, spec.get('title', ''))
    frame = find_shape(slide, contains='【图片占位】')
    if frame is not None and spec.get('image')             and os.path.exists(spec['image']):
        l, tp, w, h = frame.left, frame.top, frame.width, frame.height
        frame._element.getparent().remove(frame._element)
        add_picture_fit(slide, spec['image'], l, tp, w, h, border=True)
    b = find_shape(slide, contains='（正文）')
    if b is not None:
        write_blocks(b, spec.get('blocks') or [])


SLIDE_BUILDERS['image_text'] = build_image_text_slide
SLIDE_BUILDERS['gallery'] = build_gallery_slide
SLIDE_BUILDERS['fixed_image'] = build_fixed_image_slide



# ---------- 主流程 ----------

def build(data, out_path, template=None):
    prs = Presentation(template or TEMPLATE)
    all_slides = list(prs.slides)
    s_cover, s_rev, s_content, s_closing = all_slides[:4]
    # 模板第5页（可选）为 fixed_image 固定图文版式基页
    s_fixed = all_slides[4] if len(all_slides) > 4 else None

    specs = data.get('slides') or []
    body_slides = []
    content_used = False
    for spec in specs:
        if spec.get('type') == 'fixed_image' and s_fixed is not None:
            body_slides.append(duplicate_slide(prs, s_fixed))
        elif not content_used:
            body_slides.append(s_content)  # 复用模板自带内容页
            content_used = True
        else:
            body_slides.append(duplicate_slide(prs, s_content))
    s_toc = duplicate_slide(prs, s_content)

    fill_cover(s_cover, data)
    fill_revision(s_rev, data.get('revision'))
    fill_toc(s_toc, data.get('toc') or
             [s.get('title', '') for s in specs])
    for slide, spec in zip(body_slides, specs):
        builder = SLIDE_BUILDERS.get(spec.get('type', 'text'))
        builder(slide, spec)
    fill_closing(s_closing, data)
    global_replace(prs, data)

    ordered = [s_cover, s_rev, s_toc] + body_slides + [s_closing]
    reorder_slides(prs, ordered)  # 未使用的版式基页自动从成品中剔除
    prs.save(out_path)
    return out_path


def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        sys.exit(1)
    out = 'output.pptx'
    tpl = None
    if '-o' in args:
        out = args[args.index('-o') + 1]
    if '-t' in args:  # 指定扩展模板（如含固定图文版式的 template_v2）
        tpl = args[args.index('-t') + 1]
    with open(args[0], encoding='utf-8') as f:
        data = json.load(f)
    build(data, out, template=tpl)
    print('OK ->', out)


if __name__ == '__main__':
    main()
