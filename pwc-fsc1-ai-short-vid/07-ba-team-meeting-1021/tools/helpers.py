"""Apply Robert's 11/2 review comments to the BA meeting deck (structure already reordered in base_rev.pptx)."""
import sys, copy, math
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.oxml.ns import qn


ORANGE, MID, LIGHT, PALE, CREAM = "FD5108", "FE7C39", "FFAA72", "FFE8D4", "FFF5ED"
DARK, GREY, LGREY, TEXT, WHITE = "464646", "7D7D7D", "E8E8E8", "2D2D2D", "FFFFFF"
rgb = lambda h: RGBColor.from_string(h)


# ---------------------------------------------------------------- helpers
def all_shapes(shapes):
    for sh in shapes:
        yield sh
        if sh.shape_type == 6:
            yield from all_shapes(sh.shapes)

def find(slide, sid):
    for sh in all_shapes(slide.shapes):
        if sh.shape_id == sid:
            return sh
    raise KeyError(sid)

def set_tf(tf, items):
    """items: list of (prototype paragraph index, text); keeps template formatting."""
    body = tf._txBody
    ps = body.findall(qn('a:p'))
    protos = list(ps)
    for x in ps:
        body.remove(x)
    for idx, text in items:
        np_ = copy.deepcopy(protos[min(idx, len(protos) - 1)])
        runs = np_.findall(qn('a:r'))
        for r in runs[1:]:
            np_.remove(r)
        for tag in ('a:br', 'a:fld'):
            for x in np_.findall(qn(tag)):
                np_.remove(x)
        if runs:
            runs[0].find(qn('a:t')).text = text
        else:
            r = np_.makeelement(qn('a:r'), {}); t = r.makeelement(qn('a:t'), {}); t.text = text; r.append(t)
            end = np_.find(qn('a:endParaRPr'))
            (end.addprevious(r) if end is not None else np_.append(r))
        body.append(np_)

def st(slide, sid, *lines):
    set_tf(find(slide, sid).text_frame, [(0, l) for l in lines])

def num(slide, sid, n):
    fld = find(slide, sid).text_frame._txBody.find(qn('a:p')).find(qn('a:fld'))
    fld.find(qn('a:t')).text = str(n)

def notes(slide, lines):
    slide.notes_slide.notes_text_frame.text = "\n".join(lines)

def clear(slide, keep_ids):
    for sh in list(slide.shapes):
        if sh.shape_id not in keep_ids:
            sh._element.getparent().remove(sh._element)

def box(slide, x, y, w, h, text, fill, color=TEXT, size=11, bold=False, shape=MSO_SHAPE.RECTANGLE,
        line=None, dash=False, align=PP_ALIGN.CENTER, name=None, lw=0.75, sub_size=None):
    sp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if name: sp.name = name
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = rgb(fill)
    if line:
        sp.line.color.rgb = rgb(line); sp.line.width = Pt(lw)
        if dash: sp.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    else:
        sp.line.fill.background()
    sp.shadow.inherit = False
    tf = sp.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.05); tf.margin_top = tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    lines = text if isinstance(text, list) else [text]
    for i, t in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        r = para.add_run(); r.text = t
        r.font.size = Pt(size if i == 0 else (sub_size or size - 2))
        r.font.bold = bold if i == 0 else False
        r.font.color.rgb = rgb(color)
    return sp

def text(slide, x, y, w, h, lines, size=12, color=TEXT, bold=False, align=PP_ALIGN.LEFT, name=None):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name: tb.name = name
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.02)
    for i, t in enumerate(lines if isinstance(lines, list) else [lines]):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        r = para.add_run(); r.text = t; r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = rgb(color)
    return tb

def line(slide, x1, y1, x2, y2, color=GREY, w=1.25, dash=False):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = rgb(color); c.line.width = Pt(w)
    if dash: c.line.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT
    return c

def circle(slide, cx, cy, d, text_, fill, color=WHITE, size=11, bold=True, line_=None, dash=False, name=None, sub_size=None):
    return box(slide, cx - d / 2, cy - d / 2, d, d, text_, fill, color, size, bold, MSO_SHAPE.OVAL, line_, dash, name=name, sub_size=sub_size)

def edge_point(cx, cy, r, tx, ty):
    a = math.atan2(ty - cy, tx - cx)
    return cx + r * math.cos(a), cy + r * math.sin(a)

def link(slide, a, b, color, w=1.5, dash=False):
    (ax, ay, ar), (bx, by, br) = a, b
    x1, y1 = edge_point(ax, ay, ar, bx, by); x2, y2 = edge_point(bx, by, br, ax, ay)
    return line(slide, x1, y1, x2, y2, color, w, dash)

