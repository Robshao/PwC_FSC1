import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

src, out, idx = sys.argv[1], sys.argv[2], int(sys.argv[3])
p = Presentation(src)
s = p.slides[idx]
keep = {'Title 6', 'Footer Placeholder 1', 'Slide Number Placeholder 2'}
for sh in list(s.shapes):
    if sh.name not in keep:
        sh._element.getparent().remove(sh._element)

ORANGE, MID, LIGHT, PALE, CREAM = "FD5108", "FE7C39", "FFAA72", "FFE8D4", "FFF5ED"
DARK, GREY, LGREY, TEXT, WHITE = "464646", "7D7D7D", "E8E8E8", "2D2D2D", "FFFFFF"
rgb = lambda h: RGBColor.from_string(h)

def box(x, y, w, h, text, fill, color=TEXT, size=11, bold=False, shape=MSO_SHAPE.RECTANGLE, line=None, align=PP_ALIGN.CENTER, name=None):
    sp = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if name: sp.name = name
    sp.fill.solid(); sp.fill.fore_color.rgb = rgb(fill)
    if line: sp.line.color.rgb = rgb(line); sp.line.width = Pt(0.75)
    else: sp.line.fill.background()
    sp.shadow.inherit = False
    tf = sp.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.06); tf.margin_top = tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    lines = text if isinstance(text, list) else [text]
    for i, t in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        r = para.add_run(); r.text = t
        r.font.size = Pt(size if i == 0 else size - 2); r.font.bold = bold if i == 0 else False
        r.font.color.rgb = rgb(color)
    return sp

def line(x1, y1, x2, y2, color=GREY, w=1.25):
    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = rgb(color); c.line.width = Pt(w)
    return c

# title
t = s.shapes.title
t.text_frame.paragraphs[0].runs[0].text = "2030 的 BA：設計人與 AI 一起工作的方式"

# ----- left column: today -----
box(0.4, 1.55, 2.3, 0.45, "今天的 BA", GREY, WHITE, 13, True, name="Today header")
today = ["寫需求，交給技術團隊", "會議紀錄、文件靠人工", "知識留在個人腦中", "測試太晚，需求與驗證脫節", "AI 隨手用，客戶不敢信"]
for i, tx in enumerate(today):
    box(0.4, 2.12 + i * 0.62, 2.3, 0.52, tx, LGREY, TEXT, 11, name=f"Today {i+1}")
box(2.78, 3.55, 0.4, 0.55, "", ORANGE, shape=MSO_SHAPE.CHEVRON, name="Arrow")

# ----- center -----
cx, cy = 6.65, 4.05
box(cx - 1.1, cy - 1.1, 2.2, 2.2, ["2030 的 BA", "動手做的", "翻譯者"], ORANGE, WHITE, 16, True, shape=MSO_SHAPE.OVAL, name="Center")

# ----- four branches with leaves -----
branches = [
    ("現場共創", ORANGE, WHITE, 3.25, 1.55, [("工作坊當場做出可用原型", "需求就是會動的原型"), ("變革教練", "帶使用者走過改變")]),
    ("脈絡與 AI 代理", MID, WHITE, 8.05, 1.55, [("脈絡架構師", "業務知識變成 AI 讀得懂的資產"), ("AI 代理團隊指揮官", "AI 起草、BA 決策與核准")]),
    ("信任設計", DARK, WHITE, 3.25, 4.55, [("AI 驗收標準設計師", "定義合格與 go／no-go 關卡"), ("人在迴圈內", "對應金管會指引與公會自律規範")]),
    ("營運模式", LIGHT, TEXT, 8.05, 4.55, [("規則即程式", "改一條規則，影響一目了然"), ("業務數位分身", "上線前先模擬政策變更"), ("人機分工設計", "直通、輔助、專家三條車道")]),
]
bw = 2.35
for name, fill, fg, x, y, leaves in branches:
    b = box(x, y, bw, 0.5, name, fill, fg, 14, True, shape=MSO_SHAPE.ROUNDED_RECTANGLE, name=f"Branch {name}")
    # connect center to branch inner edge
    bx = x + bw if x < cx else x
    by = y + 0.25
    line(cx + (-0.78 if x < cx else 0.78), cy + (-0.78 if y < cy else 0.78), bx, by, fill if fill != LIGHT else MID, 2)
    lh, gap = (0.55, 0.08) if len(leaves) < 3 else (0.48, 0.06)
    for i, (h, sub) in enumerate(leaves):
        ly = y + 0.62 + i * (lh + gap)
        box(x + 0.25, ly, bw - 0.25, lh, [h, sub], CREAM, TEXT, 11, True, line=fill if fill != LIGHT else MID, align=PP_ALIGN.LEFT, name=f"Leaf {h}")
        line(x + 0.1, y + 0.5, x + 0.1, ly + lh / 2, GREY, 0.75)
        line(x + 0.1, ly + lh / 2, x + 0.25, ly + lh / 2, GREY, 0.75)

# ----- right column: skills -----
box(10.6, 1.55, 2.35, 0.45, "需要的 8 項能力", ORANGE, WHITE, 13, True, name="Skills header")
skills = ["AI 動手做原型", "產品排序思維", "業務架構視野", "脈絡工程", "驗收與評估設計", "AI 治理素養", "用業務語言說故事", "變革領導"]
for i, tx in enumerate(skills):
    box(10.6, 2.1 + i * 0.58, 2.35, 0.48, tx, PALE if i % 2 == 0 else CREAM, TEXT, 11, line=LIGHT, name=f"Skill {i+1}")

for sh in s.shapes:
    if sh.name == 'Footer Placeholder 1':
        sh.text_frame.paragraphs[0].runs[0].text = "依 graphify 知識圖譜整理（5 份公開資料、80 個概念、8 個群組）"
        for r in sh.text_frame.paragraphs[0].runs[1:]: r.text = ""

s.notes_slide.notes_text_frame.text = "\n".join([
    "【願景心智圖，約 60 秒】",
    "這張圖是我用 graphify 把五份公開資料做成知識圖譜後整理的：BlackRock 的職缺、金管會指引、Gartner 趨勢、McKinsey UWOS，還有 Alexandr Wang 的對談。",
    "左邊是今天的 BA：寫需求交給技術團隊、文件靠人工、知識在個人腦中、測試太晚、AI 隨手用客戶不敢信。",
    "中間是我想像的 2030 的 BA：一個會動手做的翻譯者，往四個方向長出新角色。",
    "現場共創：工作坊當場做出可用原型，同時當使用者的變革教練。",
    "脈絡與 AI 代理：把業務知識變成 AI 讀得懂的資產，並指揮一組 AI 代理，AI 起草、我們決策。",
    "信任設計：替 AI 產出定義合格標準和 go／no-go 關卡，這也是金管會要求的人在迴圈內。",
    "營運模式：規則即程式、業務數位分身、人機三條車道。這些聽起來很遠，但 BlackRock 的職缺已經在要求其中好幾項。",
    "圖譜裡串起最多群組的節點，就是「信任設計」，所以我認為它是我們最該先投資的能力。",
    "右邊是要具備的八項能力，下一頁我挑出三個今年就能開始的方向。",
])
p.save(out)
