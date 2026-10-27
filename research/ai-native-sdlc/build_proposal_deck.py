# AI 原生軟體開發：客戶提案簡報（以「面向誰、解決什麼問題、帶來什麼價值、可擴充與可複製」四個面向為主軸）
# 依據：Anthropic《The AI-Native SDLC Playbook》、Rex v3 五階段流程、MIT SMR 產品化框架、同業產品化盤點、保險科技研究
# 用法：python3 build_proposal_deck.py <範本.pptx> <輸出.pptx> <手冊.pdf>
#   手冊的圖在建置時從 PDF 裁切，不放進版本控制
import sys, copy, subprocess, tempfile
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from PIL import Image, ImageChops

TPL, OUT, PDF = sys.argv[1], sys.argv[2], sys.argv[3]
OR = RGBColor(0xFD, 0x51, 0x08); OR2 = RGBColor(0xFE, 0x7C, 0x39); TINT = RGBColor(0xFF, 0xE8, 0xD4)
# 範本佈景主題色：tx1 黑、lt2 淺灰、accent6 線條灰、accent4 中灰
INK = RGBColor(0x00, 0x00, 0x00); GREY = RGBColor(0x59, 0x59, 0x59); LG = RGBColor(0xEB, 0xEB, 0xEB); LINE = RGBColor(0xCB, 0xD1, 0xD6)
MID = RGBColor(0xA1, 0xA8, 0xB3); WHITE = RGBColor(0xFF, 0xFF, 0xFF)
# 範本字型：標題 Georgia、內文 Arial、中文一律微軟正黑體
FONT = 'Arial'; HEAD_FONT = 'Georgia'; EA_FONT = '微軟正黑體'
SYMBOLS = set('●◐○✓✗→↑①②③④⑤⑥⑦⚠')
FOOT = 'PwC FSC1｜提案討論稿｜標示「建議」「示意」者非實測數據或報價，實際以診斷與試點結果為準'
V3 = 'Rex《AI Agent 協作開發流程與產出範本》v3'
PB = 'Anthropic《The AI-Native SDLC Playbook》'
SMR = 'Sawhney, MIT Sloan Management Review, 2024'

def style_font(r, head=False):
    # Arial 沒有 ●◐○ 等符號，這類文字改用中文字型顯示
    r.font.name = EA_FONT if SYMBOLS & set(r.text) else (HEAD_FONT if head else FONT)
    rPr = r._r.get_or_add_rPr()
    for e in rPr.findall(qn('a:ea')): rPr.remove(e)
    rPr.find(qn('a:latin')).addnext(rPr.makeelement(qn('a:ea'), {'typeface': EA_FONT}))

p = Presentation(TPL)
lst = p.slides._sldIdLst
for sid in list(lst):
    p.part.drop_rel(sid.rId); lst.remove(sid)
L = {l.name: l for l in p.slide_masters[2].slide_layouts}
# 移除範本遺留的投影片章節（PowerPoint 左側的章節名稱）
for ext in list(p.part._element.iter('{http://schemas.openxmlformats.org/presentationml/2006/main}ext')):
    if ext.find('{http://schemas.microsoft.com/office/powerpoint/2010/main}sectionLst') is not None:
        ext.getparent().remove(ext)

def set_ph(s, idx, text, size=None, bold=None):
    for ph in s.placeholders:
        if ph.placeholder_format.idx == idx:
            ph.text = text
            for para in ph.text_frame.paragraphs:
                for r in para.runs:
                    style_font(r, head=idx == 0)
                    if size: r.font.size = Pt(size)
                    if bold is not None: r.font.bold = bold
            return ph

def tb(s, x, y, w, h, text, size=11, bold=False, color=INK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, fill=None, line=None, margin=0.06, shape=MSO_SHAPE.RECTANGLE):
    sh = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.shadow.inherit = False
    if fill is None: sh.fill.background()
    else: sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if line is None: sh.line.fill.background()
    else: sh.line.color.rgb = line; sh.line.width = Pt(0.75)
    tf = sh.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    for side in ('left', 'right', 'top', 'bottom'): setattr(tf, f'margin_{side}', Inches(margin))
    lines = text if isinstance(text, list) else [text]
    for i, ln in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        segs = ln if isinstance(ln, list) else [(ln, {})]
        for seg in segs:
            t, st = (seg, {}) if isinstance(seg, str) else seg
            r = para.add_run(); r.text = t
            style_font(r); f = r.font; f.size = Pt(st.get('size', size)); f.bold = st.get('bold', bold)
            f.color.rgb = st.get('color', color)
            if 'link' in st:
                r.hyperlink.address = st['link']; f.underline = True
        para.space_after = Pt(2)
    return sh

def circle(s, x, y, d, text, fill=OR, size=12):
    c = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    c.fill.solid(); c.fill.fore_color.rgb = fill; c.line.fill.background(); c.shadow.inherit = False
    tf = c.text_frame; tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    para = tf.paragraphs[0]; para.alignment = PP_ALIGN.CENTER
    r = para.add_run(); r.text = text; r.font.size = Pt(size); r.font.bold = True; r.font.color.rgb = WHITE; style_font(r)

def hline(s, x1, y1, x2, y2, color=LINE, width=1.0):
    ln = s.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    ln.line.color.rgb = color; ln.line.width = Pt(width)

def arrow(s, x, y, w, h, kind=MSO_SHAPE.RIGHT_ARROW, fill=OR):
    a = s.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    a.fill.solid(); a.fill.fore_color.rgb = fill; a.line.fill.background(); a.shadow.inherit = False

def footer(s):
    # python-pptx 新增投影片不會帶入頁尾與頁碼，從版面複製（頁碼為自動欄位）
    have = {ph.placeholder_format.idx for ph in s.placeholders}
    for ph in s.slide_layout.placeholders:
        if ph.placeholder_format.idx in (3, 4) and ph.placeholder_format.idx not in have:
            s.shapes._spTree.append(copy.deepcopy(ph._element))
    set_ph(s, 3, FOOT)

def content(title):
    s = p.slides.add_slide(L['只有標題'])
    set_ph(s, 0, title, 20, bold=True); footer(s)
    return s

def note(s, text, y=6.66):
    tb(s, 0.45, y, 12.45, 0.32, text, size=9, color=GREY, margin=0.03)

def section(num, title, ask, see):
    # 章節轉場頁：沿用範本「章節標題」版面，左上方說明這一章要回答什麼、會看到什麼
    s = p.slides.add_slide(L['章節標題'])
    set_ph(s, 0, title, 30); set_ph(s, 1, num, 180); footer(s)
    tb(s, 0.45, 0.7, 7.9, 2.75, [[('這一章要回答', {'bold': True, 'color': OR, 'size': 13})], [(ask, {'size': 17, 'bold': True})],
                                 [('', {'size': 6})], [('接下來你會看到', {'bold': True, 'color': OR, 'size': 13})]] + [[('・' + x, {'size': 14})] for x in see],
       fill=WHITE, margin=0.22)
    return s

def grid(s, x, y, cw, head, rows, rh, size=10.5, first_bold=True, head_h=0.4, hot=()):
    # 以文字框排成的表格：表頭橘底白字、隔列淺灰；hot 為要強調的列
    xx = x
    for j, h in enumerate(head):
        tb(s, xx, y, cw[j] - 0.04, head_h, h, size=size, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.07)
        xx += cw[j]
    for i, row in enumerate(rows):
        yy = y + head_h + 0.04 + i * rh; xx = x
        bg = TINT if i in hot else (LG if i % 2 == 0 else WHITE)
        for j, v in enumerate(row):
            first = j == 0 and first_bold
            tb(s, xx, yy, cw[j] - 0.04, rh - 0.04, v, size=size, bold=first, color=OR if first else INK,
               fill=bg, anchor=MSO_ANCHOR.MIDDLE, margin=0.07)
            xx += cw[j]

def callout(s, y, label, text, h=0.62, size=11.5):
    tb(s, 0.45, y, 0.1, h, '', fill=OR)
    tb(s, 0.55, y, 12.35, h, [[(label, {'bold': True, 'color': OR}), (text, {})]], size=size, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.16)


# ───────── 從手冊 PDF 裁切「直線變循環」圖（與 build_sdlc_deck.py 相同裁切範圍） ─────────
TMP = tempfile.mkdtemp()
subprocess.run(['pdftoppm', '-r', '220', '-f', '5', '-l', '5', '-png', '-singlefile', PDF, f'{TMP}/p5'], check=True)
im = Image.open(f'{TMP}/p5.png').convert('RGB'); w, h = im.size
im = im.crop((int(0.09 * w), int(0.335 * h), int(0.91 * w), int(0.775 * h)))
box = ImageChops.difference(im, Image.new('RGB', im.size, (255, 255, 255))).convert('L').point(lambda v: 255 if v > 12 else 0).getbbox()
FIG2 = f'{TMP}/fig2.png'
im.crop((max(box[0] - 6, 0), max(box[1] - 6, 0), min(box[2] + 6, im.size[0]), min(box[3] + 6, im.size[1]))).save(FIG2)

def pic_fit(s, path, x, y, mw, mh):
    # 依原圖比例放入，不超過 mw × mh
    iw, ih = Image.open(path).size
    pw = min(mw, mh * iw / ih); ph = pw * ih / iw
    s.shapes.add_picture(path, Inches(x), Inches(y), Inches(pw), Inches(ph))
    return pw, ph


TAGC = {'事實': INK, '建議': OR, '示意': MID, '待確認': GREY}

def tag(s, *kinds):
    # 每頁右上方標示內容性質：事實（有公開來源）、建議（我們的提案）、示意（說明用，非實測或報價）
    x = 12.9 - 0.95 * len(kinds)
    for i, k in enumerate(kinds):
        tb(s, x + i * 0.95, 1.0, 0.88, 0.3, k, size=9.5, bold=True, color=WHITE, fill=TAGC[k], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.02)

ASPECT = [('面向誰', 'Face who'), ('解決什麼問題', 'Solve what problem'), ('帶來什麼價值', 'Bring what value'), ('可擴充、可複製', 'Extendable, duplicable')]
VP = '讓保險公司用 AI 加快系統交付，同時守住法規、留下每一步紀錄，並把做法變成團隊可以重複使用的能力。'

# ───────── 1 封面 ─────────
s = p.slides.add_slide(L['標題投影片'])
set_ph(s, 0, 'AI 原生軟體開發：\n可複製的導入方案', 28)
set_ph(s, 1, 'PwC FSC1｜給保險業的提案｜2026 年 10 月', 12)

# ───────── 2 一頁看懂 ─────────
s = content('一頁看懂：我們協助保險公司把 AI 用進整個系統開發流程，交付更快，每一步都守得住、查得到')
tag(s, '建議')
tb(s, 0.45, 1.42, 12.45, 0.72, [[('核心價值主張：', {'bold': True, 'color': OR}), (VP, {})]], size=13.5, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.18)
ANS = ['有系統開發需求的台灣壽險、產險公司；決策與使用者是 IT／數位主管、業務單位主管、法遵與風控。',
       'AI 讓寫程式變快，但需求、審查、上線還是人工速度；同時擔心委外、個資、AI 責任與稽核。',
       '新功能更快上線（以試點量測）、規範自動執行、每一步有紀錄可查、團隊學會方法。',
       '共用範本、規範技能、閘門設定與交付手冊；從一個試點擴到更多團隊、情境與產業。']
for i, ((zh, en), t) in enumerate(zip(ASPECT, ANS)):
    x = 0.45 + i * 3.14
    tb(s, x, 2.3, 3.0, 0.82, [[(f'{i + 1}　{zh}', {'bold': True, 'size': 14})], [(en, {'size': 10})]], color=WHITE, fill=OR if i < 3 else INK, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
    tb(s, x, 3.17, 3.0, 2.15, t, size=12, fill=LG, margin=0.14)
    if i < 3: arrow(s, x + 3.0, 2.58, 0.14, 0.26)
callout(s, 5.55, '如何開始：', '從「AI 原生開發準備度診斷」開始，挑一個真實需求做試點，用同一套指標比較前後，數字站得住再擴大。', h=0.7, size=12)

# ───────── 3 議程 ─────────
s = content('議程：以四個面向回答三件事：為什麼需要、為什麼選我們、如何開始')
AG = [('01', '面向誰', '客戶的難題與四種角色的關心點', '4–6'),
      ('02', '解決什麼問題', '卡點與解法、運作方式、四個面向如何組成方案、為什麼選我們', '7–11'),
      ('03', '帶來什麼價值', '問題 → 解法 → 價值 → 指標、應用情境、怎麼證明', '12–15'),
      ('04', '可擴充、可複製', '服務套組、標準化與客製化、交付模式、同業做法、擴充路線', '16–21'),
      ('05', '如何開始', '分階段路線圖與下一步', '22–24'),
      ('附錄', '標示說明、來源、同業產品清單', '', '25–26')]
for i, (n, h, d, pg) in enumerate(AG):
    y = 1.5 + i * 0.78
    circle(s, 0.6, y + 0.1, 0.5, n, fill=OR if n != '附錄' else MID, size=12 if n != '附錄' else 10)
    tb(s, 1.35, y, 3.6, 0.68, h, size=16, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    tb(s, 5.0, y, 6.6, 0.68, d, size=12, color=GREY, anchor=MSO_ANCHOR.MIDDLE)
    tb(s, 11.65, y, 1.25, 0.68, pg, size=15, bold=True, color=OR, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT)
    hline(s, 1.35, y + 0.73, 12.9, y + 0.73)
note(s, '右側數字為頁碼。每頁右上方標示內容性質：事實（有公開來源）、建議（我們的提案）、示意（說明用，非實測或報價）。', y=6.4)

# ───────── 章節 01 ─────────
section('01', '面向誰\nFace who', '客戶遇到什麼難題？誰在意、各自在意什麼？', ['AI 寫程式變快，交付卻沒變快', '投資回報與法規的雙重壓力', '四種角色的關心點與我們的答案'])

# ───────── 5 痛點 ─────────
s = content('客戶的難題：AI 已經讓寫程式變快，但整體交付沒有變快，法規要求反而更多')
tag(s, '事實')
PAIN = [('瓶頸移到前後', '人工速度', '需求討論、程式審查、上線核准仍是人工速度。程式量變多之後，審查反而排得更長，整體交付時間沒有等比例縮短。', f'{PB}，第 3–4 頁'),
        ('投資看不到回報', '95%', '投資 AI 但尚未看到可衡量回報的企業比例（跨產業）。同時 57% 的保險高階主管把生成式與代理型 AI 列為 2026 年首要科技投資。', 'Bain 引述 MIT 2025；PwC 調查（Claims Journal，2025-11）'),
        ('法規要求同時到位', '2025-12-31', '保險業作業委外新辦法施行。公會 AI 自律規範（2025-04-10 備查）要求保險公司對 AI 結果負最終責任；金管會 AI 指引 2024-06 發布。', '理律法律事務所；壽險／產險公會；金管會')]
for i, (h, big, t, src) in enumerate(PAIN):
    x = 0.45 + i * 4.18
    tb(s, x, 1.45, 4.05, 0.5, h, size=14, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.08)
    tb(s, x, 2.0, 4.05, 0.95, big, size=34 if len(big) < 6 else 26, bold=True, color=OR, fill=LG, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    tb(s, x, 2.95, 4.05, 1.75, t, size=11.5, fill=LG, margin=0.14)
    tb(s, x, 4.7, 4.05, 0.55, '來源：' + src, size=9, color=GREY, fill=LG, margin=0.07)
callout(s, 5.45, '所以：', '問題不在「要不要用 AI 寫程式」，而在「怎麼讓整個開發流程跟上 AI 的速度，同時守得住法規、拿得出紀錄」。', h=0.75, size=12.5)
note(s, '完整來源與連結見附錄 A。', y=6.35)

# ───────── 6 面向誰 ─────────
s = content('面向誰：有系統開發需求的保險公司；四種角色在意的事不同，方案都要回答得了')
tag(s, '建議')
ROWS = [['IT／數位主管', '交付速度、品質、團隊產能', '五階段流程與範本，AI 參與每一步；先在一個試點量出成效'],
        ['業務單位主管\n（理賠、核保、商品）', '需求說得清楚、上線不要拖', '業務人員和 AI 一起整理需求說明，每階段交出檔案，不必反覆開會'],
        ['法遵、風控、稽核', '委外、個資、AI 責任、稽核軌跡', '規範寫進 AI 會照做的技能；關卡有指定核准人；每一步留下紀錄'],
        ['管理層', '投資回報、風險可控', '從小試點開始，用同一套指標比較前後，數字站得住才擴大']]
grid(s, 0.45, 1.45, [2.6, 3.4, 6.45], ['角色', '他們在意什麼', '我們的答案'], ROWS, 0.7, size=11.5)
tb(s, 0.45, 4.85, 12.45, 0.38, '客戶情況不同，起點不同', size=12, bold=True, color=GREY, margin=0.02)
SIT = [('已有 IT 團隊、缺方法', '→ 你自己做（工具包）'), ('想先看到成效再決定', '→ 陪你做（試點陪跑，多數客戶的起點）'), ('人力不足、時程緊', '→ 幫你做（駐點開發）')]
for i, (a, b) in enumerate(SIT):
    tb(s, 0.45 + i * 4.18, 5.25, 4.05, 0.85, [[(a, {'bold': True})], [(b, {'color': OR, 'bold': True})]], size=11.5, fill=TINT if i == 1 else LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.14)
note(s, '客戶若同時是所內審計客戶，推案前先與 Risk & Quality 確認獨立性。', y=6.3)

# ───────── 章節 02 ─────────
section('02', '解決什麼問題\nSolve what problem', '我們實際提供什麼？怎麼運作？為什麼選我們？', ['六個卡點與對應做法', '從直線到循環：AI 參與每一步，人在關鍵點決定', '四個面向如何組成完整方案', '我們的差異'])

# ───────── 8 卡點與解法 ─────────
s = content('解決什麼問題：把客戶最常遇到的六個卡點，對應到流程裡的具體做法')
tag(s, '建議')
ROWS = [['需求來回好幾輪才定稿', '業務人員和 AI 一起整理需求說明；AI 追問範圍與限制、標出缺口，不自行補需求', '需求說明範本'],
        ['規格、程式、測試對不起來', '每個階段交出一份檔案，下一階段直接接手；測試案例綁需求編號', '規格、實作計畫範本'],
        ['審查排隊、品質看人', 'AI 先做多輪審查並排出嚴重度，人專注判斷意圖與風險', '審查規則、Review Checklist'],
        ['公司規範沒人照做', '資安、個資、委外規範寫成 AI 做事時自動套用的技能', '技能（Skill）'],
        ['怕 AI 失控、稽核查不到', '沒有核准不能合併、不能上線；誰核准、何時核准都有紀錄', '自動閘門（Hook）、集中設定'],
        ['上線後同樣問題一再發生', '異常由 AI 先診斷，結果寫成新的需求說明，回到流程', '維運回饋']]
grid(s, 0.45, 1.45, [3.2, 6.45, 2.8], ['客戶卡點', '我們的做法', '用到的元件'], ROWS, 0.7, size=11.5)
note(s, f'做法依據 {PB}；五階段流程沿用 FSC1 IT 團隊的開發流程（需求分析、SA&D、開發、測試、上版／交付）。', y=6.35)

# ───────── 9 直線變循環 ─────────
s = content('我們的做法：把直線流程變成循環，AI 參與每一步，關鍵決定由客戶指定的人做')
tag(s, '事實')
pw, ph = pic_fit(s, FIG2, 0.45, 1.45, 6.6, 3.9)
tb(s, 0.45, 1.5 + ph, pw, 0.3, '圖：Anthropic 手冊圖 2「The line and the loop」', size=9, color=GREY, margin=0.02)
x0 = 0.45 + pw + 0.3
GATE = [('① 接受需求', '產品負責人'), ('② 確認規格', '產品負責人（高風險加技術主管）'), ('③ 核准計畫', '工程師（高風險找架構師）'),
        ('④ 核准合併', '程式負責人'), ('⑤ 授權上線', '發布經理')]
tb(s, x0, 1.45, 12.9 - x0, 0.42, '五個由人把關的關卡', size=12.5, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.07)
for i, (g, who) in enumerate(GATE):
    y = 1.93 + i * 0.6
    tb(s, x0, y, 1.6, 0.54, g, size=11.5, bold=True, color=OR, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.08)
    tb(s, x0 + 1.65, y, 12.9 - x0 - 1.65, 0.54, who, size=11.5, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.08)
callout(s, 5.55, '原則：', 'AI 可以一路做到正式環境的門口，但不能自己開門。寫程式的 AI 沒有核准權限，職責分離和現在一樣；上線後的異常寫成新需求，回到第一步。', h=0.78)
note(s, f'來源：{PB}，第 5、13、17、20、39、46 頁。左圖左邊為傳統直線，右邊為 AI 原生循環。', y=6.45)

# ───────── 10 四個面向如何組成方案 ─────────
s = content('四個面向環環相扣：找對客戶、解對問題，用可衡量的價值證明，再用共用資產複製擴大')
tag(s, '建議')
BOX = [(0.45, 1.45, '面向誰', '保險公司的 IT、業務、法遵與管理層', '第 4–6 頁'),
       (7.9, 1.45, '解決什麼問題', '交付卡在人工步驟，又要守住法規與稽核', '第 7–11 頁'),
       (7.9, 3.95, '帶來什麼價值', '更快上線、規範自動執行、紀錄可查、團隊學會方法', '第 12–15 頁'),
       (0.45, 3.95, '可擴充、可複製', '標準資產＋客製設定；擴到更多團隊、情境與產業', '第 16–21 頁')]
for i, (x, y, h, t, pg) in enumerate(BOX):
    tb(s, x, y, 5.0, 0.5, f'{i + 1}　{h}　{ASPECT[i][1]}', size=13, bold=True, color=WHITE, fill=OR if i < 3 else INK, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
    tb(s, x, y + 0.5, 5.0, 1.5, [t, [(pg, {'color': GREY, 'size': 10})]], size=12.5, fill=LG, margin=0.16)
tb(s, 5.65, 2.55, 2.05, 2.3, [[('核心價值主張', {'bold': True, 'color': OR, 'size': 11})], 'AI 加快交付\n守住法規\n留下紀錄\n能力留在客戶'], size=12, bold=True, fill=TINT, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.08)
arrow(s, 5.6, 1.7, 2.15, 0.32)
arrow(s, 10.25, 3.5, 0.32, 0.42, MSO_SHAPE.DOWN_ARROW)
arrow(s, 5.6, 5.3, 2.15, 0.32, MSO_SHAPE.LEFT_ARROW)
arrow(s, 2.8, 3.5, 0.32, 0.42, MSO_SHAPE.UP_ARROW)
callout(s, 6.05, '循環：', '每做完一個客戶，範本、規範技能與指標基準就回到共用資產庫，下一個客戶導入更快，也能服務更多類型的客戶。', h=0.55, size=11)

# ───────── 11 為什麼選我們 ─────────
s = content('為什麼選我們：台灣保險法規、保險業務與在地系統、可稽核的關卡，三者同時具備')
tag(s, '建議', '待確認')
DIFF = [('台灣保險法規內建', '委外辦法、公會 AI 與雲端自律規範、金管會 AI 指引，直接寫進技能與檢查清單；診斷時一併做法規檢查。'),
        ('懂保險業務與在地系統', 'FSC1 熟悉理賠、核保、核心系統與 IFRS 17，試點直接用客戶真正在痛的流程，不必從頭解釋業務。'),
        ('可稽核的人工關卡', '五個關卡誰核准、什麼時候核准，全部查得到；寫程式的 AI 不能核准自己的程式。')]
for i, (h, t) in enumerate(DIFF):
    x = 0.45 + i * 4.18
    circle(s, x, 1.5, 0.55, str(i + 1), size=14)
    tb(s, x + 0.65, 1.5, 3.4, 0.55, h, size=14, bold=True, color=OR, anchor=MSO_ANCHOR.MIDDLE)
    tb(s, x, 2.15, 4.05, 1.75, t, size=12, fill=LG, margin=0.16)
tb(s, 0.45, 4.1, 12.45, 1.6, [[('我們做得到的依據（待 IT 團隊填入）', {'bold': True, 'color': OR, 'size': 12.5})],
                              '・已經做出來的系統或工具：名稱、給誰用、解決什麼問題',
                              '・用到的做法：需求說明、規格、計畫範本，技能，自動閘門',
                              '・量到的成效與可以現場示範的內容'],
   size=12, fill=WHITE, line=OR, margin=0.18)
note(s, '虛線框內容須由 IT 團隊提供實際成果後才對外使用，不以假設數字代替。', y=5.85)

# ───────── 章節 03 ─────────
section('03', '帶來什麼價值\nBring what value', '每個問題解決之後，客戶得到什麼？怎麼證明？', ['問題 → 解法 → 商業價值 → 衡量指標', '應用情境：理賠進度自助查詢', '試點前後怎麼比較'])

# ───────── 13 價值鏈 ─────────
s = content('帶來什麼價值：每個客戶問題都對應一個可以量測的商業成果')
tag(s, '建議')
ROWS = [['需求定稿慢', '需求說明範本＋AI 追問', '新功能更快上線', '從第一次討論到需求定稿的時間'],
        ['審查排隊', 'AI 先審、人判斷', '工程師把時間花在需要判斷的工作', '第一次審查的等待時間、每個改動的返工次數'],
        ['品質不穩', 'AI 自行測試，案例綁需求', '上線事故減少', '上線失敗率、上線後才發現的缺陷'],
        ['法規與稽核壓力', '技能＋自動閘門＋紀錄', '主管機關或稽核詢問時拿得出佐證', '有完整核准紀錄的改動比例'],
        ['知識只在個人腦中', '團隊說明檔、範本與教學', '人員異動不影響交付，新人上手更快', '新成員獨立交付第一個改動的時間']]
cw = [2.4, 3.0, 3.4, 3.65]
grid(s, 0.45, 1.45, cw, ['客戶問題', '→ 解決方式', '→ 商業價值', '→ 衡量指標'], ROWS, 0.74, size=11.5)
callout(s, 5.75, '說明：', '價值以試點前後用同一套指標比較，不預先承諾數字。手冊所述預期：從想法到需求定稿可由數週縮短到數小時（非實測）。', h=0.7, size=11.5)

# ───────── 14 應用情境 ─────────
s = content('應用情境：「理賠進度自助查詢」從客服痛點到上線，人只在關鍵點做決定')
tag(s, '事實')
STEP = [('① 提出問題', '理賠同仁和 AI 整理需求', '需求說明', False), ('② 接受需求', '產品負責人決定', '接受紀錄', True),
        ('③ 起草規格', 'AI 自動套用公司規範', '規格書', False), ('④ 實作計畫', '工程師核准前 AI 不能改程式', '實作計畫', True),
        ('⑤ 程式與測試', 'AI 自己跑測試，全過才交人', '測試結果', False), ('⑥ 審查上線', '程式負責人、發布經理核准', '核准紀錄', True),
        ('⑦ 上線後監看', '異常由 AI 先診斷', '事件紀錄', False)]
W = 1.72
for i, (h, t, rec, human) in enumerate(STEP):
    x = 0.45 + i * (W + 0.07)
    tb(s, x, 1.5, W, 0.6, h, size=12, bold=True, color=WHITE, fill=OR if human else INK, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, margin=0.04)
    tb(s, x, 2.15, W, 1.3, t, size=11, fill=TINT if human else LG, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, margin=0.1)
    tb(s, x, 3.5, W, 0.5, rec, size=10.5, bold=True, color=OR, fill=WHITE, line=LINE, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, margin=0.04)
tb(s, 0.45, 4.1, 12.45, 0.3, '橘色＝人做決定的關卡；黑色＝AI 先做、人檢查。框線內為每一步留下的紀錄。', size=10, color=GREY, margin=0.02)
tb(s, 0.45, 4.55, 6.1, 1.45, [[('情境', {'bold': True, 'color': OR, 'size': 12})], '客戶常打電話問理賠進度。理賠部門同仁直接和 AI 說明問題，AI 追問範圍、使用者與限制，一路走到上線，上線後持續監看。'], size=11.5, fill=LG, margin=0.14)
tb(s, 6.8, 4.55, 6.1, 1.45, [[('對客戶的意義', {'bold': True, 'color': OR, 'size': 12})], '業務需求不必經過多輪會議才定稿；每一步都有檔案和核准紀錄，稽核時可以直接調出來。'], size=11.5, fill=TINT, margin=0.14)
note(s, f'此為 {PB} 的範例（第 12–20、37–52 頁），非客戶案例；手冊中「客服約三分之一通話時間在回答進度」為示例數字。', y=6.15)

# ───────── 15 怎麼證明 ─────────
s = content('怎麼證明價值：試點前先量基準，試點後用同一套指標比較，數字站得住才擴大')
tag(s, '示意')
ROWS = [['需求分析', '從第一次討論到需求定稿的時間', '試點前量測', '縮短'],
        ['SA&D', '開發開始後還要改規格的次數', '試點前量測', '減少'],
        ['開發', '第一次實作就能合併的比例', '試點前量測', '提高'],
        ['測試', 'AI 改動第一次就通過 CI 的比例', '試點前量測', '提高'],
        ['上版／交付', '第一次審查的等待時間、上線失敗率', '試點前量測', '縮短、降低'],
        ['維運', '同類事件重複發生的次數', '試點前量測', '減少']]
grid(s, 0.45, 1.45, [1.6, 4.6, 1.8, 1.9], ['階段', '指標', '基準', '期望方向（示意）'], ROWS, 0.6, size=11.5)
tb(s, 10.55, 1.45, 2.35, 4.1, [[('原則', {'bold': True, 'color': OR, 'size': 12.5})],
                               '數據大多已在客戶的版本紀錄與 CI 系統裡，不需要另外填表。', [('', {'size': 6})],
                               '對外只引用自己試點量到的數字，不引用手冊或他人的預期值。'], size=11.5, fill=TINT, margin=0.14)
note(s, f'指標定義依 {PB} 各階段「How to measure it」，可搭配 DORA 指標（部署頻率、交付時間、變更失敗率、復原時間）。期望方向為示意，非承諾。', y=5.75)

# ───────── 章節 04 ─────────
section('04', '可擴充、可複製\nExtendable, duplicable', '怎麼把框架變成可以重複銷售、重複交付的方案？', ['一個入門診斷加三種方案', '哪些標準化、哪些客製化', '共用資產與交付手冊', '同業怎麼做', '擴充路線'])

# ───────── 17 服務套組 ─────────
s = content('產品化：一個入門診斷加三種方案，客戶依需要的人力參與程度選擇')
tag(s, '建議', '示意')
ROWS = [['入門｜準備度診斷', '想評估是否導入、從哪裡開始', '瓶頸地圖、法規檢查、試點建議、導入藍圖', '訪談、流程盤點、工作坊', '4–6 週', '固定價'],
        ['你自己做｜工具包', '已有 IT 團隊、缺方法', '範本、保險業規範技能、閘門設定、Review Checklist、交付手冊', '授權＋導入說明', '2–4 週導入', '授權或年費'],
        ['陪你做｜試點陪跑＋教學（主推）', '想先看到成效', '完成的試點、團隊說明檔、審查規則、教材、成效報告', '與客戶團隊一起走完五階段', '6–8 週', '每個試點固定價，可加里程碑'],
        ['幫你做｜駐點開發', '人力不足、時程緊', '程式、文件與完整紀錄；期末交接工具包', 'FSC1 人力駐點', '依範圍', '人月或按交付項目']]
grid(s, 0.45, 1.45, [2.3, 1.85, 3.25, 2.05, 1.25, 1.75], ['方案', '目標客群', '交付成果', '執行方式', '期程（示意）', '收費模式（建議）'], ROWS, 0.95, size=10.5, hot=(2,))
callout(s, 5.75, '計價原則：', '依產品化研究的建議，按「每次診斷」「每個試點」等任務或成果計價，而非按工時；金額依客戶範圍另行報價。', h=0.62, size=11)
note(s, f'期程為示意，以診斷結果為準。三種方案概念出自 {SMR}（Do it yourself／with me／for me）。', y=6.45)

# ───────── 18 標準化 vs 客製化 ─────────
s = content('可複製的關鍵：流程、範本與工具標準化，客製只集中在客戶自己的規範、系統與核准人')
tag(s, '建議')
STD = ['五階段流程與五個人工關卡', '需求說明、規格、實作計畫範本', 'Review Checklist 九項', '保險業共通規範技能（個資、委外、AI 責任）', '自動閘門設定樣板', '教材與交付手冊', '成效指標定義']
CUS = ['客戶自己的資安、命名與品牌規範', '核心系統與既有工具串接（Jira、變更管理）', '資料分級：哪些可以給 AI 讀', '各關卡的核准人', '試點題目與指標基準', '雲端與 AI 工具的委外安排']
tb(s, 0.45, 1.45, 6.1, 0.5, '標準化：每個客戶都一樣', size=14, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
tb(s, 0.45, 2.0, 6.1, 3.5, [[('✓　' + t, {})] for t in STD], size=12.5, fill=TINT, margin=0.2)
tb(s, 6.8, 1.45, 6.1, 0.5, '客製化：依客戶調整', size=14, bold=True, color=WHITE, fill=INK, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
tb(s, 6.8, 2.0, 6.1, 3.5, [[('●　' + t, {})] for t in CUS], size=12.5, fill=LG, margin=0.2)
callout(s, 5.7, '好處：', '標準化的部分讓交付快、品質一致；客製化的部分讓方案符合每家公司的規範與系統。客製內容做完，經整理後回到共用資產庫。', h=0.7, size=11.5)

# ───────── 19 交付模式 ─────────
s = content('可複製的交付模式：共用資產庫加標準交付手冊，每做完一個客戶，資產就更完整')
tag(s, '建議')
tb(s, 0.45, 1.45, 1.6, 4.3, '交付手冊\n\n診斷\n↓\n試點\n↓\n擴大\n\n每一步的做法、檔案與檢查點', size=11.5, bold=True, color=WHITE, fill=INK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
LAY = [('客戶情境', ['理賠', '核保', '商品', '核心系統周邊', 'IFRS 17'], LG, INK),
       ('五階段模組', ['需求分析', 'SA&D', '開發', '測試', '上版／交付'], TINT, OR),
       ('共用資產庫', ['範本庫', '規範技能庫', '閘門設定', 'Checklist', '教材'], OR, WHITE)]
for r, (h, items, bg, fc) in enumerate(LAY):
    y = 1.45 + r * 1.45
    tb(s, 2.2, y, 1.7, 1.3, h, size=12.5, bold=True, color=OR, fill=WHITE, line=LINE, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    for j, it in enumerate(items):
        tb(s, 4.0 + j * 1.38, y, 1.32, 1.3, it, size=11.5, bold=True, color=fc, fill=bg, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, margin=0.04)
tb(s, 11.05, 1.45, 1.85, 4.3, [[('回饋', {'bold': True, 'color': OR, 'size': 12.5})], '每個案子結束，至少萃取一個可重用的範本或技能，回到共用資產庫。', [('', {'size': 6})], '下一個客戶從更完整的資產開始。'], size=11, fill=LG, margin=0.12)
arrow(s, 10.95, 3.4, 0.1, 0.3, MSO_SHAPE.LEFT_ARROW)
note(s, '上層隨客戶與情境替換；中層流程固定；下層資產跨客戶共用並持續累積。', y=5.95)

# ───────── 20 同業 ─────────
s = content('同業怎麼做：保險顧問業產品化已是常態，我們借鏡加速器模式，以在地法規與可稽核關卡區隔')
tag(s, '事實', '建議')
ROWS = [['軟體／數據產品\n（你自己做）', '國際保險顧問的定價與核保軟體；策略顧問的保險市場資料庫訂閱', '用自有數據做成訂閱產品', '長期方向：累積的規範技能庫與指標基準'],
        ['加速器＋顧問導入\n（陪你做）', '會計師事務所的保險業代理型 AI 藍圖與預建代理；IFRS 17 加速器（業者宣稱 90 天內產出）', '預建資產加明確期程，降低客戶疑慮', '主推方案：固定範圍與期程的試點'],
        ['核心系統＋AI 工具組', '系統整合商的核心系統 AI 工具組（業者宣稱 300 個以上保險加速器）', '鎖定核心系統遷移情境', '聚焦台灣在地核心系統與法規'],
        ['方法論與藍圖\n（框架）', '策略顧問的 AI-First 藍圖、代理型 AI 核心現代化方法', '清楚的成熟度階段', '五階段＋五關卡，並直接落地到試點']]
grid(s, 0.45, 1.45, [2.3, 4.45, 2.75, 2.95], ['產品化型態', '同業做法（代表例）', '值得借鏡', '我們的做法'], ROWS, 0.88, size=10.5, hot=(1,))
callout(s, 5.55, '差異化：', '台灣保險法規內建、懂保險業務與在地系統、五個可稽核的人工關卡。', h=0.55, size=12)
note(s, '具體公司、產品與來源連結見附錄 B。資料為業者公開資訊，未經獨立驗證；「宣稱」數字為業者自述。', y=6.2)

# ───────── 21 擴充路線 ─────────
s = content('可擴充：同一套資產，依序延伸到更多情境、產業與功能')
tag(s, '建議')
LANE = [('情境', ['核心系統周邊開發', 'IFRS 17／TW-ICS 系統、AI 治理（AI 清冊與模型驗證）', '代理型 AI 核保與理賠']),
        ('產業', ['壽險、產險', '銀行、證券（金融業共通法規）', '其他受監理產業']),
        ('功能', ['開發流程五階段', '測試自動化、維運閉環', '跨團隊規模化、CI/CD 自動串接'])]
for j, h in enumerate(['現在', '下一步', '未來']):
    tb(s, 2.3 + j * 3.55, 1.45, 3.45, 0.45, h, size=13, bold=True, color=WHITE, fill=[OR, OR2, MID][j], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
for i, (h, items) in enumerate(LANE):
    y = 2.0 + i * 1.15
    tb(s, 0.45, y, 1.75, 1.05, h, size=14, bold=True, color=OR, fill=LG, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    for j, it in enumerate(items):
        tb(s, 2.3 + j * 3.55, y, 3.45, 1.05, it, size=12, fill=TINT if j == 0 else LG, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, margin=0.12)
callout(s, 5.6, '前提：', '每一步擴充都以前一步的試點成效為依據；共用資產庫讓新情境不必從零開始，只需補上該情境的範本與規範。', h=0.7)

# ───────── 章節 05 ─────────
section('05', '如何開始', '從評估到落地，第一步做什麼？', ['四階段導入路線圖', '下一步與我們需要客戶提供的資料'])

# ───────── 23 路線圖 ─────────
s = content('如何開始：從準備度診斷開始，用一個試點證明價值，再逐步擴大與規模化')
tag(s, '示意')
PH = [('0　準備度診斷', '4–6 週', '盤點現況、法規檢查、挑選試點', '瓶頸地圖、導入藍圖', '選定試點與指標'),
      ('1　試點', '6–8 週', '用一個真實需求走完五階段', '完成的試點、成效報告', '指標比基準改善'),
      ('2　擴大', '1–2 季', '複製到更多團隊，加入客戶規範技能', '擴大路線圖、技能庫', '管理層核准擴大'),
      ('3　規模化與閉環', '持續', '接上 CI/CD 與維運閉環，必要時駐點', '持續運作的流程', '—')]
for i, (h, d, do, out, gate) in enumerate(PH):
    x = 0.45 + i * 3.14
    tb(s, x, 1.45, 3.0, 0.75, [[(h, {'bold': True, 'size': 14})], [(d, {'size': 10.5})]], color=WHITE, fill=[OR, OR, OR2, MID][i], anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
    if i < 3: arrow(s, x + 3.0, 1.68, 0.14, 0.3)
    tb(s, x, 2.25, 3.0, 1.15, [[('做什麼', {'bold': True, 'color': OR, 'size': 10.5})], do], size=11.5, fill=LG, margin=0.12)
    tb(s, x, 3.45, 3.0, 1.1, [[('交付', {'bold': True, 'color': OR, 'size': 10.5})], out], size=11.5, fill=LG, margin=0.12)
    tb(s, x, 4.6, 3.0, 0.95, [[('進下一階段的條件', {'bold': True, 'color': OR, 'size': 10.5})], gate], size=11.5, fill=TINT, margin=0.12)
note(s, '期程為示意，實際依客戶範圍與診斷結果調整。每一階段結束由客戶決定是否進入下一階段。', y=5.8)

# ───────── 24 下一步 ─────────
s = content('下一步：三件事就能啟動準備度診斷')
tag(s, '建議')
NX = [('對齊工作坊', '對齊目標，從客戶的需求清單挑 1–2 個候選試點。'),
      ('前置確認', '資料分級、雲端與 AI 工具的委外安排；若客戶同時是所內審計客戶，先確認獨立性。'),
      ('啟動診斷', '確認診斷範圍與雙方窗口，開始盤點現況。')]
for i, (h, t) in enumerate(NX):
    y = 1.5 + i * 1.3
    circle(s, 0.45, y + 0.25, 0.65, str(i + 1), size=18)
    tb(s, 1.3, y, 6.4, 1.15, [[(h, {'bold': True, 'size': 14})], t], size=12, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.16)
tb(s, 8.0, 1.5, 4.9, 3.75, [[('我們需要客戶提供', {'bold': True, 'color': OR, 'size': 13})],
                            '・IT 主管與業務單位窗口', '・現有開發流程與工具（Jira、CI、變更管理）', '・候選需求清單', '・公司資安、個資與法遵規範'],
   size=12.5, fill=TINT, margin=0.2)
callout(s, 5.55, '一句話：', 'AI 幫你把每一步做得更快，每一步都留下紀錄；要不要做、能不能上線，永遠由你的人決定。', h=0.7, size=12.5)

# ───────── 25 附錄 A ─────────
s = content('附錄 A｜標示說明與來源')
tag(s, '事實')
tb(s, 0.45, 1.45, 12.45, 0.55, [[('事實', {'bold': True}), ('＝有公開來源　', {}), ('建議', {'bold': True, 'color': OR}), ('＝我們的提案　', {}), ('示意', {'bold': True, 'color': GREY}), ('＝說明用，非實測數據或報價　', {}), ('待確認', {'bold': True, 'color': GREY}), ('＝需內部提供', {})]], size=12, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.16)
SRCS = [('Louis Claxton（Anthropic），The AI-Native SDLC Playbook，2026', None),
        ('Bain，Reimagining Insurance for the AI Era（引述 MIT 2025：95%）', 'https://www.bain.com/ko/results/reimagining-insurance-for-the-ai-era'),
        ('Claims Journal，PwC 保險高階主管調查（57%），2025-11-17', 'https://amp.claimsjournal.com/news/national/2025/11/17/334094.htm'),
        ('理律法律事務所，保險業作業委外新辦法（2025-12-31 施行）', 'https://www.leeandli.com/TW/Newsletters/7525.htm'),
        ('壽險／產險公會，保險業運用人工智慧系統自律規範（2025-04-10 備查）', 'https://law.lia-roc.org.tw/Law/GetFile/0000394231'),
        ('壽險／產險公會，保險業作業委外使用雲端服務自律規範', 'https://law.lia-roc.org.tw/Law/GetFile/0000419987'),
        ('金管會，金融業運用人工智慧（AI）指引，2024-06', 'https://www.fsc.gov.tw/websitedowndoc?file=chfsc/202408231741530.pdf'),
        ('Mohanbir Sawhney，How to Turn Professional Services Into Products，MIT Sloan Management Review，2024', None)]
for i, (lab, url) in enumerate(SRCS):
    seg = (lab, {'link': url}) if url else (lab, {})
    tb(s, 0.45, 2.15 + i * 0.5, 12.45, 0.44, [[(f'{i + 1}. ', {'bold': True, 'color': OR}), seg]], size=11, fill=LG if i % 2 == 0 else WHITE, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
note(s, '連結可點擊。金管會 AI 指引連結取自研究連結庫，使用前請點開確認。', y=6.2)

# ───────── 26 附錄 B ─────────
s = content('附錄 B｜同業具體產品與來源（業者公開資訊）')
tag(s, '事實')
ROWS = [('Deloitte', 'Agentic AI Blueprint for Insurance', '加速器＋導入', 'https://www.deloitte.com/global/en/what-we-do/capabilities/agentic-ai-blueprint-for-insurance.html'),
        ('Deloitte', 'IFRS 17 Regulatory Assist（與 SAS 合作）', '加速器', 'https://sas.com/nl_nl/news/press-releases/2022/may/deloitte-partners-with-sas-to-power-ifrs-17-accelerator-solution.html'),
        ('KPMG', 'IFRS 17 Express, powered by SAS', '加速器', 'https://sas.com/en_ca/news/press-releases/2022/september/pvi-insurance-deploys-kpmg-ifrs-17-express.html'),
        ('Capgemini', 'Guidewire 生成式 AI／代理工具組', '核心系統＋AI', 'https://www.capgemini.com/wp-content/uploads/2026/10/GW-datasheet_Final-version.pdf'),
        ('EY', 'EY Nexus for Insurance', '平台', 'https://www.ey.com/en_us/insurance/nexus'),
        ('WTW', 'Radar 5（保險定價與核保軟體）', '軟體產品', 'https://www.wtwco.com/en-gh/news/2025/10/wtw-launch-of-radar-5-with-gen-ai-capability-marks-a-major-milestone-in-insurance-pricing'),
        ('McKinsey', 'Global Insurance Pools（訂閱資料庫）', '數據產品', 'https://www.mckinsey.com/industries/financial-services/how-we-help-clients/insurance/global-insurance-pools'),
        ('McKinsey', '代理型 AI 核心現代化', '框架', 'https://www.mckinsey.com/industries/financial-services/our-insights/can-agentic-ai-finally-modernize-core-technologies-in-insurance'),
        ('BCG', 'AI-First 產險公司藍圖', '框架', 'https://www.bcg.com/assets/2026/executive-perspectives-ai-first-companies-win-the-future-property-casualty-insurance.pdf')]
xx = [0.45, 2.0, 6.0, 7.75]
for j, h in enumerate(['公司', '產品／方案', '型態', '來源（可點擊）']):
    tb(s, xx[j], 1.45, (xx[j + 1] if j < 3 else 12.9) - xx[j] - 0.04, 0.36, h, size=10.5, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.07)
for i, (c, prod, kind, url) in enumerate(ROWS):
    y = 1.85 + i * 0.47; bg = LG if i % 2 == 0 else WHITE
    tb(s, xx[0], y, 1.51, 0.43, c, size=10.5, bold=True, color=OR, fill=bg, anchor=MSO_ANCHOR.MIDDLE, margin=0.07)
    tb(s, xx[1], y, 3.96, 0.43, prod, size=10.5, fill=bg, anchor=MSO_ANCHOR.MIDDLE, margin=0.07)
    tb(s, xx[2], y, 1.71, 0.43, kind, size=10.5, fill=bg, anchor=MSO_ANCHOR.MIDDLE, margin=0.07)
    dom = url.split('/')[2]
    tb(s, xx[3], y, 5.11, 0.43, [[(dom + ' 原文', {'link': url})]], size=10.5, fill=bg, anchor=MSO_ANCHOR.MIDDLE, margin=0.07)
note(s, '對外提案前請確認是否保留本頁公司名稱。資料為業者新聞稿與產品頁，規模數字為業者自述，未經獨立驗證。', y=6.2)

p.save(OUT)
print(OUT, len(p.slides), 'slides')
