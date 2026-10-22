# AI 原生軟體開發流程（AI-native SDLC）白話說明簡報：給管理層與客戶的說明版（以所內簡報範本建立）
# 內容整理自 Anthropic《The AI-Native SDLC Playbook》，把工程用語翻成管理層聽得懂的說法
# 用法：python3 build_sdlc_deck.py <範本.pptx> <輸出.pptx> <手冊.pdf>
#   手冊的圖在建置時從 PDF 裁切，不放進版本控制
import sys, copy, glob, subprocess, tempfile
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
SYMBOLS = set('●◐○✓✗→①②③④⑤⑥⑦⚠↑σ')
FOOT = '對外說明版｜內容整理自 Anthropic《The AI-Native SDLC Playbook》｜時間縮短幅度為該手冊所述預期，實際以試點量測為準'
SRC = 'Anthropic《The AI-Native SDLC Playbook》'

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


# ───────── 從手冊 PDF 裁切圖片 ─────────
# (頁碼, 左, 上, 右, 下)，以頁面寬高比例表示；圖 3 是 PDF 內嵌圖片，直接抽出
CROPS = {'fig1': (4, 0.09, 0.385, 0.91, 0.62), 'fig2': (5, 0.09, 0.335, 0.91, 0.775),
         'intent': (13, 0.09, 0.125, 0.91, 0.50), 'plan': (20, 0.09, 0.345, 0.91, 0.675),
         'fig4': (52, 0.09, 0.625, 0.91, 0.975)}
TMP = tempfile.mkdtemp()

def trim(im):
    box = ImageChops.difference(im, Image.new('RGB', im.size, (255, 255, 255))).convert('L').point(lambda v: 255 if v > 12 else 0).getbbox()
    return im.crop((max(box[0] - 6, 0), max(box[1] - 6, 0), min(box[2] + 6, im.size[0]), min(box[3] + 6, im.size[1])))

FIG = {}
for key, (n, x0, y0, x1, y1) in CROPS.items():
    subprocess.run(['pdftoppm', '-r', '220', '-f', str(n), '-l', str(n), '-png', '-singlefile', PDF, f'{TMP}/p{n}'], check=True)
    im = Image.open(f'{TMP}/p{n}.png').convert('RGB'); w, h = im.size
    FIG[key] = f'{TMP}/{key}.png'
    trim(im.crop((int(x0 * w), int(y0 * h), int(x1 * w), int(y1 * h)))).save(FIG[key])
subprocess.run(['pdfimages', '-f', '9', '-l', '9', '-png', PDF, f'{TMP}/emb'], check=True)
FIG['fig3'] = sorted(glob.glob(f'{TMP}/emb-*.png'))[0]

def pic(s, key, x, y, w=None, h=None):
    # 依原圖比例放入圖片，只給寬或高其中一個
    iw, ih = Image.open(FIG[key]).size
    if w is None: w = h * iw / ih
    if h is None: h = w * ih / iw
    s.shapes.add_picture(FIG[key], Inches(x), Inches(y), Inches(w), Inches(h))
    return w, h

def caption(s, x, y, w, text):
    tb(s, x, y, w, 0.3, text, size=9, color=GREY, margin=0.02)

# ───────── 1 封面 ─────────
s = p.slides.add_slide(L['標題投影片'])
set_ph(s, 0, 'AI 原生軟體開發流程：\n實際怎麼運作', 28)
set_ph(s, 1, 'FSC1｜給管理層與客戶的白話說明｜2026 年 10 月', 12)

# ───────── 2 一頁看懂 ─────────
s = content('一頁看懂：寫程式變快之後，整個開發流程也要跟著改，關鍵決定仍由人來做')
KEY = [('現況', 'AI 寫程式的速度比一年前快很多，但需求、審查、上線還是人工速度，整體交付時間沒有等比例縮短。'),
       ('做法', '把「規劃、設計、開發、測試、上線、維運」串成一個循環。每個階段結束都留下一份檔案，下一個階段從這份檔案接手。'),
       ('把關', 'AI 負責起草、檢查與執行。接受需求、核准計畫、核准上線這類需要判斷的決定，一律由指定的人做，而且都有紀錄。')]
for i, (h, t) in enumerate(KEY):
    y = 1.45 + i * 1.32
    circle(s, 0.45, y + 0.2, 0.62, str(i + 1), size=16)
    tb(s, 1.2, y, 1.1, 1.05, h, size=15, bold=True, color=OR, anchor=MSO_ANCHOR.MIDDLE)
    tb(s, 2.3, y, 4.75, 1.05, t, size=12, anchor=MSO_ANCHOR.MIDDLE)
    if i < 2: hline(s, 1.2, y + 1.18, 7.05, y + 1.18)
tb(s, 7.45, 1.45, 5.45, 0.48, '對客戶的三個好處', size=13, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.06)
GAIN = [('更快', '手冊預期，從想法到可執行的需求，可由數週縮短到數小時。'),
        ('查得到', '誰提出、AI 產出什麼、誰核准，全部寫在版本紀錄裡，就是現成的稽核軌跡。'),
        ('守得住', '公司規範寫成 AI 會照做的規則；正式上線前一定要有人授權。')]
for i, (h, t) in enumerate(GAIN):
    y = 2.0 + i * 0.98
    tb(s, 7.45, y, 1.2, 0.9, h, size=13, bold=True, color=OR, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    tb(s, 8.7, y, 4.2, 0.9, t, size=11.5, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
tb(s, 0.45, 5.6, 0.1, 0.9, '', fill=OR)
tb(s, 0.55, 5.6, 12.35, 0.9, [[('給客戶的一句話：', {'bold': True, 'color': OR}),
                               ('「AI 幫你把每一步做得更快，每一步也都留下紀錄；要不要上線，永遠由你的人決定。」', {})]],
   size=14, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.18)
note(s, f'來源：{SRC}，第 3–7 頁。')

# ───────── 3 為什麼要改（手冊圖 1） ─────────
s = content('為什麼要改：AI 把「寫程式」縮短到幾小時，等待時間移到了前後的人工步驟')
pic(s, 'fig1', 0.45, 1.45, w=7.3)
caption(s, 0.45, 4.3, 7.3, f'圖：{SRC} 圖 1「The bottleneck moved」')
tb(s, 8.0, 1.45, 4.9, 0.4, '這張圖怎麼看', size=13, bold=True, color=OR)
tb(s, 8.0, 1.88, 4.9, 2.4, [[('上排「Before agents」：', {'bold': True})], '過去每個階段都靠人工，其中開發（橘色 Build）最花時間。',
                            [('下排「After agents」：', {'bold': True})], 'AI 讓開發縮成一小段，但規劃、設計、測試、上線、維運的長度都沒變。右邊虛線框是「可以拿回來的時間」，要靠流程一起改才拿得到。'],
   size=11, fill=LG, margin=0.14)
tb(s, 8.0, 4.3, 4.9, 0.4, '圖中英文：Plan 規劃｜Design 設計｜Build 開發｜Test 測試｜Deploy 上線｜Maintain 維運', size=8.5, color=GREY, margin=0.02)
CON = [('瓶頸移位', '需求、審查、上線還是人工速度，整體時間被這些步驟卡住。'),
       ('控制跟不上', '逐行人工審查，在人寫程式的年代合理；AI 寫出大部分程式後就追不上。'),
       ('治理成本上升', '例外事項還是要等每週或每月的會議，AI 做得越快，排隊越長。')]
for i, (h, t_) in enumerate(CON):
    x = 0.45 + i * 4.18
    tb(s, x, 4.72, 4.0, 1.05, [[(f'{i + 1}　{h}', {'bold': True, 'size': 12.5, 'color': OR})], t_], size=11, fill=LG, margin=0.12)
tb(s, 0.45, 5.88, 12.45, 0.7, [[('資安審查的例子：', {'bold': True, 'color': OR}),
                                ('資安人力是依人寫程式的產量配置。程式量倍增後，不是審查排隊變長，就是沒審完就上線。受監理的機構兩者都不能接受，所以檢查也必須跟上 AI 的速度。', {})]],
   size=11, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.14)
note(s, f'來源：{SRC}，第 3–4 頁。')

# ───────── 4 直線變循環（手冊圖 2） ─────────
s = content('從「一條直線」變成「一個循環」：每一步自動交棒，人在循環上方發起、指揮與把關')
pw, ph = pic(s, 'fig2', 0.45, 1.45, h=4.1)
caption(s, 0.45, 5.58, pw, f'圖：{SRC} 圖 2「The line and the loop」')
x0 = 0.45 + pw + 0.35; ww = 12.9 - x0
tb(s, x0, 1.45, ww, 0.4, '這張圖怎麼看', size=13, bold=True, color=OR)
tb(s, x0, 1.88, ww, 3.65, [[('左邊「Traditional — the line」傳統的直線：', {'bold': True})],
                           '六個階段由不同角色依序負責，靠文件、工單、簽核交接。只要退回一次，就等於重新開始一個版本週期。',
                           [('右邊「AI-native — the loop」AI 原生的循環：', {'bold': True})],
                           'AI（圖中的 Claude）在中間參與每一個階段。上一步核准的檔案直接觸發下一步，手冊的說法是「以小時計，不是以週計」。',
                           [('人在循環上方：', {'bold': True})], '人負責發起、指揮與治理，專心看 AI 標出來的重點，不必每一步從頭做起。'],
   size=11.5, fill=LG, margin=0.14)
tb(s, 0.45, 5.92, 12.45, 0.65, [[('白話：', {'bold': True, 'color': OR}),
                                 ('以前是「做完一段、交給下一組人、等簽核」；現在每一段做完就把檔案交出去，下一段立刻接手，人只在關卡上做決定。', {})]],
   size=12, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.14)
note(s, f'來源：{SRC}，第 5、8–9 頁。')

# ───────── 5 六階段對照 ─────────
s = content('六個階段白話對照：做法都變了，但每一步都留下一份下一步能接手的紀錄')
ROWS = [['規劃', '需求要經過多次工作坊與簽核，再由專人寫成文件', '提出想法的人直接和 AI 討論，整理成一份需求說明檔', 'intent.md\n需求說明', '產品負責人'],
        ['設計', '分析師寫需求，設計師再轉成設計，兩組人分開做', 'AI 依需求說明一次產出需求與設計規格，套用資安、法遵、品牌規範，並標出疑慮', 'spec.md\n規格書', '產品負責人（高風險找技術主管）'],
        ['開發', '工程師讀完設計就寫程式，做法只在自己腦中', 'AI 先寫實作計畫，工程師修正、核准後才動手；團隊慣例寫成 AI 每次都會讀的說明檔', 'plan.md 實作計畫\n程式與測試', '工程師（高風險找架構師）'],
        ['測試', '階段結束才由 QA 把關，問題很晚才發現', 'AI 每改一次就自己跑測試，通過才交給人；AI 的設定改了，也要重跑評測', '測試結果\n建置紀錄', '程式負責人'],
        ['上線', '每一行程式都由人審，品質看審查者忙不忙', 'AI 先做多輪審查並排出嚴重度，人專注判斷意圖與風險；正式上線一定要人授權', 'PR 審查紀錄\n核准紀錄', '程式負責人、發布經理'],
        ['維運', '人盯著系統，事件常被擱置', '監控程式發現異常才請 AI 診斷，結果寫成新的需求說明，重新進入循環', '事件紀錄\n新的 intent.md', '服務負責人']]
grid(s, 0.45, 1.4, [1.0, 3.05, 4.35, 2.15, 1.9], ['階段', '以前怎麼做', '現在怎麼做（AI 原生）', '留下的紀錄', '誰拍板'], ROWS, 0.78, size=10.5)
note(s, f'來源：{SRC}，第 6 頁「The shifts」及各階段說明。')

# ───────── 6 實際走一遍 ─────────
s = content('實際走一遍（手冊範例）：「理賠進度自助查詢」從客服痛點到上線的七個步驟')
STEPS = [('理賠作業人員', '向 AI 描述問題：客戶常打電話問理賠進度，約三分之一通話時間花在查進度', 'intent.md'),
         ('產品負責人', '檢查並修正 AI 整理的需求說明，決定是否接受', '接受紀錄'),
         ('AI＋產品負責人', 'AI 產出規格並套用資安規範（入口網站不新增個資）；疑慮交給規範負責人處理', 'spec.md'),
         ('工程師＋AI', 'AI 寫實作計畫：改哪些檔案、順序、風險（核心系統每秒 50 次請求上限，要加快取）', 'plan.md'),
         ('AI', '照計畫寫程式並自己跑測試，四種理賠狀態都要通過，畫面截圖比對設計稿', '程式＋測試結果'),
         ('AI 審查＋人核准', 'AI 依審查規則查錯誤、資安、合規；程式負責人核准合併；發布經理授權上線', '審查與核准紀錄'),
         ('監控＋AI', '錯誤率超出正常範圍時，AI 先診斷，必要時執行事先演練的回復，寫成新的需求說明', '事件紀錄')]
for i, (who, what, art) in enumerate(STEPS):
    x = 0.45 + i * 1.78
    human = 'AI' != who
    tb(s, x, 1.45, 1.72, 0.85, [[(f'{i + 1}', {'size': 16, 'bold': True})], who], size=11, bold=True,
       color=WHITE, fill=OR if i in (1, 5) else (OR2 if human else MID), anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, margin=0.05)
    tb(s, x, 2.36, 1.72, 1.95, what, size=11, fill=LG, margin=0.09)
    tb(s, x, 4.37, 1.72, 0.52, art, size=10.5, bold=True, color=OR, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, margin=0.04)
    if i < 6: arrow(s, x + 1.66, 3.2, 0.16, 0.24, fill=OR)
tb(s, 0.45, 5.15, 12.45, 0.95, [[('BA 的觀察：', {'bold': True, 'color': OR}),
                                ('整個過程中，人做了三件事：說清楚問題（第 1 步）、決定要不要做（第 2 步）、決定能不能上線（第 6 步）。其餘起草、計畫、寫程式、測試、初審都由 AI 先做，再交給人看。', {})]],
   size=12, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.16)
note(s, f'深橘為人做決定的步驟。數字（三分之一通話時間、每秒 50 次）為手冊示例，非實際客戶資料。來源：{SRC}，第 12–20、48–52 頁。')

# ───────── 6b 手冊範例原文 ─────────
s = content('手冊範例原文：需求說明與實作計畫長什麼樣子（附中文翻譯）')
EXS = [('intent', '需求說明（intent.md）：理賠部門同仁和 AI 討論後產出', 0.45,
        ['標題：理賠進度自助查詢｜提出者：理賠作業 J. Ortiz｜狀態：草稿',
         '問題：客戶打電話問理賠進度，處理人員約三分之一通話時間只在回答進度。',
         '期望結果：客戶在網站上看到理賠狀態、下一步與預計日期。',
         '影響對象：理賠人員、網站團隊、理賠核心系統 API。',
         '限制：網站不新增個資，只用現有登入方式。',
         '待釐清：外部理賠公估人是否也需要使用？']),
       ('plan', '實作計畫（plan.md）：工程師核准後才開始寫程式', 6.85,
        ['標題：理賠進度自助查詢（依 2026-06-02 的需求說明）',
         '要改的檔案：網站理賠狀態面板（新增）、理賠 API 路由與測試。',
         '工作順序：① 在現有驗證下新增狀態 API　② 面板串接 API　③ 放進網站選單。',
         '風險：理賠核心 API 每秒最多 50 次請求，面板必須做快取。',
         '完成證明：測試涵蓋四種理賠狀態；畫面截圖與核准的設計稿一致。'])]
for key, head, x, lines in EXS:
    tb(s, x, 1.42, 6.05, 0.4, head, size=12, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.06)
    iw, ih = Image.open(FIG[key]).size
    pic(s, key, x + (6.05 - 2.95 * iw / ih) / 2, 1.88, h=2.95)
    tb(s, x, 4.92, 6.05, 1.68, [[('中文翻譯', {'bold': True, 'color': OR, 'size': 10})]] + lines, size=9.5, fill=TINT, margin=0.1)
note(s, f'圖：{SRC} 第 13、20 頁。兩份檔案都是一般文字，產品負責人和 AI 都能直接閱讀與修改；內容為手冊示例，非客戶資料。')

# ───────── 7 稽核軌跡 ─────────
s = content('每一步都留下紀錄：串起來就是一條完整的稽核軌跡，回答「誰要了什麼、誰核准」')
CHAIN = [('intent.md', '需求說明', '誰提出、什麼時候、要解決什麼'), ('spec.md', '規格書', '套用了哪些規範、哪些疑慮怎麼處理'),
         ('plan.md', '實作計畫', '誰核准了做法與風險'), ('程式＋測試', '改動本身', '改了什麼、測試是否通過'),
         ('PR 審查', '審查紀錄', 'AI 找到什麼、人怎麼決定'), ('事件紀錄', '維運紀錄', '發生什麼、怎麼處理、學到什麼')]
for i, (a, b, c_) in enumerate(CHAIN):
    x = 0.45 + i * 2.1
    tb(s, x, 1.45, 1.85, 0.75, [[(a, {'bold': True, 'size': 13})], [(b, {'size': 10.5})]], color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, margin=0.04)
    tb(s, x, 2.26, 1.85, 0.8, c_, size=10.5, fill=LG, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, margin=0.08)
    if i < 5: arrow(s, x + 1.88, 1.68, 0.18, 0.28)
tb(s, 0.45, 3.3, 6.0, 0.42, '已經有 Jira、ServiceNow 或需求管理系統，怎麼辦？', size=13, bold=True, color=OR)
SOT = [('以程式庫為準', '需求說明、規格等檔案是正式紀錄，舊系統放連結。一個工具、一個時間戳記，最乾淨。'),
       ('以現有系統為準', 'Jira 等系統是正式紀錄，AI 開工時讀取、做完寫回；檔案只是工作副本。'),
       ('至少互相連結（起步建議）', '檔案記下工單編號，工單記下檔案版本號。接受兩個來源並存，先讓兩邊能對上。')]
for i, (h, t) in enumerate(SOT):
    y = 3.8 + i * 0.9
    tb(s, 0.45, y, 2.3, 0.82, h, size=11.5, bold=True, color=OR if i == 2 else INK, fill=TINT if i == 2 else LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
    tb(s, 2.8, y, 3.65, 0.82, t, size=10.5, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
tb(s, 6.85, 3.3, 6.05, 0.42, '稽核時可以直接回答的三個問題', size=13, bold=True, color=OR)
AQ = [('誰要求了什麼？', '需求說明檔的作者、時間與修改歷程'), ('AI 產出了什麼？', '規格、計畫、程式與測試結果，以及當時套用的規範版本'),
      ('誰核准？', '每個關卡的核准人與時間，記在合併與審查紀錄中')]
for i, (q, a) in enumerate(AQ):
    y = 3.8 + i * 0.9
    circle(s, 6.85, y + 0.17, 0.48, str(i + 1), size=13)
    tb(s, 7.45, y, 5.45, 0.82, [[(q, {'bold': True, 'size': 12})], a], size=11, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
note(s, f'不需要汰換現有系統，重點是每一種紀錄都說清楚「以哪個系統為準」。來源：{SRC}，第 6–7、22 頁。')

# ───────── 8 人在哪裡把關 ─────────
s = content('人在哪裡把關：五個關鍵決定由指定的人做，AI 寫的程式不能自己核准')
ROWS = [['① 接受需求', '產品負責人', '這個想法值不值得做', '接受或關閉需求說明，留下紀錄'],
        ['② 確認規格', '產品負責人（高風險找技術主管）', '規格有沒有解決原本的問題，標出的疑慮是否已找規範負責人處理', '規格與需求說明一起存檔'],
        ['③ 核准計畫', '工程師（高風險找技術主管或架構師）', '要怎麼改、最大的風險在哪', '計畫模式下，核准前 AI 不能修改任何檔案'],
        ['④ 核准合併', '程式負責人', '改動是否符合原意、風險能否接受', '分支保護：沒有程式負責人核准就不能合併'],
        ['⑤ 授權上線', '發布經理', '能不能進入正式環境', '上線閘門：沒有授權，上線指令會被擋下']]
grid(s, 0.45, 1.4, [1.75, 3.0, 4.0, 3.7], ['關卡', '誰決定', '決定什麼', '怎麼確保一定經過這一關'], ROWS, 0.78, size=11)
tb(s, 0.45, 5.75, 0.1, 0.78, '', fill=OR)
tb(s, 0.55, 5.75, 12.35, 0.78, [[('原則：', {'bold': True, 'color': OR}),
                                 ('AI 可以一路做到正式環境的門口，但不能自己開門。寫程式的 AI 沒有核准權限，職責分離和現在一樣。', {})]],
   size=13, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.18)
note(s, f'來源：{SRC}，第 13、17、20、39、46 頁。')

# ───────── 8b 事件頻道（手冊圖 4） ─────────
s = content('上線後出事怎麼辦：AI 在事件頻道當第一線回應，但「要不要回復」由人說了算')
pw, ph = pic(s, 'fig4', 0.45, 1.45, w=6.6)
caption(s, 0.45, 1.5 + ph, 6.6, f'圖：{SRC} 圖 4「A first responder in the incident channel」')
CHAT = [('22:04', 'R. Mehta（工程師）', '結帳頁的 5xx 錯誤率從 21:40 上版後一直上升，@claude 幫忙看一下？', False),
        ('22:07', 'Claude', '處理中，正在比對錯誤特徵與這次上版的差異。原因看起來是新的快取鍵漏掉了租戶 ID。今天早上已在測試環境演練過回復，要執行嗎？', False),
        ('22:08', 'R. Mehta（工程師）', '執行。', True),
        ('22:12', 'Claude', '已回復。5xx 錯誤率回到正常範圍，事後檢討已寫入 lessons/2026-06-checkout-cache.md。', False)]
tb(s, 7.35, 1.45, 5.55, 0.4, '對話中文翻譯', size=13, bold=True, color=OR)
for i, (tm, who, msg, key) in enumerate(CHAT):
    y = 1.9 + i * 0.86
    tb(s, 7.35, y, 5.55, 0.8, [[(f'{tm}　{who}', {'bold': True, 'size': 10.5, 'color': OR if key else INK})], msg],
       size=10.5, fill=TINT if key else LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
tb(s, 0.45, 5.45, 12.45, 1.1, [[('重點：', {'bold': True, 'color': OR}),
                                ('AI 先診斷原因並提出建議，但只有工程師回覆「執行」之後才回復上一版。請求、診斷、人的授權、處理結果都留在同一個頻道，頻道本身就是稽核軌跡；事後檢討寫成檔案，下次遇到類似事件可以直接參考。', {})]],
   size=12, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.16)
note(s, f'Claude Tag 在手冊撰寫時為 Slack 上的公開測試版，對客戶說明前先確認現況。來源：{SRC}，第 52 頁。')

# ───────── 9 控制機制 ─────────
s = content('控制機制白話版：規範手冊讓 AI 照著做，自動閘門讓違規做不到')
CTRL = [('技能（Skill）', 'AI 會讀的公司規範手冊', '資安、API、品牌規範寫成檔案，AI 寫程式時自動套用。規範改了，大家下一次使用就是新版。', '建議性：讓違規變少'),
        ('鉤子（Hook）', 'AI 動作前的自動閘門', 'AI 每個動作前先檢查，可以放行、要求核准或擋下。例如沒有變更單不能改基礎設施，修錯時不能改測試。', '強制性：讓違規幾乎做不到'),
        ('集中管理設定', 'IT 統一鎖定的規則', '禁止讀取密碼與金鑰檔、限制可連的網站、只用公司核准的外掛。工程師無法自行關閉。', '由 IT 管理，個人改不動'),
        ('分環境授權', '環境越重要，AI 權限越小', '開發環境 AI 可以自由部署；正式環境由 AI 準備、發布經理授權。退回上一版是最常演練的流程。', '正式環境一定要人授權')]
for i, (name, plain, body, kind) in enumerate(CTRL):
    x = 0.45 + (i % 2) * 6.3; y = 1.45 + (i // 2) * 2.35
    tb(s, x, y, 6.05, 2.2, '', fill=LG)
    tb(s, x, y, 0.1, 2.2, '', fill=OR)
    tb(s, x + 0.3, y + 0.12, 5.6, 0.45, [[(name, {'bold': True, 'size': 14}), ('　＝　' + plain, {'size': 13, 'color': OR, 'bold': True})]], anchor=MSO_ANCHOR.MIDDLE)
    tb(s, x + 0.3, y + 0.62, 5.6, 1.0, body, size=11.5)
    tb(s, x + 0.3, y + 1.66, 2.9, 0.4, kind, size=10.5, bold=True, color=WHITE, fill=OR2, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
note(s, f'手冊原話：技能讓違規變少，鉤子讓違規幾乎不可能。來源：{SRC}，第 25–27、40–46 頁。', y=6.3)

# ───────── 10 客戶常問 ─────────
s = content('客戶最常問的五個問題，以及可以怎麼回答')
QA = [['會不會失控？', 'AI 只能走核准過的路徑。正式上線一定要指定的人授權，每一次放行或阻擋都有時間戳記。'],
      ['出事誰負責？', '每個需要判斷的決定都由人做，並記在版本紀錄中。寫程式的 AI 不能核准自己的程式，職責分離不變。'],
      ['會和現有 Jira、變更管理衝突嗎？', '不需要汰換。每種紀錄指定一個「以誰為準」的系統，其他系統放連結；起步可先做互相連結。'],
      ['機密資料安全嗎？', 'IT 可統一禁止 AI 讀取金鑰與機密檔、限制對外連線、在隔離環境執行。也可透過 AWS Bedrock、Google Vertex 或 Microsoft Foundry 使用，留在公司既有的雲端合約內。'],
      ['怎麼知道有沒有效？', '用版本紀錄與 CI 系統本來就有的資料量測，例如需求定稿時間、第一次就通過測試的比例、審查等待時間、上線失敗率。']]
grid(s, 0.45, 1.4, [3.2, 9.25], ['客戶問', '可以這樣回答'], QA, 0.94, size=12)
note(s, f'來源：{SRC}，第 22、37、41–46 頁。')

# ───────── 11 怎麼衡量 ─────────
s = content('怎麼衡量成效：數據大多已經在版本紀錄和 CI 系統裡，不需要另外填表')
ROWS = [['規劃', '從第一次討論到需求說明定稿的時間（手冊預期由數週降到數小時）', '被接受、進入設計的需求比例'],
        ['設計', '需求說明到規格定稿的時間', '開發開始後還要改規格的次數'],
        ['開發', '第一次實作就能合併的比例', '每個改動的返工次數'],
        ['測試', 'AI 改動第一次就通過 CI 的比例', '每個 PR 的審查時間、上線失敗率'],
        ['上線', '第一次審查的等待時間（目標降到分鐘級）', '合併前抓到的缺陷，對比上線後才發現的'],
        ['維運', '異常發生到寫成需求說明的時間', '同類事件重複發生的次數']]
grid(s, 0.45, 1.4, [1.2, 6.0, 5.25], ['階段', '先看（領先指標）', '後看（落後指標）'], ROWS, 0.66, size=11.5)
tb(s, 0.45, 5.92, 12.45, 0.55, [[('建議：', {'bold': True, 'color': OR}), ('試點前先記下現況基準，試點後用同一套指標比較，對外只引用自己量到的數字。', {})]],
   size=12, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.16)
note(s, f'整體交付績效可再搭配業界常用的 DORA 指標（部署頻率、交付時間、變更失敗率、復原時間）。來源：{SRC}，各階段「How to measure it」。', y=6.62)

# ───────── 11b 導入順序圖（手冊圖 3） ─────────
s = content('導入順序圖：橘色項目不需要任何前置條件，可以先做；箭頭代表誰要先於誰')
pw, ph = pic(s, 'fig3', 0.45, 1.45, w=8.2)
caption(s, 0.45, 1.5 + ph, 8.2, f'圖：{SRC} 圖 3「The play dependency graph」')
tb(s, 8.95, 1.45, 3.95, 0.4, '圖中方塊中文對照', size=13, bold=True, color=OR)
DEPS = [('Capture intent', '記錄需求（需求說明）', True), ('Requirements & design', '需求與設計規格', False), ('Skills', '技能（規範手冊）', False),
        ('CLAUDE.md', '團隊說明檔', True), ('Subagents', '子代理', False), ('Plan mode', '計畫模式', True), ('Evals', '持續評測', False),
        ('Feedback loop', 'AI 自我驗證', True), ('PR review', 'PR 審查', False), ('CI/CD', '自動化流水線', False), ('Hooks', '鉤子（自動閘門）', True),
        ('Closing the loop', '閉環維運', False)]
for i, (en, zh, entry) in enumerate(DEPS):
    y = 1.9 + i * 0.37
    tb(s, 8.95, y, 1.95, 0.33, en, size=9.5, bold=True, color=WHITE if entry else INK, fill=OR if entry else LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.06)
    tb(s, 10.95, y, 1.95, 0.33, zh, size=10, anchor=MSO_ANCHOR.MIDDLE, margin=0.04)
tb(s, 0.45, 5.0, 8.2, 1.55, [[('怎麼讀：', {'bold': True, 'color': OR})],
                             '圖由左到右分成規劃、設計、開發、測試、上線、維運六欄。橘色方塊是入口，不需要先做任何事就能開始；實線箭頭表示「A 要先於 B」；虛線表示「有幫助但不是必要」。',
                             '例如：PR 審查要等團隊說明檔與技能到位，閉環維運則要等審查、鉤子與 CI/CD 都準備好，所以放在最後。'],
   size=11, fill=TINT, margin=0.14)
note(s, f'來源：{SRC}，第 8–9 頁。')

# ───────── 12 怎麼開始 ─────────
s = content('怎麼開始：先做不需要前置條件的項目，再一步一步把循環接起來')
PH = [('第一步｜先做', '不需要任何前置條件', ['需求說明範本（intent.md）', '團隊說明檔（CLAUDE.md）', '把一條常被漏掉的規範寫成技能', '讓 AI 自己跑測試再交件', '用鉤子守住現有的核准關卡']),
      ('第二步｜串起來', '以第一步為基礎', ['設計規格由 AI 起草、人審核', '計畫模式成為開發的起點', 'AI 參與 PR 審查', 'AI 設定變更要跑持續評測']),
      ('第三步｜自動化', '關卡都到位之後', ['AI 進入 CI/CD（先唯讀，再寫入）', '監控觸發 AI 診斷，結果回到規劃', '事件頻道裡由 AI 擔任第一線回應'])]
for i, (h, sub, items) in enumerate(PH):
    x = 0.45 + i * 4.18
    tb(s, x, 1.45, 3.95, 0.75, [[(h, {'bold': True, 'size': 14})], [(sub, {'size': 10.5})]], color=WHITE, fill=OR if i == 0 else (OR2 if i == 1 else MID), anchor=MSO_ANCHOR.MIDDLE, margin=0.14)
    tb(s, x, 2.25, 3.95, 2.75, ['・' + t for t in items], size=13, fill=LG, margin=0.18)
    if i < 2: arrow(s, x + 3.97, 3.4, 0.2, 0.32)
tb(s, 0.45, 5.2, 12.45, 1.3, [[('第一個試點怎麼挑：', {'bold': True, 'color': OR})],
                              '依手冊建議，挑「規格清楚、影響範圍小、已經有測試覆蓋」的日常工作。先用人工逐步帶著 AI 做，關卡穩定後，再讓核准過的檔案自動觸發下一步。',
                              [('提醒：', {'bold': True, 'color': OR}), ('自動化要在核准關卡都到位之後才做，否則只是讓沒把關的東西更快上線。', {})]],
   size=12, fill=TINT, margin=0.16)
note(s, f'各項目的先後順序見上一頁導入順序圖。來源：{SRC}，第 8–9、21、44 頁。', y=6.62)

# ───────── 13–14 翻譯小抄 ─────────
GLOSS = [('SDLC', '軟體從想法到上線、再到維護的整套流程'),
         ('intent.md（需求說明）', '提出者用自己的話寫下：要什麼、為什麼、有什麼限制'),
         ('spec.md（規格書）', 'AI 依需求說明和公司規範起草，人負責審核'),
         ('plan.md（實作計畫）', '要改哪些地方、順序、風險、怎麼證明完成'),
         ('CLAUDE.md（團隊說明檔）', '新進同仁第一天要知道的事，AI 每次開工都先讀'),
         ('Skill（技能）', 'AI 會自動套用的公司規範手冊'),
         ('Hook（鉤子）', 'AI 動作前的自動檢查站：放行、詢問或擋下'),
         ('Plan mode（計畫模式）', 'AI 只能讀不能改，等人核准計畫才動手'),
         ('Auto mode（自動模式）', '計畫核准後，AI 不必每改一處都問人'),
         ('PR（Pull request）', '改動送審單：附上改了什麼、測試結果，等人核准'),
         ('Code review（程式審查）', 'AI 先查錯誤與資安，人判斷意圖與風險'),
         ('CI/CD', '自動建置、測試、部署的流水線'),
         ('Eval（評測）', '用一組真實任務考 AI，確認換模型或改設定後品質沒有退步'),
         ('Worktree／平行工作階段', '一位工程師同時開幾個 AI 分頭做不同工作，彼此不干擾'),
         ('Subagent（子代理）', '專做一件事的小幫手，例如驗證程式能不能正常執行'),
         ('MCP', '讓 AI 安全連接其他系統（Jira、部署工具）的標準介面'),
         ('Sandbox（沙箱）', '隔離環境：AI 在裡面執行，碰不到不該碰的檔案和網路'),
         ('Managed settings（集中管理設定）', 'IT 統一下發、工程師不能改的設定'),
         ('Control band（σ）', '指標的正常波動範圍，超出越多，AI 能做的處理越多'),
         ('Rollback（回復）', '一個指令退回上一版，平時就要演練'),
         ('DORA 指標', '業界常用的交付績效指標：部署頻率、交付時間、變更失敗率、復原時間'),
         ('OpenTelemetry', '把 AI 每個動作記錄下來、送到監控系統的標準格式')]
for k in range(2):
    s = content(f'BA 翻譯小抄（{k + 1}／2）：工程團隊常用詞，對管理層和客戶可以這樣說')
    part = GLOSS[k * 11:k * 11 + 11]
    grid(s, 0.45, 1.4, [3.9, 8.55], ['工程用語', '白話說法'], [list(g) for g in part], 0.43, size=11.5, head_h=0.38)
    note(s, '用語依手冊原文整理；Claude Tag、Claude Design 等產品名稱為手冊撰寫時的預覽或測試版功能，對客戶說明前先確認現況。')

# ───────── 15 Sam 的講法 ─────────
s = content('對客戶怎麼講：30 秒版本、可以強調的重點，以及要避免的說法')
tb(s, 0.45, 1.4, 12.45, 0.42, '30 秒版本', size=13, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.06)
tb(s, 0.45, 1.86, 12.45, 1.15, '「現在 AI 寫程式很快，但如果需求、審查、上線還照舊，整體不會快多少。我們協助你把開發流程重新設計：AI 參與每一步，每一步都留下紀錄，關鍵決定仍由你的人做。先從一個小系統試，用你們既有的版本紀錄和 CI 數據量測成效。」',
   size=13, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.2)
tb(s, 0.45, 3.2, 6.05, 0.42, '可以強調', size=13, bold=True, color=WHITE, fill=OR2, anchor=MSO_ANCHOR.MIDDLE, margin=0.06)
PROS = ['每一步都有紀錄，稽核時可以直接調出來', '寫程式的 AI 不能核准自己的程式，職責分離不變', '可沿用現有的 Jira、ServiceNow 與變更管理',
        '從不需要前置條件的項目開始，風險小', '正式上線一定要指定的人授權']
tb(s, 0.45, 3.68, 6.05, 2.85, ['・' + t for t in PROS], size=12, fill=LG, margin=0.16)
tb(s, 6.85, 3.2, 6.05, 0.42, '避免的說法 → 改成', size=13, bold=True, color=WHITE, fill=INK, anchor=MSO_ANCHOR.MIDDLE, margin=0.06)
CONS = [('「AI 會取代工程師」', '工程師的工作轉為規劃、指揮與審查'),
        ('「以後不用人審了」', '人審集中在高風險與受監理的程式'),
        ('「保證從數週變數小時」', '手冊預期可大幅縮短，實際以試點量測'),
        ('「AI 會自己上線」', 'AI 只做到上線前，授權一定由人來做')]
for i, (bad, good) in enumerate(CONS):
    y = 3.68 + i * 0.72
    tb(s, 6.85, y, 2.5, 0.66, bad, size=11, color=GREY, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
    arrow(s, 9.4, y + 0.21, 0.28, 0.24)
    tb(s, 9.75, y, 3.15, 0.66, good, size=11, bold=True, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
note(s, '推案前：客戶若同時是所內審計客戶，先與 Risk & Quality 確認獨立性；客戶資料不放進任何外部 AI 工具試用。', y=6.62)

# ───────── 16 來源 ─────────
s = content('來源與延伸閱讀')
tb(s, 0.45, 1.4, 12.45, 0.8, [[('主要來源：', {'bold': True, 'color': OR}), ('Louis Claxton（Anthropic），《The AI-Native SDLC Playbook》，2026。本簡報的流程、範例與控制機制皆整理自該手冊，並改寫為白話中文。', {})]],
   size=12, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.16)
DOCS = [('組織導入總覽（先看這份）', 'https://code.claude.com/docs/en/admin-setup'), ('設定與優先順序（含集中管理設定）', 'https://code.claude.com/docs/en/settings'),
        ('權限設定', 'https://code.claude.com/docs/en/permissions'), ('沙箱：檔案與網路隔離', 'https://code.claude.com/docs/en/sandboxing'),
        ('鉤子（Hooks）指南', 'https://code.claude.com/docs/en/hooks-guide'), ('技能（Skills）', 'https://code.claude.com/docs/en/skills'),
        ('外掛與私有市集', 'https://code.claude.com/docs/en/plugin-marketplaces'), ('集中管理 MCP', 'https://code.claude.com/docs/en/managed-mcp'),
        ('企業部署：Bedrock、Vertex、Foundry', 'https://code.claude.com/docs/en/third-party-integrations'), ('監控（OpenTelemetry）', 'https://code.claude.com/docs/en/monitoring-usage'),
        ('安全模型', 'https://code.claude.com/docs/en/security'), ('合規 API', 'https://platform.claude.com/docs/en/manage-claude/compliance-api')]
tb(s, 0.45, 2.3, 12.45, 0.38, '手冊列出的官方文件（依導入順序，可點擊）', size=12, bold=True, color=GREY)
for i, (lab, url) in enumerate(DOCS):
    col, row = divmod(i, 6)
    tb(s, 0.45 + col * 6.3, 2.75 + row * 0.58, 6.1, 0.52, [[(f'{i + 1}. ', {'bold': True, 'color': OR}), (lab, {'link': url})]], size=11.5, anchor=MSO_ANCHOR.MIDDLE, fill=LG, margin=0.12)
note(s, '連結取自手冊第 54 頁；本工作環境無法開啟網頁，使用前請點開確認。')

p.save(OUT)
print(OUT, len(p.slides), 'slides')
