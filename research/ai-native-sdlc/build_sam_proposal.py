# AI 原生軟體開發：給合夥人向企業客戶提案的簡報（v3：依所內提案範本的頁面格式，修正審閱意見）
# 章節沿用四個面向（面向誰、解決什麼問題、帶來什麼價值、可擴充與可複製）；外部證據經查核後引用原始出處；
# IT 現場示範為主要實證，示範畫面待 IT 團隊提供；不含任何個人姓名，同業公司名稱只在附錄
# 版面：內容頁用範本「3__Pattern header」（灰底頁首＋左側橘條），標題下一句結論；字型 Georgia＋微軟正黑體
# 用法：python3 build_sam_proposal.py <範本.pptx> <輸出.pptx> [--client]
#   預設為合夥人審閱版（含「請 Sam 決定」內部頁與研究整合附錄）；--client 產生客戶版，不含內部頁
import sys, copy, math, re
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

TPL, OUT = sys.argv[1], sys.argv[2]
INTERNAL = '--client' not in sys.argv[3:]
OR = RGBColor(0xFD, 0x51, 0x08); OR2 = RGBColor(0xFE, 0x7C, 0x39); TINT = RGBColor(0xFF, 0xE8, 0xD4)
# 範本佈景主題色：tx1 黑、lt2 淺灰、accent6 線條灰、accent4 中灰；範本不用純黑色塊，深色塊用深灰
INK = RGBColor(0x00, 0x00, 0x00); GREY = RGBColor(0x59, 0x59, 0x59); LG = RGBColor(0xEB, 0xEB, 0xEB); LINE = RGBColor(0xCB, 0xD1, 0xD6)
MID = RGBColor(0xA1, 0xA8, 0xB3); WHITE = RGBColor(0xFF, 0xFF, 0xFF); DARK = RGBColor(0x59, 0x59, 0x59); RED = RGBColor(0xDE, 0x2A, 0x18)
# 範本提案頁字型：英數 Georgia、中文微軟正黑體
FONT = 'Georgia'; HEAD_FONT = 'Georgia'; EA_FONT = '微軟正黑體'
SYMBOLS = set('●◐○✓✗→↑↓①②③④⑤⑥⑦⚠')
PROJECT = 'PwC FSC1 AI 原生軟體開發提案'
V3 = 'FSC1〈AI 協作開發流程與產出範本〉v3'
PB = 'Anthropic《The AI-Native SDLC Playbook》'
SMR = 'MIT Sloan Management Review，2024'
TODO = '[待填]'

def style_font(r, head=False):
    # Georgia 沒有 ●◐○ 等符號，這類文字改用中文字型顯示
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

# 內容頁的版面座標：各頁以 y=1.45–6.95 的區域排版，繪製時對應到結論句下方的 Y0–6.95
Y0, YB = 1.85, 6.95
K = (YB - Y0) / (YB - 1.45)
MAP = False

def Y(y): return Y0 + (y - 1.45) * K if MAP else y
def H(h): return h * K if MAP else h

# 頁碼與附錄代號在全部頁面建好後才確定，文字中先放 ⟦p:鍵⟧、⟦app:鍵⟧、⟦r:章節⟧，最後統一代換
PAGE, APP, SEC = {}, {}, []
CUR = ['']

def mark(key): PAGE[key] = len(p.slides)

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

def tb(s, x, y, w, h, text, size=11, bold=False, color=INK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, fill=None, line=None, margin=0.06, shape=MSO_SHAPE.RECTANGLE, spacing=None):
    sh = s.shapes.add_shape(shape, Inches(x), Inches(Y(y)), Inches(w), Inches(H(h)))
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
        if spacing: para.line_spacing = spacing
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

def circle(s, x, y, d, text, fill=OR, size=12, color=WHITE, line=None):
    # 圓形只對應位置不縮放，避免變成橢圓
    c = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(Y(y + d / 2) - d / 2), Inches(d), Inches(d))
    c.fill.solid(); c.fill.fore_color.rgb = fill; c.shadow.inherit = False
    if line is None: c.line.fill.background()
    else: c.line.color.rgb = line; c.line.width = Pt(1.5)
    tf = c.text_frame; tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    para = tf.paragraphs[0]; para.alignment = PP_ALIGN.CENTER
    r = para.add_run(); r.text = text; r.font.size = Pt(size); r.font.bold = True; r.font.color.rgb = color; style_font(r)

def hline(s, x1, y1, x2, y2, color=LINE, width=1.0):
    ln = s.shapes.add_connector(1, Inches(x1), Inches(Y(y1)), Inches(x2), Inches(Y(y2)))
    ln.line.color.rgb = color; ln.line.width = Pt(width)

def arrow(s, x, y, w, h, kind=MSO_SHAPE.RIGHT_ARROW, fill=OR):
    a = s.shapes.add_shape(kind, Inches(x), Inches(Y(y)), Inches(w), Inches(H(h)))
    a.fill.solid(); a.fill.fore_color.rgb = fill; a.line.fill.background(); a.shadow.inherit = False

def footer(s):
    # python-pptx 新增投影片不會帶入頁尾與頁碼，從版面複製（頁碼為自動欄位）；頁尾寫「提案名 | 章節」
    have = {ph.placeholder_format.idx for ph in s.placeholders}
    for ph in s.slide_layout.placeholders:
        if ph.placeholder_format.idx in (3, 4, 12) and ph.placeholder_format.idx not in have:
            s.shapes._spTree.append(copy.deepcopy(ph._element))
    set_ph(s, 3, f'{PROJECT} | {CUR[0]}' if CUR[0] else PROJECT)

def width(t):
    # 估計字寬：中文全形算 1，英數與半形符號算 0.55；⟦⟧ 代換後只剩一兩個字
    t = re.sub(r'⟦[^⟧]*⟧', 'X', t)
    return sum(1 if ord(c) > 0x2E80 else 0.55 for c in t)

TAGC = {'建議': OR, '示意': MID, '假設': GREY, '待確認': INK, '內部': RED}

def content(title, lead, tags=(), layout='3__Pattern header'):
    # 範本提案頁：灰底頁首放標題，標題下一句結論；右上方為建議、示意等標示（範本的框線印章樣式）
    global MAP
    MAP = False
    s = p.slides.add_slide(L[layout])
    if layout == 'Agenda':
        # 照範本議程頁：拿掉版面的表格預留位置，標題移到左上方
        for ph in list(s.placeholders):
            if ph.placeholder_format.idx not in (0, 3, 4): ph._element.getparent().remove(ph._element)
        t = [ph for ph in s.placeholders if ph.placeholder_format.idx == 0][0]
        t.left, t.top, t.width, t.height = Inches(0.46), Inches(0.51), Inches(8.0), Inches(0.83)
    set_ph(s, 0, title, None if width(title) <= 33 else 22, bold=True); footer(s)
    lw = 12.36 - 0.95 * len(tags)
    # 議程頁的範本版面沒有灰底頁首，標題較低，結論句跟著往下
    ly = 1.18 if layout == '3__Pattern header' else 1.36
    tb(s, 0.48, ly, lw, 0.5, lead, size=16 if width(lead) <= lw * 4.4 else 14, anchor=MSO_ANCHOR.MIDDLE, margin=0.04, spacing=1.1)
    for i, k in enumerate(tags):
        tb(s, 12.9 - 0.95 * (len(tags) - i), 1.24, 0.88, 0.36, k, size=11, bold=True, color=TAGC[k], fill=WHITE, line=TAGC[k], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.02)
    MAP = True
    return s

def note(s, text, y=6.66):
    tb(s, 0.45, y, 12.45, 0.32, text, size=9.5, color=GREY, margin=0.03)

def section(num, title, name):
    # 章節轉場頁：沿用範本「章節標題」版面，只放章節編號與名稱
    global MAP
    MAP = False; CUR[0] = name
    s = p.slides.add_slide(L['章節標題'])
    set_ph(s, 0, title, 30); set_ph(s, 1, num, 180); footer(s)
    SEC.append((num, len(p.slides)))
    return s

def appendix(key):
    APP[key] = chr(ord('A') + len(APP)); return f'⟦app:{key}⟧'

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


def draw_loop(s, cx, cy, R, d=1.05):
    # 自繪「循環」圖：六個階段排成一圈，中間是 AI 協作；不使用任何廠商的圖
    names = ['需求分析', 'SA&D', '開發', '測試', '上版／交付', '維運']
    pts = [(cx + R * math.cos(math.radians(-90 + 60 * i)), cy + R * math.sin(math.radians(-90 + 60 * i))) for i in range(6)]
    for i in range(6):
        (x1, y1), (x2, y2) = pts[i], pts[(i + 1) % 6]
        hline(s, x1, y1, x2, y2, OR2, 2.5)
    circle(s, cx - 0.75, cy - 0.75, 1.5, '')
    tb(s, cx - 0.75, cy - 0.4, 1.5, 0.8, ['AI 協作', [('參與每一步', {'size': 9.5})]], size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.0)
    for (x, y), n in zip(pts, names):
        circle(s, x - d / 2, y - d / 2, d, '', fill=WHITE if n != '維運' else TINT, line=OR)
        tb(s, x - d / 2, y - 0.25, d, 0.5, n, size=10.5 if len(n) < 5 else 9.5, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.0)

CHECK = [['是否屬作業委外、是否需依規定辦理', '保險業作業委託他人處理內部作業制度及程序辦法（2025-12-31 施行）', '導入前判斷並留下紀錄；AI 工具納入委外管理', '診斷'],
         ['雲端資料存放、雲端業者管理與查核', '保險業作業委外使用雲端服務自律規範', '使用公司核准的雲端合約；保留查核所需紀錄', '診斷'],
         ['AI 工具登錄 AI 清冊並做風險評估', '保險業運用人工智慧系統自律規範（2025-04-10 備查）', '每個 AI 工具與用途登錄；高風險用途加審', '診斷'],
         ['關鍵決定由人做，保險公司負最終責任', '同上；金管會 AI 指引「治理及問責」', '五個人工關卡，各有指定核准人', '①–⑤'],
         ['個資與客戶資料不進入 AI 可讀範圍', '個人資料保護法；金管會 AI 指引「隱私及客戶權益」', '資料分級；集中設定禁止讀取機密檔與金鑰', '規格、開發'],
         ['程式經資安審查，出事可退回上一版', '金管會 AI 指引「系統穩健性與安全性」', 'AI 資安初審＋人工審查；退版流程事先演練', '④ 合併、⑤ 上線'],
         ['誰要求、AI 產出什麼、誰核准，查得到', '金管會 AI 指引「透明性與可解釋性」', '需求說明、規格、計畫與核准紀錄版本化保存', '全程']]

# 試點驗收指標：都量得到、也可能不達標；核准紀錄完整是流程內建的必要條件，不當成效指標
KPI = ['需求定稿時間', '第一次審查等待時間', '返工次數']

def cards(s, items, y=1.45, h_big=0.95, h_txt=1.6):
    # 大數字卡：(標題, 大數字, 說明, 來源)
    n = len(items); w = (12.45 - 0.13 * (n - 1)) / n
    for i, (h, big, t, src) in enumerate(items):
        x = 0.45 + i * (w + 0.13)
        tb(s, x, y, w, 0.5, h, size=13, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
        tb(s, x, y + 0.55, w, h_big, big, size=30 if len(big) < 7 else 24, bold=True, color=OR, fill=LG, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
        tb(s, x, y + 0.55 + h_big, w, h_txt, t, size=11.5, fill=LG, margin=0.14)
        tb(s, x, y + 0.55 + h_big + h_txt, w, 0.55, '來源：' + src, size=9, color=GREY, fill=LG, margin=0.07)

STAGES = ['需求分析', 'SA&D', '開發', '測試', '上版／交付', '維運']

# ───────── 封面 ─────────
s = p.slides.add_slide(L['標題投影片'])
set_ph(s, 0, 'AI 原生軟體開發：\n從一個用例開始', 28)
set_ph(s, 1, 'PwC FSC1｜給保險業的提案｜2026 年 10 月', 12)

# ───────── 內部：請 Sam 決定（只在合夥人審閱版） ─────────
if INTERNAL:
    CUR[0] = '內部審閱'
    s = content('內部｜請 Sam 決定：是否以本提案接觸首批保險客戶',
                '建議：同意以「診斷＋一個用例試點」接觸 1–2 家保險公司；第 ⟦p:demo⟧ 頁示範完成後再對外。', tags=('內部',))
    COLS = [('請您決定', ['① 是否同意以本提案接觸首批客戶（候選名單：' + TODO + '）',
                         '② 附錄 ⟦app:peers⟧ 對客戶是否保留同業公司名稱',
                         '③ 指派 IT 同仁完成第 ⟦p:demo⟧ 頁示範畫面（目標日期：' + TODO + '）']),
            ('推案前要先確認', ['・獨立性：候選客戶若為 PwC 審計客戶，先經 Risk & Quality 確認（狀態：尚未檢查）',
                              '・責任範圍：上線由客戶指定的人授權；AI 產出的責任歸屬寫進合約，條文待 Legal 確認',
                              '・示範：畫面完成前，提案不對外']),
            ('規模與資源', ['・準備度診斷約 4–6 週＋一個用例試點約 6–8 週（第 ⟦p:offer⟧ 頁）',
                           '・人力配置與報價：' + TODO,
                           '・IT 團隊：示範準備，以及試點期間的技術支援'])]
    for i, (h, items) in enumerate(COLS):
        x = 0.45 + i * 4.18
        tb(s, x, 1.45, 4.05, 0.5, h, size=13.5, bold=True, color=WHITE, fill=OR if i == 0 else DARK, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
        tb(s, x, 2.0, 4.05, 2.35, [[(t, {})] for t in items], size=12.5, fill=TINT if i == 0 else LG, margin=0.16, spacing=1.1)
    QA = [['跟國際顧問的方案比，我們贏在哪？', '在地法規與可稽核關卡綁在一起，並從客戶的一個真實需求開始（第 ⟦p:why⟧ 頁；同業見附錄 ⟦app:peers⟧）'],
          ['AI 寫的程式出事，責任算誰的？', '能不能上線由客戶指定的人授權，誰核准都查得到（第 ⟦p:solution⟧ 頁）；合約責任範圍待 Legal 確認'],
          ['怎麼證明有效？', '試點前雙方講好指標與合格線，達標才擴大，未達標就停（第 ⟦p:offer⟧ 頁）']]
    grid(s, 0.45, 4.55, [3.6, 8.85], ['預期會被問的問題', '建議回答'], QA, 0.68, size=12)

# ───────── 一頁看懂 ─────────
CUR[0] = '提案摘要'
s = content('一頁看懂：從一個用例開始，證明 AI 開發更快也可稽核',
            '先做 4–6 週診斷與 6–8 週單一用例試點；達到事先講好的標準，才擴大到下一個系統。')
WHY = [('為什麼是現在', '委外新辦法與 AI 自律規範已上路，AI 產出的最終責任在保險公司；開發流程要拿得出「誰要求、AI 做了什麼、誰核准」的紀錄。', '第 ⟦p:urgency⟧ 頁'),
       ('為什麼只買工具不夠', 'DORA 2025：AI 讓交付變快，但交付穩定度仍在下降。效益取決於需求、審查、上線與控制方式有沒有一起改。', '第 ⟦p:tools⟧ 頁'),
       ('為什麼找我們', 'FSC1 自己已經用這套流程開發；台灣保險法規對應到每個關卡；從貴公司的一個真實需求開始，不必全面轉型。', '第 ⟦p:demo⟧、⟦p:why⟧ 頁')]
for i, (h, t, ref) in enumerate(WHY):
    x = 0.45 + i * 4.18
    tb(s, x, 1.45, 4.05, 0.55, f'{i + 1}　{h}', size=14, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
    tb(s, x, 2.05, 4.05, 2.05, t, size=14, fill=LG, margin=0.18, spacing=1.15)
    tb(s, x, 4.1, 4.05, 0.4, '詳見' + ref, size=10, color=GREY, fill=LG, margin=0.16)
tb(s, 0.45, 4.75, 0.1, 1.45, '', fill=OR)
tb(s, 0.55, 4.75, 6.1, 1.45, [[('請貴公司決定', {'bold': True, 'color': OR, 'size': 13})],
                              '安排一場現場示範，指定 IT 與業務窗口，選出 1–2 個候選用例，啟動準備度診斷。'], size=13.5, fill=TINT, margin=0.18)
tb(s, 6.8, 4.75, 6.1, 1.45, [[('如果先不做', {'bold': True, 'color': GREY, 'size': 13})],
                             'AI 工具照樣在開發團隊擴散，但流程與紀錄沒跟上；等到金檢或內稽時再補紀錄，成本更高、風險更大。'], size=13.5, fill=LG, margin=0.18)

# ───────── 議程（範本 Agenda 版面：編號圓圈、章節名稱、說明、頁碼） ─────────
CUR[0] = '議程'
s = content('議程', '從為什麼現在要做、怎麼做，到帶來什麼價值與如何擴大，說明從一個用例開始的導入方式。', layout='Agenda')
AG = [('1', '面向誰', '保險公司為什麼現在要做、優先服務誰', '01'),
      ('2', '解決什麼問題', '全球趨勢與只買工具為什麼不夠、我們的解法、AI 分工、現場示範', '02'),
      ('3', '帶來什麼價值', '價值與衡量、你買到什麼、為什麼選我們', '03'),
      ('4', '可擴充、可複製', '從一個用例逐步擴大、如何開始', '04'),
      ('5', '附錄', '示範情境、延伸方式、同業參考、法規清單、' + ('研究整合、' if INTERNAL else '') + '資料來源', 'A')]
for i, (n, h, d, key) in enumerate(AG):
    y = 1.5 + i * 0.95
    circle(s, 0.6, y + 0.13, 0.55, n, fill=WHITE, color=OR, line=OR, size=14)
    if i < len(AG) - 1: hline(s, 0.875, y + 0.68, 0.875, y + 1.08, OR, 1.0)
    tb(s, 1.45, y, 3.4, 0.8, h, size=20, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    tb(s, 4.9, y, 6.4, 0.8, d, size=12, color=GREY, anchor=MSO_ANCHOR.MIDDLE)
    tb(s, 11.3, y, 1.6, 0.8, f'⟦r:{key}⟧', size=22, bold=True, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT)
note(s, '右側數字為頁碼。頁面右上方標示「建議」「示意」「假設」「待確認」者為建議或示意內容；未標示者為有來源的事實或既有設計。', y=6.45)

# ───────── 01 面向誰 ─────────
section('01', '面向誰\nWho we serve', '面向誰')

s = content('保險公司現在就要把 AI 治理設計進開發流程',
            'AI 產出的最終責任已落在保險公司；開發流程要留下「誰要求、AI 做了什麼、誰核准」的紀錄。')
mark('urgency')
cards(s, [('責任落在保險公司', '2025-12-31', '保險業作業委外新辦法施行；AI 自律規範要求保險公司對 AI 結果負最終責任。', '理律法律事務所；壽險／產險公會'),
          ('監理要求可問責', '2024-06', '金管會 AI 指引要求治理及問責、透明性；新加坡金管局 2024-12 檢視銀行後指出，多數已建 AI 清冊，高風險 AI 須獨立驗證。', '金管會；新加坡金融管理局（MAS）'),
          ('開發需求集中', '2026-01-01', 'IFRS 17 與 TW-ICS 接軌上路，財務、精算與報表相關系統都要跟著調整，新的開發需求集中出現。', '金管會保險局')])
callout(s, 5.55, '意涵：', '紀錄必須在流程裡設計好，不能事後補；新制帶來的系統調整，正好是把 AI 與治理一起導入的時機。', h=0.75, size=12)

s = content('優先客群：有自建開發團隊的壽險與產險公司',
            '依經驗推估：IT 主管決策與付費，法遵、風控與資安握有否決權；決策鏈需在診斷訪談中確認。', tags=('假設',))
ROWS = [['IT／數位主管', '決策與付費', '交付跟不上業務與新制需求；AI 工具用了，整體卻沒變快', '一個可量測的用例試點'],
        ['業務單位主管', '提出需求、共同贊助', '需求來回多輪、上線時間難預期', '業務和 AI 一起寫需求說明'],
        ['法遵、風控、資安', '否決權', '委外、個資與 AI 責任；稽核要拿得出紀錄', '法規檢查清單、人工關卡、完整紀錄'],
        ['管理層', '核准預算', 'AI 投資要看到回報', '小範圍起步，先講好成功標準']]
grid(s, 0.45, 1.45, [2.4, 1.9, 4.3, 3.85], ['角色', '採購中的角色', '他們的問題', '我們給的答案'], ROWS, 0.85, size=12)
tb(s, 0.45, 5.4, 12.45, 0.85, [[('優先客群的條件：', {'bold': True, 'color': OR}), ('有自建或駐點開發團隊；已有版本控制與 CI；近期有新制相關或核心系統周邊的開發需求。', {})]], size=12, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.16)

# ───────── 02 解決什麼問題 ─────────
section('02', '解決什麼問題\nProblems we solve', '解決什麼問題')

s = content('AI 已普及，但交付沒有跟著變好：AI 放大原有的流程與治理',
            'AI 使用率已達八至九成，但交付穩定度下降；DORA 2026 指出，只買授權不保證回報。')
mark('tools')
# 上排：AI 已普及的五個數字；下排：交付表現、既有系統的效益、DORA 2026 結論
STATS = [('開發者', '84%', '正在使用或計畫使用 AI 工具'), ('技術團隊', '約 90%', '在工作中使用 AI'),
         ('JPMorgan', '10–20%', '工程師生產力提升（公司自述）'), ('Citi', '3 萬名', '開發者導入 AI 寫程式助手'),
         ('保險業', '57%', '高階主管把生成式 AI 列為 2026 年首要科技投資')]
w = (12.45 - 0.1 * 4) / 5
for i, (h, big, t) in enumerate(STATS):
    x = 0.45 + i * (w + 0.1)
    tb(s, x, 1.45, w, 0.4, h, size=12, bold=True, color=WHITE, fill=OR if i < 4 else DARK, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
    tb(s, x, 1.9, w, 1.15, [[(big, {'bold': True, 'size': 24, 'color': OR})], [(t, {'size': 10.5})]], fill=LG if i < 4 else TINT, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.08)
COLS = [('交付表現（DORA 2025）', [('↑', '交付速度：轉為正相關'), ('↓', '交付穩定度：仍為負相關')]),
        ('既有系統的效益較小', [('35–40%', '新開發的簡單任務'), ('≤ 10%', '複雜的既有系統（常見）')]),
        ('研究結論（DORA 2026）', [('放大器', '回報來自內部平台、清楚的流程與團隊協作，不是工具本身')])]
for i, (h, nums) in enumerate(COLS):
    x = 0.45 + i * 4.18
    tb(s, x, 3.25, 4.05, 0.42, h, size=12.5, bold=True, color=WHITE, fill=OR if i < 2 else DARK, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
    for j, (big, lab) in enumerate(nums):
        hh = 1.5 / len(nums)
        tb(s, x, 3.72 + j * hh, 4.05, hh, [[(big + '　', {'bold': True, 'size': 22 if len(big) < 6 else 20, 'color': OR}), (lab, {'size': 11.5})]], fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.14)
callout(s, 5.42, '所以：', '效益取決於流程與治理，不取決於買了哪個工具。導入 AI 時，需求、審查、上線與控制方式要一起改，這正是我們協助的部分。', h=0.7, size=12)
note(s, '來源：Stack Overflow 2025；Google DORA 2024、2025、2026；JPMorgan（公司自述）；Deloitte Insights 2025；PwC 保險高階主管調查。完整連結見附錄 ⟦app:src⟧。', y=6.22)
note(s, '既有系統的生產力數字為 DORA 2026 引述的 Stanford 研究。DORA 2024 模型估計：AI 採用率每增加 25%，穩定度 −7.2%、速度 −1.5%。顧問業者已推出相關方案（附錄 ⟦app:peers⟧）。', y=6.52)

s = content('我們的解法：把開發流程串成循環，用三個機制守住',
            'AI 參與每一步；接受需求、核准合併、授權上線等五個關鍵決定，由貴公司指定的人做。')
mark('solution')
draw_loop(s, 2.75, 3.45, 1.5, d=0.95)
ROWS = [['交接靠檔案', '需求說明 → 規格 → 實作計畫 → 測試結果 → 核准紀錄，下一步直接接手', '需求反覆、交接失真'],
        ['規則自動執行', '資安、個資、委外規範寫成 AI 自動套用的規則；必守的事由自動關卡擋下', '合規與品質靠個人'],
        ['關卡有人簽', '接受需求、確認規格、核准計畫、核准合併、授權上線，五個決定由指定的人做', '怕失控、稽核查不到']]
grid(s, 5.2, 1.45, [1.75, 3.9, 2.05], ['機制', '怎麼運作', '解決的問題'], ROWS, 1.12, size=11.5)
callout(s, 5.7, '原則：', 'AI 可以一路做到正式環境的門口，但不能自己開門。本方法不綁特定 AI 工具，依貴公司政策與合約選定。', h=0.75)
note(s, f'流程整合自 {V3}。DORA 2026 建議以不可略過的檢查點、提交前自動檢查、小批次與版本控制承接 AI 帶來的審查負擔，與三個機制一致（附錄 ⟦app:src⟧）。', y=6.45)

s = content('AI 分工：重複、判斷少的先交給 AI；關鍵決定永遠由人做',
            '依任務「多常發生 × 需要判斷」分四類；第一個試點用例，優先挑左上角工作多的需求。')
mark('tasks')
MX, MY, QW, QH = 1.15, 1.5, 3.6, 2.15
Q = [(0, 0, '常發生 × 判斷少', 'AI 先做，人確認', ['欄位清單、Test Case、Test Data、', 'Unit Test、註解與文件、Release Note'], TINT),
     (1, 0, '常發生 × 判斷多', 'AI 起草，人逐項審查', ['程式碼、API 規格、ERD、', 'Code Review 風險檢查、測試失敗原因分析'], LG),
     (0, 1, '少發生 × 判斷少', '做成範本，重複使用', ['Deploy Guide、User Guide、', 'Migration Script、部署腳本'], LG),
     (1, 1, '少發生 × 判斷多', '人決定，AI 只提供資料', ['接受需求、架構選型、核准合併、', '授權上線、事件處置'], WHITE)]
for c_, r, h, rule, items, bg in Q:
    x, y = MX + c_ * (QW + 0.06), MY + r * (QH + 0.06)
    tb(s, x, y, QW, QH, '', fill=bg, line=LINE)
    tb(s, x + 0.12, y + 0.08, QW - 0.24, 0.62, [[(h, {'bold': True, 'size': 11, 'color': GREY})], [(rule, {'bold': True, 'size': 13, 'color': OR})]], margin=0.02)
    tb(s, x + 0.12, y + 0.76, QW - 0.24, QH - 0.84, items, size=11.5, margin=0.02)
tb(s, 0.45, MY, 0.65, QH, '常發生', size=11, bold=True, color=GREY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.0)
tb(s, 0.45, MY + QH + 0.06, 0.65, QH, '少發生', size=11, bold=True, color=GREY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.0)
tb(s, MX, MY + 2 * QH + 0.1, 2 * QW + 0.06, 0.3, '需要判斷：少 ─────────────→ 多', size=10.5, bold=True, color=GREY, align=PP_ALIGN.CENTER, margin=0.0)
tb(s, 8.75, 1.5, 4.15, 4.4, [[('對貴公司的意義', {'bold': True, 'color': OR, 'size': 13})],
                             [('先挑左上角：', {'bold': True}), ('第一個試點用例，優先挑這類工作多的需求，最快看到效果。', {})],
                             [('', {'size': 5})],
                             [('右下角就是五個關卡：', {'bold': True}), ('這些決定由貴公司指定的人做，AI 只提供資料（第 ⟦p:solution⟧ 頁）。', {})],
                             [('', {'size': 5})],
                             [('依實際專案調整：', {'bold': True}), ('分類在準備度診斷時，依貴公司的專案與規範確認（第 ⟦p:offer⟧ 頁）。', {})]], size=12, fill=LG, margin=0.18)
note(s, f'整理自 FSC1 開發流程實務；方法參考 {SMR}。分類為初步判斷。', y=6.45)

s = content('現場示範：用一個用例走完從需求到上線的每一步',
            'FSC1 已在使用這套做法；現場可用貴公司提供的一個需求，走完以下五個步驟。', tags=('待確認',))
mark('demo')
FR = [('需求說明', '業務需求 → AI 追問 → 定稿'), ('規格與規範', 'AI 起草，自動套用規範'), ('實作計畫', '核准前 AI 不能改程式'), ('程式與測試', 'AI 自己跑測試'), ('核准與紀錄', '誰核准、何時，查得到')]
W = 2.41
for i, (h, t) in enumerate(FR):
    x = 0.45 + i * (W + 0.1)
    tb(s, x, 1.45, W, 0.45, f'{i + 1}　{h}', size=12, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.08)
    tb(s, x, 1.95, W, 1.9, '【IT 示範畫面】', size=11, bold=True, color=MID, fill=WHITE, line=OR, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    tb(s, x, 3.9, W, 0.55, t, size=10.5, fill=LG, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, margin=0.06)
PROVE = [('不是概念', '我們自己已經在用，畫面是實際操作'), ('每一步可稽核', '每個關卡都有人核准、留下紀錄'), ('一個用例就能開始', '不需要先改造整個 IT 部門')]
for i, (h, t) in enumerate(PROVE):
    x = 0.45 + i * 4.18
    tb(s, x, 4.65, 4.05, 1.05, [[(f'證明 {i + 1}：{h}', {'bold': True, 'color': OR, 'size': 12.5})], t], size=11.5, fill=TINT, margin=0.14)
note(s, '示範畫面由 FSC1 IT 團隊提供，使用假資料，不含任何客戶資料。現場示範可用貴公司提供的一個需求進行。', y=5.9)

# ───────── 03 帶來什麼價值 ─────────
section('03', '帶來什麼價值\nValue we deliver', '帶來什麼價值')

s = content('帶來什麼價值：每個痛點對應一項成果與衡量指標',
            '試點以流程指標驗收；DORA 四項交付指標需要較長期間與足夠的上線次數，於擴大階段起追蹤。', tags=('建議',))
ROWS = [['需求來回多輪', '業務和 AI 一起寫需求說明', '新功能更快上線', '從第一次討論到需求定稿的時間'],
        ['審查排隊、品質看人', 'AI 先審、人判斷風險', '工程師把時間花在需要判斷的工作', '第一次審查的等待時間、返工次數'],
        ['上線出錯', 'AI 每次改動自行測試', '上線事故減少', '變更失敗率、復原時間（DORA，擴大階段）'],
        ['法規與稽核壓力', '規範規則＋自動關卡＋紀錄', '金檢與內稽時拿得出佐證', '內稽調閱佐證所需時間'],
        ['交付慢且難預期', '每階段交出檔案，下一步直接接手', '交付時間可預期', '交付時間、部署頻率（DORA，擴大階段）']]
grid(s, 0.45, 1.45, [2.4, 3.2, 3.2, 3.65], ['痛點', '→ 解決方式', '→ 商業價值', '→ 衡量指標'], ROWS, 0.84, size=11.5)
note(s, 'DORA 四項指標（部署頻率、交付時間、變更失敗率、復原時間）為業界通用的交付績效指標。效益不預先承諾數字，以實測為準。', y=6.25)
note(s, 'DORA 2026：回收期 6–9 個月是敏捷團隊的優異目標；需要大量治理的大型企業導入，12–18 個月屬可接受。', y=6.55)

s = content('你買到什麼：先診斷，再做一個用例的試點',
            '成功標準在試點前雙方講好，達標才擴大；固定範圍、固定價格，金額另行報價。', tags=('建議', '示意'))
mark('offer')
OFF = [('準備度診斷', '約 4–6 週', ['盤點現況與卡點、法規檢查', '挑出第一個用例', '雙方訂好成功標準'], '瓶頸地圖、法規檢查結果、用例建議、導入藍圖'),
       ('一個用例的試點', '約 6–8 週', ['用一個真實需求走完循環', '建立範本、規範規則與關卡', '同步教會各角色'], '完成的試點、範本與規則、教材、成效報告')]
for i, (h, d, items, out) in enumerate(OFF):
    x = 0.45 + i * 4.3
    tb(s, x, 1.45, 4.15, 0.6, [[(h, {'bold': True, 'size': 15}), ('　' + d, {'size': 11})]], color=WHITE, fill=OR if i == 0 else OR2, anchor=MSO_ANCHOR.MIDDLE, margin=0.14)
    tb(s, x, 2.1, 4.15, 1.6, [[('・' + t, {})] for t in items], size=12, fill=LG, margin=0.16)
    tb(s, x, 3.75, 4.15, 0.95, [[('交付：', {'bold': True, 'color': OR}), (out, {})]], size=11.5, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.14)
arrow(s, 4.62, 1.6, 0.11, 0.3)
tb(s, 9.05, 1.45, 3.85, 0.6, '成功標準（試點前講好）', size=13, bold=True, color=WHITE, fill=DARK, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
tb(s, 9.05, 2.1, 3.85, 2.6, [[('・' + k, {})] for k in KPI] + [[('・核准紀錄完整（流程內建的必要條件）', {})], [('', {'size': 5})],
                             [('達標：', {'bold': True}), ('擴大到下一個用例', {})], [('未達標：', {'bold': True}), ('檢討後再試或停止', {})]], size=11.5, fill=LG, margin=0.16)
callout(s, 4.95, '貴公司需配合：', 'IT 與業務窗口、一個真實需求與團隊、測試環境、AI 工具的合約或內部核准。', h=0.6)
note(s, '期程依範圍確認。擴大階段、工具包授權與駐點支援可依需求討論。', y=5.7)

s = content('為什麼選我們：已在用的做法，加上在地法規與可稽核關卡',
            '國際方案多以預建代理與藍圖為主；我們綁上在地法規與可稽核關卡，從貴公司的一個真實需求開始。')
mark('why')
ROWS = [['已經在用', '我們自己的開發流程已經這樣運作：各階段的產出、AI 分工與檢查點都已定義', '現場示範（第 ⟦p:demo⟧ 頁）'],
        ['台灣保險法規內建', '委外、雲端、AI 自律規範與金管會指引，對應到開發的每一步', '法規檢查清單（附錄 ⟦app:rules⟧）'],
        ['懂保險業務與在地系統', '用例直接選理賠、核保、核心系統周邊等真實流程，不必從頭解釋業務', '[待補：2–3 個 FSC1 保險業專案]'],
        ['可稽核的人工關卡', '誰核准、何時核准都查得到；寫程式的 AI 不能核准自己的程式', '關卡設計（第 ⟦p:solution⟧ 頁）']]
grid(s, 0.45, 1.45, [2.6, 6.6, 3.25], ['差異', '貴公司得到什麼', '證據'], ROWS, 1.12, size=12.5, hot=(0,))
note(s, '同業具體產品與來源見附錄 ⟦app:peers⟧。', y=6.3)

# ───────── 04 可擴充、可複製 ─────────
section('04', '可擴充、可複製\nScalable and repeatable', '可擴充、可複製')

s = content('從一個用例開始，驗證價值後再逐步擴大',
            '每一步只承諾下一步、每一步都可以停，不必一開始就承諾全面轉型。', tags=('建議',))
STAIR = [('一個用例', '試點一個範圍小、可量測的需求', '承諾：診斷＋一個試點'),
         ('驗證價值', '對照試點前講好的成功標準', '決定：擴大、調整或停止'),
         ('相鄰用例', '沿用範本、規則與關卡，只補該系統的規範', '承諾：下一個用例'),
         ('全公司', '集中管理設定，統一關卡與紀錄', '承諾：依前面成果決定')]
for i, (h, t, c) in enumerate(STAIR):
    x = 0.45 + i * 2.35; h_ = 1.4 + i * 0.75; y = 5.35 - h_
    tb(s, x, y, 2.25, h_, [[(f'{i + 1}　{h}', {'bold': True, 'size': 13, 'color': WHITE})], [(t, {'size': 10.5, 'color': WHITE})]], fill=[OR, OR2, OR, DARK][i], margin=0.12)
    tb(s, x, 5.42, 2.25, 0.6, c, size=10, bold=True, color=OR, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, margin=0.06)
tb(s, 9.95, 1.45, 2.95, 4.57, [[('為什麼這樣做', {'bold': True, 'color': OR, 'size': 13})],
                               [('風險小：', {'bold': True}), ('第一步只動一個用例。', {})],
                               [('成本可控：', {'bold': True}), ('每一步只承諾下一步。', {})],
                               [('有數據才擴大：', {'bold': True}), ('每一步都對照同一套指標。', {})],
                               [('越做越快：', {'bold': True}), ('範本與規則沿用，下一個用例不必從頭來。', {})],
                               [('先講好 J 曲線：', {'bold': True}), ('DORA 2026 指出導入初期會有短暫下滑，常見的失敗是在低谷時撤資；成功標準與期間事先講好。', {})]], size=11, fill=LG, margin=0.16)
note(s, '各用例可沿用與需客製的內容、延伸到其他系統與產業的方式，見附錄 ⟦app:ext⟧。', y=6.2)

s = content('如何開始：先看現場示範，再挑一個用例',
            '從現場示範到啟動診斷共三步；請貴公司先指定窗口，並提供候選需求清單。', tags=('示意',))
PH = [('0　準備度診斷', '約 4–6 週', '選定用例與成功標準'), ('1　用例試點', '約 6–8 週', '指標達到合格線'), ('2　相鄰用例', '依範圍', '管理層核准擴大'), ('3　全公司', '持續', '—')]
for i, (h, d, g) in enumerate(PH):
    x = 0.45 + i * 3.14
    tb(s, x, 1.45, 3.0, 0.7, [[(h, {'bold': True, 'size': 13.5})], [(d, {'size': 10.5})]], color=WHITE, fill=[OR, OR, OR2, MID][i], anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
    tb(s, x, 2.2, 3.0, 0.5, [[('進下一步：', {'bold': True, 'color': OR}), (g, {})]], size=11, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
    if i < 3: arrow(s, x + 3.0, 1.63, 0.14, 0.3)
NX = [('現場示範', '用貴公司提供的一個需求，示範從需求說明到核准紀錄。'), ('對齊工作坊', '挑 1–2 個候選用例，確認前提與限制。'), ('啟動準備度診斷', '確認範圍與雙方窗口，開始盤點現況與法規檢查。')]
for i, (h, t) in enumerate(NX):
    y = 2.95 + i * 0.92
    circle(s, 0.45, y + 0.13, 0.55, str(i + 1), size=15)
    tb(s, 1.15, y, 6.55, 0.82, [[(h, {'bold': True, 'size': 13})], t], size=11.5, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.14)
tb(s, 8.0, 2.95, 4.9, 2.66, [[('請貴公司提供', {'bold': True, 'color': OR, 'size': 13})],
                             '・IT 主管與業務單位窗口', '・現有開發流程與工具（Jira、CI、變更管理）', '・候選需求清單', '・資安、個資與法遵規範'], size=12, fill=TINT, margin=0.18)
callout(s, 5.85, '一句話：', 'AI 幫你把每一步做得更快，每一步都留下紀錄；要不要做、能不能上線，永遠由你的人決定。', h=0.6, size=12)

# ───────── 附錄 ─────────
section('A', '附錄', '附錄')

s = content(f'附錄 {appendix("scenario")}｜示範情境：「理賠進度自助查詢」導入前後',
            '以一個常見的保險需求，說明每個階段 AI 做什麼、由誰決定。', tags=('示意',))
ROWS = [['需求分析', '多次會議，再由專人寫成文件', '理賠同仁和 AI 一起整理需求說明，AI 追問範圍與限制', '產品負責人接受'],
        ['SA&D', '分析與設計分開做，規範靠人記得', 'AI 起草規格，自動套用「入口網站不新增個資」等規範，衝突處標出', '產品負責人確認'],
        ['開發', '讀完設計就寫，做法只在個人腦中', 'AI 先寫實作計畫（含核心系統流量限制與快取），核准後才動手', '工程師核准'],
        ['測試', '階段末才測，問題晚發現', 'AI 每次改動自己跑測試，各種理賠狀態都正確才交人', 'AI 自行驗證'],
        ['上版／交付', '逐行人工審查，排隊等待核准', 'AI 先審並排嚴重度，人判斷風險；沒有授權不能上線', '程式負責人、發布經理'],
        ['維運', '事件常被擱置', '錯誤率異常時 AI 先診斷並建議退版；服務負責人核准後執行，問題寫成新需求', '服務負責人核准']]
grid(s, 0.45, 1.45, [1.5, 3.2, 5.75, 2.0], ['階段', '導入前（常見做法）', '導入後', '誰決定'], ROWS, 0.72, size=10.5)
note(s, f'情境改編自 {PB} 的公開範例，非客戶案例；「導入前」為常見做法的示意。', y=6.3)

s = content(f'附錄 {appendix("ext")}｜延伸方式：同一套資產延伸到其他系統、客戶與產業',
            '流程、範本與關卡可以沿用；不同延伸方向要補的內容與前提不同。', tags=('建議',))
ROWS = [['同一客戶的其他系統', '理賠 → 核保規則、保單服務', '流程、範本、關卡、規範規則', '該系統的規範與核准人', '前一個用例達標'],
        ['新制相關專案', 'IFRS 17／TW-ICS 報表與資料調整', '流程、關卡、法規清單', '財會與精算規範、資料品質檢查', '財會、精算參與；高風險加審'],
        ['AI 治理', 'AI 清冊與模型驗證流程', '紀錄與關卡、自律規範對應', '模型驗證規則', '屬治理服務，另行界定'],
        ['其他保險公司', '另一家壽險或產險', '全部標準模組與資產', '該公司規範與系統串接', '審計客戶需先確認獨立性'],
        ['其他金融業', '銀行、證券', '流程、關卡、指標', '該產業法規規則', '以金管會 AI 指引為共同基礎']]
grid(s, 0.45, 1.45, [2.3, 2.85, 2.5, 2.4, 2.4], ['延伸方向', '例子', '沿用什麼', '要新增什麼', '前提'], ROWS, 0.75, size=11)
tb(s, 0.45, 5.7, 12.45, 0.6, [[('需客製的部分：', {'bold': True, 'color': OR}), ('貴公司的資安與命名規範、系統與工具串接、資料分級與核准人、用例題目與基準、委外安排。', {})]], size=11.5, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.14)

s = content(f'附錄 {appendix("peers")}｜同業產品化參考：具體產品與來源',
            '國際顧問與軟體業者已推出保險業 AI 相關方案；以下整理自業者公開資訊。')
ROWS = [('Deloitte', 'Agentic AI Blueprint for Insurance', '加速器＋導入', 'https://www.deloitte.com/global/en/what-we-do/capabilities/agentic-ai-blueprint-for-insurance.html'),
        ('Deloitte', 'IFRS 17 Regulatory Assist（與 SAS 合作）', '加速器', 'https://sas.com/nl_nl/news/press-releases/2022/may/deloitte-partners-with-sas-to-power-ifrs-17-accelerator-solution.html'),
        ('KPMG', 'IFRS 17 Express, powered by SAS', '加速器', 'https://sas.com/en_ca/news/press-releases/2022/september/pvi-insurance-deploys-kpmg-ifrs-17-express.html'),
        ('Capgemini', 'Guidewire 生成式 AI／代理工具組', '核心系統＋AI', 'https://www.capgemini.com/wp-content/uploads/2026/10/GW-datasheet_Final-version.pdf'),
        ('WTW', 'Radar 5（保險定價與核保軟體）', '軟體產品', 'https://www.wtwco.com/en-gh/news/2025/10/wtw-launch-of-radar-5-with-gen-ai-capability-marks-a-major-milestone-in-insurance-pricing'),
        ('McKinsey', '代理型 AI 核心現代化', '方法論', 'https://www.mckinsey.com/industries/financial-services/our-insights/can-agentic-ai-finally-modernize-core-technologies-in-insurance'),
        ('BCG', 'AI-First 產險公司藍圖', '方法論', 'https://www.bcg.com/assets/2026/executive-perspectives-ai-first-companies-win-the-future-property-casualty-insurance.pdf')]
xx = [0.45, 2.0, 6.0, 7.75]
for j, h in enumerate(['公司', '產品／方案', '型態', '來源（可點擊）']):
    tb(s, xx[j], 1.45, (xx[j + 1] if j < 3 else 12.9) - xx[j] - 0.04, 0.4, h, size=11, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.07)
for i, (c, prod, kind, url) in enumerate(ROWS):
    y = 1.9 + i * 0.56; bg = LG if i % 2 == 0 else WHITE
    tb(s, xx[0], y, 1.51, 0.52, c, size=11, bold=True, color=OR, fill=bg, anchor=MSO_ANCHOR.MIDDLE, margin=0.07)
    tb(s, xx[1], y, 3.96, 0.52, prod, size=11, fill=bg, anchor=MSO_ANCHOR.MIDDLE, margin=0.07)
    tb(s, xx[2], y, 1.71, 0.52, kind, size=11, fill=bg, anchor=MSO_ANCHOR.MIDDLE, margin=0.07)
    tb(s, xx[3], y, 5.11, 0.52, [[(url.split('/')[2] + ' 原文', {'link': url})]], size=11, fill=bg, anchor=MSO_ANCHOR.MIDDLE, margin=0.07)
note(s, '規模與期程數字為業者自述。', y=5.95)

s = content(f'附錄 {appendix("rules")}｜保險業 AI 開發法規檢查清單（草稿）',
            '把台灣保險業的委外、雲端、AI 與個資規範，對應到開發流程中的檢查點。')
grid(s, 0.45, 1.45, [3.2, 3.85, 3.85, 1.55], ['檢查項目', '依據', '開發流程中怎麼落實', '在哪一步檢查'], CHECK, 0.66, size=10.5)
note(s, '導入規劃工具，非法律意見；條文以主管機關公告為準；草稿待法遵確認。', y=6.55)

if INTERNAL:
    s = content(f'附錄 {appendix("research")}｜研究整合：兩份研究互補，衝突處的取捨',
                '內部參考：兩份研究的差異，以及本提案採用哪一邊、為什麼。提交客戶前移除本頁。', tags=('內部',))
    ROWS = [['流程階段', '五階段：需求分析、SA&D、開發、測試、上版／交付', '六階段：規劃、設計、開發、測試、上線、維運', '採研究一的名稱，加上維運：保險業 IT 熟悉這些用語；維運讓流程形成循環'],
            ['控制方式', 'Prompt 撰寫原則＋IT 人工 Review', '技能、鉤子、集中管理設定', '分層：Prompt 為第一層，往規則與自動關卡加強；受監理行業需要可強制的控制'],
            ['關卡', '每階段文件確認後進下一階段', '五個關鍵決定由指定的人核准', '保留研究一的檢查點，補上指定核准人與紀錄'],
            ['效益說法', '未提數字', '手冊預期「數週到數小時」', '不承諾數字，以試點實測；外部證據改用 DORA 等獨立研究'],
            ['實作證據', '各階段案例頁待補', '實績頁待 IT 團隊填寫', '以 IT 現場示範補上（第 ⟦p:demo⟧ 頁）'],
            ['服務內容', '未涵蓋', '四個服務模組', '對外收斂為「診斷＋一個用例試點」，其餘依需求討論']]
    grid(s, 0.45, 1.45, [1.6, 3.3, 3.2, 4.35], ['主題', '研究一：FSC1 開發流程 v3', '研究二：AI 原生 SDLC 說明', '本提案的取捨與原因'], ROWS, 0.78, size=10.5)

s = content(f'附錄 {appendix("src")}｜資料來源', '本提案引用的外部研究、法規與內部研究。')
SRCS = [('Stack Overflow，2025 Developer Survey 新聞稿（84%）', 'https://stackoverflow.co/company/press/archive/stack-overflow-2025-developer-survey/'),
        ('Google Cloud，Announcing the 2024 DORA report（−1.5%、−7.2%）', 'https://cloud.google.com/blog/products/devops-sre/announcing-the-2024-dora-report'),
        ('Google Cloud，Announcing the 2025 DORA report（約 90%、速度與穩定度）', 'https://cloud.google.com/blog/products/ai-machine-learning/announcing-the-2025-dora-report'),
        ('DORA，The ROI of AI-assisted Software Development，v.2026.1（放大器、J 曲線、既有系統效益、回收期）', 'https://services.google.com/fh/files/misc/dora-roi-of-ai-assisted-software-development-2026.pdf'),
        ('InfoQ，DORA 2025 State of AI-assisted Software Development 整理', 'https://www.infoq.com/news/2025/09/dora-state-of-ai-in-dev-2025'),
        ('Yahoo Finance，JPMorgan 工程師使用寫程式助手效率提升至多 20%，2025-03', 'https://finance.yahoo.com/news/jpmorgan-engineers-efficiency-jumps-much-190410471.html'),
        ('Deloitte Insights，AI and bank software development，2025（Citi 3 萬名開發者）', 'https://www.deloitte.com/global/en/insights/industry/financial-services/financial-services-industry-predictions/2025/ai-and-bank-software-development.html'),
        ('Claims Journal，PwC 保險高階主管調查（57%），2025-11-17', 'https://amp.claimsjournal.com/news/national/2025/11/17/334094.htm'),
        ('McKinsey，Technology Trends Outlook 2026（代理式軟體開發列為新趨勢）', 'https://www.mckinsey.com/capabilities/tech-and-ai/our-insights/the-top-trends-in-tech'),
        ('Allen & Gledhill，MAS 發布 AI 模型風險管理資訊文件（2024-12-05）', 'https://www.allenandgledhill.com/sg/publication/articles/29565/mas-publishes-information-paper-on-ai-model-risk-management'),
        ('理律法律事務所，保險業作業委外新辦法（2025-12-31 施行）', 'https://www.leeandli.com/TW/Newsletters/7525.htm'),
        ('壽險／產險公會，保險業運用人工智慧系統自律規範（2025-04-10 備查）', 'https://law.lia-roc.org.tw/Law/GetFile/0000394231'),
        ('金管會，金融業運用人工智慧（AI）指引，2024-06', 'https://www.fsc.gov.tw/websitedowndoc?file=chfsc/202408231741530.pdf'),
        (f'研究一：{V3}；研究二：AI 原生 SDLC 說明，整理自 {PB}', None)]
for i, (lab, url) in enumerate(SRCS):
    seg = (lab, {'link': url}) if url else (lab, {})
    tb(s, 0.45, 1.45 + i * 0.355, 12.45, 0.33, [[(f'{i + 1}. ', {'bold': True, 'color': OR}), seg]], size=10, fill=LG if i % 2 == 0 else WHITE, anchor=MSO_ANCHOR.MIDDLE, margin=0.08)
note(s, '「−1.5%」「−7.2%」為 AI 採用率每增加 25% 的模型估計；JPMorgan 數字為公司自述。', y=6.5)

# ───────── 代換頁碼、附錄代號與議程頁碼範圍 ─────────
total = len(p.slides)
RANGE = {}
for i, (num, start) in enumerate(SEC):
    end = SEC[i + 1][1] - 1 if i + 1 < len(SEC) else total
    RANGE[num] = f'{start}–{end}'

def resolve(m):
    kind, key = m.group(1), m.group(2)
    return str(PAGE[key]) if kind == 'p' else APP[key] if kind == 'app' else RANGE[key]

for sl in p.slides:
    for sh in sl.shapes:
        if not sh.has_text_frame: continue
        for para in sh.text_frame.paragraphs:
            for r in para.runs:
                if '⟦' in r.text: r.text = re.sub(r'⟦(p|app|r):(\w+)⟧', resolve, r.text)

p.save(OUT)
print(OUT, total, 'slides', 'internal' if INTERNAL else 'client')
