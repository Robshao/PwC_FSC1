# 把 AI 原生開發做成產品：做得快，也管得住（給合夥人向保險業客戶提案的簡報；v4 起改用此名稱，含講者備註）
# 章節沿用四個面向（面向誰、解決什麼問題、帶來什麼價值、可擴充與可複製）；外部證據經查核後引用原始出處；
# IT 現場示範為主要實證，示範畫面待 IT 團隊提供；不含任何個人姓名，同業公司名稱只在附錄
# 版面：內容頁用範本「3__Pattern header」（灰底頁首＋左側橘條），標題下一句結論；字型 Georgia＋微軟正黑體
# 用法：python3 build_sam_proposal.py <範本.pptx> <輸出.pptx> [--client] [--demo <示範素材資料夾>]
#   示範素材由 build_it_demo_assets.py 產生；沒有帶 --demo 時，現場示範頁保留待填框
#   預設為合夥人審閱版（含「請 Sam 決定」內部頁與研究整合附錄）；--client 產生客戶版，不含內部頁
import sys, os, copy, math, re
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

TPL, OUT = sys.argv[1], sys.argv[2]
INTERNAL = '--client' not in sys.argv[3:]
DEMO = sys.argv[sys.argv.index('--demo') + 1] if '--demo' in sys.argv else None
OR = RGBColor(0xFD, 0x51, 0x08); OR2 = RGBColor(0xFE, 0x7C, 0x39); TINT = RGBColor(0xFF, 0xE8, 0xD4)
# 範本佈景主題色：tx1 黑、lt2 淺灰、accent6 線條灰、accent4 中灰；範本不用純黑色塊，深色塊用深灰
INK = RGBColor(0x00, 0x00, 0x00); GREY = RGBColor(0x59, 0x59, 0x59); LG = RGBColor(0xEB, 0xEB, 0xEB); LINE = RGBColor(0xCB, 0xD1, 0xD6)
MID = RGBColor(0xA1, 0xA8, 0xB3); WHITE = RGBColor(0xFF, 0xFF, 0xFF); DARK = RGBColor(0x59, 0x59, 0x59); RED = RGBColor(0xDE, 0x2A, 0x18)
# 範本提案頁字型：英數 Georgia、中文微軟正黑體
FONT = 'Georgia'; HEAD_FONT = 'Georgia'; EA_FONT = '微軟正黑體'
SYMBOLS = set('●◐○✓✗→↑↓①②③④⑤⑥⑦⚠')
PROJECT = 'PwC FSC1｜AI 原生開發：做得快，也管得住'
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

def pic(s, path, x, y, w, h):
    # 圖片等比縮放後置中放進方框（方框以內容頁座標表示）
    from PIL import Image
    iw, ih = Image.open(path).size
    bx, by, bw, bh = x, Y(y), w, H(h)
    k = min(bw / iw, bh / ih); pw, ph = iw * k, ih * k
    return s.shapes.add_picture(path, Inches(bx + (bw - pw) / 2), Inches(by + (bh - ph) / 2), Inches(pw), Inches(ph))

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
set_ph(s, 0, '把 AI 原生開發做成產品：\n做得快，也管得住', 28)
set_ph(s, 1, '從框架到實務：可重複的流程、內建治理、關鍵決定由人做\nPwC FSC1｜給保險業的提案｜2026 年 10 月', 12)

# ───────── 內部：請 Sam 決定（只在合夥人審閱版） ─────────
if INTERNAL:
    CUR[0] = '內部審閱'
    s = content('內部｜請 Sam 決定：是否以本提案接觸首批保險客戶',
                '建議：同意以「診斷＋一個用例試點」接觸 1–2 家保險公司；第 ⟦p:demo⟧ 頁示範素材確認可對外後即出門。', tags=('內部',))
    mark('sam')
    COLS = [('請您決定', ['① 是否同意以本提案接觸首批客戶（候選名單：' + TODO + '）',
                         '② 附錄 ⟦app:peers⟧ 對客戶是否保留同業公司名稱',
                         '③ 請 IT 確認第 ⟦p:demo⟧ 頁示範素材可對外使用，並補一張合併核准紀錄截圖' if DEMO else '③ 指派 IT 同仁完成第 ⟦p:demo⟧ 頁示範畫面（目標日期：' + TODO + '）',
                         '④ 試點帶隊人：' + TODO]),
            ('推案前要先確認', ['・獨立性：候選客戶若為審計客戶，先經 R&Q 確認（尚未檢查）',
                              '・客戶關係：先取得負責該客戶的合夥人同意',
                              '・責任範圍：AI 產出的責任歸屬寫進合約，待 Legal 確認',
                              '・駐點外包廠商：AI 會減少人天，試點前先和客戶 IT 談好廠商角色']),
            ('規模與資源', ['・診斷約 4–6 週＋試點約 6–8 週（第 ⟦p:offer⟧ 頁）',
                           '・人力配置與報價：' + TODO,
                           '・團隊分工與決定權見第 ⟦p:team⟧ 頁'])]
    for i, (h, items) in enumerate(COLS):
        x = 0.45 + i * 4.18
        tb(s, x, 1.45, 4.05, 0.5, h, size=13.5, bold=True, color=WHITE, fill=OR if i == 0 else DARK, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
        tb(s, x, 2.0, 4.05, 2.45, [[(t, {})] for t in items], size=11.5, fill=TINT if i == 0 else LG, margin=0.14, spacing=1.05)
    OPT = [['A　診斷＋一個用例試點（建議）', '金額小、客戶容易說好；試點同時產出對外實績與工具包第一版', '採用'],
           ['B　直接賣工具包，客戶自己做', '我們還沒有對外實績，客戶難以信任方法；維護成本也較高', '實績累積後再推'],
           ['C　等示範完成再決定', '風險最低，但獨立性與客戶關係的確認可以先並行', '不等：先做確認，示範完成即出門']]
    grid(s, 0.45, 4.65, [3.9, 6.25, 2.3], ['考慮過的做法', '評估', '結論'], OPT, 0.6, size=11.5, hot=(0,))

    # ───────── 內部：為什麼要做成產品 ─────────
    s = content('內部｜為什麼要做成產品：開發變便宜，治理才是能賣的能力',
                '做出來不再是門檻；可重複交付、內建治理與人工把關的做法，才是客戶願意付費、我們能規模化的東西。', tags=('內部',))
    mark('product')
    PCOL = [('市場在發生什麼', ['・DORA 2026 引 Stanford AI Index：模型使用成本 2022-11 到 2024-10 降為 1/280；導入的主要負擔轉到治理：審查負擔、流程調整、人才培訓',
                              '・DORA 2026：只買授權不保證回報，回報來自內部平台、清楚流程與團隊協作',
                              '・同業已把服務做成產品：加速器、工具組、導入藍圖（附錄 ⟦app:peers⟧）']),
            ('對我們的意義', ['・只接專案：每案從頭來，營收只能靠加人成長',
                            '・做成產品：範本、規則、關卡與法規清單每案沿用，越做越快',
                            '・治理與人工把關是產品的核心，不是附加服務']),
            ('我們已經有的資產', ['・六份規格範本與任務格式（驗收場景、最多改 5 個檔案）' + ('，見第 ⟦p:demo⟧ 頁' if DEMO else ''),
                                '・品質閘門與 AI 任務四象限（第 ⟦p:tasks⟧ 頁）',
                                '・保險業法規檢查清單（附錄 ⟦app:rules⟧）'])]
    for i, (h, items) in enumerate(PCOL):
        x = 0.45 + i * 4.18
        tb(s, x, 1.45, 4.05, 0.5, h, size=13.5, bold=True, color=WHITE, fill=OR if i < 2 else DARK, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
        tb(s, x, 2.0, 4.05, 3.25, [[(t, {})] for t in items], size=12, fill=TINT if i == 1 else LG, margin=0.16, spacing=1.1)
    callout(s, 5.5, '做法：', '先賣「診斷＋一個用例試點」，每個試點至少萃取一個可重用的範本或規則；資產成熟、有實績後，再推客戶自己做的工具包（第 ⟦p:sam⟧ 頁做法 B）。', h=0.85, size=12)
    note(s, '成本數字為 DORA 2026 引述 Stanford AI Index；「只接專案靠加人成長」整理自內部產品化討論（MIT Sloan Management Review，2024）。', y=6.55)

    # ───────── 內部：FSC1 團隊分工 ─────────
    s = content('內部｜FSC1 團隊分工：BA 管業務與法規，IT 管實作與證明',
                '客戶買的是兩種別人難以複製的能力；決定權先講清楚，示範與報價才不會卡住。', tags=('內部',))
    mark('team')
    TEAM = [['BA', '懂保險業務與法規，讓 AI 一開始就拿到對的需求', '業務單位、法遵與風控', '主導需求分析與「接受需求」紀錄；提出法規檢查清單（法遵審定）；把工程語言翻成管理層聽得懂的話'],
            ['Rex（IT）', '這套方法的設計者', 'IT 主管、架構師', 'v3 流程與五個關卡；決定示範用例與四象限分類；試點中的技術決定'],
            ['Kenny（IT）', '每天實際在用的人', '工程師', '現場示範與實績表；建置範本、規則與關卡；帶客戶工程師實作'],
            ['Sam', '決策者，也是開門的人', '管理層', '決定推案、對外承諾、報價與人力'],
            ['Luke', '定價與人力（職稱待確認）', '採購與財務', '提出報價與人力配置'],
            ['R&Q、法遵', '把關', '—', '決定獨立性；審定法規檢查清單']]
    grid(s, 0.45, 1.45, [1.55, 3.75, 2.2, 4.95], ['角色', '對客戶的定位', '對口的客戶角色', '負責與決定'], TEAM, 0.47, size=10.5, hot=(0,), head_h=0.36)
    QA = [['跟國際顧問的方案比，我們贏在哪？', '在地法規與可稽核關卡綁在一起，BA 與 IT 同一組人，從客戶的一個真實需求開始（第 ⟦p:why⟧ 頁）'],
          ['AI 寫的程式出事，責任算誰的？', '能不能上線由客戶指定的人授權，誰核准都查得到（第 ⟦p:compliance⟧ 頁）；合約責任範圍待 Legal 確認'],
          ['怎麼證明有效？', '試點前以基準訂好合格線，達標才擴大，未達標就停（第 ⟦p:offer⟧ 頁）']]
    grid(s, 0.45, 4.8, [3.6, 8.85], ['客戶可能會問', '建議回答'], QA, 0.47, size=10.5, head_h=0.36)
    note(s, '交接點：IT→BA 的示範截圖與數字未交，不約推案會議；BA→IT 的需求說明沒有核准紀錄，不進 SA&D。名字與分工依內部會議資料整理，待確認。', y=6.68)

# ───────── 一頁看懂 ─────────
CUR[0] = '提案摘要'
s = content('一頁看懂：做得快也管得住，從一個用例開始',
            '先做 4–6 週診斷與 6–8 週單一用例試點；達標後，範本、規則與關卡沿用到下一個系統。')
WHY = [('為什麼是現在', '委外新辦法與 AI 自律規範已上路，AI 產出的最終責任在保險公司；開發流程要拿得出「誰要求、AI 做了什麼、誰核准」的紀錄。', '第 ⟦p:urgency⟧、⟦p:compliance⟧ 頁'),
       ('為什麼只買工具不夠', 'DORA 2025：AI 讓交付變快，但交付穩定度仍在下降。效益取決於需求、審查、上線與控制方式有沒有一起改。', '第 ⟦p:tools⟧ 頁'),
       ('為什麼找我們', 'FSC1 自己已經用這套流程開發；懂保險業務與法規的 BA 和實作的 IT 同一組人，從貴公司的一個真實需求開始。', '第 ⟦p:demo⟧、⟦p:why⟧ 頁')]
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
s = content('議程', '為什麼現在要做、怎麼落地、帶來什麼價值，以及怎麼做成可沿用的做法。', layout='Agenda')
AG = [('1', '面向誰', '保險公司為什麼現在要做、優先服務誰', '01'),
      ('2', '解決什麼問題', '全球趨勢與只買工具為什麼不夠、我們的解法、AI 分工、現場示範', '02'),
      ('3', '帶來什麼價值', '價值與衡量、給法遵與風控、你買到什麼、為什麼選我們', '03'),
      ('4', '可擴充、可複製', '做成可沿用的做法：逐步擴大、如何開始', '04'),
      ('5', '附錄', '示範情境、延伸方式、同業參考、法規清單、' + ('示範追溯、' if DEMO else '') + '試點風險、' + ('研究整合、' if INTERNAL else '') + '資料來源', 'A')]
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

if DEMO:
    # IT 團隊交付的示範：再保險系統主檔模組（TASK-000013 任務包與錄影），文件為原文節錄
    s = content('現場示範：再保險系統主檔模組，從業務規則到上線畫面',
                '以 FSC1 IT 團隊的實作為例：業務規則、規格、任務與測試一路相連，每一步都留下文件與紀錄。', tags=('待確認',))
    mark('demo')
    FR = [('需求說明', '7 條業務規則，含制裁主管覆核'), ('規格與規範', '6 份規格，各有審閱週期'), ('實作計畫', '列明驗收場景，最多改 5 個檔案'),
          ('程式與測試', '10 個測試案例，覆蓋率 > 90%'), ('核准與紀錄', '任務回溯規則代碼，寫入審計')]
    W = 2.41
    for i, (h, t) in enumerate(FR):
        x = 0.45 + i * (W + 0.1)
        tb(s, x, 1.45, W, 0.42, f'{i + 1}　{h}', size=12, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.08)
        tb(s, x, 1.9, W, 1.85, '', fill=WHITE, line=LINE)
        pic(s, os.path.join(DEMO, f'step{i + 1}.png'), x + 0.03, 1.93, W - 0.06, 1.79)
        tb(s, x, 3.8, W, 0.55, t, size=10.5, fill=LG, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, margin=0.06)
    SCR = [('currencies', '結算幣別選單：清單'), ('countries', '國家選單：新增'), ('regions', '區域選單：新增與提醒')]
    for i, (k, cap) in enumerate(SCR):
        x = 0.45 + i * 4.18
        tb(s, x, 4.5, 4.05, 1.62, '', fill=WHITE, line=LINE)
        pic(s, os.path.join(DEMO, f'screen_{k}.png'), x + 0.03, 4.53, 3.99, 1.56)
        tb(s, x, 6.14, 4.05, 0.3, cap, size=10, bold=True, color=GREY, align=PP_ALIGN.CENTER, margin=0.02)
    note(s, '文件為 TASK-000013「金融機構主檔」任務包原文節錄；畫面取自同一系統的示範錄影，使用測試資料。追溯表見附錄 ⟦app:trace⟧。', y=6.55)
else:
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

s = content('給法遵與風控：每一個 AI 產出，都查得到誰要求、誰核准',
            '金檢與內稽會問的五個問題，在流程中都有對應的做法與紀錄；條文對照見附錄 ⟦app:rules⟧' + ('，實例見附錄 ⟦app:trace⟧。' if DEMO else '。'))
mark('compliance')
ROWS = [['AI 工具有沒有納入委外管理與 AI 清冊？', '診斷時判斷是否屬委外並登錄 AI 清冊；高風險用途加審', '委外判斷紀錄、AI 清冊'],
        ['客戶個資有沒有進到 AI？', '資料分級；集中設定禁止讀取機密檔與金鑰；試點用假資料', '設定檔與存取紀錄'],
        ['這個改動是誰核准上線的？', '五個關卡各有指定核准人；寫程式的 AI 不能核准自己的程式', '核准紀錄（人、時間、版本）'],
        ['出事能不能退回？', '退版流程事先演練；AI 建議退版，服務負責人核准後執行', '演練與事件紀錄'],
        ['AI 做了什麼，事後查得到嗎？', '需求說明、規格、實作計畫與測試結果版本化保存', '完整的文件軌跡']]
grid(s, 0.45, 1.45, [3.6, 5.6, 3.25], ['金檢與內稽會問', '流程中的做法', '留下的紀錄'], ROWS, 0.74, size=12)
callout(s, 5.75, '對法遵的意義：', 'AI 使用從「難以稽核的個人行為」變成「有紀錄的流程」。法規檢查清單為草稿，診斷時與貴公司法遵逐條確認。', h=0.75)

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
tb(s, 9.05, 2.1, 3.85, 2.85, [[('・' + k, {})] for k in KPI] + [[('・核准紀錄完整（流程內建的必要條件）', {})], [('・合格線：以診斷期量到的基準，試點前約定', {})], [('', {'size': 5})],
                             [('達標：', {'bold': True}), ('擴大到下一個用例', {})], [('未達標：', {'bold': True}), ('檢討後再試或停止', {})]], size=11.5, fill=LG, margin=0.16)
callout(s, 5.1, '貴公司需配合：', 'IT 與業務窗口、一個真實需求與團隊、測試環境、AI 工具的合約或內部核准。', h=0.6)
callout(s, 5.8, '我們的團隊：', '懂保險業務與法規的 BA 主導需求與法規檢查，IT 主導實作與示範；同一組人從診斷做到試點。', h=0.6)
note(s, '期程依範圍確認。擴大階段、工具包授權與駐點支援可依需求討論。試點風險與因應見附錄 ⟦app:risk⟧。', y=6.5)

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

if DEMO:
    s = content(f'附錄 {appendix("trace")}｜示範案例：每條業務規則都追溯到測試與任務',
                '以再保險系統「金融機構主檔」為例；稽核時從任一條規則，都找得到對應的使用案例、測試與程式範圍。')
    ROWS = [['BR-FI-01　代碼唯一性', 'UC-FI-02', 'TC-FI-03', 'TASK-000013-001 場景 3'],
            ['BR-FI-02　引用後鎖定唯讀', 'UC-FI-04', '—', 'TASK-000013-005 場景 2'],
            ['BR-FI-03　引用防刪與軟作廢', 'UC-FI-04、UC-FI-06', 'TC-FI-07、TC-FI-08', 'TASK-000013-005 場景 2'],
            ['BR-FI-04　SWIFT/BIC 格式檢核', 'UC-FI-02', 'TC-FI-04', 'TASK-000013-001 場景 4'],
            ['BR-FI-05　作廢資料隔離與追溯', 'UC-FI-06', '—', '任務包未附'],
            ['BR-FI-06　制裁連動與主管覆核', 'UC-FI-03', 'TC-FI-05、TC-FI-06', '任務包未附'],
            ['BR-FI-07　順序自動計算', 'UC-FI-02', 'TC-FI-02', 'TASK-000013-001 場景 2']]
    grid(s, 0.45, 1.45, [4.0, 2.6, 2.6, 3.25], ['業務規則', '使用案例', '測試案例', '開發任務'], ROWS, 0.56, size=11, hot=(5,))
    callout(s, 5.85, '追溯表的用處：', '缺口一眼看得出來：BR-FI-02、BR-FI-05 目前沒有對應的測試案例；任務包只附 001 與 005 兩個任務。', h=0.62)
    note(s, '整理自 TASK-000013 任務包（SA、使用案例、QA 規格與任務文件）；任務與規則的對應依任務內容判讀，待 IT 團隊確認。', y=6.6)

s = content(f'附錄 {appendix("risk")}｜試點風險與因應',
            '試點前就講好可能出問題的地方、怎麼提早發現，以及誰負責處理。', tags=('建議',))
ROWS = [['用例效益偏低', '用例落在複雜的既有系統，重複性工作少', '診斷時優先挑「常發生、判斷少」工作多的需求', '調整範圍或換用例，不硬撐到底', 'PwC（BA、IT）＋貴公司 IT 主管'],
        ['初期下滑被誤判（J 曲線）', '前幾週需求定稿與審查時間變長', '試點前講好預期的下滑與觀察期間', '依約定期間判斷，不在低谷時中止', '貴公司管理層＋PwC 帶隊人'],
        ['個資或機密進入 AI', '設定或存取紀錄出現例外', '資料分級；試點用假資料；集中設定禁止讀取機密檔', '立即停用該工具、通報並檢討設定', '貴公司資安＋PwC IT'],
        ['AI 工具未獲內部核准，時程延誤', '診斷第 2 週仍未完成委外判斷與清冊登錄', '診斷第一週就啟動委外判斷與清冊登錄', '先用已核准的工具試點，或調整時程', '貴公司法遵、IT'],
        ['駐點團隊配合度低', '試點任務排不進工作、出席率低', '試點前與 IT 主管談好駐點團隊的角色與工時', '改由貴公司自有人員或 PwC 補位', '貴公司 IT 主管']]
grid(s, 0.45, 1.45, [2.3, 2.6, 2.85, 2.65, 2.05], ['風險', '提早發現的訊號', '事前降低', '發生時的應變', '負責'], ROWS, 0.85, size=10.5)
note(s, '依 DORA 2026（J 曲線、既有系統效益）與試點規劃整理；正式清單於診斷時與貴公司共同確認。', y=6.35)

if INTERNAL and DEMO:
    s = content(f'附錄 {appendix("video")}｜示範錄影：再保險系統主檔模組',
                '內部參考：IT 團隊的三段示範錄影（各約 5 秒），點擊畫面播放；對外以現場示範為主。', tags=('內部',))
    for i, (k, cap) in enumerate([('countries', '國家選單'), ('regions', '區域選單'), ('currencies', '結算幣別選單')]):
        x = 0.45 + i * 4.18
        w, h = 4.05, 4.05 * 450 / 800
        s.shapes.add_movie(os.path.join(DEMO, f'video_{k}.mp4'), Inches(x), Inches(Y(1.6)), Inches(w), Inches(h),
                           poster_frame_image=os.path.join(DEMO, f'screen_{k}.png'), mime_type='video/mp4')
        tb(s, x, 1.6 + h / K + 0.08, w, 0.35, cap, size=12, bold=True, color=GREY, align=PP_ALIGN.CENTER, margin=0.02)
    note(s, '錄影使用測試資料；對外使用前請 IT 確認不含客戶資料或客戶專屬設計。', y=6.3)

if INTERNAL:
    s = content(f'附錄 {appendix("research")}｜研究整合：兩份研究互補，衝突處的取捨',
                '內部參考：兩份研究的差異，以及本提案採用哪一邊、為什麼。提交客戶前移除本頁。', tags=('內部',))
    ROWS = [['流程階段', '五階段：需求分析、SA&D、開發、測試、上版／交付', '六階段：規劃、設計、開發、測試、上線、維運', '採研究一的名稱，加上維運：保險業 IT 熟悉這些用語；維運讓流程形成循環'],
            ['控制方式', 'Prompt 撰寫原則＋IT 人工 Review', '技能、鉤子、集中管理設定', '分層：Prompt 為第一層，往規則與自動關卡加強；受監理行業需要可強制的控制'],
            ['關卡', '每階段文件確認後進下一階段', '五個關鍵決定由指定的人核准', '保留研究一的檢查點，補上指定核准人與紀錄'],
            ['效益說法', '未提數字', '手冊預期「數週到數小時」', '不承諾數字，以試點實測；外部證據改用 DORA 等獨立研究'],
            ['實作證據', '各階段案例頁待補', '實績頁待 IT 團隊填寫', ('以 IT 示範補上：再保險系統主檔模組（第 ⟦p:demo⟧ 頁）' if DEMO else '以 IT 現場示範補上（第 ⟦p:demo⟧ 頁）')],
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

# ───────── 講者備註 ─────────
# 依頁面標題比對（先比對較具體的鍵）；第二段只在合夥人審閱版加入，客戶版不含內部提醒
NOTES = [
    ('把 AI 原生開發做成產品', """【這頁的重點】30 秒開場，先講結論。
【講法】「AI 寫程式的工具，各位的團隊大概都已經在用了。今天不談工具，談三件事：第一，這套 AI 原生的開發做法在實務上怎麼跑；第二，怎麼把它做成每個案子都能沿用的東西，而不是每次從頭來；第三，做得快的同時，治理和人工把關放在哪裡。結論先講：從一個用例開始，做得快，也管得住。」
【轉場】「我先用一頁把整件事講完，後面再展開。」""",
     """【內部提醒】向 Sam 說明有兩個版本：審閱版多三頁內部頁（請您決定、為什麼要做成產品、團隊分工）與兩個內部附錄（示範錄影、研究整合）；客戶版不含。"""),
    ('內部｜請 Sam 決定', """【這頁的重點】請 Sam 今天做四個決定。
【講法】先講建議：同意以「診斷＋一個用例試點」接觸 1–2 家保險公司。接著逐項講「請您決定」：首批客戶、附錄同業名稱是否保留、示範素材對外確認、試點帶隊人。中間欄是推案前要先確認的事，獨立性與客戶關係最可能卡住，所以要先做。下方是考慮過的三種做法：不直接賣工具包，因為還沒有對外實績；不等示範全部完成，因為獨立性與客戶關係的確認可以並行。
【可能被問】「第一家找誰？」→ 候選名單請您決定，原則是先找非審計客戶。「報價多少？」→ 人力與報價待提出，採固定範圍、固定價格。
【轉場】「為什麼封面寫『做成產品』，下一頁說明。」""", ''),
    ('內部｜為什麼要做成產品', """【這頁的重點】AI 讓「做出來」變便宜，能賣、能規模化的是可重複的治理做法。
【講法】「封面寫『做成產品』，原因在左欄：DORA 2026 引 Stanford 的數據，模型使用成本兩年內降到原本的 1/280，導入的主要負擔變成治理：審查負擔、流程調整、人才培訓。報告也說，只買授權不保證回報。客戶缺的不是寫程式的能力，而是一套可重複、管得住的做法。中間欄是對我們的意義：如果每案從頭來，營收只能靠加人；把範本、規則、關卡和法規清單做成每案沿用的資產，才會越做越快，而治理和人工把關就是這個產品的核心。右欄是我們已經有的資產，來自 IT 團隊的實作。」
【做法】「不是現在就賣工具包。先賣診斷加試點，每個試點萃取一個可重用的資產；有實績後，再推客戶自己做的版本。」
【可能被問】「跟 Deloitte 的加速器有什麼不同？」→ 他們偏預建代理與藍圖；我們把在地法規、可稽核關卡和人工把關綁在一起（附錄同業參考）。""", ''),
    ('內部｜FSC1 團隊分工', """【這頁的重點】客戶買的是兩種別人難以複製的能力：BA 的業務與法規、IT 的實作與證明。
【講法】逐列講定位與對口：BA 對口業務單位與法遵，主導需求分析並提出法規檢查清單；Rex 是方法設計者，對口客戶 IT 主管；Kenny 是每天實際在用的人，對口工程師；Sam 決定推案與報價；R&Q 與法遵把關。下方是客戶最常問的三題與建議回答。頁底是兩個交接點：IT 的示範素材沒交，不約推案會議；需求說明沒有核准紀錄，不進入設計。
【可能被問】「試點誰帶隊？」→ 請 Sam 指定，客戶需要固定窗口。""", ''),
    ('一頁看懂', """【這頁的重點】如果對方只看一頁，就是這頁：建議、三個理由、請對方決定的事、不做的代價。
【講法】「我們的建議是：先做 4 到 6 週的準備度診斷，再做 6 到 8 週、一個用例的試點；達到事先講好的標準才擴大，而且範本、規則與關卡會沿用到下一個系統。理由有三個。第一，為什麼是現在：委外辦法和 AI 自律規範已經上路，AI 產出的最終責任在保險公司，流程要拿得出誰要求、AI 做了什麼、誰核准的紀錄。第二，為什麼只買工具不夠：DORA 2025 發現，AI 讓交付變快，但交付穩定度在下降。第三，為什麼找我們：我們自己已經這樣在做，懂保險業務與法規的 BA 和實作的 IT 是同一組人。」最後指向下方：今天想請貴公司決定的，是安排一場現場示範並選出 1–2 個候選用例；如果先不做，工具照樣擴散，但紀錄要等到金檢時再補。
【轉場】「接下來四章依序展開。」
【可能被問】「試點要多少錢？」→ 固定範圍、固定價格，另行報價。""", ''),
    ('議程', """【這頁的重點】說明結構與時間安排。
【講法】四章加附錄：面向誰、解決什麼問題、帶來什麼價值、可擴充與可複製。時間不夠時，優先講第 02 章的解法、AI 分工與現場示範，以及第 03 章的「你買到什麼」；附錄依提問使用。""", ''),
    ('面向誰\nWho', """【轉場】「先講為什麼保險公司現在就要處理這件事，以及我們優先服務誰。」""", ''),
    ('保險公司現在就要', """【這頁的重點】責任已經落在保險公司，紀錄必須設計進流程。
【講法】三張卡由左到右：第一，委外新辦法 2025 年底施行，AI 自律規範要求保險公司對 AI 結果負最終責任。第二，金管會 AI 指引要求治理及問責、透明性；新加坡金管局檢查銀行後，發現多數已建 AI 清冊、高風險 AI 要獨立驗證，這是區域監理的方向。第三，IFRS 17 與 TW-ICS 上路後，相關系統要跟著調整，開發需求集中。意涵：紀錄不能事後補，要在流程裡設計好；新制帶來的系統調整，正好是一起導入的時機。
【可能被問】「新加坡的例子是銀行？」→ 是，當作區域監理方向的參考；台灣的直接依據是金管會指引與公會自律規範。
【注意】不要說保險業「比其他產業更急」，證據不支持。""", ''),
    ('優先客群', """【這頁的重點】先服務有自建或駐點開發團隊的壽險與產險公司。
【講法】表格是採購中的四種角色：IT 主管決策與付費，業務單位提出需求，法遵、風控與資安握有否決權，管理層核准預算；每個角色各有在意的問題與我們的答案。這是依經驗推估的決策鏈，會在診斷訪談中確認。下方是優先客群的三個條件。
【可能被問】「開發全部外包的公司呢？」→ 可以做，但要先談好外包廠商在試點中的角色（見附錄試點風險）。""", ''),
    ('解決什麼問題\nProblems', """【轉場】「接下來講問題：工具已經普及，為什麼交付沒有跟著變好，以及我們怎麼解。」""", ''),
    ('AI 已普及', """【這頁的重點】工具已普及，但效益取決於流程與治理。
【講法】上排五個數字說明普及程度：開發者 84%、技術團隊約九成在用 AI，JPMorgan、Citi 已大規模導入，保險業 57% 的高階主管把生成式 AI 列為首要投資。下排是三個發現：DORA 2025，速度轉為正相關，但穩定度仍是負相關；DORA 2026 引述 Stanford 研究，新開發的簡單任務可提升 35–40%，但複雜的既有系統常在 10% 以下，而保險公司的核心系統多屬後者；DORA 2026 的結論是，AI 是放大器，回報來自內部平台、清楚的流程與團隊協作，不是工具本身。所以導入 AI 時，需求、審查、上線與控制方式要一起改。
【可能被問】「JPMorgan 不是提升 10–20% 嗎？」→ 那是個人生產力，而且是公司自述；DORA 看的是整體交付，兩者不衝突。
【轉場】「那要怎麼一起改？看我們的解法。」""",
     """【內部提醒】57% 目前轉引 Claims Journal，正式對外前改引 PwC 原始報告。"""),
    ('我們的解法', """【這頁的重點】把開發流程串成循環，用三個機制守住；關鍵決定由人做。
【講法】左邊是循環：需求分析、設計、開發、測試、上線、維運，維運發現的問題回到下一個需求。右邊是三個機制：交接靠檔案，每一步交出文件給下一步；規則自動執行，資安、個資、委外規範寫成 AI 自動套用的規則；關卡有人簽，接受需求、確認規格、核准計畫、核准合併、授權上線，五個決定由貴公司指定的人做。原則只有一句：AI 可以一路做到正式環境的門口，但不能自己開門。頁底：DORA 2026 建議的不可略過檢查點、提交前檢查、小批次與版本控制，和這三個機制一致。
【可能被問】「要換掉我們現有的 Jira 或變更管理嗎？」→ 不用，關卡與紀錄接在既有工具上，診斷時確認做法。「綁哪個 AI 工具？」→ 不綁，依貴公司政策與合約選定。
【轉場】「人到底在什麼時候介入？下一頁。」""", ''),
    ('AI 分工', """【這頁的重點】回答「人在什麼時候介入」：依多常發生、需要多少判斷，把工作分四類。
【講法】左上，常發生、判斷少，例如測試案例、文件、Release Note：AI 先做、人確認，這類最先交給 AI，也最先做成範本。右上，常發生、判斷多，例如程式碼、API 規格：AI 起草、人逐項審查。左下，少發生、判斷少：做成範本重複使用。右下，少發生、判斷多，例如接受需求、核准合併、授權上線：人決定，AI 只提供資料，這就是前一頁的五個關卡。右側的意義：第一個試點，優先挑左上角工作多的需求，最快看到效果。
【轉場】「說了這麼多，實際長什麼樣子？看示範。」""", ''),
    ('現場示範', """【這頁的重點】不是概念，是我們 IT 團隊在再保險系統上實際做出來的。
【講法】「五個框由左到右是同一個功能的五個步驟。第一，需求說明：金融機構主檔有 7 條業務規則，包括制裁國家的機構建檔要主管覆核，這是洗錢防制的規則。第二，依範本產出六份規格，從系統分析、資料庫、API、畫面到測試，每份都有審閱週期。第三，實作計畫：每個任務寫明驗收場景，而且最多只能改 5 個檔案，修改範圍事先鎖定。第四，10 個測試案例，覆蓋率要超過 90% 才能過閘。第五，每個任務都回溯到規則代碼，制裁建檔會寫入審計紀錄。下方是同一系統的實際畫面。」最後邀請：「現場示範可以用貴公司提供的一個需求，走一遍。」
【可能被問】「這是哪家客戶的系統？」→ 依 IT 確認的說法回答。「追溯怎麼查？」→ 翻到附錄追溯表。""",
     """【內部提醒】(1) 素材對外前，請 IT 確認不含客戶資料或客戶專屬設計；(2) 錄影拍的是國家、區域、幣別，規格是金融機構，最好補一段金融機構選單的錄影；(3) 第 5 步還缺「合併核准」截圖；(4) IT 確認前，不要說規格是 AI 產出的。"""),
    ('帶來什麼價值\nValue', """【轉場】「看完怎麼做，接下來講帶來什麼價值、給法遵的答案，以及你買到什麼。」""", ''),
    ('帶來什麼價值：每個痛點', """【這頁的重點】每個痛點對應一個做法、一項價值、一個量得到的指標。
【講法】逐列講：需求來回多輪，對應業務和 AI 一起寫需求說明，看需求定稿時間；審查排隊，對應 AI 先審、人判斷風險，看第一次審查等待時間與返工次數；上線出錯，對應每次改動自行測試，看變更失敗率；法規與稽核壓力，對應規則、關卡與紀錄，看內稽調閱佐證所需時間；交付難預期，對應每階段交出檔案，看交付時間。試點用流程指標驗收；DORA 四項指標需要較長期間，擴大後再追蹤。我們不先承諾數字。頁底的回收期基準：需要大量治理的大型企業導入，12 到 18 個月屬可接受。
【可能被問】「能保證快多少嗎？」→ 不承諾數字，試點前量基準、訂合格線，用實測證明。""", ''),
    ('給法遵與風控', """【這頁的重點】寫給握有否決權的人：金檢與內稽會問的五個問題，流程中都有答案與紀錄。
【講法】逐列講：AI 工具有沒有納入委外管理與 AI 清冊；客戶個資有沒有進到 AI；改動是誰核准上線的；出事能不能退回；AI 做了什麼，事後查不查得到。每一題對應流程中的做法，以及會留下的紀錄。結論：AI 的使用從難以稽核的個人行為，變成有紀錄的流程。條文對照在附錄法規清單，實例在附錄追溯表。
【建議】這頁最好在法遵在場時講，或會前先單獨給法遵看。""", ''),
    ('你買到什麼', """【這頁的重點】先診斷、再試點；成功標準事先講好，達標才擴大。
【講法】左邊，準備度診斷 4–6 週：盤點現況與卡點、做法規檢查、挑第一個用例、訂成功標準；交付瓶頸地圖、法規檢查結果、用例建議與導入藍圖。中間，一個用例的試點 6–8 週：用真實需求走完循環，建立範本、規則與關卡，同時教會各角色；交付完成的試點、範本與規則、教材與成效報告。右邊是成功標準：需求定稿時間、第一次審查等待時間、返工次數；核准紀錄完整是必要條件；合格線用診斷期量到的基準在試點前約定。達標擴大，未達標就檢討後再試或停止。下方是需要貴公司配合的事項，以及我們的團隊組成；試點風險見附錄。
【可能被問】「價格？」→ 固定範圍、固定價格，另行報價。""", ''),
    ('為什麼選我們', """【這頁的重點】已在用的做法，加上在地法規與可稽核關卡。
【講法】四個差異，每個都有證據：已經在用，見現場示範；台灣保險法規內建，見法規檢查清單；懂保險業務與在地系統，用例直接挑理賠、核保、核心系統周邊；可稽核的人工關卡，寫程式的 AI 不能核准自己的程式。和國際方案的差別：國際方案多以預建代理與藍圖為主，我們從貴公司的一個真實需求開始。""",
     """【內部提醒】「懂保險業務」的證據欄仍待補 2–3 個 FSC1 保險業專案；如果再保險系統是 FSC1 自己的專案，可以直接補上。"""),
    ('可擴充、可複製\nScalable', """【轉場】「最後講怎麼擴大，這就是封面說的『做成產品』：每一步都沿用前一步的資產。」""", ''),
    ('從一個用例開始，驗證價值', """【這頁的重點】每一步只承諾下一步；範本與規則沿用，越做越快。
【講法】四個台階：一個用例，然後驗證價值、決定擴大、調整或停止，接著是相鄰用例，只補該系統的規範，最後才是全公司。右側是為什麼這樣做：風險小、成本可控、有數據才擴大、越做越快。最後一點是 J 曲線：DORA 2026 指出，導入初期會有短暫下滑，常見的失敗是在低谷時撤資，所以成功標準與觀察期間要事先講好。
【可能被問】「多久能全公司導入？」→ 依前一階段的成果決定，不預先承諾時程。""", ''),
    ('如何開始', """【這頁的重點】三個具體的下一步，以及需要貴公司提供什麼。
【講法】上方是四個階段與進入下一步的條件。下一步只有三件事：第一，現場示範，用貴公司提供的一個需求；第二，對齊工作坊，挑 1–2 個候選用例；第三，啟動準備度診斷。右邊是需要貴公司提供的：IT 與業務窗口、現有開發流程與工具、候選需求清單、資安個資與法遵規範。
【收尾】「一句話：AI 幫你把每一步做得更快，每一步都留下紀錄；要不要做、能不能上線，永遠由你的人決定。我們想請貴公司先約一場現場示範。」""", ''),
    ('示範情境', """【用途】對方想看一個完整流程的例子時使用。說明每個階段導入前後的差別，以及由誰決定。這是改編自公開範例的示意，不是客戶案例。""", ''),
    ('延伸方式', """【用途】被問「之後能延伸到哪裡」時使用：同一客戶的其他系統、新制專案、AI 治理、其他保險公司、其他金融業，各有沿用的資產與前提。""",
     """【內部提醒】延伸到其他保險公司時，審計客戶要先確認獨立性。"""),
    ('同業產品化參考', """【用途】被問到競爭對手時使用。這些都是業者公開資訊，規模與期程數字為業者自述。""",
     """【內部提醒】對客戶是否保留同業公司名稱，待 Sam 決定。"""),
    ('法規檢查清單', """【用途】法遵追問條文時使用。說明每條規範在開發流程的哪一步落實。這是導入規劃工具，不是法律意見，草稿待法遵確認。""", ''),
    ('示範案例', """【用途】示範之後，法遵或稽核想看證據時使用。從任一條業務規則，都能找到對應的使用案例、測試案例與開發任務。也可以強調：缺口一眼看得出來，例如有兩條規則還沒有對應的測試。""",
     """【內部提醒】任務與規則的對應是依任務內容判讀，待 IT 確認；任務包只附 001 與 005。"""),
    ('試點風險與因應', """【用途】被問「會出什麼問題」時使用。五個風險各有提早發現的訊號、事前降低的做法、發生時的應變與負責人。正式清單於診斷時與貴公司共同確認。""", ''),
    ('示範錄影', """【用途】內部參考。三段錄影各約 5 秒，點擊畫面播放；對外以現場示範為主。錄影使用測試資料。""", ''),
    ('研究整合', """【用途】內部參考：兩份研究的差異，以及本提案採用哪一邊、為什麼。提交客戶前移除。""", ''),
    ('資料來源', """【用途】被問數字出處時使用。−1.5%、−7.2% 是 AI 採用率每增加 25% 的模型估計；JPMorgan 數字為公司自述。""", ''),
]

def speaker(sl, title):
    for key, pub, inner in NOTES:
        if key in title:
            if key == '現場示範' and not DEMO:
                pub, inner = '【這頁的重點】用一個用例走完從需求到上線的五個步驟；畫面由 FSC1 IT 團隊提供。', '【內部提醒】示範畫面尚未放入；帶 --demo 建置即可。'
            text = pub + ('\n\n' + inner if INTERNAL and inner else '')
            sl.notes_slide.notes_text_frame.text = text
            return True
    return False

missing = []
for sl in p.slides:
    t = sl.shapes.title.text if sl.shapes.title is not None else ''
    if t.strip() == '附錄':
        sl.notes_slide.notes_text_frame.text = '【用途】附錄依提問使用，不主動講。'
    elif not speaker(sl, t):
        missing.append(t)
if missing: print('slides without notes:', missing)

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
    if sl.has_notes_slide:
        for para in sl.notes_slide.notes_text_frame.paragraphs:
            for r in para.runs:
                if '⟦' in r.text: r.text = re.sub(r'⟦(p|app|r):(\w+)⟧', resolve, r.text)

p.save(OUT)
print(OUT, total, 'slides', 'internal' if INTERNAL else 'client')
