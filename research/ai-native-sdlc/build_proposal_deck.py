# AI 原生軟體開發：客戶提案簡報（正文 14 頁＋附錄 2 頁）
# 以四個面向（面向誰、解決什麼問題、帶來什麼價值、可擴充與可複製）為主軸；依 Sam 審閱意見修改：
#   不標內部標籤、不點名同業、不使用廠商的圖、對外只推「診斷 → 陪你做試點」、價值以試點成功標準回答
# 內部版（故事線、修改紀錄、待決事項）見 build_proposal_internal.py
# 用法：python3 build_proposal_deck.py <範本.pptx> <輸出.pptx>
import sys, copy, math
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

TPL, OUT = sys.argv[1], sys.argv[2]
OR = RGBColor(0xFD, 0x51, 0x08); OR2 = RGBColor(0xFE, 0x7C, 0x39); TINT = RGBColor(0xFF, 0xE8, 0xD4)
# 範本佈景主題色：tx1 黑、lt2 淺灰、accent6 線條灰、accent4 中灰
INK = RGBColor(0x00, 0x00, 0x00); GREY = RGBColor(0x59, 0x59, 0x59); LG = RGBColor(0xEB, 0xEB, 0xEB); LINE = RGBColor(0xCB, 0xD1, 0xD6)
MID = RGBColor(0xA1, 0xA8, 0xB3); WHITE = RGBColor(0xFF, 0xFF, 0xFF)
# 範本字型：標題 Georgia、內文 Arial、中文一律微軟正黑體
FONT = 'Arial'; HEAD_FONT = 'Georgia'; EA_FONT = '微軟正黑體'
SYMBOLS = set('●◐○✓✗→↑①②③④⑤⑥⑦⚠')
FOOT = 'PwC FSC1｜AI 原生軟體開發提案｜期程為建議，效益依試點實測'
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


def draw_loop(s, cx, cy, R, d=1.05):
    # 自繪「循環」圖：六個階段排成一圈，中間是 AI 協作；不使用任何廠商的圖
    names = ['需求分析', 'SA&D', '開發', '測試', '上版／交付', '維運']
    pts = [(cx + R * math.cos(math.radians(-90 + 60 * i)), cy + R * math.sin(math.radians(-90 + 60 * i))) for i in range(6)]
    for i in range(6):
        (x1, y1), (x2, y2) = pts[i], pts[(i + 1) % 6]
        hline(s, x1, y1, x2, y2, OR2, 2.5)
    c = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 0.75), Inches(cy - 0.75), Inches(1.5), Inches(1.5))
    c.fill.solid(); c.fill.fore_color.rgb = OR; c.line.fill.background(); c.shadow.inherit = False
    tb(s, cx - 0.75, cy - 0.4, 1.5, 0.8, ['AI 協作', [('參與每一步', {'size': 9.5})]], size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.0)
    for (x, y), n in zip(pts, names):
        o = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x - d / 2), Inches(y - d / 2), Inches(d), Inches(d))
        o.fill.solid(); o.fill.fore_color.rgb = WHITE if n != '維運' else TINT; o.line.color.rgb = OR; o.line.width = Pt(1.5); o.shadow.inherit = False
        tb(s, x - d / 2, y - 0.25, d, 0.5, n, size=10.5 if len(n) < 5 else 9.5, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.0)

CHECK = [['是否屬作業委外、是否需依規定辦理', '保險業作業委託他人處理內部作業制度及程序辦法（2025-12-31 施行）', '導入前判斷並留下紀錄；AI 工具納入委外管理', '診斷'],
         ['雲端資料存放、雲端業者管理與查核', '保險業作業委外使用雲端服務自律規範', '使用公司核准的雲端合約；保留查核所需紀錄', '診斷'],
         ['AI 工具登錄 AI 清冊並做風險評估', '保險業運用人工智慧系統自律規範（2025-04-10 備查）', '每個 AI 工具與用途登錄；高風險用途加審', '診斷'],
         ['關鍵決定由人做，保險公司負最終責任', '同上；金管會 AI 指引「治理及問責」', '五個人工關卡，各有指定核准人', '①–⑤'],
         ['個資與客戶資料不進入 AI 可讀範圍', '個人資料保護法；金管會 AI 指引「隱私及客戶權益」', '資料分級；集中設定禁止讀取機密檔與金鑰', '規格、開發'],
         ['程式經資安審查，出事可退回上一版', '金管會 AI 指引「系統穩健性與安全性」', 'AI 資安初審＋人工審查；退版流程事先演練', '④ 合併、⑤ 上線'],
         ['誰要求、AI 產出什麼、誰核准，查得到', '金管會 AI 指引「透明性與可解釋性」', '需求說明、規格、計畫與核准紀錄版本化保存', '全程']]

def checklist_page(title, foot):
    s = content(title)
    grid(s, 0.45, 1.45, [3.2, 3.85, 3.85, 1.55], ['檢查項目', '依據', '開發流程中怎麼落實', '在哪一步檢查'], CHECK, 0.62, size=10.5)
    note(s, foot, y=6.42)
    return s

KPI = [['需求定稿時間', '從第一次討論到需求說明定稿', '試點前量測', '雙方約定'],
       ['第一次審查等待時間', '從提交程式到收到第一次審查意見', '試點前量測', '雙方約定'],
       ['有完整核准紀錄的改動比例', '比對版本紀錄與核准紀錄', '試點前量測', '雙方約定']]

def criteria_page(title):
    s = content(title)
    grid(s, 0.45, 1.45, [2.8, 3.6, 1.6, 1.6], ['指標（可依痛點替換）', '怎麼量', '基準', '合格線'], KPI, 0.75, size=11.5)
    tb(s, 10.25, 1.45, 2.65, 2.7, [[('結果怎麼用', {'bold': True, 'color': OR, 'size': 12.5})],
                                   [('達標：', {'bold': True}), ('擴大到下一個團隊。', {})],
                                   [('未達標：', {'bold': True}), ('一起檢討原因，調整後再試或停止，不擴大。', {})]], size=11.5, fill=TINT, margin=0.14)
    tb(s, 0.45, 4.45, 12.45, 1.2, [[('為什麼這樣做', {'bold': True, 'color': OR, 'size': 12.5})],
                                   '效益不先寫數字，而是試點開始前一起訂好指標和合格線；數據大多已在版本紀錄與 CI 系統中，不需要另外填表。試點結束，用同一套指標比較前後。'],
       size=12, fill=LG, margin=0.16)
    return s

ASPECT = [('面向誰', 'Who we serve'), ('解決什麼問題', 'Problems we solve'), ('帶來什麼價值', 'Value we deliver'), ('可擴充、可複製', 'Scalable and repeatable')]
VP = '讓保險公司用 AI 加快系統交付，同時守住法規、留下每一步紀錄，並把做法變成團隊可以重複使用的能力。'

def aspect_bar(s, i):
    # 頁面右上方標示本頁屬於哪個面向
    zh, en = ASPECT[i]
    tb(s, 9.9, 1.0, 3.0, 0.3, f'{i + 1}　{zh}｜{en}', size=9.5, bold=True, color=WHITE, fill=OR if i < 3 else INK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.02)

# ───────── 1 封面 ─────────
s = p.slides.add_slide(L['標題投影片'])
set_ph(s, 0, 'AI 原生軟體開發：\n從一個試點開始', 28)
set_ph(s, 1, 'PwC FSC1｜給保險業的提案｜2026 年 10 月', 12)

# ───────── 2 一頁看懂 ─────────
s = content('一頁看懂：協助保險公司把 AI 用進整個系統開發流程，交付更快，每一步都守得住、查得到')
tb(s, 0.45, 1.45, 12.45, 0.72, [[('我們的主張：', {'bold': True, 'color': OR}), (VP, {})]], size=13.5, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.18)
ANS = ['有系統開發需求的壽險、產險公司：IT／數位主管、業務單位、法遵與風控、管理層。',
       'AI 讓寫程式變快，但需求、審查、上線還是人工速度；又要面對新制上路與 AI 法規。',
       '新功能更快上線、規範自動執行、每一步有紀錄可查；成效用雙方先講好的指標證明。',
       '同一套範本、規範與關卡，從一個團隊推到全公司，下一個系統不必從頭來。']
for i, ((zh, en), t) in enumerate(zip(ASPECT, ANS)):
    x = 0.45 + i * 3.14
    tb(s, x, 2.32, 3.0, 0.82, [[(f'{i + 1}　{zh}', {'bold': True, 'size': 14})], [(en, {'size': 10})]], color=WHITE, fill=OR if i < 3 else INK, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
    tb(s, x, 3.19, 3.0, 2.1, t, size=12, fill=LG, margin=0.14)
    if i < 3: arrow(s, x + 3.0, 2.6, 0.14, 0.26)
callout(s, 5.5, '如何開始：', '先做一次準備度診斷，再挑一個真實需求做試點。試點前雙方先講好三個指標和合格線，達標才擴大。', h=0.75, size=12.5)

# ───────── 3 痛點 ─────────
s = content('保險公司的難題：同時要上新制、要用 AI、要守新規，開發流程卻還是舊的')
aspect_bar(s, 0)
PAIN = [('新制上路', '2026-01-01', 'IFRS 17 與 TW-ICS 接軌，財務、精算與報表系統都要跟著調整，系統開發需求增加。', '金管會保險局'),
        ('AI 投資要看到回報', '57%', '保險高階主管把生成式與代理型 AI 列為 2026 年首要科技投資；但跨產業有 95% 投資 AI 的企業尚未看到可衡量回報。', 'PwC 調查（Claims Journal，2025-11）；Bain 引述 MIT 2025'),
        ('法規要求同時到位', '2025-12-31', '保險業作業委外新辦法施行；AI 自律規範要求保險公司對 AI 結果負最終責任。用外部 AI 工具開發，要先想清楚。', '理律法律事務所；壽險／產險公會')]
for i, (h, big, t, src) in enumerate(PAIN):
    x = 0.45 + i * 4.18
    tb(s, x, 1.45, 4.05, 0.5, h, size=14, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.08)
    tb(s, x, 2.0, 4.05, 0.95, big, size=32 if len(big) < 6 else 26, bold=True, color=OR, fill=LG, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    tb(s, x, 2.95, 4.05, 1.75, t, size=12, fill=LG, margin=0.14)
    tb(s, x, 4.7, 4.05, 0.55, '來源：' + src, size=9, color=GREY, fill=LG, margin=0.07)
callout(s, 5.45, '所以：', '問題不在「要不要用 AI 寫程式」，而在讓整個開發流程跟上 AI 的速度，同時守得住法規、拿得出紀錄。', h=0.75, size=12.5)

# ───────── 4 面向誰 ─────────
s = content('面向誰：IT、業務、法遵、管理層在意的事不同，這個方案都要回答得了')
aspect_bar(s, 0)
ROWS = [['IT／數位主管', '交付速度、品質、團隊產能', '五階段流程與範本，AI 參與每一步；先在一個試點量出成效'],
        ['業務單位主管\n（理賠、核保、商品）', '需求說得清楚、上線不要拖', '業務人員和 AI 一起整理需求說明，每階段交出檔案，不必反覆開會'],
        ['法遵、風控、稽核', '委外、個資、AI 責任、稽核軌跡', '規範寫進 AI 會照做的規則；關卡有指定核准人；每一步留下紀錄'],
        ['管理層', '投資回報、風險可控', '從小試點開始，用先講好的指標比較前後，達標才擴大']]
grid(s, 0.45, 1.45, [2.6, 3.4, 6.45], ['角色', '他們在意什麼', '我們的答案'], ROWS, 0.75, size=12)
tb(s, 0.45, 5.05, 6.1, 0.95, [[('想先評估要不要做', {'bold': True})], [('→ 準備度診斷', {'color': OR, 'bold': True})]], size=12, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.16)
tb(s, 6.8, 5.05, 6.1, 0.95, [[('想先看到成效', {'bold': True})], [('→ 陪你做一個試點', {'color': OR, 'bold': True})]], size=12, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.16)

# ───────── 5 卡點與解法 ─────────
s = content('解決什麼問題：把最常見的六個卡點，對應到流程裡的具體做法')
aspect_bar(s, 1)
ROWS = [['需求來回好幾輪才定稿', '業務人員和 AI 一起整理需求說明；AI 追問範圍與限制、標出缺口，不自行補需求', '需求說明範本'],
        ['規格、程式、測試對不起來', '每個階段交出一份檔案，下一階段直接接手；測試案例綁需求編號', '規格、實作計畫範本'],
        ['審查排隊、品質看人', 'AI 先做多輪審查並排出嚴重度，人專注判斷意圖與風險', '審查規則、檢查清單'],
        ['公司規範沒人照做', '資安、個資、委外規範寫成 AI 做事時自動套用的規則', '規範規則庫'],
        ['怕 AI 失控、稽核查不到', '沒有核准不能合併、不能上線；誰核准、何時核准都有紀錄', '自動關卡與集中設定'],
        ['上線後同樣問題一再發生', '異常先由 AI 診斷，結果寫成新的需求，回到流程', '維運回饋']]
grid(s, 0.45, 1.45, [3.2, 6.65, 2.6], ['卡點', '我們的做法', '用到的元件'], ROWS, 0.72, size=11.5)

# ───────── 6 直線變循環 ─────────
s = content('我們的做法：把直線流程變成循環，AI 參與每一步，關鍵決定由貴公司指定的人做')
aspect_bar(s, 1)
tb(s, 0.45, 1.45, 2.1, 0.4, '以前：一條直線', size=12, bold=True, color=GREY, margin=0.02)
for i, n in enumerate(['需求分析', 'SA&D', '開發', '測試', '上版／交付', '維運']):
    tb(s, 0.45, 1.9 + i * 0.66, 2.1, 0.5, n, size=11, color=WHITE, fill=MID, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
tb(s, 0.45, 5.9, 2.1, 0.5, '每一段做完才交棒，退回一次就要重來', size=9.5, color=GREY, margin=0.02)
tb(s, 3.1, 1.45, 4.5, 0.4, '現在：一個循環', size=12, bold=True, color=OR, margin=0.02)
draw_loop(s, 5.35, 3.85, 1.75)
GATE = [('① 接受需求', '產品負責人'), ('② 確認規格', '產品負責人（高風險加技術主管）'), ('③ 核准計畫', '工程師（高風險找架構師）'),
        ('④ 核准合併', '程式負責人'), ('⑤ 授權上線', '發布經理')]
tb(s, 8.0, 1.45, 4.9, 0.42, '五個由人把關的關卡', size=12.5, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.08)
for i, (g, who) in enumerate(GATE):
    y = 1.93 + i * 0.58
    tb(s, 8.0, y, 1.55, 0.52, g, size=11.5, bold=True, color=OR, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.08)
    tb(s, 9.6, y, 3.3, 0.52, who, size=11, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.08)
tb(s, 8.0, 4.9, 4.9, 1.0, [[('原則：', {'bold': True, 'color': OR}), ('AI 可以一路做到正式環境的門口，但不能自己開門；寫程式的 AI 沒有核准權限。', {})]], size=11.5, fill=TINT, margin=0.14)
note(s, '本方法不綁特定 AI 工具，使用哪一套工具依貴公司政策與合約決定。', y=6.3)

# ───────── 7 四個面向 ─────────
s = content('四個面向環環相扣：找對人、解對問題，用先講好的指標證明價值，再推到全公司')
BOX = [(0.45, 1.45, '保險公司的 IT、業務、法遵與管理層', '第 3–4 頁'),
       (7.9, 1.45, '交付卡在人工步驟，又要守住新制與法規', '第 5–9 頁'),
       (7.9, 3.95, '更快上線、規範自動執行、紀錄可查；以試點成功標準證明', '第 10–11 頁'),
       (0.45, 3.95, '同一套範本、規範與關卡，從一個團隊推到全公司', '第 12 頁')]
for i, (x, y, t, pg) in enumerate(BOX):
    zh, en = ASPECT[i]
    tb(s, x, y, 5.0, 0.5, f'{i + 1}　{zh}　{en}', size=13, bold=True, color=WHITE, fill=OR if i < 3 else INK, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
    tb(s, x, y + 0.5, 5.0, 1.5, [t, [(pg, {'color': GREY, 'size': 10})]], size=12.5, fill=LG, margin=0.16)
tb(s, 5.65, 2.55, 2.05, 2.3, [[('我們的主張', {'bold': True, 'color': OR, 'size': 11})], 'AI 加快交付\n守住法規\n留下紀錄\n能力留在貴公司'], size=12, bold=True, fill=TINT, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.08)
arrow(s, 5.6, 1.7, 2.15, 0.32)
arrow(s, 10.25, 3.5, 0.32, 0.42, MSO_SHAPE.DOWN_ARROW)
arrow(s, 5.6, 5.3, 2.15, 0.32, MSO_SHAPE.LEFT_ARROW)
arrow(s, 2.8, 3.5, 0.32, 0.42, MSO_SHAPE.UP_ARROW)
callout(s, 6.05, '循環：', '每推到一個新團隊，範本與規範就更完整；下一個系統導入更快，也能服務更多業務單位。', h=0.55, size=11)

# ───────── 8 為什麼選我們 ─────────
s = content('為什麼選我們：台灣保險法規、保險業務與在地系統、可稽核的關卡，三者同時具備')
aspect_bar(s, 1)
DIFF = [('台灣保險法規內建', '委外辦法、公會 AI 與雲端自律規範、金管會 AI 指引，已整理成開發流程的檢查清單（下一頁），寫進 AI 會照做的規則。'),
        ('懂保險業務與在地系統', 'FSC1 熟悉理賠、核保、核心系統與 IFRS 17，試點直接用貴公司真正在痛的流程，不必從頭解釋業務。'),
        ('可稽核的人工關卡', '五個關卡誰核准、什麼時候核准，全部查得到；寫程式的 AI 不能核准自己的程式。')]
for i, (h, t) in enumerate(DIFF):
    x = 0.45 + i * 4.18
    circle(s, x, 1.5, 0.55, str(i + 1), size=14)
    tb(s, x + 0.65, 1.5, 3.4, 0.55, h, size=14, bold=True, color=OR, anchor=MSO_ANCHOR.MIDDLE)
    tb(s, x, 2.15, 4.05, 1.75, t, size=12, fill=LG, margin=0.16)
tb(s, 0.45, 4.1, 12.45, 1.6, [[('我們的實作成果與現場示範【待 IT 團隊填入，填好前本頁不對外】', {'bold': True, 'color': OR, 'size': 12.5})],
                              '・已經做出來的系統或工具：名稱、給誰用、解決什麼問題',
                              '・用到的做法與量到的成效',
                              '・現場可以示範的內容'],
   size=12, fill=WHITE, line=OR, margin=0.18)

# ───────── 9 法規檢查清單 ─────────
s = checklist_page('看得到的差異：保險業 AI 開發法規檢查清單，把法規要求對應到開發流程的每一步',
                   '本清單為導入規劃工具，非法律意見；是否屬委外、是否需申報，依貴公司個案與實際架構判斷，並與貴公司法遵確認。')
aspect_bar(s, 1)

# ───────── 10 價值鏈 ─────────
s = content('帶來什麼價值：每個問題都對應一個可以量測的商業成果')
aspect_bar(s, 2)
ROWS = [['需求定稿慢', '需求說明範本＋AI 追問', '新功能更快上線', '從第一次討論到需求定稿的時間'],
        ['審查排隊', 'AI 先審、人判斷', '工程師把時間花在需要判斷的工作', '第一次審查的等待時間、返工次數'],
        ['品質不穩', 'AI 自行測試，案例綁需求', '上線事故減少', '上線失敗率、上線後才發現的缺陷'],
        ['法規與稽核壓力', '規範規則＋自動關卡＋紀錄', '主管機關或稽核詢問時拿得出佐證', '有完整核准紀錄的改動比例'],
        ['知識只在個人腦中', '團隊說明檔、範本與教學', '人員異動不影響交付，新人上手更快', '新成員獨立交付第一個改動的時間']]
grid(s, 0.45, 1.45, [2.4, 3.0, 3.4, 3.65], ['客戶問題', '→ 解決方式', '→ 商業價值', '→ 衡量指標'], ROWS, 0.76, size=11.5)

# ───────── 11 試點成功標準 ─────────
s = criteria_page('試點成功標準：開始之前，雙方先講好三個指標和合格線，達標才擴大')
aspect_bar(s, 2)

# ───────── 12 從一個團隊到全公司 ─────────
s = content('可擴充、可複製：同一套做法從一個團隊推到全公司，下一個系統不必從頭來')
aspect_bar(s, 3)
STG = [('一個試點團隊', '用一個真實需求走完五個階段；建立範本、規範規則與團隊說明檔。'),
       ('更多團隊', '範本、規範規則與關卡設定直接沿用；只補該系統自己的規範與核准人。'),
       ('全公司', '集中管理設定、統一關卡與紀錄格式；AI 工具納入 AI 清冊統一管理。')]
for i, (h, t) in enumerate(STG):
    x = 0.45 + i * 4.18
    tb(s, x, 1.45, 4.05, 0.6, h, size=15, bold=True, color=WHITE, fill=[OR, OR2, INK][i], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    tb(s, x, 2.1, 4.05, 1.3, t, size=12, fill=LG, margin=0.16)
    if i < 2: arrow(s, x + 4.06, 1.6, 0.11, 0.3)
tb(s, 0.45, 3.6, 6.1, 0.45, '每個客戶都一樣（直接沿用）', size=12.5, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
tb(s, 0.45, 4.08, 6.1, 1.6, [[('✓　' + t, {})] for t in ['五階段流程與五個人工關卡', '需求說明、規格、實作計畫範本', '保險業法規檢查清單', '成效指標定義與教材']], size=12, fill=TINT, margin=0.16)
tb(s, 6.8, 3.6, 6.1, 0.45, '依貴公司調整', size=12.5, bold=True, color=WHITE, fill=INK, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
tb(s, 6.8, 4.08, 6.1, 1.6, [[('●　' + t, {})] for t in ['貴公司的資安、命名與品牌規範', '核心系統與既有工具串接（Jira、變更管理）', '資料分級與各關卡的核准人', '試點題目與指標基準']], size=12, fill=LG, margin=0.16)
callout(s, 5.85, '延伸：', '同一套做法可延伸到 IFRS 17／TW-ICS 相關系統、AI 治理，以及核保、理賠等更多業務情境。', h=0.6, size=11.5)

# ───────── 13 合作方式 ─────────
s = content('合作方式：從準備度診斷開始，再陪貴公司做一個試點')
OFF = [('準備度診斷', ['盤點六個階段的現況與卡點', '法規檢查：委外、雲端、AI 清冊、資料分級', '挑選試點並訂好成功標準'], '瓶頸地圖、法規檢查結果、試點建議與導入藍圖', '約 4–6 週'),
       ('陪你做一個試點', ['用一個真實需求走完五個階段', '同步教會各角色，建立範本與審查規則', '用先講好的指標比較前後'], '完成的試點、成效報告、範本與規範規則、教材', '約 6–8 週')]
for i, (h, items, out, d) in enumerate(OFF):
    x = 0.45 + i * 6.35
    tb(s, x, 1.45, 6.1, 0.6, h, size=16, bold=True, color=WHITE, fill=OR if i == 0 else OR2, anchor=MSO_ANCHOR.MIDDLE, margin=0.16)
    tb(s, x, 2.1, 6.1, 1.6, [[('・' + t, {})] for t in items], size=12.5, fill=LG, margin=0.18)
    tb(s, x, 3.75, 6.1, 0.85, [[('交付：', {'bold': True, 'color': OR}), (out, {})]], size=12, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.16)
    tb(s, x, 4.65, 6.1, 0.5, [[('建議期程：', {'bold': True, 'color': OR}), (d + '，依範圍確認；固定範圍、固定價格', {})]], size=11.5, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.16)
arrow(s, 6.57, 1.6, 0.2, 0.3)
callout(s, 5.4, '之後：', '試點達標後，依貴公司需要擴大到更多團隊；工具包授權與駐點支援等合作方式可再討論。', h=0.65)

# ───────── 14 路線圖與下一步 ─────────
s = content('如何開始：三件事啟動診斷，四個階段從評估走到全公司')
PH = [('0　準備度診斷', '選定試點與指標'), ('1　試點', '指標達到合格線'), ('2　擴大', '管理層核准擴大'), ('3　全公司', '持續運作與改善')]
for i, (h, g) in enumerate(PH):
    x = 0.45 + i * 3.14
    tb(s, x, 1.45, 3.0, 0.55, h, size=13.5, bold=True, color=WHITE, fill=[OR, OR, OR2, MID][i], anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
    tb(s, x, 2.05, 3.0, 0.5, [[('進下一步：', {'bold': True, 'color': OR}), (g, {})]], size=11, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
    if i < 3: arrow(s, x + 3.0, 1.6, 0.14, 0.26)
NX = [('對齊工作坊', '對齊目標，從貴公司的需求清單挑 1–2 個候選試點。'), ('前置確認', '資料分級、雲端與 AI 工具的委外安排。'), ('啟動診斷', '確認範圍與雙方窗口，開始盤點現況。')]
for i, (h, t) in enumerate(NX):
    y = 2.85 + i * 0.95
    circle(s, 0.45, y + 0.15, 0.55, str(i + 1), size=15)
    tb(s, 1.15, y, 6.5, 0.85, [[(h, {'bold': True, 'size': 13})], t], size=11.5, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.14)
tb(s, 8.0, 2.85, 4.9, 2.75, [[('請貴公司提供', {'bold': True, 'color': OR, 'size': 13})],
                             '・IT 主管與業務單位窗口', '・現有開發流程與工具（Jira、CI、變更管理）', '・候選需求清單', '・資安、個資與法遵規範'], size=12, fill=TINT, margin=0.18)
callout(s, 5.85, '一句話：', 'AI 幫你把每一步做得更快，每一步都留下紀錄；要不要做、能不能上線，永遠由你的人決定。', h=0.6, size=12)

# ───────── 附錄 A 示範情境 ─────────
s = content('附錄 A｜示範情境：「理賠進度自助查詢」從客服痛點到上線，人只在關鍵點做決定')
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
callout(s, 4.6, '說明：', '此為示範情境，取自公開的 AI 原生開發實務範例，非客戶案例；實際試點題目於診斷時與貴公司共同挑選。', h=0.7)

# ───────── 附錄 B 資料與來源 ─────────
s = content('附錄 B｜資料與假設說明、來源')
tb(s, 0.45, 1.45, 12.45, 0.95, [[('資料與假設：', {'bold': True, 'color': OR}),
                                 ('數字皆引自下列公開來源。期程為建議，依範圍確認；效益不預先承諾，以試點前雙方約定的指標與合格線實測。法規檢查清單為導入規劃工具，非法律意見。', {})]],
   size=12, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.16)
SRCS = [('金管會保險局，IFRS 17／TW-ICS 專區', None),
        ('Claims Journal，PwC 保險高階主管調查（57%），2025-11-17', 'https://amp.claimsjournal.com/news/national/2025/11/17/334094.htm'),
        ('Bain，Reimagining Insurance for the AI Era（引述 MIT 2025：95%）', 'https://www.bain.com/ko/results/reimagining-insurance-for-the-ai-era'),
        ('理律法律事務所，保險業作業委外新辦法（2025-12-31 施行）', 'https://www.leeandli.com/TW/Newsletters/7525.htm'),
        ('壽險／產險公會，保險業運用人工智慧系統自律規範（2025-04-10 備查）', 'https://law.lia-roc.org.tw/Law/GetFile/0000394231'),
        ('壽險／產險公會，保險業作業委外使用雲端服務自律規範', 'https://law.lia-roc.org.tw/Law/GetFile/0000419987'),
        ('金管會，金融業運用人工智慧（AI）指引，2024-06', 'https://www.fsc.gov.tw/websitedowndoc?file=chfsc/202408231741530.pdf'),
        (f'流程概念參考：Louis Claxton，{PB}，2026', None)]
for i, (lab, url) in enumerate(SRCS):
    seg = (lab, {'link': url}) if url else (lab, {})
    tb(s, 0.45, 2.55 + i * 0.47, 12.45, 0.42, [[(f'{i + 1}. ', {'bold': True, 'color': OR}), seg]], size=11, fill=LG if i % 2 == 0 else WHITE, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)

p.save(OUT)
print(OUT, len(p.slides), 'slides')
