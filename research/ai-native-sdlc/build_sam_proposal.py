# AI 原生軟體開發：給合夥人向企業客戶提案的簡報（正文 26 頁＋附錄 6 頁）
# 整合兩份研究：FSC1 IT 團隊〈AI 協作開發流程與產出範本〉v3、〈AI 原生 SDLC 白話說明〉（整理自 Anthropic 手冊），
# 以四個面向（面向誰、解決什麼問題、帶來什麼價值、可擴充與可複製）為主軸；不含任何個人姓名，同業公司名稱只在附錄
# 正文 27 頁＋附錄 6 頁
# 用法：python3 build_sam_proposal.py <範本.pptx> <輸出.pptx>
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
FOOT = 'PwC FSC1｜AI 原生軟體開發提案｜標示「建議」「示意」者非承諾，效益以試點實測'
V3 = 'FSC1〈AI 協作開發流程與產出範本〉v3'
PB = 'Anthropic《The AI-Native SDLC Playbook》'
SMR = 'MIT Sloan Management Review，2024'

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

def section(num, title):
    # 章節轉場頁：沿用範本「章節標題」版面，只放章節編號與名稱
    s = p.slides.add_slide(L['章節標題'])
    set_ph(s, 0, title, 30); set_ph(s, 1, num, 180); footer(s)
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

TAGC = {'建議': OR, '示意': MID, '假設': GREY, '待確認': INK}

def tag(s, *kinds):
    # 標示本頁含有建議、示意、假設或待確認內容；沒有標示者為有來源的事實或既有設計
    x = 12.9 - 0.95 * len(kinds)
    for i, k in enumerate(kinds):
        tb(s, x + i * 0.95, 1.0, 0.88, 0.3, k, size=9.5, bold=True, color=WHITE, fill=TAGC[k], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.02)

ASPECT = [('面向誰', 'Who we serve'), ('解決什麼問題', 'Problems we solve'), ('帶來什麼價值', 'Value we deliver'), ('可擴充、可複製', 'Scalable and repeatable')]
VP = '協助保險公司把 AI 用進整個系統開發流程：交付更快、每一步可稽核，並把做法留在貴公司的團隊裡。'
STAGES = ['需求分析', 'SA&D', '開發', '測試', '上版／交付', '維運']

# ───────── 1 封面 ─────────
s = p.slides.add_slide(L['標題投影片'])
set_ph(s, 0, 'AI 原生軟體開發：\n從一個試點開始', 28)
set_ph(s, 1, 'PwC FSC1｜給保險業的提案｜2026 年 10 月', 12)

# ───────── 2 一頁看懂 ─────────
s = content('一頁看懂：從一個試點開始，把 AI 用進整個開發流程，交付更快、每一步都查得到')
tag(s, '建議')
tb(s, 0.45, 1.45, 12.45, 0.72, [[('我們的主張：', {'bold': True, 'color': OR}), (VP, {})]], size=13.5, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.18)
ANS = ['有自建或駐點開發團隊的壽險、產險公司；IT 主管付費，法遵與風控把關，業務單位提出需求。',
       'AI 讓寫程式變快，但需求、審查、上線還是人工速度；新制上路與 AI 法規責任同時到來。',
       '交接靠檔案、規範自動執行、關卡有人簽：更快上線、紀錄可查；成效以雙方先講好的指標證明。',
       '流程、範本、規範規則與關卡可以沿用；從一個試點推到其他系統、團隊與新制相關專案。']
for i, ((zh, en), t) in enumerate(zip(ASPECT, ANS)):
    x = 0.45 + i * 3.14
    tb(s, x, 2.32, 3.0, 0.82, [[(f'{i + 1}　{zh}', {'bold': True, 'size': 14})], [(en, {'size': 10})]], color=WHITE, fill=OR if i < 3 else INK, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
    tb(s, x, 3.19, 3.0, 2.0, t, size=12, fill=LG, margin=0.14)
    if i < 3: arrow(s, x + 3.0, 2.6, 0.14, 0.26)
STEP3 = [('現場示範', '用貴公司一個需求示範'), ('準備度診斷', '挑試點、訂成功標準'), ('試點陪跑', '達標才擴大')]
for i, (h, t) in enumerate(STEP3):
    x = 0.45 + i * 4.18
    tb(s, x, 5.38, 4.05, 0.8, [[(f'{i + 1}　{h}', {'bold': True, 'size': 13, 'color': WHITE})], [(t, {'color': WHITE, 'size': 11})]], fill=[OR, OR2, INK][i], anchor=MSO_ANCHOR.MIDDLE, margin=0.14)
    if i < 2: arrow(s, x + 4.06, 5.62, 0.11, 0.3)

# ───────── 3 議程 ─────────
s = content('議程')
AG = [('01', '面向誰', '優先服務誰、誰決策與付費、為什麼是現在、問題的根源', '4–7'),
      ('02', '解決什麼問題', '框架怎麼運作、四個面向如何組成方案、具體賣什麼、為什麼選我們、我們自己已在用', '8–15'),
      ('03', '帶來什麼價值', '價值鏈、導入前後對比情境、怎麼衡量', '16–19'),
      ('04', '可擴充、可複製', '可重用與需客製的部分、延伸例子、前提與限制', '20–23'),
      ('05', '如何開始', '路線圖、成功標準與擴展方式、下一步', '24–27'),
      ('附錄', '附錄', '研究整合、競爭參考來源、法規檢查清單、AI 任務分類、資料來源', '28–33')]
for i, (n, h, d, pg) in enumerate(AG):
    y = 1.45 + i * 0.8
    circle(s, 0.6, y + 0.1, 0.52, n, fill=OR if n != '附錄' else MID, size=12 if n != '附錄' else 10)
    tb(s, 1.35, y, 3.0, 0.7, h, size=16, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    tb(s, 4.4, y, 7.2, 0.7, d, size=12, color=GREY, anchor=MSO_ANCHOR.MIDDLE)
    tb(s, 11.65, y, 1.25, 0.7, pg, size=15, bold=True, color=OR, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT)
    hline(s, 1.35, y + 0.75, 12.9, y + 0.75)
note(s, '右側數字為頁碼。頁面右上方標示「建議」「示意」「假設」「待確認」；未標示者為有來源的事實或既有設計。', y=6.42)

# ───────── 章節 01 ─────────
section('01', '面向誰\nWho we serve')

# ───────── 5 客戶與決策鏈 ─────────
s = content('優先服務有自建開發團隊的壽險與產險公司：IT 主管付費，法遵與風控握有否決權')
tag(s, '假設')
ROWS = [['IT／數位主管', '決策與付費', '交付跟不上業務與新制需求；AI 工具用了，整體卻沒變快', '流程改造＋一個可量測的試點'],
        ['業務單位主管\n（理賠、核保、商品）', '提出需求、共同贊助', '需求來回多輪、上線時間難預期', '業務人員和 AI 一起寫需求說明，每階段交出檔案'],
        ['法遵、風控、資安', '否決權', '委外、個資與 AI 責任；稽核時要拿得出紀錄', '法規檢查清單、人工關卡、完整核准紀錄'],
        ['管理層', '核准預算', 'AI 投資看不到回報', '小範圍試點，先講好成功標準']]
grid(s, 0.45, 1.45, [2.45, 1.85, 4.15, 4.0], ['角色', '採購中的角色', '他們的問題', '我們給的答案'], ROWS, 0.8, size=11.5)
tb(s, 0.45, 5.2, 12.45, 0.9, [[('優先客群的條件：', {'bold': True, 'color': OR}), ('有自建或駐點開發團隊；已有版本控制與 CI；近期有新制相關或核心系統周邊的開發需求。', {})]], size=12, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.16)
note(s, '決策鏈為依經驗的假設，需在診斷訪談中與客戶確認。', y=6.3)

# ───────── 6 為什麼是現在 ─────────
s = content('為什麼是現在：新制上路、AI 投資要看到回報、法規責任落在保險公司身上，三件事同時發生')
PAIN = [('新制上路', '2026-01-01', 'IFRS 17 與 TW-ICS 接軌，財務、精算與報表相關系統都要跟著調整。', '金管會保險局'),
        ('AI 投資要看到回報', '57%', '保險高階主管把生成式與代理型 AI 列為 2026 年首要科技投資；跨產業則有 95% 投資 AI 的企業尚未看到可衡量回報。', 'PwC 調查（Claims Journal，2025-11）；Bain 引述 MIT 2025'),
        ('法規責任到位', '2025-12-31', '保險業作業委外新辦法施行；AI 自律規範要求保險公司對 AI 結果負最終責任。', '理律法律事務所；壽險／產險公會')]
for i, (h, big, t, src) in enumerate(PAIN):
    x = 0.45 + i * 4.18
    tb(s, x, 1.45, 4.05, 0.5, h, size=14, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.08)
    tb(s, x, 2.0, 4.05, 0.95, big, size=32 if len(big) < 6 else 26, bold=True, color=OR, fill=LG, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    tb(s, x, 2.95, 4.05, 1.75, t, size=12, fill=LG, margin=0.14)
    tb(s, x, 4.7, 4.05, 0.55, '來源：' + src, size=9, color=GREY, fill=LG, margin=0.07)
callout(s, 5.45, '意涵：', '開發需求變多、AI 預算要交代成果、監理要求留下紀錄。單靠買 AI 工具解決不了，流程要跟著改。', h=0.75, size=12.5)
note(s, '完整來源與連結見附錄 E。', y=6.35)

# ───────── 7 根源 ─────────
s = content('問題的根源：寫程式變快了，但需求、審查、上線還是人工速度，品質與合規仍靠個人')
tb(s, 0.45, 1.45, 6.4, 0.4, '寫程式只是六個階段中的一段', size=12.5, bold=True, color=GREY, margin=0.02)
for i, n in enumerate(STAGES):
    y = 1.9 + i * 0.66
    fast = n == '開發'
    tb(s, 0.45, y, 1.8, 0.56, n, size=12, bold=True, color=WHITE, fill=OR if fast else MID, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    tb(s, 2.3, y, 4.55, 0.56, 'AI 已大幅加速' if fast else '仍是人工速度：開會、等待、交接', size=11.5, bold=fast, color=OR if fast else GREY, fill=TINT if fast else LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.14)
RC = [('瓶頸移位', '寫程式變快後，等待時間移到需求、審查與上線核准；程式量變多，審查反而排得更長。'),
      ('交接靠開會與文件', '每一段做完交給下一組人，退回一次就要重來。'),
      ('規則靠個人', 'AI 產出的品質取決於每個人的 Prompt 寫得好不好；公司規範有沒有被遵守，要靠人記得。')]
for i, (h, t) in enumerate(RC):
    tb(s, 7.1, 1.45 + i * 1.45, 5.8, 1.3, [[(f'{i + 1}　{h}', {'bold': True, 'color': OR, 'size': 13})], t], size=11.5, fill=LG, margin=0.14)
note(s, f'整理自兩份研究：瓶頸移位見 {PB}；「規則靠個人」見 {V3} 的 Prompt 撰寫原則。', y=6.0)

# ───────── 章節 02 ─────────
section('02', '解決什麼問題\nProblems we solve')

# ───────── 9 框架 ─────────
s = content('我們的做法：五個開發階段加上維運串成循環，AI 參與每一步，關鍵決定由指定的人做')
tb(s, 0.45, 1.45, 2.1, 0.4, '以前：一條直線', size=12, bold=True, color=GREY, margin=0.02)
for i, n in enumerate(STAGES):
    tb(s, 0.45, 1.9 + i * 0.66, 2.1, 0.5, n, size=11, color=WHITE, fill=MID, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
tb(s, 0.45, 5.9, 2.1, 0.5, '做完一段才交棒，退回一次就要重來', size=9.5, color=GREY, margin=0.02)
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
note(s, '階段名稱沿用台灣保險業 IT 熟悉的用語，再加上維運形成循環。本方法不綁特定 AI 工具，依貴公司政策與合約選定。', y=6.3)

# ───────── 10 三個機制 ─────────
s = content('三個機制讓循環守得住：交接靠檔案、規則自動執行、關卡有人簽')
ROWS = [['交接靠檔案', '需求說明 → 規格 → 實作計畫 → 測試結果 → 核准紀錄，下一步直接讀上一步的檔案', '需求反覆、交接失真', '三份範本（需求說明、規格、實作計畫）'],
        ['規則自動執行', '資安、個資、委外規範寫成 AI 自動套用的規則；必守的事由自動關卡擋下違規', '合規與品質靠個人', '規範規則庫、自動關卡設定、Review Checklist'],
        ['關卡有人簽', '五個關鍵決定由指定的人做，每次核准都有時間紀錄', '怕失控、稽核查不到', '關卡設計、核准紀錄格式']]
grid(s, 0.45, 1.45, [2.0, 5.3, 2.3, 2.85], ['機制', '怎麼運作', '解決的問題', '交付的元件'], ROWS, 0.95, size=11.5)
tb(s, 0.45, 4.85, 12.45, 0.38, '控制從「靠個人」走向「靠制度」的四層', size=12, bold=True, color=GREY, margin=0.02)
LV = [('Prompt 撰寫原則', '因人而異'), ('範本與團隊說明檔', '格式一致'), ('規範規則', '讓違規變少'), ('自動關卡', '讓違規做不到')]
for i, (h, t) in enumerate(LV):
    x = 0.45 + i * 3.14
    tb(s, x, 5.25, 3.0, 0.75, [[(f'{i + 1}　{h}', {'bold': True, 'size': 12, 'color': WHITE})], [(t, {'color': WHITE, 'size': 10.5})]], fill=MID if i == 0 else OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
    if i < 3: arrow(s, x + 3.0, 5.48, 0.14, 0.28)
note(s, f'第 1 層出自 {V3}；第 2–4 層依 {PB} 的技能、鉤子與集中管理設定。', y=6.15)

# ───────── 11 四個面向 ─────────
s = content('四個面向環環相扣：找對人、解對問題，用先講好的指標證明價值，再推到更多系統與團隊')
BOX = [(0.45, 1.45, '有自建開發團隊的壽險、產險公司；IT 付費、法遵把關', '第 4–7 頁'),
       (7.9, 1.45, '循環流程＋三個機制；三個服務套組', '第 8–15 頁'),
       (7.9, 3.95, '更快上線、規範自動執行、紀錄可查；以試點成功標準證明', '第 16–19 頁'),
       (0.45, 3.95, '流程、範本、規則與關卡沿用；延伸到其他系統、團隊與新制專案', '第 20–23 頁')]
for i, (x, y, t, pg) in enumerate(BOX):
    zh, en = ASPECT[i]
    tb(s, x, y, 5.0, 0.5, f'{i + 1}　{zh}　{en}', size=13, bold=True, color=WHITE, fill=OR if i < 3 else INK, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
    tb(s, x, y + 0.5, 5.0, 1.5, [t, [(pg, {'color': GREY, 'size': 10})]], size=12.5, fill=LG, margin=0.16)
tb(s, 5.65, 2.55, 2.05, 2.3, [[('我們的主張', {'bold': True, 'color': OR, 'size': 11})], 'AI 加快交付\n每一步可稽核\n能力留在\n貴公司'], size=12, bold=True, fill=TINT, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.08)
arrow(s, 5.6, 1.7, 2.15, 0.32)
arrow(s, 10.25, 3.5, 0.32, 0.42, MSO_SHAPE.DOWN_ARROW)
arrow(s, 5.6, 5.3, 2.15, 0.32, MSO_SHAPE.LEFT_ARROW)
arrow(s, 2.8, 3.5, 0.32, 0.42, MSO_SHAPE.UP_ARROW)
callout(s, 6.05, '循環：', '每擴展到一個新系統或團隊，範本與規範規則就更完整，下一次導入更快。', h=0.55, size=11)

# ───────── 12 服務套組 ─────────
s = content('具體賣什麼：三個服務套組，從診斷、試點到擴大，每一步都有明確交付與客戶配合事項')
tag(s, '建議', '示意')
ROWS = [['A　準備度診斷', '想評估是否導入、從哪裡開始', '瓶頸地圖、法規檢查結果、試點建議與成功標準、導入藍圖', '約 4–6 週', 'IT 與業務窗口；現行流程與工具說明；候選需求清單', '固定範圍、固定價'],
        ['B　試點陪跑', '想先看到成效', '完成的試點；三份範本；規範規則與關卡設定；各角色教材；成效報告', '約 6–8 週', '一個真實需求與團隊；測試環境；AI 工具合約或內部核准', '每個試點固定價，可設里程碑'],
        ['C　擴大與內化', '試點達標，要推到更多團隊', '擴大路線圖；規範規則庫；集中管理設定；各團隊教學', '依範圍', '管理層核准；各團隊窗口與核准人', '按團隊或階段計價'],
        ['加購', '有特定需求時', '工具包授權；駐點支援', '依需求', '—', '授權費；人月（駐點需先確認獨立性）']]
grid(s, 0.45, 1.45, [1.85, 1.95, 3.3, 1.2, 2.45, 1.7], ['套組', '適合誰', '交付成果', '期程（示意）', '客戶需配合', '商業模式（建議）'], ROWS, 0.95, size=10.5, hot=(1,))
note(s, '套組對應研究二的四個服務模組：A＝現況診斷；B＝規範與控制設計＋試點陪跑與教學；C＝量測與擴大。期程依範圍確認，金額另行報價；按任務或成果計價的思路參考 ' + SMR + '。', y=5.95)

# ───────── 13 市場參考 ─────────
s = content('市場參考：同業多以「預建資產＋顧問導入」包裝，我們借鏡其包裝方式，以在地落地區隔')
ROWS = [['加速器＋顧問導入', '保險業代理型 AI 導入藍圖與預建代理 [B1]；IFRS 17 加速器 [B2][B3]', '預建資產；評估 → 設計 → 試點 → 規模化；IFRS 17 加速器宣稱 90 天內產出', '固定範圍與期程，降低客戶疑慮', '預建的是台灣保險法規規則與關卡'],
        ['核心系統＋AI 工具組', '核心系統 AI 工具組 [B4]', '圍繞特定核心系統產品的加速器', '鎖定明確情境', '不綁特定核心系統或 AI 工具'],
        ['軟體／數據產品', '定價與核保軟體 [B5]；保險市場資料庫訂閱 [B6]', '授權或訂閱', '標準化程度高', '先不做軟體，以試點累積可重用資產'],
        ['方法論與藍圖', '代理型 AI 核心現代化方法 [B7]；AI-First 產險藍圖 [B8]', '顧問專案', '清楚的成熟度階段', '直接落地到可量測的試點']]
grid(s, 0.45, 1.45, [2.1, 3.0, 3.05, 2.1, 2.2], ['包裝型態', '代表產品（附錄 B）', '包裝與交付方式', '值得借鏡', '我們的差異'], ROWS, 0.95, size=10.5, hot=(0,))
note(s, '代表產品的公司名稱與來源連結見附錄 B；資料為業者公開資訊，未經獨立驗證，「宣稱」為業者自述。', y=5.85)

# ───────── 14 為什麼選我們 ─────────
s = content('為什麼選我們：三個差異都有對應的證據；尚缺的實作成果，以第一步的現場示範補上')
tag(s, '待確認')
ROWS = [['台灣保險法規內建', '委外、雲端、AI 自律規範與金管會指引對應到開發的每一步', '法規檢查清單（附錄 C）', '草稿完成，待法遵確認'],
        ['懂保險業務與在地系統', '試點直接用理賠、核保、核心系統等真實流程，不必從頭解釋業務', 'FSC1 已建立的 AI 協作開發流程與產出範本', '流程已建立；實作成果待補'],
        ['可稽核的人工關卡', '誰核准、何時核准都查得到；寫程式的 AI 不能核准自己的程式', '關卡設計（第 9 頁）', '設計完成，試點驗證'],
        ['現場示範', '用貴公司一個需求，示範從需求說明到實作計畫', 'FSC1 IT 團隊', '待安排']]
grid(s, 0.45, 1.45, [2.5, 4.6, 3.1, 2.25], ['差異', '貴公司得到什麼', '證據', '狀態'], ROWS, 0.85, size=11.5, hot=(3,))
callout(s, 5.4, '和同業的根本差別：', '國際方案以預建代理與藍圖為主；我們把在地法規、在地系統與可稽核關卡綁在一起，從貴公司的真實需求開始。', h=0.75)
note(s, '「國際方案套到台灣客製化舊核心與在地法規需大量調整」為依公開資料的推論。', y=6.3)

# ───────── 15 我們自己已在用 ─────────
s = content('我們先用在自己身上：FSC1 已建立 AI 協作開發流程，產出、分工與檢查點都已定義')
tag(s, '待確認')
ROWS = [['需求分析', 'Prototype、欄位清單、業務流程圖／BPMN', '依需求產出 Prototype、整理欄位、繪製流程圖', '畫面、欄位、流程與需求一致'],
        ['SA&D', 'Use Case、ERD、API 設計文件、架構圖、Test Plan', '產出 Use Case、ERD、OpenAPI 初稿，檢查資料結構', '技術可行性、資料關聯、需求可追溯'],
        ['開發', '程式、API 實作、SQL／Migration、版本紀錄', '依規格產生程式與 Unit Test，協助 Code Review', '正確性、影響範圍、安全性、可維護性'],
        ['測試', 'Test Case、Test Data、Test Report、Bug List、UAT 紀錄', '產生正向、反向、邊界案例與測試資料，分析失敗原因', '覆蓋率、預期結果、缺陷原因'],
        ['上版／交付', 'Deploy Guide、Release Note、User Guide、Log 監控', '整理 Release Note 與 Deploy Guide，分析 Log', '與實際版本一致、包含 rollback']]
grid(s, 0.45, 1.45, [1.6, 3.85, 3.85, 3.15], ['階段', '主要產出', 'AI 協作', 'IT Review 重點'], ROWS, 0.68, size=10.5)
tb(s, 0.45, 5.35, 6.1, 0.85, [[('已具備：', {'bold': True, 'color': OR}), ('五階段流程、各階段產出清單、IT 與 AI 分工、Review Checklist、Prompt 撰寫原則與範例。', {})]], size=11, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.14)
tb(s, 6.8, 5.35, 6.1, 0.85, [[('待補：', {'bold': True, 'color': INK}), ('各階段的實際案例畫面與量到的成效，提案前由 IT 團隊提供，或以現場示範呈現。', {})]], size=11, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.14)
note(s, f'來源：研究一 {V3}。', y=6.35)

# ───────── 章節 03 ─────────
section('03', '帶來什麼價值\nValue we deliver')

# ───────── 16 價值鏈 ─────────
s = content('價值鏈：每個痛點都對應一項交付成果、一個商業效益和一個可以量測的指標')
tag(s, '建議')
ROWS = [['需求來回多輪', '業務和 AI 一起寫需求說明', '需求說明範本', '新功能更快上線', '從第一次討論到需求定稿的時間'],
        ['審查排隊、品質看人', 'AI 先審、人判斷風險', '審查規則、Review Checklist', '工程師時間花在需要判斷的工作', '第一次審查的等待時間、返工次數'],
        ['上線出錯', 'AI 每次改動自行測試', '測試做法與合併條件', '上線事故減少', '上線失敗率、上線後才發現的缺陷'],
        ['法規與稽核壓力', '規範規則＋自動關卡＋紀錄', '規範規則庫、法規檢查清單', '金檢與內稽時拿得出佐證', '有完整核准紀錄的改動比例'],
        ['知識只在個人腦中', '範本、團隊說明檔與教學', '教材、團隊說明檔', '人員異動不影響交付', '新成員獨立交付第一個改動的時間']]
grid(s, 0.45, 1.45, [2.15, 2.6, 2.5, 2.55, 2.65], ['痛點', '→ 解決方式', '→ 交付成果', '→ 商業效益', '→ 衡量指標'], ROWS, 0.8, size=11)
note(s, '商業效益以試點前後同一套指標比較，不預先承諾數字。', y=5.95)

# ───────── 17 前後對比 ─────────
s = content('導入前後對比：以「理賠進度自助查詢」為例，每一步都有人負責、有檔案交接、有紀錄可查')
tag(s, '示意')
ROWS = [['需求分析', '多次會議，再由專人寫成文件', '理賠同仁和 AI 一起整理需求說明，AI 追問範圍與限制', '產品負責人接受'],
        ['SA&D', '分析與設計分開做，規範靠人記得', 'AI 起草規格，自動套用「入口網站不新增個資」等規範，衝突處標出', '產品負責人確認'],
        ['開發', '讀完設計就寫，做法只在個人腦中', 'AI 先寫實作計畫（含核心系統流量限制與快取），核准後才動手', '工程師核准'],
        ['測試', '階段末才測，問題晚發現', 'AI 每次改動自己跑測試，各種理賠狀態都正確才交人', 'AI 自行驗證'],
        ['上版／交付', '逐行人工審查，排隊等待核准', 'AI 先審並排嚴重度，人判斷風險；沒有授權不能上線', '程式負責人、發布經理'],
        ['維運', '事件常被擱置', '錯誤率異常時 AI 先診斷，必要時退回上一版，寫成新需求', '服務負責人']]
grid(s, 0.45, 1.45, [1.5, 3.2, 5.75, 2.0], ['階段', '導入前（常見做法）', '導入後', '誰決定'], ROWS, 0.66, size=10.5)
tb(s, 0.45, 5.9, 12.45, 0.55, [[('為什麼選這個情境：', {'bold': True, 'color': OR}), ('客服天天被問、業務看得懂、範圍小風險低、成果可量測。', {})]], size=11.5, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.14)
note(s, f'示範情境取自 {PB} 的公開範例，非客戶案例；「導入前」為常見做法的示意。', y=6.5)

# ───────── 18 衡量 ─────────
s = content('怎麼衡量：每個階段一個指標，數據大多已在版本紀錄與 CI 系統中，不需要另外填表')
tag(s, '示意')
ROWS = [['需求分析', '從第一次討論到需求說明定稿的時間', '試點前量測', '縮短'],
        ['SA&D', '開發開始後還要改規格的次數', '試點前量測', '減少'],
        ['開發', '第一次實作就能合併的比例', '試點前量測', '提高'],
        ['測試', 'AI 改動第一次就通過 CI 的比例', '試點前量測', '提高'],
        ['上版／交付', '第一次審查的等待時間、上線失敗率', '試點前量測', '縮短、降低'],
        ['維運', '同類事件重複發生的次數', '試點前量測', '減少']]
grid(s, 0.45, 1.45, [1.6, 5.0, 1.9, 1.9], ['階段', '指標', '基準', '期望方向'], ROWS, 0.6, size=11.5)
tb(s, 10.95, 1.45, 1.95, 4.0, [[('原則', {'bold': True, 'color': OR, 'size': 12.5})], '對外只引用自己試點量到的數字。', [('', {'size': 6})], '手冊所述「數週縮短到數小時」僅為預期，不作承諾。'], size=11, fill=TINT, margin=0.12)
note(s, f'指標定義依 {PB} 各階段「How to measure it」，可搭配 DORA 指標。', y=5.85)

# ───────── 章節 04 ─────────
section('04', '可擴充、可複製\nScalable and repeatable')

# ───────── 20 可重用 vs 客製 ─────────
s = content('可複製的關鍵：模組、資產與流程可以沿用，客製只集中在貴公司自己的規範、系統與核准人')
tag(s, '建議')
COLS = [('可重用模組', ['五階段加維運的循環流程', '五個人工關卡', '三份範本', 'Review Checklist 九項'], OR),
        ('可重用資產', ['保險業規範規則（個資、委外、AI 責任）', '法規檢查清單', '指標定義與教材', 'AI 任務分類（附錄 D）'], OR),
        ('可重用流程', ['診斷方法', '試點陪跑步驟', '成功標準設定', '擴大決策規則'], OR),
        ('需要客製', ['貴公司的資安、命名規範', '核心系統與 Jira、變更管理串接', '資料分級與各關卡核准人', '試點題目、基準與委外安排'], INK)]
for i, (h, items, c) in enumerate(COLS):
    x = 0.45 + i * 3.14
    tb(s, x, 1.45, 3.0, 0.55, h, size=14, bold=True, color=WHITE, fill=c, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    tb(s, x, 2.05, 3.0, 3.0, [[(('✓　' if i < 3 else '●　') + t, {})] for t in items], size=12, fill=TINT if i < 3 else LG, margin=0.16)
callout(s, 5.3, '好處：', '標準部分讓交付快、品質一致；客製部分讓方案符合貴公司規範。每個新系統都從更完整的資產開始，導入更快。', h=0.75)

# ───────── 21 延伸例子 ─────────
s = content('延伸方式：同一套資產，依序延伸到其他系統、新制專案、其他客戶與產業，各有前提')
tag(s, '建議')
ROWS = [['同一客戶的其他系統', '理賠 → 核保規則調整、保單服務', '流程、範本、關卡、規範規則', '該系統的規範與核准人', '前一個試點達標'],
        ['新制相關專案', 'IFRS 17／TW-ICS 報表與資料調整', '流程、關卡、法規清單', '財會與精算規範、資料品質檢查', '財會、精算參與；高風險變更加審'],
        ['AI 治理', 'AI 清冊與模型驗證流程', '紀錄與關卡、自律規範對應', '模型驗證規則', '屬治理服務範疇，另行界定'],
        ['其他保險公司', '另一家壽險或產險', '全部標準模組與資產', '該公司規範與系統串接', '審計客戶需先確認獨立性'],
        ['其他金融業', '銀行、證券', '流程、關卡、指標', '該產業法規規則', '法規清單需重做；以金管會 AI 指引為共同基礎']]
grid(s, 0.45, 1.45, [2.3, 2.85, 2.5, 2.4, 2.4], ['延伸方向', '例子', '沿用什麼', '要新增什麼', '前提／限制'], ROWS, 0.8, size=11)

# ───────── 22 前提與限制 ─────────
s = content('前提與限制：先確認五個前提，並清楚說明這個方法做不到的地方')
tag(s, '建議')
PRE = ['有版本控制與 CI，或願意在試點中建立', '試點範圍有自動測試，或先補上', '可使用的 AI 工具與合約，並完成資料分級', '各關卡有指定的核准人', '管理層支持一個小範圍試點']
LIM = ['缺文件或高度客製的舊系統，AI 產出準確度要在試點中驗證', '是否屬委外、是否需申報，依個案判斷，需貴公司法遵確認', 'AI 工具的使用依貴公司與所內政策決定', '效益依試點實測，不預先承諾數字']
tb(s, 0.45, 1.45, 6.1, 0.5, '導入前提', size=14, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
tb(s, 0.45, 2.0, 6.1, 3.4, [[(f'{i + 1}　', {'bold': True, 'color': OR}), (t, {})] for i, t in enumerate(PRE)], size=12.5, fill=TINT, margin=0.2)
tb(s, 6.8, 1.45, 6.1, 0.5, '限制', size=14, bold=True, color=WHITE, fill=INK, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
tb(s, 6.8, 2.0, 6.1, 3.4, [[('●　' + t, {})] for t in LIM], size=12.5, fill=LG, margin=0.2)
callout(s, 5.6, '做法：', '前提不足的部分列入診斷範圍，先補齊再進試點，避免讓沒把關的東西更快上線。', h=0.6)

# ───────── 章節 05 ─────────
section('05', '如何開始')

# ───────── 24 路線圖 ─────────
s = content('路線圖：從準備度診斷到全公司內化，每一階段結束由貴公司決定是否往下走')
tag(s, '示意')
PH = [('0　準備度診斷', '約 4–6 週', '盤點現況與卡點、法規檢查、挑選試點', '瓶頸地圖、導入藍圖', '選定試點與成功標準'),
      ('1　試點陪跑', '約 6–8 週', '用一個真實需求走完循環，同步教學', '完成的試點、成效報告', '指標達到合格線'),
      ('2　擴大', '依範圍', '推到更多團隊，加入貴公司規範規則', '規範規則庫、擴大路線圖', '管理層核准擴大'),
      ('3　全公司內化', '持續', '集中管理設定、統一關卡與紀錄，逐步自動化', '持續運作的流程', '—')]
for i, (h, d, do, out, gate) in enumerate(PH):
    x = 0.45 + i * 3.14
    tb(s, x, 1.45, 3.0, 0.75, [[(h, {'bold': True, 'size': 14})], [(d, {'size': 10.5})]], color=WHITE, fill=[OR, OR, OR2, MID][i], anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
    if i < 3: arrow(s, x + 3.0, 1.68, 0.14, 0.3)
    tb(s, x, 2.25, 3.0, 1.15, [[('做什麼', {'bold': True, 'color': OR, 'size': 10.5})], do], size=11.5, fill=LG, margin=0.12)
    tb(s, x, 3.45, 3.0, 1.1, [[('交付', {'bold': True, 'color': OR, 'size': 10.5})], out], size=11.5, fill=LG, margin=0.12)
    tb(s, x, 4.6, 3.0, 0.95, [[('進下一階段的條件', {'bold': True, 'color': OR, 'size': 10.5})], gate], size=11.5, fill=TINT, margin=0.12)
note(s, '期程為建議，依範圍確認。自動化在關卡都到位後才進行。', y=5.8)

# ───────── 25 成功標準 ─────────
s = content('成功標準：試點前雙方講好三個指標和合格線，達標才擴大，未達標一起檢討')
tag(s, '建議')
grid(s, 0.45, 1.45, [2.8, 3.6, 1.6, 1.6], ['指標（可依痛點替換）', '怎麼量', '基準', '合格線'], KPI, 0.75, size=11.5)
tb(s, 10.25, 1.45, 2.65, 2.7, [[('結果怎麼用', {'bold': True, 'color': OR, 'size': 12.5})],
                               [('達標：', {'bold': True}), ('擴大到下一個系統或團隊。', {})],
                               [('未達標：', {'bold': True}), ('一起檢討原因，調整後再試或停止，不擴大。', {})]], size=11.5, fill=TINT, margin=0.14)
tb(s, 0.45, 4.45, 12.45, 1.2, [[('擴展方式', {'bold': True, 'color': OR, 'size': 12.5})],
                               '達標後依第 22 頁的延伸順序推進：先同一客戶的其他系統，再到新制相關專案；每次擴展都沿用同一套指標，持續比較。'],
   size=12, fill=LG, margin=0.16)
note(s, '合格線於診斷階段與貴公司共同訂定。', y=5.85)

# ───────── 26 下一步 ─────────
s = content('下一步：三件事，從一場現場示範開始')
NX = [('現場示範', '用貴公司提供的一個需求，示範從需求說明到實作計畫。'),
      ('對齊工作坊', '對齊目標，從需求清單挑 1–2 個候選試點，確認前提與限制。'),
      ('啟動準備度診斷', '確認範圍與雙方窗口，開始盤點現況與法規檢查。')]
for i, (h, t) in enumerate(NX):
    y = 1.5 + i * 1.3
    circle(s, 0.45, y + 0.25, 0.65, str(i + 1), size=18)
    tb(s, 1.3, y, 6.4, 1.15, [[(h, {'bold': True, 'size': 14})], t], size=12, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.16)
tb(s, 8.0, 1.5, 4.9, 3.75, [[('請貴公司提供', {'bold': True, 'color': OR, 'size': 13})],
                            '・IT 主管與業務單位窗口', '・現有開發流程與工具（Jira、CI、變更管理）', '・候選需求清單（示範與試點用）', '・資安、個資與法遵規範'],
   size=12.5, fill=TINT, margin=0.2)
callout(s, 5.55, '一句話：', 'AI 幫你把每一步做得更快，每一步都留下紀錄；要不要做、能不能上線，永遠由你的人決定。', h=0.7, size=12.5)
note(s, '示範與診斷均不使用貴公司真實客戶資料。', y=6.35)

# ───────── 附錄轉場 ─────────
section('A', '附錄')

# ───────── 28 附錄 A 研究整合 ─────────
s = content('附錄 A｜研究整合：兩份研究互補，衝突處的取捨與原因')
ROWS = [['流程階段', '五階段：需求分析、SA&D、開發、測試、上版／交付', '六階段：規劃、設計、開發、測試、上線、維運', '採研究一的名稱，加上維運：保險業 IT 熟悉這些用語；維運讓流程形成循環'],
        ['控制方式', 'Prompt 撰寫原則＋IT 人工 Review', '技能、鉤子、集中管理設定', '兩者分層：Prompt 為第一層，往規則與自動關卡加強；受監理行業需要可強制的控制'],
        ['關卡', '每階段文件確認後進下一階段', '五個關鍵決定由指定的人核准', '保留研究一的檢查點，補上研究二的指定核准人與紀錄'],
        ['效益說法', '未提數字', '手冊預期「數週到數小時」', '不承諾數字，以試點前後實測；手冊數字僅作背景'],
        ['實作證據', '各階段案例頁待補', '實績頁待 IT 團隊填寫', '兩份都缺，列為待確認；以現場示範作為第一步'],
        ['服務內容', '未涵蓋', '四個服務模組', '整合為三個服務套組（第 12 頁）']]
grid(s, 0.45, 1.45, [1.6, 3.3, 3.2, 4.35], ['主題', '研究一：FSC1 開發流程 v3', '研究二：AI 原生 SDLC 說明', '本提案的取捨與原因'], ROWS, 0.7, size=10.5)
note(s, '互補：研究一提供各階段的輸入、產出與 IT 動作；研究二提供治理、法規、成效衡量與服務模組。', y=6.2)

# ───────── 29 附錄 B 競爭來源 ─────────
s = content('附錄 B｜競爭參考：具體產品與來源（業者公開資訊）')
ROWS = [('B1', 'Deloitte', 'Agentic AI Blueprint for Insurance', 'https://www.deloitte.com/global/en/what-we-do/capabilities/agentic-ai-blueprint-for-insurance.html'),
        ('B2', 'Deloitte', 'IFRS 17 Regulatory Assist（與 SAS 合作）', 'https://sas.com/nl_nl/news/press-releases/2022/may/deloitte-partners-with-sas-to-power-ifrs-17-accelerator-solution.html'),
        ('B3', 'KPMG', 'IFRS 17 Express, powered by SAS', 'https://sas.com/en_ca/news/press-releases/2022/september/pvi-insurance-deploys-kpmg-ifrs-17-express.html'),
        ('B4', 'Capgemini', 'Guidewire 生成式 AI／代理工具組', 'https://www.capgemini.com/wp-content/uploads/2026/10/GW-datasheet_Final-version.pdf'),
        ('B5', 'WTW', 'Radar 5（保險定價與核保軟體）', 'https://www.wtwco.com/en-gh/news/2025/10/wtw-launch-of-radar-5-with-gen-ai-capability-marks-a-major-milestone-in-insurance-pricing'),
        ('B6', 'McKinsey', 'Global Insurance Pools（訂閱資料庫）', 'https://www.mckinsey.com/industries/financial-services/how-we-help-clients/insurance/global-insurance-pools'),
        ('B7', 'McKinsey', '代理型 AI 核心現代化', 'https://www.mckinsey.com/industries/financial-services/our-insights/can-agentic-ai-finally-modernize-core-technologies-in-insurance'),
        ('B8', 'BCG', 'AI-First 產險公司藍圖', 'https://www.bcg.com/assets/2026/executive-perspectives-ai-first-companies-win-the-future-property-casualty-insurance.pdf')]
xx = [0.45, 1.15, 2.7, 7.0]
for j, h in enumerate(['編號', '公司', '產品／方案', '來源（可點擊）']):
    tb(s, xx[j], 1.45, (xx[j + 1] if j < 3 else 12.9) - xx[j] - 0.04, 0.36, h, size=10.5, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.07)
for i, (n, c, prod, url) in enumerate(ROWS):
    y = 1.85 + i * 0.5; bg = LG if i % 2 == 0 else WHITE
    tb(s, xx[0], y, 0.66, 0.46, n, size=10.5, bold=True, color=OR, fill=bg, anchor=MSO_ANCHOR.MIDDLE, margin=0.07)
    tb(s, xx[1], y, 1.51, 0.46, c, size=10.5, bold=True, fill=bg, anchor=MSO_ANCHOR.MIDDLE, margin=0.07)
    tb(s, xx[2], y, 4.26, 0.46, prod, size=10.5, fill=bg, anchor=MSO_ANCHOR.MIDDLE, margin=0.07)
    tb(s, xx[3], y, 5.86, 0.46, [[(url.split('/')[2] + ' 原文', {'link': url})]], size=10.5, fill=bg, anchor=MSO_ANCHOR.MIDDLE, margin=0.07)
note(s, '對客戶提案時是否保留本頁公司名稱，由提案合夥人決定；規模與期程數字為業者自述。', y=6.0)

# ───────── 30 附錄 C 法規檢查清單 ─────────
checklist_page('附錄 C｜保險業 AI 開發法規檢查清單（草稿）',
               '導入規劃工具，非法律意見；條文以主管機關公告為準；草稿待法遵確認。')

# ───────── 31 附錄 D AI 任務分類 ─────────
s = content('附錄 D｜AI 任務分類：用「多常發生 × 需要判斷」決定哪些工作先交給 AI、先標準化')
MX, MY, QW, QH = 1.15, 1.5, 3.6, 2.15
Q = [(0, 0, '常發生 × 判斷少', 'AI 先做，人確認', '欄位清單、Test Case、Test Data、Unit Test、註解與文件、Release Note', TINT),
     (1, 0, '常發生 × 判斷多', 'AI 起草，人逐項審查', '程式碼、API 規格、ERD、Code Review 風險檢查、測試失敗原因分析', LG),
     (0, 1, '少發生 × 判斷少', '做成範本，重複使用', 'Deploy Guide、User Guide、Migration Script、部署腳本', LG),
     (1, 1, '少發生 × 判斷多', '人決定，AI 只提供資料', '接受需求、架構選型、核准合併、授權上線、事件處置', WHITE)]
for c, r, h, rule, items, bg in Q:
    x, y = MX + c * (QW + 0.06), MY + r * (QH + 0.06)
    tb(s, x, y, QW, QH, '', fill=bg, line=LINE)
    tb(s, x + 0.12, y + 0.08, QW - 0.24, 0.62, [[(h, {'bold': True, 'size': 11, 'color': GREY})], [(rule, {'bold': True, 'size': 13, 'color': OR})]], margin=0.02)
    tb(s, x + 0.12, y + 0.76, QW - 0.24, QH - 0.84, items, size=11.5, margin=0.02)
tb(s, 0.45, MY, 0.65, QH, '常發生', size=11, bold=True, color=GREY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.0)
tb(s, 0.45, MY + QH + 0.06, 0.65, QH, '少發生', size=11, bold=True, color=GREY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, margin=0.0)
tb(s, MX, MY + 2 * QH + 0.1, 2 * QW + 0.06, 0.3, '需要判斷：少 ─────────────→ 多', size=10.5, bold=True, color=GREY, align=PP_ALIGN.CENTER, margin=0.0)
tb(s, 8.75, 1.5, 4.15, 4.4, [[('怎麼用', {'bold': True, 'color': OR, 'size': 13})],
                             '・左上角最先交給 AI，也最先做成標準範本與規則。', '・右上角由 AI 起草，人逐項審查。', '・右下角永遠由人決定，對應五個人工關卡。',
                             [('', {'size': 6})], [('依據：', {'bold': True, 'color': OR}), (f'任務取自研究一各階段「AI 可以做什麼」；方法參考 {SMR}。', {})]], size=11.5, fill=LG, margin=0.16)
note(s, '分類為初步判斷，診斷時依貴公司實際專案調整。', y=6.45)

# ───────── 32 附錄 E 資料來源 ─────────
s = content('附錄 E｜資料來源與標示說明')
tb(s, 0.45, 1.45, 12.45, 0.7, [[('標示：', {'bold': True, 'color': OR}), ('建議＝我們的提案；示意＝說明用，非實測或報價；假設＝需訪談確認；待確認＝需內部或客戶提供。未標示者為有來源的事實或既有設計。', {})]], size=11.5, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.16)
SRCS = [(f'研究一：{V3}（19 頁）', None), (f'研究二：AI 原生 SDLC 白話說明（25 頁），整理自 Louis Claxton，{PB}，2026', None),
        ('金管會保險局，IFRS 17／TW-ICS 專區', None),
        ('Claims Journal，PwC 保險高階主管調查（57%），2025-11-17', 'https://amp.claimsjournal.com/news/national/2025/11/17/334094.htm'),
        ('Bain，Reimagining Insurance for the AI Era（引述 MIT 2025：95%）', 'https://www.bain.com/ko/results/reimagining-insurance-for-the-ai-era'),
        ('理律法律事務所，保險業作業委外新辦法（2025-12-31 施行）', 'https://www.leeandli.com/TW/Newsletters/7525.htm'),
        ('壽險／產險公會，保險業運用人工智慧系統自律規範（2025-04-10 備查）', 'https://law.lia-roc.org.tw/Law/GetFile/0000394231'),
        ('壽險／產險公會，保險業作業委外使用雲端服務自律規範', 'https://law.lia-roc.org.tw/Law/GetFile/0000419987'),
        ('金管會，金融業運用人工智慧（AI）指引，2024-06', 'https://www.fsc.gov.tw/websitedowndoc?file=chfsc/202408231741530.pdf'),
        (f'Mohanbir Sawhney，How to Turn Professional Services Into Products，{SMR}', None)]
for i, (lab, url) in enumerate(SRCS):
    seg = (lab, {'link': url}) if url else (lab, {})
    tb(s, 0.45, 2.3 + i * 0.42, 12.45, 0.38, [[(f'{i + 1}. ', {'bold': True, 'color': OR}), seg]], size=10.5, fill=LG if i % 2 == 0 else WHITE, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)

p.save(OUT)
print(OUT, len(p.slides), 'slides')
