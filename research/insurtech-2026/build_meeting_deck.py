# 保險科技服務優先順序：會議決策簡報（顧問式故事線，以所內簡報範本建立）
# 方法：mckinsey-consultant 技能（問題定義 → Issue Tree → 假設 → 論點式標題 → 決策頁）
# 用法：python3 build_meeting_deck.py <範本.pptx> <輸出.pptx>
import sys, copy
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

TPL, OUT = sys.argv[1], sys.argv[2]
OR = RGBColor(0xFD, 0x51, 0x08); OR2 = RGBColor(0xFE, 0x7C, 0x39); TINT = RGBColor(0xFF, 0xE8, 0xD4)
# 範本佈景主題色：tx1 黑、lt2 淺灰、accent6 線條灰
INK = RGBColor(0x00, 0x00, 0x00); GREY = RGBColor(0x59, 0x59, 0x59); LG = RGBColor(0xEB, 0xEB, 0xEB); LINE = RGBColor(0xCB, 0xD1, 0xD6)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
# 範本字型：標題 Georgia、內文 Arial、中文一律微軟正黑體（與所內範本範例頁一致）
FONT = 'Arial'; HEAD_FONT = 'Georgia'; EA_FONT = '微軟正黑體'

SYMBOLS = set('●◐○✓✗→①②③④⑤⚠↑')

def style_font(r, head=False):
    # Arial 沒有 ●◐○ 等符號，這類文字改用中文字型顯示
    r.font.name = EA_FONT if SYMBOLS & set(r.text) else (HEAD_FONT if head else FONT)
    rPr = r._r.get_or_add_rPr()
    for e in rPr.findall(qn('a:ea')): rPr.remove(e)
    rPr.find(qn('a:latin')).addnext(rPr.makeelement(qn('a:ea'), {'typeface': EA_FONT}))
FOOT = '內部討論稿｜僅供 FSC1 保險組內部使用｜資料來源為公開研究，數字待回原文核對'

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

def footer(s):
    # python-pptx 新增投影片不會帶入頁尾與頁碼，從版面複製（頁碼為自動欄位）
    have = {ph.placeholder_format.idx for ph in s.placeholders}
    for ph in s.slide_layout.placeholders:
        if ph.placeholder_format.idx in (3, 4) and ph.placeholder_format.idx not in have:
            s.shapes._spTree.append(copy.deepcopy(ph._element))
    set_ph(s, 3, FOOT)

def tb(s, x, y, w, h, text, size=11, bold=False, color=INK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, fill=None, line=None, margin=0.06):
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
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
        para.space_after = Pt(3)
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
    return ln

def arrow(s, x, y, w, h, kind=MSO_SHAPE.RIGHT_ARROW, fill=OR):
    ar = s.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    ar.fill.solid(); ar.fill.fore_color.rgb = fill; ar.line.fill.background(); ar.shadow.inherit = False

def table(s, x, y, w, col_w, rows, size=9.5, head_size=10, row_h=0.3, first_col_bold=True):
    nr, nc = len(rows), len(rows[0])
    gt = s.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(row_h * nr)).table
    for j, cw in enumerate(col_w): gt.columns[j].width = Inches(cw)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            c = gt.cell(i, j); c.text = ''
            c.margin_left = c.margin_right = Inches(0.06); c.margin_top = c.margin_bottom = Inches(0.035)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            c.fill.solid(); c.fill.fore_color.rgb = OR if i == 0 else (LG if i % 2 == 0 else WHITE)
            tf = c.text_frame; tf.word_wrap = True
            parts = val if isinstance(val, list) else [val]
            for k, part in enumerate(parts):
                para = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                r = para.add_run(); r.text = part; style_font(r)
                r.font.size = Pt(head_size if i == 0 else size)
                r.font.bold = i == 0 or (j == 0 and first_col_bold)
                r.font.color.rgb = WHITE if i == 0 else (INK if not (j == 0 and first_col_bold) else OR)
    return gt

def content(title):
    s = p.slides.add_slide(L['只有標題'])
    set_ph(s, 0, title, 20, bold=True); footer(s)
    return s

def section(num, title, desc):
    s = p.slides.add_slide(L['章節標題'])
    set_ph(s, 0, title, 30); set_ph(s, 1, num, 180); footer(s)
    tb(s, 0.45, 6.68, 8.5, 0.35, desc, size=10.5, color=GREY)
    return s

def source(s, text):
    tb(s, 0.45, 6.68, 12.4, 0.3, text, size=9, color=GREY)

def takeaway(s, x, y, w, h, text, size=12):
    tb(s, x, y, 0.1, h, '', fill=OR)
    tb(s, x + 0.1, y, w - 0.1, h, [[('所以呢：', {'bold': True, 'color': OR}), (text, {})]], size=size, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.09)

# ───────── 1 封面 ─────────
s = p.slides.add_slide(L['標題投影片'])
set_ph(s, 0, '保險 AI 進入規模化落地期：\nFSC1 建議先推兩項服務包', 28)
set_ph(s, 1, 'FSC1 保險組｜服務優先順序討論會｜2026 年 10 月', 12)

# ───────── 2 問題定義 ─────────
s = content('今天要回答一個問題、做四個決定：先推哪兩項服務、誰負責、90 天內交付什麼')
tb(s, 0.45, 1.4, 12.45, 0.8, [[('核心問題　', {'bold': True, 'size': 13}), ('未來 12 個月，FSC1 保險組應優先把哪些保險科技機會做成可銷售的服務，並在 2026 年的需求窗口內拿下首案？', {})]],
   size=15, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.2)
cols = [('今天討論的範圍 ✓', OR, ['全球與亞太過去 18 個月的保險科技變化，對台灣壽險與產險的意涵',
                                 '四條既有服務線（AI 治理、數據應用、核心系統、人力委外）可以承接的機會',
                                 '第一波服務包的切入方式、負責人與 90 天行動計畫']),
        ('不在今天範圍 ✗', INK, ['個別客戶的提案內容與報價（另案處理）',
                                 '自建保險科技產品或對外投資',
                                 '對外發布研究報告（本稿數字尚待回原文核對）'])]
for i, (h, c, items) in enumerate(cols):
    x = 0.45 + i * 6.3
    tb(s, x, 2.45, 6.05, 0.45, h, size=13, bold=True, color=c, anchor=MSO_ANCHOR.MIDDLE)
    hline(s, x, 2.92, x + 6.05, 2.92, color=c, width=1.5)
    for k, t in enumerate(items):
        tb(s, x, 3.02 + k * 0.62, 6.05, 0.58, '・' + t, size=12.5, anchor=MSO_ANCHOR.MIDDLE)
tb(s, 0.45, 5.0, 12.45, 0.4, '今天請決定', size=13, bold=True, color=OR)
decs = ['第一波推哪兩項服務', '每項服務包的負責人', '90 天計畫與檢查點', '研討會 AI 短片的主題']
for i, t in enumerate(decs):
    x = 0.45 + i * 3.15
    circle(s, x, 5.5, 0.5, str(i + 1), size=13)
    tb(s, x + 0.6, 5.45, 2.4, 0.6, t, size=13, bold=True, anchor=MSO_ANCHOR.MIDDLE, fill=LG, margin=0.1)
source(s, '建議方案見第 3 頁（重點摘要）與第 19 頁（決策事項）。')

# ───────── 3 重點摘要（SCQA） ─────────
s = content('重點摘要：需求已到、窗口在 2026 年，建議以 IFRS 17 與 AI 治理兩項服務包先行')
scq = [('情境', '保險業 AI 已從試點走向規模化：2026 Q1 全球保險科技募資 95.2% 流向 AI；63% 企業已把 AI 投入營運。'),
       ('衝突', '但 95% 投資 AI 的企業還看不到財務回報；台灣同時面對 IFRS 17／TW-ICS 首年出表與 AI 指引落地查核。'),
       ('問題', 'FSC1 該先把哪些機會做成服務，才能在 2026 年的窗口內拿下首案？')]
for i, (h, t) in enumerate(scq):
    y = 1.4 + i * 1.72
    tb(s, 0.45, y, 3.7, 0.4, h, size=12, bold=True, color=WHITE, fill=INK, anchor=MSO_ANCHOR.MIDDLE, margin=0.04)
    tb(s, 0.45, y + 0.4, 3.7, 1.22, t, size=11.5, fill=LG, margin=0.12)
arrow(s, 4.25, 3.4, 0.35, 0.5)
tb(s, 4.7, 1.4, 8.2, 1.05, [[('建議　', {'bold': True, 'size': 12}), ('第一波先推 2 項服務包：① IFRS 17／TW-ICS 後財務轉型　② AI 清冊、模型驗證與護欄。另 3 項列為第二波。', {})]],
   size=15, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.18)
args = [('需求確定、有時限', ['2026-01-01 正式接軌，首年結帳仍高度人工', '金管會 AI 指引與公會自律規範帶來查核壓力']),
        ('離既有服務最近', ['延伸既有駐點團隊與 AI 治理服務', '不需新增技術夥伴，可用短期診斷或健檢切入（AI 治理健檢 4–6 週）']),
        ('能帶出長期合作', ['治理與財務資料是代理型 AI、核心現代化的前置條件', '可延伸為第二波三項服務與 1–3 年人力委外'])]
for i, (h, bl) in enumerate(args):
    x = 4.7 + i * 2.76
    hline(s, x + 1.3, 2.45, x + 1.3, 2.7, color=OR, width=1.5)
    tb(s, x, 2.7, 2.62, 0.55, h, size=13, bold=True, color=OR, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    tb(s, x, 3.3, 2.62, 2.0, ['・' + b for b in bl], size=11.5, fill=LG, margin=0.12)
tb(s, 4.7, 5.5, 8.2, 0.95, [[('今天請決定：', {'bold': True, 'color': OR})], '同意第一波兩項服務、指派各 1 位負責人、核定 90 天計畫與第 4／8／12 週檢查點。'],
   size=12, fill=TINT, margin=0.14)
source(s, '來源：Gallagher Re Global InsurTech Report Q1 2026（2026-05）；Bain 引述 MIT（2026-01）；金管會公告。數字待回原文核對。')

# ───────── 4 Issue Tree ─────────
s = content('我們把問題拆成三個子問題：需求在哪裡、我們能不能贏、怎麼開始')
for x, w, h in [(0.45, 2.3, '核心問題'), (3.0, 2.6, '子問題'), (5.85, 3.3, '要回答的問題'), (9.3, 2.75, '初步答案（假設）'), (12.1, 0.8, '狀態')]:
    tb(s, x, 1.35, w, 0.32, h, size=10.5, bold=True, color=GREY)
tb(s, 0.45, 1.75, 2.3, 4.74, '未來 12 個月，FSC1 應優先把哪些保險科技機會做成可銷售的服務？', size=14, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.15)
TREE = [('A　需求', '哪些機會的需求明確且有時限？',
         [('全球資金與預算往哪裡走？', '預算全面轉向 AI，但落地回報不足', '✓', 'p.6'),
          ('台灣的監理觸發點何時發生？', '2026 年是 IFRS 17 與 AI 治理雙重觸發年', '✓', 'p.7')]),
        ('B　能力', '哪些機會我們比對手更有勝算？',
         [('與四條既有服務線距離多遠？', '九項趨勢中五項可直接承接', '✓', 'p.9–11'),
          ('對手的做法有哪些侷限？', '對手重藍圖與平台，缺在地確信與駐點人力', '◐', 'p.15')]),
        ('C　路徑', '如何最快拿下首案並擴大？',
         [('用什麼小案切入？', '4–8 週健檢、診斷或 PoC 可降低決策門檻', '?', 'p.14'),
          ('如何擴大成長期合作？', '從顧問藍圖接上人力委外，形成 1–3 年經常性營收', '?', 'p.14')])]
hline(s, 2.88, 2.5, 2.88, 5.74, color=OR, width=1.5)
for i, (b, q, subs) in enumerate(TREE):
    y = 1.75 + i * 1.62
    hline(s, 2.75, y + 0.75, 3.0, y + 0.75, color=OR, width=1.5)
    tb(s, 3.0, y, 2.6, 1.5, [[(b, {'bold': True, 'color': OR, 'size': 13})], q], size=12, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
    hline(s, 5.72, y + 0.37, 5.72, y + 1.13, color=GREY)
    for k, (sq, hy, st, ref) in enumerate(subs):
        yy = y + k * 0.78
        hline(s, 5.6, yy + 0.37, 5.85, yy + 0.37, color=GREY)
        tb(s, 5.85, yy, 3.3, 0.72, sq, size=11.5, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
        tb(s, 9.3, yy, 2.75, 0.72, [hy, [(ref, {'size': 9, 'color': GREY})]], size=11, line=LINE, anchor=MSO_ANCHOR.MIDDLE, margin=0.08)
        circle(s, 12.25, yy + 0.13, 0.46, st, fill={'✓': OR, '◐': OR2, '?': GREY}[st], size=13)
source(s, '狀態：✓ 研究資料支持　◐ 部分支持（公開資料推論）　? 待驗證（需客戶訪談）。分類依 MECE 原則：需求（市場）× 能力（我們）× 路徑（執行）。')

# ───────── 第一章 需求 ─────────
section('01', '需求：預算已轉向 AI，\n台灣 2026 年出現監理窗口', '回答子問題 A：哪些機會的需求明確且有時限？')

s = content('資金與預算已全面轉向 AI，但 95% 企業還看不到回報：落地缺口就是服務需求')
stats = [('95.2%', '2026 Q1 全球保險科技募資流向 AI 相關企業的比例', 'Gallagher Re, 2026-05'),
         ('63%', '2025 年底已把 AI 投入實際營運的企業（2024 年為 45%）', 'Gallagher Re, 2026-05'),
         ('57%', '保險高階主管把生成式與代理型 AI 列為 2026 首要科技投資', 'PwC 調查／Insurance Journal, 2025-11'),
         ('95%', '投資 AI 但尚未看到可衡量財務回報的企業', 'Bain 引述 MIT, 2026-01'),
         ('25–50%', '用代理型 AI 做核心系統現代化的生產力提升', 'McKinsey, 2026-04'),
         ('92%', '曾因深偽相關事件遭受財務損失的企業', 'Swiss Re SONAR 2025, 2025-06')]
for i, (n, d, src) in enumerate(stats):
    x = 0.45 + (i % 3) * 4.2; y = 1.4 + (i // 3) * 2.2
    hot = i == 3
    tb(s, x, y, 3.95, 2.05, '', fill=OR if hot else LG)
    tb(s, x + 0.2, y + 0.1, 3.6, 0.8, n, size=36, color=WHITE if hot else OR)
    tb(s, x + 0.2, y + 0.92, 3.6, 0.75, d, size=12, color=WHITE if hot else INK)
    tb(s, x + 0.2, y + 1.68, 3.6, 0.3, '來源：' + src, size=9, color=WHITE if hot else GREY)
takeaway(s, 0.45, 5.88, 12.45, 0.7, '客戶缺的不是 AI 工具，而是讓 AI 進入流程、通過查核、產生回報的落地能力；這正是四條服務線的交集。')

s = content('台灣 2026 年同時出現三個需求觸發點，合規與轉型需求將集中在未來 12–18 個月')
EV = [('2024-08', '金管會《金融業運用 AI 指引》', True), ('2024-08', '保險業作業委外規範修正（雲端鬆綁）', True),
      ('2025-10', '內部作業制度辦法草案預告（提升法令位階）', True), ('2025-11', 'MAS AI 風險管理指引（區域標竿）', False),
      ('2026-01-01', 'IFRS 17／TW-ICS 正式接軌', True), ('2026', 'NAIC 12 州試點 AI 系統評估工具（美國標竿）', False)]
hline(s, 0.6, 2.35, 12.75, 2.35, color=OR, width=2.5)
for i, (d, t, tw) in enumerate(EV):
    x = 0.55 + i * 2.06
    circle(s, x + 0.82, 2.2, 0.3, '', fill=OR if tw else GREY)
    tb(s, x, 1.55, 1.95, 0.5, d, size=13, bold=True, color=OR if tw else GREY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM)
    tb(s, x, 2.65, 1.95, 0.95, t, size=11, align=PP_ALIGN.CENTER, fill=LG if tw else None, line=None if tw else LINE, margin=0.08)
tb(s, 0.45, 3.85, 12.45, 0.4, '三個觸發點 → 對應的服務需求', size=13, bold=True, color=OR)
TRIG = [('①', 'IFRS 17／TW-ICS 首年出表', '結帳仍高度人工；覆蓋法退場造成損益波動', 'IFRS 17 後財務轉型、CSM 與動態 ALM'),
        ('②', 'AI 指引與自律規範進入查核', '多數只有紙本政策，缺模型驗證與 AI 清冊', 'AI 清冊、模型驗證與執行期護欄'),
        ('③', '雲端委外鬆綁並提升法令位階', '舊核心無法支援 API 與 AI 串接', 'AI 加速核心現代化與雲端聯合查核')]
for i, (n, a, b, c) in enumerate(TRIG):
    y = 4.32 + i * 0.75
    circle(s, 0.45, y + 0.08, 0.5, n, size=13)
    tb(s, 1.1, y, 3.5, 0.66, a, size=12, bold=True, anchor=MSO_ANCHOR.MIDDLE, fill=TINT, margin=0.1)
    tb(s, 4.7, y, 4.4, 0.66, b, size=11.5, anchor=MSO_ANCHOR.MIDDLE, fill=LG, margin=0.1)
    arrow(s, 9.2, y + 0.18, 0.4, 0.3)
    tb(s, 9.7, y, 3.2, 0.66, c, size=11.5, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, fill=OR, margin=0.1)
source(s, '橘色為台灣監理事件，灰色為國際標竿。「12–18 個月」為團隊依首年出表與查核時程的判斷。來源：金管會公告、MAS、NAIC（2026-04）、工商時報。')

# ───────── 第二章 機會 ─────────
section('02', '機會：九項趨勢\n篩選出五項優先服務', '回答子問題 B：哪些機會離我們最近、我們最有勝算？')

s = content('用兩道篩選，把九項趨勢收斂成五項優先服務，其中兩項可立即推動')
NINE = ['1 代理型 AI 核保與理賠', '2 AI 加速核心現代化與雲端', '3 營運化 AI 治理與護欄', '4 IFRS 17／TW-ICS 後財務轉型',
        '5 深偽理賠反詐欺', '6 AI 精算定價與數據平台', '7 嵌入式保險 API 平台', '8 參數型氣候巨災保險', '9 AI 表現保證保險']
tb(s, 0.45, 1.4, 3.3, 0.5, '九項趨勢', size=14, bold=True, color=WHITE, fill=INK, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
for i, t in enumerate(NINE):
    tb(s, 0.45, 1.98 + i * 0.47, 3.3, 0.42, t, size=11, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
FIL = [(3.9, '篩選一', ['台灣有明確的監理或損益壓力', '與四條既有服務線的距離近']),
       (8.55, '篩選二', ['首案週期短（≤ 8 週）', '不需新增技術夥伴', '可先處理獨立性問題'])]
for x, h, crit in FIL:
    arrow(s, x, 3.35, 0.95, 0.7, fill=OR2)
    tb(s, x - 0.05, 4.15, 1.05, 0.35, h, size=11, bold=True, color=OR, align=PP_ALIGN.CENTER)
    tb(s, x - 0.05, 4.5, 1.05, 1.6, ['・' + c for c in crit], size=9.5, color=GREY)
FIVE = ['4 IFRS 17／TW-ICS 後財務轉型', '3 營運化 AI 治理與護欄', '1 代理型 AI 核保與理賠', '2 AI 加速核心現代化與雲端', '5 深偽理賠反詐欺']
tb(s, 5.0, 1.4, 3.4, 0.5, '五項優先服務', size=14, bold=True, color=WHITE, fill=OR2, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
for i, t in enumerate(FIVE):
    tb(s, 5.0, 1.98 + i * 0.52, 3.4, 0.46, t, size=11, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
tb(s, 5.0, 4.68, 3.4, 1.85, [[('未入選（6–9）', {'bold': True})], '6、7：需求分散，要個案診斷', '8、9：台灣市場仍在概念或早期採用，需求待驗證', '→ 持續觀察，客戶主動詢問時再接'],
   size=10.5, color=GREY, line=LINE, margin=0.1)
tb(s, 9.65, 1.4, 3.25, 0.5, '第一波：立即推動', size=14, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
for i, t in enumerate(FIVE[:2]):
    tb(s, 9.65, 1.98 + i * 0.6, 3.25, 0.52, t, size=12, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
tb(s, 9.65, 3.3, 3.25, 1.5, [[('第二波（1、2、5）', {'bold': True})], '1：需先有數據與治理基礎', '2：週期長，需雲端業者配合', '5：需整合技術夥伴'],
   size=10.5, color=INK, fill=LG, margin=0.1)
takeaway(s, 9.65, 4.95, 3.25, 1.58, '第一波兩項是第二波的地基：先拿到財務資料與 AI 治理的入口。', size=11.5)
source(s, '趨勢定義與案例見附錄 A1–A3。篩選準則為團隊判斷，請會議校準。')

s = content('五項優先服務都落在「需求高 × 離既有能力近」的象限，IFRS 17 與 AI 治理最靠近右上角')
X0, Y0, W, H = 1.2, 1.45, 7.6, 4.85
tb(s, X0, Y0, W / 2, H / 2, '', fill=LG); tb(s, X0 + W / 2, Y0, W / 2, H / 2, '', fill=TINT)
tb(s, X0, Y0 + H / 2, W / 2, H / 2, '', fill=WHITE, line=LINE); tb(s, X0 + W / 2, Y0 + H / 2, W / 2, H / 2, '', fill=LG)
for qx, qy, t, c in [(0, 0, '找夥伴或觀察', GREY), (1, 0, '優先推動', OR), (0, 1, '持續觀察', GREY), (1, 1, '順帶銷售', GREY)]:
    tb(s, X0 + qx * W / 2 + 0.1, Y0 + qy * H / 2 + 0.08, 2.5, 0.35, t, size=12, bold=True, color=c)
hline(s, X0, Y0 + H, X0 + W, Y0 + H, color=INK, width=1.25); hline(s, X0, Y0, X0, Y0 + H, color=INK, width=1.25)
tb(s, X0, Y0 + H + 0.05, W, 0.32, '與既有服務的距離　遠 → 近', size=11, color=GREY, align=PP_ALIGN.RIGHT)
tb(s, 0.3, Y0 + H / 2 - 0.3, 0.85, 0.62, '需求\n迫切度', size=10.5, color=GREY, align=PP_ALIGN.CENTER)
tb(s, 0.3, Y0 - 0.02, 0.85, 0.3, '高', size=10.5, color=GREY, align=PP_ALIGN.CENTER)
tb(s, 0.3, Y0 + H - 0.3, 0.85, 0.3, '低', size=10.5, color=GREY, align=PP_ALIGN.CENTER)
POS = [('4', 0.92, 0.92, OR), ('3', 0.84, 0.80, OR), ('1', 0.72, 0.74, OR2), ('2', 0.66, 0.60, OR2), ('5', 0.56, 0.62, OR2),
       ('6', 0.72, 0.34, GREY), ('7', 0.42, 0.36, GREY), ('8', 0.22, 0.28, GREY), ('9', 0.13, 0.13, GREY)]
for n, fx, fy, c in POS:
    d = 0.52 if c == OR else 0.44
    circle(s, X0 + fx * W - d / 2, Y0 + (1 - fy) * H - d / 2, d, n, fill=c, size=13 if c == OR else 11)
tb(s, 9.2, 1.4, 3.7, 0.4, '圖例', size=12, bold=True, color=OR)
for i, t in enumerate(NINE):
    n = t.split(' ')[0]
    c = OR if n in '43' else (OR2 if n in '125' else GREY)
    circle(s, 9.2, 1.88 + i * 0.4, 0.3, n, fill=c, size=9)
    tb(s, 9.6, 1.83 + i * 0.4, 3.3, 0.38, t[2:], size=10.5, anchor=MSO_ANCHOR.MIDDLE)
takeaway(s, 9.2, 5.55, 3.7, 0.95, '右上角兩項不需新能力即可開賣；其餘三項靠第一波累積的資料與信任推進。', size=11)
source(s, '位置為團隊依研究報告與既有服務經驗的判斷，非量化評分；請會議校準。')

s = content('五項準則綜合比較：IFRS 17 後財務轉型與 AI 治理在需求、距離與首案週期都最有利')
F, M, N = '●', '◐', '○'
rows = [['服務', '需求迫切度', '與既有服務距離', '法規與市場時機', '首案週期', '獨立性風險', '建議'],
        ['4　IFRS 17／TW-ICS 後財務轉型', [F, '首年正式出表'], [F, '延伸既有駐點'], [F, '覆蓋法退場'], [F, '結帳瓶頸診斷'], [M, '審計客戶需評估'], '第一波'],
        ['3　AI 清冊、模型驗證與護欄', [F, '查核壓力'], [F, 'AI 治理升級'], [F, '金管會指引上路'], [F, '4–6 週健檢'], [M, '確信角色需 R&Q'], '第一波'],
        ['1　代理型 AI 核保與理賠', [F, '突破單點助理'], [M, '需數據平台'], [F, '全球轉向規模化'], [M, '6–8 週藍圖'], [F, '低'], '第二波'],
        ['2　AI 加速核心現代化與雲端', [M, '核心老舊'], [M, '核心導入＋委外'], [F, '委外規範放寬'], [N, '導入期長'], [M, '聯合查核需確認'], '第二波'],
        ['5　深偽理賠反詐欺', [M, '損率壓力'], [N, '需技術夥伴'], [M, '行動理賠普及'], [M, 'PoC'], [F, '低'], '第二波']]
gt = table(s, 0.45, 1.4, 12.45, [3.25, 1.6, 1.65, 1.65, 1.5, 1.65, 1.15], rows, size=10.5, head_size=11.5, row_h=0.78)
for i in range(1, 6):
    for j in range(1, 6):
        paras = gt.cell(i, j).text_frame.paragraphs
        for k, para in enumerate(paras):
            para.alignment = PP_ALIGN.CENTER
            if k == 0:
                for r in para.runs: r.font.size = Pt(18); r.font.color.rgb = OR
    c = gt.cell(i, 6); first = i <= 2
    c.fill.solid(); c.fill.fore_color.rgb = OR if first else TINT
    for para in c.text_frame.paragraphs:
        para.alignment = PP_ALIGN.CENTER
        for r in para.runs: r.font.bold = True; r.font.size = Pt(12); r.font.color.rgb = WHITE if first else OR
takeaway(s, 0.45, 6.2, 12.45, 0.45, '兩項第一波服務唯一的黃燈都是獨立性，會議後第一週就要找 R&Q 確認。', size=11.5)
source(s, '● 有利　◐ 中等　○ 不利（獨立性風險欄：● 代表風險低）。評等為團隊判斷。')

# ───────── 第三章 建議 ─────────
section('03', '建議：第一波兩項服務包，\n第二波三項', '回答子問題 C：用什麼小案切入、如何擴大成長期合作？')

s = content('第一波兩項服務包：用短期診斷或健檢切入，再延伸為財務與治理的長期合作')
LAB = ['目標買方', '客戶痛點', '切入案（Land）', '延伸（Expand）', '我們的優勢', '要先確認']
PK = [('① IFRS 17／TW-ICS 後財務轉型', ['財務長、精算長、會計主管',
                                        '首年結帳仰賴人工與外包；覆蓋法退場使損益波動；通路仍以初年度保費考核',
                                        '首年結帳瓶頸診斷＋CSM 管理儀表板',
                                        '結帳自動化、以 KRD 為核心的動態 ALM、通路 CSM KPI、精算財會人力委外',
                                        '延伸既有駐點團隊；IFRS 17 資料庫可串 ALM、定價與通路 KPI，科技廠商難以跨過',
                                        '與所內 IFRS 17、審計團隊的分工與獨立性']),
      ('② AI 清冊、模型驗證與護欄', ['風控長、法遵長、資訊長',
                                    '多為紙本政策，缺跨精算、核保、IT 的模型驗證，看不到部門自用的外部 AI 工具',
                                    '4–6 週 AI 清冊盤點＋金管會指引合規健檢',
                                    '治理平台、模型驗證機制、代理變數偏見測試、緊急關閉機制、LLM 紅隊演練',
                                    'Assurance-by-Design：同一個案子同時交付創新與可稽核軌跡',
                                    '確信類服務的獨立性（R&Q）'])]
for r, lab in enumerate(LAB):
    y = 2.0 + r * 0.76
    tb(s, 0.45, y, 1.65, 0.7, lab, size=11.5, bold=True, color=OR, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
for c, (h, cells) in enumerate(PK):
    x = 2.2 + c * 5.37
    tb(s, x, 1.4, 5.27, 0.52, h, size=14, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
    for r, t in enumerate(cells):
        y = 2.0 + r * 0.76
        hot = r == 2
        tb(s, x, y, 5.27, 0.7, t, size=11.5, bold=hot, fill=LG if r % 2 == 0 else WHITE, line=OR if hot else LINE, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
source(s, '切入案週期與內容取自研究報告建議；報價區間由負責人在第 4 週提出。')

s = content('五項服務都走「先小案、後延伸」，第二波三項以第一波的數據與治理成果為基礎')
lands = [('第一波', 'IFRS 17 後財務轉型', '首年結帳瓶頸診斷＋CSM 管理儀表板', 'KRD 動態 ALM、通路 CSM KPI、常態精算財會人力委外'),
         ('第一波', 'AI 治理', '4–6 週 AI 清冊盤點與合規落差健檢', '治理平台、代理變數偏見測試、緊急關閉機制、第三方 AI 認證'),
         ('第二波', '代理型 AI 核保理賠', '6–8 週單一險種代理型工作流藍圖', 'AI 代理編排層、核心系統 API 串接、人機協作作業導入'),
         ('第二波', '核心現代化與雲端', '雲端委外合規評估，或舊系統 AI 重寫 PoC', '核心中台解耦、同業聯合查核、遷移與測試人力委外'),
         ('第二波', '深偽反詐欺', '車損照片／醫療單據偽造偵測 PoC', '即時反詐欺評分、關聯圖譜分析、調查派案整合')]
tb(s, 3.6, 1.35, 4.0, 0.4, '切入（Land）', size=12, bold=True, color=OR)
tb(s, 8.3, 1.35, 4.6, 0.4, '延伸（Expand）', size=12, bold=True, color=OR)
for i, (wv, n, a, b) in enumerate(lands):
    y = 1.8 + i * 0.9
    first = wv == '第一波'
    tb(s, 0.45, y + 0.17, 0.95, 0.45, wv, size=10.5, bold=True, color=WHITE if first else OR, fill=OR if first else TINT, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    tb(s, 1.5, y, 2.0, 0.8, n, size=12, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    tb(s, 3.6, y, 4.0, 0.8, a, size=11.5, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
    arrow(s, 7.7, y + 0.24, 0.45, 0.32)
    tb(s, 8.3, y, 4.6, 0.8, b, size=11.5, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
takeaway(s, 0.45, 6.4 - 0.15, 12.45, 0.45, '第一波兩項首案成立後，用同一批客戶關係與資料基礎提案第二波，降低每案的取得成本。', size=11.5)
source(s, '來源：研究報告各趨勢的「第一步切入」建議。')

s = content('國際顧問都在推代理型 AI 藍圖與平台，我們以「落地＋確信＋駐點人力」做出區隔')
rows = [['機構', '主打', '可能的交付侷限'],
        ['Deloitte', 'Agentic AI Blueprint；預置代理庫', '套到台灣客製化舊核心與繁中條款時，資料對接與在地合規需調適'],
        ['EY', 'EY Nexus 平台；AI 治理與確信', '偏新業務快速上線，大型壽險舊核心遷移與 IFRS 17 整合較難只靠平台'],
        ['McKinsey', '核保作業系統 UWOS；選擇性重寫', '缺少在地法規確信團隊與長期駐點人力'],
        ['Bain', '12 週 Future-back；三層代理架構', '在台交付規模較小，大型核心導入人力有限']]
table(s, 0.45, 1.4, 7.3, [1.3, 2.6, 3.4], rows, size=10.5, head_size=11.5, row_h=0.9)
tb(s, 8.0, 1.4, 4.9, 0.45, '我們的四個差異化', size=13, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.07)
diffs = [('Assurance-by-Design', '設計 AI 工作流時同步內建稽核軌跡與護欄'),
         ('顧問藍圖＋人力委外', '小案切入，接上 1–3 年經常性人力'),
         ('以 IFRS 17 資料庫為樞紐', '把合規資料延伸到 ALM、定價與通路 KPI'),
         ('雲端與第三方 AI 聯合查核', '多家保險公司共同委託，分攤成本 ⚠ 需確認獨立性')]
for i, (h, d) in enumerate(diffs):
    y = 1.95 + i * 0.85
    circle(s, 8.0, y + 0.15, 0.48, str(i + 1), size=13)
    tb(s, 8.6, y, 4.3, 0.78, [[(h, {'bold': True, 'size': 12})], d], size=11, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
takeaway(s, 0.45, 5.95, 12.45, 0.6, '對手賣的是藍圖與平台；我們賣「做得出來、查得過、有人長期顧」。此對比為公開資料推論，對外勿點名。', size=11.5)
source(s, '來源：各機構 2025–2026 公開發布的服務與研究（Deloitte 2026-05、EY 2026-09、McKinsey 2026-04／06、Bain 2026）。')

# ───────── 第四章 行動 ─────────
section('04', '行動：90 天計畫、\n風險與今天的決定', '把建議轉成負責人、時程與檢查點')

s = content('90 天行動計畫：第 4 週服務包定稿、第 8 週完成內部審查、第 12 週提出首案建議書')
PH = ['第 1–4 週｜打包', '第 5–8 週｜驗證', '第 9–12 週｜開賣']
GX, GW = 3.05, 9.85
for k, h in enumerate(PH):
    tb(s, GX + k * GW / 3, 1.4, GW / 3 - 0.06, 0.45, h, size=12, bold=True, color=WHITE, fill=INK, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
WS = [('① IFRS 17 服務包', ['服務一頁紙、首案範圍與報價區間', '與所內 IFRS 17 團隊確認分工、選定目標客戶', '客戶對談、首案建議書'], OR),
      ('② AI 治理服務包', ['健檢方法與 AI 清冊範本', 'R&Q 獨立性確認、內部試做', '客戶對談、首案建議書'], OR),
      ('第二波準備', [None, '代理型 AI 藍圖方法、核心／雲端評估框架、反詐欺技術夥伴初篩', None], OR2),
      ('素材與研究', ['研究數字回原文核對', '研討會 AI 短片定稿', '研討會與客戶拜訪使用'], GREY)]
for i, (n, cells, c) in enumerate(WS):
    y = 1.98 + i * 0.95
    tb(s, 0.45, y, 2.5, 0.85, n, size=12.5, bold=True, color=c if c != GREY else INK, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
    if cells[0] is None:
        tb(s, GX + GW / 3, y + 0.05, 2 * GW / 3 - 0.06, 0.75, cells[1], size=11, fill=TINT, line=OR2, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
        continue
    for k, t in enumerate(cells):
        tb(s, GX + k * GW / 3, y + 0.05, GW / 3 - 0.06, 0.75, t, size=11, fill=TINT if c != GREY else LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
MS = [(1, '服務包定稿'), (2, '內部審查'), (3, '首案建議書')]
hline(s, GX, 5.95, GX + GW, 5.95, color=GREY)
for k, t in MS:
    x = GX + k * GW / 3 - 0.2
    d = s.shapes.add_shape(MSO_SHAPE.DIAMOND, Inches(x), Inches(5.8), Inches(0.3), Inches(0.3))
    d.fill.solid(); d.fill.fore_color.rgb = OR; d.line.fill.background(); d.shadow.inherit = False
    tb(s, x - 1.6, 6.12, 1.85, 0.35, f'W{k * 4}　{t}', size=10.5, bold=True, color=OR, align=PP_ALIGN.RIGHT)
tb(s, 0.45, 5.78, 2.5, 0.35, '檢查點', size=11, bold=True, color=GREY)
source(s, '週數自會議決定日起算。負責人與人力配置待今天指派。')

s = content('四個主要風險都有對應做法，最需要先處理的是資料可信度與獨立性')
rows = [['風險', '可能的影響', '對應做法', '時點'],
        ['資料可信度', '研究由 Gemini Deep Research 彙整，部分數字為單一來源；直接引用可能被客戶質疑', '對外前逐一回原文核對；單一來源數字標註或不用', '第 1–4 週'],
        ['獨立性', 'AI 治理健檢、第三方 AI 認證與聯合查核涉及確信角色；IFRS 17 顧問可能碰到審計客戶', '推動前與 Risk & Quality 確認；審計客戶另案評估', '第 1 週起'],
        ['交付產能', '第一波與既有專案爭用同一批精算、財會與 AI 治理人力', '首案前由既有團隊兼任；首案成立後再補人力', '第 5–8 週'],
        ['競爭者先行', '國際顧問已推出代理型 AI 藍圖與平台，客戶可能已在洽談', '主打在地確信與駐點人力；優先鎖定既有客戶', '持續']]
gt = table(s, 0.45, 1.4, 12.45, [1.8, 4.75, 4.4, 1.5], rows, size=11.5, head_size=12, row_h=0.86)
for i in (1, 2):
    gt.cell(i, 0).fill.solid(); gt.cell(i, 0).fill.fore_color.rgb = TINT
takeaway(s, 0.45, 6.0, 12.45, 0.55, '前兩項風險不先處理，第一波服務就無法對外；建議列入今天的決策事項。', size=11.5)
source(s, '單一來源待驗證數字：Celent 22% 計畫部署代理型 AI、亞太深偽網路險理賠 +233%、嵌入式保險市場規模、AI 精算定價 +30% 效率。')

s = content('今天請決定四件事：以下為建議方案')
rows = [['#', '決策事項', '建議方案', '負責人', '期限'],
        ['1', '第一波推哪兩項服務', 'IFRS 17／TW-ICS 後財務轉型、AI 清冊與模型驗證健檢', '主管決定', '今天'],
        ['2', '每項服務包的負責人', '各指派 1 位負責人＋1 位支援；負責人第 1 週找 R&Q 確認獨立性', '主管指派', '今天'],
        ['3', '90 天計畫與檢查點', '核定第 4、8、12 週三次檢查點，每次 30 分鐘', '服務包負責人', '第 1 週'],
        ['4', '研討會 AI 短片的主題', '以第一波服務為主題；原版／新版與結尾 CTA 二選一', '主管', '第 2 週']]
gt = table(s, 0.45, 1.4, 12.45, [0.6, 3.0, 5.85, 1.7, 1.3], rows, size=13, head_size=12.5, row_h=0.86)
for i in range(1, 5):
    gt.cell(i, 0).text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    gt.cell(i, 2).fill.solid(); gt.cell(i, 2).fill.fore_color.rgb = TINT
tb(s, 0.45, 6.0, 12.45, 0.55, [[('會議後 48 小時內：', {'bold': True, 'color': OR}), ('寄出會議紀錄與決策結果；負責人確認第 1 週工作與 R&Q 會議時間。', {})]],
   size=12, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.14)

# ───────── 附錄 ─────────
section('A', '附錄', '九項趨勢的定義、國際案例與台灣痛點；主要來源')
TR = [
 ('1｜代理型 AI 核保與理賠', '多代理協作自主完成收件解析、外部數據查核、風險分流與理賠審核，人只處理複雜案件', 'Manulife MAUDE 核保系統、AmTrust 生成式 AI 理賠（Celent 2025／2026 獎）', '已有核保、理賠智能助理，但停在單點知識檢索，無法跨核心系統自動觸發照會或給付', '核保／理賠三車道分流藍圖、工具目錄與 API 架構、AI 代理權限與護欄', '6–8 週單一險種「價值量化與工作流重塑」工作坊'),
 ('2｜AI 加速核心現代化與雲端', 'AI 自動解析舊程式、資料映射與測試生成；客製化移出核心，改放雲端中台', 'John Hancock、MSIG USA 雲原生核心與數據平台（Celent 獎）', '核心替換動輒 5–8 年且超支；舊語言維護人力不足；舊核心無法支援 API 與 AI 串接', 'AI 加速核心翻新評估、中台解耦架構、遷移測試人力、雲端業者聯合查核', '雲端委外合規落差評估，或周邊系統 AI 重寫 PoC'),
 ('3｜營運化 AI 治理', '把原則轉成可查核的 AI 清冊、偏見測試與執行期護欄（權限、防注入、自動回滾、緊急關閉）', 'NAIC 12 州試點評估工具；MAS AI 風險指引；香港保監局 AI Cohort', '多由法遵寫紙本政策，缺少跨精算、核保、IT 的模型驗證，也看不到部門自用的外部 AI 工具', 'AI 清冊盤點與風險分級、模型驗證機制、第三方 AI 盡職調查範本、LLM 紅隊演練', '4–6 週 AI 清冊盤點＋金管會指引合規健檢'),
 ('4｜IFRS 17 後財務轉型', '開帳後轉向結帳與內控自動化、以 KRD 為核心的動態 ALM、新契約 CSM 納入通路 KPI', 'Manulife、Zurich 將 CSM 納入核心 KPI；韓國接軌 K-ICS 後調整商品結構', '結帳仰賴人工與外包；覆蓋法退場使損益劇烈波動；通路仍以初年度保費考核', '結帳自動化、ALM 分析平台、CSM 通路 KPI 設計、精算財會人力委外', '首年結帳瓶頸診斷，或 CSM 管理儀表板'),
 ('5｜深偽理賠反詐欺', '影像鑑識、文件中繼資料檢查、敘事矛盾比對，嵌入理賠調查工作流', 'Swiss Re SONAR 2025；FRISS、Verisk 多模態偽造偵測', '行動理賠普及，修圖車損、偽造醫療收據門檻變低，規則引擎辨識不了', '多模態偵測模型、跨保單與院所關聯圖譜、調查派案流程重塑', '歷史車損照片／醫療單據 AI 偽造偵測 PoC'),
 ('6｜AI 精算定價與數據平台', 'Lakehouse 數據平台＋AI 費率開發，自動化費率模擬、規則生成與版本控管', 'Roamly Actuarial Rater、AIA、Income Insurance（Celent 獎）', '費率配置進核心仰賴人工轉檔與核對，上市慢且有配置錯誤裁罰風險', '三層數據架構（中繼資料、血緣、資料契約）、定價到上線自動化、數據工程人力', '新商品費率上線流程診斷與自動比對'),
 ('7｜嵌入式保險 API 平台', '標準 Open API 把場景化保單嵌入銀行、車廠、電信、旅遊平台，即時定價出單理賠', 'bolttech、Chubb Studio、Qover', '每接一家異業就要數月客製開發；缺少商品工廠與 API 閘道', 'API 中台架構、異業分潤與資料共享合規架構、內嵌式理賠流程', '高頻場景 API 商品工廠與快速串接藍圖'),
 ('8｜參數型氣候巨災保險', '以風速、震度、淹水面積等參數觸發理賠，30 天內撥款', 'Swiss Re STORM／QUAKE／HAIL；紐約市參數型洪水險', '颱風、地震高風險；科技廠與綠能資產需求大；缺少地理空間建模與基差風險量化', '氣候實體風險數據整合、參數型商品設計、自動觸發數據管道', '企業巨災自負額填補可行性評估'),
 ('9｜AI 表現保證保險', '模型表現不如預期或 AI 侵權時的保證與責任保險，以技術盡職調查為前提', 'Armilla AI（Lloyd\'s）、Munich Re aiSure', '企業採購 AI 擔心求償無門；產險承保 AI 資料中心面臨價值累積風險', '第三方 AI 模型技術盡職調查、AI 責任險批單與資料中心核保顧問', '第三方 AI 供應商風險認證與合約保證條款設計'),
]
for k in range(3):
    s = content(f'A{k + 1}｜九項趨勢明細（{k + 1}／3）')
    labels = ['是什麼', '國際案例', '台灣痛點', '服務包裝', '第一步切入']
    for r, lab in enumerate(labels):
        y = 2.0 + r * 0.92
        tb(s, 0.45, y, 1.35, 0.86, lab, size=11, bold=True, color=OR, anchor=MSO_ANCHOR.MIDDLE, fill=TINT, margin=0.1)
    for c, t in enumerate(TR[k * 3:k * 3 + 3]):
        x = 1.9 + c * 3.68
        tb(s, x, 1.4, 3.55, 0.55, t[0], size=12.5, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
        for r in range(5):
            tb(s, x, 2.0 + r * 0.92, 3.55, 0.86, t[r + 1], size=10.5, fill=LG if r % 2 == 0 else WHITE, line=LINE, anchor=MSO_ANCHOR.MIDDLE, margin=0.08)

s = content('A4｜主要來源')
src = [['類別', '來源（發布日期）'],
       ['監理機關', ['金管會《金融業運用 AI 指引》（2024-08）', '金管會《保險業作業委託他人處理應注意事項》修正（2024-08）；內部作業制度及程序辦法草案（2025-10 預告）', '產壽險公會《保險業使用雲端服務自律規範》', 'NAIC AI Model Bulletin 實施追蹤（2026-04）', 'MAS AI 風險管理指引（2025-11）；香港保監局 AI Cohort（2025-08）']],
       ['產業研究', ['Gallagher Re《Global InsurTech Report Q1 2026》（2026-05）、Q2 2026（2026-08）', 'Swiss Re SONAR 2025（2025-06）；參數型洪水保險（2026-03）', 'Celent Model Insurer 2025／2026；Shedding Light on Agentic AI（2025-05）', 'Armilla AI；Munich Re《Insuring Generative AI》']],
       ['顧問與事務所', ['McKinsey：Agentic AI 核心現代化（2026-04）、UWOS（2026-06）', 'Bain：Reimagining Insurance for the AI Era（2026）、How to Architect for Agentic AI', 'Deloitte：Agentic AI Blueprint for Insurance；亞太壽險論文（2026-05）', 'EY：EY Nexus for Insurance；AI Risk and Governance Survey（2026-09）', 'PwC Taiwan：IFRS 17 關鍵議題（2025-11）、首份 IFRS 17 財報導讀（2026）']],
       ['媒體', ['Insurance Journal（2025-11）；工商時報（2024-06、2024-08、2025-09、2026-01、2026-02、2026-05）；Insurance Innovation Reporter（2026-06）']]]
table(s, 0.45, 1.4, 12.45, [1.8, 10.65], src, size=11, head_size=11.5, row_h=0.6)
source(s, '完整連結見研究報告原文。')

p.save(OUT)
print(OUT, len(p.slides), 'slides')
