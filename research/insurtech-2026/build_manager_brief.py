# 保險科技服務優先順序：主管三頁精簡版（附可點擊來源連結，以所內簡報範本建立）
# 用法：python3 build_manager_brief.py <範本.pptx> <輸出.pptx>
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

TPL, OUT = sys.argv[1], sys.argv[2]
OR = RGBColor(0xFD, 0x51, 0x08); OR2 = RGBColor(0xFE, 0x7C, 0x39); TINT = RGBColor(0xFF, 0xE8, 0xD4)
INK = RGBColor(0x1E, 0x1E, 0x1E); GREY = RGBColor(0x6B, 0x6B, 0x6B); LG = RGBColor(0xF3, 0xF3, 0xF3); LINE = RGBColor(0xD9, 0xD9, 0xD9)
WHITE = RGBColor(0xFF, 0xFF, 0xFF); LINK = RGBColor(0x1F, 0x5F, 0xBF)
FONT = 'Microsoft JhengHei'
FOOT = '內部討論稿｜僅供 FSC1 保險組內部使用｜資料來源為公開研究，點選來源可開啟原文'

# 來源連結：清單（insurtech_research_links_directory.md）有具體頁面者用清單連結；
# 清單沒有、或只給入口首頁者，用網路搜尋找到的具體頁面
URL = {
    'gallagher': 'https://www.ajg.com/gallagherre/news-and-insights/global-insurtech-report-q1-2026/',
    'gallagher_news': 'https://www.insurancebusinessmag.com/reinsurance/news/breaking-news/ai-startups-capture-record-95-2-of-insurtech-funding--gallagher-re-574376.aspx',
    'pwc57': 'https://amp.claimsjournal.com/news/national/2025/11/17/334094.htm',
    'bain': 'https://www.bain.com/ko/results/reimagining-insurance-for-the-ai-era',
    'mckinsey': 'https://www.mckinsey.com/industries/financial-services/our-insights/can-agentic-ai-finally-modernize-core-technologies-in-insurance',
    'deloitte': 'https://www.deloitte.com/ap/en/about/press-room/deloitte-launches-new-paper-on-scaling-agentic-ai-in-life-insurance.html',
    'fsc_ai': 'https://www.fsc.gov.tw/websitedowndoc?file=chfsc/202408231741530.pdf',
    'fsc_ai_news': 'https://www.leeandli.com/TW/Newsletters/7287.htm',
    'fsc_outsource': 'https://www.fsc.gov.tw/ch/home.jsp?id=96&parentpath=0,2&mcustomize=news_view.jsp&dataserno=202408200001',
    'outsource_rule': 'https://www.leeandli.com/TW/Newsletters/7525.htm',
    'cloud_selfreg': 'https://law.lia-roc.org.tw/Law/GetFile/0000419987',
    'pwc_ifrs17': 'https://www.pwc.tw/zh/topics/trends/ifrs-17-key-issues.html',
    'naic': 'https://content.naic.org/cmte_h_bdaiwg.htm',
    'naic_pilot': 'https://www.fenwick.com/insights/publications/naic-expands-ai-systems-evaluation-tool-pilot-program-to-12-states-key-updates-for-insurers-and-ai-vendors-supporting-insurers',
    'mas': 'https://www.mas.gov.sg/regulation/guidelines/guidelines-on-artificial-intelligence-risk-management-for-financial-institutions',
}

p = Presentation(TPL)
lst = p.slides._sldIdLst
for sid in list(lst):
    p.part.drop_rel(sid.rId); lst.remove(sid)
L = {l.name: l for l in p.slide_masters[2].slide_layouts}
# 移除範本遺留的投影片章節（PowerPoint 左側的章節名稱）
for ext in list(p.part._element.iter('{http://schemas.openxmlformats.org/presentationml/2006/main}ext')):
    if ext.find('{http://schemas.microsoft.com/office/powerpoint/2010/main}sectionLst') is not None:
        ext.getparent().remove(ext)

def set_ph(s, idx, text, size=None):
    for ph in s.placeholders:
        if ph.placeholder_format.idx == idx:
            ph.text = text
            for para in ph.text_frame.paragraphs:
                for r in para.runs:
                    r.font.name = FONT
                    if size: r.font.size = Pt(size)
            return ph

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
            f = r.font; f.name = FONT; f.size = Pt(st.get('size', size)); f.bold = st.get('bold', bold)
            f.color.rgb = st.get('color', color)
            if 'link' in st:
                r.hyperlink.address = st['link']; f.underline = True
                if 'color' not in st: f.color.rgb = LINK
        para.space_after = Pt(2)
    return sh

def circle(s, x, y, d, text, fill=OR, size=12):
    c = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    c.fill.solid(); c.fill.fore_color.rgb = fill; c.line.fill.background(); c.shadow.inherit = False
    tf = c.text_frame; tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    para = tf.paragraphs[0]; para.alignment = PP_ALIGN.CENTER
    r = para.add_run(); r.text = text; r.font.size = Pt(size); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = FONT

def hline(s, x1, y1, x2, y2, color=LINE, width=1.0):
    ln = s.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    ln.line.color.rgb = color; ln.line.width = Pt(width)

def content(title):
    s = p.slides.add_slide(L['3__Pattern header'])
    set_ph(s, 0, title, 20); set_ph(s, 3, FOOT)
    return s

def sources(s, items):
    segs = [('來源：', {'bold': True, 'color': GREY})]
    for i, (label, key) in enumerate(items):
        if i: segs.append(('　', {}))
        segs.append((label, {'link': URL[key]}))
    tb(s, 0.45, 6.62, 12.45, 0.42, [segs], size=9, color=GREY, margin=0.03)

# ───────── 第 1 頁 重點摘要 ─────────
s = content('重點摘要：保險 AI 已進入規模化，建議第一波先推 IFRS 17 與 AI 治理兩項服務包')
STATS = [('95.2%', '2026 Q1 全球保險科技募資流向 AI 企業的比例（歷史新高）', 'Gallagher Re, 2026-05', 'gallagher'),
         ('57%', '保險高階主管把生成式與代理型 AI 列為 2026 首要科技投資', 'PwC 調查, 2025-11', 'pwc57'),
         ('95%', '投資 AI 但尚未看到可衡量回報的企業（跨產業）', 'Bain 引述 MIT', 'bain')]
for i, (n, d, src, key) in enumerate(STATS):
    y = 1.4 + i * 1.72
    hot = i == 2
    tb(s, 0.45, y, 3.9, 1.58, '', fill=OR if hot else LG)
    tb(s, 0.6, y + 0.06, 3.6, 0.65, n, size=30, bold=True, color=WHITE if hot else OR)
    tb(s, 0.6, y + 0.72, 3.6, 0.5, d, size=11, color=WHITE if hot else INK, margin=0.03)
    tb(s, 0.6, y + 1.24, 3.6, 0.3, [[(src, {'link': URL[key], 'color': WHITE if hot else LINK})]], size=9, margin=0.03)
tb(s, 4.6, 1.4, 8.3, 1.1, [[('情境　', {'bold': True, 'color': OR}), ('資金與預算全面轉向 AI，保險業已從試點走向規模化。', {})],
                           [('衝突　', {'bold': True, 'color': OR}), ('多數投資還看不到回報；台灣同時面對 IFRS 17／TW-ICS 首年出表與 AI 治理要求。', {})]],
   size=12, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.14)
tb(s, 4.6, 2.62, 8.3, 0.9, [[('建議　', {'bold': True, 'size': 12}), ('第一波先推 ① IFRS 17／TW-ICS 後財務轉型　② AI 清冊、模型驗證與護欄；代理型 AI 核保理賠、核心現代化、深偽反詐欺列為第二波。', {})]],
   size=13.5, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.16)
WHY = [('需求有時限', 'IFRS 17／TW-ICS 2026-01-01 接軌；委外新辦法 2025-12-31 施行；AI 指引進入落地'),
       ('離既有服務最近', '延伸既有駐點團隊與 AI 治理服務，不需新技術夥伴；AI 治理健檢 4–6 週可交付'),
       ('能帶出第二波', '財務資料與 AI 治理是代理型 AI、核心現代化的前置條件，可延伸為 1–3 年人力委外')]
for i, (h, d) in enumerate(WHY):
    x = 4.6 + i * 2.8
    tb(s, x, 3.65, 2.68, 0.48, h, size=12.5, bold=True, color=OR, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    tb(s, x, 4.18, 2.68, 1.4, d, size=11, fill=LG, margin=0.12)
tb(s, 4.6, 5.72, 8.3, 0.75, [[('今天請決定：', {'bold': True, 'color': OR}), ('同意第一波兩項服務、各指派 1 位負責人、核定 90 天計畫（第 4／8／12 週檢查）。', {})]],
   size=12, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.14)
sources(s, [('Gallagher Re Global InsurTech Report Q1 2026', 'gallagher'), ('Insurance Business 報導（95.2%）', 'gallagher_news'),
            ('PwC 保險高階主管調查（Claims Journal）', 'pwc57'), ('Bain：Reimagining Insurance for the AI Era', 'bain')])

# ───────── 第 2 頁 為什麼是這兩項 ─────────
s = content('為什麼是這兩項：台灣監理時程已經確定，而且兩項服務離我們的既有能力最近')
EV = [('2024-06', '金管會《金融業運用 AI 指引》', 'fsc_ai', True), ('2024-08', '委外注意事項修正：放寬雲端聯合查核', 'fsc_outsource', True),
      ('2025-12-31', '委外新辦法施行，提升為法規命令', 'outsource_rule', True), ('2026-01-01', 'IFRS 17／TW-ICS 正式接軌', 'pwc_ifrs17', True),
      ('2026-03～09', 'NAIC 12 州 AI 評估工具試點（美國）', 'naic_pilot', False), ('2026-10', 'MAS AI 風險管理指引定稿（新加坡）', 'mas', False)]
hline(s, 0.6, 2.1, 12.75, 2.1, color=OR, width=2.5)
for i, (d, t, key, tw) in enumerate(EV):
    x = 0.5 + i * 2.07
    circle(s, x + 0.84, 1.96, 0.28, '', fill=OR if tw else GREY)
    tb(s, x, 1.38, 1.96, 0.5, d, size=12.5, bold=True, color=OR if tw else GREY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM)
    tb(s, x, 2.38, 1.96, 0.95, [t, [('原文', {'link': URL[key], 'size': 9})]], size=10.5, align=PP_ALIGN.CENTER,
       fill=LG if tw else None, line=None if tw else LINE, margin=0.07)
tb(s, 0.45, 3.42, 12.4, 0.3, '時間軸：橘色為台灣監理事件，灰色為國際標竿。　評等：● 有利　◐ 中等　○ 不利，為團隊依研究報告判斷，請會議校準。', size=9, color=GREY)
F, M, N = '●', '◐', '○'
ROWS = [('IFRS 17／TW-ICS 後財務轉型', F, F, F, '首年結帳瓶頸診斷＋CSM 儀表板', '第一波'),
        ('AI 清冊、模型驗證與護欄', F, F, F, '4–6 週 AI 清冊盤點與合規健檢', '第一波'),
        ('代理型 AI 核保與理賠', F, M, M, '6–8 週單一險種工作流藍圖', '第二波'),
        ('AI 加速核心現代化與雲端', M, M, N, '雲端委外合規評估或 AI 重寫 PoC', '第二波'),
        ('深偽理賠反詐欺', M, N, M, '車損照片／醫療單據偽造偵測 PoC', '第二波')]
HEAD = ['服務', '需求迫切度', '離既有服務', '首案週期', '切入案', '建議']
CW = [3.4, 1.35, 1.35, 1.35, 3.85, 1.1]
y0 = 3.78; rh = 0.47
x = 0.45
for j, h in enumerate(HEAD):
    tb(s, x, y0, CW[j] - 0.04, 0.4, h, size=11, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.LEFT if j in (0, 4) else PP_ALIGN.CENTER, margin=0.08)
    x += CW[j]
for i, row in enumerate(ROWS):
    y = y0 + 0.44 + i * rh; x = 0.45
    first = row[-1] == '第一波'
    for j, v in enumerate(row):
        glyph = j in (1, 2, 3)
        fill = (OR if first else TINT) if j == 5 else (LG if i % 2 == 0 else WHITE)
        tb(s, x, y, CW[j] - 0.04, rh - 0.04, v, size=16 if glyph else 11, bold=j in (0, 5),
           color=WHITE if (j == 5 and first) else (OR if glyph or j == 5 else INK), fill=fill, anchor=MSO_ANCHOR.MIDDLE,
           align=PP_ALIGN.LEFT if j in (0, 4) else PP_ALIGN.CENTER, margin=0.08)
        x += CW[j]
sources(s, [('金管會 AI 指引', 'fsc_ai'), ('理律：AI 指引解析', 'fsc_ai_news'), ('金管會委外修正', 'fsc_outsource'), ('理律：委外新辦法', 'outsource_rule'),
            ('公會雲端自律規範', 'cloud_selfreg'), ('資誠 IFRS 17 關鍵議題', 'pwc_ifrs17'), ('NAIC 工作小組', 'naic'), ('MAS AI 指引', 'mas')])

# ───────── 第 3 頁 行動與決策 ─────────
s = content('90 天內把兩項服務包做成可銷售的提案，今天請決定四件事')
PH = ['第 1–4 週｜打包', '第 5–8 週｜驗證', '第 9–12 週｜開賣']
PLAN = [('① IFRS 17', ['服務一頁紙、首案範圍與報價區間', '與所內 IFRS 17 團隊確認分工、選目標客戶', '客戶對談、首案建議書']),
        ('② AI 治理', ['健檢方法與 AI 清冊範本', 'R&Q 獨立性確認、內部試做', '客戶對談、首案建議書'])]
GX, GW = 1.75, 11.15
for k, h in enumerate(PH):
    tb(s, GX + k * GW / 3, 1.4, GW / 3 - 0.06, 0.42, h, size=12, bold=True, color=WHITE, fill=INK, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
for i, (n, cells) in enumerate(PLAN):
    y = 1.9 + i * 0.68
    tb(s, 0.45, y, 1.24, 0.6, n, size=12, bold=True, color=OR, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.08)
    for k, t in enumerate(cells):
        tb(s, GX + k * GW / 3, y, GW / 3 - 0.06, 0.6, t, size=11, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
DEC = [('1', '第一波推哪兩項服務', 'IFRS 17／TW-ICS 後財務轉型、AI 清冊與模型驗證健檢', '今天'),
       ('2', '每項服務包的負責人', '各 1 位負責人＋1 位支援；第 1 週找 R&Q 確認獨立性', '今天'),
       ('3', '90 天計畫與檢查點', '第 4、8、12 週各一次，每次 30 分鐘', '第 1 週'),
       ('4', '研討會 AI 短片主題', '以第一波服務為主題', '第 2 週')]
tb(s, 0.45, 3.4, 8.2, 0.42, '今天請決定', size=13, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.06)
for i, (n, a, b, d) in enumerate(DEC):
    y = 3.9 + i * 0.66
    circle(s, 0.5, y + 0.1, 0.42, n, size=12)
    tb(s, 1.05, y, 2.35, 0.58, a, size=11.5, bold=True, anchor=MSO_ANCHOR.MIDDLE, fill=LG, margin=0.08)
    tb(s, 3.45, y, 4.25, 0.58, b, size=11, anchor=MSO_ANCHOR.MIDDLE, fill=TINT, margin=0.08)
    tb(s, 7.75, y, 0.9, 0.58, d, size=11, bold=True, color=OR, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, fill=LG, margin=0.04)
tb(s, 8.9, 3.4, 4.0, 0.42, '使用前注意', size=13, bold=True, color=WHITE, fill=INK, anchor=MSO_ANCHOR.MIDDLE, margin=0.06)
CAV = [[('數字核對　', {'bold': True}), ('對外引用前回原文確認；95% 為跨產業數字', {})],
       [('獨立性　', {'bold': True}), ('AI 治理健檢與 IFRS 17 顧問碰到審計客戶時，先與 R&Q 確認', {})],
       [('對外不點名　', {'bold': True}), ('研究中的同業與競爭者僅供內部討論', {})],
       [('完整版　', {'bold': True}), ('24 頁會議簡報含問題樹、矩陣與風險', {})]]
for i, t in enumerate(CAV):
    tb(s, 8.9, 3.9 + i * 0.66, 4.0, 0.58, [t], size=10.5, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
sources(s, [('McKinsey：代理型 AI 核心現代化', 'mckinsey'), ('Deloitte：亞太壽險規模化代理型 AI', 'deloitte'),
            ('資誠 IFRS 17 關鍵議題', 'pwc_ifrs17'), ('MAS AI 風險管理指引', 'mas')])

p.save(OUT)
print(OUT, len(p.slides), 'slides')
