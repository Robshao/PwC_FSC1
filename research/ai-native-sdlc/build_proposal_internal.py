# AI 原生軟體開發提案：內部小組會議版（11 頁）
# 對齊客戶提案的故事線（四個面向）、Sam 審閱後的修改，以及客戶版定稿前要拍板的事項；客戶版見 build_proposal_deck.py
# 用法：python3 build_proposal_internal.py <範本.pptx> <輸出.pptx>
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
FOOT = '內部小組會議｜AI 原生軟體開發提案籌備｜AI 協助產出，需人工確認後再對外'
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

# ───────── 1 封面 ─────────
s = p.slides.add_slide(L['標題投影片'])
set_ph(s, 0, 'AI 原生軟體開發提案：\n內部小組會議', 28)
set_ph(s, 1, 'FSC1 內部｜客戶版定稿前的對齊｜2026 年 10 月', 12)

# ───────── 2 今天的目的 ─────────
s = content('今天的目的：對齊提案故事線，拍板七件事，客戶版才能定稿')
tb(s, 0.45, 1.45, 6.1, 0.5, '今天要達成', size=14, bold=True, color=WHITE, fill=INK, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
GOAL = ['確認客戶版以四個面向為主軸的故事線（第 3 頁）', '確認 Sam 審閱後的 11 項修改（第 4 頁）', '審閱兩個新增頁：法規檢查清單、試點成功標準（第 5–6 頁）', '拍板價格、人力、目標客戶等七件事（第 9 頁）']
tb(s, 0.45, 2.0, 6.1, 3.3, [[(f'{i + 1}　', {'bold': True, 'color': OR}), (t, {})] for i, t in enumerate(GOAL)], size=13, fill=LG, margin=0.2)
tb(s, 6.8, 1.45, 6.1, 0.5, '提案的三個關鍵選擇', size=14, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
KEY = [('對外只推兩步', '準備度診斷 → 陪你做一個試點；工具包與駐點等客戶問了再談。'),
       ('差異化要看得到', '用「法規檢查清單」和 IT 團隊的實作成果或現場示範證明，不只用說的。'),
       ('價值先講好標準', '不先寫效益數字，試點前雙方訂好三個指標和合格線。')]
tb(s, 6.8, 2.0, 6.1, 3.3, sum([[[(h, {'bold': True, 'color': OR, 'size': 13})], t, [('', {'size': 5})]] for h, t in KEY], []), size=12, fill=TINT, margin=0.2)
callout(s, 5.55, '核心價值主張：', '讓保險公司用 AI 加快系統交付，同時守住法規、留下每一步紀錄，並把做法變成團隊可以重複使用的能力。', h=0.75, size=12)

# ───────── 3 故事線 ─────────
s = content('提案故事線：客戶版正文 14 頁，四個面向各回答一個問題，最後落在「如何開始」')
COLS = [['第 3 頁　保險公司的難題', '第 4 頁　四種角色在意什麼'],
        ['第 5 頁　六個卡點與做法', '第 6 頁　從直線變循環', '第 7 頁　四個面向如何組成方案', '第 8 頁　為什麼選我們', '第 9 頁　法規檢查清單'],
        ['第 10 頁　問題 → 價值 → 指標', '第 11 頁　試點成功標準'],
        ['第 12 頁　從一個團隊到全公司']]
for i, ((zh, en), items) in enumerate(zip(ASPECT, COLS)):
    x = 0.45 + i * 3.14
    tb(s, x, 1.45, 3.0, 0.8, [[(f'{i + 1}　{zh}', {'bold': True, 'size': 14})], [(en, {'size': 10})]], color=WHITE, fill=OR if i < 3 else INK, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
    tb(s, x, 2.3, 3.0, 2.85, [[('・' + t, {})] for t in items], size=11.5, fill=LG, margin=0.14)
    if i < 3: arrow(s, x + 3.0, 1.7, 0.14, 0.28)
tb(s, 0.45, 5.3, 12.45, 0.6, [[('如何開始　', {'bold': True, 'color': WHITE, 'size': 13}), ('第 13 頁　合作方式：診斷與陪你做試點　｜　第 14 頁　路線圖與下一步', {'color': WHITE})]], size=12, fill=OR2, anchor=MSO_ANCHOR.MIDDLE, margin=0.16)
note(s, '附錄：A 示範情境（理賠進度自助查詢）、B 資料與假設說明及來源。英文標籤依 Sam 建議改為較自然的說法，中文名稱不變。', y=6.05)

# ───────── 4 Sam 審閱 ─────────
s = content('Sam 審閱後的 11 項修改：8 項已改進客戶版，3 項要等 IT、法遵或 Sam')
ROWS = [['第 8 頁「為什麼選我們」是空的', '保留待填框；來不及就改成現場示範', '待 IT'],
        ['「懂台灣法規」同業也會說', '起草法規檢查清單當證據（第 5 頁）', '待法遵'],
        ['每頁標籤像在說自己不確定', '客戶版拿掉，改在附錄說明資料與假設', '已改'],
        ['圖中有「Claude」，像在賣工具', '改用自己畫的循環圖，寫明不綁特定工具', '已改'],
        ['理賠情境不是我們的案例', '標成「示範情境」，移到附錄', '已改'],
        ['26 頁太長', '客戶版正文 14 頁＋附錄 2 頁', '已改'],
        ['擴充章在講我們自己', '改成客戶角度「從一個團隊到全公司」', '已改'],
        ['四個方案太多；駐點有獨立性疑慮', '對外只推診斷＋陪你做試點', '已改'],
        ['價值沒有數字', '新增「試點成功標準」（第 6 頁）', '已改'],
        ['95% 是跨產業數字', '痛點改以保險業新制與法規為主', '已改'],
        ['「Face who」英文不自然', '改為 Who we serve 等，中文名稱不變', '已改']]
grid(s, 0.45, 1.45, [0.55, 4.3, 6.1, 1.5], ['#', 'Sam 的疑問', '改法', '狀態'], [[str(i + 1)] + r for i, r in enumerate(ROWS)], 0.42, size=10.5, hot=(0, 1))
note(s, '淺橘底兩項是客戶版能不能拿出去的關鍵。', y=6.62)

# ───────── 5 法規檢查清單 ─────────
checklist_page('新增：保險業 AI 開發法規檢查清單（草稿），差異化的證據，對外前需法遵確認',
               '草稿，請法遵與 R&Q 審閱；條文以主管機關公告為準，非法律意見。金管會 AI 指引六大原則名稱依公開資料整理。')

# ───────── 6 試點成功標準 ─────────
s = criteria_page('新增：試點成功標準，用「先講好指標和合格線」回答「能省多少」')
note(s, '待 Sam 確認：是否願意對客戶承諾「未達標不擴大」。指標可依客戶痛點替換，以三個為限。', y=5.85)

# ───────── 7 方案全貌 ─────────
s = content('方案全貌：四種方案都保留在內部，對外只推「診斷 → 陪你做試點」')
ROWS = [['準備度診斷', '瓶頸地圖、法規檢查、試點建議、導入藍圖', '固定價', '對外主推', '低門檻，客戶容易說好'],
        ['陪你做｜試點陪跑＋教學', '完成的試點、成效報告、範本與審查規則、教材', '每個試點固定價，可加里程碑', '對外主推', '看得到成效，最能證明價值'],
        ['你自己做｜工具包', '範本、規範技能、閘門設定、交付手冊', '授權或年費', '客戶問了再談', '要先從試點累積成熟的資產'],
        ['幫你做｜駐點開發', '程式、文件與完整紀錄', '人月或按交付項目', '客戶問了再談', '偏人力派遣；審計客戶有獨立性限制']]
grid(s, 0.45, 1.45, [2.6, 3.7, 2.3, 1.55, 2.3], ['方案', '交付', '收費模式（建議）', '對外策略', '原因'], ROWS, 0.82, size=11, hot=(0, 1))
callout(s, 5.0, '計價原則（建議）：', '按「每次診斷」「每個試點」等任務或成果計價，不按工時。金額與期程待 Sam、Luke 決定（第 9 頁）。', h=0.65)
note(s, f'三種方案概念出自 {SMR}。', y=5.8)

# ───────── 8 我們的產品化邏輯 ─────────
s = content('我們自己的產品化邏輯（不放客戶版）：同業已在做，我們借鏡加速器模式，用資產累積拉開差距')
ROWS = [['Deloitte', 'Agentic AI Blueprint for Insurance', '最接近的競品：AI 導入方法＋預建代理'],
        ['Capgemini', 'Guidewire 生成式 AI／代理工具組', '核心系統遷移＋AI，和第 3 集影片主題重疊'],
        ['Deloitte、KPMG', 'IFRS 17 加速器（與 SAS 合作）', '預建資產＋明確期程，降低客戶疑慮']]
grid(s, 0.45, 1.45, [2.0, 4.2, 6.25], ['同業', '產品／方案', '對我們的意思'], ROWS, 0.62, size=11)
tb(s, 0.45, 3.85, 6.1, 1.85, [[('為什麼要產品化', {'bold': True, 'color': OR, 'size': 12.5})],
                              '・只接專案，營收靠多請人才能成長', '・客戶會拿同業的現成方案來比較', '・開發工作多半常發生、規則清楚，適合做成產品'], size=11.5, fill=LG, margin=0.16)
tb(s, 6.8, 3.85, 6.1, 1.85, [[('資產怎麼累積', {'bold': True, 'color': OR, 'size': 12.5})],
                             '每個案子結束，至少萃取一個可重用的範本或規範技能，回到共用資產庫；累積成熟後才推出「你自己做」的工具包。'], size=11.5, fill=TINT, margin=0.16)
note(s, '同業資料為業者公開資訊，未經獨立驗證；連結見附錄。對外不點名同業。', y=5.85)

# ───────── 9 要拍板的事 ─────────
s = content('客戶版定稿前要拍板的七件事')
ROWS = [['診斷與試點的價格區間', '客戶第一個問題就會問', 'Sam、Luke', '本次會議'],
        ['人力：診斷與試點由誰負責、排不排得出來', '承諾期程前要先確認', 'Sam', '本次會議'],
        ['IT 實作成果（客戶版第 8 頁）或現場示範', '差異化的證據', 'Kenny／IT 團隊', '本次會議'],
        ['法規檢查清單審閱', '對外前需法遵確認', '法遵、R&Q', '本次會議'],
        ['試點成功標準：是否承諾「未達標不擴大」', '影響客戶信任與我們的風險', 'Sam', '本次會議'],
        ['目標客戶名單與獨立性確認', '審計客戶需先經 R&Q', 'Sam、R&Q', '推案前'],
        ['對外使用的 AI 工具與所內政策', '客戶會問用什麼工具、資料去哪裡', 'IT、R&Q', '推案前']]
grid(s, 0.45, 1.45, [4.6, 3.6, 2.4, 1.85], ['事項', '為什麼重要', '負責', '期限'], ROWS, 0.62, size=11, hot=(0, 1, 2))
note(s, '負責人與期限為提案，以會議決定為準。', y=6.3)

# ───────── 10 時程 ─────────
s = content('接下來：內部確認完成後才定稿客戶版，再約第一場客戶工作坊')
FLOW = [('本次會議', '對齊故事線，拍板七件事', '全員'), ('補證據', 'IT 實績或示範；法遵審清單', 'Kenny、法遵'),
        ('定價與人力', '價格區間、負責人', 'Sam、Luke'), ('客戶版定稿', '更新第 8、9 頁並人工確認', 'BA'), ('第一場工作坊', '對齊目標、挑候選試點', 'Sam＋團隊')]
for i, (h, t, who) in enumerate(FLOW):
    x = 0.45 + i * 2.52
    tb(s, x, 1.6, 2.35, 0.65, h, size=13.5, bold=True, color=WHITE, fill=OR if i == 0 else (INK if i == 4 else OR2), anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    tb(s, x, 2.3, 2.35, 1.4, t, size=12, fill=LG, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, margin=0.12)
    tb(s, x, 3.75, 2.35, 0.5, who, size=11.5, bold=True, color=OR, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    if i < 4: arrow(s, x + 2.36, 1.78, 0.15, 0.3)
callout(s, 4.6, '原則：', '客戶資料不放進外部 AI 工具；AI 協助產出的內容對外前一定人工確認；審計客戶推案前先和 R&Q 確認獨立性。', h=0.7)
note(s, '各步驟日期待本次會議決定。', y=5.5)

# ───────── 11 附錄 ─────────
s = content('附錄｜同業產品與來源（業者公開資訊）')
ROWS = [('Deloitte', 'Agentic AI Blueprint for Insurance', 'https://www.deloitte.com/global/en/what-we-do/capabilities/agentic-ai-blueprint-for-insurance.html'),
        ('Deloitte', 'IFRS 17 Regulatory Assist（與 SAS 合作）', 'https://sas.com/nl_nl/news/press-releases/2022/may/deloitte-partners-with-sas-to-power-ifrs-17-accelerator-solution.html'),
        ('KPMG', 'IFRS 17 Express, powered by SAS', 'https://sas.com/en_ca/news/press-releases/2022/september/pvi-insurance-deploys-kpmg-ifrs-17-express.html'),
        ('Capgemini', 'Guidewire 生成式 AI／代理工具組', 'https://www.capgemini.com/wp-content/uploads/2026/10/GW-datasheet_Final-version.pdf'),
        ('EY', 'EY Nexus for Insurance', 'https://www.ey.com/en_us/insurance/nexus'),
        ('McKinsey', '代理型 AI 核心現代化', 'https://www.mckinsey.com/industries/financial-services/our-insights/can-agentic-ai-finally-modernize-core-technologies-in-insurance'),
        ('BCG', 'AI-First 產險公司藍圖', 'https://www.bcg.com/assets/2026/executive-perspectives-ai-first-companies-win-the-future-property-casualty-insurance.pdf')]
for i, (c, prod, url) in enumerate(ROWS):
    y = 1.45 + i * 0.6; bg = LG if i % 2 == 0 else WHITE
    tb(s, 0.45, y, 1.9, 0.54, c, size=11, bold=True, color=OR, fill=bg, anchor=MSO_ANCHOR.MIDDLE, margin=0.08)
    tb(s, 2.4, y, 4.6, 0.54, prod, size=11, fill=bg, anchor=MSO_ANCHOR.MIDDLE, margin=0.08)
    tb(s, 7.05, y, 5.85, 0.54, [[(url.split('/')[2] + ' 原文', {'link': url})]], size=11, fill=bg, anchor=MSO_ANCHOR.MIDDLE, margin=0.08)
note(s, '完整盤點見 research/productization/保險顧問產品化服務盤點.md。', y=5.8)

p.save(OUT)
print(OUT, len(p.slides), 'slides')
