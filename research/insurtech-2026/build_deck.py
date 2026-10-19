# 全球與亞太保險科技研究：內部討論簡報（以所內簡報範本建立）
# 用法：python3 build_deck.py <範本.pptx> <輸出.pptx>
import sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

TPL, OUT = sys.argv[1], sys.argv[2]
OR = RGBColor(0xFD, 0x51, 0x08); OR2 = RGBColor(0xFE, 0x7C, 0x39); TINT = RGBColor(0xFF, 0xE8, 0xD4)
INK = RGBColor(0x1E, 0x1E, 0x1E); GREY = RGBColor(0x6B, 0x6B, 0x6B); LG = RGBColor(0xF3, 0xF3, 0xF3); LINE = RGBColor(0xD9, 0xD9, 0xD9)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = 'Microsoft JhengHei'
FOOT = '內部討論稿｜僅供 FSC1 保險組內部使用｜資料來源為公開研究，數字待回原文核對'

p = Presentation(TPL)
lst = p.slides._sldIdLst
for sid in list(lst):
    p.part.drop_rel(sid.rId); lst.remove(sid)
L = {l.name: l for l in p.slide_masters[2].slide_layouts}

def set_ph(s, idx, text, size=None):
    for ph in s.placeholders:
        if ph.placeholder_format.idx == idx:
            ph.text = text
            for para in ph.text_frame.paragraphs:
                for r in para.runs:
                    r.font.name = FONT
                    if size: r.font.size = Pt(size)
            return ph

def footer(s):
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
            f = r.font; f.name = FONT; f.size = Pt(st.get('size', size)); f.bold = st.get('bold', bold)
            f.color.rgb = st.get('color', color)
        para.space_after = Pt(st.get('after', 3) if lines else 0)
    return sh

def circle(s, x, y, d, text, fill=OR, size=12):
    c = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    c.fill.solid(); c.fill.fore_color.rgb = fill; c.line.fill.background(); c.shadow.inherit = False
    tf = c.text_frame; tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    para = tf.paragraphs[0]; para.alignment = PP_ALIGN.CENTER
    r = para.add_run(); r.text = text; r.font.size = Pt(size); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = FONT

def table(s, x, y, w, col_w, rows, size=9.5, head_size=10, row_h=0.3, first_col_bold=True):
    nr, nc = len(rows), len(rows[0])
    gt = s.shapes.add_table(nr, nc, Inches(x), Inches(y), Inches(w), Inches(row_h * nr)).table
    for j, cw in enumerate(col_w): gt.columns[j].width = Inches(cw)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            c = gt.cell(i, j); c.text = ''
            c.margin_left = c.margin_right = Inches(0.06); c.margin_top = c.margin_bottom = Inches(0.035)
            c.fill.solid(); c.fill.fore_color.rgb = OR if i == 0 else (LG if i % 2 == 0 else WHITE)
            tf = c.text_frame; tf.word_wrap = True
            parts = val if isinstance(val, list) else [val]
            for k, part in enumerate(parts):
                para = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                r = para.add_run(); r.text = part; r.font.name = FONT
                r.font.size = Pt(head_size if i == 0 else size)
                r.font.bold = i == 0 or (j == 0 and first_col_bold)
                r.font.color.rgb = WHITE if i == 0 else (INK if not (j == 0 and first_col_bold) else OR)
    return gt

def content(title):
    s = p.slides.add_slide(L['3__Pattern header'])
    set_ph(s, 0, title, 20); footer(s)
    return s

def section(num, title, desc):
    s = p.slides.add_slide(L['章節標題'])
    set_ph(s, 0, title, 30); set_ph(s, 1, num); footer(s)
    tb(s, 0.45, 6.68, 8.5, 0.35, desc, size=10.5, color=GREY)
    return s

# 1 封面
s = p.slides.add_slide(L['標題投影片'])
set_ph(s, 0, '全球與亞太保險科技新應用\n暨顧問服務機會研究', 30)
set_ph(s, 1, 'FSC1 保險組｜內部討論稿｜2026 年 10 月', 12)

# 2 議程
s = content('議程')
items = [('01', '五大轉折', '過去 18 個月保險科技發生了什麼'), ('02', '九大趨勢與服務機會', '每項趨勢對應的台灣痛點與服務包裝'),
         ('03', '前五大優先服務', '評估方法、切入案與延伸路徑'), ('04', '競爭者與差異化', '國際顧問公司在做什麼、我們怎麼贏'),
         ('05', '下一步與待驗證事項', '請團隊決定的事、資料可信度')]
for i, (n, t, d) in enumerate(items):
    y = 1.55 + i * 1.0
    circle(s, 1.0, y, 0.62, n, size=14)
    tb(s, 1.9, y + 0.02, 5.0, 0.4, t, size=18, bold=True)
    tb(s, 1.9, y + 0.42, 8.0, 0.35, d, size=12, color=GREY)
    ln = s.shapes.add_connector(1, Inches(1.9), Inches(y + 0.86), Inches(12.3), Inches(y + 0.86)); ln.line.color.rgb = LINE

# 3 重點摘要
s = content('重點摘要：保險 AI 已從「試點」走到「規模化與回報」，我們的機會在落地與確信')
msgs = [('市場', '資金與預算全面轉向 AI。2026 Q1 全球保險科技募資有 95.2% 流向 AI 相關企業；57% 保險高階主管把生成式與代理型 AI 列為 2026 首要科技投資。'),
        ('痛點', '多數企業 AI 投資還看不到財務回報（Bain 引述 MIT：95%），卡點在流程沒重塑、核心系統過度客製化、數據語意不一致。'),
        ('台灣時機', '2026-01-01 正式接軌 IFRS 17 與 TW-ICS，金管會《金融業運用 AI 指引》與公會自律規範上路，雲端委外規範放寬：合規與轉型需求同時出現。'),
        ('我們的機會', '四條服務線（AI 治理、數據應用、核心系統、人力委外）正好對應落地所需。差異化在「技術落地＋可稽核確信＋長期駐點人力」。')]
for i, (h, t) in enumerate(msgs):
    y = 1.45 + i * 1.18
    tb(s, 0.45, y, 0.12, 1.0, '', fill=OR)
    tb(s, 0.7, y, 1.6, 1.0, h, size=14, bold=True, color=OR, anchor=MSO_ANCHOR.MIDDLE)
    tb(s, 2.3, y, 5.4, 1.0, t, size=12, anchor=MSO_ANCHOR.MIDDLE)
tb(s, 8.1, 1.45, 4.8, 0.45, '建議優先推動的五項服務', size=13, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
pri = ['IFRS 17／TW-ICS 開帳後財務轉型、CSM 與動態 ALM', 'AI 系統清冊、模型驗證與執行期護欄', '代理型 AI 核保與理賠流程重塑',
       'AI 加速核心系統現代化與雲端聯合查核', '深偽理賠詐欺偵測與調查工作流']
for i, t in enumerate(pri):
    y = 2.0 + i * 0.7
    circle(s, 8.2, y + 0.08, 0.42, str(i + 1), fill=OR if i < 2 else OR2, size=12)
    tb(s, 8.75, y, 4.1, 0.6, t, size=11, anchor=MSO_ANCHOR.MIDDLE, fill=LG)
tb(s, 8.1, 5.6, 4.8, 0.95, [[('今天要決定：', {'bold': True, 'color': OR})], '挑 2 項先做成可銷售的服務包。建議 #1、#2：需求確定、離現有服務最近、首案週期短。'],
   size=11, fill=TINT, margin=0.12)

# 4 數字
section('01', '五大轉折', '過去 18 個月（2025 年初至 2026 年 10 月）全球與亞太保險科技的關鍵變化')
s = content('關鍵數字：資金、採用率與風險同時升溫')
stats = [('95.2%', '2026 Q1 全球保險科技募資流向 AI 相關企業的比例', 'Gallagher Re, 2026-05'),
         ('63%', '2025 年底已把 AI 投入實際營運的企業（2024 年為 45%）', 'Gallagher Re, 2026-05'),
         ('57%', '保險高階主管把生成式與代理型 AI 列為 2026 首要科技投資', 'PwC 調查／Insurance Journal, 2025-11'),
         ('95%', '投資 AI 但尚未看到可衡量財務回報的企業', 'Bain 引述 MIT, 2026-01'),
         ('25–50%', '用代理型 AI 做核心系統現代化的生產力提升', 'McKinsey, 2026-04'),
         ('92%', '曾因深偽相關事件遭受財務損失的企業', 'Swiss Re SONAR 2025, 2025-06')]
for i, (n, d, src) in enumerate(stats):
    x = 0.45 + (i % 3) * 4.2; y = 1.45 + (i // 3) * 2.75
    hot = i in (0, 3)
    tb(s, x, y, 3.95, 2.5, '', fill=OR if hot else LG)
    tb(s, x + 0.2, y + 0.15, 3.6, 0.9, n, size=40, bold=True, color=WHITE if hot else OR)
    tb(s, x + 0.2, y + 1.1, 3.6, 0.9, d, size=12, color=WHITE if hot else INK)
    tb(s, x + 0.2, y + 2.05, 3.6, 0.35, '來源：' + src, size=9, color=WHITE if hot else GREY)

s = content('五大轉折與對台灣保險業的意涵')
rows = [['轉折', '已發生的事實（摘要）', '對台灣保險業的意涵'],
        ['① 代理型 AI 重構核保與理賠', 'Manulife、AmTrust 以 AI 核保理賠獲 Celent 獎；22% 保險公司計畫部署代理型 AI（Celent，單一來源）', '從單點智能助理升級為端到端流程，必須同時處理系統串接與人機分工'],
        ['② AI 改寫核心現代化經濟學', '代理型 AI 用於舊程式解析、資料映射、測試生成，生產力 +25–50%（McKinsey）；金管會 2024-08 修正委外規範、2025-10 預告提升法令位階', '「潔淨核心上雲＋外部 AI 代理層」的重構窗口打開'],
        ['③ AI 治理進入實地查核', 'NAIC 2026 年 12 州試點 AI 系統評估工具；MAS 2025-11 發布涵蓋 AI 代理的風險管理指引；金管會 AI 指引＋公會自律規範', '從紙本政策走向可查核的 AI 清冊、模型驗證與緊急關閉機制'],
        ['④ IFRS 17／TW-ICS 正式上路', '2026-01-01 正式適用；IFRS 4 覆蓋法退場；英國經驗顯示第二年結帳仍高度人工', '需求由「合規開帳」轉向結帳自動化、動態 ALM 與 CSM 經營管理'],
        ['⑤ 深偽詐欺與 AI 風險保險', '92% 企業曾因深偽受損（Swiss Re）；Armilla、aiSure 推出 AI 表現保證保險', '理賠反詐欺升級；AI 責任與表現保險成為新商品議題']]
table(s, 0.45, 1.4, 12.45, [2.6, 5.5, 4.35], rows, size=11.5, head_size=12, row_h=0.9)

# 九大趨勢
section('02', '九大趨勢與服務機會', '每項趨勢都對應到台灣痛點、監理要求、我們的服務線與第一個切入案')
s = content('九大趨勢總覽：與我們四條服務線的對應')
D = '●'
rows = [['#', '趨勢', '成熟度', 'AI 治理', '數據應用', '核心系統', '人力委外', '新服務'],
        ['1', '代理型 AI 核保與理賠作業系統', '早期採用', D, D, D, '', ''],
        ['2', 'AI 加速核心系統現代化與雲端合規轉移', '早期採用', D, '', D, D, ''],
        ['3', '營運化 AI 治理、即時護欄與第三方查核', '早期採用', D, D, '', '', ''],
        ['4', 'IFRS 17／TW-ICS 後財務轉型、動態 ALM、CSM', '普及／早期', '', D, D, D, ''],
        ['5', '深偽理賠詐欺偵測與調查工作流', '早期採用', '', D, D, '', ''],
        ['6', 'AI 精算定價自動化與現代數據平台', '早期採用', '', D, D, D, ''],
        ['7', 'API 優先的嵌入式保險生態系平台', '普及／早期', '', D, D, '', ''],
        ['8', '參數型氣候與巨災保險', '早期採用', '', D, '', '', D],
        ['9', 'AI 表現保證保險與資料中心累積風險', '概念／早期', D, '', '', '', D]]
gt = table(s, 0.45, 1.4, 12.45, [0.45, 5.0, 1.4, 1.12, 1.12, 1.12, 1.12, 1.12], rows, size=11, row_h=0.5)
for i in range(1, 10):
    for j in range(3, 8):
        for para in gt.cell(i, j).text_frame.paragraphs:
            para.alignment = PP_ALIGN.CENTER
            for r in para.runs: r.font.color.rgb = OR; r.font.size = Pt(14)
tb(s, 0.45, 6.55, 12.4, 0.35, '● 表示對應的服務線。成熟度依研究報告判斷：概念階段 → 早期採用 → 已普及。', size=9, color=GREY)

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
    s = content(f'趨勢深入（{k + 1}／3）')
    trio = TR[k * 3:k * 3 + 3]
    labels = ['是什麼', '國際案例', '台灣痛點', '服務包裝', '第一步切入']
    tb(s, 0.45, 1.4, 1.35, 0.55, '', fill=None)
    for r, lab in enumerate(labels):
        y = 2.0 + r * 0.95
        tb(s, 0.45, y, 1.35, 0.88, lab, size=11, bold=True, color=OR, anchor=MSO_ANCHOR.MIDDLE, fill=TINT, margin=0.1)
    for c, t in enumerate(trio):
        x = 1.9 + c * 3.68
        tb(s, x, 1.4, 3.55, 0.55, t[0], size=12.5, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.1)
        for r in range(5):
            y = 2.0 + r * 0.95
            tb(s, x, y, 3.55, 0.88, t[r + 1], size=10.5, fill=LG if r % 2 == 0 else WHITE, line=LINE, anchor=MSO_ANCHOR.MIDDLE, margin=0.08)

# 優先順序
section('03', '前五大優先服務', '評估維度：台灣市場需求迫切度 × 與既有服務的距離 × 法規與市場時機')
s = content('優先順序評估：五項服務的三個維度')
rows = [['優先', '服務', '需求迫切度', '與既有服務距離', '法規與市場時機'],
        ['1', 'IFRS 17／TW-ICS 開帳後財務轉型、CSM 與動態 ALM', '極高：2026 首年正式出表', '零距離：延伸既有駐點團隊', '2026-01 接軌、覆蓋法退場'],
        ['2', 'AI 系統清冊、模型驗證與執行期護欄', '高：公會自律規範與金檢壓力', '極近：AI 治理服務升級', '金管會 AI 指引上路；可對標 NAIC 2026 查核工具'],
        ['3', '代理型 AI 核保與理賠流程重塑', '高：要突破單點助理的回報瓶頸', '近：流程數位化＋數據平台', '2026 全球轉向規模化落地'],
        ['4', 'AI 加速核心現代化與雲端聯合查核', '中高：核心老舊、雲端法規鬆綁', '近：核心系統導入＋IT 委外', '委外規範放寬並提升法令位階'],
        ['5', '深偽理賠詐欺偵測與調查工作流', '中高：醫療險與車險損率壓力', '中：需整合技術夥伴', '行動理賠普及，偽造單據增加']]
table(s, 0.45, 1.4, 12.45, [0.7, 3.9, 2.6, 2.6, 2.65], rows, size=12.5, head_size=12.5, row_h=0.88)

s = content('切入與延伸路徑：先用小案取得信任，再擴大為長期合作')
lands = [('IFRS 17 後財務轉型', '首年結帳瓶頸診斷＋CSM 管理儀表板', 'KRD 動態 ALM、通路 CSM KPI、常態精算財會人力委外'),
         ('AI 治理', '4–6 週 AI 清冊盤點與合規落差健檢', '治理平台、代理變數偏見測試、緊急關閉機制、第三方 AI 認證'),
         ('代理型 AI 核保理賠', '6–8 週單一險種代理型工作流藍圖', 'AI 代理編排層、核心系統 API 串接、人機協作作業導入'),
         ('核心現代化與雲端', '雲端委外合規評估，或舊系統 AI 重寫 PoC', '核心中台解耦、同業聯合查核、遷移與測試人力委外'),
         ('深偽反詐欺', '車損照片／醫療單據偽造偵測 PoC', '即時反詐欺評分、關聯圖譜分析、調查派案整合')]
tb(s, 3.0, 1.35, 4.3, 0.4, '切入（Land）', size=12, bold=True, color=OR)
tb(s, 8.0, 1.35, 4.9, 0.4, '延伸（Expand）', size=12, bold=True, color=OR)
for i, (n, a, b) in enumerate(lands):
    y = 1.8 + i * 0.98
    circle(s, 0.45, y + 0.12, 0.6, str(i + 1), fill=OR if i < 2 else OR2, size=14)
    tb(s, 1.2, y, 1.7, 0.85, n, size=11.5, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    tb(s, 3.0, y, 4.3, 0.85, a, size=11, fill=TINT, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
    ar = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(7.4), Inches(y + 0.27), Inches(0.45), Inches(0.32))
    ar.fill.solid(); ar.fill.fore_color.rgb = OR; ar.line.fill.background()
    tb(s, 8.0, y, 4.9, 0.85, b, size=11, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)

s = content('跨趨勢洞察：三條因果鏈決定機會大小')
chains = [('A', '潔淨核心＋數據中台', '代理型 AI 能否跨過「95% 無回報」', '代理型 AI 要跨系統讀寫即時資料；核心過度客製化與語意不一致，會讓多步驟工作流產生複合錯誤。前端 AI 案會帶出第二波核心去客製化與數據平台需求。'),
          ('B', 'IFRS 17 細顆粒度資料庫', '成為 AI 與動態經營的數據底座', '為合規建置的保單現金流、假設與 CSM 資料，可向上串投資部門 ALM、串定價引擎、串通路 KPI，把合規成本轉成經營工具。'),
          ('C', 'AI 攻防＋監理趨同', '治理即服務＋風險移轉的閉環', '快速理賠帶來 AI 偽造攻擊，需要反詐欺 AI；各國監理要求對第三方 AI 負責；通過獨立查核的模型才能被承保。確信能力成為切入點。')]
for i, (k, a, b, d) in enumerate(chains):
    x = 0.45 + i * 4.2
    circle(s, x, 1.45, 0.6, k, size=16)
    tb(s, x + 0.75, 1.42, 3.2, 0.66, a, size=13, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    ar = s.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(x + 1.75), Inches(2.2), Inches(0.4), Inches(0.45))
    ar.fill.solid(); ar.fill.fore_color.rgb = OR2; ar.line.fill.background()
    tb(s, x, 2.75, 3.95, 0.8, b, size=13, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, margin=0.12)
    tb(s, x, 3.7, 3.95, 2.0, d, size=13, fill=LG, margin=0.15)

# 競爭者
section('04', '競爭者與差異化', '國際顧問公司過去 18 個月的公開服務與方法論，以及我們的切入空間')
s = content('競爭者觀察：大家都在做代理型 AI，差別在能不能落地')
rows = [['機構', '公開推出的服務或研究', '方法論重點', '可能的交付侷限'],
        ['Deloitte', 'Agentic AI Blueprint for Insurance；亞太壽險規模化代理型 AI 論文（2026-05）', '四個模組化工具：能力評估、影響力分析、工作流設計、預置代理庫', '預置代理庫直接套到台灣客製化舊核心與繁中條款時，資料對接與在地合規需調適'],
        ['EY', 'EY Nexus for Insurance；AI 風險治理調查與代理型 AI ROI 研究（2026-09）', '平台＋生態系；強調 AI 治理與確信', '偏向新業務快速上線，大型壽險既有核心遷移與 IFRS 17 深度整合較難只靠外掛平台'],
        ['McKinsey', '核保作業系統 UWOS（2026-06）；代理型 AI 核心現代化（2026-04）', '三車道核保分流；原子化代理能力庫；選擇性重寫', '缺少在地法規確信團隊與長期駐點人力'],
        ['Bain', '12 週 Future-back 方法論；三層代理型 AI 架構', '流程重塑結合工程標準（MCP、執行期政策、緊急關閉）', '在台交付規模較小，法規文書與大型核心導入人力有限'],
        ['所內相關團隊（可協作）', 'IFRS 17 提前審計、ALM／CSM 顧問；GenAI ROI 研究', '財會、風險與高階營運指標', '可直接協作；代理型 AI 工程實作可由我們補強']]
table(s, 0.45, 1.4, 12.45, [1.9, 3.7, 3.35, 3.5], rows, size=11.5, head_size=12, row_h=0.9)

s = content('我們的四個差異化策略')
diffs = [('Assurance-by-Design', '設計每個 AI 工作流與核心 API 時，同步內建可稽核軌跡、權限控管與護欄。讓資訊長與風控、法遵長在同一個案子裡買到創新與金檢安心。'),
         ('顧問藍圖＋人力委外', '前端用小規模高價值的健檢或藍圖切入，確立架構後無縫接上數據工程、核心系統與 IFRS 17 委外人力，形成 1–3 年經常性營收。'),
         ('以 IFRS 17 資料庫為樞紐', '從既有財會與精算專案出發，把合規資料延伸到 ALM、定價與通路 KPI。沒有 IFRS 17 實作經驗的科技廠商很難跨過。'),
         ('雲端與第三方 AI 聯合查核', '依委外規範放寬的方向，召集多家保險公司共同委託查核同一雲端或 AI 業者，分攤成本並掌握市場動態。⚠ 需先確認獨立性。')]
for i, (h, d) in enumerate(diffs):
    x = 0.45 + (i % 2) * 6.3; y = 1.45 + (i // 2) * 2.65
    tb(s, x, y, 6.05, 2.45, '', fill=LG)
    tb(s, x, y, 0.12, 2.45, '', fill=OR)
    circle(s, x + 0.35, y + 0.25, 0.55, str(i + 1), size=14)
    tb(s, x + 1.05, y + 0.25, 4.8, 0.55, h, size=15, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    tb(s, x + 0.35, y + 0.95, 5.5, 1.4, d, size=13)

# 下一步
section('05', '下一步與待驗證事項', '請團隊決定的事，以及這份研究的資料可信度')
s = content('下一步：請團隊討論與決定')
tb(s, 0.45, 1.4, 6.1, 0.45, '要決定的事', size=13, bold=True, color=WHITE, fill=OR, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
dec = ['選 2 項服務先做成服務包（建議 #1 IFRS 17 後財務轉型、#2 AI 治理健檢）',
       '每項指派負責人，4 週內產出：服務一頁紙、首案範圍與報價區間、案例素材',
       '選 1 項做成拜訪客戶或研討會用的 AI 短片',
       '確認與所內 IFRS 17、風險團隊的協作方式']
for i, t in enumerate(dec):
    y = 2.0 + i * 0.95
    circle(s, 0.55, y + 0.15, 0.45, str(i + 1), size=12)
    tb(s, 1.15, y, 5.4, 0.8, t, size=13, anchor=MSO_ANCHOR.MIDDLE)
tb(s, 6.85, 1.4, 6.05, 0.45, '使用這份研究前要注意', size=13, bold=True, color=WHITE, fill=INK, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)
cav = [[('資料來源', {'bold': True}), '：由 Gemini Deep Research 彙整公開資料，所有數字與日期放進提案前都要回原文核對'],
       [('單一來源待驗證', {'bold': True}), '：Celent 22% 計畫部署代理型 AI、亞太深偽網路險理賠 +233%、嵌入式保險市場規模、AI 精算定價 +30% 效率'],
       [('點名台灣同業', {'bold': True}), '：報告引用的國內保險公司資訊來自公開新聞，僅供內部討論，對外不點名'],
       [('獨立性', {'bold': True}), '：聯合查核、第三方 AI 認證涉及確信角色，推動前先與 Risk & Quality 確認']]
for i, t in enumerate(cav):
    tb(s, 6.85, 2.0 + i * 1.12, 6.05, 1.02, [t], size=12, fill=LG, anchor=MSO_ANCHOR.MIDDLE, margin=0.12)

# 附錄
s = content('附錄：主要來源')
src = [['類別', '來源（發布日期）'],
       ['監理機關', ['金管會《金融業運用 AI 指引》（2024-08）', '金管會《保險業作業委託他人處理應注意事項》修正（2024-08）；內部作業制度及程序辦法草案（2025-10 預告）', '產壽險公會《保險業使用雲端服務自律規範》', 'NAIC AI Model Bulletin 實施追蹤（2026-04）', 'MAS AI 風險管理指引（2025-11）；香港保監局 AI Cohort（2025-08）']],
       ['產業研究', ['Gallagher Re《Global InsurTech Report Q1 2026》（2026-05）、Q2 2026（2026-08）', 'Swiss Re SONAR 2025（2025-06）；參數型洪水保險（2026-03）', 'Celent Model Insurer 2025／2026；Shedding Light on Agentic AI（2025-05）', 'Armilla AI；Munich Re《Insuring Generative AI》']],
       ['顧問與事務所', ['McKinsey：Agentic AI 核心現代化（2026-04）、UWOS（2026-06）、SaaS in the age of agentic AI（2026-10）', 'Bain：Reimagining Insurance for the AI Era（2026）、How to Architect for Agentic AI', 'Deloitte：Agentic AI Blueprint for Insurance；亞太壽險論文（2026-05）', 'EY：EY Nexus for Insurance；AI Risk and Governance Survey（2026-09）', 'PwC Taiwan：IFRS 17 關鍵議題（2025-11）、首份 IFRS 17 財報導讀（2026）']],
       ['媒體', ['Insurance Journal（2025-11）；工商時報（2024-06、2024-08、2025-09、2026-01、2026-02、2026-05）；Insurance Innovation Reporter（2026-06）']]]
table(s, 0.45, 1.4, 12.45, [1.8, 10.65], src, size=11, head_size=11.5, row_h=0.6)
tb(s, 0.45, 6.6, 12.4, 0.3, '完整連結見研究報告原文。', size=9, color=GREY)

p.save(OUT)
print(OUT, len(p.slides), 'slides')
