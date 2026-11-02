"""Apply Robert's 11/2 review comments to the BA meeting deck (structure already reordered in base_rev.pptx)."""
import sys, copy, math
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.oxml.ns import qn

src, out, shot = sys.argv[1], sys.argv[2], sys.argv[3]
p = Presentation(src)
SL = p.slides

ORANGE, MID, LIGHT, PALE, CREAM = "FD5108", "FE7C39", "FFAA72", "FFE8D4", "FFF5ED"
DARK, GREY, LGREY, TEXT, WHITE = "464646", "7D7D7D", "E8E8E8", "2D2D2D", "FFFFFF"
rgb = lambda h: RGBColor.from_string(h)


# ---------------------------------------------------------------- helpers
def all_shapes(shapes):
    for sh in shapes:
        yield sh
        if sh.shape_type == 6:
            yield from all_shapes(sh.shapes)

def find(slide, sid):
    for sh in all_shapes(slide.shapes):
        if sh.shape_id == sid:
            return sh
    raise KeyError(sid)

def set_tf(tf, items):
    """items: list of (prototype paragraph index, text); keeps template formatting."""
    body = tf._txBody
    ps = body.findall(qn('a:p'))
    protos = list(ps)
    for x in ps:
        body.remove(x)
    for idx, text in items:
        np_ = copy.deepcopy(protos[min(idx, len(protos) - 1)])
        runs = np_.findall(qn('a:r'))
        for r in runs[1:]:
            np_.remove(r)
        for tag in ('a:br', 'a:fld'):
            for x in np_.findall(qn(tag)):
                np_.remove(x)
        if runs:
            runs[0].find(qn('a:t')).text = text
        else:
            r = np_.makeelement(qn('a:r'), {}); t = r.makeelement(qn('a:t'), {}); t.text = text; r.append(t)
            end = np_.find(qn('a:endParaRPr'))
            (end.addprevious(r) if end is not None else np_.append(r))
        body.append(np_)

def st(slide, sid, *lines):
    set_tf(find(slide, sid).text_frame, [(0, l) for l in lines])

def num(slide, sid, n):
    fld = find(slide, sid).text_frame._txBody.find(qn('a:p')).find(qn('a:fld'))
    fld.find(qn('a:t')).text = str(n)

def notes(slide, lines):
    slide.notes_slide.notes_text_frame.text = "\n".join(lines)

def clear(slide, keep_ids):
    for sh in list(slide.shapes):
        if sh.shape_id not in keep_ids:
            sh._element.getparent().remove(sh._element)

def box(slide, x, y, w, h, text, fill, color=TEXT, size=11, bold=False, shape=MSO_SHAPE.RECTANGLE,
        line=None, dash=False, align=PP_ALIGN.CENTER, name=None, lw=0.75, sub_size=None):
    sp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if name: sp.name = name
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = rgb(fill)
    if line:
        sp.line.color.rgb = rgb(line); sp.line.width = Pt(lw)
        if dash: sp.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    else:
        sp.line.fill.background()
    sp.shadow.inherit = False
    tf = sp.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.05); tf.margin_top = tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    lines = text if isinstance(text, list) else [text]
    for i, t in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        r = para.add_run(); r.text = t
        r.font.size = Pt(size if i == 0 else (sub_size or size - 2))
        r.font.bold = bold if i == 0 else False
        r.font.color.rgb = rgb(color)
    return sp

def text(slide, x, y, w, h, lines, size=12, color=TEXT, bold=False, align=PP_ALIGN.LEFT, name=None):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name: tb.name = name
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.02)
    for i, t in enumerate(lines if isinstance(lines, list) else [lines]):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        r = para.add_run(); r.text = t; r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = rgb(color)
    return tb

def line(slide, x1, y1, x2, y2, color=GREY, w=1.25, dash=False):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = rgb(color); c.line.width = Pt(w)
    if dash: c.line.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT
    return c

def circle(slide, cx, cy, d, text_, fill, color=WHITE, size=11, bold=True, line_=None, dash=False, name=None, sub_size=None):
    return box(slide, cx - d / 2, cy - d / 2, d, d, text_, fill, color, size, bold, MSO_SHAPE.OVAL, line_, dash, name=name, sub_size=sub_size)

def edge_point(cx, cy, r, tx, ty):
    a = math.atan2(ty - cy, tx - cx)
    return cx + r * math.cos(a), cy + r * math.sin(a)

def link(slide, a, b, color, w=1.5, dash=False):
    (ax, ay, ar), (bx, by, br) = a, b
    x1, y1 = edge_point(ax, ay, ar, bx, by); x2, y2 = edge_point(bx, by, br, ax, ay)
    return line(slide, x1, y1, x2, y2, color, w, dash)

FOOT = "FSC1 BA 組會（2026/11/06）｜第二次組會分享"

# ================================================================ 1 Agenda
a = SL[0]
rows = [("工作角色盤點", "三條工作線與關鍵成果", "02"),
        ("文章分享", "McKinsey UWOS 與模組化趨勢", "03"),
        ("10 月議題心得", "用 Intent Map 回答四個問題", "05"),
        ("成為 AI builder", "從意圖到動手做", "06"),
        ("團隊發展方向", "2030 的 BA 與三個方向", "07"),
        ("附件", "參考資料", "11")]
tb = find(a, 14)
for i, row in enumerate(rows):
    for j, v in enumerate(row):
        set_tf(tb.table.cell(i, j).text_frame, [(0, v)])
st(a, 2, FOOT)
num(a, 3, 1)
notes(a, ["今天分享五件事：我的三條工作線、Luke 指定的 McKinsey 文章、用 Intent Map 回答 10 月的四個問題、為什麼要成為 AI builder，最後是我對 BA 團隊發展方向的建議。"])

# ================================================================ 2 Work network
w = SL[1]
clear(w, {2, 6, 12, 9})
st(w, 6, "工作角色盤點：三條工作線、三個關鍵成果，彼此互相支援")
st(w, 9, "大圓是關鍵成果，小圓是背後的思考過程與小工作，虛線圓是等待中的事項；點線代表工作線之間互相支援的地方。")
sub = find(w, 9); sub.height = Inches(0.8)

A = (2.75, 4.75, 0.98); B = (6.65, 3.75, 0.98); C = (10.55, 4.75, 0.98); H = (6.65, 5.95, 0.62)
circle(w, A[0], A[1], 1.96, ["IC Care", "對客簡報 v0.5", "13 頁"], ORANGE, size=15, sub_size=11, name="Key IC Care")
circle(w, B[0], B[1], 1.96, ["AI 短影片", "人物誌短片定稿", "Sam 認可"], MID, size=15, sub_size=11, name="Key Video")
circle(w, C[0], C[1], 1.96, ["AI 開發框架", "白話簡報 32 頁", "＋3 頁摘要"], DARK, size=15, sub_size=11, name="Key Framework")
circle(w, H[0], H[1], 1.3, ["我", "動手做的", "翻譯者"], PALE, TEXT, 13, sub_size=10, line_=ORANGE, name="Hub")
for k in (A, B, C):
    link(w, H, k, GREY, 1.25)

def smalls(parent, color, items):
    for (x, y, t, waiting) in items:
        node = (x, y, 0.5)
        if waiting:
            circle(w, x, y, 1.0, t, WHITE, TEXT, 9, False, GREY, True, name="Wait " + t[0])
            link(w, parent, node, GREY, 1, dash=True)
        else:
            circle(w, x, y, 1.0, t, CREAM, TEXT, 9, False, color, name="Step " + t[0])
            link(w, parent, node, color, 1.25)

smalls(A, ORANGE, [(1.0, 3.3, ["標竿研究", "5 家"], False), (0.85, 5.15, ["五模組", "拆解"], False),
                   (1.75, 6.45, ["人物誌", "旅程圖"], False), (3.85, 6.45, ["11 個", "使用情境"], False),
                   (4.25, 3.2, ["等待", "Luke 補充"], True)])
smalls(B, MID, [(4.85, 2.85, ["趨勢研究", "選題前五名"], False), (8.45, 2.85, ["與 Rex、Kenny", "對齊三段架構"], False),
                (8.55, 4.4, ["等待", "Luke 效益構想"], True)])
smalls(C, DARK, [(12.35, 3.3, ["客戶角色", "扮演測試"], False), (12.45, 5.05, ["法規頁", "人的變化頁"], False),
                 (11.75, 6.45, ["給 Sam 的", "3 頁摘要"], False), (9.6, 6.45, ["等待 IT", "實績與報價"], True)])
# cross-links: persona work feeds the persona film; human gates feed the video's go/no-go part
link(w, (1.75, 6.45, 0.5), B, ORANGE, 1.25, dash=True)
link(w, (8.45, 2.85, 0.5), C, DARK, 1.25, dash=True)
text(w, 2.4, 2.35, 3.3, 0.3, "點線：人物誌研究 → 人物誌短片", 9, GREY)
text(w, 9.2, 2.35, 3.6, 0.3, "點線：人工關卡 → 影片第三段 go／no-go", 9, GREY)
tree = w.shapes._spTree
for cx_ in [e for e in tree if e.tag == qn("p:cxnSp")]:
    tree.remove(cx_); tree.insert(2, cx_)
st(w, 12, FOOT)
num(w, 2, 2)
notes(w, ["我用一張網絡圖說明目前的工作。三個大圓是三個關鍵成果：IC Care 的對客簡報做到 v0.5；AI 短影片的人物誌短片定稿，Sam 認可；AI 開發框架完成 32 頁白話簡報和給 Sam 的 3 頁摘要。",
          "小圓是背後的思考過程：IC Care 從標竿研究、五模組拆解，到人物誌和 11 個情境；影片先做趨勢研究，再和 Rex、Kenny 對齊三段架構；框架則用客戶角色扮演找缺口，補上法規頁和人的變化頁。",
          "虛線圓是等待中的事：Luke 的補充與效益構想、IT 團隊的實績與報價。",
          "三條線也互相支援：IC Care 的人物誌研究，直接用在人物誌短片；框架裡的人工關卡，會是影片第三段的 go／no-go 示範。",
          "中間是我自己：一個動手做的翻譯者，把複雜內容變成客戶看得懂、能驗證的成果。"])

# ================================================================ 3 UWOS p1: modular
u = SL[2]
st(u, 19, "趨勢：走向模組化")
st(u, 4, "產業趨勢"); st(u, 5, "AI 規模化"); st(u, 7, "UWOS 怎麼接")
st(u, 8, "從垂直整合走向模組化", "更模組化，但不更一致", "463 位保險高階主管調查")
st(u, 10, "核保與定價 54%", "理賠、詐欺與服務 53%", "合規與資料治理 52%")
st(u, 11, "45 項能力像積木", "先上 1–2 個流程", "驗證後再一塊塊擴大")
st(u, 17, "“Insurers are transitioning from vertically integrated operating models toward more modular architectures.” — BlackRock, 2026 Global Insurance Report")
st(u, 12, "FSC1 BA 組會｜文章分享（資料來源：McKinsey & Company, 2026；BlackRock 2026 Global Insurance Report）")
num(u, 2, 3)
notes(u, ["【一個重點】McKinsey 提出 UWOS：用 Agentic AI 把核保從逐案人工處理，改成機器優先、人類治理。",
          "我想特別強調「模組化」。BlackRock 2026 年全球保險報告調查了 463 位保險高階主管，結論是保險業正從垂直整合的營運模式，走向模組化架構，而且是「更模組化，但不更一致」，每家公司組合不同。",
          "AI 也從試驗走向落地：受訪者中，已在核保與定價規模化或實際營運使用 AI 的比例是 54%，理賠、詐欺與服務 53%，合規與資料治理 52%。",
          "所以 UWOS 的 45 項能力不是一次全部導入，而是像積木：先挑 1–2 個流程上線，驗證後再一塊塊擴大。這和下一頁的落地四步驟是同一個邏輯。",
          "三條車道：直通、輔助決策、專家判斷；人的角色從處理個案轉向監督與例外處理。"])

# ================================================================ 4 UWOS p2
u2 = SL[3]
num(u2, 1092, 4)

# ================================================================ 5 Reflection x Intent Map
q = SL[4]
cols = [(56, "聽懂指令｜誰的問題", ["原本：自己整理主題，做完一版才發現不夠好", "改變：先問清楚給誰看、要改變什麼再動手", "下一步：分清問路與派工，問路要早問"]),
        (58, "持續對齊｜成功訊號", ["原本：做完才拿給主管看", "改變：寫腳本前先和 Rex、Kenny 對齊三段架構", "下一步：事先和主管講好怎樣算成功"]),
        (48, "轉成業務語言｜三個答案", ["原本：簡報用縮寫和工程用語", "改變：站在客戶角度改寫，例如「金檢時拿得出紀錄」", "下一步：demo＋做什麼、不做什麼、怎麼知道成了"]),
        (44, "成為有效成員｜邊界", ["原本：專注做完自己的兩條工作線", "改變：主動補位，整理影片框架與組會筆記", "下一步：接任務先確認做與不做、期限"])]
for sid, head, body in cols:
    set_tf(find(q, sid).text_frame, [(0, head)] + [(1, b) for b in body])
st(q, 7, "10 月議題心得：用 Intent Map 回答四個問題")
text(q, 0.45, 1.3, 12.4, 0.5, "四個問題其實都在問同一件事：需求交出去之前，意圖有沒有說清楚。", 16, DARK, name="Intent subtitle")
st(q, 2, "議題：AI 影片製作、AI 開發管理框架｜對照 Shaun《Lesson 04 產品意圖卡（Intent Map）》")
num(q, 3, 5)
notes(q, ["我參與的是 AI 影片製作和 AI 開發管理框架兩個議題。這次我用自己整理的 Intent Map（產品意圖卡）回頭看這四個問題，發現它們其實都在問：意圖有沒有說清楚。",
          "【草稿，請改成自己的話】",
          "聽懂指令，對應「誰的問題」：原本我自己整理主題，做完一版才覺得不夠好；現在先問清楚給誰看、要改變什麼。也要分清楚「問路」和「派工」：問現有能力到哪、誰負責是問路，成本很低，應該早問。",
          "持續對齊，對應「成功訊號」：沒有成功訊號的需求，無法被驗證，也無法委派。所以這次寫腳本前，先和 Rex、Kenny 對齊三段架構，下一步是事先和主管講好怎樣算成功。",
          "轉成業務語言，對應「三個答案」：對外說明時，demo 一定搭配三句話：做什麼、不做什麼、怎麼知道成了。",
          "成為有效成員，對應「邊界」：接任務先確認做與不做，以及期限，避免把需求讀成無限責任。",
          "這也帶到下一頁：為什麼寫好意圖，是成為 AI builder 的第一步。"])

# ================================================================ 6 Bridge: intent -> AI builder
g = SL[5]
st(g, 4, "從意圖到 AI builder：先說清楚要什麼，AI 才幫得上忙")
st(g, 49, "AI-native SDLC 裡，提出想法的人直接和 AI 討論，把需求寫成一份意圖說明（intent.md）；起草、計畫、寫程式、測試都可以交給 AI 先做。人只做三件事：說清楚問題、決定要不要做、決定能不能上線。")
st(g, 1077, "Intent Map")
set_tf(find(g, 1076).text_frame, [(0, "需求交出去之前，分清「我想要世界變成什麼樣」和「我猜要怎麼做」，只交出前者。"), (0, "沒有成功訊號的需求，無法被驗證，也無法委派。")])
st(g, 1078, "為什麼重要")
st(g, 1075, "會寫意圖，就能把工作交給同事、資料團隊或 AI 代理，再用 AI 快速做出 demo 驗證。這就是 AI builder 的核心能力。")
st(g, 1062, "Intent Map 的關鍵欄位：")
st(g, 8, "誰的問題", "叫得出名字的人")
st(g, 13, "改變什麼", "行為與數字")
st(g, 21, "成功訊號", "事先講好的門檻")
st(g, 28, "邊界", "不做與期限")
st(g, 50, "過去：交出解法…")
st(g, 61, "…將會變成…")
st(g, 51, "AI builder：交出意圖＋demo")
st(g, 62, "主詞是「請你＋動詞」")
st(g, 1026, "只有過程指標，沒有成功訊號")
st(g, 1029, "沒有終點，被讀成無限責任")
st(g, 1042, "主詞是使用者")
st(g, 1045, "成功訊號事先訂好")
st(g, 1093, "寫出刻意不做的清單")
st(g, 1094, "需求訪談：用五個欄位整理客戶的話，當場確認")
st(g, 1095, "委派工作：交給同事或 AI 代理前，先過意圖閘")
st(g, 1096, "對客提案：demo＋做什麼、不做什麼、怎麼知道成了")
st(g, 2, "資料來源：Shaun《Lesson 04 產品意圖卡（Intent Map）》；Anthropic《The AI-Native SDLC Playbook》")
num(g, 1092, 6)
notes(g, ["這一頁把上一頁的 Intent Map，和 AI 開發框架連起來。",
          "在 AI-native SDLC 裡，提出想法的人直接和 AI 討論，把需求寫成一份意圖說明；後面的起草、計畫、寫程式、測試都可以交給 AI 先做。整個過程中，人只做三件事：說清楚問題、決定要不要做、決定能不能上線。",
          "所以 AI 時代最重要的能力，不是自己寫每一份文件，而是把意圖說清楚。Intent Map 有五個欄位：誰的問題、改變什麼、成功訊號、不做會怎樣、邊界；這頁列出最關鍵的四個。",
          "常見的壞意圖有三種：主詞是「請你做某件事」，其實是在交解法；只有覆蓋率這類過程指標，沒有成功訊號；沒有終點，被讀成無限責任。",
          "會寫意圖的人，就能把工作交給同事、資料團隊或 AI 代理，再用 AI 快速做出 demo 驗證，這就是 AI builder。",
          "這個能力也能延伸到其他工作：需求訪談時用五個欄位整理客戶的話；委派工作前先過意圖閘；對客提案時永遠是 demo 加三個答案。",
          "接下來我想談，為什麼現在正是成為 AI builder 的時候。"])

# ================================================================ 7 Why now (Wang)
num(SL[6], 3, 7)

# ================================================================ 8 2030 BA loop
m = SL[7]
keep8 = {sh.shape_id for sh in m.shapes if sh.is_placeholder}
clear(m, keep8)
m.shapes.title.text_frame.paragraphs[0].runs[0].text = "2030 的 BA：角色在內圈，能力成外圈的循環"
# before
circle(m, 1.75, 4.35, 2.7, ["今天的 BA", "寫需求，交給技術團隊", "文件靠人工", "知識在個人腦中", "測試太晚", "AI 隨手用，客戶不敢信"], LGREY, TEXT, 15, True, GREY, name="Before", sub_size=10)
box(m, 3.3, 4.05, 0.42, 0.6, "", ORANGE, shape=MSO_SHAPE.CHEVRON, name="Arrow")
cx, cy, R = 7.75, 4.35, 2.5
ring = m.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - R), Inches(cy - R), Inches(2 * R), Inches(2 * R))
ring.fill.background(); ring.line.color.rgb = rgb(LIGHT); ring.line.width = Pt(3); ring.name = "Capability loop"
disk = m.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 1.72), Inches(cy - 1.72), Inches(3.44), Inches(3.44))
disk.fill.solid(); disk.fill.fore_color.rgb = rgb(PALE); disk.line.fill.background(); disk.name = "Inner ring"
circle(m, cx, cy, 1.6, ["2030 的 BA", "動手做的", "翻譯者"], ORANGE, WHITE, 13, sub_size=11, name="Core")
roles = [("現場共創", ORANGE, WHITE, -135), ("脈絡與 AI 代理", MID, WHITE, -45), ("營運模式", LIGHT, TEXT, 45), ("信任設計", DARK, WHITE, 135)]
for name, fill, fg, ang in roles:
    rx, ry = cx + 1.2 * math.cos(math.radians(ang)), cy + 1.2 * math.sin(math.radians(ang))
    box(m, rx - 0.72, ry - 0.2, 1.44, 0.4, name, fill, fg, 11, True, MSO_SHAPE.ROUNDED_RECTANGLE, name="Role " + name)
skills = ["AI 動手做原型", "產品排序思維", "業務架構視野", "脈絡工程", "驗收與評估設計", "AI 治理素養", "用業務語言說故事", "變革領導"]
for i, s in enumerate(skills):
    ang = -90 + i * 45
    sx, sy = cx + R * math.cos(math.radians(ang)), cy + R * math.sin(math.radians(ang))
    box(m, sx - 0.8, sy - 0.24, 1.6, 0.48, s, WHITE, TEXT, 11, True, MSO_SHAPE.ROUNDED_RECTANGLE, line=ORANGE, lw=1.25, name="Skill " + s)
    mid = math.radians(ang + 22.5)
    tx, ty = cx + R * math.cos(mid), cy + R * math.sin(mid)
    tri = m.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(tx - 0.09), Inches(ty - 0.08), Inches(0.18), Inches(0.16))
    tri.rotation = ang + 22.5 + 90 + 90  # point clockwise along the ring
    tri.fill.solid(); tri.fill.fore_color.rgb = rgb(ORANGE); tri.line.fill.background(); tri.name = "Loop arrow"
box(m, 11.0, 1.95, 1.95, 0.42, "怎麼讀這張圖", DARK, WHITE, 12, True, name="Legend head")
text(m, 11.0, 2.45, 1.95, 3.6, ["內圈：四個新角色", "", "外圈：八項能力，彼此帶動，形成持續精進的循環", "", "和今天相比：從交出文件，變成交出看得到、能驗證的成果"], 11, TEXT, name="Legend")
for sh in m.shapes:
    if sh.is_placeholder and sh.placeholder_format.type is not None and 'FOOTER' in str(sh.placeholder_format.type):
        sh.text_frame.paragraphs[0].runs[0].text = "內圈角色與外圈能力依 graphify 知識圖譜整理（5 份公開資料、80 個概念、8 個群組）"
        for r in sh.text_frame.paragraphs[0].runs[1:]: r.text = ""
for sh in m.shapes:
    if sh.is_placeholder and 'SLIDE_NUMBER' in str(sh.placeholder_format.type):
        sh.text_frame._txBody.find(qn('a:p')).find(qn('a:fld')).find(qn('a:t')).text = "8"
notes(m, ["【願景，約 60 秒】左邊是今天的 BA：寫需求交給技術團隊、文件靠人工、知識在個人腦中、測試太晚、AI 隨手用但客戶不敢信。",
          "右邊是我想像的 2030 的 BA。中心是一個動手做的翻譯者；內圈是四個新角色：現場共創、脈絡與 AI 代理、信任設計、營運模式。",
          "外圈是八項能力，排成一個循環：動手做原型、產品排序、業務架構、脈絡工程、驗收與評估設計、AI 治理、用業務語言說故事、變革領導。它們會彼此帶動，例如原型做得越多，越懂業務架構；越懂架構，脈絡整理得越好。",
          "這張圖是我用 graphify 把五份公開資料做成知識圖譜後整理的，屬於願景，不是現有職缺的要求。",
          "下一頁說明，全球市場上已經出現在要求這些能力的新職缺：Forward Deployed BA。"])

# ================================================================ 9 Forward deployed
f = SL[8]
clear(f, {2, 6, 12, 9})
st(f, 6, "Forward Deployed 是什麼：走到使用者身邊，一起把東西做出來")
st(f, 9, "Forward 是往前，走到第一線；Deployed 是派駐，嵌進客戶或業務團隊一起工作。這個說法原本用在工程師（Forward Deployed Engineer），現在全球金融業也開始用在 BA。")
find(f, 9).height = Inches(1.1)
# left: why it matters
box(f, 0.48, 2.55, 3.75, 0.45, "為什麼對任何領域的 BA 都重要", ORANGE, WHITE, 13, True, name="Why head")
why = [("1", "寫需求變便宜了", "價值移到現場驗證"), ("2", "客戶要的是成果", "看得到、能測試的東西"), ("3", "懂業務又能動手", "這樣的人最難被取代")]
for i, (n_, h, s) in enumerate(why):
    y = 3.15 + i * 1.22
    box(f, 0.48, y, 0.5, 1.05, n_, ORANGE, WHITE, 18, True, MSO_SHAPE.SNIP_2_DIAG_RECTANGLE, name="Why num " + n_)
    box(f, 1.05, y, 3.18, 1.05, [h, s], CREAM, TEXT, 13, True, line=LIGHT, align=PP_ALIGN.LEFT, name="Why " + n_, sub_size=11)
# middle: screenshot of duties
text(f, 4.45, 2.55, 2.9, 0.35, "職缺「職責」原文截圖", 11, GREY, True, name="Shot caption")
pic = f.shapes.add_picture(shot, Inches(4.45), Inches(2.9), height=Inches(4.05))
pic.line.color.rgb = rgb(LGREY); pic.line.width = Pt(0.75); pic.name = "Job duties screenshot"
# right: mapping to capabilities
box(f, 7.55, 2.55, 5.3, 0.45, "職責 → 對應上一頁的能力", DARK, WHITE, 13, True, name="Map head")
maprows = [("把模糊問題轉成可測試的假設", "驗收與評估設計"), ("建立結構化業務脈絡給 AI 用", "脈絡工程"),
           ("把流程拆成邏輯、控制與例外", "業務架構視野"), ("用 AI 開發工具做可運作原型", "AI 動手做原型"),
           ("形塑目標營運模式與治理", "AI 治理素養"), ("從探索到上線，持續和各方協作", "用業務語言說故事"),
           ("訂可衡量目標、凝聚共識", "產品排序思維")]
for i, (d, c) in enumerate(maprows):
    y = 3.08 + i * 0.555
    box(f, 7.55, y, 3.3, 0.48, f"{i+1}  {d}", LGREY, TEXT, 11, align=PP_ALIGN.LEFT, name=f"Duty {i+1}")
    box(f, 10.9, y + 0.12, 0.25, 0.24, "", ORANGE, shape=MSO_SHAPE.CHEVRON, name=f"Duty arrow {i+1}")
    box(f, 11.2, y, 1.65, 0.48, c, WHITE, TEXT, 11, True, MSO_SHAPE.ROUNDED_RECTANGLE, line=ORANGE, lw=1.25, name=f"Cap {i+1}")
st(f, 12, "職缺：國際金融機構 Forward Deployed BA（2026-10）職責原文截圖；對應為解讀，「變革領導」為另外加上")
num(f, 2, 9)
notes(f, ["Forward Deployed 可以拆成兩個字：Forward 是往前，走到第一線，不在後方等需求；Deployed 是派駐，嵌進客戶或業務團隊裡一起工作。",
          "這個說法原本用在工程師，叫 Forward Deployed Engineer；現在全球金融業開始用在 BA，出現了 Forward Deployed Business Analyst 這種職缺。",
          "中間是職缺的職責原文截圖。右邊我把七條職責，一一對應到上一頁外圈的能力：模糊問題轉成可測試假設，對應驗收與評估設計；結構化業務脈絡，對應脈絡工程；拆流程，對應業務架構；用 AI 做原型，對應動手做原型；形塑營運模式與治理，對應 AI 治理素養；持續協作，對應用業務語言說故事；訂目標、凝聚共識，對應產品排序思維。第八項變革領導，是我們另外加上的。",
          "為什麼對任何領域的 BA 都重要？第一，寫需求變便宜了，價值移到現場驗證；第二，客戶要的是看得到、能測試的成果；第三，懂業務又能動手的人，最難被取代。",
          "下一頁是我提的三個團隊方向。"])

# ================================================================ 10 Directions (no company name)
d = SL[9]
st(d, 9, "全球市場已經出現「Forward Deployed Business Analyst」這類新職缺，要求 BA 結合結構化分析、產品思維與動手做原型的能力。職缺列出的職責，正好對應我們可以發展的三個方向，也對應一年目標中的價值、品質與效率。")
st(d, 17, "“This is not a traditional business analysis role centered on gathering requirements and handing them to technology.” — Forward Deployed BA job posting (2026)")
st(d, 8, "用 AI 做可操作原型", "發掘需求、挑戰假設", "全球新職缺要求")
st(d, 12, "FSC1 BA 組會｜團隊發展方向提案（資料來源：全球金融業職缺 2026-10、金管會 AI 指引 2024-06、Gartner 2025-10）")
num(d, 2, 10)
notes(d, ["【約 2 分鐘】這三個方向，職缺原文也講得很直接：這不是傳統只蒐集需求、再交給技術團隊的 BA。",
          "方向一，用 AI 現場做原型，和客戶一起確認需求。國發專案已經在做，客戶當場就指出哪些欄位不需要、哪些流程太複雜。建議整理成團隊標準流程，一年內至少兩個專案在工作坊用上，衡量原型階段抓到的需求落差。",
          "方向二，替 AI 的產出定義驗收標準，設計哪一步由人把關。客戶最大的疑慮是 AI 做的東西能不能信；金管會 AI 指引也要求「人在迴圈內」。寫驗收標準本來就是 BA 的工作，建議先做一份 BA 版的 AI 產出審查清單。",
          "方向三，把業務規則整理成 AI 讀得懂的知識。AI 讀不到公司內部規格，問題在於缺少脈絡。Gartner 預測到 2028 年，企業用的生成式 AI 模型超過一半會是領域專用的。第一步挑一個專案，把需求格式、業務規則、名詞定義整理成知識檔。",
          "總結：現場做原型對應價值，AI 驗收標準對應品質，領域脈絡對應效率；AI 省下來的時間拿來沉澱成團隊範本，就是成長。"])

# ================================================================ 11 References (9 rows)
rf = SL[10]
tbl = find(rf, 15).table
trs = tbl._tbl.findall(qn('a:tr'))
for _ in range(2):
    tbl._tbl.append(copy.deepcopy(trs[-1]))
refs = [("#", "資料", "發布單位", "日期", "用在哪一頁"),
        ("1", "Forward Deployed Business Analyst 職缺", "國際金融機構", "2026-10", "9、10"),
        ("2", "金融業運用人工智慧（AI）指引", "金管會", "2024-06", "10"),
        ("3", "保險業運用人工智慧系統自律規範", "壽險／產險公會", "2025-04 備查", "8"),
        ("4", "Top Strategic Technology Trends for 2026", "Gartner", "2025-10", "10"),
        ("5", "This is a Once-in-a-Civilization Opportunity", "Y Combinator", "2026-07", "7"),
        ("6", "Underwriting Operating System（UWOS）", "McKinsey & Company", "2026", "3、4"),
        ("7", "Different by Design: 2026 Global Insurance Report", "BlackRock", "2026-10", "3"),
        ("8", "Lesson 04 產品意圖卡（Intent Map）", "Shaun AI Lesson Learning", "—", "5、6"),
        ("9", "The AI-Native SDLC Playbook", "Anthropic", "—", "6")]
for i, row in enumerate(refs):
    for j, v in enumerate(row):
        set_tf(tbl.cell(i, j).text_frame, [(0, v)])
find(rf, 15).top = Inches(1.55)
st(rf, 2, FOOT)
num(rf, 3, 11)
notes(rf, ["參考資料都已打開原文核對過；Alexandr Wang 的對談只確認過大意，沒有引用原話；職缺截圖只擷取職責段落。"])

p.save(out)
print("saved", out)
