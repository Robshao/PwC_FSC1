"""Apply Robert's second round of review comments (11/2) to the BA meeting deck."""
import sys, math
from lxml import etree
exec(open(sys.argv[4], encoding='utf8').read())   # helpers: find, set_tf, st, num, notes, clear, box, text, line, circle, colours

src, out, shot = sys.argv[1], sys.argv[2], sys.argv[3]
p = Presentation(src)
SL = p.slides
FOOT = "FSC1 BA 組會（2026/11/06）｜第二次組會分享"

def ring_arrows(slide, cx, cy, rx, ry, angles, color=ORANGE):
    for a in angles:
        r = math.radians(a)
        x, y = cx + rx * math.cos(r), cy + ry * math.sin(r)
        tri = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(x - 0.1), Inches(y - 0.09), Inches(0.2), Inches(0.18))
        tri.rotation = a + 180
        tri.fill.solid(); tri.fill.fore_color.rgb = rgb(color); tri.line.fill.background(); tri.name = "Loop arrow"

def ph(slide, kind):
    for sh in slide.shapes:
        if sh.is_placeholder and kind in str(sh.placeholder_format.type):
            return sh

def set_ph_text(sh, t):
    runs = sh.text_frame.paragraphs[0].runs
    runs[0].text = t
    for r in runs[1:]: r.text = ""

# ================================================================ 1 Agenda
a = SL[0]
rows = [("工作角色盤點", "三條工作線與關鍵成果", "02"),
        ("文章分享", "McKinsey UWOS 與模組化趨勢", "03"),
        ("10 月議題心得", "什麼是意圖：Intent Map", "05"),
        ("釐清意圖", "減少衝突，成為 AI builder", "06"),
        ("新時代的 BA", "2030 的 BA 與三個方向", "07"),
        ("附件", "參考資料", "09")]
tb = find(a, 14)
for i, row in enumerate(rows):
    for j, v in enumerate(row):
        set_tf(tb.table.cell(i, j).text_frame, [(0, v)])
notes(a, ["今天分享五件事：我的三條工作線、Luke 指定的 McKinsey 文章、什麼是意圖以及我怎麼用它回答 10 月的四個問題、為什麼釐清意圖能減少衝突並成為 AI builder 的第一步，最後是新時代 BA 的樣子與三個團隊方向。"])

# ================================================================ 3 UWOS: modular in plain words
u = SL[2]
st(u, 19, "模組化：像組積木，不是蓋大樓")
st(u, 4, "以前：一體成形"); st(u, 5, "現在：組合積木"); st(u, 7, "對核保的意思")
st(u, 8, "一家公司從頭做到尾", "系統綁在一起，難改動", "要改就得整套大改")
st(u, 10, "核心自己做，其他找夥伴", "可以一塊一塊替換", "每家的組合都不一樣")
st(u, 11, "能力一塊一塊上線", "先做最痛的那一塊", "驗證有效再接下一塊")
notes(u, ["【一個重點】McKinsey 提出 UWOS：用 Agentic AI 把核保從逐案人工處理，改成機器優先、人類治理。",
          "我想特別強調「模組化」。BlackRock 今年的全球保險報告指出，保險業正從垂直整合的營運模式，走向模組化架構。",
          "白話來說：以前像蓋一棟大樓，一家公司從頭做到尾，系統綁在一起，要改就得整套大改；現在像組積木，核心能力自己做，其他找夥伴，可以一塊一塊替換，而且每家的組合都不一樣。",
          "放到核保上，UWOS 的能力也不是一次全部導入，而是先挑最痛的那一塊上線，驗證有效再接下一塊。下一頁的落地四步驟就是這個邏輯。",
          "中間是三條車道：直通、輔助決策、專家判斷；右邊是人的角色，從處理個案轉向監督與例外處理。"])

# ================================================================ 5 What is intent
q = SL[4]
clear(q, {2, 3, 7})
st(q, 7, "10 月議題心得：四個問題，其實都在問意圖有沒有說清楚")
# left: dictionary definition
box(q, 0.45, 1.55, 3.9, 0.42, "Cambridge Dictionary：intent（名詞）", DARK, WHITE, 12, True, name="Dict head")
box(q, 0.45, 2.0, 3.9, 1.3, ["“the fact that you want and plan to do something”", "你「想要」、並且「打算」去做某件事"],
    CREAM, TEXT, 13, True, line=LIGHT, align=PP_ALIGN.LEFT, name="Dict quote", sub_size=12)
box(q, 0.45, 3.45, 1.9, 1.25, ["想要（want）", "我希望世界變成什麼樣", "→ 交出去"], ORANGE, WHITE, 13, True, name="Want", sub_size=11)
box(q, 2.45, 3.45, 1.9, 1.25, ["打算（plan）", "我猜要怎麼做", "→ 先留著，一起找"], LGREY, TEXT, 13, True, name="Plan", sub_size=11)
text(q, 0.45, 4.78, 3.9, 0.55, "Intent Map 的關鍵：把這兩件事分開，只交出「想要」。", 11, DARK, True, name="Key")
box(q, 0.45, 5.35, 3.9, 0.36, "對應 10 月的四個問題", MID, WHITE, 11, True, name="Q head")
text(q, 0.5, 5.75, 3.85, 1.2, ["聽懂指令：先問「誰的問題」", "持續對齊：事先講好「成功訊號」", "轉成業務語言：demo＋三個答案", "成為有效成員：講清楚「邊界」"], 11, TEXT, name="Q list")
# right: discussion loop
text(q, 4.65, 1.5, 8.2, 0.4, "在金融科技公司，我們這樣討論一個新想法，一輪一輪迭代出共識", 13, DARK, True, name="Loop title")
cx, cy, rx, ry = 8.75, 4.35, 3.0, 1.95
ring = q.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - rx), Inches(cy - ry), Inches(2 * rx), Inches(2 * ry))
ring.fill.background(); ring.line.color.rgb = rgb(LIGHT); ring.line.width = Pt(3); ring.name = "Discussion loop"
circle(q, cx, cy, 1.35, ["共識", "大家對齊", "同一件事"], ORANGE, WHITE, 16, sub_size=10, name="Consensus")
steps = [(-90, "① 新想法", "先分清：問路，還是派工？", WHITE),
         (-18, "② 填意圖卡", "誰的問題・改變什麼・成功訊號・不做會怎樣・邊界", CREAM),
         (54, "③ 過三道閘", "主詞是使用者？成功訊號先訂了？邊界畫了？", PALE),
         (126, "④ 出手", "demo＋做什麼／不做什麼／怎麼知道成了", CREAM),
         (198, "⑤ 回饋", "記下填不出來的格子，下一輪先補", WHITE)]
for ang, head, sub, fill in steps:
    r = math.radians(ang)
    x, y = cx + rx * math.cos(r), cy + ry * math.sin(r)
    box(q, x - 1.25, y - 0.52, 2.5, 1.04, [head, sub], fill, TEXT, 13, True, MSO_SHAPE.ROUNDED_RECTANGLE,
        line=ORANGE, lw=1.25, name="Step " + head, sub_size=10)
ring_arrows(q, cx, cy, rx, ry, [-54 + 0, 18, 90, 162, 234])
set_ph_text(find(q, 2), "資料來源：Cambridge Dictionary「intent」；Shaun《Lesson 04 產品意圖卡（Intent Map）》")
num(q, 3, 5)
notes(q, ["10 月的四個問題，我發現其實都在問同一件事：意圖有沒有說清楚。",
          "先看字典。Cambridge Dictionary 對 intent 的解釋是：the fact that you want and plan to do something，你「想要」、並且「打算」去做某件事。",
          "這個定義裡其實藏著兩件事：想要什麼，和打算怎麼做。Intent Map 的關鍵就是把它們分開，只交出「想要」，因為「打算」常常是我們自己猜的做法，收到的人沒辦法幫你修正一個看不見的意圖。",
          "右邊是我在金融科技公司學到的討論方式：有新想法時，先分清是在問路還是派工；派工就填意圖卡的五個欄位；再過三道閘：主詞是不是使用者、成功訊號有沒有事先訂、邊界有沒有畫；然後一定是 demo 加三個答案一起出手；最後把填不出來的格子記下來，下一輪先補。這樣一輪一輪轉，大家就會對齊同一件事。",
          "回到四個問題：聽懂指令，是先問誰的問題；持續對齊，是事先講好成功訊號；轉成業務語言，是 demo 加三個答案；成為有效成員，是講清楚邊界。"])

# ================================================================ 6 Intent reduces conflict -> AI builder
g = SL[5]
st(g, 4, "釐清意圖能減少衝突，也是成為 AI builder 的第一步")
st(g, 49, "很多衝突不是能力問題，而是每個人心裡的「想要」不一樣：業務講結果、技術講做法、主管看風險。把意圖寫成同一張卡，大家先對齊要什麼，再討論怎麼做；AI 也要靠這張卡才幫得上忙。")
st(g, 1077, "衝突從哪裡來")
set_tf(find(g, 1076).text_frame, [(0, "業務說「要快」，技術聽成「要省」，主管在意「別出事」。"), (0, "每個人都對，但講的不是同一件事。")])
st(g, 1078, "為什麼能減少衝突")
st(g, 1075, "意圖卡讓大家用同一套語言：誰的問題、改變什麼、怎樣算成功、做到哪裡停。技術與非技術背景的人都看得懂。")
st(g, 1062, "意圖卡讓每個角色都看到自己在意的那一格：")
st(g, 8, "誰的問題", "業務最在意")
st(g, 13, "改變什麼", "使用者看得到")
st(g, 21, "成功訊號", "主管能驗收")
st(g, 28, "邊界", "技術知道停點")
st(g, 50, "沒有意圖卡：各說各話…")
st(g, 61, "…將會變成…")
st(g, 51, "有意圖卡：先對齊要什麼")
st(g, 62, "需求被讀成無限責任")
st(g, 1026, "做完才發現不是要的")
st(g, 1029, "業務與技術覺得對方不懂")
st(g, 1042, "邊界寫清楚，知道做到哪停")
st(g, 1045, "成功訊號事先講好")
st(g, 1093, "大家用同一張卡溝通")
st(g, 1094, "AI builder：意圖寫清楚，AI 才能起草、計畫、做出 demo")
st(g, 1095, "人只做三件事：說清楚問題、決定做不做、決定能不能上線")
st(g, 1096, "同一套方法，延伸到需求訪談、委派工作與對客提案")
notes(g, ["為什麼要花時間釐清意圖？因為很多衝突不是能力問題，而是每個人心裡的「想要」不一樣。",
          "例如業務說要快，技術聽成要省，主管在意別出事；每個人都對，但講的不是同一件事。",
          "意圖卡讓每個角色都看到自己在意的那一格：業務看誰的問題，使用者看得到改變，主管用成功訊號驗收，技術從邊界知道做到哪裡停。技術與非技術背景的人，用的是同一套語言。",
          "沒有意圖卡時，常見三種衝突：需求被讀成無限責任、做完才發現不是要的、業務和技術互相覺得對方不懂。有了意圖卡，邊界寫清楚、成功訊號事先講好、大家用同一張卡溝通。",
          "這也是成為 AI builder 的第一步。在 AI-native SDLC 裡，意圖寫清楚，AI 才能起草、計畫、做出 demo；人只做三件事：說清楚問題、決定做不做、決定能不能上線。",
          "同一套方法也能延伸到需求訪談、委派工作和對客提案。"])

# ================================================================ 7 Loop: small fading "before" circle
m = SL[6]
bef = find(m, 8)
bef.left, bef.top, bef.width, bef.height = Inches(0.75), Inches(3.5), Inches(1.7), Inches(1.7)
set_tf(bef.text_frame, [(0, "今天的 BA"), (1, "寫需求、交出去")])
arw = find(m, 9)
arw.left, arw.top = Inches(2.7), Inches(4.05)
notes(m, ["左邊小小的灰色圓，是今天的 BA：寫需求、交出去。按一下，它會淡出，代表這種做法正在退場。",
          "右邊是 2030 的 BA。中心是一個動手做的翻譯者；內圈是四個新角色：現場共創、脈絡與 AI 代理、信任設計、營運模式。",
          "外圈是八項能力，排成一個循環：動手做原型、產品排序、業務架構、脈絡工程、驗收與評估設計、AI 治理、用業務語言說故事、變革領導。它們會彼此帶動，例如原型做得越多，越懂業務架構；越懂架構，脈絡整理得越好。",
          "這張圖是我用 graphify 把五份公開資料做成知識圖譜後整理的，屬於願景，不是現有職缺的要求。下一頁說明，全球市場已經出現在要求這些能力的新職缺。"])
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
timing = f'''<p:timing xmlns:p="{P}"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>
<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>
<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>
<p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
{''.join(f"""<p:par><p:cTn id="{5 + k * 3}" presetID="10" presetClass="exit" presetSubtype="0" fill="hold" grpId="0" nodeType="{'clickEffect' if k == 0 else 'withEffect'}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
<p:animEffect transition="out" filter="fade"><p:cBhvr><p:cTn id="{6 + k * 3}" dur="1200"/><p:tgtEl><p:spTgt spid="{sid}"/></p:tgtEl></p:cBhvr></p:animEffect>
<p:set><p:cBhvr><p:cTn id="{7 + k * 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="1199"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{sid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="hidden"/></p:to></p:set>
</p:childTnLst></p:cTn></p:par>""" for k, sid in enumerate([8, 9]))}
</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>
<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>
</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst><p:bldP spid="8" grpId="0" animBg="1"/><p:bldP spid="9" grpId="0" animBg="1"/></p:bldLst></p:timing>'''
sld = m._element
for old in sld.findall(qn('p:timing')):
    sld.remove(old)
tnode = etree.fromstring(timing)
anchor = sld.find(qn('p:clrMapOvr'))
if anchor is None: anchor = sld.find(qn('p:cSld'))
tr = sld.find(qn('p:transition'))
(tr if tr is not None else anchor).addnext(tnode)

# ================================================================ 8 New-era BA (FD + three directions)
f = SL[7]
clear(f, {2, 6, 12, 9})
st(f, 6, "新時代的 BA：走到使用者身邊，從三個方向開始")
st(f, 9, "Forward 是往前、走到第一線；Deployed 是派駐、嵌進團隊一起工作。全球金融業已出現 Forward Deployed BA 職缺，它的職責正好對應我們可以發展的三個方向。")
find(f, 9).height = Inches(0.85)
text(f, 0.48, 2.3, 2.6, 0.3, "職缺「職責」原文截圖", 10, GREY, True, name="Shot caption")
pic = f.shapes.add_picture(shot, Inches(0.48), Inches(2.6), height=Inches(4.3))
pic.line.color.rgb = rgb(LGREY); pic.line.width = Pt(0.75); pic.name = "Job duties screenshot"
cols = [("① 現場共創", ORANGE, WHITE, "職責第 4 條：用 AI 開發工具做可運作原型，發掘需求", "我們的基礎", "國發專案已用互動原型，客戶當場刪減欄位", "整理成團隊標準流程，至少 2 個專案的工作坊使用"),
        ("② AI 驗收把關", MID, WHITE, "職責第 1 條：把模糊問題轉成可測試、可衡量的假設", "為什麼需要", "金管會 AI 指引要求「人在迴圈內」，由人監督並保留控制權", "做 BA 版 AI 產出審查清單，團隊先用再帶到客戶"),
        ("③ 領域脈絡", DARK, WHITE, "職責第 2 條：建立結構化業務脈絡，給專家、AI 與工程團隊使用", "為什麼需要", "Gartner 預測 2028 年企業的生成式 AI 模型，超過一半是領域專用", "挑一個專案，把需求格式、業務規則、名詞定義整理成知識檔")]
x0, cw, gap = 3.25, 3.1, 0.2
for i, (head, fill, fg, duty, lab2, why, step) in enumerate(cols):
    x = x0 + i * (cw + gap)
    box(f, x, 2.3, cw, 0.5, head, fill, fg, 15, True, name="Dir head " + head)
    for j, (lab, body) in enumerate([("對應職責", duty), (lab2, why), ("一年內第一步", step)]):
        y = 2.9 + j * 1.08
        box(f, x, y, cw, 0.3, lab, LGREY, TEXT, 10, True, name=f"Dir {i+1} lab {j+1}")
        box(f, x, y + 0.3, cw, 0.7, body, WHITE if j < 2 else CREAM, TEXT, 11, False, line=LGREY if j < 2 else LIGHT, align=PP_ALIGN.LEFT, name=f"Dir {i+1} body {j+1}")
box(f, x0, 6.2, 3 * cw + 2 * gap, 0.7, ["為什麼對任何領域的 BA 都重要", "寫需求變便宜了，價值移到現場驗證；客戶要的是看得到、能測試的成果；懂業務又能動手的人，最難被取代。"],
    PALE, TEXT, 12, True, align=PP_ALIGN.LEFT, name="Why it matters", sub_size=11)
st(f, 12, "資料來源：國際金融機構 Forward Deployed BA 職缺（2026-10，職責原文截圖）；金管會 AI 指引（2024-06）；Gartner（2025-10）")
num(f, 2, 8)
notes(f, ["【約 2 分鐘】上一頁的能力，全球市場已經在要了。Forward Deployed 可以拆成兩個字：Forward 是往前，走到第一線；Deployed 是派駐，嵌進客戶或業務團隊一起工作。這個說法原本用在工程師，現在全球金融業開始用在 BA。",
          "左邊是職缺的職責原文。我從裡面挑出三條，對應三個團隊方向。",
          "方向一，現場共創，對應第 4 條：用 AI 做可運作原型、發掘需求。國發專案已經在做，下一步是整理成團隊標準流程，一年內至少兩個專案的工作坊用上。",
          "方向二，AI 驗收把關，對應第 1 條：把模糊問題轉成可測試、可衡量的假設。金管會 AI 指引要求人在迴圈內；第一步是做一份 BA 版的 AI 產出審查清單。",
          "方向三，領域脈絡，對應第 2 條：建立結構化的業務脈絡給 AI 用。Gartner 預測到 2028 年，企業用的生成式 AI 模型超過一半是領域專用；第一步是挑一個專案，把需求格式、業務規則、名詞定義整理成知識檔。",
          "為什麼對任何領域的 BA 都重要？寫需求變便宜了，價值移到現場驗證；客戶要的是看得到、能測試的成果；懂業務又能動手的人，最難被取代。",
          "三個方向分別對應一年目標的價值、品質與效率。"])

# ================================================================ 9 References
rf = SL[8]
tbl = find(rf, 15).table
refs = [("#", "資料", "發布單位", "日期", "用在哪一頁"),
        ("1", "Forward Deployed Business Analyst 職缺", "國際金融機構", "2026-10", "8"),
        ("2", "金融業運用人工智慧（AI）指引", "金管會", "2024-06", "8"),
        ("3", "保險業運用人工智慧系統自律規範", "壽險／產險公會", "2025-04 備查", "7"),
        ("4", "Top Strategic Technology Trends for 2026", "Gartner", "2025-10", "8"),
        ("5", "intent（詞條）", "Cambridge Dictionary", "—", "5"),
        ("6", "Underwriting Operating System（UWOS）", "McKinsey & Company", "2026", "3、4"),
        ("7", "Different by Design: 2026 Global Insurance Report", "BlackRock", "2026-10", "3"),
        ("8", "Lesson 04 產品意圖卡（Intent Map）", "Shaun AI Lesson Learning", "—", "5、6"),
        ("9", "The AI-Native SDLC Playbook", "Anthropic", "—", "6")]
for i, row in enumerate(refs):
    for j, v in enumerate(row):
        set_tf(tbl.cell(i, j).text_frame, [(0, v)])
num(rf, 3, 9)

for s in SL:
    for e in s._element.iter(qn('p:extLst')):
        pass
    for e in [e for e in s._element.iter(qn('p:extLst')) if len(e) == 0]:
        e.getparent().remove(e)
p.save(out)
print("saved", out)
