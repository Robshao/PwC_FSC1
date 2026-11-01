import sys, copy
from pptx import Presentation
from pptx.oxml.ns import qn

src, out = sys.argv[1], sys.argv[2]
p = Presentation(src)

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

def set_text(sh, paras):
    """Replace text, keeping the first paragraph's pPr and first run's rPr."""
    txBody = sh.text_frame._txBody
    ps = txBody.findall(qn('a:p'))
    proto = ps[0]
    for x in ps:
        txBody.remove(x)
    for text in paras:
        np_ = copy.deepcopy(proto)
        runs = np_.findall(qn('a:r'))
        for r in runs[1:]:
            np_.remove(r)
        for tag in ('a:br', 'a:fld'):
            for x in np_.findall(qn(tag)):
                np_.remove(x)
        if runs:
            runs[0].find(qn('a:t')).text = text
        else:
            r = np_.makeelement(qn('a:r'), {})
            t = r.makeelement(qn('a:t'), {}); t.text = text; r.append(t)
            end = np_.find(qn('a:endParaRPr'))
            if end is not None: end.addprevious(r)
            else: np_.append(r)
        txBody.append(np_)

def delete(slide, sid):
    el = find(slide, sid)._element
    el.getparent().remove(el)

def set_notes(slide, text):
    tf = slide.notes_slide.notes_text_frame
    tf.text = text


def set_hb(sh, header, body):
    """Header paragraph + body paragraph, each keeping its own template formatting."""
    txBody = sh.text_frame._txBody
    ps = txBody.findall(qn('a:p'))
    protos = [ps[0], ps[1]]
    for x in ps:
        txBody.remove(x)
    for proto, text in ((protos[0], header), (protos[1], body)):
        np_ = copy.deepcopy(proto)
        runs = np_.findall(qn('a:r'))
        for r in runs[1:]:
            np_.remove(r)
        for x in np_.findall(qn('a:br')):
            np_.remove(x)
        runs[0].find(qn('a:t')).text = text
        txBody.append(np_)

def set_num(slide, sid, n):
    fld = find(slide, sid).text_frame._txBody.find(qn('a:p')).find(qn('a:fld'))
    fld.find(qn('a:t')).text = str(n)

# ---------- Slide 1: opener (template "Four segments") ----------
o = p.slides[5]
set_text(o.shapes.title, ["為什麼現在是「動手做」的時代"])
set_text(find(o, 20), ["一生一次的機會", "Once-in-a-civilization"])
set_hb(find(o, 15), "智慧變便宜、門檻變低", "AI 讓分析與寫程式的能力變得便宜又容易取得，小團隊甚至一個人，也能和大公司競爭。")
set_hb(find(o, 17), "解決過去解不了的問題", "生物、醫療與基礎科學裡長期無解的難題，現在開始有機會被解決。")
set_hb(find(o, 16), "全新的商業與創作模式", "以前想像不到的產品、個人化的 AI 助理與商業模式，現在都做得出來。")
set_hb(find(o, 18), "引導風險，是 builder 的責任", "世界還沒準備好迎接快速發展的 AI，動手做的人有責任協助企業與政府調適，並處理資安等風險。")
set_text(find(o, 2), ["資料來源：Alexandr Wang 與 Garry Tan 對談〈This is a Once-in-a-Civilization Opportunity〉，Y Combinator，2026"])
set_num(o, 3, 6)
set_notes(o, "\n".join([
    "【開場，約 40 秒】",
    "Scale AI 創辦人 Alexandr Wang 今年在 Y Combinator 的對談裡說，現在是一生一次的機會。他講了四個理由：",
    "第一，AI 讓分析和寫程式的能力變得便宜，小團隊、甚至一個人也能做出以前要一整個團隊才做得出的東西。",
    "第二，過去解不了的科學與醫療問題，現在開始有機會解決。",
    "第三，全新的產品和商業模式變得可能。",
    "第四，世界還沒準備好，所以動手做的人也有責任幫企業管好風險。",
    "對我們 BA 來說，第一點代表我們也能自己動手做出東西；第四點代表客戶需要有人幫他們把關。這兩點就是接下來我想講的 BA 轉變。",
]))

# ---------- Slide 2: mind map (template hexagon "Process") ----------
m = p.slides[6]
set_text(m.shapes.title, ["BA 的轉變"])
set_text(find(m, 150), ["從寫需求", "到做出成果"])
set_hb(find(m, 118), "過去｜蒐集需求", "訪談後寫成需求文件，交給技術團隊，等開發完才看到成果。")
set_hb(find(m, 119), "過去｜人工整理文件", "會議紀錄、需求文件都靠人工撰寫與檢查，品質看個人經驗。")
set_hb(find(m, 120), "過去｜知識在個人腦中", "業務規則和專案經驗留在個人身上，換人就要從頭來。")
set_hb(find(m, 123), "未來｜現場共創", "用 AI 做出可操作的原型，和客戶當場驗證、當場修改。")
set_hb(find(m, 124), "未來｜AI 驗收把關", "定義 AI 產出的合格標準，設計哪一步由人做 go／no-go 決定。")
set_hb(find(m, 125), "未來｜領域脈絡", "把業務規則、名詞與例外整理成 AI 讀得懂的知識，讓經驗可以複用。")
labels = {146: "動手做原型", 143: "產品排序思維", 144: "AI 驗收把關", 147: "業務脈絡建模", 145: "業務架構視野", 142: "溝通與共創"}
for sid, text in labels.items():
    set_text(find(m, sid), [text])
set_text(find(m, 2), ["六項能力整理自 BlackRock「Forward Deployed Business Analyst」職缺說明（2026-10）"])
set_num(m, 3, 7)
set_notes(m, "\n".join([
    "【轉變心智圖，約 40 秒】",
    "左邊是我們過去的工作方式：蒐集需求交給技術團隊、文件靠人工整理、知識留在個人腦中。",
    "右邊是未來的 BA：現場用 AI 做原型和客戶共創、替 AI 產出把關、把業務規則整理成 AI 讀得懂的知識。",
    "中間六項能力，是我從 BlackRock 最新的 Forward Deployed BA 職缺整理出來的：動手做原型、產品排序思維、AI 驗收把關、業務脈絡建模、業務架構視野、溝通與共創。",
    "右邊這三個轉變，就是我想提的三個團隊方向，下一頁說明。",
]))

NOTES_DIR = sys.argv[3]
FOOT = "FSC1 BA 組會（2026/11/06）｜10/21 組會作業分享"

def set_tf(tf, items):
    """items: list of (proto_index, text). Rebuild paragraphs from prototypes in tf."""
    txBody = tf._txBody
    ps = txBody.findall(qn('a:p'))
    protos = list(ps)
    for x in ps:
        txBody.remove(x)
    for idx, text in items:
        np_ = copy.deepcopy(protos[idx])
        runs = np_.findall(qn('a:r'))
        for r in runs[1:]:
            np_.remove(r)
        for x in np_.findall(qn('a:br')):
            np_.remove(x)
        if runs:
            runs[0].find(qn('a:t')).text = text
        elif text:
            r = np_.makeelement(qn('a:r'), {}); t = r.makeelement(qn('a:t'), {}); t.text = text; r.append(t)
            end = np_.find(qn('a:endParaRPr'))
            (end.addprevious(r) if end is not None else np_.append(r))
        txBody.append(np_)

def cell(tbl_shape, r, c):
    return tbl_shape.table.cell(r, c).text_frame

# ---------- 1 Agenda ----------
a = p.slides[0]
set_text(a.shapes.title, ["議程"])
set_text(find(a, 6), ["10/21 組會後的作業：工作盤點、議題心得、文章分享，以及團隊發展方向提案"])
rows = [("工作角色盤點", "三條工作線與最新進度", "02"),
        ("10 月議題心得", "四個問題的反思", "03"),
        ("文章分享", "McKinsey UWOS 核保作業系統", "04"),
        ("團隊發展方向", "AI 時代 BA 的三個方向", "06"),
        ("需要的支援", "待確認事項與請求協助", "09"),
        ("附件", "參考資料", "10")]
tb = find(a, 14)
for i, row in enumerate(rows):
    for j, v in enumerate(row):
        set_tf(cell(tb, i, j), [(0, v)])
set_text(find(a, 2), [FOOT])
set_num(a, 3, 1)
set_notes(a, "今天分享四件事：先說明我目前的三條工作線，再分享 10 月議題的心得，接著是 Luke 指定的 McKinsey 文章，最後提出我對 BA 團隊發展方向的建議，以及需要大家協助的事項。")

# ---------- 2 Role inventory ----------
r2 = p.slides[1]
set_text(find(r2, 6), ["工作角色盤點：三條工作線與最新進度"])
lt = find(r2, 1382)
set_tf(cell(lt, 0, 0), [(0, "我的角色")])
set_tf(cell(lt, 1, 0), [(0, "負責 IC Care 商業模式設計與 AI 短影片製作，並協助把 IT 團隊的 AI 開發方式轉成業務端聽得懂的說法。")])
set_tf(cell(lt, 3, 0), [(0, "未來方向（1–3 年）"), (1, ""), (2, "# 會用 AI 動手做的 BA"), (3, "# 能分析需求，也能快速做出成果"), (4, "# 讓客戶看到可驗證的價值")])
mt = find(r2, 1391)
set_tf(cell(mt, 0, 0), [(0, "IC Care 商業模式")])
set_tf(cell(mt, 0, 1), [(0, "AI 短影片")])
set_tf(cell(mt, 0, 2), [(0, "AI 開發框架")])
set_tf(cell(mt, 1, 0), [(0, "完成標竿研究、人物誌與旅程圖"), (1, "對客簡報 v0.5（13 頁）"), (2, "11 個使用情境已在 Teams 討論"), (2, "下一步：整合 Luke 的補充")])
set_tf(cell(mt, 1, 1), [(0, "人物誌短片定稿，Sam 認可"), (1, "第一階段主題：核心系統 × AI 轉型"), (2, "已和 Rex、Kenny 對齊三段架構"), (2, "等待：Luke 的主題效益構想")])
set_tf(cell(mt, 1, 2), [(0, "32 頁白話版簡報與 3 頁摘要"), (1, "已補上法規頁與人的變化頁"), (2, "等待：IT 實績、期程與費用"), (2, "待確認：正式負責人")])
set_text(find(r2, 10), [FOOT])
set_num(r2, 2, 2)
set_notes(r2, "\n".join([
    "我目前有三條工作線。",
    "IC Care：從標竿研究做到人物誌和旅程圖，對客簡報做到 v0.5，11 個情境已經在 Teams 討論過。",
    "AI 短影片：人物誌短片已定稿；第一階段主題是核心系統在 AI 下的轉型，已和 Rex、Kenny 對齊三段架構，正在等 Luke 的效益構想。",
    "AI 開發框架：完成 32 頁白話版簡報和給 Sam 的 3 頁版，還在等 IT 團隊的實績和報價區間，正式負責人也待確認。",
    "未來方向：成為會用 AI 動手做的 BA，既能分析需求，也能快速做出成果。",
]))

# ---------- 3 Reflection: four questions ----------
q = p.slides[2]
set_text(q.shapes.title, ["10 月議題心得：四個問題的反思"])
cols = [(56, "聽懂指令", ["原本：自己整理主題，做完一版才發現主題不夠好", "改變：先做趨勢研究，和主管確認優先主題再動手", "下一步：動手前先確認給誰看、要達成什麼"]),
        (58, "持續對齊", ["原本：研究和主題先自己做完，再拿給主管看", "改變：寫腳本前先和 Rex、Kenny 對齊三段架構", "下一步：腳本階段就對齊，不等到成片"]),
        (48, "轉成業務語言", ["原本：簡報用縮寫和工程用語", "改變：站在客戶角度改寫，例如「金檢時拿得出紀錄」", "下一步：對外內容先問：非專業的人聽得懂嗎"]),
        (44, "成為有效成員", ["原本：專注做完自己的兩條工作線", "改變：主動補位，整理影片框架與組會筆記", "下一步：AI 產出一律人工核對；整理短影片 SOP"])]
from pptx.util import Inches
for sid, head, body in cols:
    sh = find(q, sid)
    set_tf(sh.text_frame, [(0, head)] + [(1, b) for b in body])
    sh.top = Inches(3.55); sh.height = Inches(3.3)
    if sid != 56:
        sh.left = sh.left + Inches(0.25); sh.width = sh.width - Inches(0.2)
set_text(find(q, 2), ["議題：AI 影片製作、AI 開發管理框架（10/16 進度會議）"])
set_num(q, 3, 3)
set_notes(q, "\n".join([
    "我參與的是 AI 影片製作和 AI 開發管理框架兩個議題，以下依「原本怎麼想、10/16 後改變了什麼、下一步」來分享。",
    "【草稿，請改成自己的話】",
    "聽懂指令：原本自己整理主題，做完一版才覺得主題不夠好；後來先做趨勢研究、和主管確認優先主題才動手。",
    "持續對齊：原本研究做完才拿給主管看；這次寫腳本前先和 Rex、Kenny 對齊三段架構。",
    "轉成業務語言：AI 開發框架簡報裡，我先假裝是客戶資訊長找缺口，再把我們的專業改寫成客戶得到的好處。",
    "成為有效成員：主動幫忙整理影片框架和組會筆記；之後 AI 產出一律人工核對，並把短影片做法整理成 SOP。",
]))

# ---------- 9 Support needed ----------
sp9 = p.slides[8]
set_text(sp9.shapes.title, ["需要的支援與待確認事項"])
set_tf(find(sp9, 6).text_frame, [(0, "三條工作線都有部分進度卡在外部輸入。"), (1, "右表是需要主管與相關同仁協助的事項，希望今天能確認負責人與時間。"), (2, "")])
t9 = find(sp9, 10)
data = [("事項", "需要誰", "為什麼", "時間"),
        ("AI 短影片主題效益構想", "Luke", "腳本要等效益構想才能開工", "盡快"),
        ("AI 開發框架正式負責人", "Luke", "確認我的角色與預期產出", "今天確認"),
        ("參加客戶端會議", "Luke", "了解實際業務情境", "今天討論"),
        ("IT 團隊實績與示範", "IT 團隊", "簡報需要實際成果佐證", "待定"),
        ("診斷與試點的期程、費用", "Sam／IT 團隊", "客戶需要知道採購核決層級", "待定"),
        ("法規頁內容確認", "所內法遵", "對外使用前確認", "對外前")]
for i, row in enumerate(data):
    for j, v in enumerate(row):
        set_tf(cell(t9, i, j), [(0, v)])
set_text(find(sp9, 2), [FOOT])
set_num(sp9, 3, 9)
set_notes(sp9, "最後是需要協助的事項。最急的是 Luke 的主題效益構想，影片腳本要等這個才能開工；另外想確認 AI 開發框架的正式負責人，以及我是否適合參加客戶端會議。")

# ---------- 10 References ----------
rf = p.slides[9]
set_text(rf.shapes.title, ["附件：參考資料"])
t10 = find(rf, 15)
refs = [("#", "資料", "發布單位", "日期", "用在哪一頁"),
        ("1", "Forward Deployed Business Analyst 職缺", "BlackRock", "2026-10", "6、7、8"),
        ("2", "金融業運用人工智慧（AI）指引", "金管會", "2024-06", "8"),
        ("3", "保險業運用人工智慧系統自律規範", "壽險／產險公會", "2025-04 備查", "8"),
        ("4", "Top Strategic Technology Trends for 2026", "Gartner", "2025-10", "8"),
        ("5", "This is a Once-in-a-Civilization Opportunity", "Y Combinator", "2026-07", "6"),
        ("6", "Underwriting Operating System（UWOS）", "McKinsey & Company", "2026", "4、5"),
        ("7", "AI 協助整理，數字與日期已人工查證", "—", "2026-11", "全份")]
for i, row in enumerate(refs):
    for j, v in enumerate(row):
        set_tf(cell(t10, i, j), [(0, v)])
set_text(find(rf, 2), [FOOT])
set_num(rf, 3, 10)
set_notes(rf, "參考資料都已經人工打開原文核對過；Alexandr Wang 的對談只確認過大意，沒有引用原話。")
UW_FOOTER = "FSC1 BA 組會｜McKinsey UWOS 文章分享（資料來源：McKinsey & Company, 2026）"

# ---------- Slide 1 (layout from template p.23) ----------
s1 = p.slides[3]
set_text(find(s1, 6), ["McKinsey UWOS：核保轉向「機器優先、人類治理」"])
set_text(find(s1, 9), ["商業險與特種險的核保流程，多半和十年前差不多。McKinsey 認為 Agentic AI 已能執行多步驟工作，保險公司可以把核保重新設計成 UWOS（Underwriting Operating System）：由一個智慧層統籌從受理到出單的流程，人只在需要判斷的地方介入。"])
set_text(find(s1, 17), ["Machine-first, human-governed: AI handles routine work end to end; underwriters focus on portfolio management and complex decisions"])
set_text(find(s1, 19), ["為什麼現在要改"])
set_text(find(s1, 4), ["流程停滯"])
set_text(find(s1, 5), ["營運壓力"])
set_text(find(s1, 7), ["技術轉折"])
set_text(find(s1, 8), ["送件在佇列中排隊", "人工摘事實、追缺漏", "一件要數天到數週"])
set_text(find(s1, 10), ["擴大規模不能加成本", "核保人才短缺", "舊系統、流程零碎"])
set_text(find(s1, 11), ["Agentic AI 跑多步驟", "核對證據、調出前例", "自動準備核保建議"])
set_text(find(s1, 20), ["三條車道：依案件動態分派"])
set_text(find(s1, 14), ["直通車道：簡單、可重複的案件"])
set_text(find(s1, 35), ["在風險胃納、授權與信心門檻內，系統完成受理到出單，人幾乎不介入。"])
set_text(find(s1, 15), ["輔助決策車道：中等複雜度"])
set_text(find(s1, 33), ["系統整理事實、標出不確定處並提出建議，人快速核准、修改或拒絕。"])
set_text(find(s1, 13), ["專家判斷車道：複雜或策略性"])
set_text(find(s1, 36), ["系統備好證據與選項，核保專家主導判斷與談判，時間從行政轉向決策。"])
set_text(find(s1, 21), ["人的角色改變，", "但沒有消失"])
set_text(find(s1, 37), ["核保主管：從處理個案到監督"])
set_text(find(s1, 42), ["形塑業務組合、確保核保品質，並調整風險胃納、定價與經紀人策略。"])
set_text(find(s1, 38), ["資深專家：成為例外處理者"])
set_text(find(s1, 41), ["專注真正需要判斷、談判與經驗的複雜或客製化案件。"])
set_text(find(s1, 39), ["機器：承擔大部分處理工作"])
set_text(find(s1, 40), ["六大領域、45 項能力由智慧層串起，決策更快、更一致，重工更少。"])
set_text(find(s1, 12), [UW_FOOTER])
find(s1, 2).text_frame._txBody.find(qn('a:p')).find(qn('a:fld')).find(qn('a:t')).text = "4"
set_notes(s1, "\n".join([
    "【一個重點】McKinsey 提出 UWOS（核保作業系統）：用 Agentic AI 把核保從逐案人工處理，改成機器優先、人類治理。",
    "為什麼現在：流程十年沒變、成本與人才壓力、Agentic AI 已能跑多步驟流程。",
    "怎麼運作：案件依複雜度、信心程度、所需授權層級分到三條車道，直通、輔助決策、專家判斷。",
    "人的角色：核保主管轉為監督，資深專家處理例外，機器承擔大部分處理。",
    "原文把 UWOS 拆成 6 大領域（受理、風險選擇、定價、報價出單、業務組合管理、營運績效管理）共 45 項能力。",
]))

# ---------- Slide 2 (layout from template p.114) ----------
s2 = p.slides[4]
for sid in (1049, 1051, 1052, 1053, 1056, 1057, 1074, 1082):   # hidden leftovers from the source deck
    delete(s2, sid)
set_text(find(s2, 4), ["從願景到落地：先選一個燈塔案例，再逐步擴大"])
set_text(find(s2, 49), ["McKinsey 建議不要等整個企業資料模型完美才開始，而是從高件量、資料齊全的流程切入，先定好控制機制，再擴大到更多領域。先行動的保險公司，將主導未來十年的速度、服務與核保紀律。"])
set_text(find(s2, 1077), ["一個重點"])
set_text(find(s2, 1076), ["UWOS 不是多加一個工具，而是重新分配人與機器的工作：例行交給系統，判斷留給人。", "原文拆成 6 大領域、45 項能力，由一個智慧層串起。"])
set_text(find(s2, 1078), ["一個啟發"])
set_text(find(s2, 1075), ["AI 導入的關鍵在分派規則與授權門檻：哪些自動通過、哪些交給人，要先講清楚並保留依據，客戶才會信任結果。"])
set_text(find(s2, 1062), ["從願景到落地的四個步驟："])
set_text(find(s2, 8), ["1 選一個", "燈塔案例"])
set_text(find(s2, 13), ["2 建立最低", "限度資料基礎"])
set_text(find(s2, 21), ["3 定義控制、", "授權與覆核"])
set_text(find(s2, 28), ["4 先代理化", "1–2 個流程"])
set_text(find(s2, 50), ["過去：逐案人工處理…"])
set_text(find(s2, 61), ["…將會變成…"])
set_text(find(s2, 51), ["UWOS：機器優先、人類治理"])
set_text(find(s2, 62), ["送件在佇列中排隊等人處理"])
set_text(find(s2, 1026), ["核保人員手動摘資料、追補缺件"])
set_text(find(s2, 1029), ["案件靠人工在各階段之間移動"])
set_text(find(s2, 1042), ["依複雜度與信心程度動態分派"])
set_text(find(s2, 1045), ["例行案件從受理到出單一路處理"])
set_text(find(s2, 1093), ["人專注判斷、談判與業務組合"])
set_text(find(s2, 1094), ["可應用做法：寫需求時，標出每一步是「系統自動」、「AI 輔助＋人確認」或「人判斷」"])
set_text(find(s2, 1095), ["AI 產出要可追溯：保留依據與前例，人工 review 才有基礎"])
set_text(find(s2, 1096), ["先挑一個高件量、資料齊全的流程試做，驗證後再擴大"])
set_text(find(s2, 2), [UW_FOOTER])
find(s2, 1092).text_frame._txBody.find(qn('a:p')).find(qn('a:fld')).find(qn('a:t')).text = "5"
set_notes(s2, "\n".join([
    "落地四步驟：",
    "1. 選燈塔案例：送件頻繁、資料齊全、處理週期或一致性很重要的險種、客群或續保業務。",
    "2. 最低資料基礎：只做第一批流程需要的文件、欄位、規則、定價工具與決策紀錄。",
    "3. 定義控制：信心門檻、授權規則、人工推翻流程，以及 human-in-the-loop 的角色。",
    "4. 啟動目標流程：先代理化 1–2 個高影響力領域（受理、風險選擇或定價），依商業價值、資料準備度與可行性排序。",
    "【一個啟發】AI 輔助能否被客戶接受，取決於分派規則、授權門檻與可追溯的依據，這和我們組會談到的人工 review、建立客戶信任是同一件事。",
    "【一個可應用做法】寫需求或設計流程時，標出每一步由系統自動、AI 輔助＋人確認，或由人判斷。",
]))

DIR_FOOTER = "FSC1 BA 組會｜團隊發展方向提案（資料來源：BlackRock 職缺 2026-10、金管會 AI 指引 2024-06、Gartner 2025-10）"
s1 = p.slides[7]
set_text(find(s1, 6), ["AI 時代的 BA：三個團隊發展方向"])
set_text(find(s1, 9), ["BlackRock 2026 年 10 月在新加坡開出「Forward Deployed Business Analyst」職缺，要求 BA 結合結構化分析、產品思維與動手做原型的能力。它列出的三項職責，正好對應我們可以發展的三個方向，也對應一年目標中的價值、品質與效率。"])
set_text(find(s1, 17), ["“This is not a traditional business analysis role centered on gathering requirements and handing them to technology.” — BlackRock, Forward Deployed BA job posting (2026)"])
set_text(find(s1, 19), ["現場共創：AI 原型確認需求"])
set_text(find(s1, 4), ["業界怎麼做"])
set_text(find(s1, 5), ["我們已經在做"])
set_text(find(s1, 7), ["一年內第一步"])
set_text(find(s1, 8), ["用 AI 做可操作原型", "發掘需求、挑戰假設", "BlackRock 職缺要求"])
set_text(find(s1, 10), ["國發專案互動原型", "客戶當場刪減欄位", "提早簡化流程"])
set_text(find(s1, 11), ["整理成團隊標準流程", "≥2 個專案工作坊使用", "衡量：抓到的需求落差"])
set_text(find(s1, 20), ["AI 驗收標準與人工把關"])
set_text(find(s1, 14), ["為什麼需要"])
set_text(find(s1, 35), ["客戶擔心 AI 產出能不能信；金管會 AI 指引要求「人在迴圈內」，由人監督並保留控制權。"])
set_text(find(s1, 15), ["BA 的優勢"])
set_text(find(s1, 33), ["寫驗收標準本來就是 BA 的工作，對象從系統功能延伸到 AI 產出。"])
set_text(find(s1, 13), ["一年內第一步"])
set_text(find(s1, 36), ["做 BA 版 AI 產出審查清單，團隊先用，再帶到客戶專案。"])
set_text(find(s1, 21), ["領域脈絡工程：", "讓 AI 讀懂業務規則"])
set_text(find(s1, 37), ["要解決的問題"])
set_text(find(s1, 42), ["AI 讀不到公司內部規格，無法判斷文件是否合規，問題在於缺少脈絡。"])
set_text(find(s1, 38), ["外部佐證"])
set_text(find(s1, 41), ["Gartner 預測 2028 年企業使用的生成式 AI 模型，超過一半是領域專用。"])
set_text(find(s1, 39), ["一年內第一步"])
set_text(find(s1, 40), ["挑一個專案，把需求格式、業務規則、名詞定義整理成知識檔。"])
set_text(find(s1, 12), [DIR_FOOTER])
find(s1, 2).text_frame._txBody.find(qn('a:p')).find(qn('a:fld')).find(qn('a:t')).text = "8"
set_notes(s1, open(NOTES_DIR, encoding='utf8').read())


p.save(out)
