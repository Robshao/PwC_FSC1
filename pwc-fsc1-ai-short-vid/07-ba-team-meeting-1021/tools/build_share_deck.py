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
o = p.slides[0]
set_text(o.shapes.title, ["為什麼現在是「動手做」的時代"])
set_text(find(o, 20), ["一生一次的機會", "Once-in-a-civilization"])
set_hb(find(o, 15), "智慧變便宜、門檻變低", "AI 讓分析與寫程式的能力變得便宜又容易取得，小團隊甚至一個人，也能和大公司競爭。")
set_hb(find(o, 17), "解決過去解不了的問題", "生物、醫療與基礎科學裡長期無解的難題，現在開始有機會被解決。")
set_hb(find(o, 16), "全新的商業與創作模式", "以前想像不到的產品、個人化的 AI 助理與商業模式，現在都做得出來。")
set_hb(find(o, 18), "引導風險，是 builder 的責任", "世界還沒準備好迎接快速發展的 AI，動手做的人有責任協助企業與政府調適，並處理資安等風險。")
set_text(find(o, 2), ["資料來源：Alexandr Wang 與 Garry Tan 對談〈This is a Once-in-a-Civilization Opportunity〉，Y Combinator，2026"])
set_num(o, 3, 1)
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
m = p.slides[1]
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
set_num(m, 3, 2)
set_notes(m, "\n".join([
    "【轉變心智圖，約 40 秒】",
    "左邊是我們過去的工作方式：蒐集需求交給技術團隊、文件靠人工整理、知識留在個人腦中。",
    "右邊是未來的 BA：現場用 AI 做原型和客戶共創、替 AI 產出把關、把業務規則整理成 AI 讀得懂的知識。",
    "中間六項能力，是我從 BlackRock 最新的 Forward Deployed BA 職缺整理出來的：動手做原型、產品排序思維、AI 驗收把關、業務脈絡建模、業務架構視野、溝通與共創。",
    "右邊這三個轉變，就是我想提的三個團隊方向，下一頁說明。",
]))
FOOTER = "FSC1 BA 組會｜團隊發展方向提案（資料來源：BlackRock 職缺 2026-10、金管會 AI 指引 2024-06、Gartner 2025-10）"
s1 = p.slides[2]
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
set_text(find(s1, 12), [FOOTER])
find(s1, 2).text_frame._txBody.find(qn('a:p')).find(qn('a:fld')).find(qn('a:t')).text = "3"
set_notes(s1, open(sys.argv[3], encoding='utf8').read())

p.save(out)
