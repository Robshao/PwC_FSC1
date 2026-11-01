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


FOOTER = "FSC1 BA 組會｜團隊發展方向提案（資料來源：BlackRock 職缺 2026-10、金管會 AI 指引 2024-06、Gartner 2025-10）"
s1 = p.slides[0]
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
find(s1, 2).text_frame._txBody.find(qn('a:p')).find(qn('a:fld')).find(qn('a:t')).text = "1"
set_notes(s1, open(sys.argv[3], encoding='utf8').read())
p.save(out)
