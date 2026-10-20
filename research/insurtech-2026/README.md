# 全球與亞太保險科技新應用暨顧問服務機會研究（內部討論稿）

- `保險科技研究_內部討論稿.pptx`／`.pdf`：21 頁，所內簡報範本
- `build_deck.py`：產生簡報的程式。範本檔不放進版本控制（含其他專案頁面），重建時需指定範本路徑：
  `python3 build_deck.py <範本.pptx> 保險科技研究_內部討論稿.pptx`
- 內容來源：Gemini Deep Research 彙整之公開研究（2026-10）。所有數字與日期放進對外提案前須回原文核對；標註「單一來源」者待驗證。

## 會議決策簡報（顧問式故事線）

- `保險科技服務優先順序_會議簡報.pptx`／`.pdf`：24 頁，給同事與主管的服務優先順序討論會使用（SCQA 摘要、Issue Tree、2×2 矩陣、90 天計畫、決策頁）
- `build_meeting_deck.py`：`python3 build_meeting_deck.py <範本.pptx> 保險科技服務優先順序_會議簡報.pptx`
- `會議簡報_故事線.md`：問題定義、Issue Tree、假設與每頁論點式標題
- 方法來自 `.claude/skills/mckinsey-consultant`（MIT，來源見該資料夾 SOURCE.md）

## 主管三頁精簡版

- `保險科技服務優先順序_主管精簡版.pptx`／`.pdf`：3 頁，重點摘要、為什麼是這兩項、90 天計畫與決策；每頁底部來源可點擊
- `build_manager_brief.py`：`python3 build_manager_brief.py <範本.pptx> 保險科技服務優先順序_主管精簡版.pptx`
- `主管精簡版_來源連結.md`：每個連結的出處與查核結果（含日期更正）
- `check_fit.py`：檢查文字是否超出文字框或頁面
