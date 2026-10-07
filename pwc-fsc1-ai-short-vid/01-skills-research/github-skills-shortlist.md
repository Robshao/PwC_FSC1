# GitHub AI 影片技能（Skills）評估

調查日期：2026-10-07。星數與狀態以當日 GitHub 頁面為準。

## 篩選原則

1. **純提示詞型優先**：只有 `SKILL.md` 與參考文件、不呼叫外部付費 API、不上傳資料的技能，最符合「所內 AI 影片製作底線」（步驟 1）。
2. **對應 Dreamina / Seedance**：我們實際產片的工具是 Dreamina，提示詞要貼合它的模型與限制。
3. **非專業人員也能用**：能把模糊想法轉成分鏡與提示詞，降低每天 3 小時投入的負擔。
4. **授權**：MIT 等寬鬆授權優先；AGPL 之類的授權在顧問公司環境要先確認。

## 對應九步驟的推薦清單

| 步驟 | 推薦技能 | 用途 | 星數 | 授權 | 類型 | 評價 |
|---|---|---|---|---|---|---|
| 3 腳本設計 | [aicontentskills/ai-video-storyboard-skill](https://github.com/aicontentskills/ai-video-storyboard-skill) | 把簡短需求變成 6–18 鏡頭的分鏡、視覺一致性指引、後製清單 | 60 | MIT | 純提示詞（單一 SKILL.md） | ⭐ **首選**：最輕量，也可以直接貼進 ChatGPT 當系統指令 |
| 3 腳本設計 | [yipingheijiang/DirectorSKILL](https://github.com/yipingheijiang/DirectorSKILL) | 節拍表、鏡位、關鍵影格提示詞、剪輯時間軸、品質修正；附中英雙語模板 | 0 | MIT | 純提示詞 | 功能完整、支援中文；但很新、無社群驗證，適合進階使用 |
| 4 提示詞產出 | [sjinn-ai/seedance2.5-skills](https://github.com/sjinn-ai/seedance2.5-skills) | 依 Dreamina Seedance 2.5 的實測規格產生提示詞（4–30 秒、480p/720p、參考素材上限） | — | 未標示 | 純提示詞 | ⭐ **首選**：規格最貼近 Dreamina；授權未標示，限內部參考 |
| 4 提示詞產出 | [silentbuilds/seedance-prompt-forge](https://github.com/silentbuilds/seedance-prompt-forge) | 撰寫、檢查、修復 Seedance 2.5 提示詞；附 Python 檢查器 | 2 | MIT | 提示詞 + 本機腳本 | 能在生成前先檢查提示詞，省下 Dreamina 點數 |
| 4 提示詞產出 | [gbeyrouti/seedance-prompting-claude-skill](https://github.com/gbeyrouti/seedance-prompting-claude-skill) | 六段式公式：主體＋動作＋場景＋風格＋鏡頭＋音訊 | 6 | MIT | 純提示詞 | 公式好記；說明文件是法文，需要翻譯 |
| 4 提示詞產出 | [Square-Zero-Labs/video-prompting-skill](https://github.com/Square-Zero-Labs/video-prompting-skill) | 文生影片、圖生影片提示詞，以及角色設定圖提示詞 | — | — | 純提示詞 | 角色設定圖對「參考圖（ref）」很有用 |

## 暫不建議（列出供參考）

| 技能 | 原因 |
|---|---|
| [0xadvait/ai-video-skill](https://github.com/0xadvait/ai-video-skill) | 需要 Replicate、FAL、Gemini、ElevenLabs 等付費 API 金鑰，資料會送到外部服務，與 Dreamina 流程重複 |
| [op7418/guizang-product-video-skill](https://github.com/op7418/guizang-product-video-skill) | 專做軟體產品展示動畫（程式碼渲染），不是生成式敘事影片；AGPL-3.0 授權 |
| [OSideMedia/higgsfield-ai-prompt-skill](https://github.com/OSideMedia/higgsfield-ai-prompt-skill) | 綁定 Higgsfield 平台，不是 Dreamina |

## 延伸索引

- [zhuyansen/awesome-claude-video-skills](https://github.com/zhuyansen/awesome-claude-video-skills)：180 個影片相關技能，按類型分類並附安全評級，有中英文版本。之後需要剪輯、字幕類工具時可以從這裡找。

## 建議做法

1. **10/8 前**：只用兩個首選技能——`ai-video-storyboard-skill`（分鏡）和 `seedance2.5-skills`（Dreamina 提示詞），先跑完一次流程。
2. **安裝前**：每個技能的 `SKILL.md` 都要完整讀過，確認沒有對外連線、沒有執行未知腳本。
3. **輸入資料**：不要把客戶資料或機密內容輸入任何外部 AI 工具，個人年度目標也只放可以公開的內容。
