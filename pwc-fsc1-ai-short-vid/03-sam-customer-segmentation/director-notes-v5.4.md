# 導演暨剪輯指導筆記｜客戶分層概念樣片 v5.4

> 對象：Sam 與管理層提案前的製作修正
> 依據版本：`customer-segmentation-v5-45s`（使用者標註的時間碼）→ 已修正為 `customer-segmentation-v5.4.mp4`
> 時間碼標示：**[v5]** 指使用者提到的舊版時間，**[v5.4]** 指修正後的新版時間

---

## 任務一：陽台 → 咖啡店轉場（已完成，做進 v5.4）

### 問題診斷 [v5] 00:20–00:22

| 問題 | 原因 |
|---|---|
| 動線斷裂 | 志明把手機拿開後，下一格直接是咖啡店業務員已經在說話，觀眾沒有時間理解「這是隔天的另一場見面」 |
| 音訊斷裂 | 電話裡業務員的尾音被硬切，緊接著咖啡店業務員本人開口，兩個聲音幾乎疊在一起 |
| 敘事缺緩衝 | 從「被打斷的電話」進入「錯誤示範」，中間沒有換場訊號 |

### 剪輯修剪秒數

| 片段 | 原始素材取用 | 處理 | [v5.4] 時間軸 |
|---|---|---|---|
| 陽台電話 | 素材 0.6–7.0 秒（剪掉開頭 0.6 秒空白） | 全程緩慢數位推近 1.00→1.06；最後 0.3 秒淡出到深藍；聲音最後 0.15 秒淡出 | 14.0–20.4 |
| 過渡字卡 | 0.8 秒 | 見下方版面 | 20.4–21.2 |
| 咖啡店 | shot_4.5 素材 0.3–6.6 秒 | 0.2 秒淡入；業務員在進場後 0.13 秒開口 | 21.2–27.5 |

為了維持 45 秒，倒帶字卡由 1.2 秒縮為 0.9 秒。

### 音訊橋接

| 時間 [v5.4] | 聲音層 | 手法 |
|---|---|---|
| 18.0–20.4 | 電話裡業務員的推銷聲 | 保留完整句子「⋯我明天帶過去給你看！」，不切斷 |
| 20.45–20.65、20.80–21.00 | **掛斷提示音**「嘟、嘟」兩聲 | 480 Hz＋620 Hz 雙頻音，經 300–3400 Hz 電話頻寬濾波，音量低於對白 |
| 20.4–21.2 | 咖啡店環境音（杯盤、街道機車聲） | **J-cut**：在字卡期間先 0.6 秒淡入，畫面還沒到、聲音先到 |
| 21.2 起 | 咖啡店現場音＋業務員對白 | 畫面淡入接上 |
| 全程 | 配樂 | 緊張段落的低音持續音延續到字卡下，不中斷，讓兩場戲在情緒上連成一段 |

### 0.8 秒過渡字卡

0.8 秒只夠讀 6–8 個字，所以只放「時間＋地點」和一個預告標籤：

```
┌───────────────────────────────────────────┐
│                                           │
│                                           │
│            隔天・巷口咖啡店                 │  ← 72px、粗體、白色、字距加寬、置中（y=505）
│                                           │
│             [ ✕ 錯誤示範 ]                 │  ← 32px、白字紅底標籤、置中（y=610）
│                                           │
│                                           │
└───────────────────────────────────────────┘
  底色：深海軍藍 #0B1A2E（與倒帶字卡、結尾 CTA 卡同色系）
  動態：文字 0.12 秒淡入、0.12 秒淡出
```

| 設計決定 | 理由 |
|---|---|
| 只寫「隔天・巷口咖啡店」 | 對應電話裡「我明天帶過去」，觀眾立刻知道這是約好的見面 |
| 預告紅色「錯誤示範」標籤 | 主管進入咖啡店前就知道接下來是反例，看的時候會帶著批判眼光 |
| 與倒帶字卡、CTA 卡同一個深藍底色 | 全片的「章節字卡」有一致的視覺語言 |

---

## 任務二：關鍵畫面重繪與 Inpainting 提示詞

### 2-1 桌上手部特寫（[v5] 00:36 ／ [v5.4] 35.5–37.5）

**實測結果**：素材第 1.0–1.5 秒兩人手指交疊處融成一團，看起來像多了手指；第 2.0–4.0 秒大致正常。
**v5.4 已先改用 2.0–4.0 秒的段落**，管理層提案可先用這版。若要徹底修正，建議流程：

1. 從素材擷取一張最清楚的畫格，在 Midjourney 用 **Vary (Region)／Editor** 框選志明的右手重繪（只動手，不動其他）。
2. 用修好的圖當首幀，丟回 Seedance／Runway 重新生成 4 秒動態。
3. 不建議直接在影片上逐格 inpaint：手指在動，每格修補結果不一致，會出現閃爍。

**Midjourney（局部重繪，框選志明右手）**

中文（理解用）：
```
寫實電影劇照，俯角特寫木頭茶几。一隻中年台灣男性的右手平放按在白色醫院信封上，五根手指分開清楚、解剖結構正確：一根拇指、四根手指，每根手指三個指節，指甲自然，手背有輕微青筋與細紋。淺藍色襯衫袖子捲到前臂，左手腕戴銀色智慧手錶，無名指戴素面銀色婚戒。一隻女性的手（米白色針織袖口、戴婚戒）輕輕覆蓋在他的手背上，兩人的手指不交纏、輪廓清楚分開。溫暖金黃色午後側光，淺景深。
```

英文（貼到 Midjourney）：
```
photorealistic cinematic still, overhead close-up of a wooden coffee table, a middle-aged Taiwanese man's right hand resting flat on a white hospital envelope, anatomically correct hand with exactly five fingers, one thumb and four fingers, three knuckles per finger, natural nails, subtle veins and skin texture, light blue dress shirt sleeve rolled to the forearm, silver smartwatch on the left wrist, plain silver wedding band on the ring finger, a woman's hand with a cream knit sleeve and wedding ring gently resting on top of the back of his hand, fingers clearly separated and not intertwined, warm golden afternoon side light, shallow depth of field, 85mm lens --ar 16:9 --style raw --no extra fingers, fused fingers, deformed hands, text, watermark
```

**Runway（影片編修／首幀生成動態，4 秒）**

中文：
```
以上傳的圖片作為第一幀。俯角固定特寫。女性的手輕輕握緊男性的手背，男性的手指慢慢放鬆，不再用力按著信封。右側一隻穿米色西裝袖口的手把彩色勾選表格擺正後收回。全程每隻手保持五根手指、指節結構正確，手指不融合、不增減。溫暖午後光線，鏡頭極輕微下推。
```

英文：
```
Use the uploaded image as the first frame. Locked overhead close-up. The woman's hand gently squeezes the back of the man's hand; his fingers slowly relax and stop pressing on the envelope. From the right, a hand in a beige blazer sleeve straightens the colorful checklist sheet and withdraws. Every hand keeps exactly five anatomically correct fingers throughout, no fused, missing or extra fingers. Warm afternoon light, very slight slow push-in. No text, no watermark.
```

### 2-2 林顧問平板 UI 替換（[v5] 00:34 ／ [v5.4] 32.6–35.5）

**實測結果**：平板在畫面中會跟著鏡頭移動，林顧問的手指也不時遮住螢幕。

**建議做法（比 AI 重繪可靠）**：螢幕畫面用後製替換，不用 AI 重繪。

| 做法 | 優點 | 缺點 |
|---|---|---|
| ✅ **後製螢幕替換**：做一張高解析度儀表板圖，用 After Effects（Mocha 平面追蹤）或 CapCut 的追蹤功能貼上螢幕，手指用遮罩摳出來 | 數字與文字百分之百清楚、可控、可以給法遵審 | 需要 1–2 小時的追蹤與遮罩 |
| AI 重繪螢幕（Runway 影片編修） | 快 | 文字與數字容易亂碼，每格不一致；金融數字不能出錯，不建議 |

**儀表板內容規格（示意數字，需由業務單位確認）**

> 關鍵洞察：志明已確診，**新保單幾乎無法通過核保**。所以專業顧問的儀表板，第一欄應該是「現有保單可理賠多少」，缺口的補強對象要轉向**太太雅婷與小宇**（PwC 工具 B：關係人保障缺口）。這一點正是錯誤示範（推新保單給志明）與正確示範的根本差異。

| 區塊 | 內容 |
|---|---|
| 標題列 | 家庭保障缺口分析｜志明一家（示意） |
| 第一欄：現有保單可理賠 | 重大傷病一次金 100 萬（可立即申請）／醫療實支實付：每次住院上限 10 萬 |
| 第二欄：三大缺口 | **壽險缺口**：需求 1,600 萬（房貸 800 萬＋子女教育＋5 年生活費）／已有 600 萬／缺口 1,000 萬<br>**重大傷病缺口**：需求 300 萬（3 年收入補償）／已有 100 萬／缺口 200 萬<br>**醫療實支實付缺口**：每次住院需求 30 萬／已有 10 萬／缺口 20 萬 |
| 第三欄：建議優先順序 | ① 申請現有理賠 ② 檢視配偶保障 ③ 子女教育金規劃 |
| 頁尾 | 數字為示意，非實際試算 |

**Midjourney（產出平面儀表板視覺，再交給設計排上正確文字）**

中文：
```
專業金融機構風格的平板電腦儀表板介面設計，正面平視、扁平 UI。深海軍藍標題列，白色背景。三欄式版面：左欄「現有保單可理賠」卡片；中欄三條橫向缺口長條圖，分別為壽險、重大傷病、醫療實支實付，每條以灰色表示已有保障、橘紅色表示缺口；右欄為編號優先順序清單。配色為海軍藍、青綠色、橘紅色點綴，留白充足，資訊層級清楚，像銀行理財顧問使用的系統。
```

英文：
```
professional financial advisory tablet dashboard UI, flat front-facing screen design, family protection gap analysis, navy header bar, white background, three-column layout, left card for existing claimable benefits, center three horizontal gap bars for life insurance, critical illness, and medical reimbursement, grey segment for existing coverage and coral-orange segment for the gap, right column numbered priority list, navy, teal and coral accent palette, generous whitespace, clean information hierarchy, bank-grade wealth advisor software, crisp vector UI --ar 4:3 --style raw --no photo, hands, people
```
注意：Midjourney 產生的中文字與數字通常不正確，**請把它當版型參考**，最後的文字與數字由設計在 Figma／PowerPoint 排上，再做螢幕替換。

**Runway（若仍要用 AI 換螢幕，限文字少的版本）**

英文：
```
Replace only the content on the tablet screen with a clean financial dashboard: navy header, three horizontal bar charts with grey existing coverage and coral-orange gaps, minimal text. Keep the woman's hand, the tablet frame, the couple and the room exactly unchanged. The screen content stays locked to the tablet surface as it moves. No garbled text.
```

---

## 任務三：台詞與配音（VO）情緒精修

### 3-1 林顧問 [v5] 00:30–00:35 ／ [v5.4] 28.9–37.1

| | 原台詞 | 精修台詞 |
|---|---|---|
| 鏡頭 4b（正面，有對嘴） | 志明哥，我們先不急著買新保單。 | **志明哥，這陣子辛苦了。新保單⋯我們先不談。** |
| 鏡頭 5（背對鏡頭，可直接換配音） | 先算清楚現有保單能賠多少，再補家裡的缺口。 | **先把你手上這幾張能賠的，一筆一筆幫你申請下來。家裡還缺的，我們再一起補。** |

**配音指導**

| 句子 | 語速 | 重音（粗體） | 停頓 | 情緒 |
|---|---|---|---|---|
| 志明哥，這陣子**辛苦**了。 | 慢，約 3.5 字／秒 | 辛苦 | 句尾停 0.4 秒 | 先接住情緒，不急著談專業；聲音放低、柔 |
| 新保單⋯我們**先不談**。 | 慢 | 先不談 | 「新保單」後停 0.3 秒 | 篤定、讓對方放心，像在幫他擋掉壓力 |
| 先把你手上這幾張**能賠的**，**一筆一筆**幫你申請下來。 | 中速，約 4.5 字／秒 | 能賠的、一筆一筆 | 逗號處停 0.2 秒 | 具體、有行動力，專業感來自「一筆一筆」的細節 |
| 家裡還缺的，我們**再一起**補。 | 放慢收尾 | 一起 | 「缺的」後停 0.2 秒 | 溫暖，「一起」表示陪伴，不是推銷 |

**落地方式**：鏡頭 5 林顧問背對鏡頭，**可以直接換真人配音**，不需重新生成；新句約 5 秒，與原句長度相近。鏡頭 4b 是正面對嘴：新台詞含停頓約 4.5 秒，比原鏡頭的 3.7 秒長 0.8 秒，需要用原圖在 Seedance 重新生成 5 秒版本。片長會變成約 45.8 秒，可把結尾 CTA 卡前的溶接縮短，或接受略超過 45 秒。

### 3-2 推銷業務員 [v5] 00:22–00:25 ／ [v5.4] 21.3–24.0

| | 台詞 |
|---|---|
| 原台詞 | 現在更要趕快加買這張高保費的！ |
| **精修台詞** | **哥，這張先送核保看看啦，早買早安心！** |

| 設計 | 說明 |
|---|---|
| 「哥」 | 台灣業務常見的拉近距離稱呼，親切裡帶著壓迫 |
| 「先送核保看看」 | 業界真實用語。對剛確診的人說這句，懂行的主管一聽就知道荒謬：明知核保很難過還硬送，只為業績 |
| 「早買早安心」 | 經典話術，用「安心」包裝銷售，和顧問的真安心形成對比 |
| 長度 | 16 字，約 2.7 秒，和原句長度一致 |

**配音指導**：語速快（約 6 字／秒），音調上揚、熱情過頭；「送」和「早買」加重；句尾「啦」拉長帶笑。

**落地方式**：咖啡店業務員是側面中景、嘴型可見。換配音會有輕微對不上，建議用 shot_4.5 的原圖重新生成 Seedance，或接受輕微不同步（側臉、距離較遠，觀眾不易察覺）。

---

## 任務四：管理層 1 頁 Executive Summary

完整版已發布為網頁（見對話中的連結）。架構：

1. **一句話結論**：同一位客戶，產品導向會推錯產品、流失信任；先分層、看懂需求，才能把顧問服務做成長期關係。
2. **商業戰略價值**：產品導向 vs. 客戶分層顧問模式的對照；預期影響以假說呈現，並列出試點應追蹤的 KPI（成約率、13／25 個月繼續率、退保率、NPS），不引用未經驗證的數字。
3. **AI 製程效益**：本專案實際數據（工作天數、生成次數、工具）＋與實拍的比較架構；實拍成本欄位需以製作公司報價填入。
4. **規模化複製**：角色設定圖、場景庫、提示詞結構、後製模板四類資產；衍生 30 歲新婚首購族、55 歲退休準備族的變數替換表。
5. **請管理層決定的事**：試點範圍、法遵審查、品牌露出、配樂授權。
