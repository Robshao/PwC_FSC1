# 動態字卡產生器（Remotion）

後製用的動態字卡：章節卡、問題卡、服務卡、結尾卡。做法參考開源專案 OpenMontage 的 Remotion 合成方式，程式碼是我們自己寫的。不需要任何 API 金鑰，也不會把內容送到外部服務。

## 第一次使用

```bash
cd pwc-fsc1-ai-short-vid/tools/remotion-cards
npm install
```

需要 Node.js 18 以上，以及 Chromium（雲端環境已內建；在自己電腦上，Remotion 會自動下載）。

## 產生字卡

```bash
node render.mjs <字卡名稱> <內容檔.json> [輸出檔] [--overlay]
```

| 字卡名稱 | 用途 | 輸出 |
|---|---|---|
| `ChapterCard` | 章節字卡（深色底、置中大字） | `out/ChapterCard.mp4` |
| `QuestionCards` | 問題卡，依序浮現；**加 `--overlay` 輸出透明背景**，疊在影片上 | `out/QuestionCards.mov`（ProRes 4444） |
| `ServiceCard` | 服務對照表，「團隊展望」自動用不同顏色 | `out/ServiceCard.mp4` |
| `EndCard` | 結尾卡：主標語＋CTA＋署名＋小字 | `out/EndCard.mp4` |

內容寫在 `props/*.json`，改文字不用改程式。每個 JSON 都可以加 `"seconds": 數字` 決定片長。

範例（研討會短片）：

```bash
node render.mjs ChapterCard props/seminar-chapter.json
node render.mjs QuestionCards props/seminar-questions.json --overlay
node render.mjs ServiceCard props/seminar-services.json
node render.mjs EndCard props/seminar-end.json
```

### 問題卡的位置

`QuestionCards` 可以設定：
- `side`：`"left"` 或 `"right"`
- `vAlign`：`"top"`、`"center"`、`"bottom"`
- `compact`：`true` 時卡片小一號，避免擋到人臉

### 疊到影片上（ffmpeg）

```bash
ffmpeg -i s5b.mp4 -i out/QuestionCards.mov -filter_complex "[0][1]overlay=format=auto:enable='between(t,0,4)'" -c:a copy s5b-cards.mp4
```

## 預覽與調整

```bash
npm run studio
```

會開啟 Remotion Studio，可以在瀏覽器即時預覽、拖時間軸。

## 修改樣式

- 顏色與字型：`src/theme.ts`（與製作流程簡報同一組顏色）
- 各字卡版面：`src/ChapterCard.tsx`、`src/QuestionCards.tsx`、`src/ServiceCard.tsx`、`src/EndCard.tsx`
