# 製作分工：Gemini 生參考圖 → Seedance 生影片

| 階段 | 工具 | 提示詞 | 產出 |
|---|---|---|---|
| 1 角色設定圖 | 已完成 | `../character-sheet.txt` | `character.png` |
| 2 鏡頭圖片 | Gemini 圖片生成（上傳 `character.png`） | `1-keyframe1.txt`、`2-keyframe2.txt`、`3-keyframe3.txt` | `shot1–3.png` |
| 3 圖生影片 | Dreamina Seedance（首幀模式，5 秒，16:9） | `../i2v1.zh.txt`、`../i2v2.zh.txt`、`../i2v3.zh.txt` | 三段 5 秒影片 |
| 4 剪輯 | CapCut | 見 `../../storyboard.md` | 15 秒成品 |

臉跑掉時：Dreamina 改用全能參考模式，`@Image 1` 放鏡頭圖片（首幀）、`@Image 2` 放 `character.png`，並在 i2v 提示詞開頭加上：「@Image 2 只控制主角的臉、髮型和服裝，不使用它的背景與構圖。」
