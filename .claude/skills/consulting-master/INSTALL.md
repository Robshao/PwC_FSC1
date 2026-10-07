# consulting-master skill — 安裝說明

這是一個 Claude Code 的 Skill（技能套件），用於以管理顧問的方式拆解模糊的商業問題：
框定決策、驗證指標、建立 MECE 結構、定位問題、寫可證偽的假設、排序、規劃分析、找根因並提出建議。

## 檔案結構

```
consulting-master/
├── SKILL.md                         # 主流程（Claude 讀這個檔觸發技能）
└── references/
    ├── first-cut-identities.md      # 各產業常用的指標拆解恆等式
    ├── root-cause.md                # 根因分析方法
    ├── templates.md                 # 交付物範本與評分表
    └── worked-example.md            # 完整示範案例
```

## 安裝（Claude Code）

把整個 `consulting-master` 資料夾放到下列其中一個位置：

- 全域（所有專案可用）：`~/.claude/skills/consulting-master/`
- 單一專案：`<專案根目錄>/.claude/skills/consulting-master/`

放好後重新開啟 Claude Code，輸入 `/consulting-master` 即可呼叫，
或直接描述一個商業問題（例如「業績掉 20%，老闆說是行銷的問題」），Claude 會自動套用。

## 在其他 AI 工具使用

若環境不支援 Skill，可直接把 `SKILL.md` 的內容貼為系統提示或對話開頭，
再視需要貼上 `references/` 裡的檔案。

## 四種模式

| 模式 | 用途 |
|---|---|
| Quick | 在對話中快速給出結構化看法 |
| Full | 產出完整交付物（文件、工作表、案例報告） |
| Coach | 練習模式，先讓使用者作答再評分 |
| Review | 檢查既有的問題結構或分析計畫 |
