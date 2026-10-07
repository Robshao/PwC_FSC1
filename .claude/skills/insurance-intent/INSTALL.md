# insurance-intent skill — 安裝說明

這是一個 Claude Code 的 Skill（技能套件），用於把保險客戶的業務問題轉成：
intent.md → 規格（spec）→ 使用假資料的可點擊原型 → 業務規則驗證（evals）→ 給 RD 的工程交接包。
讓業務分析師（BA）能在工程開發前先和客戶驗證需求。

適用範圍：壽險（Life）與產險（Non-life / P&C），涵蓋核保、承保發單、收費、保全／批改、續保、理賠、再保等環節。

## 檔案結構

```
insurance-intent/
└── SKILL.md    # 完整流程、intent.md 範本、原型與交接規範
```

## 相依技能

SKILL.md 內會引用 `ai-native-sdlc` 技能的流程（Plan → Design → Build → Test → Deploy → Maintain）。
沒有安裝 `ai-native-sdlc` 也能單獨使用，但搭配安裝效果較完整。

## 安裝（Claude Code）

把整個 `insurance-intent` 資料夾放到下列其中一個位置：

- 全域（所有專案可用）：`~/.claude/skills/insurance-intent/`
- 單一專案：`<專案根目錄>/.claude/skills/insurance-intent/`

放好後重新開啟 Claude Code，輸入 `/insurance-intent` 即可呼叫，
或直接描述一個保險客戶的問題（例如「理賠補件率 30%，想縮短結案天數」），Claude 會自動套用。

## 在其他 AI 工具使用

若環境不支援 Skill，可直接把 `SKILL.md` 的內容貼為系統提示或對話開頭。

## 使用前的守則（SKILL.md 內建）

- 不放任何真實客戶資料，一律用虛構公司與假資料。
- 遵守事務所與客戶的 AI 工具政策。
- 原型不得使用雇主或客戶的品牌，每一頁都標示「原型（Prototype）— 非正式系統」。
