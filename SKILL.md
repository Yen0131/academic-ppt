---
name: academic-ppt
description: "Use when asking to generate academic presentations or using /academic-ppt (keyword: academic-ppt). 學術簡報生成器 — 回答幾個問題，從無到有生成 PowerPoint/Keynote 相容的學術簡報
triggers:
  - /academic-ppt
  - skill_view(name='academic-ppt')
inputs: none
outputs: |
  在桌面產生 .pptx 檔案（格式：YYYY-MM-DD_主題.pptx）
---

# 學術簡報生成器

回答以下問題，AI 會在桌面產生一份完整的學術簡報（.pptx，16:9，適用於 PowerPoint / Keynote）。

## 問題清單

| # | 問題 | 說明 | 預設值 |
|---|------|------|--------|
| 1 | 簡報主題 | 一句話描述簡報標題 | — |
| 2 | 場合 | 研討會 / 課堂報告 / 論文答辯 / 其他 | 研討會 |
| 3 | 總時長 | 5 / 10 / 15 / 20 / 30 分鐘 | 15 分鐘 |
| 4 | 投影片數量 | 數字（6-30） | 10 頁 |
| 5 | 大綱 | 自己輸入章節，或說「由 AI 規劃」 | AI 規劃 |
| 6 | 特殊需求 | 圖表 / 資料視覺化 / Q&A 頁 / 附錄 | 無 |
| 7 | 配色 | 深藍（學術）/ 深綠 / 灰黑 / 自訂色票 | 深藍 |

## 生成結構邏輯

根據時長與頁數自動調整內容：
- 標題頁 → 大綱頁 → 研究背景 → 研究動機 → 方法 → 結果 → 討論 → 結論 → Q&A（可選）→ 參考文獻（可選）
- 頁數越多，各章節內容越豐富

## 配色方案

- **深藍（學術）**：主色 `#1F4E79`，背景 `#FFFFFF`，文字 `#1F4E79`
- **深綠**：主色 `#2E7D32`，背景 `#FFFFFF`，文字 `#2E7D32`
- **灰黑**：主色 `#333333`，背景 `#F5F5F5`，文字 `#333333`

## 字體與格式

- 尺寸：16:9（標準寬螢幕）
- 標題：32-36pt，深色
- 內文：18-24pt，深色
- 留白充足，符合學術簡報美學

## 使用方式

執行以下命令啟動互動式問答流程：

```bash
python3 ~/.hermes/skills/productivity/academic-ppt/scripts/generate_ppt.py
```

腳本會依序詢問問題，生成完畢後在桌面產生 `.pptx` 檔案。

## 安裝依賴

第一次執行時會自動安裝 `python-pptx`，如需手動安裝：

```bash
pip install python-pptx
```

## 跨設備同步

此 Skill 位於 `~/.hermes/skills/productivity/academic-ppt/`，
只要將整個 `.hermes` 資料夾同步到新設備（透過 iCloud、Git、RSync 等），
即可在新設備上使用。