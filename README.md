# Academic PPT - 學術簡報生成器

從無到有生成學術簡報（.pptx），回答幾個問題即可，適用於 PowerPoint / Keynote。

---

## 安裝方式（選一種）

### 方式一：懶人指令（建議，macOS / Linux 通用）

```bash
# 下載並執行，自動安裝到 ~/.hermes/skills/
curl -fsSL https://raw.githubusercontent.com/Yen032/academic-ppt/main/install.sh | bash -s -- -y
```

> 如果還沒有建 GitHub repo，先手動下載 zip 再用方式二。

### 方式二：下載 zip 後執行 script

1. 下載 zip 並解壓
2. 終端機進入目錄，執行：
```bash
chmod +x install.sh
./install.sh -y
```

---

## 使用方式

```bash
python3 ~/.hermes/skills/productivity/academic-ppt/scripts/generate_ppt.py
```

腳本會依序詢問 7 個問題，生成完畢後在桌面輸出 `.pptx` 檔案。

---

## 功能特色

- **7 個互動式問題**：主題、場合、時長、頁數、大綱、特殊需求、配色
- **大綱二次確認**：AI 生成大綱後需確認才會製作
- **3 種配色**：深藍（學術）、深綠、灰黑
- **自動章節對應**：根據頁數自動調整內容結構
- **跨平台支援**：Mac / Windows 皆可（自動偵測桌面路徑）
- **附錄支援**：Q&A 頁、參考文獻頁（可選）
- **學習功能**：使用後評分，長期越用越順

---

## 輸出格式

- 副檔名：`.pptx`
- 比例：16:9（寬螢幕）
- 標題字：32-36pt
- 內文字：18-24pt
- 檔名：`YYYY-MM-DD_主題.pptx`

---

## 自行上架到 GitHub（首次設定）

```bash
cd academic-ppt
gh repo create academic-ppt --public --push
git push -u origin main
```

之後同學只要：
```bash
curl -fsSL https://raw.githubusercontent.com/Yen032/academic-ppt/main/install.sh | bash -s -- -y
```

就能完成安裝。