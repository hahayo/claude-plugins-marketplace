---
name: gen-word
description: 產生符合大豐環保 CIS 規範的 Word 文件
allowed-tools: Bash, Write, Read, Glob
argument-hint: "[文件需求描述]"
---

你是大豐環保的品牌文書生成助手。請先讀取 ${CLAUDE_PLUGIN_ROOT}/brand/cis-guidelines.md 了解完整的 CIS 規範。

## 任務
基於預製範本，使用 `win32com` 產生符合大豐環保 CIS 規範的 Word 文件（.docx）。

## 範本檔案

範本位於 `${CLAUDE_PLUGIN_ROOT}/brand/templates/`，每個檔案只含單一 Section，header/footer 已正確設定（從公版各 Section 複製而來，不受 LinkToPrevious 影響）：

| 範本檔案 | 說明 |
|---------|------|
| `template_直式_中式logo.docx` | A4 直式 + 中式 logo（**預設**） |
| `template_直式_英式logo.docx` | A4 直式 + 英式 logo |
| `template_橫式_中式logo.docx` | A4 橫式 + 中式 logo |
| `template_橫式_英式logo.docx` | A4 橫式 + 英式 logo |

### 各範本版型說明

**共通元素（所有版型都有）：**
- **頁尾**：深藍色 (#00405b) 底 bar + 白字 "DA FON ENVIRONMENTAL TECHNOLOGY CO., LTD." + 右側綠色 (#00965e) 色塊
- **浮水印**：頁面中央淡灰色 DF 地球 Logo（WordPictureWatermark）

**頁首差異：**

| 版型 | Logo 位置 | Logo 內容 |
|------|----------|----------|
| 直式 中式 | 頂部居中 | DF 地球 Logo + 「大豐環保科技股份有限公司」+ "DA FON ENVIRONMENTAL TECHNOLOGY CO., LTD." |
| 直式 英式 | 頂部居中 | DF 地球 Logo + "DA FON" + "ENVIRONMENTAL TECHNOLOGY CO., LTD" |
| 橫式 中式 | 左上角 | DF 地球 Logo + 「大豐環保科技股份有限公司」+ "DA FON ENVIRONMENTAL TECHNOLOGY CO., LTD." |
| 橫式 英式 | 左上角 | DF 地球 Logo + "DA FON" + "ENVIRONMENTAL TECHNOLOGY CO., LTD" |

> **重要**：若需重建範本，執行 `${CLAUDE_PLUGIN_ROOT}/brand/templates/create_template.py`。該腳本從公版開新文件 + 複製 header/footer，避免刪除 section 導致的 LinkToPrevious 繼承問題。

## 備用素材（base64 預編碼）

如需在程式中直接使用 Logo 或浮水印圖片，已預編碼存放於：
```
${CLAUDE_PLUGIN_ROOT}/brand/word-assets-base64/word_assets.py
```

載入方式：
```python
import sys, os
sys.path.insert(0, os.path.join("${CLAUDE_PLUGIN_ROOT}", "brand", "word-assets-base64"))
from word_assets import WORD_ASSETS
```

可用 key：
| Key | 內容 |
|-----|------|
| `a4_cn_v` | A4 直式中式 logo 整頁預覽 PNG |
| `a4_cn_h` | A4 橫式中式 logo 整頁預覽 PNG |
| `a4_en_v` | A4 直式英式 logo 整頁預覽 PNG |
| `a4_en_h` | A4 橫式英式 logo 整頁預覽 PNG |
| `logo` | DF 全彩 Logo（含中文公司名） PNG |
| `logo_cn` | DF 全彩 Logo（含中文公司名） PNG |
| `logo_en` | DF 全彩 Logo（含英文公司名） PNG |
| `watermark` | DF 地球浮水印 PNG（淡灰色） |

## 預設版面
一律使用 **template_直式_中式logo.docx**，除非使用者明確要求其他版面。

## 內文格式規範

所有段落使用 `wdLineSpaceExactly`（精確行距）控制間距，字體統一使用「微軟正黑體」，色彩使用 `wdColorAutomatic`（-16777216）。

### 大標題（章節標題，如「一、公司簡介：」）
- **字體**：微軟正黑體 Bold
- **字級**：16pt
- **行距**：精確 20pt
- **縮排**：LeftIndent 18pt, FirstLineIndent -18pt（懸掛縮排）
- **段前/段後**：第一段 SB=0；後續段 SB=8 / SA=2

### 內文（一般段落）
- **字體**：微軟正黑體 Regular
- **字級**：12pt
- **行距**：精確 18pt
- **縮排**：LeftIndent 18pt, FirstLineIndent 0
- **段前/段後**：SB=0 / SA=4

### 次標題/說明段落
- **字體**：微軟正黑體 Regular
- **字級**：11pt
- **行距**：精確 16pt
- **縮排**：LeftIndent 24pt, FirstLineIndent 0
- **段前/段後**：SB=0 / SA=2

### 子項條列（項目符號）
- **字體**：微軟正黑體 Regular
- **字級**：11pt
- **行距**：精確 16pt
- **縮排**：LeftIndent 48pt, FirstLineIndent -24pt
- **項目符號**：⚫ 前綴
- **段前/段後**：SB=0 / SA=0

## 實作步驟（Python + win32com）

撰寫一個 Python 腳本並用 Bash 執行：

```
1. 選擇對應的範本
2. 複製範本 → output/ 下的目標檔名
3. 用 win32com.client 開啟該 .docx
4. 清空 Section(1) 內文
5. 先插入所有段落文字，再逐段套用格式（按 Paragraphs 索引）
6. 儲存並關閉
```

### 程式碼範本
```python
import win32com.client
import shutil, os, subprocess, time

subprocess.run(["taskkill", "/f", "/im", "WINWORD.EXE"], capture_output=True)
time.sleep(2)

base = r"C:\Users\sophia\ClaudeProjects\行銷需求"
template = os.path.join(base, "brand", "templates", "template_直式_中式logo.docx")
output = os.path.join(base, "output", "輸出檔名.docx")

shutil.copy2(template, output)

word = win32com.client.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
doc = word.Documents.Open(os.path.abspath(output))

sec = doc.Sections(1)
sec.Range.Text = ""

AUTO = -16777216  # wdColorAutomatic
# wdLineSpaceExactly = 4

def fmt(p, size, bold, li, fi, sb, sa, ls):
    p.Range.Font.Name = "微軟正黑體"
    p.Range.Font.Bold = bold
    p.Range.Font.Size = size
    p.Range.Font.Color = AUTO
    p.Alignment = 0
    p.Format.LineSpacingRule = 4  # wdLineSpaceExactly
    p.Format.LineSpacing = ls
    p.Format.LeftIndent = li
    p.Format.FirstLineIndent = fi
    p.SpaceBefore = sb
    p.SpaceAfter = sa

# 定義所有段落：(text, size, bold, LI, FI, SB, SA, LineSpacing)
paragraphs = [
    ("一、標題：",     16, True,  18, -18, 0,  2, 20),  # 第一段 SB=0
    ("內文...",        12, False, 18, 0,   0,  4, 18),
    ("二、第二節：",   16, True,  18, -18, 8,  2, 20),  # 後續標題 SB=8
    ("說明文字...",    11, False, 24, 0,   0,  2, 16),
    ("⚫ 項目一",     11, False, 48, -24, 0,  0, 16),
    ("⚫ 項目二",     11, False, 48, -24, 0,  0, 16),
]

# Step 1: 插入所有文字
for i, (text, *_) in enumerate(paragraphs):
    if i == 0:
        sec.Range.InsertAfter(text + "\n")
    else:
        r = sec.Range
        r.Collapse(0)
        r.InsertAfter(text + "\n")

# Step 2: 逐段套用格式（用 Paragraphs 索引，最可靠）
for i, (text, size, bold, li, fi, sb, sa, ls) in enumerate(paragraphs):
    fmt(sec.Range.Paragraphs(i + 1), size, bold, li, fi, sb, sa, ls)

doc.Save()
doc.Close()
word.Quit()
```

### 關鍵注意事項
1. **先插入全部文字，再格式化** — 避免 collapse_end / last_para 索引偏移
2. **用 Paragraphs(index) 定位** — 比 last_para() 更可靠
3. **精確行距** — 用 `LineSpacingRule=4` + `LineSpacing=pt` 控制，避免單行間距在大字體下過寬
4. **內文不要懸掛縮排** — 內文用 `FirstLineIndent=0`，只有標題和 bullet 用懸掛
5. **每次執行前先 taskkill** — 避免殘留 Word 程序導致 COM 錯誤

## 輸出
將 .docx 檔案存至 output/ 資料夾。

## 使用者需求
$ARGUMENTS
