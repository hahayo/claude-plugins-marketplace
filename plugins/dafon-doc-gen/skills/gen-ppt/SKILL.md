---
name: gen-ppt
description: 產生符合大豐環保 CIS 規範的 PowerPoint 簡報
allowed-tools: Bash, Write, Read, Glob
argument-hint: "[簡報需求描述]"
---

你是大豐環保的品牌簡報生成助手。請先讀取 ${CLAUDE_PLUGIN_ROOT}/brand/cis-guidelines.md 了解完整的 CIS 規範。

## 任務
使用 **HTML + python-pptx 混合方案** 產生符合大豐環保 CIS 規範的 PowerPoint 簡報（.pptx）。

## 混合方案策略

根據投影片內容複雜度，選擇不同的渲染方式：

### 方式 A：python-pptx 原生文字框（簡單文字頁面）
適用於：封面、章節頁、尾頁、純文字內容頁
- 使用公版 JPG 背景 + 疊加 python-pptx 文字框
- **優點**：文字可編輯
- **用於版型**：1, 2, 3, 4, 5, 6, 7, 8, 9, 18

### 方式 B：HTML→截圖→嵌入（複雜排版頁面）
適用於：圖表、表格、infographic、流程圖等複雜排版
- 用 HTML/CSS 精確排版 → playwright 截圖 1920x1080 PNG → 作為整頁圖片嵌入 PPT
- **優點**：CSS 排版精確度高，可用 flexbox/grid，支援圓餅圖、長條圖等
- **用於版型**：10, 11, 12, 13, 14, 15, 16, 17

## 公版投影片背景素材（base64 預編碼）

所有背景圖片和 Logo 已預先轉為 base64，存放於：
```
${CLAUDE_PLUGIN_ROOT}/brand/ppt-assets-base64/slide_assets.py
```

載入方式：
```python
import sys, os
sys.path.insert(0, os.path.join("${CLAUDE_PLUGIN_ROOT}", "brand", "ppt-assets-base64"))
from slide_assets import SLIDE_BACKGROUNDS, LOGOS
```

- `SLIDE_BACKGROUNDS[n]` — 投影片 n (1~18) 的 JPG base64 字串
- `LOGOS["logo_color_cn"]` — 彩色中文 Logo PNG base64
- `LOGOS["logo_color_en"]` — 彩色英文 Logo PNG base64
- `LOGOS["logo_white_cn"]` — 白色中文 Logo PNG base64
- `LOGOS["logo_white_en"]` — 白色英文 Logo PNG base64

### 投影片版型對照表

| 編號 | 用途 | 背景特徵 | 渲染方式 |
|------|------|----------|----------|
| 1 | **封面A** — 有公司建築照片 | 左藍漸層 + 右弧形照片 + 右下 DF Logo | A (原生) |
| 2 | **封面B** — 產品照片 | 左藍漸層 + 右弧形照片 + 右下 DF Logo | A (原生) |
| 3 | **全藍內頁** — 章節轉場頁 | 全版深藍漸層 + 底部藍色橫條 + 右下 DF Logo | A (原生) |
| 4 | **左藍右白 A** — 左半圓弧大 | 左藍大弧 + 綠色弧邊 + 右白 | A (原生) |
| 5 | **左藍右白 B** — 左半圓弧中 | 左藍中弧 + 綠色弧邊 + 右白 | A (原生) |
| 6 | **底部弧形** — 上白下藍綠弧 | 白底 + 底部藍綠弧線 + 右下 DF Logo | A (原生) |
| 7 | **底部色帶** — 純白+底部多色條 | 白底 + 底部4色帶 + 右下 DF Logo | A (原生) |
| 8 | **頂部色帶** — 純白+頂部多色條 | 白底 + 頂部左側4色帶 + 右下 DF Logo | A (原生) |
| 9 | **純白頁** — 極簡白底 | 全白 + 右下 DF Logo | A (原生) |
| 10 | **清單頁** — 數字條列 | 白底 + 5條深藍橫條 + 綠色數字 | B (HTML) |
| 11 | **流程圖頁** — Infographic Layout | 白底 + 4步驟箭頭流程 | B (HTML) |
| 12 | **表格+圓餅圖** — Table & Chart | 左藍漸層描述 + 表格 + 圓餅圖 | B (HTML) |
| 13 | **大表格頁** — Table Layout | 白底 + 大表格 + 雙欄說明 | B (HTML) |
| 14 | **圖片網格** — 4欄圖片+說明格 | 白底 + 4x2 網格色塊 | B (HTML) |
| 15 | **3欄圓形圖** — Infographic Layout | 白底 + 3個圓形資訊卡 | B (HTML) |
| 16 | **長條圖頁** — Chart Layout | 白底 + 雙欄橫條圖 | B (HTML) |
| 17 | **世界地圖頁** — Worldmap Infographic | 左藍+世界地圖+圓環圖 | B (HTML) |
| 18 | **尾頁** — 結尾感謝頁 | 左藍漸層 + 右弧形照片 + 左下 DF Logo | A (原生) |

## 方式 A：python-pptx 原生文字框

```python
import base64, io
from pptx import Presentation
from pptx.util import Inches, Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

prs = Presentation()
prs.slide_width = Inches(13.333)   # 16:9
prs.slide_height = Inches(7.5)

def add_bg(slide, slide_num):
    """將公版背景圖設為投影片全版底圖"""
    img_data = base64.b64decode(SLIDE_BACKGROUNDS[slide_num])
    img_stream = io.BytesIO(img_data)
    slide.shapes.add_picture(
        img_stream, Emu(0), Emu(0),
        prs.slide_width, prs.slide_height
    )

def add_textbox(slide, text, x, y, w, h, font_size=24, bold=False,
                color='FFFFFF', font_name='Noto Sans CJK TC', alignment=PP_ALIGN.LEFT):
    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = RGBColor.from_string(color)
    p.font.name = font_name
    p.alignment = alignment
    return txBox

# 範例：封面
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, 1)
add_textbox(slide, '簡報標題', 0.8, 2.0, 4, 1, font_size=44, bold=True, color='FFFFFF')
```

## 方式 B：HTML→截圖→嵌入 PPT

### 步驟概覽
1. 產生一個完整的 HTML 頁面，使用公版背景（base64 嵌入）+ CSS 排版內容
2. 用 playwright 截圖為 1920x1080 PNG
3. 將 PNG 嵌入 PPT 作為整頁圖片

### HTML 模板

```python
def render_html_slide(slide_num, content_html, output_png):
    """用 HTML 渲染複雜版面，截圖為 PNG"""
    bg_b64 = SLIDE_BACKGROUNDS[slide_num]

    html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;700&family=Inter:wght@300;400;700;900&display=swap');
  * {{ margin: 0; padding: 0; box-sizing: border-box; }}
  body {{
    width: 1920px; height: 1080px;
    font-family: 'Noto Sans TC', sans-serif;
    position: relative; overflow: hidden;
  }}
  .bg {{
    position: absolute; top: 0; left: 0;
    width: 100%; height: 100%;
    z-index: 0;
  }}
  .bg img {{ width: 100%; height: 100%; object-fit: cover; }}
  .content {{
    position: relative; z-index: 1;
    width: 100%; height: 100%;
    padding: 60px 80px;
  }}
  /* 品牌色 */
  .dafon-green {{ color: #00965e; }}
  .dafon-blue {{ color: #00405b; }}
  .dafon-white {{ color: #ffffff; }}
  /* 字體 */
  .title {{ font-size: 48px; font-weight: 700; }}
  .subtitle {{ font-size: 28px; font-weight: 400; }}
  .body-text {{ font-size: 20px; font-weight: 400; line-height: 1.6; }}
  .number {{ font-family: 'Inter', sans-serif; font-weight: 900; font-size: 64px; }}
</style>
</head><body>
<div class="bg"><img src="data:image/jpeg;base64,{{bg_b64}}"></div>
<div class="content">
  {content_html}
</div>
</body></html>"""

    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={{'width': 1920, 'height': 1080}})
        page.set_content(html)
        page.wait_for_load_state('networkidle')
        page.screenshot(path=output_png, full_page=False)
        browser.close()
```

### 將截圖嵌入 PPT

```python
def add_html_slide(prs, png_path):
    """將 HTML 截圖作為整頁圖片加入 PPT"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_picture(
        png_path, Emu(0), Emu(0),
        prs.slide_width, prs.slide_height
    )
    return slide
```

### HTML 版面範例

#### 清單頁（版型 10）
```html
<div style="display:flex; flex-direction:column; gap:20px; padding-top:80px;">
  <div style="display:flex; align-items:center; gap:20px;">
    <span class="number dafon-green">1</span>
    <div style="background:#00405b; color:white; padding:16px 40px; flex:1; font-size:22px;">
      清單項目內容
    </div>
  </div>
  <!-- 重複 2~5... -->
</div>
```

#### 表格頁（版型 12/13）
```html
<h2 class="title dafon-blue" style="margin-bottom:30px;">Table & Chart</h2>
<div style="display:flex; gap:40px;">
  <table style="border-collapse:collapse; flex:1;">
    <thead><tr style="background:#00405b; color:white;">
      <th style="padding:12px;">項目</th><th>數值</th>
    </tr></thead>
    <tbody>
      <tr style="background:#e8f4f0;"><td style="padding:10px;">項目A</td><td>100</td></tr>
    </tbody>
  </table>
  <!-- 圓餅圖用 CSS conic-gradient -->
  <div style="width:300px; height:300px; border-radius:50%;
    background: conic-gradient(#00965e 0% 70%, #00405b 70% 90%, #0977d1 90% 100%);">
  </div>
</div>
```

#### 長條圖頁（版型 16）
```html
<h2 class="title dafon-blue">Chart Layout</h2>
<div style="display:flex; gap:60px; margin-top:40px;">
  <div style="flex:1;">
    <div style="display:flex; align-items:center; gap:10px; margin:12px 0;">
      <span style="width:140px; text-align:right; font-weight:bold;">項目 A</span>
      <div style="background:#00405b; height:30px; width:65%;"></div>
      <span>65%</span>
    </div>
    <!-- 更多橫條... -->
  </div>
</div>
```

## 品牌規則（必須遵守）

- **中文字體**：Noto Sans CJK TC（大標 Bold，內文 Regular）— HTML 用 Google Fonts 載入
- **英文字體**：Inter（大標 Bold/Black，內文 Regular）— HTML 用 Google Fonts 載入
- **主色**：Da Fon Green #00965e、Da Fon Blue #00405b
- **輔助色**：依 CIS 規範搭配
- **背景**：必須使用公版背景圖，不可自行繪製替代
- **投影片尺寸**：16:9（python-pptx: 13.333" x 7.5"，HTML: 1920x1080px）

### 各版型文字安全區域建議

| 版型 | 文字區域 |
|------|----------|
| 封面 (1,2) | 左側藍色區域放標題（x: 0.5"~4.5", y: 1.5"~5.5"） |
| 全藍內頁 (3) | 中央偏上（x: 1"~11", y: 1"~5.5"） |
| 左藍右白 (4,5) | 左藍區放標題，右白區放內容 |
| 白底頁 (6-9) | 大部分區域可用（x: 0.8"~12", y: 0.5"~6"），避開底部 Logo |
| 清單頁 (10) | HTML：跟隨藍色橫條位置（padding-top: 80px） |
| 圖表頁 (12,13,16) | HTML：標題頂部 60px，內容 flex 佈局 |
| 尾頁 (18) | 左側藍色區域放感謝文字 |

## 建議的投影片結構

一份典型簡報應包含：
1. **封面** → 版型 1 或 2（方式 A）
2. **目錄/章節頁** → 版型 3（方式 A）
3. **內容頁** → 版型 4~9 文字為主（方式 A）、10~17 圖表數據（方式 B）
4. **尾頁** → 版型 18（方式 A）

## 完整流程

```python
# 1. 建立 Presentation
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# 2. 方式 A 頁面（封面、章節、文字頁、尾頁）
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, 1)
add_textbox(slide, '標題', ...)

# 3. 方式 B 頁面（圖表、infographic）
render_html_slide(12, '<h2>Table</h2>...', 'output/tmp_slide_3.png')
add_html_slide(prs, 'output/tmp_slide_3.png')

# 4. 儲存
prs.save('output/簡報名稱.pptx')

# 5. 清理暫存截圖
import glob
for f in glob.glob('output/tmp_slide_*.png'):
    os.remove(f)
```

## 輸出
將 .pptx 檔案存至 output/ 資料夾。

## 使用者需求
$ARGUMENTS
