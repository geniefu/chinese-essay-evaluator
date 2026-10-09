# 中文作文與國中會考寫作測驗專業評審・教學診斷技能 (chinese-essay-evaluator)

本技能 (Skill) 專門用於指導 Google Antigravity 等 AI 代理人依據台灣**「國中教育會考寫作測驗標準評分規準」**與**「第一線資深國文名師閱卷實戰指標」**，對各類國中與國小引導式寫作、命題作文進行嚴格客觀的級分判定（0～6 級分）、篇幅行數檢核、天頭錯別字挑錯、寫作進階鷹架、逐段眉批與段落升級改寫示範。

支援單篇作文深度評閱與全班手寫掃描檔（PDF/圖片）批次自動批改，並一鍵產出具備「封面總覽表＋強制獨立分頁」之標準 **A4 作文批閱診斷手冊 (`.docx`)**。

---

## 🏛️ 專案模組架構 (Architecture)

本技能遵循漸進揭示 (Progressive Disclosure) 設計規範，採用三層式架構：

```text
chinese-essay-evaluator/
├── SKILL.md                              # 核心引導文檔：提詞模板、雙軌學段、評閱工作流、自檢清單
├── README.md                             # 專案介紹、安裝指南、提詞範例與使用說明
├── requirements.txt                      # 相依套件清單 (python-docx, PyMuPDF)
├── LICENSE                               # MIT 開源授權條款
├── scripts/                              # 自動化輔助工具庫
│   ├── essay_report_builder.py           # Word 批閱單/全班手冊自動生成引擎 (含天頭訂正格、強制分頁、鷹架卡片)
│   └── pdf_essay_extractor.py            # 手寫作文掃描 PDF/圖片批次頁面萃取預處理工具
└── references/                           # 核心評分法規與批閱寶典庫
    ├── cap_essay_rubric.md               # 心測中心官方寫作測驗四大向度 0~6 級分標準規準手冊
    ├── teacher_grading_secrets.md        # 現場名師批閱心法：篇幅防守線、一例一段、首尾呼應、由物入理
    └── sample_prompts_and_upgrades.md    # 歷屆會考經典考題排他性審題拆解與段落升級示範
```

---

## 📦 如何安裝本技能

### 方式 A：全域安裝（推薦，所有專案與新對話皆可隨時調用）

#### 1. 在 Antigravity 對話框貼上一句話自動安裝：
```text
請幫我把這個 GitHub 倉庫安裝成全域 Skill：
https://github.com/geniefu/chinese-essay-evaluator
並幫我檢查安裝所需的 Python 套件 (pip install -r requirements.txt)
```

#### 2. 透過 Git 指令一鍵安裝：
- **Windows (PowerShell)**:
  ```powershell
  git clone https://github.com/geniefu/chinese-essay-evaluator.git "$HOME\.gemini\config\skills\chinese-essay-evaluator"
  ```
- **macOS / Linux**:
  ```bash
  git clone https://github.com/geniefu/chinese-essay-evaluator.git ~/.gemini/config/skills/chinese-essay-evaluator
  ```

---

## 🚀 快速開始：提詞範例 (Prompt)

在對話框中輸入 `/chinese-essay-evaluator` 並提供寫作材料：

```text
請使用 /chinese-essay-evaluator 技能為我進行作文專業批閱與診斷：
1、學段與年級：【國中九年級 或 國小六年級】
2、寫作題目：〈當我停下來〉
3、引導材料與限定條件：【以身邊天天見得到的人事物為觀察核心，不可寫罕見出遊】
4、學生作文全文（或提供手寫掃描 PDF 檔案路徑）：
【在此貼上學生文章原文，或提供 PDF 檔案】
5、卷面篇幅預估：【單頁寫了約 18 行（佔比約 85%）】
6、輸出格式：【請在對話中輸出完整專業評閱報告，並產出 Word (.docx) 診斷報告供列印發放】
```

---

## 💻 自動產生 Word 批閱單與全班手冊 (CLI Usage)

```bash
# 1. 產出單生標準診斷單
python scripts/essay_report_builder.py --topic "當我停下來" --student "901班 05號 李同學" --score 4 --output "李同學作文診斷單.docx"

# 2. 傳入 JSON 資料生成全班彙整手冊（含封面總覽表、每生強制獨立分頁、全班教學診斷專頁）
python scripts/essay_report_builder.py --data-json "class_results.json" --output "全班寫作評閱手冊.docx"

# 3. 處理學生手寫掃描檔 PDF
python scripts/pdf_essay_extractor.py "學生手寫稿卷.pdf" --output-dir "scanned_pages"
```

---

## ✨ 核心特色與升級功能

1. **嚴格執行考場篇幅行數防守線**：
   - 國中教育會考稿紙：$<75\%$ 行數上限鎖定 $\le 3$ 級分；$<90\%$ 行數上限鎖定 $\le 4$ 級分；$\ge 90\%$ 方具備 5~6 級分（上卷）評定資格。
2. **字數不足強制啟動「寫作進階鷹架」**：
   - 當字數 $<500$ 字時，主動給予「細節擴寫提問」、「五感描摹錦囊」與「邏輯/情感深度反思」，給予具體寫作支撐。
3. **800~1200 字完整版昇華範文**：
   - 以學生原創靈光為骨幹，重構成 6 級分（A++）完整範文，開頭、轉折、收束絕不節錄縮水。
4. **手寫掃描檔批次自動連續管線**：
   - 支援整份多頁手寫 PDF 一鍵批改到底，自動統計全班成績矩陣。
5. **Word 排版極致規範（每生強制換頁）**：
   - 全班報告具備「封面成績矩陣表 ➔ 強制換頁 ➔ 學生個人專章（每生結束後強制換頁） ➔ 強制換頁 ➔ 全班整體教學診斷與研討活動」，可直接印製發放。
6. **考場天頭錯別字回寫訂正格**：
   - 實體還原考場批改實務，以 `「錯字」口 ➔ 【 正字 】` 呈現，輔導學生落實文字正確性。

---

## 🛠️ 環境相依需求 (Dependencies)

```bash
pip install -r requirements.txt
```
包含：`python-docx` 與 `PyMuPDF`。
