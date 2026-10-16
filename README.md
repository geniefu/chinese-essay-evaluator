# 中文作文與國中會考寫作測驗專業評審・教學診斷技能 (chinese-essay-evaluator)

本技能 (Skill) 專門用於指導 Google Antigravity 等 AI 代理人依據台灣**「國中教育會考寫作測驗標準評分規準」**與**「第一線資深國文名師閱卷實戰指標」**，對各類國中與國小引導式寫作、命題作文進行嚴格客觀的級分判定（0～6 級分）、篇幅行數檢核、天頭錯別字挑錯、寫作進階鷹架、逐段眉批與段落升級改寫示範。

支援單篇作文深度評閱與全班手寫掃描檔（PDF/圖片）批次自動批改，並一鍵產出具備「封面總覽表＋強制獨立分頁」之標準 **A4 作文批閱診斷手冊 (`.docx`)**。

特別支援**「兩階段人機協同自動化工作流 (HITL)」**：
1. **階段一（命題與規畫）**：指定題目與引導說明，由 AI 自動產出具備排他性審題、0~6級分矩陣與防守線之《嚴格評審準則規畫書》Word 檔（`.docx`）。
2. **人機確認節點 (HITL)**：由老師審查、確認或直接微調規畫書。
3. **階段二（批量批閱與同目錄存檔）**：指定規畫書與學生作文資料夾，AI 自動連續批改到底，並將個人 Word/PDF 診斷單及全班 Excel 總表**直接儲存在同一作文資料夾下**。

---

## 🏛️ 專案模組架構 (Architecture)

本技能遵循漸進揭示 (Progressive Disclosure) 設計規範，採用三層式架構：

```text
chinese-essay-evaluator/
├── SKILL.md                              # 核心引導文檔：提詞模板、雙軌學段、評閱工作流、自檢清單
├── README.md                             # 專案介紹、安裝指南、提詞範例與使用說明
├── requirements.txt                      # 相依套件清單 (python-docx, PyMuPDF, openpyxl, docx2pdf)
├── LICENSE                               # MIT 開源授權條款
├── scripts/                              # 自動化輔助工具庫
│   ├── rubric_plan_builder.py            # 《嚴格評審準則規畫書》Word (.docx) 自動生成引擎
│   ├── essay_report_builder.py           # Word 批閱單/全班手冊自動生成引擎 (含錯別字訂正表、強制分頁、鷹架卡片)
│   ├── class_excel_generator.py          # 班級成績總表 Excel 生成引擎 (雙工作表：向度明細表 + 常模統計表)
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

### 範例一：標準評審模式（適中給分 / 會考官方常態規準）

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

### 範例二：嚴格評審模式（會考高標閱卷 / 衝刺拔尖篩選）

```markdown
/chinese-essay-evaluator
請依據【嚴格給分標準 / 會考高標閱卷模式】進行專業批改：

1. 班級/座號/姓名：901 班 15 號 陳同學
2. 作文題目：〈當我停下來〉
3. 給分尺度：【從嚴評定（高標篩選）】
   - 俗套選材（如考試、生病）若無新穎視角，立意上限鎖定 4 級分。
   - 單一事例絕不給 5 級分以上。
   - 錯字、病句與口語贅字從嚴挑出並在級分邊界從嚴降級。
4. 學生文章內容（或上傳手寫 PDF / 圖片）：
   [在此貼上文章內容或檔案路徑]
```

### 範例三：兩階段工作流－階段一【命題與嚴格評審準則規畫】（產出規畫書 .docx）

```markdown
/chinese-essay-evaluator 【階段一：命題與評審準則規畫】

請針對以下寫作題目與引導材料，產出一份《嚴格評審準則規畫書》Word 檔（.docx）：

1. 學段與年級：國中九年級（會考衝刺標準）
2. 作文題目：〈當我停下來〉
3. 引導寫作說明／情境限制：
   [在此貼上引導說明，例如：以身邊天天見得到的人事物為觀察核心，不可寫罕見出遊或補習班生活]
4. 給分尺度要求：
   - 採【會考高標／嚴格閱卷】模式
   - 俗套選材（如生病發燒、段考失利）若無深刻新視角，立意上限鎖定 4 級分
   - 單一事例上限鎖定 4 級分，必須有客觀具體畫面特寫（視覺/聽覺等）
   - 明確定義切題、踩線、偏題之排他性條件
5. 預計儲存路徑：[例如：D:\作文批改\九年級模考\當我停下來_嚴格評審準則規畫書.docx]
```

### 範例四：兩階段工作流－階段二【依規畫書批量批閱】（指定資料夾與規畫書，同目錄存檔）

```markdown
/chinese-essay-evaluator 【階段二：依規畫書批量批閱】

我已確認評審準則，請嚴格依照該規畫書，對指定資料夾內的所有作文進行自動批閱：

1. 評審準則規畫書路徑：
   [例如：D:\作文批改\九年級模考\當我停下來_嚴格評審準則規畫書.docx]
2. 學生作文資料夾路徑：
   [例如：D:\作文批改\九年級模考\901班作文\]
   （資料夾內包含學生手寫掃描 PDF 或圖檔）
3. 執行與輸出要求：
   - 自動連續批閱到底，無需每篇中斷確認
   - 嚴格落實規畫書中的篇幅防守線、排他性審題與錯字天頭座標定位
   - 成果檔案請【直接儲存在與學生作文相同的目錄下】：
     * 每生個人作文評閱診斷單（.docx 與同名高畫質向量 .pdf）
     * 全班成績總表（.xlsx，雙 Sheet：向度明細表 + 常模統計表）
```

---

## 💻 自動產生 Word 批閱單與全班手冊 (CLI Usage)

```bash
# 1. 產出《嚴格評審準則規畫書》Word (.docx)
python scripts/rubric_plan_builder.py --topic "當我停下來" --stage "國中九年級（會考衝刺）" --guidance "以身邊天天見得到的人事物為觀察核心，不可寫罕見出遊" --output "當我停下來_嚴格評審準則規畫書.docx"

# 2. 產出單生標準診斷單 (Word + PDF)
python scripts/essay_report_builder.py --topic "當我停下來" --student "901班 05號 李同學" --score 4 --output "李同學作文診斷單.docx"

# 3. 傳入 JSON 資料生成全班彙整手冊（含封面總覽表、每生強制獨立分頁、全班教學診斷專頁）
python scripts/essay_report_builder.py --data-json "class_results.json" --output "全班寫作評閱手冊.docx"
python scripts/class_excel_generator.py "class_results.json" "班級作文成績總表.xlsx"

# 4. 處理學生手寫掃描檔 PDF (支援局部頁面轉正，如 --pages even 或 --pages 2,4)
python scripts/pdf_essay_extractor.py "學生手寫稿卷.pdf" --output-dir "scanned_pages"
```

---

## ✨ 核心特色與升級功能

1. **兩階段人機協同 (HITL) 批閱管線**：
   - 支援「命題規畫書生成」➔「教師審閱/微調確認」➔「資料夾批量批閱」完整自動化流。
   - 批量批閱成果報告（個人 Word/PDF 與全班 Excel 總表）一律自動存放在作文資料夾同層目錄。
2. **嚴格執行考場篇幅行數防守線**：
   - 國中教育會考稿紙：$<75\%$ 行數上限鎖定 $\le 3$ 級分；$<90\%$ 行數上限鎖定 $\le 4$ 級分；$\ge 90\%$ 方具備 5~6 級分（上卷）評定資格。
3. **字數不足強制啟動「寫作進階鷹架」**：
   - 當字數 $<500$ 字時，主動給予「細節擴寫提問」、「五感描摹錦囊」與「邏輯/情感深度反思」，給予具體寫作支撐。
4. **800~1200 字完整版昇華範文**：
   - 以學生原創靈光為骨幹，重構成 6 級分（A++）完整範文，開頭、轉折、收束絕不節錄縮水。
5. **手寫掃描檔批次自動連續管線**：
   - 支援整份多頁手寫 PDF 一鍵批改到底，自動統計全班成績矩陣。
6. **Word 排版極致規範（每生強制換頁）**：
7. **考場天頭錯別字精準座標訂正表**：
   - 實體還原考場批改實務，以教學直覺型 4 欄表格（錯字詞語、卷面座標 `P1 第 X 行第 Y 格`、錯誤字修正、語境訂正與字義解析）精準呈現，輔導學生落實文字正確性。

---

## 🛠️ 環境相依需求 (Dependencies)

```bash
pip install -r requirements.txt
```
包含：`python-docx` 與 `PyMuPDF`。
