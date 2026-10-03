# 國中會考寫作測驗標準評審與深度評閱技能 (chinese-essay-evaluator)

本技能 (Skill) 專門用於指導 Google Antigravity 等 AI 代理人依據台灣**「國中教育會考寫作測驗標準評分規準」**與**「第一線資深國文名師閱卷實戰指標」**，對各類國中引導式寫作或命題作文進行嚴格客觀的級分判定（0～6 級分）、篇幅行數檢核、天頭錯別字挑錯、逐段紅筆眉批與段落升級改寫示範，並支援自動產出可直接列印發放之 **A4 作文批閱診斷單 (`.docx`)**。

---

## 🏛️ 專案模組架構 (Architecture)

本技能遵循 Antigravity 漸進揭示 (Progressive Disclosure) 設計規範，採用精簡三層架構：

```text
chinese-essay-evaluator/
├── SKILL.md                              # 核心引導文檔：提詞模板、四階段評閱工作流、自檢清單
├── README.md                             # 專案介紹、安裝指南、提詞範例與使用說明
├── requirements.txt                      # 相依套件清單 (python-docx)
├── LICENSE                               # MIT 開源授權條款
├── scripts/                              # 自動化輔助工具庫
│   └── essay_report_builder.py           # Word 批閱診斷報告單自動產出引擎 (排版精美、天頭訂正格、可列印發放)
└── references/                           # 核心評分法規與批閱寶典庫
    ├── cap_essay_rubric.md               # 心測中心官方寫作測驗四大向度 0~6 級分標準規準
    ├── teacher_grading_secrets.md        # 現場名師批閱心法：篇幅防守線、一例一段、首尾呼應、由物入理
    └── sample_prompts_and_upgrades.md    # 歷屆會考經典考題排他性審題拆解與段落升級示範
```

---

## 📦 如何安裝本技能

### 方式 A：全域安裝（推薦，所有專案與新對話皆可隨時調用）

#### 1. 在 Antigravity 對話框貼上一句話自動安裝（最推薦）：
```text
請幫我把這個 GitHub 倉庫安裝成全域 Skill：
https://github.com/geniefu/chinese-essay-evaluator
並幫我檢查安裝所需的 Python 套件 (python-docx)
```

#### 2. 透過 Git 指令一鍵安裝（未來更新只需 `git pull`）：
- **Windows (PowerShell)**:
  ```powershell
  git clone https://github.com/geniefu/chinese-essay-evaluator.git "$HOME\.gemini\config\skills\chinese-essay-evaluator"
  ```
- **macOS / Linux**:
  ```bash
  git clone https://github.com/geniefu/chinese-essay-evaluator.git ~/.gemini/config/skills/chinese-essay-evaluator
  ```

#### 3. 手動解壓縮安裝：
將 `chinese-essay-evaluator` 資料夾複製至個人全域設定目錄：
- **Windows**: `C:\Users\<使用者名稱>\.gemini\config\skills\chinese-essay-evaluator\`
- **macOS / Linux**: `~/.gemini/config/skills/chinese-essay-evaluator/`

---

## 🚀 快速開始：提詞範例 (Prompt)

安裝完成後，在對話框中輸入 `/chinese-essay-evaluator` 並貼上作文題目與學生文章：

```text
請使用 /chinese-essay-evaluator 技能為我進行國中會考作文專業批閱與診斷，規格參數如下：
1、標題：【例如：新北市立錦和高中九年級第1次模考寫作測驗 或 113年會考作文模擬評閱】。
2、寫作題目：【請填寫題目，例如：〈當我停下來〉 或 〈在這次相遇中〉】。
3、引導材料與限定條件：【請填寫題目附帶之引導說明，例如：請以身邊天天見得到的人事物為觀察核心，並寫出個人具體看見的畫面與體悟】。
4、學生作文全文：
【請在此貼上學生文章原文】
5、卷面篇幅預估：【例如：單頁寫了約 18 行（佔比約 85%），或未滿一頁約 14 行（約 65%），或寫滿一頁跨入第二頁】。
6、輸出格式：【例如：請在對話中輸出完整專業評閱報告，並額外產出 Word (.docx) 精美批閱診斷單供列印發放】。
7、其它：【選填，例如：針對該生常常出現的「通篇抽象心靈雞湯」加強具體畫面描摹示範】。
```

---

## 💻 自動產生可列印 Word 批閱單 (CLI Usage)

若需要將評閱結果轉化為可印製發放之 A4 批閱診斷單：

```bash
# 產出基礎診斷單
python scripts/essay_report_builder.py --topic "當我停下來" --student "909班 01號 張同學" --score 4 --output "作文評閱診斷單.docx"

# 依據完整評閱診斷 JSON 檔案產生詳細報告單
python scripts/essay_report_builder.py --data-json "eval_result.json" --output "評閱單_90901.docx"
```

---

## ✨ 五大核心特色

1. **嚴格執行考場篇幅行數防守線**：
   - 未滿一頁且行數未達 75% ➔ **最高鎖定 $\le 3$ 級分**。
   - 未滿一頁且行數未達 90% ➔ **最高鎖定 $\le 4$ 級分**。
   - 達 90% 以上或跨入第 2 頁（約 550 字以上）➔ 具備評定 5～6 級分（上卷）之基本門檻。
2. **排他性審題與客觀畫面檢核**：
   - 嚴格抓出「看似切題實則跑題」之舉例（如日常慣常題目卻寫罕見出遊）；杜絕「通篇空洞心靈雞湯」給高分，強制檢核視覺/感官之具體客觀畫面。
3. **例證數量硬指標與「一例一段」**：
   - 僅舉單一事例一律限縮 4 級分以下；倡導「2~3 個事例層次展開、夾敘夾議、首尾呼應」之高分架構。
4. **考場天頭錯別字訂正格**：
   - 實體還原考場批閱實務，明確標註 `錯字 口 ➔ 正字：【 正字 】`，具備實質輔導補救效益。
   - 嚴格糾舉「一逗到底」、標點混亂與段落空兩格格式。
5. **原創段落升級改寫示範**：
   - 絕非套話！針對學生最弱段落，以「生活細節特寫、由物入理、映襯修辭」示範如何將 4 級分句子昇華為 5~6 級分典範。

---

## 🛠️ 環境相依需求 (Dependencies)

```bash
pip install -r requirements.txt
```
或直接安裝：
```bash
pip install python-docx
```
