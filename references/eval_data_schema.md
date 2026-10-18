# 學生作文評審資料交換結構規範 (eval_data_schema.md)

本文件定義 `/chinese-essay-evaluator` 技能在批閱單篇或全班學生作文時的標準 JSON 資料交換格式（Data Contract），確保多子代理並行批閱、資料整併與報表工具鏈（Word, PDF, Excel）100% 欄位一致。

---

## 一、 全班資料集結構 (Batch Class Dataset)

當執行全班批量批閱時，請將所有學生評估資料彙整至以下根結構：

```json
{
  "title": "新北市立錦和高中 115-1 國中部七年五班 寫作測驗評閱手冊",
  "subtitle": "國中教育會考評分規準 ｜ 心測中心四大向度 + 篇幅防守線 + 原句錯字訂正",
  "class_name": "705班",
  "default_topic": "記憶中的滋味",
  "include_images": false,
  "students": [
    {
      "seat_no": "01",
      "name": "丁唯硯",
      "paper_seat_no": "01",
      "topic": "記憶中的滋味",
      "score": 3,
      "tier": "中卷待加強 (B)",
      "page_status": "內容略少（第1頁寫18行未滿，篇幅上限受限）",
      "one_sentence_review": "生活取材平實親切，可惜文末受限時間稍嫌倉促，未及展開深刻哲思。",
      "dimensions": [
        ["立意取材", "3 級分", "題材貼合生活但停留在表面事件敘述，缺乏由物入理之深刻體悟。"],
        ["結構組織", "3 級分", "段落劃分尚稱清楚，但結尾段過於短促，未能首尾呼應。"],
        ["遣詞造句", "4 級分", "文句通順自然，偶有口語贅字，整體表意清楚。"],
        ["錯別字與標點", "4 級分", "全卷錯字極少，標點符號運用大致合規。"]
      ],
      "wrong_words": [
        {
          "word": "自已",
          "context": "第 1 段：「……常常看見自已的影子……」",
          "correction": "已 ➔ 己",
          "note": "應為「自己」（指自身），不可誤寫為已經之「已」。"
        }
      ],
      "section_comments": [
        {"section": "起（開篇破題）", "comment": "以熟悉的家常氣味破題，自然流暢。"},
        {"section": "承（情境展開）", "comment": "細緻刻畫廚房熱氣與家人的互動，生活感強。"},
        {"section": "轉（心境遞進）", "comment": "帶出成長過程中的酸甜苦辣，情感真摯。"},
        {"section": "合（收束昇華）", "comment": "稍顯急促，若能多花三兩行深化體悟會更圓滿。"}
      ],
      "scaffold": {
        "need_scaffold": true,
        "detail_prompts": ["當時廚房傳來的聲音是什麼樣的？", "看到那道菜端上桌時，對方的眼神有何變化？"],
        "sensory_tips": "視覺：金黃微焦的蛋皮；嗅覺：濃郁撲鼻的麻油香氣。",
        "reflection_guide": "從單純的味覺享受，昇華為對家人長年辛勞的感念。"
      },
      "upgrade_demo": {
        "pain_point": "情感感悟較為口號化，缺乏生活細節映襯。",
        "original_text": "這道菜真的很好吃，每次吃完我都覺得很幸福，也感謝媽媽的辛苦。",
        "upgraded_text": "當那抹焦香在舌尖化開，滾燙的不只是喉間的溫度，更是母親在晨光微熹中洗手作羹湯的默默守候。原來記憶中最深邃的滋味，從來不是調味料的堆疊，而是歲月沉澱下的深情。",
        "tips": "以具體動作與感官特寫取代空泛的『覺得很幸福』，由物入理。"
      },
      "model_essay": "完整 800~1200 字示範範文全文……"
    }
  ],
  "class_diagnosis": {
    "instructor": "國文領域教學研究會",
    "common_blindspots": "全班普遍存在事例單一、篇幅受限於第1頁之問題；優良卷多能以具體細節由物入理。",
    "activities": [
      "1. 【感官特寫微寫作】：限時 100 字具體刻畫食物熱氣與人物微表情。",
      "2. 【雙例遞進佈局對讀】：引導學生落實『一例一段、雙例映襯』結構。"
    ]
  }
}
```

---

## 二、 欄位別名智慧相容表 (Smart Aliases)

為避免子代理輸出慣用詞不同，報表產生工具支援以下別名自動映射：

| 標準欄位 | 支援別名 | 型態 | 說明 |
| :--- | :--- | :--- | :--- |
| `seat_no` | `seat`, `seat_number` | string / int | 學生座號，腳本會自動補零為兩位數（如 `"01"`） |
| `name` | `student_name` | string | 學生姓名 |
| `paper_seat_no` | `seat_no_on_paper`, `actual_seat_no` | string | 學生在卷面上誤填的座號（若與 `seat_no` 不同自動標記勘誤） |
| `topic` | `title` | string | 寫作題目 |
| `score` | `final_score`, `overall_score`, `grade` | int | 綜合會考級分（純 0~6 之整數） |
| `tier` | `grade_level`, `rank_level` | string | 等第類別（如 `"上卷優良 (A+)"`） |
| `page_status` | `length_status`, `length_check` | string | 卷面篇幅檢核質性用語（如 `"篇幅充裕"`、`"內容略少"`） |
| `one_sentence_review` | `teacher_summary`, `teacher_comment`, `review` | string | 名師總評一句話短評 |
| `wrong_words` | `corrections`, `typos` | list | 錯別字訂正清單（支援 dict 或 tuple） |
| `section_comments` | `paragraphs`, `paragraph_annotations` | list | 逐段批註與脈絡診斷 |
| `upgrade_demo` | `upgrade_section`, `upgrade_prescription` | dict | 段落升級改寫前後對照與要訣 |
| `include_images` | `with_images` | bool | 是否插入原卷圖檔，**預設為 false** |
