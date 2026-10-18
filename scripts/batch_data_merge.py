# -*- coding: utf-8 -*-
"""
batch_data_merge.py
專業全班作文評閱資料智慧合併、正規化校準與報表一鍵自動產生器
功能：
1. 接收多個子代理產出之部分 JSON (如 eval_part1.json, eval_part2.json...) 或資料夾
2. 智慧映射所有常見欄位別名 (score, wrong_words, section_comments 等)
3. 自動按座號數值排序 (01, 02, ..., 27)，偵測座號異常與缺考
4. 匯出標準 class_results.json
5. 一鍵調用 essay_report_builder.py (產出 Word + PDF) 與 class_excel_generator.py (產出 Excel)
"""

import os
import sys
import glob
import json
import argparse
import subprocess

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

def normalize_student(raw, index=1):
    """將任意子代理格式之單生資料正規化為標準規格"""
    s = {}
    
    # 座號提取與補零
    seat_raw = raw.get('seat_no', raw.get('seat', raw.get('seat_number', '')))
    if str(seat_raw).isdigit():
        s['seat_no'] = f"{int(seat_raw):02d}"
    elif seat_raw:
        s['seat_no'] = str(seat_raw).replace('號', '').strip()
    else:
        # 從 student_info 提取
        info_str = raw.get('student_info', f"{index:02d}")
        parts = info_str.split()
        if len(parts) >= 2 and any(c.isdigit() for c in parts[0] + parts[1]):
            s['seat_no'] = parts[0].replace('號', '').zfill(2) if '號' in parts[0] or parts[0].isdigit() else parts[1].replace('號', '').zfill(2)
        else:
            s['seat_no'] = f"{index:02d}"
            
    # 姓名提取
    name = raw.get('name', raw.get('student_name', ''))
    if not name:
        info_str = raw.get('student_info', f"學生 {index}")
        parts = info_str.split()
        if len(parts) >= 3:
            name = parts[2]
        elif len(parts) == 2:
            name = parts[1]
        else:
            name = info_str
    s['name'] = str(name).strip()
    s['student_info'] = f"{s['seat_no']}號 {s['name']}"
    
    # 卷面座號筆誤記註
    paper_seat_no = raw.get('paper_seat_no', raw.get('seat_no_on_paper', raw.get('actual_seat_no', '')))
    if paper_seat_no and str(paper_seat_no).strip() != s['seat_no'] and str(paper_seat_no).replace('號', '').strip() != s['seat_no'].lstrip('0'):
        s['paper_seat_no'] = str(paper_seat_no).replace('號', '').strip()
    else:
        s['paper_seat_no'] = s['seat_no']
        
    # 題目
    s['topic'] = raw.get('topic', raw.get('title', '未定題'))
    
    # 成績 (0~6)
    score_val = raw.get('score', raw.get('final_score', raw.get('overall_score', raw.get('grade', 4))))
    if isinstance(score_val, dict):
        score_val = score_val.get('score', 4)
    try:
        s['score'] = int(score_val)
    except:
        s['score'] = 4
        
    # 等第推導
    tier_defaults = {
        6: "上卷精熟 (A++)",
        5: "上卷優良 (A+)",
        4: "中卷穩健 (B++)",
        3: "中卷待加強 (B)",
        2: "下卷基礎弱 (C)",
        1: "下卷基礎弱 (C)",
        0: "零級分判定"
    }
    s['tier'] = raw.get('tier', raw.get('grade_level', raw.get('rank_level', tier_defaults.get(s['score'], "中卷穩健 (B++)"))))
    
    # 篇幅檢核質性用語
    s['page_status'] = raw.get('page_status', raw.get('length_status', raw.get('length_check', '篇幅充裕')))
    
    # 一句話名師總評
    s['one_sentence_review'] = raw.get('one_sentence_review', raw.get('teacher_summary', raw.get('teacher_comment', raw.get('review', '論述切題通暢。'))))
    if isinstance(s['one_sentence_review'], dict):
        s['one_sentence_review'] = s['one_sentence_review'].get('summary', '論述切題通暢。')
        
    # 四大向度
    dims = raw.get('dimensions', [])
    if not dims and 'score_matrix' in raw:
        sm = raw['score_matrix']
        dims = [
            ["立意取材", f"{sm.get('立意取材', {}).get('score', s['score'])} 級分", sm.get('立意取材', {}).get('comment', '切合題意。')],
            ["結構組織", f"{sm.get('結構組織', {}).get('score', s['score'])} 級分", sm.get('結構組織', {}).get('comment', '佈局條理分明。')],
            ["遣詞造句", f"{sm.get('遣詞造句', {}).get('score', s['score'])} 級分", sm.get('遣詞造句', {}).get('comment', '修辭自然通順。')],
            ["錯別字與標點", f"{sm.get('錯別字與標點', {}).get('score', s['score'])} 級分", sm.get('錯別字與標點', {}).get('comment', '卷面工整。')]
        ]
    elif not dims and 'subscores' in raw:
        sub = raw['subscores']
        dims = [
            ["立意取材", f"{sub.get('content', s['score'])} 級分", "題材貼合度與哲思轉化。"],
            ["結構組織", f"{sub.get('structure', s['score'])} 級分", "篇章完整度與段落佈局。"],
            ["遣詞造句", f"{sub.get('language', s['score'])} 級分", "文句通暢度與修辭表現。"],
            ["錯別字與標點", f"{sub.get('convention', s['score'])} 級分", "錯字率與標點精確度。"]
        ]
    s['dimensions'] = dims
    
    # 錯別字列表 (段落與前後文定位)
    raw_ww = raw.get('wrong_words', raw.get('corrections', raw.get('typos', [])))
    ww_list = []
    for item in raw_ww:
        if isinstance(item, dict):
            w = item.get('word', item.get('wrong', ''))
            c = item.get('context', item.get('location', ''))
            r = item.get('correction', item.get('correct', ''))
            n = item.get('note', '')
            if w or r:
                ww_list.append({
                    "word": w,
                    "context": c if c else f"出處原句：「……{w}……」",
                    "correction": r if '➔' in str(r) else f"{w} ➔ {r}",
                    "note": n
                })
        elif isinstance(item, (list, tuple)) and len(item) >= 2:
            ww_list.append({
                "word": item[0],
                "context": item[1] if len(item)>=3 else "出處原句",
                "correction": item[2] if len(item)>=3 else item[1],
                "note": item[3] if len(item)>=4 else ""
            })
        elif isinstance(item, str) and item.strip():
            ww_list.append({
                "word": item.strip(),
                "context": "文中段落",
                "correction": item.strip(),
                "note": ""
            })
    s['wrong_words'] = ww_list
    
    # 逐段批註
    s['section_comments'] = raw.get('section_comments', raw.get('paragraphs', raw.get('paragraph_annotations', raw.get('paragraph_critiques', []))))
    
    # 寫作鷹架
    s['scaffold'] = raw.get('scaffold', None)
    
    # 升級改寫
    s['upgrade_demo'] = raw.get('upgrade_demo', raw.get('upgrade_section', raw.get('upgrade_prescription', {})))
    
    # 範文示範
    s['model_essay'] = raw.get('model_essay', raw.get('full_upgraded_essay', None))
    
    # 預設不插入原卷圖檔
    s['include_images'] = raw.get('include_images', False)
    if 'essay_images' in raw and s['include_images']:
        s['essay_images'] = raw['essay_images']
    else:
        s['essay_images'] = []
        
    return s

def run_pipeline(inputs, output_dir, class_name=None, topic=None, include_images=False):
    """執行全流程整併與報表產生"""
    os.makedirs(output_dir, exist_ok=True)
    
    # 1. 搜集所有 JSON 檔案
    json_files = []
    for pattern in inputs:
        matches = glob.glob(pattern)
        if matches:
            json_files.extend(matches)
        elif os.path.exists(pattern):
            json_files.append(pattern)
            
    if not json_files:
        print(f"錯誤：找不到任何輸入 JSON 檔案 (輸入參數: {inputs})")
        sys.exit(1)
        
    print(f"找到 {len(json_files)} 個評閱資料檔，開始讀取與整併...")
    
    raw_students = []
    class_diag = None
    default_title = None
    
    for jf in json_files:
        try:
            with open(jf, 'r', encoding='utf-8') as f:
                d = json.load(f)
            if isinstance(d, list):
                raw_students.extend(d)
            elif isinstance(d, dict):
                if 'students' in d:
                    raw_students.extend(d['students'])
                    if not class_diag and 'class_diagnosis' in d:
                        class_diag = d['class_diagnosis']
                    if not default_title and 'title' in d:
                        default_title = d['title']
                elif 'evaluations' in d:
                    raw_students.extend(d['evaluations'])
                else:
                    # 嘗試抓取 values 中的 list
                    for v in d.values():
                        if isinstance(v, list) and len(v) > 0 and isinstance(v[0], dict):
                            raw_students.extend(v)
                            break
                    else:
                        raw_students.append(d)
        except Exception as e:
            print(f"警告：讀取 {jf} 失敗 ({e})，略過該檔。")
            
    print(f"成功搜集 {len(raw_students)} 位學生評閱數據，正在進行正規化校準...")
    
    # 2. 正規化
    normalized = []
    for idx, r in enumerate(raw_students, 1):
        norm_s = normalize_student(r, idx)
        if topic and (not norm_s.get('topic') or norm_s.get('topic') == '未定題'):
            norm_s['topic'] = topic
        norm_s['include_images'] = include_images
        normalized.append(norm_s)
        
    # 3. 按座號數值排序
    def seat_sort_key(s):
        try:
            return int(s['seat_no'])
        except:
            return 999
    normalized.sort(key=seat_sort_key)
    
    # 4. 偵測座號重複或勘誤
    seat_counts = {}
    anomalies = []
    for s in normalized:
        st = s['seat_no']
        seat_counts[st] = seat_counts.get(st, 0) + 1
        if s.get('paper_seat_no') and s['paper_seat_no'] != st:
            anomalies.append(f"座號 {st} {s['name']} (原卷誤寫為 {s['paper_seat_no']} 號)")
            
    dupes = [k for k, v in seat_counts.items() if v > 1]
    if dupes:
        print(f"⚠️ 提示：發現重複座號：{dupes}，請確認。")
    if anomalies:
        print(f"💡 卷面座號筆誤勘誤：{', '.join(anomalies)}")
        
    # 5. 組裝全班 Master JSON
    inferred_class = class_name if class_name else (normalized[0]['student_info'].split()[0] if normalized else "國中部")
    inferred_topic = topic if topic else (normalized[0].get('topic', '段考作文') if normalized else "段考作文")
    
    master_title = default_title if default_title else f"國中部 {inferred_class} 國文寫作測驗評閱手冊〈{inferred_topic}〉"
    master_data = {
        "title": master_title,
        "subtitle": "國中教育會考評分規準 ｜ 心測中心四大向度 + 篇幅防守線 + 原句錯字訂正",
        "class_name": inferred_class,
        "topic": inferred_topic,
        "include_images": include_images,
        "students": normalized,
        "class_diagnosis": class_diag if class_diag else {
            "instructor": "國文領域教學研究會",
            "common_blindspots": "普遍缺乏具體生活客觀畫面，感悟流於口號化；少數同學未寫滿第1頁篇幅受限。",
            "activities": [
                "1. 【感官特寫盲寫練習】：播放 3 分鐘生活片段，限時刻畫人物眼神與環境聲響。",
                "2. 【雙例對讀遞進佈局】：訓練學生一例一段，層次推進生命哲思。"
            ]
        }
    }
    
    master_json_path = os.path.join(output_dir, f"{inferred_class}_class_results.json")
    with open(master_json_path, 'w', encoding='utf-8') as f:
        json.dump(master_data, f, ensure_ascii=False, indent=2)
    print(f"✅ 全班正規化資料已儲存：{master_json_path}")
    
    # 6. 一鍵調用 essay_report_builder.py
    script_dir = os.path.dirname(os.path.abspath(__file__))
    builder_script = os.path.join(script_dir, "essay_report_builder.py")
    excel_script = os.path.join(script_dir, "class_excel_generator.py")
    
    out_docx_name = f"{inferred_class}作文評閱診斷彙整手冊.docx"
    out_docx_path = os.path.join(output_dir, out_docx_name)
    
    out_xlsx_name = f"{inferred_class}作文成績總表.xlsx"
    out_xlsx_path = os.path.join(output_dir, out_xlsx_name)
    
    print("\n🚀 開始自動產出三合一標準成果報告...")
    
    # 生成 Word + 向量 PDF
    cmd_doc = [
        sys.executable, builder_script,
        "--data-json", master_json_path,
        "--output", out_docx_path
    ]
    if include_images:
        cmd_doc.append("--include-images")
    subprocess.run(cmd_doc, check=True)
    
    # 生成 Excel 雙 Sheet 總表
    cmd_xls = [
        sys.executable, excel_script,
        master_json_path,
        out_xlsx_path
    ]
    subprocess.run(cmd_xls, check=True)
    
    # 7. 產出成績統計看板
    scores = [s['score'] for s in normalized]
    avg_score = sum(scores) / len(scores) if scores else 0
    dist = {i: scores.count(i) for i in range(6, -1, -1)}
    
    print("\n" + "="*50)
    print(f"🎉 【{inferred_class}】全班作文評閱完成看板")
    print("="*50)
    print(f"• 應試人數：{len(normalized)} 人")
    print(f"• 全班平均：{avg_score:.2f} 級分")
    print("• 級分分佈：")
    for g in range(6, -1, -1):
        cnt = dist.get(g, 0)
        pct = (cnt / len(normalized) * 100) if normalized else 0
        bar = "█" * cnt
        print(f"  {g} 級分: {cnt:2d} 人 ({pct:5.1f}%) {bar}")
    if anomalies:
        print(f"• 卷面筆誤勘誤：{len(anomalies)} 件")
        for a in anomalies:
            print(f"  - {a}")
    print("\n📁 最終成果交付檔案（全部存放於作文同層目錄）：")
    print(f"1. 📄 Word 彙整手冊：{out_docx_path}")
    pdf_path = os.path.splitext(out_docx_path)[0] + ".pdf"
    if os.path.exists(pdf_path):
        print(f"2. 📑 向量 PDF 手冊：{pdf_path}")
    print(f"3. 📊 Excel 成績總表：{out_xlsx_path}")
    print("="*50 + "\n")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="全班作文評閱成果自動合併與報表產生管線 (batch_data_merge)")
    parser.add_argument('--inputs', nargs='+', required=True, help="輸入之 JSON 檔案路徑（支援通配符如 eval_part*.json）")
    parser.add_argument('--output-dir', required=True, help="成果報告輸出目錄（通常為學生作文資料夾）")
    parser.add_argument('--class-name', default=None, help="班級名稱（如 705班）")
    parser.add_argument('--topic', default=None, help="寫作題目")
    parser.add_argument('--include-images', action='store_true', default=False, help="是否插入學生手寫原卷圖檔（預設 False 不插入）")
    
    args = parser.parse_args()
    run_pipeline(args.inputs, args.output_dir, args.class_name, args.topic, args.include_images)
