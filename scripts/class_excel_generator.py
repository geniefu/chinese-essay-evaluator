# -*- coding: utf-8 -*-
"""
class_excel_generator.py
專業生成班級學生作文成績總表 Excel (*.xlsx)
包含：
1. Sheet 1: 【全班成績與向度明細表】(含凍結窗格、自動篩選、條件色彩高亮、12大向度明細)
2. Sheet 2: 【級分常模與統計圖表】(全班平均、級分分佈表、四大向度強弱統計、高頻錯字榜)
"""

import os
import sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_class_excel_report(batch_data, output_path):
    wb = Workbook()
    
    # -------------------------------------------------------------
    # 樣式定義
    # -------------------------------------------------------------
    font_name = '微軟正黑體'
    
    font_title = Font(name=font_name, size=16, bold=True, color='1E3A8A')
    font_subtitle = Font(name=font_name, size=10, bold=False, color='64748B')
    font_header = Font(name=font_name, size=10, bold=True, color='FFFFFF')
    font_sub_header = Font(name=font_name, size=11, bold=True, color='1E3A8A')
    
    fill_header = PatternFill(start_color='1E3A8A', end_color='1E3A8A', fill_type='solid')
    fill_sub_hdr = PatternFill(start_color='F1F5F9', end_color='F1F5F9', fill_type='solid')
    fill_alt = PatternFill(start_color='F8FAFC', end_color='F8FAFC', fill_type='solid')
    fill_white = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    
    # 級分高亮底色
    fill_score_6 = PatternFill(start_color='DCFCE7', end_color='DCFCE7', fill_type='solid')
    font_score_6 = Font(name=font_name, size=11, bold=True, color='15803D')
    
    fill_score_5 = PatternFill(start_color='ECFDF5', end_color='ECFDF5', fill_type='solid')
    font_score_5 = Font(name=font_name, size=11, bold=True, color='047857')
    
    fill_score_4 = PatternFill(start_color='EFF6FF', end_color='EFF6FF', fill_type='solid')
    font_score_4 = Font(name=font_name, size=11, bold=True, color='1D4ED8')
    
    fill_score_3 = PatternFill(start_color='FEF9C3', end_color='FEF9C3', fill_type='solid')
    font_score_3 = Font(name=font_name, size=11, bold=True, color='B45309')
    
    fill_score_low = PatternFill(start_color='FEE2E2', end_color='FEE2E2', fill_type='solid')
    font_score_low = Font(name=font_name, size=11, bold=True, color='B91C1C')
    
    # 邊框
    thin_border_side = Side(border_style='thin', color='CBD5E1')
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    
    # -------------------------------------------------------------
    # Sheet 1: 【全班成績與向度明細表】
    # -------------------------------------------------------------
    ws1 = wb.active
    ws1.title = "全班成績與向度明細表"
    ws1.views.sheetView[0].showGridLines = True
    
    title_text = batch_data.get('title', "國中會考寫作測驗 班級成績總表")
    subtitle_text = batch_data.get('subtitle', "國中教育會考評分規準 ｜ 心測中心四大向度 + 篇幅防守線 + 錯別字訂正")
    students = batch_data.get('students', [])
    
    # 標題列
    ws1.merge_cells('A1:L1')
    ws1['A1'] = f"📊 {title_text}"
    ws1['A1'].font = font_title
    ws1['A1'].alignment = Alignment(horizontal='left', vertical='center')
    ws1.row_dimensions[1].height = 32
    
    ws1.merge_cells('A2:L2')
    ws1['A2'] = subtitle_text
    ws1['A2'].font = font_subtitle
    ws1['A2'].alignment = Alignment(horizontal='left', vertical='center')
    ws1.row_dimensions[2].height = 18
    
    # 表頭欄位
    headers = [
        "座號", "姓名", "寫作題目", "綜合會考級分",
        "立意取材", "結構組織", "遣詞造句", "錯別字與標點",
        "卷面篇幅檢核", "錯別字數", "錯字訂正摘要", "名師總評"
    ]
    
    header_row = 3
    ws1.row_dimensions[header_row].height = 26
    for col_idx, h in enumerate(headers, 1):
        cell = ws1.cell(row=header_row, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = cell_border
        
    # 填入學生資料
    for row_idx, s in enumerate(students, start=4):
        ws1.row_dimensions[row_idx].height = 28
        is_even = (row_idx % 2 == 0)
        base_fill = fill_alt if is_even else fill_white
        
        # 智慧提取座號與姓名 (相容 seat_no, name, student_info 各種欄位形式)
        seat_raw = s.get('seat_no', s.get('seat', s.get('seat_number', '')))
        name_raw = s.get('name', s.get('student_name', ''))
        seat_no = f"{int(seat_raw):02d}" if str(seat_raw).isdigit() else str(seat_raw).replace('號', '').strip()
        name = str(name_raw).strip()
        
        if not name or not seat_no:
            info_str = s.get('student_info', f"學生 {row_idx-3}")
            parts = info_str.split()
            if len(parts) >= 3:
                if not seat_no: seat_no = parts[1].replace('號', '').zfill(2)
                if not name: name = parts[2]
            elif len(parts) == 2:
                if not seat_no: seat_no = parts[0].replace('號', '').zfill(2)
                if not name: name = parts[1]
            elif not name:
                name = info_str
                
        # 標註卷面座號筆誤（若有）
        paper_seat_no = s.get('paper_seat_no', s.get('seat_no_on_paper', ''))
        if paper_seat_no and str(paper_seat_no).strip() != seat_no and str(paper_seat_no).replace('號', '').strip() != seat_no.lstrip('0'):
            name = f"{name} (卷面誤填{paper_seat_no}號)"
            
        topic = f"〈{s.get('topic', s.get('title', '未定題'))}〉"
        is_absent = s.get('is_absent', False) or (s.get('score', 0) == 0 and '缺考' in str(s.get('page_status', '')))
        
        score_val = s.get('score', s.get('final_score', s.get('grade', 4)))
        try:
            score = int(score_val)
        except:
            score = 4
        score_display = "缺考" if is_absent else f"{score} 級分"
        
        # 拆解四大向度 (相容 dimensions, score_matrix, subscores)
        dims = s.get('dimensions', [])
        dim_map = {}
        for d in dims:
            if isinstance(d, (list, tuple)) and len(d) >= 2:
                dim_map[str(d[0]).strip()] = str(d[1]).strip()
            elif isinstance(d, dict):
                dim_map[d.get('name', '')] = d.get('grade', f"{score} 級分")
                
        # 若未提供 dimensions，嘗試從 subscores 提取
        if not dim_map and 'subscores' in s:
            sub = s['subscores']
            dim_map['立意取材'] = f"{sub.get('content', score)} 級分"
            dim_map['結構組織'] = f"{sub.get('structure', score)} 級分"
            dim_map['遣詞造句'] = f"{sub.get('language', score)} 級分"
            dim_map['錯別字與標點'] = f"{sub.get('convention', score)} 級分"
                
        d_liyi = dim_map.get('立意取材', f"{score} 級分" if not is_absent else "缺考")
        d_jiegou = dim_map.get('結構組織', f"{score} 級分" if not is_absent else "缺考")
        d_qianci = dim_map.get('遣詞造句', f"{score} 級分" if not is_absent else "缺考")
        d_cuozi = dim_map.get('錯別字與標點', f"{score} 級分" if not is_absent else "缺考")
        
        page_status = s.get('page_status', s.get('length_status', s.get('length_check', '篇幅充裕')))
        
        # 錯別字統計與摘要 (相容 wrong_words, corrections)
        ww_list = s.get('wrong_words', s.get('corrections', []))
        ww_count = len(ww_list)
        ww_summary_items = []
        for item in ww_list:
            if isinstance(item, dict):
                w_orig = item.get('word', item.get('wrong', ''))
                w_corr = item.get('correction', item.get('correct', ''))
                if w_orig and w_corr:
                    ww_summary_items.append(f"{w_orig}➔{w_corr}")
                elif w_orig:
                    ww_summary_items.append(w_orig)
            elif isinstance(item, (list, tuple)) and len(item) >= 3:
                ww_summary_items.append(f"{item[0]}➔{item[2]}")
            elif isinstance(item, str):
                ww_summary_items.append(item)
        ww_summary = "、".join(ww_summary_items) if ww_summary_items else ("缺考無作答" if is_absent else "卷面嚴謹零錯字")
        
        review = s.get('one_sentence_review', s.get('teacher_summary', s.get('review', s.get('teacher_comment', ''))))
        
        score_highlight = None if is_absent else score
        row_vals = [
            (seat_no, align_center, None, None),
            (name, align_center, Font(name=font_name, size=10, bold=True, color='1E3A8A'), None),
            (topic, align_center, None, None),
            (score_display, align_center, None, score_highlight), # 級分高亮
            (d_liyi, align_center, None, None),
            (d_jiegou, align_center, None, None),
            (d_qianci, align_center, None, None),
            (d_cuozi, align_center, None, None),
            (page_status, align_center, None, 'page'),
            (f"{ww_count} 個" if ww_count > 0 else "0 (嚴謹)", align_center, None, 'ww_cnt' if ww_count==0 else None),
            (ww_summary, align_left, Font(name=font_name, size=9, color='B91C1C' if ww_count>0 else '15803D'), None),
            (review, align_left, Font(name=font_name, size=9, color='334155'), None)
        ]
        
        for col_idx, (val, alignment, custom_font, flag) in enumerate(row_vals, 1):
            cell = ws1.cell(row=row_idx, column=col_idx, value=val)
            cell.alignment = alignment
            cell.border = cell_border
            cell.fill = base_fill
            cell.font = custom_font if custom_font else Font(name=font_name, size=9.5, color='0F172A')
            
            # 特殊高亮
            if flag in [6, 5, 4, 3, 2, 1, 0]:
                if flag == 6:
                    cell.fill = fill_score_6; cell.font = font_score_6
                elif flag == 5:
                    cell.fill = fill_score_5; cell.font = font_score_5
                elif flag == 4:
                    cell.fill = fill_score_4; cell.font = font_score_4
                elif flag == 3:
                    cell.fill = fill_score_3; cell.font = font_score_3
                else:
                    cell.fill = fill_score_low; cell.font = font_score_low
            elif flag == 'page':
                if val == '篇幅充裕':
                    cell.fill = fill_score_6; cell.font = font_score_6
                elif val == '內容略少':
                    cell.fill = fill_score_4; cell.font = font_score_4
                else:
                    cell.fill = fill_score_3; cell.font = font_score_3
            elif flag == 'ww_cnt':
                cell.fill = fill_score_6; cell.font = font_score_6

    # 欄寬設定
    col_widths = {
        'A': 8,   # 座號
        'B': 12,  # 姓名
        'C': 16,  # 題目
        'D': 14,  # 綜合會考級分
        'E': 11,  # 立意取材
        'F': 11,  # 結構組織
        'G': 11,  # 遣詞造句
        'H': 13,  # 錯別字與標點
        'I': 14,  # 卷面篇幅檢核
        'J': 11,  # 錯別字數
        'K': 32,  # 錯字訂正摘要
        'L': 48   # 名師總評
    }
    for col_letter, width in col_widths.items():
        ws1.column_dimensions[col_letter].width = width
        
    # 凍結窗格：凍結前兩欄（A、B座號姓名）及表頭
    ws1.freeze_panes = 'C4'
    # 啟用自動篩選
    last_row = len(students) + 3
    ws1.auto_filter.ref = f"A3:L{last_row}"
    
    # -------------------------------------------------------------
    # Sheet 2: 【級分常模與統計圖表】
    # -------------------------------------------------------------
    ws2 = wb.create_sheet(title="級分常模與統計分析")
    ws2.views.sheetView[0].showGridLines = True
    
    ws2.merge_cells('A1:F1')
    ws2['A1'] = f"📈 {title_text} ── 級分常模與向度分析"
    ws2['A1'].font = font_title
    ws2['A1'].alignment = Alignment(horizontal='left', vertical='center')
    ws2.row_dimensions[1].height = 32
    
    # 1. 班級整體量化看板
    tested_students = [s for s in students if not (s.get('is_absent', False) or (s.get('score', 0) == 0 and '缺考' in s.get('page_status', '')))]
    scores = [s.get('score', 4) for s in tested_students]
    total_students = len(students)
    tested_count = len(tested_students)
    absent_count = total_students - tested_count
    avg_score = round(sum(scores) / tested_count, 2) if tested_count > 0 else 0
    
    ws2.cell(row=3, column=1, value="一、 班級整體成績量化概況").font = font_sub_header
    ws2.row_dimensions[3].height = 24
    
    summary_kpis = [
        ("應試總人數 (實到 / 應到)", f"{tested_count} 人 / {total_students} 人 (缺考 {absent_count} 人)"),
        ("全班平均會考級分 (實考)", f"{avg_score} 級分"),
        ("滿級分 (6級分) 人數", f"{scores.count(6)} 人 ({round(scores.count(6)/tested_count*100, 1) if tested_count else 0}%)"),
        ("上卷精熟 (5~6級分) 人數", f"{scores.count(6)+scores.count(5)} 人 ({round((scores.count(6)+scores.count(5))/tested_count*100, 1) if tested_count else 0}%)"),
        ("穩健中卷 (4級分) 人數", f"{scores.count(4)} 人 ({round(scores.count(4)/tested_count*100, 1) if tested_count else 0}%)"),
        ("待加強 (3級分以下) 人數", f"{sum(1 for sc in scores if sc <= 3)} 人 ({round(sum(1 for sc in scores if sc <= 3)/tested_count*100, 1) if tested_count else 0}%)")
    ]
    
    for idx_k, (k_lbl, k_val) in enumerate(summary_kpis, start=4):
        ws2.row_dimensions[idx_k].height = 22
        c_l = ws2.cell(row=idx_k, column=1, value=k_lbl)
        c_l.font = Font(name=font_name, size=10, bold=False, color='475569')
        c_l.fill = fill_sub_hdr; c_l.border = cell_border; c_l.alignment = align_left
        
        c_v = ws2.cell(row=idx_k, column=2, value=k_val)
        c_v.font = Font(name=font_name, size=10, bold=True, color='1E3A8A')
        c_v.border = cell_border; c_v.alignment = align_center
        
    # 2. 0~6 級分分佈常模表
    row_dist_start = 12
    ws2.cell(row=row_dist_start-1, column=1, value="二、 全班會考級分分佈表 (0 ~ 6 級分)").font = font_sub_header
    dist_headers = ["會考級分", "人數", "百分比佔比", "常模階梯評定"]
    for c_i, dh in enumerate(dist_headers, 1):
        c = ws2.cell(row=row_dist_start, column=c_i, value=dh)
        c.font = font_header; c.fill = fill_header; c.border = cell_border; c.alignment = align_center
        
    tier_descs = {
        6: "精熟頂尖（雙事例展開、由物入理）",
        5: "優良流暢（取材深刻、生活特寫）",
        4: "穩健及格（單一事例、內容略少）",
        3: "待加強（篇幅太少、感悟口號化）",
        2: "基礎薄弱（偏題或文句不順）",
        1: "嚴重不足（字數過少）",
        0: "不予給分（離題或空白）"
    }
    
    for r_offset, sc in enumerate(range(6, -1, -1)):
        curr_r = row_dist_start + 1 + r_offset
        ws2.row_dimensions[curr_r].height = 20
        cnt = scores.count(sc)
        pct = f"{round(cnt/total_students*100, 1)}%" if total_students>0 else "0%"
        
        c_sc = ws2.cell(row=curr_r, column=1, value=f"{sc} 級分")
        c_sc.font = Font(name=font_name, size=10, bold=True, color='1E3A8A')
        c_sc.alignment = align_center; c_sc.border = cell_border
        if sc == 6: c_sc.fill = fill_score_6
        elif sc == 5: c_sc.fill = fill_score_5
        elif sc == 4: c_sc.fill = fill_score_4
        elif sc == 3: c_sc.fill = fill_score_3
        
        c_cnt = ws2.cell(row=curr_r, column=2, value=cnt)
        c_cnt.alignment = align_center; c_cnt.border = cell_border
        
        c_pct = ws2.cell(row=curr_r, column=3, value=pct)
        c_pct.alignment = align_center; c_pct.border = cell_border
        
        c_desc = ws2.cell(row=curr_r, column=4, value=tier_descs.get(sc, ""))
        c_desc.alignment = align_left; c_desc.border = cell_border; c_desc.font = Font(name=font_name, size=9, color='64748B')

    # 3. 四大向度平均表現
    row_dim_start = 22
    ws2.cell(row=row_dim_start-1, column=1, value="三、 心測中心四大評審向度平均得分").font = font_sub_header
    dim_names = ["立意取材", "結構組織", "遣詞造句", "錯別字與標點"]
    for c_i, dh in enumerate(["評審向度", "向度平均", "教學診斷方針"], 1):
        c = ws2.cell(row=row_dim_start, column=c_i, value=dh)
        c.font = font_header; c.fill = fill_header; c.border = cell_border; c.alignment = align_center
        
    for r_offset, dname in enumerate(dim_names):
        curr_r = row_dim_start + 1 + r_offset
        ws2.row_dimensions[curr_r].height = 20
        # 簡單推算該向度平均
        d_scores = []
        for s in students:
            dims = s.get('dimensions', [])
            for d in dims:
                if isinstance(d, (list, tuple)) and len(d) >= 2 and dname in str(d[0]):
                    try:
                        d_scores.append(int(str(d[1]).replace('級分', '').strip()))
                    except:
                        pass
        d_avg = round(sum(d_scores)/len(d_scores), 2) if d_scores else avg_score
        
        ws2.cell(row=curr_r, column=1, value=dname).alignment = align_center
        ws2.cell(row=curr_r, column=1).border = cell_border
        ws2.cell(row=curr_r, column=1).font = Font(name=font_name, size=10, bold=True, color='1E3A8A')
        
        ws2.cell(row=curr_r, column=2, value=f"{d_avg} 級分").alignment = align_center
        ws2.cell(row=curr_r, column=2).border = cell_border
        
        advice_map = {
            "立意取材": "加強生活微觀特寫與『由物入理』思維，避免空泛心靈雞湯口號。",
            "結構組織": "落實『一例一段、2~3個事例』，突破單一事例篇幅防守線受限。",
            "遣詞造句": "強化五感動態白描，減少口語化贅詞，增添譬喻與修辭節奏。",
            "錯別字與標點": "加強常見形近字（已/己、及/急、歷/厲）筆畫區辨與逗號分界。"
        }
        ws2.cell(row=curr_r, column=3, value=advice_map.get(dname, "")).alignment = align_left
        ws2.cell(row=curr_r, column=3).border = cell_border
        ws2.cell(row=curr_r, column=3).font = Font(name=font_name, size=9, color='475569')

    # 4. 全班常見錯別字盲點排行
    row_ww_start = 28
    ws2.cell(row=row_ww_start-1, column=1, value="四、 全班常見錯別字盲點 Top 清單 (課堂小考命題庫)").font = font_sub_header
    for c_i, dh in enumerate(["錯字詞語", "錯誤字修正", "語境訂正與字義解析"], 1):
        c = ws2.cell(row=row_ww_start, column=c_i, value=dh)
        c.font = font_header; c.fill = fill_header; c.border = cell_border; c.alignment = align_center
        
    all_ww = []
    for s in students:
        for item in s.get('wrong_words', []):
            if isinstance(item, dict):
                all_ww.append((item.get('word',''), item.get('correction',''), item.get('note','')))
            elif isinstance(item, (list, tuple)) and len(item)>=4:
                all_ww.append((item[0], item[2], item[3]))
                
    # 去重
    unique_ww = []
    seen_w = set()
    for w_o, w_c, w_n in all_ww:
        if w_o not in seen_w:
            seen_w.add(w_o)
            unique_ww.append((w_o, w_c, w_n))
            
    if not unique_ww:
        unique_ww = [("無", "卷面嚴謹", "全班書寫狀況良好，無明顯高頻錯字。")]
        
    for r_offset, (w_o, w_c, w_n) in enumerate(unique_ww[:6]):
        curr_r = row_ww_start + 1 + r_offset
        ws2.row_dimensions[curr_r].height = 20
        c1 = ws2.cell(row=curr_r, column=1, value=w_o); c1.alignment = align_center; c1.border = cell_border
        c1.font = Font(name=font_name, size=10, bold=True, color='B91C1C')
        
        c2 = ws2.cell(row=curr_r, column=2, value=w_c); c2.alignment = align_center; c2.border = cell_border
        c2.font = Font(name=font_name, size=10, bold=True, color='15803D')
        
        c3 = ws2.cell(row=curr_r, column=3, value=w_n); c3.alignment = align_left; c3.border = cell_border
        c3.font = Font(name=font_name, size=9, color='334155')

    # 設定 Sheet 2 欄寬
    ws2.column_dimensions['A'].width = 24
    ws2.column_dimensions['B'].width = 16
    ws2.column_dimensions['C'].width = 22
    ws2.column_dimensions['D'].width = 45
    ws2.column_dimensions['E'].width = 20
    ws2.column_dimensions['F'].width = 20
    
    # 儲存
    wb.save(output_path)
    print(f"班級學生作文成績總表 Excel 已成功產出：{output_path}")
    return output_path

if __name__ == '__main__':
    import json
    if len(sys.argv) >= 3:
        json_path = sys.argv[1]
        out_excel = sys.argv[2]
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        build_class_excel_report(data, out_excel)
    elif len(sys.argv) == 2:
        json_path = sys.argv[1]
        out_excel = os.path.splitext(json_path)[0] + ".xlsx"
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        build_class_excel_report(data, out_excel)
    else:
        print("Usage: python class_excel_generator.py <class_results.json> [output.xlsx]")

