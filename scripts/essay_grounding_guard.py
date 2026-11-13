# -*- coding: utf-8 -*-
"""
essay_grounding_guard.py
中文作文防虛構、防空白、防文字污染「四大剛性防禦機制」核心檢核模組
1. 實體墨水像素比例閥值檢驗（物理防空白硬攔截）：採用面積自適應相對比例，免疫不同掃描 DPI 差異。
2. 座標式防污染裁切（Bounding Box Masking）：支援多版型遮罩試卷印刷題目、引導說明與評分表格。
3. 原文錨定比對（Text Grounding Assert）：成語採嚴格比對、錯字採上下文脈絡比對，徹底杜絕幻覺評語。
4. 全班錯字常模統計防禦閘門（Statistical Typo Gate）：強制檢核錯字檢出率與顯式證明，杜絕流水線惰性。
"""

import os
import re
import sys
import numpy as np
from PIL import Image

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')


# -------------------------------------------------------------
# 防禦機制一：實體墨水像素比例閥值檢驗（物理防空白硬攔截）
# -------------------------------------------------------------
def analyze_ink_density(image_path_or_img, page_type='p1', layout='cap_standard'):
    """
    計算作答格線區域之實質手寫墨水像素比例（面積自適應，免疫不同 DPI 影響）。
    :param image_path_or_img: 圖片路徑或 PIL Image
    :param page_type: 'p1' (第一頁) 或 'p2' (第二頁)
    :param layout: 'cap_standard' (會考雙欄題本稿紙) 或 'plain_grid' (純稿紙)
    :return: dict 包含 is_blank, ink_ratio, net_ink_ratio, net_ink
    """
    if isinstance(image_path_or_img, str):
        img = Image.open(image_path_or_img).convert('L')
    else:
        img = image_path_or_img.convert('L')

    arr = np.array(img)
    h, w = arr.shape

    if layout == 'plain_grid':
        grid = arr[int(h * 0.10):int(h * 0.95), int(w * 0.05):int(w * 0.95)]
        baseline_ratio = 0.025  # 純稿紙空框線約佔 2.5%
    else:
        # 標準直式會考答案紙座標規範：
        # 第一頁：右上 30% 為題目與引導說明，格線作答區位於 x: 5% ~ 70%, y: 15% ~ 95%
        # 第二頁：左側 20% 為評審標準欄，格線作答區位於 x: 20% ~ 95%, y: 15% ~ 95%
        if page_type == 'p1':
            grid = arr[int(h * 0.15):int(h * 0.95), int(w * 0.05):int(w * 0.70)]
            baseline_ratio = 0.028
        else:
            grid = arr[int(h * 0.15):int(h * 0.95), int(w * 0.20):int(w * 0.95)]
            baseline_ratio = 0.025

    total_pixels = max(1, grid.size)
    # 深色墨水閥值 (< 140)
    ink_count = int(np.sum(grid < 140))
    ink_ratio = ink_count / total_pixels
    net_ink_ratio = max(0.0, ink_ratio - baseline_ratio)

    # 換算為等效 200 DPI 淨像素數以維持統計可讀性
    net_ink_equiv = int(net_ink_ratio * total_pixels)

    # 淨墨水比例小於 0.45%（約少於 5 個完整漢字筆劃）判定為空白
    is_blank = (net_ink_ratio < 0.0045)

    return {
        'page_type': page_type,
        'layout': layout,
        'ink_count': ink_count,
        'ink_ratio': round(ink_ratio, 5),
        'net_ink_ratio': round(net_ink_ratio, 5),
        'net_ink': net_ink_equiv,
        'is_blank': is_blank
    }


def check_student_blank(p1_path, p2_path=None, layout='cap_standard'):
    """
    整合兩頁判斷學生答卷是否為完全空白卷
    """
    res1 = analyze_ink_density(p1_path, 'p1', layout=layout)
    res2 = (
        analyze_ink_density(p2_path, 'p2', layout=layout)
        if (p2_path and os.path.exists(p2_path))
        else {'is_blank': True, 'net_ink': 0, 'net_ink_ratio': 0.0}
    )

    is_totally_blank = res1['is_blank'] and res2['is_blank']
    return {
        'is_blank': is_totally_blank,
        'p1_net_ink': res1['net_ink'],
        'p1_net_ink_ratio': res1['net_ink_ratio'],
        'p2_net_ink': res2['net_ink'],
        'p2_net_ink_ratio': res2.get('net_ink_ratio', 0.0)
    }


def build_blank_record(class_name, seat_no, name, topic="未定題"):
    """
    當觸發物理防空白硬攔截時，直接生成官方標準 0 級分紀錄，嚴禁呼叫 LLM 虛構內容。
    """
    st_str = f"{int(seat_no):02d}" if str(seat_no).isdigit() else str(seat_no)
    return {
        'class_name': class_name,
        'seat_no': st_str,
        'paper_seat_no': st_str,
        'name': name,
        'student_info': f"{class_name} {st_str}號 {name}",
        'topic': topic,
        'score': 0,
        'tier': '0 級分（空白卷）',
        'page_status': '卷面完全空白（未作答）',
        'one_sentence_review': '【空白卷／未作答】答案紙除個人基本資料外，全卷格線內完全空白，未書寫任何字句。依國中教育會考評分規準，核定為「0 級分」。',
        'dimensions': [
            ['立意取材', '0 級分', '【空白卷】全卷未書寫任何文字，未成篇，無法判定立意取材。'],
            ['結構組織', '0 級分', '【空白卷】篇幅 0 行，完全未作答，無篇章結構可言。'],
            ['遣詞造句', '0 級分', '【空白卷】無任何文字與修辭表達。'],
            ['錯別字與標點', '0 級分', '【空白卷】無書寫內容可供評定。']
        ],
        'wrong_words': [],
        'idioms': [],
        'section_comments': [
            ("全卷作答狀態", "【完全空白】答案紙第一頁與第二頁之寫作格線內均未書寫任何字句，全卷無作答內容。")
        ],
        'scaffold': {
            'detail_guidance': '先求有再求好：下階段考試第一步，強迫自己在前 5 分鐘內動筆寫出第 1 句話。',
            'senses_guidance': '由最熟悉的感官著手：寫一寫今天走進考場的溫度、聽到的鐘聲，或早晨喝到熱豆漿的味道。',
            'depth_reflection': '會考計分關鍵：在免試入學比序中，作文 0 級分將嚴重失分；哪怕只寫三行，也能取得基本級分！'
        },
        'upgrade_demo': {
            'orig': '（卷面完全空白，未書寫任何文字）',
            'upgrade': '猛然煞住趕赴補習班的腳步，氣喘吁吁地倚在紅磚巷口。褪去耳機裡的喧囂，一陣規律而清脆的鋼鐵敲擊聲悄然鑽入耳際。那是巷口修鞋的老伯，正赤著黝黑的臂膀，在微塵浮動的光暈裡專注穿引麻線。粗糙龜裂的手指一次次沉穩穿透厚硬的皮掌，不慌不忙。那一刻我驀地驚覺，自己日復一日在考卷與分數間的倉皇奔逐，究竟耗損了多少專注的魂魄？唯有停下盲奔，方能讀懂在時光沉澱下的真實生命。',
            'tips': '【突破零分第一步】：只要動筆寫下身邊觀察到的任一真實生活畫面，就能打破 0 級分僵局，跨越至 2~3 級分！'
        }
    }


# -------------------------------------------------------------
# 防禦機制二：座標式防污染裁切（Bounding Box Masking）
# -------------------------------------------------------------
def get_masked_writing_area(image_path_or_img, page_type='p1', layout='cap_standard'):
    """
    對試卷印刷區域進行純白遮罩，回傳純學生作答區域影像，杜絕 OCR 讀取印刷題目與表格。
    :param layout: 'cap_standard' (會考標準雙欄), 'plain_grid' (純稿紙僅遮頂欄), 'none' (不遮罩)
    """
    if isinstance(image_path_or_img, str):
        img = Image.open(image_path_or_img).convert('RGB')
    else:
        img = image_path_or_img.convert('RGB')

    if layout == 'none':
        return img

    arr = np.array(img)
    h, w, _ = arr.shape
    masked = arr.copy()

    if layout == 'plain_grid':
        # 純稿紙僅遮罩頂部 10% 欄頭
        masked[:int(h * 0.10), :, :] = 255
    else:
        if page_type == 'p1':
            # 遮罩右側 30% 寬度（印刷題目與引導說明）
            masked[:, int(w * 0.70):, :] = 255
            # 遮罩頂部 15% 欄頭（校名、答案紙名稱）
            masked[:int(h * 0.15), :, :] = 255
        else:
            # 遮罩左側 20% 寬度（評審標準表）
            masked[:, :int(w * 0.20), :] = 255
            # 遮罩頂部 15% 欄頭
            masked[:int(h * 0.15), :, :] = 255

    return Image.fromarray(masked)


# -------------------------------------------------------------
# 防禦機制三：原文錨定比對（Text Grounding Assert）
# -------------------------------------------------------------
def _clean_text_for_match(text):
    """移除空白與常見標點符號，保留實質漢字與英數以進行穩健比對"""
    return re.sub(r'[\s\u3000，。、；：「」『』（）！？…\.\,\!\?\-\"\']+', '', str(text))


def verify_text_grounding(student_raw_text, item_list, item_type='idiom'):
    """
    比對評審輸出之詞句是否真實存在於學生手寫文本中。
    - 'idiom' (成語)：嚴格比對完整成語（或至少 3 字連續重疊），杜絕拆分拼湊的幻覺。
    - 'wrong_word' (錯別字)：因 OCR 可能將錯字自動糾正，改採「前後文語境片段比對」，
      避免誤殺真實錯字，同時攔截完全虛構的整句引文。
    """
    if not student_raw_text or not item_list:
        return []

    cleaned_text = _clean_text_for_match(student_raw_text)
    verified = []

    for item in item_list:
        word = ""
        context = ""
        if isinstance(item, dict):
            word = item.get('idiom', item.get('word', ''))
            context = item.get('context', item.get('location', ''))
        elif isinstance(item, (list, tuple)) and len(item) >= 1:
            word = str(item[0])
            context = str(item[1]) if len(item) >= 2 else ""

        word_clean = _clean_text_for_match(word)
        passed = False

        if item_type == 'idiom':
            # 成語採嚴格比對：全文出現完整成語，或 4 字以上成語有連續 3 字相符（容許單一字手寫辨識誤差）
            if word_clean:
                if word_clean in cleaned_text:
                    passed = True
                elif len(word_clean) >= 4:
                    for i in range(len(word_clean) - 2):
                        if word_clean[i:i + 3] in cleaned_text:
                            passed = True
                            break
        else:
            # 錯別字比對：若錯字本身直接命中原文，直接通過
            if word_clean and word_clean in cleaned_text:
                passed = True
            elif context:
                # 剝除「第 X 段：」前綴與省略號，比對上下文相鄰短語（長度 >= 3 的連續片段）
                ctx_body = re.sub(r'^第\s*\d+\s*段[：:]?\s*', '', str(context))
                ctx_clean = _clean_text_for_match(ctx_body)
                if word_clean and word_clean in ctx_clean:
                    # 切成錯字前後的語境片段
                    parts = [p for p in ctx_clean.split(word_clean) if len(p) >= 2]
                    if any(p in cleaned_text for p in parts):
                        passed = True
                elif len(ctx_clean) >= 4:
                    # 滑動檢查是否有至少連續 4 字的上下文出現在原文中
                    for i in range(len(ctx_clean) - 3):
                        if ctx_clean[i:i + 4] in cleaned_text:
                            passed = True
                            break

        if passed:
            verified.append(item)
        else:
            print(f"⚠️ [防虛構攔截] 發現未通過原文錨定驗證之項目：{word}（類型：{item_type}），已自動剔除！")

    return verified


# -------------------------------------------------------------
# 防禦機制四：全班錯字常模統計防禦閘門（Statistical Typo Gate）
# -------------------------------------------------------------
def verify_class_typo_statistical_gate(students, min_ratio=0.30, min_total_typos=10):
    """
    全班錯字合理性常模統計檢驗：
    在全班性（10人以上）作文評閱中，檢核錯字檢出率是否符合教學現場常模。
    若檢出率過低或錯字總數偏離常態，強制中斷流水線，杜絕「使用空清單預設值跳過微觀手寫排查」之重大疏漏。
    """
    non_blank_students = [s for s in students if s.get('score', 0) > 0]
    total_students = len(non_blank_students)
    if total_students < 10:
        return {'passes_gate': True, 'msg': '實評學生人數少於 10 人，略過常模統計防禦閘門。'}

    students_with_typos = 0
    total_typos = 0
    lazy_blank_defaults = []

    for s in non_blank_students:
        seat = s.get('seat_no', s.get('seat', '?'))
        name = s.get('name', '未知')
        score = s.get('score', 0)
        ww = s.get('wrong_words', [])

        if len(ww) > 0:
            students_with_typos += 1
            total_typos += len(ww)
        else:
            justification = s.get('zero_typo_justification', '')
            if score < 5 and not justification:
                lazy_blank_defaults.append(f"座號 {seat} {name} ({score}級分)")

    ratio = students_with_typos / total_students if total_students > 0 else 0
    passes = (ratio >= min_ratio and total_typos >= min_total_typos and len(lazy_blank_defaults) == 0)

    error_msg = ""
    if not passes:
        reasons = []
        if ratio < min_ratio:
            reasons.append(f"全班錯字檢出率僅 {ratio*100:.1f}%（低於常模安全門檻 {min_ratio*100:.0f}%）")
        if total_typos < min_total_typos:
            reasons.append(f"全班錯字總數僅 {total_typos} 處（低於常態門檻 {min_total_typos} 處）")
        if lazy_blank_defaults:
            reasons.append(f"發現未經實質手寫審查之空白佔位符名單：{', '.join(lazy_blank_defaults[:5])}...")
        error_msg = (
            "【🛑 觸發第四大剛性防禦：錯字常模統計防禦閘門攔截！】\n"
            + "\n".join(f"- {r}" for r in reasons)
            + "\n系統判定涉嫌批閱流水線跳過微觀手寫字元級排查，強制拒絕生成發行報告！"
        )

    return {
        'passes_gate': passes,
        'total_students': total_students,
        'students_with_typos': students_with_typos,
        'total_typos': total_typos,
        'ratio': ratio,
        'error_msg': error_msg
    }
