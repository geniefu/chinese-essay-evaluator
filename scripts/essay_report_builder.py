# -*- coding: utf-8 -*-
"""
會考作文專業評閱診斷單 Word (.docx) 自動生成器 (essay_report_builder.py)
產出排版精美、可直接列印發放之標準 A4 作文評閱診斷報告卡
"""

import os
import sys
import argparse
import json
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

def set_font(run, name='標楷體', size_pt=10.0, bold=False, color_rgb=(15, 23, 42)):
    """設定文字樣式 (中文字體預設標楷體，英數字 Times New Roman)"""
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), name)
    run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
    run._element.rPr.rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(*color_rgb)

def set_cell_shd(cell, color_hex):
    """設定表格單元格背景色"""
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shd)

def set_cell_pad(cell, top=50, bottom=50, left=70, right=70):
    """設定單元格內邊距"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_borders(cell, top=None, bottom=None, left=None, right=None):
    """設定單元格框線"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="{top.get('val', 'none') if top else 'none'}" w:sz="{top.get('sz', '4') if top else '0'}" w:space="0" w:color="{top.get('color', 'auto') if top else 'auto'}"/>
            <w:left w:val="{left.get('val', 'none') if left else 'none'}" w:sz="{left.get('sz', '4') if left else '0'}" w:space="0" w:color="{left.get('color', 'auto') if left else 'auto'}"/>
            <w:bottom w:val="{bottom.get('val', 'none') if bottom else 'none'}" w:sz="{bottom.get('sz', '4') if bottom else '0'}" w:space="0" w:color="{bottom.get('color', 'auto') if bottom else 'auto'}"/>
            <w:right w:val="{right.get('val', 'none') if right else 'none'}" w:sz="{right.get('sz', '4') if right else '0'}" w:space="0" w:color="{right.get('color', 'auto') if right else 'auto'}"/>
        </w:tcBorders>
    ''')
    tcPr.append(tcBorders)

def add_callout(doc, text, title="重點提示", border_color="1E3A8A", bg_color="EFF6FF"):
    """加入精緻 Callout 提示框"""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Inches(7.1)
    set_cell_shd(cell, bg_color)
    set_cell_pad(cell, top=120, bottom=120, left=150, right=150)
    set_borders(cell, left={'val': 'single', 'sz': '24', 'color': border_color},
                      top={'val': 'none'}, bottom={'val': 'none'}, right={'val': 'none'})
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = 1.2
    r_t = p.add_run(f"【{title}】\n")
    c_rgb = (30, 58, 138) if border_color=="1E3A8A" else ((180, 83, 9) if border_color=="D97706" else (185, 28, 28))
    set_font(r_t, name='標楷體', size_pt=10.5, bold=True, color_rgb=c_rgb)
    r_c = p.add_run(text)
    set_font(r_c, name='標楷體', size_pt=9.5, bold=False, color_rgb=(51, 65, 85))

def format_heading(doc, text, level=1):
    """加入章節標題"""
    p = doc.add_paragraph()
    p.paragraph_format.keep_with_next = True
    if level == 1:
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        set_font(run, name='標楷體', size_pt=13, bold=True, color_rgb=(30, 58, 138))
    elif level == 2:
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(3)
        run = p.add_run(text)
        set_font(run, name='標楷體', size_pt=11, bold=True, color_rgb=(37, 99, 235))
    return p

def add_p(doc, text="", space_after=3, bold=False, color_rgb=(30, 41, 59)):
    """加入正文段落"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.2
    if text:
        run = p.add_run(text)
        set_font(run, name='標楷體', size_pt=9.5, bold=bold, color_rgb=color_rgb)
    return p

def build_essay_report(report_data, output_path):
    """
    依據評閱資料字典產生標準 A4 作文評閱診斷單 (.docx)
    """
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(0.40)
        s.bottom_margin = Inches(0.40)
        s.left_margin = Inches(0.45)
        s.right_margin = Inches(0.45)

    title_main = report_data.get('exam_title', "國中教育會考寫作測驗專業評閱診斷單")
    topic = report_data.get('topic', "未定題")
    student_info = report_data.get('student_info', "九年級 學生")
    score = report_data.get('score', 4)
    level_tier = report_data.get('tier', "中卷穩健 (B++)")
    one_sentence_review = report_data.get('one_sentence_review', "符合題意，但取材多停留在表面觀察，感悟流於口號化。")
    page_status = report_data.get('page_status', "未滿一頁，行數約 85%（最高上限鎖定 4 級分）")
    dimensions = report_data.get('dimensions', [])
    wrong_words = report_data.get('wrong_words', [])
    section_comments = report_data.get('section_comments', [])
    upgrade_demo = report_data.get('upgrade_demo', {})

    # 1. 頂部抬頭
    p_t = doc.add_paragraph()
    p_t.paragraph_format.space_before = Pt(2)
    p_t.paragraph_format.space_after = Pt(1)
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_t = p_t.add_run(title_main)
    set_font(r_t, name='標楷體', size_pt=16, bold=True, color_rgb=(30, 58, 138))

    # 2. 學生與試卷基本資訊卡 (Banner)
    tbl_info = doc.add_table(rows=1, cols=1)
    tbl_info.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_info.autofit = False
    c_info = tbl_info.cell(0, 0)
    c_info.width = Inches(7.3)
    set_cell_shd(c_info, "1E3A8A")
    set_cell_pad(c_info, top=45, bottom=45, left=70, right=70)
    p_inf = c_info.paragraphs[0]
    p_inf.paragraph_format.space_before = Pt(0)
    p_inf.paragraph_format.space_after = Pt(0)
    r_i1 = p_inf.add_run(f"寫作題目：〈{topic}〉 ｜ 評閱對象：{student_info} ｜ 評閱標準：心測中心會考量規 + 名師閱卷指標\n")
    set_font(r_i1, name='標楷體', size_pt=10.5, bold=True, color_rgb=(255, 255, 255))
    r_i2 = p_inf.add_run(f"卷面篇幅檢核：{page_status}")
    set_font(r_i2, name='標楷體', size_pt=9.0, bold=False, color_rgb=(224, 231, 255))

    add_p(doc, "")

    # 3. 最終評定級分與紅筆總評 (Score Banner)
    tbl_score = doc.add_table(rows=1, cols=2)
    tbl_score.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_score.autofit = False
    c_sc1 = tbl_score.cell(0, 0); c_sc1.width = Inches(2.2)
    c_sc2 = tbl_score.cell(0, 1); c_sc2.width = Inches(5.1)
    
    sc_bg = "EFF6FF" if score >= 5 else ("FEF3C7" if score == 4 else "FEF2F2")
    sc_border = "1D4ED8" if score >= 5 else ("D97706" if score == 4 else "DC2626")
    set_cell_shd(c_sc1, sc_bg); set_cell_pad(c_sc1, top=60, bottom=60, left=40, right=40)
    set_borders(c_sc1, left={'val': 'single', 'sz': '24', 'color': sc_border},
                       top={'val': 'single', 'sz': '4', 'color': 'CBD5E1'},
                       bottom={'val': 'single', 'sz': '4', 'color': 'CBD5E1'})

    p_sc = c_sc1.paragraphs[0]
    p_sc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sc_lbl = p_sc.add_run("綜合評定\n")
    set_font(r_sc_lbl, name='標楷體', size_pt=9.5, bold=False, color_rgb=(100, 116, 139))
    r_sc_val = p_sc.add_run(f"【 {score} 級分 】\n")
    set_font(r_sc_val, name='標楷體', size_pt=20, bold=True, color_rgb=(30, 58, 138) if score>=5 else ((180, 83, 9) if score==4 else (185, 28, 28)))
    r_sc_tier = p_sc.add_run(level_tier)
    set_font(r_sc_tier, name='標楷體', size_pt=9.0, bold=True, color_rgb=(71, 85, 105))

    set_cell_shd(c_sc2, "FFFFFF"); set_cell_pad(c_sc2, top=60, bottom=60, left=60, right=60)
    set_borders(c_sc2, right={'val': 'single', 'sz': '4', 'color': 'CBD5E1'},
                       top={'val': 'single', 'sz': '4', 'color': 'CBD5E1'},
                       bottom={'val': 'single', 'sz': '4', 'color': 'CBD5E1'})
    p_rev = c_sc2.paragraphs[0]
    r_rev_t = p_rev.add_run("國文名師考場紅筆總評：\n")
    set_font(r_rev_t, name='標楷體', size_pt=10, bold=True, color_rgb=(30, 58, 138))
    r_rev_c = p_rev.add_run(one_sentence_review)
    set_font(r_rev_c, name='標楷體', size_pt=9.5, bold=False, color_rgb=(51, 65, 85))

    # 4. 四大向度評量矩陣表
    format_heading(doc, "壹、 會考四大評分向度量化考核", 1)
    tbl_dim = doc.add_table(rows=1, cols=3)
    tbl_dim.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_dim.autofit = False
    dim_w = [Inches(1.5), Inches(1.2), Inches(4.6)]
    for i, h in enumerate(["評審向度", "向度等第", "核心診斷亮點與硬傷分析"]):
        c = tbl_dim.rows[0].cells[i]; c.text = h; c.width = dim_w[i]
        set_cell_shd(c, "1E3A8A"); set_cell_pad(c, top=60, bottom=60, left=50, right=50)
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font(p.runs[0], name='標楷體', size_pt=9, bold=True, color_rgb=(255, 255, 255))

    default_dims = [
        ("立意取材", f"{min(score, 4)} 級分", "取材切合題意，但多停留在表面客觀描述，缺乏獨特生命發現與深層哲理體悟。"),
        ("結構組織", f"{score} 級分", "結構大致完整，但僅舉單一事例略顯單薄，建議落實『一例一段、2~3個事例』層次展開。"),
        ("遣詞造句", f"{score} 級分", "文句通順清楚，能傳達意思，但句型稍平鋪，偶有口語化贅字。"),
        ("錯別字與標點", f"{max(score, 4)} 級分", "格式大致正確，有少數錯別字，標點符號請注意逗號與句號之適當分界。")
    ]
    target_dims = dimensions if dimensions else default_dims

    for r_i, d_row in enumerate(target_dims):
        row = tbl_dim.add_row()
        bg = "F8FAFC" if r_i % 2 == 1 else "FFFFFF"
        for c_i, val in enumerate(d_row):
            c = row.cells[c_i]; c.text = val; c.width = dim_w[c_i]
            set_cell_shd(c, bg); set_cell_pad(c, top=50, bottom=50, left=50, right=50)
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_i in [0, 1] else WD_ALIGN_PARAGRAPH.LEFT
            f_color = (30, 58, 138) if c_i == 0 else ((180, 83, 9) if c_i == 1 else (15, 23, 42))
            set_font(p.runs[0], name='標楷體', size_pt=8.5, bold=(c_i in [0, 1]), color_rgb=f_color)
            set_borders(c, bottom={'val': 'single', 'sz': '4', 'color': 'CBD5E1'}, top={'val': 'single', 'sz': '4', 'color': 'CBD5E1'})

    # 5. 考場天頭錯別字訂正格
    format_heading(doc, "貳、 考場天頭錯別字回寫訂正格", 1)
    if wrong_words:
        tbl_ww = doc.add_table(rows=1, cols=4)
        tbl_ww.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_ww.autofit = False
        ww_w = [Inches(1.8), Inches(1.8), Inches(1.8), Inches(1.8)]
        for i, h in enumerate(["錯字回寫欄 ①", "錯字回寫欄 ②", "錯字回寫欄 ③", "錯字回寫欄 ④"]):
            c = tbl_ww.rows[0].cells[i]; c.text = h; c.width = ww_w[i]
            set_cell_shd(c, "334155"); set_cell_pad(c, top=50, bottom=50, left=40, right=40)
            p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_font(p.runs[0], name='標楷體', size_pt=8.5, bold=True, color_rgb=(255, 255, 255))
        
        row_w = tbl_ww.add_row()
        for idx_w in range(4):
            c = row_w.cells[idx_w]; c.width = ww_w[idx_w]
            set_cell_shd(c, "FEF2F2"); set_cell_pad(c, top=60, bottom=60, left=40, right=40)
            p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if idx_w < len(wrong_words):
                w_orig, w_corr = wrong_words[idx_w]
                r_txt = p.add_run(f"「{w_orig}」 口 ➔ 【 {w_corr} 】")
                set_font(r_txt, name='標楷體', size_pt=9.0, bold=True, color_rgb=(185, 28, 28))
            else:
                r_txt = p.add_run("---")
                set_font(r_txt, name='標楷體', size_pt=9.0, bold=False, color_rgb=(148, 163, 184))
            set_borders(c, bottom={'val': 'single', 'sz': '4', 'color': 'CBD5E1'}, top={'val': 'single', 'sz': '4', 'color': 'CBD5E1'})
    else:
        add_p(doc, "太棒了！本次卷面書寫極為嚴謹，無明顯錯別字，請繼續保持高水準之文字正確性！", bold=True, color_rgb=(21, 128, 61))

    # 6. 逐段眉批與精準改進方針
    format_heading(doc, "參、 國文老師逐段批註與脈絡診斷", 1)
    default_secs = [
        ("第一段（起 / 開門破題）", "破題清楚，能點出核心主旨，但引導鋪陳稍急躁，建議加強生活情境之帶入感。"),
        ("第二段（承 / 第一事例展開）", "描寫具有生活細節，能具體呈現人事事物，唯在心理轉折處收尾較倉促。"),
        ("第三段（轉 / 第二事例深入）", "事例與第二段性質略有重複，建議採對比視角，深化『由物入理』之哲思轉化。"),
        ("第四段（合 / 總結回扣昇華）", "有總結意識，但感悟稍流於口號化，結論段必須實質回扣前文事例，形成首尾呼應。")
    ]
    target_secs = section_comments if section_comments else default_secs
    for s_title, s_comment in target_secs:
        add_p(doc, f"• 【{s_title}】：{s_comment}", space_after=3)

    # 7. 段落升級改寫示範
    format_heading(doc, "肆、 突破瓶頸：段落升級改寫示範（衝刺 5~6 級分）", 1)
    orig_p = upgrade_demo.get('orig', "當我停下來的時候，我發現世界其實很美麗。以前我每天都走得很匆忙，沒有好好看路上的風景。現在停下腳步，我感受到陽光很溫暖，微風吹過臉龐很舒服，我學會了不要給自己太大壓力。")
    up_p = upgrade_demo.get('upgrade', "猛然煞住趕赴補習班的腳步，氣喘吁吁地倚在紅磚巷口。褪去耳機裡的喧囂，一陣規律而清脆的鋼鐵敲擊聲悄然鑽入耳際。那是巷口修鞋的老伯，正赤著黝黑的臂膀，在微塵浮動的光暈裡專注穿引麻線。粗糙龜裂的手指一次次沉穩穿透厚硬的皮掌，不慌不忙。那一刻我驀地驚覺，自己日復一日在考卷與分數間的倉皇奔逐，究竟耗損了多少專注的魂魄？唯有停下盲奔，方能讀懂在時光沉澱下的真實生命。")
    tips = upgrade_demo.get('tips', "【改寫要訣點撥】：將原先『陽光溫暖、微風舒服』等抽象空泛雞湯，轉化為具體客觀畫面（修鞋老伯、穿引麻線、光暈微塵），完成『客觀畫面特寫 ➔ 人物專注同理 ➔ 自我心靈省思』之三層昇華！")

    tbl_up = doc.add_table(rows=2, cols=1)
    tbl_up.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl_up.autofit = False
    c_u1 = tbl_up.cell(0, 0); c_u1.width = Inches(7.1)
    set_cell_shd(c_u1, "FEF2F2"); set_cell_pad(c_u1, top=50, bottom=50, left=60, right=60)
    set_borders(c_u1, left={'val': 'single', 'sz': '16', 'color': 'DC2626'})
    p_u1 = c_u1.paragraphs[0]; p_u1.paragraph_format.space_before = Pt(0); p_u1.paragraph_format.space_after = Pt(0)
    r_u1_t = p_u1.add_run("【學生原始段落（4 級分常見通病：抽象抒情多、無客觀細節）】：\n")
    set_font(r_u1_t, name='標楷體', size_pt=9.0, bold=True, color_rgb=(185, 28, 28))
    r_u1_c = p_u1.add_run(orig_p)
    set_font(r_u1_c, name='標楷體', size_pt=8.5, bold=False, color_rgb=(153, 27, 27))

    c_u2 = tbl_up.cell(1, 0); c_u2.width = Inches(7.1)
    set_cell_shd(c_u2, "EFF6FF"); set_cell_pad(c_u2, top=50, bottom=50, left=60, right=60)
    set_borders(c_u2, left={'val': 'single', 'sz': '16', 'color': '2563EB'})
    p_u2 = c_u2.paragraphs[0]; p_u2.paragraph_format.space_before = Pt(0); p_u2.paragraph_format.space_after = Pt(0)
    r_u2_t = p_u2.add_run("【名師升級示範（5~6 級分典範：注入生活特寫、由物入理）】：\n")
    set_font(r_u2_t, name='標楷體', size_pt=9.0, bold=True, color_rgb=(30, 58, 138))
    r_u2_c = p_u2.add_run(up_p)
    set_font(r_u2_c, name='標楷體', size_pt=8.5, bold=False, color_rgb=(30, 41, 59))

    add_p(doc, tips, bold=True, color_rgb=(30, 58, 138))

    doc.save(output_path)
    print(f"作文評閱診斷單已成功產出：{output_path}")
    return output_path

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="國中會考作文專業評閱診斷單 Word 生成工具 (essay_report_builder)")
    parser.add_argument('--topic', default="當我停下來", help="寫作題目")
    parser.add_argument('--student', default="九年級 學生", help="學生座號姓名或年級")
    parser.add_argument('--score', type=int, default=4, help="評定級分 (0~6)")
    parser.add_argument('--tier', default="中卷穩健 (B++)", help="等第類別")
    parser.add_argument('--output', default="會考作文評閱診斷單.docx", help="輸出 Word 檔案路徑")
    parser.add_argument('--data-json', default=None, help="傳入包含完整評審診斷之 JSON 檔案")
    
    args = parser.parse_args()
    
    data = {
        'topic': args.topic,
        'student_info': args.student,
        'score': args.score,
        'tier': args.tier
    }
    if args.data_json and os.path.exists(args.data_json):
        with open(args.data_json, 'r', encoding='utf-8') as f:
            data.update(json.load(f))
            
    build_essay_report(data, args.output)
