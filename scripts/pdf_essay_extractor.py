# -*- coding: utf-8 -*-
"""
手寫作文掃描 PDF / 圖片批次萃取與辨識預處理工具 (pdf_essay_extractor.py)
利用 PyMuPDF (fitz) 讀取多頁學生手寫稿卷 PDF，自動拆解並提取頁面圖片，
並整合 essay_grounding_guard 執行「物理防空白檢測」與「座標式防文字污染遮罩」。
"""

import os
import sys
import json
import argparse

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    from essay_grounding_guard import check_student_blank, get_masked_writing_area
except ImportError:
    # 支援從不同工作目錄以 module 執行
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if script_dir not in sys.path:
        sys.path.insert(0, script_dir)
    try:
        from essay_grounding_guard import check_student_blank, get_masked_writing_area
    except ImportError:
        check_student_blank = None
        get_masked_writing_area = None


def parse_target_pages(target_str, total_pages):
    """
    解析指定旋轉的頁碼（1-based），支援：
    - None 或 'all': 全部頁面
    - 'even': 所有偶數頁 (2, 4, 6...)，常見於雙面掃描背面顛倒
    - 'odd': 所有奇數頁 (1, 3, 5...)
    - '2,4,7-9': 指定特定頁碼與區間
    - list 或 set: 直接指定頁碼集合
    """
    if not target_str or target_str == 'all':
        return set(range(1, total_pages + 1))
    if isinstance(target_str, (list, tuple, set)):
        return set(int(x) for x in target_str)

    target_str = str(target_str).strip().lower()
    if target_str == 'even':
        return set(p for p in range(1, total_pages + 1) if p % 2 == 0)
    if target_str == 'odd':
        return set(p for p in range(1, total_pages + 1) if p % 2 == 1)

    pages = set()
    parts = target_str.replace(' ', '').split(',')
    for part in parts:
        if '-' in part:
            try:
                start, end = part.split('-', 1)
                pages.update(range(int(start), int(end) + 1))
            except ValueError:
                pass
        elif part.isdigit():
            pages.add(int(part))
    return pages


def extract_pdf_pages_to_images(
    pdf_path,
    output_dir=None,
    zoom=2.0,
    rotate=0,
    target_pages=None,
    layout='cap_standard',
    apply_mask=True,
    check_blank=True,
):
    """
    將 PDF 各頁轉換為高解析度圖片供 AI 辨識手寫文章，支援「指定特定頁面」旋轉轉正、
    物理空白攔截檢測與座標式印刷區遮罩。
    :return: dict {'images': [...], 'masked_images': [...], 'blank_check': {...}}
    """
    if not fitz:
        raise ImportError("請先安裝 PyMuPDF：pip install PyMuPDF")

    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"找不到 PDF 檔案：{pdf_path}")

    if not output_dir:
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_dir = os.path.join(os.path.dirname(pdf_path), f"{base_name}_pages")

    os.makedirs(output_dir, exist_ok=True)

    doc = fitz.open(pdf_path)
    total_pages = len(doc)
    target_set = parse_target_pages(target_pages, total_pages) if rotate != 0 else set()

    rotated_info = f"（第 {sorted(list(target_set))} 頁旋轉 {rotate}°）" if target_set else "（全部保持原方向）"
    print(f"📖 讀取 PDF 成功：共 {total_pages} 頁 {rotated_info}")

    image_paths = []
    masked_paths = []

    for page_idx in range(total_pages):
        page_num = page_idx + 1
        page = doc[page_idx]
        deg = rotate if (page_num in target_set) else 0
        mat = fitz.Matrix(zoom, zoom).prerotate(deg)

        pix = page.get_pixmap(matrix=mat)
        img_filename = f"page_{page_num:03d}.png"
        img_path = os.path.join(output_dir, img_filename)
        pix.save(img_path)
        image_paths.append(img_path)

        if apply_mask and layout != 'none' and get_masked_writing_area is not None:
            page_type = 'p1' if (page_num % 2 == 1) else 'p2'
            masked_img = get_masked_writing_area(img_path, page_type=page_type, layout=layout)
            masked_path = os.path.join(output_dir, f"page_{page_num:03d}_masked.png")
            masked_img.save(masked_path)
            masked_paths.append(masked_path)

    doc.close()
    print(f"✅ 已成功將 {total_pages} 頁轉換為圖片，存放於：{output_dir}")
    if masked_paths:
        print(f"🛡️ 已套用座標防污染遮罩 ({layout})，產出 {len(masked_paths)} 張純作答區圖檔 (*_masked.png)")

    blank_res = None
    if check_blank and check_student_blank is not None and len(image_paths) >= 1:
        p1 = image_paths[0]
        p2 = image_paths[1] if len(image_paths) >= 2 else None
        blank_res = check_student_blank(p1, p2, layout=layout if layout != 'none' else 'cap_standard')
        if blank_res['is_blank']:
            print(f"🛑 [物理防空白硬攔截] 偵測到空白卷！P1淨墨水比={blank_res['p1_net_ink_ratio']:.4%}, P2淨墨水比={blank_res['p2_net_ink_ratio']:.4%}，嚴禁送入 LLM！")
        else:
            print(f"✍️ [墨水檢核通過] P1淨墨水比={blank_res['p1_net_ink_ratio']:.4%}, P2淨墨水比={blank_res['p2_net_ink_ratio']:.4%}")

    meta = {
        'pdf_path': pdf_path,
        'output_dir': output_dir,
        'total_pages': total_pages,
        'images': image_paths,
        'masked_images': masked_paths,
        'blank_check': blank_res,
    }
    meta_path = os.path.join(output_dir, "_extract_meta.json")
    with open(meta_path, 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

    return meta


def rotate_pdf_file(pdf_path, output_pdf_path=None, rotate=180, target_pages=None):
    """
    直接將 PDF 實體檔案中「指定的頁面」旋轉指定角度並另存新檔（其餘頁面不動）
    """
    if not fitz:
        raise ImportError("請先安裝 PyMuPDF：pip install PyMuPDF")

    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"找不到 PDF 檔案：{pdf_path}")

    if not output_pdf_path:
        base, ext = os.path.splitext(pdf_path)
        output_pdf_path = f"{base}_已轉正{ext}"

    doc = fitz.open(pdf_path)
    total_pages = len(doc)
    target_set = parse_target_pages(target_pages, total_pages)

    for page_idx in range(total_pages):
        page_num = page_idx + 1
        if page_num in target_set:
            page = doc[page_idx]
            page.set_rotation((page.rotation + rotate) % 360)

    doc.save(output_pdf_path)
    doc.close()
    print(f"🔄 PDF 第 {sorted(list(target_set))} 頁已成功旋轉 {rotate}° 並另存為：{output_pdf_path}")
    return output_pdf_path


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="PDF 稿紙批次圖片萃取與局部轉正工具")
    parser.add_argument('pdf_path', help="PDF 檔案路徑")
    parser.add_argument('--output-dir', default=None, help="圖片輸出資料夾")
    parser.add_argument('--zoom', type=float, default=2.0, help="縮放係數 (預設 2.0)")
    parser.add_argument('--rotate', type=int, default=0, choices=[0, 90, 180, 270], help="旋轉角度 (例如 180 為上下顛倒轉正)")
    parser.add_argument('--pages', default=None, help="指定需旋轉的頁碼，支援 'even'(偶數頁)、'odd'(奇數頁)、'2,4,7-9' 或留空(全部)")
    parser.add_argument('--save-rotated-pdf', action='store_true', help="是否額外另存已轉正的 PDF 檔案")
    parser.add_argument('--layout', default='cap_standard', choices=['cap_standard', 'plain_grid', 'none'], help="稿紙遮罩版型 (預設 cap_standard)")
    parser.add_argument('--no-mask', action='store_true', help="停用座標防污染遮罩")
    parser.add_argument('--no-blank-check', action='store_true', help="停用物理防空白檢測")

    args = parser.parse_args()

    if args.save_rotated_pdf and args.rotate != 0:
        rotate_pdf_file(args.pdf_path, rotate=args.rotate, target_pages=args.pages)

    extract_pdf_pages_to_images(
        args.pdf_path,
        args.output_dir,
        args.zoom,
        args.rotate,
        target_pages=args.pages,
        layout=args.layout,
        apply_mask=(not args.no_mask),
        check_blank=(not args.no_blank_check),
    )
