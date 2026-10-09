# -*- coding: utf-8 -*-
"""
手寫作文掃描 PDF / 圖片批次萃取與辨識預處理工具 (pdf_essay_extractor.py)
利用 PyMuPDF (fitz) 讀取多頁學生手寫稿卷 PDF，自動拆解並提取頁面圖片以利批次評閱。
"""

import os
import sys
import argparse

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

def extract_pdf_pages_to_images(pdf_path, output_dir=None, zoom=2.0):
    """
    將 PDF 各頁轉換為高解析度圖片供 AI 辨識手寫文章
    :param pdf_path: PDF 檔案路徑
    :param output_dir: 暫存輸出資料夾
    :param zoom: 縮放比率 (2.0 = 約 150~200 DPI，適合手寫字跡辨識)
    :return: 產出的圖片路徑清單
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
    print(f"📖 讀取 PDF 成功：共 {total_pages} 頁")

    image_paths = []
    mat = fitz.Matrix(zoom, zoom)

    for page_idx in range(total_pages):
        page = doc[page_idx]
        pix = page.get_pixmap(matrix=mat)
        img_filename = f"page_{page_idx + 1:03d}.png"
        img_path = os.path.join(output_dir, img_filename)
        pix.save(img_path)
        image_paths.append(img_path)

    doc.close()
    print(f"✅ 已成功將 {total_pages} 頁轉換為圖片，存放於：{output_dir}")
    return image_paths

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="PDF 稿紙批次圖片萃取工具")
    parser.add_argument('pdf_path', help="PDF 檔案路徑")
    parser.add_argument('--output-dir', default=None, help="圖片輸出資料夾")
    parser.add_argument('--zoom', type=float, default=2.0, help="縮放係數 (預設 2.0)")
    
    args = parser.parse_args()
    extract_pdf_pages_to_images(args.pdf_path, args.output_dir, args.zoom)
