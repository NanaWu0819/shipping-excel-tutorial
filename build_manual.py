import os
import sys

# Ensure current directory is in path
sys.path.insert(0, r"c:\shipping_tutorial")

import docx
from doc_helpers import setup_document_styles
from content_frontmatter import build_cover_page, build_executive_guide, build_toc_and_lists
from content_ch1_ch3 import build_chapter_1, build_chapter_2, build_chapter_3
from content_ch4_ch5 import build_chapter_4, build_chapter_5
from content_ch6_ch7 import build_chapter_6, build_chapter_7, build_references

def main():
    print("Initializing document...")
    doc = docx.Document()

    print("Configuring styles and margins...")
    setup_document_styles(doc)

    print("Building Cover Page...")
    build_cover_page(doc)

    print("Building Executive Guide...")
    build_executive_guide(doc)

    print("Building Table of Contents & List of Figures/Tables...")
    build_toc_and_lists(doc)

    print("Building Chapter 1 (MS Excel 操作介面全景導覽)...")
    build_chapter_1(doc)

    print("Building Chapter 2 (航運資料前處理核心常識與標準規範)...")
    build_chapter_2(doc)

    print("Building Chapter 3 (30筆航運示範資料集——問題診斷與設計)...")
    build_chapter_3(doc)

    print("Building Chapter 4 (實戰操作示範——Excel 內建選單功能篇)...")
    build_chapter_4(doc)

    print("Building Chapter 5 (實戰操作示範——函數公式清洗篇)...")
    build_chapter_5(doc)

    print("Building Chapter 6 (300筆航運案例情境實戰練習)...")
    build_chapter_6(doc)

    print("Building Chapter 7 (300筆練習題參考解題步驟與驗證指南)...")
    build_chapter_7(doc)

    print("Building References & Citations...")
    build_references(doc)

    out_file1 = r"c:\shipping_tutorial\航運行政人員_Excel基礎資料前處理操作教學手冊.docx"

    print(f"Saving document to {out_file1}...")
    doc.save(out_file1)

    print("All operations completed successfully!")
    print(f"Generated Document Size: {os.path.getsize(out_file1) / 1024:.1f} KB")

if __name__ == "__main__":
    main()
