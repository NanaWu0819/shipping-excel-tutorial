from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from doc_helpers import (
    NAVY_HEX, TEAL_HEX, COLOR_NAVY, COLOR_TEAL, COLOR_BODY, COLOR_MUTED,
    add_body_p, add_callout
)

def build_cover_page(doc):
    # Top spacing
    p_top = doc.add_paragraph()
    p_top.paragraph_format.space_before = Pt(36)
    p_top.paragraph_format.space_after = Pt(12)

    # Organization / Series Header
    p_org = doc.add_paragraph()
    r_org = p_org.add_run("全球海運物流暨供應鏈數位賦能系列教材  |  實戰技術指引")
    r_org.font.name = "微軟正黑體"
    r_org.font.size = Pt(11)
    r_org.font.bold = True
    r_org.font.color.rgb = COLOR_TEAL
    p_org.paragraph_format.space_after = Pt(20)

    # Main Title
    p_title = doc.add_paragraph()
    r_title = p_title.add_run("航運行政人員必修：\nMS Excel 基礎資料分析前處理\n實戰電腦操作教學手冊")
    r_title.font.name = "微軟正黑體"
    r_title.font.size = Pt(26)
    r_title.font.bold = True
    r_title.font.color.rgb = COLOR_NAVY
    p_title.paragraph_format.line_spacing = 1.2
    p_title.paragraph_format.space_after = Pt(16)

    # Accent line
    p_line = doc.add_paragraph()
    pBdr = parse_xml(r'<w:pBdr %s><w:bottom w:val="single" w:sz="36" w:space="8" w:color="1B365D"/></w:pBdr>' % nsdecls('w'))
    p_line._p.get_or_add_pPr().append(pBdr)
    p_line.paragraph_format.space_after = Pt(20)

    # Subtitle
    p_sub = doc.add_paragraph()
    r_sub = p_sub.add_run("從零打底！航運單據資料清理、格式標準化、重複排查與分析前置技能全攻略\n聚焦於海運提單、貨櫃清單、港口代碼與費用數據清洗的核心操作")
    r_sub.font.name = "微軟正黑體"
    r_sub.font.size = Pt(13)
    r_sub.font.color.rgb = COLOR_BODY
    p_sub.paragraph_format.line_spacing = 1.3
    p_sub.paragraph_format.space_after = Pt(80)

    # Metadata Card at Bottom
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.line_spacing = 1.3
    
    meta_items = [
        ("適用對象：", "航運公司、船務代理、國際貨代（Freight Forwarder）、報關行之 Excel 初學行政同仁"),
        ("教學工具：", "Microsoft Excel 2016 / 2019 / 2021 / Microsoft 365 繁體中文版"),
        ("示範案例：", "30 筆航運進出口實務示範檔 ＆ 300 筆多情境海運案例挑戰題庫"),
        ("手冊版本：", "第 1.0 版（專業標準實戰發行版）"),
        ("發行日期：", "2026 年 11 月"),
        ("編撰單位：", "航運數位化轉型與商業智慧數據推動小組")
    ]
    for label, val in meta_items:
        r_l = p_meta.add_run(f"{label} ")
        r_l.font.name = "微軟正黑體"
        r_l.font.bold = True
        r_l.font.size = Pt(9.5)
        r_l.font.color.rgb = COLOR_NAVY
        
        r_v = p_meta.add_run(f"{val}\n")
        r_v.font.name = "微軟正黑體"
        r_v.font.size = Pt(9.5)
        r_v.font.color.rgb = COLOR_BODY

def build_executive_guide(doc):
    doc.add_page_break()
    p_h = doc.add_paragraph()
    r_h = p_h.add_run("導讀與手冊使用指南")
    r_h.font.name = "微軟正黑體"
    r_h.font.size = Pt(18)
    r_h.font.bold = True
    r_h.font.color.rgb = COLOR_NAVY
    pBdr = parse_xml(r'<w:pBdr %s><w:bottom w:val="single" w:sz="18" w:space="4" w:color="1B365D"/></w:pBdr>' % nsdecls('w'))
    p_h._p.get_or_add_pPr().append(pBdr)
    p_h.paragraph_format.space_after = Pt(12)

    add_body_p(doc, "在現代國際海運物流與船務代理業務中，行政與文件人員每日必須面對由不同航運公司（Carrier）、貨運承攬業（Forwarder）、海關 EDI 通關系統、貨櫃碼頭集散站（CY/CFS）以及海外代理所拋轉而來的龐雜單據與資料檔。然而，這些跨系統匯入的資料，往往充斥著前後多餘空白、全半形混雜、港口名稱不統一、船名航次黏合、重量運費帶有文字單位、甚至夾帶重複提單等「髒資料（Dirty Data）」。")
    
    add_body_p(doc, "若未經嚴謹的「資料前處理（Data Preprocessing）」，直接將這些資料投入統計加總、製作樞紐分析表或對接商業智慧（BI）報表，極易導致運費計算錯誤、報關數據異常、甚至是重複放行貨櫃的嚴重航運疏失。這就是資料科學界著名的鐵律：「Garbage In, Garbage Out（垃圾進，垃圾出）」。")

    add_callout(doc, "本手冊的核心學習宗旨", 
                "本教材絕非泛泛而談的 Excel 通用功能說明書，而是專門針對『航運行政新手』量身打造的電腦操作實戰手冊。手冊聚焦於基礎資料分析中最關鍵的 20% 前處理基本功——包括 Excel 介面掌握、資料清洗常識、內建指令工具、關鍵清理函數、30 筆真實航運資料逐步演練，以及 300 筆多情境綜合實戰練習。只要按部就班演練，即可在最短時間內具備專業的資料前處理操作實戰能力！", 
                "concept")

    add_body_p(doc, "為了讓學習成效最大化，本教學手冊特別附帶兩份標準實習活頁簿檔案，建議讀者一邊閱讀本手冊，一邊開啟 Excel 實習檔進行對照操作：")
    
    p_files = doc.add_paragraph()
    p_files.paragraph_format.left_indent = Inches(0.2)
    p_files.paragraph_format.space_after = Pt(8)
    r_f1 = p_files.add_run("1. 示範檔：【航運資料前處理教學_30筆示範檔.xlsx】\n")
    r_f1.font.bold = True
    r_f1.font.color.rgb = COLOR_NAVY
    r_f1_d = p_files.add_run("   內含第四章與第五章示範專用之 30 筆進出口提單資料，完整呈現清洗前之髒資料、港口代碼對照表以及清洗後標準資料。\n")
    r_f2 = p_files.add_run("2. 練習檔：【航運資料前處理實務_300筆情境練習檔.xlsx】\n")
    r_f2.font.bold = True
    r_f2.font.color.rgb = COLOR_TEAL
    r_f2_d = p_files.add_run("   內含第六章與第七章專用之 300 筆多情境（整櫃 FCL、併櫃 LCL、冷藏櫃、危險品）真實航運練習題、完整解答庫與稽核驗收 KPI 表。")

def build_toc_and_lists(doc):
    doc.add_page_break()
    
    # 1. 章節目錄
    p_toc = doc.add_paragraph()
    r_toc = p_toc.add_run("章節目錄 (Table of Contents)")
    r_toc.font.name = "微軟正黑體"
    r_toc.font.size = Pt(16)
    r_toc.font.bold = True
    r_toc.font.color.rgb = COLOR_NAVY
    pBdr = parse_xml(r'<w:pBdr %s><w:bottom w:val="single" w:sz="18" w:space="4" w:color="1B365D"/></w:pBdr>' % nsdecls('w'))
    p_toc._p.get_or_add_pPr().append(pBdr)
    p_toc.paragraph_format.space_after = Pt(12)

    toc_items = [
        ("導讀與手冊使用指南", "i"),
        ("第一章：MS Excel 操作介面全景導覽", "1"),
        ("  1.1 航運行政人員與 Excel 的初次相遇", "1"),
        ("  1.2 標題列與快速存取工具列 (Quick Access Toolbar)", "2"),
        ("  1.3 功能區 (Ribbon) 與常用索引標籤功能剖析", "3"),
        ("  1.4 名稱方塊 (Name Box) 與資料編輯列 (Formula Bar)", "4"),
        ("  1.5 儲存格座標系統與範圍參照觀念", "5"),
        ("  1.6 工作表標籤 (Sheet Tabs) 與狀態列 (Status Bar) 實務密技", "6"),
        ("第二章：航運資料前處理核心常識與標準規範", "8"),
        ("  2.1 為什麼航運行政人員必須學習資料前處理？(GIGO 原則)", "8"),
        ("  2.2 守則一：原始資料保護原則（建立備份副本）", "9"),
        ("  2.3 守則二：標準二維表格結構 (Tidy Data 規範結構)", "10"),
        ("  2.4 守則三：資料型態判別與對齊法則 (文字靠左、數值靠右)", "12"),
        ("  2.5 守則四：航運常見八大「髒資料」特徵診斷", "13"),
        ("  2.6 守則五：資料前處理標準作業流程 (Preprocessing Pipeline)", "15"),
        ("第三章：30筆航運示範資料集——問題診斷與設計", "17"),
        ("  3.1 航運業務情境設定與欄位規格定義", "17"),
        ("  3.2 30筆原始資料檔全貌與問題點逐筆深度標註", "19"),
        ("  3.3 髒資料成因與清理對策總表", "22"),
        ("第四章：實戰操作示範——Excel 內建選單功能篇", "24"),
        ("  4.1 副本建立與工作環境準備（凍結窗格與自動篩選）", "24"),
        ("  4.2 尋找與取代 (Find & Replace)：快速清除單位字尾與貨幣符號", "26"),
        ("  4.3 資料剖析精靈 (Text to Columns)：一分為二拆分「船名與航次」", "28"),
        ("  4.4 資料剖析精靈 (Text to Columns)：一鍵修復「文字型雜亂日期」", "30"),
        ("  4.5 移除重複項 (Remove Duplicates)：提單與櫃號唯一性檢驗", "32"),
        ("  4.6 選擇性貼上 (Paste Special)：乘運算快速消除文字數字綠色三角", "34"),
        ("  4.7 設定格式化的條件 (Conditional Formatting)：視覺化抓出漏填值與重複", "36"),
        ("第五章：實戰操作示範——函數公式清洗篇", "38"),
        ("  5.1 文字清理三劍客：TRIM()、CLEAN()、UPPER() 打造標準化代碼", "38"),
        ("  5.2 欄位拆分與擷取：LEFT()、RIGHT()、MID()、FIND()、LEN() 深入解析", "40"),
        ("  5.3 字串替換與正規化：SUBSTITUTE() 處理混雜分隔符號", "42"),
        ("  5.4 日期格式重構：DATE()、YEAR()、MONTH()、DAY() 標準化 YMD", "44"),
        ("  5.5 數值強制轉換：VALUE() 與 NUMBERVALUE() 確保運費可計算", "46"),
        ("  5.6 遺漏值與防錯判斷：IF()、ISBLANK()、IFERROR() 補齊預設標記", "48"),
        ("  5.7 跨表關聯標準化：VLOOKUP() / XLOOKUP() 參照國際港口代碼 (UN/LOCODE)", "50"),
        ("  5.8 綜合清洗管線建立與清洗成果 (Before vs. After) 總對照", "52"),
        ("第六章：300筆航運案例情境實戰練習", "55"),
        ("  6.1 練習情境介紹（多元業務型態：整櫃、併櫃、冷藏、危險品）", "55"),
        ("  6.2 300筆資料清洗規格書與任務驗收五大指標", "56"),
        ("  6.3 300筆資料集欄位與資料缺陷分佈診斷", "58"),
        ("第七章：300筆練習題參考解題步驟與驗證指南", "60"),
        ("  7.1 階段一：結構診斷與備份建置", "60"),
        ("  7.2 階段二：重複提單清理與唯一識別碼驗證 (300筆降至285筆)", "61"),
        ("  7.3 階段三：文字標準化、前後空白清理與大小寫統一", "62"),
        ("  7.4 階段四：港口名稱與國際代碼標準化匹配 (VLOOKUP 實作)", "64"),
        ("  7.5 階段五：數值與單位分離（重量、體積、運費數值化）", "65"),
        ("  7.6 階段六：日期正規化與時效合理性檢核 (ETD vs. ETA)", "67"),
        ("  7.7 階段七：遺漏值處置與報關放行狀態註記", "68"),
        ("  7.8 解題成果關鍵指標查核對帳表 (KPIs Audit Sheet)", "70"),
        ("參考文獻與延伸閱讀", "72")
    ]

    for title, pno in toc_items:
        p_item = doc.add_paragraph()
        p_item.paragraph_format.line_spacing = 1.15
        p_item.paragraph_format.space_after = Pt(2)
        r_t = p_item.add_run(title)
        r_t.font.name = "微軟正黑體"
        r_t.font.size = Pt(9.5)
        if not title.startswith("  "):
            r_t.font.bold = True
            r_t.font.color.rgb = COLOR_NAVY
        else:
            r_t.font.color.rgb = COLOR_BODY
        
        # Dots and Page
        r_dots = p_item.add_run(f" {'·' * max(3, 65 - len(title)*2)} ")
        r_dots.font.name = "Consolas"
        r_dots.font.size = Pt(8.5)
        r_dots.font.color.rgb = COLOR_MUTED
        
        r_p = p_item.add_run(pno)
        r_p.font.name = "微軟正黑體"
        r_p.font.bold = True
        r_p.font.size = Pt(9.5)
        r_p.font.color.rgb = COLOR_NAVY

    # 2. 圖表目錄
    p_lof = doc.add_paragraph()
    p_lof.paragraph_format.space_before = Pt(16)
    r_lof = p_lof.add_run("圖目錄與表目錄 (List of Figures & Tables)")
    r_lof.font.name = "微軟正黑體"
    r_lof.font.size = Pt(14)
    r_lof.font.bold = True
    r_lof.font.color.rgb = COLOR_TEAL
    pBdr2 = parse_xml(r'<w:pBdr %s><w:bottom w:val="single" w:sz="12" w:space="3" w:color="008080"/></w:pBdr>' % nsdecls('w'))
    p_lof._p.get_or_add_pPr().append(pBdr2)
    p_lof.paragraph_format.space_after = Pt(8)

    figures = [
        ("圖 1-1：MS Excel 2021 / 365 核心操作介面區域導覽", "2"),
        ("圖 2-1：航運業務資料前處理黃金標準作業流程 (Preprocessing SOP)", "15"),
        ("圖 2-2：標準二維表格 (Tidy Data) vs. 錯誤結構對照圖", "11"),
        ("圖 4-1：Excel 核心防護技能：工作表完整備份操作視窗", "25"),
        ("圖 4-2：常用檢視與探索功能：凍結首列與啟用自動篩選", "26"),
        ("圖 4-3：批次清理利器：【尋找與取代】剔除單位字尾與符號", "27"),
        ("圖 4-4：Excel 核心技能：【資料剖析精靈】三步驟拆分『船名與航次』", "29"),
        ("圖 4-5：日期大修復：利用【資料剖析精靈】一鍵將雜亂文字日期轉為標準日期", "31"),
        ("圖 4-6：移除重複項 (Remove Duplicates) 核取方塊設定與完成提示", "33"),
        ("圖 4-7：神級技巧：【選擇性貼上】乘運算瞬間消除綠色錯誤三角提示", "35"),
        ("圖 4-8：視覺化查核：【設定格式化的條件】標記重複提單與遺漏值", "37"),
        ("圖 5-1：文字清洗核心管線：TRIM()、CLEAN() 與 UPPER() 串聯運算", "39"),
        ("圖 5-2：代碼標準化大師：VLOOKUP / XLOOKUP 跨表自動精準關聯", "51"),
        ("圖 6-1：300筆情境練習檔之業務結構、資料缺陷分佈與清洗指標藍圖", "59"),
        ("表 2-1：航運行政常見八大「髒資料」特徵診斷與衝擊分析表", "14"),
        ("表 3-1：30筆航運示範資料集欄位屬性規格與前處理需求定義表", "18"),
        ("表 3-2：30筆航運示範原始資料全表與髒資料標註對照表", "20"),
        ("表 4-1：Excel 內建選單清理工具優缺點與適用情境總覽表", "24"),
        ("表 5-1：航運資料清洗必備核心函數功能與語法快速查閱表", "38"),
        ("表 5-2：30筆示範檔清洗前後 (Before vs. After) 數據完整對照驗收表", "53"),
        ("表 7-1：300筆情境練習題清洗成果關鍵指標 (KPI) 查核對帳表", "71")
    ]

    for item, pno in figures:
        p_f = doc.add_paragraph()
        p_f.paragraph_format.line_spacing = 1.15
        p_f.paragraph_format.space_after = Pt(2)
        r_it = p_f.add_run(item)
        r_it.font.name = "微軟正黑體"
        r_it.font.size = Pt(9.0)
        r_it.font.color.rgb = COLOR_BODY
        
        r_dots = p_f.add_run(f" {'·' * max(3, 62 - len(item)*2)} ")
        r_dots.font.name = "Consolas"
        r_dots.font.size = Pt(8.0)
        r_dots.font.color.rgb = COLOR_MUTED
        
        r_pn = p_f.add_run(pno)
        r_pn.font.name = "微軟正黑體"
        r_pn.font.bold = True
        r_pn.font.size = Pt(9.0)
        r_pn.font.color.rgb = COLOR_TEAL

print("content_frontmatter.py created successfully.")
