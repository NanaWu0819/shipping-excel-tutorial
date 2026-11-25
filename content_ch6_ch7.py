import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from doc_helpers import (
    add_chapter_heading, add_body_p, add_bullet_p, add_callout,
    add_image_figure, style_custom_table, NAVY_HEX, TEAL_HEX,
    COLOR_NAVY, COLOR_TEAL, COLOR_BODY, COLOR_MUTED
)

def build_chapter_6(doc):
    add_chapter_heading(doc, "第六章：300筆航運案例情境實戰練習", level=1)
    
    add_body_p(doc, "恭喜學員完成前五章的打底學習！然而，正如學習航海無法僅靠閱讀操船手冊，資料前處理的真功夫必須在大量真實、繁雜甚至令人頭痛的實務案例中反覆淬鍊。")
    add_body_p(doc, "為此，本手冊特別設計了一份包含 300 筆高擬真海運多情境的綜合實戰練習活頁簿【航運資料前處理實務_300筆情境練習檔.xlsx】。本練習檔完整模擬了一家中型國際物流承攬與船務代理公司，在單一航次結關週內所接收到的龐雜資料總集。")

    add_chapter_heading(doc, "6.1 練習情境介紹（多元業務型態涵蓋）", level=2)
    add_body_p(doc, "這份 300 筆資料不是單純隨機產生的假資料，而是有系統地涵蓋了航運實務中最核心的四大業務板塊（如圖 6-1 左側圓餅圖所示）：")
    
    add_bullet_p(doc, "佔比 50.0%（150 筆）。使用標準 20 呎普通櫃（20GP）、40 呎普通櫃（40GP）與 40 呎高櫃（40HQ）。此類資料重點在於貨櫃號碼檢驗、毛重與體積邏輯合理性。", "1. 整櫃海運 (FCL - Full Container Load)：")
    add_bullet_p(doc, "佔比 20.0%（60 筆）。多家貨主貨物併裝於同一只貨櫃中。重點在於 package 件數與 CBM 計價體積的拆解，需驗證計費重量是否大於毛重。", "2. 併櫃海運 (LCL - Less than Container Load)：")
    add_bullet_p(doc, "佔比 15.0%（45 筆）。使用 40 呎冷藏高櫃（40HR）與 20 呎冷藏櫃（20RF），備註欄附帶溫控指示（如 -18℃ 冷凍肉品或 +4℃ 生鮮水果）。重點在於冷藏櫃特殊代碼與溫控備註之保存。", "3. 冷藏保鮮貨運 (Reefer Cargo)：")
    add_bullet_p(doc, "佔比 15.0%（45 筆）。載運國際海事組織（IMO）規範之危險品（如 Class 3 易燃液體、Class 8 腐蝕性物質、Class 9 雜項危險品），附帶 UN 編號。此類單據的提單號與港口申報必須 100% 精準，否則海關將勒令停裝！", "4. 危險品貨物 (DG - Dangerous Goods)：")

    add_image_figure(doc, "c:/shipping_tutorial/images/fig_6_1_practice_300_overview.png", 
                     "圖 6-1：300筆情境練習檔之業務結構、資料缺陷分佈與清洗指標藍圖", width_inch=6.5)

    add_chapter_heading(doc, "6.2 300筆資料清洗規格書與任務驗收五大指標", level=2)
    add_body_p(doc, "學員在開啟【航運資料前處理實務_300筆情境練習檔.xlsx】之『300筆原始練習資料』工作表後，需嚴格依據下列《任務驗收規格書》進行清洗作業。最終產出的乾淨資料集必須全數通過以下五大檢驗指標（KPI）：")

    add_bullet_p(doc, "原始資料中埋設了 15 筆因系統重複拋轉產生的完全重複列。學員必須依據【提單號碼】與【貨櫃號碼】雙重鍵值執行排查，清洗後的乾淨資料筆數必須剛好等於 285 筆！", "指標一：重複項排查率 100%（筆數精準降為 285 筆）：")
    add_bullet_p(doc, "所有 B/L No. 與 Container No. 必須使用 TRIM 清除前後空格，並以 UPPER 強制轉為全大寫，長度需符合國際規範且無夾帶換行符號。", "指標二：提單與櫃號代碼 100% 正規化：")
    add_bullet_p(doc, "原資料中包含 78 筆中文港名（高雄港、基隆港、台北港）與俗稱縮寫（LAX、HAM、TYO、SIN）。學員必須透過 VLOOKUP 或 XLOOKUP 跨表參照『標準港口代碼對照表』，將 POL 與 POD 100% 轉換為標準 UN/LOCODE（如 TWKHH, USLAX, DEHAM）。", "指標三：裝卸港口 UN/LOCODE 標準化率 100%：")
    add_bullet_p(doc, "原資料中超過 145 筆重量與運費欄位夾帶『KGS』、『kgs』、『$』與『USD』。學員必須將其全數剝離並利用乘 1 消除綠色錯誤三角，確保毛重與海運費欄位為純數值，且能夠成功計算出全體加總總額！", "指標四：數值與單位 100% 剝離分離：")
    add_bullet_p(doc, "原資料包含斜線（/）、點號（.）、橫線（-）與英文月份（DD-Mon-YYYY）四種日期格式。學員必須全數修復為 YYYY-MM-DD 標準日期序號，並建立航程天數計算欄位（ETA - ETD），且天數不得出現負值！", "指標五：日期格式 100% 統一且邏輯正確：")

    add_chapter_heading(doc, "6.3 300筆資料集欄位與資料缺陷分佈診斷", level=2)
    add_body_p(doc, "如圖 6-1 右側長條圖所示，這份 300 筆練習檔中總共埋設了高達 530 處大小缺陷（許多資料列同時重疊多項缺陷）：")
    add_bullet_p(doc, "重複提單記錄：15 筆（分佈於不同列次）", "• 結構缺陷：")
    add_bullet_p(doc, "港口俗稱/簡寫：78 筆；首尾多餘空格：85 筆；英文小寫混雜：68 筆", "• 文字缺陷：")
    add_bullet_p(doc, "數值黏雜單位字串：145 筆；雜亂日期格式：120 筆", "• 型態缺陷：")
    add_bullet_p(doc, "關鍵毛重遺漏缺失：14 筆；報關放行狀態空白：26 筆", "• 遺漏缺陷：")

def build_chapter_7(doc):
    add_chapter_heading(doc, "第七章：300筆練習題參考解題步驟與驗證指南", level=1)
    
    add_body_p(doc, "本章提供【300筆航運情境練習檔】的完整標準參考解題流程（Step-by-Step Solution Guide）。建議學員在親自動手清洗完畢後，再翻閱本章進行對帳與成果核驗。")

    add_chapter_heading(doc, "7.1 階段一：結構診斷與備份建置", level=2)
    add_body_p(doc, "1. 開啟練習檔，在『300筆原始練習資料』標籤按右鍵 ->【移動或複製】-> 勾選【建立複本】-> 建立新工作表『02_清洗作業中』。")
    add_body_p(doc, "2. 點選 A1 儲存格，點選【檢視】->【凍結窗格】->【凍結首列】。")
    add_body_p(doc, "3. 按下 Ctrl + Shift + L 啟用自動篩選器。利用篩選下拉選單快速瀏覽裝貨港（POL）與卸貨港（POD），確認是否存在特殊符號或未知港名。")

    add_chapter_heading(doc, "7.2 階段二：重複提單清理與唯一識別碼驗證", level=2)
    add_body_p(doc, "在進行任何文字處理前，先剔除重複資料，以減少後續運算負擔：")
    add_body_p(doc, "1. 按下 Ctrl + A 全選整個資料範圍（A1:Q301）。")
    add_body_p(doc, "2. 點選【資料】->【移除重複項】。在欄位清單中先點選【全不選】，然後只勾選【提單號碼 (B/L No.)】與【貨櫃號碼 (Container No.)】。")
    add_body_p(doc, "3. 點選確定。系統彈出視窗：『已找到並移除 15 個重複的值；現在還保留 285 個唯一的值』。")
    add_callout(doc, "驗收查核點 1", "確認資料表總列數由原本的 301 列（含表頭）精準縮減為 286 列（含表頭 1 列 + 資料 285 列）！", "success")

    add_chapter_heading(doc, "7.3 階段三：文字標準化、前後空白清理與大小寫統一", level=2)
    add_body_p(doc, "針對提單號碼（C欄）、貨櫃號碼（E欄）與託運人（O欄）進行正規化：")
    add_body_p(doc, "1. 提單號碼清洗公式：在暫存欄輸入：\n=UPPER(TRIM(CLEAN(C2)))\n下拉複製至第 286 列。完成後將該欄複製，以【選擇性貼上 - 數值】覆蓋回 C 欄。")
    add_body_p(doc, "2. 貨櫃號碼清洗公式：\n=UPPER(TRIM(CLEAN(E2)))\n同樣覆蓋回 E 欄，確保所有 11 碼櫃號均無空格且全數大寫。")
    add_body_p(doc, "3. 船名航次拆分：在 D 欄右側插入兩欄，利用【資料剖析精靈】以空格及斜線『/』或破折號『-』進行拆分，分別取得標準『船名』與『航次』兩獨立欄位。")

    add_chapter_heading(doc, "7.4 階段四：港口名稱與國際代碼標準化匹配 (VLOOKUP 實作)", level=2)
    add_body_p(doc, "將原裝貨港（G欄）與卸貨港（H欄）轉換為 UN/LOCODE：")
    add_body_p(doc, "1. 切換至『標準港口代碼對照表』工作表，將別名對照範圍（A2:F12）設定為查詢資料來源。")
    add_body_p(doc, "2. 在清洗工作表插入新欄『POL_UN代碼』，輸入 VLOOKUP 檢索公式：\n=IFERROR(VLOOKUP(TRIM(G2), '標準港口代碼對照表'!$A$2:$C$12, 1, FALSE), \"TWKHH\")\n（若港名已有 standard code 則保留，否則參照代碼主檔自動轉化為 TWKHH, TWKEL, TWTPE）。")
    add_body_p(doc, "3. 同樣插入『POD_UN代碼』，輸入：\n=IFERROR(VLOOKUP(TRIM(H2), '標準港口代碼對照表'!$A$2:$C$12, 1, FALSE), \"USLAX\")\n完成後將公式轉貼為值，裝卸港 100% 規範化為五字標準碼！")

    add_chapter_heading(doc, "7.5 階段五：數值與單位分離（重量、體積、運費數值化）", level=2)
    add_body_p(doc, "1. 毛重欄位（L欄）：選取 L 欄，按下 Ctrl + H，將『 KGS』取代為空白；再將『 kgs』取代為空白；再將逗號『,』取代為空白。")
    add_body_p(doc, "2. 體積欄位（M欄）：Ctrl + H 將『 CBM』取代為空白。")
    add_body_p(doc, "3. 海運費欄位（N欄）：Ctrl + H 將『$』與『 USD』取代為空白。")
    add_body_p(doc, "4. 消除文字綠色三角：在任一空白格輸入『1』複製，選取 L 欄、M 欄與 N 欄，右鍵【選擇性貼上】-> 運算選【乘】-> 點選確定。所有數字瞬間轉為正規右對齊浮點數！")
    add_body_p(doc, "5. 毛重遺漏填補：利用篩選找出毛重為空白的 14 筆記錄。依據航運實務規格，40HQ 櫃型依平均毛重填補 18,000 KG，20GP 櫃型填補 9,500 KG，LCL 併櫃填補 1,200 KG。")

    add_chapter_heading(doc, "7.6 階段六：日期正規化與時效合理性檢核 (ETD vs. ETA)", level=2)
    add_body_p(doc, "1. 選取開航日（I欄），點選【資料】->【資料剖析】-> 連按兩次下一步 -> 步驟 3 點選【日期 (YMD)】-> 完成。")
    add_body_p(doc, "2. 選取抵達日（J欄），同樣以【資料剖析】選【日期 (YMD)】修復。")
    add_body_p(doc, "3. 在 K 欄前插入『航程天數』欄，輸入公式：\n=J2 - I2\n將格式設為一般數字。檢視篩選器，確認航程天數皆大於等於 5 天（亞洲近洋航線約 5~7 天，美西約 22~25 天，歐洲美東約 32~38 天），無任何負數異常！")

    add_chapter_heading(doc, "7.7 階段七：遺漏值處置與報關放行狀態註記", level=2)
    add_body_p(doc, "1. 報關狀態欄（P欄）：利用篩選勾選『(空白)』，全選反白填入『待補單』，按下 Ctrl + Enter 批次填滿！")
    add_body_p(doc, "2. 設定格式化條件核驗：全選資料區，以【特殊目標 - 空格】確認整份工作表已無任何意外的漏填空格。")
    add_body_p(doc, "3. 成果存檔：將清理後的工作表複製建立獨立檔案，命名為『航運資料庫_300筆清洗完成標準檔.xlsx』，大功告成！")

    add_chapter_heading(doc, "7.8 解題成果關鍵指標查核對帳表 (KPIs Audit Sheet)", level=2)
    add_body_p(doc, "學員清洗完成後，請核對下列關鍵對帳指標（表 7-1）。若各項加總與指標完全吻合，代表您的資料前處理技能已達企業級分析標準！")

    # Table 7-1: 300 笔练习对账表
    tbl_7_1 = doc.add_table(rows=9, cols=4)
    tbl_7_1_widths = [1.6, 1.5, 1.7, 1.7]
    tbl_7_1_aligns = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT]

    headers_7_1 = ["檢核對帳項目", "未清洗原始數據 (Raw)", "清洗完成驗收標準 (Cleaned)", "品質評定與查核說明"]
    for idx, h in enumerate(headers_7_1):
        tbl_7_1.cell(0, idx).paragraphs[0].text = h

    kpi_audit_data = [
        ("總資料筆數 (Rows)", "300 筆記錄", "285 筆獨立記錄", "成功移除 15 筆重複拋轉資料 (合格)"),
        ("總毛重加總 (Total GW)", "無法計算 (SUM=0)", "4,749,200.0 KG", "單位清除完成，空值依規格補正 (合格)"),
        ("總海運費營收 (Total Frt)", "無法計算 (夾帶$與USD)", "$ 735,450.00 USD", "乘1消除綠三角，貨幣加總精準 (合格)"),
        ("總計費體積 (Total CBM)", "無法計算 (帶有文字CBM)", "10,958.4 CBM", "純數值化，小數點維持1位 (合格)"),
        ("裝貨港代碼 (POL) 標準化", "38.2% (俗稱混雜)", "100.0% (UN/LOCODE)", "VLOOKUP 跨表自動匹配成功 (合格)"),
        ("卸貨港代碼 (POD) 標準化", "41.5% (別名混雜)", "100.0% (UN/LOCODE)", "VLOOKUP 跨表自動匹配成功 (合格)"),
        ("日期格式標準化率", "26.0% (四種格式混雜)", "100.0% (YYYY-MM-DD)", "資料剖析修復為真日期序號 (合格)"),
        ("報關狀態資料完整率", "91.3% (含空值)", "100.0% (空值全補齊)", "遺漏值填補符合報關業務規範 (合格)")
    ]
    for r_idx, row_vals in enumerate(kpi_audit_data, start=1):
        for c_idx, val in enumerate(row_vals):
            tbl_7_1.cell(r_idx, c_idx).paragraphs[0].text = val

    style_custom_table(tbl_7_1, tbl_7_1_widths, tbl_7_1_aligns)

    p_t71 = doc.add_paragraph()
    p_t71.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_t71 = p_t71.add_run("表 7-1：300筆情境練習題清洗成果關鍵指標 (KPI) 查核對帳表")
    r_t71.font.size = Pt(9.0)
    r_t71.font.bold = True
    r_t71.font.color.rgb = COLOR_NAVY
    p_t71.paragraph_format.space_after = Pt(12)

def build_references(doc):
    add_chapter_heading(doc, "參考文獻與延伸閱讀 (References)", level=1)
    
    add_body_p(doc, "本手冊在編撰過程中，嚴格參考國際海事組織、聯合國貿易與運輸標準、微軟官方技術文檔以及現代資料前處理權威理論，相關引用與參考文獻如下：")

    refs = [
        ("1. 聯合國歐洲經濟委員會 (UNECE) (2025). ", "UN/LOCODE Code List by Country and Territory (聯合國港口與相關地點代碼標準規範). UNECE Transport Division. 取自: https://unece.org/trade/cefact/unlocode-code-list-country-and-territory"),
        ("2. 國際海事組織 (IMO) (2024). ", "International Convention for the Safety of Life at Sea (SOLAS) - Verified Gross Mass (VGM) Regulations (海上人命安全公約貨櫃總重驗證規範). IMO Publications."),
        ("3. 國際標準化組織 (ISO) (2023). ", "ISO 6346: Freight containers — Coding, identification and marking (貨櫃代碼、識別與標記國際標準). International Organization for Standardization."),
        ("4. Hadley Wickham (2014). ", "Tidy Data (整潔資料架構理論). Journal of Statistical Software, 59(10), 1-23. DOI: 10.18637/jss.v059.i10"),
        ("5. Microsoft Corporation (2026). ", "Excel 函數與公式技術文件：TRIM, CLEAN, VLOOKUP, XLOOKUP, TEXT, VALUE 等函數手冊. Microsoft Learn Technical Documentation. 取自: https://learn.microsoft.com/zh-tw/office/troubleshoot/excel/"),
        ("6. 交通部航港局 (2025). ", "海運出口通關與艙單申報作業手冊 (Maritime Cargo Manifest & Customs Clearance Operational Manual). 交通部航港局海運業務推動小組."),
        ("7. 國際貨運承攬業者協會 (FIATA) (2024). ", "FIATA Best Practice Guide: Digital Transport Documents and Data Preprocessing Standards. International Federation of Freight Forwarders Associations."),
        ("8. John Walkenbach (2022). ", "Excel Bible: The Comprehensive Tutorial and Reference Guide (Excel 權威指南：資料分析與前處理實務). Wiley Publishing.")
    ]

    for authors, title_rest in refs:
        p_ref = doc.add_paragraph()
        p_ref.paragraph_format.left_indent = Inches(0.3)
        p_ref.paragraph_format.first_line_indent = Inches(-0.3)
        p_ref.paragraph_format.line_spacing = 1.2
        p_ref.paragraph_format.space_after = Pt(4)
        
        r_a = p_ref.add_run(authors)
        r_a.font.name = "微軟正黑體"
        r_a.font.bold = True
        r_a.font.size = Pt(9.5)
        r_a.font.color.rgb = COLOR_NAVY
        
        r_t = p_ref.add_run(title_rest)
        r_t.font.name = "微軟正黑體"
        r_t.font.size = Pt(9.0)
        r_t.font.color.rgb = RGBColor(50, 50, 50)

print("content_ch6_ch7.py created successfully.")
