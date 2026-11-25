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

def build_chapter_1(doc):
    add_chapter_heading(doc, "第一章：MS Excel 操作介面全景導覽", level=1)
    
    add_body_p(doc, "對於初入國際航運物流產業的行政與文件人員而言，Microsoft Excel 不僅僅是一套辦公室試算表軟體，更是整座全球貨櫃航運供應鏈的核心溝通語言。無論是船務代理的艙單申報（Manifest）、貨運承攬（Forwarder）的訂艙確認書（Booking Confirmation）、貨櫃集散站（CY）的進出場進度表，或是海關放行清冊，幾乎全數以 Excel 活頁簿的形式流轉。")

    add_body_p(doc, "工欲善其事，必先利其器。在深入學習資料前處理技巧前，必須徹底摸透 Excel 的工作環境與八大關鍵介面區域。唯有熟悉各區域的配置與運作邏輯，才能在面對數萬筆海運單據時，迅速完成視圖定位、資料篩選與格式診斷。")

    add_image_figure(doc, "c:/shipping_tutorial/images/fig_1_1_excel_ui.png", 
                     "圖 1-1：MS Excel 2021 / 365 核心操作介面區域導覽（八大關鍵區域圖解）", width_inch=6.5)

    add_chapter_heading(doc, "1.1 啟動 Excel 與核心介面佈局", level=2)
    add_body_p(doc, "當我們開啟 Excel 活頁簿時，整體工作畫面主要由八大功能區塊組成（如圖 1-1 所示）。對於資料前處理工作而言，特別需要關注以下八大區域：")

    add_bullet_p(doc, "位於視窗最頂端。顯示目前開啟的活頁簿檔案名稱（如『航運出口提單總表.xlsx』）。左側設有『快速存取工具列』，包含復原（Ctrl+Z）、重做（Ctrl+Y）與自動儲存開關。", "1. 標題列 (Title Bar)：")
    add_bullet_p(doc, "橫跨視窗上方的核心指令面板。由多個『索引標籤』（常用、公式、資料等）組成，群組化收納所有資料處理指令按鈕。", "2. 功能區索引標籤 (Ribbon Tabs)：")
    add_bullet_p(doc, "位於功能區左下方的小型文字方塊。即時顯示當前作用儲存格（Active Cell）的座標代碼（如 C5），亦可用於快速跳轉至指定儲存格或已命名的資料範圍。", "3. 名稱方塊 (Name Box)：")
    add_bullet_p(doc, "位於名稱方塊右側的狹長文字輸入列。顯示當前儲存格的真實內容、文字或計算公式。在此可檢驗儲存格是否含有不可見的前後空格或未顯示的底層值。", "4. 資料編輯列 (Formula Bar)：")
    add_bullet_p(doc, "水平頂端以英文字母標示的『欄號』（A, B, C... 至 XFD），以及垂直左側以阿拉伯數字標示的『列號』（1, 2, 3... 至 1,048,576）。兩者交叉定位出唯一儲存格。", "5. 欄號與列號座標系統 (Row & Column Headers)：")
    add_bullet_p(doc, "由無數矩形儲存格構成的工作主體。當前被綠色外框包圍並反藍的儲存格稱為『作用儲存格』，是資料輸入與公式計算的目標主體。", "6. 儲存格二維網格區域 (Grid Area)：")
    add_bullet_p(doc, "位於視窗左下方。顯示目前活頁簿內包含的所有工作表名稱（如『01_原始資料』、『02_港口代碼表』）。點選『+』按鈕可新增工作表，右鍵點擊可進行複製、重新命名與填色。", "7. 工作表標籤 (Sheet Tabs)：")
    add_bullet_p(doc, "位於視窗最底部。即時顯示目前系統狀態（如『就緒』或『篩選中』）。當在資料區圈選數值儲存格時，狀態列右側會自動顯示『平均值』、『計數』與『加總』，是航運快速對帳的神器。", "8. 狀態列 (Status Bar)：")

    add_chapter_heading(doc, "1.2 標題列與快速存取工具列實戰設定", level=2)
    add_body_p(doc, "航運行政人員每天需要處理大量單據，稍有不慎按錯按鍵可能導致數小時的清理成果毀於一旦。因此，強烈建議將最常用的指令固定於『快速存取工具列』：")
    add_bullet_p(doc, "點選標題列左上角的小箭頭，勾選『自動存檔』與『儲存』。隨時按下 Ctrl+S，養成存檔好習慣。", "存檔自保防護：")
    add_bullet_p(doc, "如果不小心誤刪了提單欄位，連續按下 Ctrl+Z 可即時退回前一步驟；若退過頭則按 Ctrl+Y 恢復。", "神級防錯快捷鍵：")

    add_chapter_heading(doc, "1.3 功能區 (Ribbon) 與常用索引標籤功能剖析", level=2)
    add_body_p(doc, "在資料前處理的生命週期中，有三個索引標籤是航運人員最密集使用的核心根據地：")
    add_bullet_p(doc, "包含字型設定、靠左/靠右對齊、文字與數值格式切換（通用格式、貨幣、小數位數），以及最重要的【設定格式化的條件】與【排序與篩選】。", "【常用】(Home) 標籤：")
    add_bullet_p(doc, "包含文字清洗函數（TRIM, CLEAN, UPPER）、參照尋找函數（VLOOKUP, XLOOKUP）以及條件判斷函數（IF, ISBLANK）。點選『fx 插入函數』可開啟函數引數對話框。", "【公式】(Formulas) 標籤：")
    add_bullet_p(doc, "資料前處理的重裝武器庫！內含【從文字/CSV取得資料】、【資料剖析】（拆分船名與修復日期）、【移除重複項】（排查重複提單）與【資料驗證】。", "【資料】(Data) 標籤：")

    add_callout(doc, "快捷展開與隱藏功能區", 
                "若筆記型電腦螢幕較小，按下快捷鍵 Ctrl+F1 可切換功能區的展開與收合，讓資料顯示區域瞬間擴大 30%，方便檢視長達數十欄的航運出口總清單！", 
                "tip")

    add_chapter_heading(doc, "1.4 名稱方塊與資料編輯列：診斷髒資料的照妖鏡", level=2)
    add_body_p(doc, "初學者常犯的錯誤是『只看儲存格表面顯示的樣子』。例如，某儲存格表面顯示『2026/11/05』，但在資料編輯列中可能其實是文字字串；或者表面看起來是一片空白，但資料編輯列中卻藏著一個不可見的半形空格！")
    add_body_p(doc, "資料編輯列就是航運人員的『顯微鏡』。只要將游標移至資料編輯列內，就能精準看見字串最前方是否有偷偷塞入的空白、公式底層引數設定，以及數值的真實小數點位數。")

    add_chapter_heading(doc, "1.5 儲存格座標系統與範圍參照觀念", level=2)
    add_body_p(doc, "Excel 是以欄（直向 Alphabet）與列（橫向 Number）構成的座標系。在後續編寫資料清洗公式時，必須透徹理解相對參照與絕對參照的差異：")
    add_bullet_p(doc, "如 A1、B2。當公式向下或向右複製拖曳時，參照列號與欄號會隨之自動遞增。適用於逐列清理提單號碼。", "相對參照 (Relative Reference)：")
    add_bullet_p(doc, "如 $A$1 或 $A$2:$C$12。在欄號與列號前加上錢字號『$』，鎖定參照目標。在跨表參照【港口標準代碼表】使用 VLOOKUP 時，必須按下 F4 鍵將代碼表範圍鎖定為絕對參照，否則公式下拉時對照範圍會錯位！", "絕對參照 (Absolute Reference)：")

    add_chapter_heading(doc, "1.6 工作表標籤與狀態列進階實務密技", level=2)
    add_body_p(doc, "工作表標籤代表一本活頁簿內的不同頁籤。航運處理大原則是：『原始資料』獨立一頁、『代碼對照表』獨立一頁、『清洗作業區』獨立一頁。右鍵點擊工作表標籤可設定『索引標籤色彩』，將原始資料標記為灰色（禁止更動）、標準代碼表標記為藍色、清洗成果標記為綠色。")
    add_body_p(doc, "狀態列更是快速對帳的無名英雄。當報關行詢問：『這批貨目前放行的貨櫃總毛重是多少？』你只需利用自動篩選勾選『已放行』，然後用滑鼠框選毛重欄位，目光掃向視窗右下角狀態列，立即就能看到加總金額與貨櫃口數，完全不需要額外輸入 SUM 函數！")

def build_chapter_2(doc):
    add_chapter_heading(doc, "第二章：航運資料前處理核心常識與標準規範", level=1)
    
    add_body_p(doc, "資料前處理（Data Preprocessing）在現代商業數據分析中佔據了 70% 至 80% 的工作時間。對於航運行政人員而言，資料前處理不是單純的美化排版，而是守護航運合約、報關合規與運費結算的防線。")

    add_chapter_heading(doc, "2.1 為什麼航運行政人員必須學習資料前處理？", level=2)
    add_body_p(doc, "國際海運涉及託運人（Shipper）、收貨人（Consignee）、船公司（Carrier）、海關、港務局與保險公司等多方機構。一旦資料來源存在缺陷，將引發連鎖反應：")
    add_bullet_p(doc, "若毛重欄位混入文字『KGS』，Excel 無法計算總重，導致向船公司申報的 VGM（貨櫃總重驗證）錯誤，貨櫃將被碼頭拒絕裝船（No VGM, No Load）！", "VGM 申報不合規被拒裝船：")
    add_bullet_p(doc, "提單若存在前後空格，VLOOKUP 函數將回傳 #N/A 錯誤，導致目的港港口代碼抓錯，貨物恐被錯運至其他國家！", "跨表查詢失效導致錯運：")
    add_bullet_p(doc, "若海運提單清冊中混入重複列，行政人員可能向客戶重複開立運費發票，或重複放行提單（B/L Release），造成公司法律糾紛與財務損失。", "重複放行引發重大索賠：")

    add_callout(doc, "航運資料前處理金律：GIGO 原則", 
                "Garbage In, Garbage Out（垃圾進，垃圾出）是資訊科學的真理。如果你輸入的資料是雜亂無章的髒資料，無論使用多麼強大的樞紐分析表或高級商業智慧軟體，產出的統計報表與圖表必定是毫無價值甚至誤導決策的垃圾！", 
                "concept")

    add_chapter_heading(doc, "2.2 守則一：原始資料保護原則（建立備份副本）", level=2)
    add_body_p(doc, "資料清洗的第一條鐵律：【永遠不要直接在原始資料（Raw Data）上進行不可逆的修改】！")
    add_body_p(doc, "當系統匯入或收到客戶寄來的 Excel 檔案時，首要動作是在工作表標籤上按右鍵，點選【移動或複製】，勾選【建立複本】。將原始工作表命名為『01_原始資料(未清理)』並將其鎖定或標記為保護狀態。所有清洗、拆分與刪除動作，只能在複製出的新工作表（如『02_清洗作業區』）中執行。這樣若清洗中途公式出錯，隨時能與原始版本進行核對回溯。")

    add_chapter_heading(doc, "2.3 守則二：標準二維表格結構 (Tidy Data 規範結構)", level=2)
    add_body_p(doc, "許多初學 Excel 的航運人員習慣製作『人類視覺閱讀友善』的報表，例如將船名跨欄置中合併、在資料中間插入小計列、建立多層次複雜標題（如圖 2-2 左側）。這種表格結構在資料分析領域被稱為『報表型表格』，完全無法被 Excel 函數、排序、篩選或樞紐分析表正確讀取！")

    add_image_figure(doc, "c:/shipping_tutorial/images/fig_2_2_tidy_data.png", 
                     "圖 2-2：標準二維表格 (Tidy Data) 與常見錯誤報表結構對比圖", width_inch=6.5)

    add_body_p(doc, "標準資料分析必須遵循 Hadley Wickham 提出的【Tidy Data（整潔資料）】原則（如圖 2-2 右側）：")
    add_bullet_p(doc, "每一筆橫列只能代表一張提單（B/L）或一個貨櫃的獨立交易記錄，絕對不可垂直合併多列。", "原則一：一列代表一筆獨立記錄 (Each row is an observation)：")
    add_bullet_p(doc, "每個直欄只能記錄一種單一業務屬性（例如船名獨立一欄、航次獨立一欄、重量數值獨立一欄），不可將多種資訊混在一格。", "原則二：一行代表一個單一屬性 (Each column is a variable)：")
    add_bullet_p(doc, "工作表的第一列（Row 1）必須是唯一且無合併儲存格的欄位名稱列，嚴禁出現跨欄置中或多層表頭。", "原則三：單一標題列 (Single Header Row)：")
    add_bullet_p(doc, "所有合計或平均值應透過樞紐分析表計算，絕對不可在原始資料行間穿插小計列，否則加總時會導致數據重複計算翻倍！", "原則四：嚴禁在資料區穿插小計列 (No inline subtotals)：")

    add_callout(doc, "為什麼要全面禁止跨欄合併儲存格？", 
                "合併儲存格（Merge Cells）是資料前處理的最大災難！當你合併 A2:A3 時，Excel 實際上只保留 A2 格位有值，而 A3 變成了完全的空值（Blank）。這會導致點選排序時跳出警告錯誤『此操作需要所有合併的儲存格具有相同大小』，同時自動篩選時 A3 會直接消失！", 
                "warning")

    add_chapter_heading(doc, "2.4 守則三：資料型態判別與對齊法則", level=2)
    add_body_p(doc, "Excel 具備智慧型態辨識機制，初學者必須具備透過『預設對齊方式』秒懂資料型態的基本功：")
    add_bullet_p(doc, "如提單號碼『OOCLU9823101』、船名『EVER GIVEN』。預設一律【靠左對齊】。文字無法進行四則運算。", "1. 文字型態 (Text)：")
    add_bullet_p(doc, "如件數 520、毛重 18500.0、海運費 2850.00。預設一律【靠右對齊】。數值可進行加減乘除與 SUM 加總。", "2. 數值型態 (Number)：")
    add_bullet_p(doc, "如開航日 2026/11/05。在 Excel 底層，日期本質上是一個『數值序號』（例如 1900/1/1 為 1，2026/11/5 為 46331）。真正的日期預設【靠右對齊】；若日期靠左對齊，代表它是文字，無法計算航程天數！", "3. 日期與時間 (Date/Time)：")
    add_bullet_p(doc, "公式回傳 TRUE 或 FALSE。預設【置中對齊】。", "4. 邏輯值 (Boolean)：")
    add_bullet_p(doc, "如 #N/A（查無代碼）、#VALUE!（文字無法相加）、#DIV/0!（除以零）。預設【置中對齊】，代表公式運算邏輯異常需進行排除。", "5. 錯誤值 (Error Values)：")

    add_chapter_heading(doc, "2.5 守則四：航運常見八大「髒資料」特徵診斷", level=2)
    add_body_p(doc, "在航運實務作業中，匯入的原始資料常潛藏以下八大典型缺陷：")

    # Table 2-1: 髒資料特徵表
    tbl_2_1 = doc.add_table(rows=9, cols=4)
    tbl_2_1_widths = [1.2, 1.6, 1.8, 1.9]
    tbl_2_1_aligns = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT]

    headers_2_1 = ["缺陷類型", "航運實務典型案例", "對業務與分析的衝擊", "建議前處理對策"]
    for idx, h in enumerate(headers_2_1):
        tbl_2_1.cell(0, idx).paragraphs[0].text = h

    dirty_diagnostics = [
        ("前後多餘空格", "  \"  TWKHH  \"  \n\"EMCU1234567 \"", "VLOOKUP 查無代碼，比對條件失效", "TRIM() 函數清除首尾空格"),
        ("不可見字元", "從 EDI 拋轉帶入之換行符或 Tab 鍵", "匯出報關文本時系統解析錯誤", "CLEAN() 函數清除 ASCII 0~31"),
        ("英文字母大小寫混雜", "ooclu123 vs. OOCLU123", "資料透視表視為不同類別無法合併統計", "UPPER() 強制轉換為全大寫"),
        ("港口代碼俗稱混用", "高雄港、KHH、TWKHH、Kaohsiung", "無法按航線港口統計裝卸貨量", "跨表參照港口對照表統一五字碼"),
        ("數值黏雜文字單位", "18,500 KGS、$2,850 USD", "資料型態變成文字，SUM 加總回傳 0", "【尋找與取代】剔除單位字尾"),
        ("文字格式數字", "綠色三角提示：'2400 (靠左對齊)", "無法運算，各儲存格加總統計失真", "【選擇性貼上】乘 1 轉回真數值"),
        ("日期格式多元雜亂", "2026.11.08 / 12-Nov-2026 / 20261115", "無法按年月份群組統計，無法計算航程", "【資料剖析】選 YMD 格式轉序號"),
        ("重複資料拋轉", "同一提單號碼因 EDI 重送出現多次", "運費計算重複翻倍，提單重複放行風險", "【移除重複項】依提單+櫃號唯一鍵值")
    ]
    for r_idx, row_vals in enumerate(dirty_diagnostics, start=1):
        for c_idx, val in enumerate(row_vals):
            tbl_2_1.cell(r_idx, c_idx).paragraphs[0].text = val

    style_custom_table(tbl_2_1, tbl_2_1_widths, tbl_2_1_aligns)

    p_t21 = doc.add_paragraph()
    p_t21.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_t21 = p_t21.add_run("表 2-1：航運行政常見八大「髒資料」特徵診斷與衝擊分析表")
    r_t21.font.size = Pt(9.0)
    r_t21.font.bold = True
    r_t21.font.color.rgb = COLOR_NAVY
    p_t21.paragraph_format.space_after = Pt(8)

    add_chapter_heading(doc, "2.6 守則五：資料前處理標準作業流程 (SOP)", level=2)
    add_body_p(doc, "面對龐雜的航運資料，不可想到哪做到哪。必須建立結構化的五步標準作業流程（Pipeline），如圖 2-1 所示：")

    add_image_figure(doc, "c:/shipping_tutorial/images/fig_2_1_preprocessing_sop.png", 
                     "圖 2-1：航運業務資料前處理黃金標準作業流程 (Preprocessing SOP)", width_inch=6.5)

    add_bullet_p(doc, "複製原始工作表，另存乾淨備份檔，確立版本命名規範（如_V1_Clean）。", "步驟 1：防護備份 (Backup Protection)：")
    add_bullet_p(doc, "檢查第一列是否為單一標題列，解開所有合併儲存格，啟用凍結窗格與自動篩選，並依據提單號碼執行重複項排查。", "步驟 2：結構與唯一性檢驗 (Structure & Uniqueness)：")
    add_bullet_p(doc, "運用 TRIM/CLEAN 清除隱形字元與多餘空格，UPPER 統一英文大寫，並藉由港口對照表將各地俗稱匹配為國際標準碼。", "步驟 3：字串與代碼標準化 (String & Code Normalization)：")
    add_bullet_p(doc, "剔除重量與運費中的單位文字符號，以運算乘 1 修復綠色三角，並利用資料剖析將雜亂文字日期正規化為真日期序號。", "步驟 4：數值與日期修復 (Numeric & Date Repair)：")
    add_bullet_p(doc, "運用條件判斷填補遺漏值（如空值補齊為『待補單』），設定格式化條件視覺化抓漏，最後驗證加總與筆數並封存為分析標準庫。", "步驟 5：遺漏填補與驗收封存 (Imputation & Archive)：")

def build_chapter_3(doc):
    add_chapter_heading(doc, "第三章：30筆航運示範資料集——問題診斷與設計", level=1)
    
    add_body_p(doc, "為了讓學習者身歷其境體會航運實務痛點，本手冊精心杜撰並設計了一份包含 30 筆真實進出口海運櫃貨的示範資料檔【航運資料前處理教學_30筆示範檔.xlsx】。本章將為學員詳細拆解該資料集的業務背景、欄位定義，並逐筆標記所埋設的各項經典髒資料問題。")

    add_chapter_heading(doc, "3.1 航運情境設定與欄位規格定義", level=2)
    add_body_p(doc, "本資料集模擬台灣主要港口（高雄港 TWKHH、基隆港 TWKEL、台北港 TWTPE）出口至北美（洛杉磯 USLAX、奧克蘭 USOAK、紐約 USNYC）、西北歐（漢堡 DEHAM、鹿特丹 NLRTM）、日本（東京 JPTYO、大阪 JPOSA）與東南亞樞紐港（新加坡 SGSIN）的定期貨櫃航線總清單。")
    add_body_p(doc, "資料集中涵蓋各大海運聯盟航運巨擘（長榮海運 EVERGREEN、陽明海運 YANG MING、萬海航運 WAN HAI、東方海外 OOCL、馬士基 MAERSK、赫伯羅德 HAPAG-LLOYD、海洋網聯 ONE、中遠海運 COSCO）之運作單據，總共包含 15 大關鍵欄位（詳見表 3-1）：")

    # Table 3-1: 欄位定義表
    tbl_3_1 = doc.add_table(rows=16, cols=4)
    tbl_3_1_widths = [1.2, 1.5, 1.4, 2.4]
    tbl_3_1_aligns = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT]

    headers_3_1 = ["欄位名稱", "英文代碼", "預期標準型態", "業務涵義與前處理清洗需求說明"]
    for idx, h in enumerate(headers_3_1):
        tbl_3_1.cell(0, idx).paragraphs[0].text = h

    field_specs = [
        ("項次", "No.", "整數 (Integer)", "流水編號 1 至 30，便於追蹤排序前的位置"),
        ("提單號碼", "B/L No.", "文字 (Text)", "海運唯一提單代碼，需排查前後空格與大小寫"),
        ("船名航次", "Vessel/Voyage", "文字 (Text)", "船名與航次黏在一起，需剖析拆分為兩獨立欄"),
        ("貨櫃號碼", "Container No.", "文字 (Text)", "11位國際ISO櫃號（4字首+7碼數字），需轉大寫"),
        ("櫃型", "Type", "文字 (Text)", "標準櫃型（20GP, 40GP, 40HQ, 45HQ）"),
        ("裝貨港", "POL", "文字 (UN/LOCODE)", "起運港，目前混雜中文與俗稱，需轉五字標準碼"),
        ("卸貨港", "POD", "文字 (UN/LOCODE)", "目的港，包含英文別名與縮寫，需統一標準碼"),
        ("開航日", "ETD", "日期 (YYYY-MM-DD)", "預計離港結關日，格式多樣混雜，需修復為真日期"),
        ("抵達日", "ETA", "日期 (YYYY-MM-DD)", "預計抵達目的港日，需確保日期格式正確可計算天數"),
        ("件數", "Packages", "數值與包裝單位", "原欄混雜『520 CTNS』，需分離純數量與單位"),
        ("毛重", "Gross Weight", "數值 (KG 浮點數)", "原欄混雜『18,500 KGS』與空值，需剔除單位轉數值"),
        ("體積", "CBM", "數值 (CBM 浮點數)", "原欄帶有『CBM』字尾，需分離為純數值"),
        ("海運費", "Freight USD", "貨幣數值 (USD)", "原欄帶有『$』與『USD』，且夾帶文字型數字"),
        ("託運人", "Shipper", "文字 (Text)", "出口廠商名稱，帶有首尾空格或全小寫，需正規化"),
        ("報關狀態", "Customs Status", "文字 (Text)", "通關放行註記，存在遺漏空值，需進行補齊標註")
    ]
    for r_idx, row_vals in enumerate(field_specs, start=1):
        for c_idx, val in enumerate(row_vals):
            tbl_3_1.cell(r_idx, c_idx).paragraphs[0].text = val

    style_custom_table(tbl_3_1, tbl_3_1_widths, tbl_3_1_aligns)

    p_t31 = doc.add_paragraph()
    p_t31.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_t31 = p_t31.add_run("表 3-1：30筆航運示範資料集欄位屬性規格與前處理需求定義表")
    r_t31.font.size = Pt(9.0)
    r_t31.font.bold = True
    r_t31.font.color.rgb = COLOR_NAVY
    p_t31.paragraph_format.space_after = Pt(8)

    add_chapter_heading(doc, "3.2 30筆原始資料檔全貌呈現與髒資料標註", level=2)
    add_body_p(doc, "以下為 30 筆示範檔在未經任何清洗前的原始面貌。讀者可特別觀察標籤標註之問題點（如表 3-2 所示）：")

    # Table 3-2: 30筆完整原始資料
    tbl_3_2 = doc.add_table(rows=31, cols=7)
    tbl_3_2_widths = [0.4, 1.1, 1.4, 0.9, 0.9, 1.0, 0.8]
    tbl_3_2_aligns = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.CENTER]

    headers_3_2 = ["No", "提單號碼", "船名航次 (需拆分)", "裝貨港", "卸貨港", "毛重 (含單位)", "缺陷診斷標籤"]
    for idx, h in enumerate(headers_3_2):
        tbl_3_2.cell(0, idx).paragraphs[0].text = h

    raw_30_table = [
        ("1", "  OOCLU9823101  ", "EVER GIVEN V.0123W", "TWKHH", "USLAX", "18,500 KGS", "首尾空格 / 單位文字"),
        ("2", "EMCE8901234", "YM UNICORN / 045E", "高雄", "LAX", "9,800 kgs", "港口俗稱 / 點號日期"),
        ("3", "MSK7654321", "OOCL POLAND - 102N", "KHH", "Los Angeles", "14,200", "英文港名 / 日期DD-Mon"),
        ("4", "whl9876543", "WAN HAI 315 V.S088", "TWKEL", "Tokyo", "21500 KGS", "小寫提單 / 之後與第28重複"),
        ("5", "ONE8877665", "ONE HARBOR V.022W", "基隆", "DEHAM", "16,800 KGS", "之後與第12筆完全重複"),
        ("6", "CMA1122334", "APL LE HAVRE / 091E", "TWKHH", "Hamburg", "8500", "純數字但缺少單位標示"),
        ("7", "HLC9988776", "VALPARAISO EXPRESS V.004S", "高雄港", "NLRTM", "15,400 KGS", "中文港名需標準化"),
        ("8", "YML0011223", "YM WELLBEING / 033W", "TWKEL", "Rotterdam", "13,600 kgs", "英文小寫單位 kgs"),
        ("9", "OOCLU9823102", "OOCL ROTTERDAM V.115E", "台北港", "JPTYO", "7,200 KGS", "中文台北港需對應"),
        ("10", "MSK7654322", "MAERSK MC-KINNEY V.088E", "TWTPE", "TYO", "19,800 KGS", "三字機場碼 TYO"),
        ("11", "EMCE8901235", "EVER GOLDEN V.009E", "TWKHH", "USOAK", "22,400 KGS", "正常標準樣品"),
        ("12", "ONE8877665", "ONE HARBOR V.022W", "基隆", "DEHAM", "16,800 KGS", "★與第5筆完全重複！"),
        ("13", "  WHL9876544  ", "WAN HAI 502 / W112", "TWKEL", "SGSIN", "9,100 kgs", "提單號包夾空白"),
        ("14", "CMA1122335", "CMA CGM ANTOINE V.044W", "Kaohsiung", "Singapore", "14,800 KGS", "英文港口全名"),
        ("15", "OOCLU9823103", "OOCL SHANGHAI V.078N", "KHH", "SGSIN", "", "★關鍵毛重遺漏缺失！"),
        ("16", "cosu6211223", "COSCO SHIPPING TAURUS / 012E", "TWKHH", "USNYC", "20,500 KGS", "提單小寫 / 櫃號小寫"),
        ("17", "HLC9988777", "HAMBURG EXPRESS V.025W", "基隆港", "New York", "8,900 kgs", "點號日期 / 港口全名"),
        ("18", "MSK7654323", "MADRID MAERSK V.062E", "TWKEL", "USNYC", "23,100 KGS", "日期跨年度計算"),
        ("19", "YML0011224", "YM TRAVEL / 019S", "TWKHH", "JPOSA", "9,400", "文字型帶逗號數字"),
        ("20", "EMCE8901236", "EVER FORTUNE / 088N", "高雄", "Osaka", "15,200 KGS", "港口中文與英文混合"),
        ("21", "ONE8877666", "ONE OLYMPUS V.016W", "TWKEL", "DEHAM", "17,900 kgs", "航次斜線與破折號"),
        ("22", "OOCLU9823104", "OOCL GERMANY V.003E", "TWTPE", "Hamburg", "20,100 KGS", "★報關放行狀態缺失！"),
        ("23", "WHL9876545", "WAN HAI 272 V.N015", "台北港", "Tokyo", "7,800 KGS", "近洋航線短天數"),
        ("24", "CMA1122336", "CMA CGM JACQUES SAADE / 055W", "TWKHH", "NLRTM", "21,800 kgs", "超大型貨櫃船標記"),
        ("25", "HLC9988778", "BERLIN EXPRESS V.002W", "Kaohsiung", "Rotterdam", "16,200 KGS", "港名大小寫混雜"),
        ("26", "MSK7654324", "MUNICH MAERSK V.031E", "KHH", "USLAX", "9,600", "千分位文字型數字"),
        ("27", "cosu6211224", "COSCO PRIDE / 041E", "TWKEL", "USLAX", "21,000 KGS", "提單與櫃號小寫"),
        ("28", "whl9876543", "WAN HAI 315 V.S088", "TWKEL", "Tokyo", "21500 KGS", "★與第4筆完全重複！"),
        ("29", "YML0011225", "YM WIND / 008S", "TWKHH", "JPTYO", "18,400 kgs", "日式日期格式點號"),
        ("30", "EMCE8901237", "EVER APEX V.001W", "高雄港", "USOAK", "22,900 KGS", "貨主名稱帶首尾空格")
    ]
    for r_idx, row_vals in enumerate(raw_30_table, start=1):
        for c_idx, val in enumerate(row_vals):
            tbl_3_2.cell(r_idx, c_idx).paragraphs[0].text = val

    style_custom_table(tbl_3_2, tbl_3_2_widths, tbl_3_2_aligns)

    p_t32 = doc.add_paragraph()
    p_t32.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_t32 = p_t32.add_run("表 3-2：30筆航運示範原始資料全表與髒資料標註對照表")
    r_t32.font.size = Pt(9.0)
    r_t32.font.bold = True
    r_t32.font.color.rgb = COLOR_NAVY
    p_t32.paragraph_format.space_after = Pt(8)

    add_chapter_heading(doc, "3.3 髒資料成因與清理對策總表", level=2)
    add_body_p(doc, "經由深度診斷，這份 30 筆示範檔總共潛藏以下幾項關鍵問題：")
    add_bullet_p(doc, "第 12 筆與第 5 筆（ONE8877665）、第 28 筆與第 4 筆（whl9876543）為系統拋轉重複產生之完全重複記錄，清洗後資料筆數必須從 30 筆降為 28 筆。", "1. 完全重複項（2 筆）：")
    add_bullet_p(doc, "裝貨港與卸貨港出現『高雄』、『KHH』、『Kaohsiung』、『LAX』、『Tokyo』等六種不同俗稱與縮寫，必須透過代碼表對照為 UN/LOCODE（如 TWKHH、USLAX、JPTYO）。", "2. 港口名稱多頭馬車（18 筆）：")
    add_bullet_p(doc, "『毛重』與『海運費』因夾帶『KGS』、『kgs』、『$』與『USD』，導致 Excel 視為文字，無法利用 SUM 函數進行營收與總重統計。", "3. 數值與單位黏著（24 筆）：")
    add_bullet_p(doc, "包含『2026/11/05』、『2026.11.08』與『12-Nov-2026』，格式混亂，無法以年/月進行報表群組彙整。", "4. 雜亂日期格式（15 筆）：")
    add_bullet_p(doc, "第 15 筆毛重遺漏為空值；第 22 筆報關狀態未填寫。需依業務規範進行統計填補與狀態標記。", "5. 關鍵數值缺漏（2 筆）：")

print("content_ch1_ch3.py created successfully.")
