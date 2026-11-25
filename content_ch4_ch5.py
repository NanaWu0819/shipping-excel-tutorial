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

def build_chapter_4(doc):
    add_chapter_heading(doc, "第四章：實戰操作示範——Excel 內建選單功能篇", level=1)
    
    add_body_p(doc, "在處理航運資料時，Excel 內建的圖形化選單指令具備『直覺、快速、免寫公式、立竿見影』的極大優勢。對於剛接觸資料前處理的初學者而言，掌握選單核心功能即可在三分鐘內解決超過 60% 的基礎資料品質問題。")
    add_body_p(doc, "本章將以第三章準備好的【30筆示範檔】作為實戰標的，帶領學員一步一步操作各項功能選單，逐一消滅髒資料！")

    add_chapter_heading(doc, "4.1 副本建立與工作環境準備（凍結窗格與自動篩選）", level=2)
    add_body_p(doc, "在展開任何清理前，必須嚴格落實『防護第一』原則，建立可隨時復原的備份機制：")
    
    add_bullet_p(doc, "在活頁簿底部的『01_原始資料(未清理)』標籤上按滑鼠右鍵，點選【移動或複製】。在彈出的對話框中，務必勾選左下角的【建立複本】（如圖 4-1 所示），點選確定。接著將新產生的工作表重新命名為『02_清洗示範區』。", "步驟一：建立工作表副本：")
    add_image_figure(doc, "c:/shipping_tutorial/images/fig_4_1_sheet_backup.png", 
                     "圖 4-1：Excel 核心防護技能：工作表完整備份操作視窗（務必勾選建立複本）", width_inch=5.5)

    add_bullet_p(doc, "點選功能區【檢視】索引標籤，在『視窗』群組中點選【凍結窗格】▼ 下拉按鈕，選取【凍結首列】（如圖 4-2 步驟 A）。此時第一列表頭下方會出現一條細灰色分隔線，無論向下滾動至第幾列，表頭欄位始終清晰可見，徹底避免欄位看錯的尷尬。", "步驟二：啟用首列凍結：")
    add_bullet_p(doc, "以滑鼠點選資料表內任一儲存格，切換至【資料】索引標籤，點選【篩選】按鈕（或按下鍵盤快捷鍵 Ctrl + Shift + L）。第一列各欄標題右側立即出現下拉篩選箭頭（▼，如圖 4-2 步驟 B）。點擊篩選箭頭可即時檢視該欄所包含的所有不重複值，是診斷港口名稱是否雜亂的極速探測器！", "步驟三：啟動自動篩選器：")
    add_image_figure(doc, "c:/shipping_tutorial/images/fig_4_2_freeze_panes_filter.png", 
                     "圖 4-2：常用檢視與探索功能：凍結首列與啟用自動篩選", width_inch=6.2)

    add_chapter_heading(doc, "4.2 尋找與取代 (Find & Replace)：快速清除單位字尾與貨幣符號", level=2)
    add_body_p(doc, "示範檔中『毛重』欄位含有『18,500 KGS』、『9,800 kgs』；『海運費』欄位含有『$2,850』與『$ 1,600 USD』。這些夾帶文字的數字完全無法以 SUM 函數計算。透過【尋找與取代】功能，可在十秒內批次清除所有干擾符號：")
    
    add_bullet_p(doc, "滑鼠點選欄標題 K（毛重欄），框選整欄。按下快捷鍵 Ctrl + H 開啟【尋找與取代】對話框（如圖 4-3 所示）。在『尋找目標』輸入『 KGS』（注意前面有一個半形空格），『取代為』完全保持空白（不輸入任何字元），按下【全部取代】。接著以同樣方式將小寫『 kgs』取代為空白。", "實戰操作步驟 1（清除重量單位）：")
    add_bullet_p(doc, "選取 M 欄（海運費欄），Ctrl + H 開啟對話框。在『尋找目標』輸入『$』，『取代為』保持空白，點選【全部取代】；再將『 USD』與空格取代為空白。", "實戰操作步驟 2（清除貨幣符號）：")
    add_image_figure(doc, "c:/shipping_tutorial/images/fig_4_3_find_replace.png", 
                     "圖 4-3：批次清理利器：【尋找與取代】剔除單位字尾與符號", width_inch=5.8)

    add_chapter_heading(doc, "4.3 資料剖析精靈 (Text to Columns)：一分為二拆分「船名與航次」", level=2)
    add_body_p(doc, "示範檔 C 欄將船名與航次黏在一起（如『EVER GIVEN V.0123W』或『YM UNICORN / 045E』）。在航運統計中，船名代表船舶資產，航次代表特定航程梯次，必須拆分成兩欄。此時【資料剖析精靈】是最佳解決方案：")
    
    add_bullet_p(doc, "在 C 欄與 D 欄之間，先按右鍵【插入】一欄空白欄，避免剖析後的航次直接覆蓋右側的起運港資料！將 C 欄命名為『船名 (Vessel)』，新插入欄命名為『航次 (Voyage)』。", "前置防護步驟（插入預備空白欄）：")
    add_bullet_p(doc, "選取 C 欄資料範圍，點選【資料】->【資料剖析】。在彈出的精靈視窗中依序執行（如圖 4-4 所示）：\n- 步驟 1/3：選取【分隔符號】-> 點選下一步。\n- 步驟 2/3：勾選【空格】，並在【其他】輸入『/』或『V.』-> 點選下一步。\n- 步驟 3/3：指定【目標儲存格】為 $C$2，點選【完成】！", "三步驟剖析執行：")
    add_image_figure(doc, "c:/shipping_tutorial/images/fig_4_4_text_to_columns.png", 
                     "圖 4-4：Excel 核心技能：【資料剖析精靈】三步驟拆分『船名與航次』", width_inch=6.5)

    add_chapter_heading(doc, "4.4 資料剖析精靈 (Text to Columns)：一鍵修復「文字型雜亂日期」", level=2)
    add_body_p(doc, "許多人以為【資料剖析】只能用來拆分欄位，其實它還隱藏著一個鮮為人知的特技——【修復文字型日期】！示範檔中的開航日包含『2026.11.08』與『12-Nov-2026』，如果直接用滑鼠點選儲存格格式改為日期，Excel 根本毫無反應。")
    add_body_p(doc, "只要選取日期欄位，點選【資料】->【資料剖析】，直接按兩次【下一步】進入第 3 步驟。在『欄位的資料格式』點選【日期】，下拉選單挑選【YMD】（如圖 4-5 所示），按下【完成】。Excel 會在底層自動將所有點號、橫線與英文月份字串強制重構解析為標準的日期序號，瞬間全部向右對齊！")
    add_image_figure(doc, "c:/shipping_tutorial/images/fig_4_5_date_parsing.png", 
                     "圖 4-5：日期大修復：利用【資料剖析精靈】一鍵將雜亂文字日期轉為標準日期", width_inch=5.8)

    add_chapter_heading(doc, "4.5 移除重複項 (Remove Duplicates)：提單與櫃號唯一性檢驗", level=2)
    add_body_p(doc, "示範檔中第 5 筆與第 12 筆（提單 ONE8877665）、第 4 筆與第 28 筆（提單 whl9876543）是完全重複的整列記錄。若不剔除，整份報表的總重量與海運費將虛增數萬美元！")
    
    add_bullet_p(doc, "選取整個資料表區域（按下 Ctrl + A 全選）。切換至【資料】索引標籤，在『資料工具』群組中點選【移除重複項】（如圖 4-6 步驟一）。", "執行步驟 1：")
    add_bullet_p(doc, "在彈出視窗中，務必先點選【全不選】。在航運業務中，判斷重複提單的黃金準則是：【提單號碼 (B/L No.)】與【貨櫃號碼 (Container No.)】！因此，只單獨勾選這兩欄即可（其他欄位不勾）。點選確定。", "執行步驟 2（關鍵唯一鍵設定）：")
    add_bullet_p(doc, "Excel 立即彈出提示視窗：『已找到並移除 2 個重複的值；現在還保留 28 個唯一的值』（如圖 4-6 步驟二）。點選確定，總筆數精準降為 28 筆！", "執行步驟 3（驗收成果）：")
    add_image_figure(doc, "c:/shipping_tutorial/images/fig_4_6_remove_duplicates.png", 
                     "圖 4-6：移除重複項 (Remove Duplicates) 核取方塊設定與完成提示", width_inch=6.5)

    add_chapter_heading(doc, "4.6 選擇性貼上 (Paste Special)：乘運算快速消除文字數字綠色三角", level=2)
    add_body_p(doc, "示範檔中第 3 筆的運費『2400』與第 10 筆的『3200』，其儲存格左上角出現了一顆刺眼的『綠色小三角形』，且內容靠左對齊。這是因為從 ERP 拋出時被標記為文字格式。當你使用 =SUM(M2:M29) 計算運費時，Excel 會自動忽略這兩筆文字，導致總運費少算 5,600 美元！")
    
    add_bullet_p(doc, "在工作表任何一個完全空白的儲存格中輸入數字『1』，按下 Enter。接著點選該儲存格按下 Ctrl + C 複製。", "實戰解法步驟 1：")
    add_bullet_p(doc, "框選所有運費與毛重欄位。在選取區域上按滑鼠右鍵，點選【選擇性貼上】（Paste Special）。", "實戰解法步驟 2：")
    add_bullet_p(doc, "在對話框中，『運算』區域點選【乘 (Multiply)】（如圖 4-7 所示），按下確定！", "實戰解法步驟 3：")
    add_body_p(doc, "任何數值乘上 1 都不會改變其數值大小，但 Excel 在執行乘法運算的瞬間，會被強制啟動型態轉換，將文字字串瞬間轉回真數值，綠色錯誤小三角全數消滅，所有數字瞬間乖乖靠右對齊！")
    add_image_figure(doc, "c:/shipping_tutorial/images/fig_4_7_paste_special_multiply.png", 
                     "圖 4-7：神級技巧：【選擇性貼上】乘運算瞬間消除綠色錯誤三角提示", width_inch=6.2)

    add_chapter_heading(doc, "4.7 設定格式化的條件 (Conditional Formatting)：視覺化抓出漏填值與重複", level=2)
    add_body_p(doc, "在清洗資料的驗收階段，千萬不要用肉眼一筆一筆核對。利用【設定格式化的條件】可以自動讓缺陷無所遁形（如圖 4-8 所示）：")
    add_bullet_p(doc, "選取 B 欄（提單號碼）。點選【常用】->【設定格式化的條件】->【醒目提示儲存格規則】->【重複的值】。Excel 會自動將重複的提單以淺紅色網底突顯出來。", "防呆技法 A（標示重複提單）：")
    add_bullet_p(doc, "選取整張表格，點選【常用】->【尋找與選取】->【特殊目標】-> 勾選【空格】-> 點選確定。此時全表所有的空白格會被一次性同時選取！接著在字型群組中點選填滿黃色色彩。第 15 筆毛重空白、第 22 筆報關狀態空白立即被標上刺眼的黃色警戒，方便行政人員立刻致電報關行補單！", "防呆技法 B（標記空白漏填值）：")
    add_image_figure(doc, "c:/shipping_tutorial/images/fig_4_8_conditional_formatting.png", 
                     "圖 4-8：視覺化查核：【設定格式化的條件】標記重複提單與遺漏值", width_inch=6.0)

def build_chapter_5(doc):
    add_chapter_heading(doc, "第五章：實戰操作示範——函數公式清洗篇", level=1)
    
    add_body_p(doc, "雖然第四章介紹的選單功能非常方便，但在航運實務中，我們經常每週或每天都會收到相同格式的最新單據。若每次都手動點選選單，重複勞動效率極低。此時，建立一組『可重複使用的資料清洗函數公式管線（Data Cleaning Formula Pipeline）』，就是頂尖航運行政人員的必殺技！")

    add_chapter_heading(doc, "5.1 文字清理三劍客：TRIM()、CLEAN()、UPPER() 打造標準化代碼", level=2)
    add_body_p(doc, "國際海運提單號碼（B/L No.）與櫃號（Container No.）是串聯所有系統的唯一鍵值。然而從前台網頁複製或 EDI 拋轉時，字串往往帶有多餘空白或換行符號（如圖 5-1 所示）。")
    
    add_image_figure(doc, "c:/shipping_tutorial/images/fig_5_1_trim_upper_clean.png", 
                     "圖 5-1：文字清洗核心管線：TRIM()、CLEAN() 與 UPPER() 串聯運算原理", width_inch=6.5)

    add_bullet_p(doc, "語法：=TRIM(字串)。功能：專門刪除文字開頭與結尾的所有空格。若單字中間有連續多個空格，TRIM 會智慧地將其縮減為單一空格。例如 =TRIM(\"  OOCLU9823101  \") 會回傳乾淨的 \"OOCLU9823101\"。", "1. TRIM() 函數（刪除多餘空格）：")
    add_bullet_p(doc, "語法：=CLEAN(字串)。功能：專門移除文字中無法列印的 ASCII 控制字元（代碼 0 至 31），例如隱藏的換行符號（Alt+Enter 產生的 CHAR(10)）。能徹底根除儲存格高度異常撐開的問題。", "2. CLEAN() 函數（消除不可列印字元）：")
    add_bullet_p(doc, "語法：=UPPER(字串)。功能：將所有英文字母強制轉換為大寫。例如 =UPPER(\"whl9876543\") 會轉為標準大寫 \"WHL9876543\"。", "3. UPPER() 函數（英文字母全大寫）：")
    
    add_callout(doc, "航運代碼終極黃金整合公式", 
                "在實務清理新欄位輸入：\n=UPPER(TRIM(CLEAN(B2)))\n此三合一嵌套公式可一氣呵成剔除不可見換行符、清除外圍空格、並將字母統一轉為大寫，產出 100% 完美的航運標準識別碼！", 
                "tip")

    add_chapter_heading(doc, "5.2 欄位拆分與擷取：LEFT()、RIGHT()、MID()、FIND()、LEN() 深入解析", level=2)
    add_body_p(doc, "當無法使用資料剖析或需要透過公式動態拆解編碼時，字串擷取函數組合就是無往不利的手術刀：")
    add_bullet_p(doc, "語法：=LEFT(字串, 字元數)。從最左側開始擷取指定長度。例如從 11 碼貨櫃號碼『EMCU1234567』擷取前 4 碼船公司貨櫃字首碼：=LEFT(D2, 4) -> 回傳 \"EMCU\"（代表長榮海運貨櫃）。", "LEFT() 函數（左側截取）：")
    add_bullet_p(doc, "語法：=RIGHT(字串, 字元數)。從最右側開始截取。例如擷取櫃號最後 1 碼檢查碼（Check Digit）：=RIGHT(D2, 1) -> 回傳 \"7\"。", "RIGHT() 函數（右側截取）：")
    add_bullet_p(doc, "語法：=MID(字串, 起始位置, 字元數)。從中間指定位置截取指定長度。例如擷取櫃號中間 6 位流水號：=MID(D2, 5, 6) -> 回傳 \"123456\"。", "MID() 函數（中間截取）：")
    add_bullet_p(doc, "語法：=FIND(搜尋文字, 目標字串)。回傳特定符號在字串中的字元位置。若船名航次為『WAN HAI 315 / 045E』，=FIND(\"/\", C2) 會回傳數字 13。搭配 LEFT 與 MID 即可動態拆分任何長度的船名！", "FIND() 函數（定位字元）：")

    add_chapter_heading(doc, "5.3 字串替換與正規化：SUBSTITUTE() 處理混雜分隔符號", level=2)
    add_body_p(doc, "語法：=SUBSTITUTE(目標文字, 舊字串, 新字串, [指定序號])。")
    add_body_p(doc, "不同船公司的航次編寫習慣大不相同，有的寫破折號『-』，有的寫斜線『/』，有的寫縮寫『V.』。為了統一格式，可使用 SUBSTITUTE 進行替換。例如將航次欄中的『V.』全數移除：=TRIM(SUBSTITUTE(D2, \"V.\", \"\"))。")
    add_body_p(doc, "若要將毛重文字『18,500 KGS』在公式中直接轉為純數字，可寫為：\n=SUBSTITUTE(SUBSTITUTE(K2, \" KGS\", \"\"), \",\", \"\")*1")

    add_chapter_heading(doc, "5.4 日期格式重構：DATE()、YEAR()、MONTH()、DAY() 標準化 YMD", level=2)
    add_body_p(doc, "當系統匯出的日期字串格式極為特殊（例如緊湊文字『20261105』），Excel 無法自動辨識時，必須利用 DATE 函數重新組裝：")
    add_bullet_p(doc, "年：=LEFT(\"20261105\", 4) -> 2026", "步驟 1（截取年份）：")
    add_bullet_p(doc, "月：=MID(\"20261105\", 5, 2) -> 11", "步驟 2（截取月份）：")
    add_bullet_p(doc, "日：=RIGHT(\"20261105\", 2) -> 05", "步驟 3（截取日份）：")
    add_bullet_p(doc, "=DATE(LEFT(H2, 4), MID(H2, 5, 2), RIGHT(H2, 2))。回傳 2026-11-05 標準真日期序號！接著航程天數即可直接相減：=ETA儲存格 - ETD儲存格，立即算出海上航行天數！", "步驟 4（DATE 函數重組）：")

    add_chapter_heading(doc, "5.5 數值強制轉換：VALUE() 與 NUMBERVALUE() 確保運費可計算", level=2)
    add_body_p(doc, "當文字型態的數字經過 SUBSTITUTE 去除單位後，依然是文字。若要確保其能 100% 參與加總運算，可使用 VALUE() 函數：")
    add_body_p(doc, "語法：=VALUE(文字字串)。例如 =VALUE(\"2400\") 回傳數值 2400。")
    add_body_p(doc, "在處理歐陸航線報表時，常遇到歐洲小數點以逗號表示、千分位以句點表示（例如『18.500,50』）。此時應使用進階版 =NUMBERVALUE(文字, 小數點符號, 分隔符號)，例如：\n=NUMBERVALUE(\"18.500,50\", \",\", \".\") 即可完美轉為標準數值 18500.5！")

    add_chapter_heading(doc, "5.6 遺漏值與防錯判斷：IF()、ISBLANK()、IFERROR() 補齊預設標記", level=2)
    add_body_p(doc, "航運單據中若有欄位漏填（如示範檔第 22 筆放行狀態空白），若放任空白不管，在資料透視表中會被列為『(空白)』，影響統計精確度。")
    add_bullet_p(doc, "語法：=IF(條件, 成立時回傳, 不成立時回傳)。例如檢驗狀態是否空白：\n=IF(ISBLANK(O2), \"待補單\", O2)\n若 O2 是空白，公式自動補上警示字串『待補單』；若已有值則維持原內容。", "1. IF 與 ISBLANK 缺漏填補：")
    add_bullet_p(doc, "語法：=IFERROR(原計算公式, 發生錯誤時代替值)。在跨表查詢時，若港口代碼不在對照表內，VLOOKUP 會噴出 #N/A。為了避免報表出現難看的錯誤代碼，應套上 IFERROR 防呆：\n=IFERROR(VLOOKUP(...), \"港名待人工確認\")", "2. IFERROR 防呆保護罩：")

    add_chapter_heading(doc, "5.7 跨表關聯標準化：VLOOKUP() / XLOOKUP() 參照國際港口代碼 (UN/LOCODE)", level=2)
    add_body_p(doc, "這是整份資料前處理手冊中，含金量最高、最關鍵的一項核心技能！國際港口具有聯合國標準五字代碼（UN/LOCODE），例如台灣高雄港為『TWKHH』、美國洛杉磯港為『USLAX』。但示範檔中卻充斥著『高雄』、『KHH』、『Kaohsiung』等多種寫法（如圖 5-2 所示）。")

    add_image_figure(doc, "c:/shipping_tutorial/images/fig_5_2_vlookup_port_mapping.png", 
                     "圖 5-2：代碼標準化大師：VLOOKUP / XLOOKUP 跨表自動精準關聯原理", width_inch=6.5)

    add_bullet_p(doc, "在活頁簿中建立獨立的『02_港口標準代碼表』工作表，A 欄放所有可能出現的別名（高雄、KHH、Kaohsiung），B 欄放標準 UN/LOCODE（TWKHH），C 欄放標準中文名（高雄港），D 欄放標準英文名（Kaohsiung）。", "步驟一：建構代碼主檔對照表：")
    add_bullet_p(doc, "語法：=VLOOKUP(查閱值, 對照表範圍, 回傳第幾欄, 精準或近似匹配)。在清洗工作表輸入：\n=VLOOKUP(F2, '02_港口標準代碼表'!$A$2:$C$12, 2, FALSE)\n注意：對照表範圍必須按下 F4 加上絕對鎖定符號 $A$2:$C$12，最後一個參數必須填寫 FALSE（或 0）代表完全精準匹配！", "步驟二：套用 VLOOKUP 函數：")
    add_bullet_p(doc, "若使用 Office 2021 或 Microsoft 365，推薦使用更直覺強大的新一代 XLOOKUP 函數：\n=XLOOKUP(F2, '02_港口標準代碼表'!$A$2:$A$12, '02_港口標準代碼表'!$B$2:$B$12, \"未列管港口\")\nXLOOKUP 語法更簡潔、查詢欄位不需排在最左欄，且自帶未找到時的防錯回傳值！", "步驟三：進階 XLOOKUP 新世代解法：")

    add_chapter_heading(doc, "5.8 綜合清洗管線建立與清洗成果 (Before vs. After) 總對照", level=2)
    add_body_p(doc, "經過第四章選單功能與第五章函數公式的洗禮，30 筆示範檔已全數脫胎換骨！重複項目順利剔除（降為 28 筆），所有代碼大寫且無空白，港口全數對應為 UN/LOCODE 標準五字碼，毛重與海運費轉為可計算之浮點數，日期全數標準化。以下為清洗成果關鍵指標總核驗（表 5-2）：")

    # Table 5-2: Before vs After 对照表
    tbl_5_2 = doc.add_table(rows=8, cols=4)
    tbl_5_2_widths = [1.5, 1.6, 1.6, 1.8]
    tbl_5_2_aligns = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT]

    headers_5_2 = ["驗收指標", "清洗前 (01_原始資料)", "清洗後 (03_清洗完成標準)", "品質提升成果判定"]
    for idx, h in enumerate(headers_5_2):
        tbl_5_2.cell(0, idx).paragraphs[0].text = h

    before_after_data = [
        ("總資料筆數", "30 筆記錄", "28 筆獨立記錄", "成功移除第 12 與 28 筆重複提單 (合格)"),
        ("船名航次獨立性", "混雜在同一欄 (C欄)", "拆分為『船名』與『航次』兩欄", "已建立單一屬性二維表格結構 (合格)"),
        ("港口代碼標準化率", "33.3% (中文俗稱混用)", "100.0% (全數為 UN/LOCODE)", "VLOOKUP 跨表自動關聯成功 (合格)"),
        ("總毛重加總運算", "無法計算 (SUM回傳0)", "445,400.0 KG", "清除 KGS 單位，空值完成補正 (合格)"),
        ("總海運費營收加總", "無法計算 (夾帶貨幣字串)", "$ 79,400.00 USD", "乘 1 消除綠三角，數值成功加總 (合格)"),
        ("日期格式標準化率", "30.0% (格式多元錯亂)", "100.0% (YYYY-MM-DD 序號)", "資料剖析轉換為真日期序號 (合格)"),
        ("平均海上航程天數", "無法計算 (文字無法相減)", "21.6 天", "ETA - ETD 成功計算出運渡時效 (合格)")
    ]
    for r_idx, row_vals in enumerate(before_after_data, start=1):
        for c_idx, val in enumerate(row_vals):
            tbl_5_2.cell(r_idx, c_idx).paragraphs[0].text = val

    style_custom_table(tbl_5_2, tbl_5_2_widths, tbl_5_2_aligns)

    p_t52 = doc.add_paragraph()
    p_t52.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_t52 = p_t52.add_run("表 5-2：30筆示範檔清洗前後 (Before vs. After) 數據完整對照驗收表")
    r_t52.font.size = Pt(9.0)
    r_t52.font.bold = True
    r_t52.font.color.rgb = COLOR_NAVY
    p_t52.paragraph_format.space_after = Pt(8)

print("content_ch4_ch5.py created successfully.")
