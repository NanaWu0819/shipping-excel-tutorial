# -*- coding: utf-8 -*-
import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# Initialize 16:9 widescreen presentation
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

blank_slide_layout = prs.slide_layouts[6]

# Theme Colors
C_DARK_BG   = RGBColor(10, 25, 47)       # #0A192F Deep Navy Background
C_CARD_BG   = RGBColor(23, 37, 65)       # #172541 Card Background
C_CARD_BDR  = RGBColor(30, 58, 102)      # #1E3A66 Card Border
C_CYAN      = RGBColor(6, 182, 212)      # #06B6D4 Vibrant Cyan
C_TEAL      = RGBColor(13, 148, 136)     # #0D9488 Ocean Teal
C_BLUE      = RGBColor(59, 130, 246)     # #3B82F6 Blue
C_AMBER     = RGBColor(245, 158, 11)     # #F59E0B Warning Amber
C_ROSE      = RGBColor(244, 63, 94)      # #F43F5E Alert Rose
C_EMERALD   = RGBColor(16, 185, 129)     # #10B981 Success Green
C_WHITE     = RGBColor(255, 255, 255)    # Text Primary
C_LIGHT_GRY = RGBColor(203, 213, 225)    # Text Secondary
C_MUTED     = RGBColor(148, 163, 184)    # Text Muted

FONT_MAIN = "微軟正黑體"
FONT_CODE = "Consolas"
IMG_DIR = r"c:\shipping_tutorial\images"

def add_base_slide(prs):
    slide = prs.slides.add_slide(blank_slide_layout)
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = C_DARK_BG
    bg.line.fill.background()
    return slide

def add_header(slide, category_text, title_text, slide_num, total_slides=14):
    # Top banner bar
    top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.4), Inches(11.733), Inches(0.9))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = RGBColor(15, 31, 58)
    top_bar.line.color.rgb = C_CARD_BDR
    top_bar.line.width = Pt(1)

    # Category Badge
    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(0.55), Inches(1.8), Inches(0.32))
    badge.fill.solid()
    badge.fill.fore_color.rgb = RGBColor(6, 44, 73)
    badge.line.color.rgb = C_CYAN
    badge.line.width = Pt(1)
    tf_b = badge.text_frame
    tf_b.word_wrap = False
    p_b = tf_b.paragraphs[0]
    p_b.alignment = PP_ALIGN.CENTER
    r_b = p_b.add_run()
    r_b.text = category_text
    r_b.font.name = FONT_MAIN
    r_b.font.size = Pt(10)
    r_b.font.bold = True
    r_b.font.color.rgb = C_CYAN

    # Title Text
    tx_box = slide.shapes.add_textbox(Inches(3.0), Inches(0.45), Inches(8.0), Inches(0.5))
    tf = tx_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = title_text
    r.font.name = FONT_MAIN
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = C_WHITE

    # Slide Number
    num_box = slide.shapes.add_textbox(Inches(11.0), Inches(0.5), Inches(1.4), Inches(0.4))
    tf_n = num_box.text_frame
    p_n = tf_n.paragraphs[0]
    p_n.alignment = PP_ALIGN.RIGHT
    r_n = p_n.add_run()
    r_n.text = f"{slide_num:02d} / {total_slides:02d}"
    r_n.font.name = FONT_CODE
    r_n.font.size = Pt(12)
    r_n.font.bold = True
    r_n.font.color.rgb = C_MUTED

def add_card(slide, left, top, width, height, title=None, title_color=C_CYAN, bg_color=C_CARD_BG, border_color=C_CARD_BDR):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = border_color
    card.line.width = Pt(1.2)
    
    if title:
        tx_box = slide.shapes.add_textbox(left + Inches(0.15), top + Inches(0.12), width - Inches(0.3), Inches(0.4))
        tf = tx_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = title
        r.font.name = FONT_MAIN
        r.font.size = Pt(12)
        r.font.bold = True
        r.font.color.rgb = title_color
    return card

# -------------------------------------------------------------
# SLIDE 1: Title & Cover
# -------------------------------------------------------------
s1 = add_base_slide(prs)

# Decorative background glow circle
glow = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(3.66), Inches(0.8), Inches(6.0), Inches(5.5))
glow.fill.solid()
glow.fill.fore_color.rgb = RGBColor(12, 35, 70)
glow.line.fill.background()

# Badge
badge1 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.16), Inches(1.2), Inches(5.0), Inches(0.4))
badge1.fill.solid()
badge1.fill.fore_color.rgb = RGBColor(6, 44, 73)
badge1.line.color.rgb = C_CYAN
badge1.line.width = Pt(1.5)
p_b1 = badge1.text_frame.paragraphs[0]
p_b1.alignment = PP_ALIGN.CENTER
r_b1 = p_b1.add_run()
r_b1.text = "全球海運物流暨供應鏈數位賦能系列教材"
r_b1.font.name = FONT_MAIN
r_b1.font.size = Pt(11)
r_b1.font.bold = True
r_b1.font.color.rgb = C_CYAN

# Main Title
t_box = s1.shapes.add_textbox(Inches(1.5), Inches(1.8), Inches(10.33), Inches(1.8))
tf = t_box.text_frame
tf.word_wrap = True
p1 = tf.paragraphs[0]
p1.alignment = PP_ALIGN.CENTER
r1 = p1.add_run()
r1.text = "航運行政人員必修：\nMS Excel 基礎資料分析前處理實戰指南"
r1.font.name = FONT_MAIN
r1.font.size = Pt(32)
r1.font.bold = True
r1.font.color.rgb = C_WHITE

# Subtitle
sub_box = s1.shapes.add_textbox(Inches(2.0), Inches(3.8), Inches(9.33), Inches(1.0))
tf_sub = sub_box.text_frame
tf_sub.word_wrap = True
p_sub = tf_sub.paragraphs[0]
p_sub.alignment = PP_ALIGN.CENTER
r_sub = p_sub.add_run()
r_sub.text = "從零打底！海運提單、貨櫃清單、港口代碼與費用數據清洗的核心操作全攻略\n聚焦於資料清洗常識、內建選單功能、關鍵函數管線與多情境實戰驗收"
r_sub.font.name = FONT_MAIN
r_sub.font.size = Pt(13)
r_sub.font.color.rgb = C_LIGHT_GRY

# 3 Info Cards
info_items = [
    ("適用對象", "航運公司、船務代理、貨運承攬初學者", C_CYAN),
    ("示範工具", "MS Excel 2016 / 2019 / 2021 / 365", C_TEAL),
    ("配套實習", "30 筆示範檔 ＆ 300 筆多情境挑戰案例", C_BLUE)
]
for idx, (label, val, col) in enumerate(info_items):
    c = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(2.2 + idx * 3.1), Inches(5.1), Inches(2.8), Inches(1.1))
    c.fill.solid()
    c.fill.fore_color.rgb = C_CARD_BG
    c.line.color.rgb = col
    c.line.width = Pt(1.2)
    tf_c = c.text_frame
    tf_c.word_wrap = True
    p_l = tf_c.paragraphs[0]
    p_l.alignment = PP_ALIGN.CENTER
    r_l = p_l.add_run()
    r_l.text = label
    r_l.font.name = FONT_MAIN
    r_l.font.size = Pt(10)
    r_l.font.bold = True
    r_l.font.color.rgb = col
    
    p_v = tf_c.add_paragraph()
    p_v.alignment = PP_ALIGN.CENTER
    r_v = p_v.add_run()
    r_v.text = val
    r_v.font.name = FONT_MAIN
    r_v.font.size = Pt(9.5)
    r_v.font.color.rgb = C_WHITE

# Footer info
f_box = s1.shapes.add_textbox(Inches(1.5), Inches(6.6), Inches(10.33), Inches(0.5))
p_f = f_box.text_frame.paragraphs[0]
p_f.alignment = PP_ALIGN.CENTER
r_f = p_f.add_run()
r_f.text = "2026 年 11 月專業版  |  航運數位化轉型與商業智慧數據推動小組編撰"
r_f.font.name = FONT_MAIN
r_f.font.size = Pt(10)
r_f.font.color.rgb = C_MUTED

# -------------------------------------------------------------
# SLIDE 2: 學習地圖與四大航運資料災難
# -------------------------------------------------------------
s2 = add_base_slide(prs)
add_header(s2, "學習導讀", "為什麼航運行政人員必須精通「資料前處理」？", 2)

# Left Card: 四大航運災難
add_card(s2, Inches(0.8), Inches(1.5), Inches(5.7), Inches(4.5), "⚠️ 航運實務常見四大資料災難 (Data Disasters)", C_ROSE, RGBColor(28, 20, 35), C_ROSE)
pain_box = s2.shapes.add_textbox(Inches(0.95), Inches(2.0), Inches(5.4), Inches(3.8))
tf_p = pain_box.text_frame
tf_p.word_wrap = True

pains = [
    ("1. VGM 申報不合規被拒裝船：", "毛重欄位夾帶文字「18,500 KGS」，Excel 視為文字無法加總，造成向船公司申報錯誤，貨櫃被碼頭拒裝！"),
    ("2. 跨表查詢失效導致貨物錯運：", "提單號碼首尾包夾隱藏空白，VLOOKUP 回傳 #N/A 錯誤，代碼對應失誤造成航線錯配。"),
    ("3. 重複放行引發重大索賠：", "EDI 系統重複傳送提單，未經重複項排查導致向客戶重複請款，甚至二次放行提單。"),
    ("4. 合併儲存格使樞紐分析表崩潰：", "船名跨欄置中合併，導致排序跳出警告錯誤，自動篩選時非首列資料全數消失！")
]
for title, desc in pains:
    p_t = tf_p.add_paragraph()
    r_t = p_t.add_run()
    r_t.text = title
    r_t.font.name = FONT_MAIN
    r_t.font.size = Pt(10.5)
    r_t.font.bold = True
    r_t.font.color.rgb = C_ROSE
    
    p_d = tf_p.add_paragraph()
    r_d = p_d.add_run()
    r_d.text = desc
    r_d.font.name = FONT_MAIN
    r_d.font.size = Pt(9.5)
    r_d.font.color.rgb = C_LIGHT_GRY

# Right Card: 四階段學習地圖
add_card(s2, Inches(6.8), Inches(1.5), Inches(5.7), Inches(4.5), "🧭 四階段循序漸進學習地圖 (Learning Roadmap)", C_CYAN)
road_box = s2.shapes.add_textbox(Inches(6.95), Inches(2.0), Inches(5.4), Inches(3.8))
tf_r = road_box.text_frame
tf_r.word_wrap = True

roadmap = [
    ("第一階段：環境佈局 (第一章)", "熟悉 Excel 八大區域，掌握座標系統與相對/絕對參照鎖定。"),
    ("第二階段：常識守則 (第二章)", "掌握 GIGO 核心觀念、Tidy Data 二維表格架構與黃金五步 SOP。"),
    ("第三階段：實戰打底 (第三至五章)", "以 30 筆示範檔演練選單指令（凍結、取代、剖析）與函數清洗管線。"),
    ("第四階段：情境通關 (第六至七章)", "挑戰 300 筆多情境案例（FCL/LCL/冷藏/危險品），以 KPI 對帳驗收。")
]
for title, desc in roadmap:
    p_t = tf_r.add_paragraph()
    r_t = p_t.add_run()
    r_t.text = title
    r_t.font.name = FONT_MAIN
    r_t.font.size = Pt(10.5)
    r_t.font.bold = True
    r_t.font.color.rgb = C_CYAN
    
    p_d = tf_r.add_paragraph()
    r_d = p_d.add_run()
    r_d.text = desc
    r_d.font.name = FONT_MAIN
    r_d.font.size = Pt(9.5)
    r_d.font.color.rgb = C_LIGHT_GRY

# Bottom GIGO Banner
b_box = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.15), Inches(11.733), Inches(0.8))
b_box.fill.solid()
b_box.fill.fore_color.rgb = RGBColor(12, 38, 64)
b_box.line.color.rgb = C_TEAL
p_g = b_box.text_frame.paragraphs[0]
p_g.alignment = PP_ALIGN.CENTER
r_g = p_g.add_run()
r_g.text = "💡 航運前處理金律：Garbage In, Garbage Out（垃圾進，垃圾出）。唯有標準化的整潔資料，才能產出值得信賴的決策報表！"
r_g.font.name = FONT_MAIN
r_g.font.size = Pt(10.5)
r_g.font.bold = True
r_g.font.color.rgb = C_TEAL

# -------------------------------------------------------------
# SLIDE 3: 第一章 MS Excel 八大核心操作區域導覽
# -------------------------------------------------------------
s3 = add_base_slide(prs)
add_header(s3, "第一章：介面導覽", "MS Excel 操作介面全景導覽 (八大核心操作區域)", 3)

# Left Column: 8 Zones
add_card(s3, Inches(0.8), Inches(1.5), Inches(5.8), Inches(5.5), "八大區域功能與航運實務操作速記", C_CYAN)
z_box = s3.shapes.add_textbox(Inches(0.95), Inches(1.95), Inches(5.5), Inches(4.9))
tf_z = z_box.text_frame
tf_z.word_wrap = True

zones = [
    ("① 標題列與存取列：", "顯示檔名，必設自動儲存與 Ctrl+Z 復原救命鍵。"),
    ("② 功能區 (Ribbon)：", "常用、公式、資料三大重裝標籤，Ctrl+F1 可收发展開。"),
    ("③ 名稱方塊：", "顯示格位座標，輸入 A1:Q301 瞬間全選三萬筆資料。"),
    ("④ 資料編輯列：", "顯微照妖鏡，可看穿字串最前方是否有偷塞隱藏空白。"),
    ("⑤ 欄號列號座標：", "掌握參照原理，按 F4 一鍵鎖定為絕對參照 $A$2:$C$12。"),
    ("⑥ 儲存格網格：", "作用儲存格；雙擊右下角黑十字填滿控點，公式瞬間複製！"),
    ("⑦ 工作表標籤：", "右鍵【移動或複製】建立備份，標籤填色區分原始與清洗區。"),
    ("⑧ 狀態列：", "框選儲存格秒看「平均、計數、加總」，客戶電話對帳神技！")
]
for title, desc in zones:
    p = tf_z.add_paragraph()
    r_t = p.add_run()
    r_t.text = title
    r_t.font.name = FONT_MAIN
    r_t.font.size = Pt(9.5)
    r_t.font.bold = True
    r_t.font.color.rgb = C_CYAN
    r_d = p.add_run()
    r_d.text = desc
    r_d.font.name = FONT_MAIN
    r_d.font.size = Pt(9.0)
    r_d.font.color.rgb = C_LIGHT_GRY

# Right Column: UI Diagram Image
img_path_ui = os.path.join(IMG_DIR, "fig_1_1_excel_ui.png")
if os.path.exists(img_path_ui):
    add_card(s3, Inches(6.8), Inches(1.5), Inches(5.733), Inches(5.5), "介面各區塊全覽標註圖解", C_TEAL)
    s3.shapes.add_picture(img_path_ui, Inches(6.95), Inches(2.0), width=Inches(5.433))

# -------------------------------------------------------------
# SLIDE 4: 第二章 前處理三大不可動搖守則
# -------------------------------------------------------------
s4 = add_base_slide(prs)
add_header(s4, "第二章：核心規範", "航運資料前處理「三大不可動搖守則」", 4)

rules_data = [
    ("守則一：原始資料保護原則", "建立工作表複本，永不在 Raw Data 上直接動刀！", "資料清洗的第一鐵律。匯入後立即右鍵【建立複本】命名為 01_原始資料(未清理)。所有拆分修改只在複本執行，確保清洗失誤時隨時能核對復原。", C_BLUE),
    ("守則二：標準二維表格 (Tidy Data)", "一列一筆記錄、一行一屬性、單一標題列、嚴禁合併！", "Hadley Wickham 現代資料整潔標準。禁止跨欄置中合併儲存格！合併格只有左上角有值，會使排序、自動篩選與樞紐分析表全面失效！", C_ROSE),
    ("守則三：型態判別與對齊法則", "文字預設靠左，數值與真日期預設靠右！", "日期底層本質是「數值序號」，靠左代表是無法計算的文字！數值靠右才能使用 SUM 加總，靠左代表文字型數字，需乘 1 消除綠色錯誤三角！", C_EMERALD)
]

for idx, (title, highlight, desc, col) in enumerate(rules_data):
    card = add_card(s4, Inches(0.8 + idx * 3.95), Inches(1.5), Inches(3.8), Inches(2.6), title, col)
    tb = s4.shapes.add_textbox(Inches(0.95 + idx * 3.95), Inches(2.0), Inches(3.5), Inches(1.9))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p_h = tf.paragraphs[0]
    r_h = p_h.add_run()
    r_h.text = highlight + "\n"
    r_h.font.name = FONT_MAIN
    r_h.font.size = Pt(10)
    r_h.font.bold = True
    r_h.font.color.rgb = col
    
    p_d = tf.add_paragraph()
    r_d = p_d.add_run()
    r_d.text = desc
    r_d.font.name = FONT_MAIN
    r_d.font.size = Pt(9.0)
    r_d.font.color.rgb = C_LIGHT_GRY

# Bottom Image: Tidy Data comparison
img_path_tidy = os.path.join(IMG_DIR, "fig_2_2_tidy_data.png")
if os.path.exists(img_path_tidy):
    add_card(s4, Inches(0.8), Inches(4.25), Inches(11.733), Inches(2.8), "Tidy Data 整潔資料結構 vs 常見錯誤結構對照圖", C_CYAN)
    s4.shapes.add_picture(img_path_tidy, Inches(0.95), Inches(4.7), width=Inches(11.433))

# -------------------------------------------------------------
# SLIDE 5: 第二章 航運常見八大「髒資料」特徵診斷
# -------------------------------------------------------------
s5 = add_base_slide(prs)
add_header(s5, "第二章：缺陷診斷", "航運常見八大「髒資料」特徵診斷與對策表", 5)

dirty_items_pptx = [
    ("1. 前後多餘空格", "\"  OOCLU9823101  \"", "VLOOKUP 查無代碼", "=TRIM() 函數清除首尾空格", C_CYAN),
    ("2. 不可見換行符號", "EDI 帶入 ASCII 0~31", "儲存格撐高、XML破圖", "=CLEAN() 清除控制字元", C_TEAL),
    ("3. 英文大小寫混用", "whl1234 vs WHL1234", "樞紐分析表視為不同類別", "=UPPER() 強制轉換為全大寫", C_BLUE),
    ("4. 港口俗稱縮寫混用", "高雄、KHH、LAX、TYO", "無法按國際航線港口統計", "VLOOKUP 參照代碼表轉 UN碼", C_AMBER),
    ("5. 數值黏雜文字單位", "18,500 KGS / $2,850 USD", "判為文字，SUM 加總回傳 0", "Ctrl+H 尋找取代剔除單位", C_ROSE),
    ("6. 綠色三角文字數字", "'2400 (靠左對齊無法SUM)", "加總自動忽略，運費短少", "選擇性貼上【乘 1】轉回真數值", C_EMERALD),
    ("7. 雜亂日期格式", "2026.11.08 / 12-Nov-2026", "無法群組統計，無法算天數", "【資料剖析】選 YMD 轉序號", C_CYAN),
    ("8. 重複拋轉提單", "同一櫃號出現兩次以上", "重複請款爭議，雙重放行", "【移除重複項】勾提單+櫃號", C_ROSE)
]

for idx, (title, ex, impact, sol, col) in enumerate(dirty_items_pptx):
    r_idx = idx // 4
    c_idx = idx % 4
    left = Inches(0.8 + c_idx * 2.98)
    top = Inches(1.5 + r_idx * 2.7)
    
    add_card(s5, left, top, Inches(2.85), Inches(2.55), title, col)
    tb = s5.shapes.add_textbox(left + Inches(0.1), top + Inches(0.45), Inches(2.65), Inches(2.0))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p1 = tf.paragraphs[0]
    r1 = p1.add_run()
    r1.text = "範例：" + ex + "\n"
    r1.font.name = FONT_CODE
    r1.font.size = Pt(8.5)
    r1.font.color.rgb = col
    
    p2 = tf.add_paragraph()
    r2 = p2.add_run()
    r2.text = "衝擊：" + impact + "\n"
    r2.font.name = FONT_MAIN
    r2.font.size = Pt(8.5)
    r2.font.color.rgb = C_LIGHT_GRY
    
    p3 = tf.add_paragraph()
    r3 = p3.add_run()
    r3.text = "解法：" + sol
    r3.font.name = FONT_MAIN
    r3.font.size = Pt(8.5)
    r3.font.bold = True
    r3.font.color.rgb = C_WHITE

# -------------------------------------------------------------
# SLIDE 6: 第二章 黃金五步標準前處理 SOP
# -------------------------------------------------------------
s6 = add_base_slide(prs)
add_header(s6, "第二章：標準流程", "航運資料前處理黃金標準作業流程 (SOP Pipeline)", 6)

sop_steps = [
    ("第 1 步：防護備份", "原始工作表複製", "建立 Raw Data 複本並鎖定原檔，命名為 01_原始資料，嚴禁直接動刀修改。", C_BLUE),
    ("第 2 步：結構檢驗", "唯一性與凍結窗格", "凍結首列保持標題可見，啟用自動篩選，依【提單+櫃號】雙鍵值移除重複項。", C_TEAL),
    ("第 3 步：字串代碼", "TRIM & 代碼標準化", "TRIM 清空格，UPPER 轉大寫，VLOOKUP 跨表將港口俗稱正規化為 UN/LOCODE。", C_CYAN),
    ("第 4 步：數值日期", "剔除單位與日期修復", "取代 KGS/USD 單位，乘 1 消除綠三角轉數值，資料剖析轉 YMD 標準日期序號。", C_AMBER),
    ("第 5 步：遺漏封存", "缺漏補齊與品質驗收", "IF/ISBLANK 補齊狀態標記，條件式格式抓漏，核對 KPI 加總並封存標準庫。", C_EMERALD)
]

for idx, (step_no, title, desc, col) in enumerate(sop_steps):
    left = Inches(0.8 + idx * 2.38)
    add_card(s6, left, Inches(1.5), Inches(2.26), Inches(2.3), step_no, col)
    tb = s6.shapes.add_textbox(left + Inches(0.1), Inches(1.95), Inches(2.06), Inches(1.7))
    tf = tb.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    r1 = p1.add_run()
    r1.text = title + "\n"
    r1.font.name = FONT_MAIN
    r1.font.size = Pt(10)
    r1.font.bold = True
    r1.font.color.rgb = C_WHITE
    
    p2 = tf.add_paragraph()
    r2 = p2.add_run()
    r2.text = desc
    r2.font.name = FONT_MAIN
    r2.font.size = Pt(8.5)
    r2.font.color.rgb = C_LIGHT_GRY

# Bottom Image: SOP visual
img_path_sop = os.path.join(IMG_DIR, "fig_2_1_preprocessing_sop.png")
if os.path.exists(img_path_sop):
    add_card(s6, Inches(0.8), Inches(4.0), Inches(11.733), Inches(3.0), "資料前處理標準作業管線流程圖解", C_CYAN)
    s6.shapes.add_picture(img_path_sop, Inches(0.95), Inches(4.4), width=Inches(11.433))

# -------------------------------------------------------------
# SLIDE 7: 第三章 30筆示範檔深度診斷
# -------------------------------------------------------------
s7 = add_base_slide(prs)
add_header(s7, "第三章：示範資料", "30 筆真實航運示範資料集欄位設計與問題診斷", 7)

# Left Box: 15 Fields summary
add_card(s7, Inches(0.8), Inches(1.5), Inches(4.0), Inches(5.5), "15 大關鍵欄位規格定義", C_CYAN)
f_box = s7.shapes.add_textbox(Inches(0.95), Inches(1.95), Inches(3.7), Inches(4.9))
tf_f = f_box.text_frame
tf_f.word_wrap = True

field_cats = [
    ("【核心識別類】", "項次、提單號碼 (B/L No.)、船名航次、貨櫃號碼 (11碼 ISO)、櫃型 (20GP/40HQ)"),
    ("【航運路線類】", "裝貨港 (POL)、卸貨港 (POD) -> 需轉 UN/LOCODE；開航日 (ETD)、抵達日 (ETA) -> 需轉 YMD"),
    ("【物流計費類】", "件數 (Package)、毛重 (Gross Weight)、體積 (CBM)、海運費 (USD) -> 需剔除文字單位"),
    ("【管理狀態類】", "託運人名稱 (Shipper) -> 需轉全大寫；報關放行狀態 -> 需填補遺漏空值")
]
for cat, content in field_cats:
    p1 = tf_f.add_paragraph()
    r1 = p1.add_run()
    r1.text = cat
    r1.font.name = FONT_MAIN
    r1.font.size = Pt(10)
    r1.font.bold = True
    r1.font.color.rgb = C_CYAN
    
    p2 = tf_f.add_paragraph()
    r2 = p2.add_run()
    r2.text = content + "\n"
    r2.font.name = FONT_MAIN
    r2.font.size = Pt(8.8)
    r2.font.color.rgb = C_LIGHT_GRY

# Right Table: Sample Raw Rows
add_card(s7, Inches(5.0), Inches(1.5), Inches(7.533), Inches(5.5), "30 筆示範檔原始資料樣本與缺陷標註", C_TEAL)

# Add Table for sample rows
rows_cnt = 8
cols_cnt = 6
tbl_shape = s7.shapes.add_table(rows_cnt, cols_cnt, Inches(5.15), Inches(2.0), Inches(7.233), Inches(4.8))
tbl = tbl_shape.table
tbl.columns[0].width = Inches(0.5)
tbl.columns[1].width = Inches(1.5)
tbl.columns[2].width = Inches(1.5)
tbl.columns[3].width = Inches(0.9)
tbl.columns[4].width = Inches(1.1)
tbl.columns[5].width = Inches(1.733)

headers_sample = ["No", "提單號碼", "船名航次", "裝/卸港", "毛重/運費", "埋設之缺陷特徵"]
for c_idx, h in enumerate(headers_sample):
    cell = tbl.cell(0, c_idx)
    cell.fill.solid()
    cell.fill.fore_color.rgb = RGBColor(12, 44, 78)
    p = cell.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = h
    r.font.name = FONT_MAIN
    r.font.size = Pt(9.5)
    r.font.bold = True
    r.font.color.rgb = C_WHITE

sample_rows_data = [
    ("1", "  OOCLU9823101  ", "EVER GIVEN V.0123W", "TWKHH/USLAX", "18,500 KGS / $2,850", "前後空格、單位黏雜"),
    ("2", "EMCE8901234", "YM UNICORN / 045E", "高雄 / LAX", "9,800 kgs / $1,600", "港口俗稱、點號日期"),
    ("3", "MSK7654321", "OOCL POLAND - 102N", "KHH / Los Angeles", "14,200 / 2400 (文字)", "文字數字綠三角"),
    ("4", "whl9876543", "WAN HAI 315 V.S088", "TWKEL / Tokyo", "21500 KGS / $3,100", "小寫提單 (與28重複)"),
    ("5", "ONE8877665", "ONE HARBOR V.022W", "基隆 / DEHAM", "16,800 KGS / $3,850", "★與第12筆完全重複！"),
    ("15", "OOCLU9823103", "OOCL SHANGHAI V.078N", "KHH / SGSIN", "[空白漏填!] / 2750", "★關鍵毛重遺漏缺失！"),
    ("22", "OOCLU9823104", "OOCL GERMANY V.003E", "TWTPE / Hamburg", "20,100 KGS / 4050", "★報關放行狀態缺失！")
]

for r_idx, rvals in enumerate(sample_rows_data, start=1):
    bg_c = RGBColor(35, 20, 30) if "★" in rvals[5] else (RGBColor(18, 30, 52) if r_idx % 2 == 0 else RGBColor(14, 24, 42))
    for c_idx, val in enumerate(rvals):
        cell = tbl.cell(r_idx, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg_c
        p = cell.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = val
        r.font.name = FONT_CODE if c_idx in [1, 2, 4] else FONT_MAIN
        r.font.size = Pt(8.5)
        r.font.color.rgb = C_ROSE if "★" in val else C_LIGHT_GRY

# -------------------------------------------------------------
# SLIDE 8: 第四章 內建選單功能篇 (一)
# -------------------------------------------------------------
s8 = add_base_slide(prs)
add_header(s8, "第四章：選單操作", "實戰示範——Excel 內建選單清理技巧 (一)", 8)

col_cards = [
    ("功能 1：凍結首列與啟用自動篩選", "選單路徑：【檢視】->【凍結首列】\n快捷鍵：Ctrl + Shift + L 啟用篩選器下拉鈕。\n滾動資料時欄位永不消失，篩選選單可秒速檢視所有俗稱碼與空格異常！", "fig_4_2_freeze_panes_filter.png", C_BLUE),
    ("功能 2：尋找與取代清除單位符號", "選單路徑：選取欄位 -> 按下 Ctrl + H\n尋找目標輸入「 KGS」、「 kgs」、「$」、「 USD」\n取代為完全保持空白，點選【全部取代】，十秒徹底將文字字串還原為純數值！", "fig_4_3_find_replace.png", C_TEAL),
    ("功能 3：資料剖析拆分「船名與航次」", "選單路徑：【資料】->【資料剖析】\n步驟 1 選【分隔符號】-> 步驟 2 勾【空格】或填斜線『/』-> 步驟 3 指定新欄位 C5:D5。\n一分為二獨立為「船名」與「航次」！", "fig_4_4_text_to_columns.png", C_CYAN)
]

for idx, (title, desc, img_name, col) in enumerate(col_cards):
    left = Inches(0.8 + idx * 3.95)
    add_card(s8, left, Inches(1.5), Inches(3.8), Inches(5.5), title, col)
    tb = s8.shapes.add_textbox(left + Inches(0.15), Inches(2.0), Inches(3.5), Inches(1.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = desc
    r.font.name = FONT_MAIN
    r.font.size = Pt(9.5)
    r.font.color.rgb = C_LIGHT_GRY
    
    img_p = os.path.join(IMG_DIR, img_name)
    if os.path.exists(img_p):
        s8.shapes.add_picture(img_p, left + Inches(0.15), Inches(3.9), width=Inches(3.5))

# -------------------------------------------------------------
# SLIDE 9: 第四章 內建選單功能篇 (二)
# -------------------------------------------------------------
s9 = add_base_slide(prs)
add_header(s9, "第四章：選單操作", "實戰示範——Excel 內建選單清理技巧 (二)", 9)

col_cards_2 = [
    ("功能 4：資料剖析修復「文字日期」", "選單路徑：【資料】->【資料剖析】\n連按兩次下一步進入步驟 3 -> 點選【日期】下拉挑選【YMD】-> 完成！\nExcel 在底層自動解析點號、橫線與英文月份，瞬間修復為靠右真日期序號！", "fig_4_5_date_parsing.png", C_AMBER),
    ("功能 5：移除重複項 (雙鍵值排查)", "選單路徑：Ctrl+A 全選 ->【資料】->【移除重複項】\n點選【全不選】，僅單獨勾選【提單號碼】與【貨櫃號碼】雙唯一鍵值！\n系統自動剔除 2 筆完全重複列，保留 28 筆獨立記錄！", "fig_4_6_remove_duplicates.png", C_ROSE),
    ("功能 6：選擇性貼上(乘1) 消除綠三角", "神級密技：在任一空格輸入「1」按下 Ctrl+C 複製。\n框選運費與毛重範圍 -> 右鍵【選擇性貼上】-> 運算選【乘】！\n瞬間強制將靠左文字型數字轉為右對齊真數值，綠色錯誤小三角全數消滅！", "fig_4_7_paste_special_multiply.png", C_EMERALD)
]

for idx, (title, desc, img_name, col) in enumerate(col_cards_2):
    left = Inches(0.8 + idx * 3.95)
    add_card(s9, left, Inches(1.5), Inches(3.8), Inches(5.5), title, col)
    tb = s9.shapes.add_textbox(left + Inches(0.15), Inches(2.0), Inches(3.5), Inches(1.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = desc
    r.font.name = FONT_MAIN
    r.font.size = Pt(9.5)
    r.font.color.rgb = C_LIGHT_GRY
    
    img_p = os.path.join(IMG_DIR, img_name)
    if os.path.exists(img_p):
        s9.shapes.add_picture(img_p, left + Inches(0.15), Inches(3.9), width=Inches(3.5))

# -------------------------------------------------------------
# SLIDE 10: 第五章 函數清洗管線篇 (一)
# -------------------------------------------------------------
s10 = add_base_slide(prs)
add_header(s10, "第五章：函數管線", "實戰示範——核心清洗函數與公式管線 (一)", 10)

# Left: Text Trio
add_card(s10, Inches(0.8), Inches(1.5), Inches(5.7), Inches(5.5), "文字清洗三劍客：TRIM + CLEAN + UPPER", C_CYAN)
tb10_l = s10.shapes.add_textbox(Inches(0.95), Inches(2.0), Inches(5.4), Inches(2.2))
tf10_l = tb10_l.text_frame
tf10_l.word_wrap = True

trio = [
    ("• TRIM(text)：", "刪除字串首尾所有多餘空格，中間多個空格縮減為單一格。"),
    ("• CLEAN(text)：", "消除不可列印之 ASCII 0~31 控制字元（如 Alt+Enter 換行符）。"),
    ("• UPPER(text)：", "將所有英文字母強制轉換為標準大寫代碼。")
]
for t, d in trio:
    p = tf10_l.add_paragraph()
    r_t = p.add_run()
    r_t.text = t
    r_t.font.name = FONT_MAIN
    r_t.font.size = Pt(10)
    r_t.font.bold = True
    r_t.font.color.rgb = C_CYAN
    r_d = p.add_run()
    r_d.text = d
    r_d.font.name = FONT_MAIN
    r_d.font.size = Pt(9.5)
    r_d.font.color.rgb = C_LIGHT_GRY

# Formula callout
fc = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.95), Inches(3.5), Inches(5.4), Inches(0.75))
fc.fill.solid()
fc.fill.fore_color.rgb = RGBColor(6, 44, 73)
fc.line.color.rgb = C_CYAN
p_fc = fc.text_frame.paragraphs[0]
p_fc.alignment = PP_ALIGN.CENTER
r_fc = p_fc.add_run()
r_fc.text = "終極單行整合公式： =UPPER(TRIM(CLEAN(B2)))"
r_fc.font.name = FONT_CODE
r_fc.font.size = Pt(11)
r_fc.font.bold = True
r_fc.font.color.rgb = C_CYAN

img_path_trio = os.path.join(IMG_DIR, "fig_5_1_trim_upper_clean.png")
if os.path.exists(img_path_trio):
    s10.shapes.add_picture(img_path_trio, Inches(0.95), Inches(4.35), width=Inches(5.4))

# Right: String Extraction
add_card(s10, Inches(6.8), Inches(1.5), Inches(5.733), Inches(5.5), "字串截取手術刀：LEFT / RIGHT / MID / FIND", C_TEAL)
tb10_r = s10.shapes.add_textbox(Inches(6.95), Inches(2.0), Inches(5.4), Inches(4.8))
tf10_r = tb10_r.text_frame
tf10_r.word_wrap = True

extracts = [
    ("• LEFT(D2, 4)：", "從 EMCU1234567 截取前 4 碼船公司碼 ➜ \"EMCU\""),
    ("• RIGHT(D2, 1)：", "截取最後 1 碼檢查碼 ➜ \"7\""),
    ("• MID(D2, 5, 6)：", "從第 5 位截取 6 碼貨櫃流水號 ➜ \"123456\""),
    ("• FIND(\"/\", C2)：", "動態定位符號位置，搭配 LEFT/MID 彈性拆分任何長度船名。"),
    ("• SUBSTITUTE(D2, \"V.\", \"\")：", "替換消除航次中的雜亂前綴字元。"),
    ("• LEN(B2)：", "計算字元長度，即時檢驗提單或櫃號是否缺碼或多塞空格！")
]
for t, d in extracts:
    p = tf10_r.add_paragraph()
    r_t = p.add_run()
    r_t.text = t
    r_t.font.name = FONT_CODE
    r_t.font.size = Pt(10)
    r_t.font.bold = True
    r_t.font.color.rgb = C_TEAL
    r_d = p.add_run()
    r_d.text = d
    r_d.font.name = FONT_MAIN
    r_d.font.size = Pt(9.5)
    r_d.font.color.rgb = C_LIGHT_GRY

# -------------------------------------------------------------
# SLIDE 11: 第五章 函數清洗管線篇 (二)
# -------------------------------------------------------------
s11 = add_base_slide(prs)
add_header(s11, "第五章：函數管線", "實戰示範——核心清洗函數與公式管線 (二)", 11)

# Left: VLOOKUP
add_card(s11, Inches(0.8), Inches(1.5), Inches(5.7), Inches(5.5), "跨表代碼標準化大師：VLOOKUP / XLOOKUP", C_AMBER)
tb11_l = s11.shapes.add_textbox(Inches(0.95), Inches(2.0), Inches(5.4), Inches(1.5))
tf11_l = tb11_l.text_frame
tf11_l.word_wrap = True

p_vl = tf11_l.paragraphs[0]
r_vl = p_vl.add_run()
r_vl.text = "參照『港口標準代碼表』，將「高雄、KHH、Kaohsiung」統一轉換為標準聯合國五字碼 TWKHH：\n"
r_vl.font.name = FONT_MAIN
r_vl.font.size = Pt(9.5)
r_vl.font.color.rgb = C_LIGHT_GRY

r_vlc = p_vl.add_run()
r_vlc.text = "=VLOOKUP(F2, 代碼表!$A$2:$C$12, 2, FALSE)"
r_vlc.font.name = FONT_CODE
r_vlc.font.size = Pt(10.5)
r_vlc.font.bold = True
r_vlc.font.color.rgb = C_AMBER

p_vln = tf11_l.add_paragraph()
r_vln = p_vln.add_run()
r_vln.text = "★ 關鍵防錯：對照範圍按 F4 鎖定為絕對參照 $A$2:$C$12，最後參數填 FALSE 代表完全精準匹配！"
r_vln.font.name = FONT_MAIN
r_vln.font.size = Pt(8.8)
r_vln.font.color.rgb = C_LIGHT_GRY

img_path_vl = os.path.join(IMG_DIR, "fig_5_2_vlookup_port_mapping.png")
if os.path.exists(img_path_vl):
    s11.shapes.add_picture(img_path_vl, Inches(0.95), Inches(3.6), width=Inches(5.4))

# Right: Numeric & Date & KPI table
add_card(s11, Inches(6.8), Inches(1.5), Inches(5.733), Inches(5.5), "數值日期重構 ＆ 30 筆清洗成果驗收", C_EMERALD)
tb11_r = s11.shapes.add_textbox(Inches(6.95), Inches(2.0), Inches(5.4), Inches(1.8))
tf11_r = tb11_r.text_frame
tf11_r.word_wrap = True

fn_items = [
    ("• =DATE(Y, M, D)：", "重構日期序號；=ETA - ETD 成功計算海上航程天數！"),
    ("• =VALUE(text)：", "強制將純文字數字轉化為真數值。"),
    ("• =IF(ISBLANK(O2), \"待補單\", O2)：", "自動補齊遺漏通關註記。")
]
for t, d in fn_items:
    p = tf11_r.add_paragraph()
    r_t = p.add_run()
    r_t.text = t
    r_t.font.name = FONT_CODE
    r_t.font.size = Pt(9.5)
    r_t.font.bold = True
    r_t.font.color.rgb = C_EMERALD
    r_d = p.add_run()
    r_d.text = d
    r_d.font.name = FONT_MAIN
    r_d.font.size = Pt(9.0)
    r_d.font.color.rgb = C_LIGHT_GRY

# Mini KPI Table
tbl_kpi_shape = s11.shapes.add_table(5, 3, Inches(6.95), Inches(3.7), Inches(5.4), Inches(3.0))
t_kpi = tbl_kpi_shape.table
t_kpi.columns[0].width = Inches(1.6)
t_kpi.columns[1].width = Inches(1.9)
t_kpi.columns[2].width = Inches(1.9)

kpi_headers = ["驗收指標", "清洗前 (01_原始資料)", "清洗後 (03_標準完成)"]
for c_i, h in enumerate(kpi_headers):
    cell = t_kpi.cell(0, c_i)
    cell.fill.solid()
    cell.fill.fore_color.rgb = RGBColor(16, 50, 40)
    p = cell.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = h
    r.font.name = FONT_MAIN
    r.font.size = Pt(9.0)
    r.font.bold = True
    r.font.color.rgb = C_WHITE

kpi_rows_data = [
    ("總資料筆數", "30 筆記錄", "28 筆 (剔除2筆重複)"),
    ("總毛重加總", "無法計算 (SUM=0)", "445,400.0 KG"),
    ("總海運費加總", "無法計算 (帶符號)", "$ 79,400.00 USD"),
    ("港口代碼標準化", "33.3% (俗稱混雜)", "100.0% (UN/LOCODE)")
]
for r_i, rvals in enumerate(kpi_rows_data, start=1):
    for c_i, val in enumerate(rvals):
        cell = t_kpi.cell(r_i, c_i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(16, 32, 45) if r_i % 2 == 0 else RGBColor(12, 24, 35)
        p = cell.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = val
        r.font.name = FONT_MAIN
        r.font.size = Pt(8.5)
        r.font.color.rgb = C_EMERALD if c_i == 2 else (C_ROSE if c_i == 1 and "無法" in val else C_LIGHT_GRY)

# -------------------------------------------------------------
# SLIDE 12: 第六章 300筆多情境航運案例練習
# -------------------------------------------------------------
s12 = add_base_slide(prs)
add_header(s12, "第六章：實戰練習", "300 筆多情境航運案例挑戰題庫與驗收規格", 12)

# Left: 4 Trades
add_card(s12, Inches(0.8), Inches(1.5), Inches(5.7), Inches(2.6), "四大海運業務板塊分佈 (300 筆總集)", C_CYAN)
tb12_l = s12.shapes.add_textbox(Inches(0.95), Inches(1.95), Inches(5.4), Inches(2.0))
tf12_l = tb12_l.text_frame
tf12_l.word_wrap = True

trades_list = [
    ("🚢 整櫃海運 (FCL)：", "150 筆 (50.0%) - 20GP, 40GP, 40HQ 標準定期航線。"),
    ("📦 併櫃海運 (LCL)：", "60 筆 (20.0%) - 多貨主散貨拼裝，CBM 與件數分離。"),
    ("❄️ 冷藏保鮮 (Reefer)：", "45 筆 (15.0%) - 40HR 特殊櫃型，-18℃ 生鮮肉品溫控指示。"),
    ("☣️ 危險品 (DG)：", "45 筆 (15.0%) - IMO Class 3/8/9，UN編號精準申報。")
]
for t, d in trades_list:
    p = tf12_l.add_paragraph()
    r_t = p.add_run()
    r_t.text = t
    r_t.font.name = FONT_MAIN
    r_t.font.size = Pt(9.5)
    r_t.font.bold = True
    r_t.font.color.rgb = C_CYAN
    r_d = p.add_run()
    r_d.text = d
    r_d.font.name = FONT_MAIN
    r_d.font.size = Pt(9.0)
    r_d.font.color.rgb = C_LIGHT_GRY

# Right: 5 KPIs
add_card(s12, Inches(6.8), Inches(1.5), Inches(5.733), Inches(2.6), "任務驗收五大指標規格書 (KPI Check)", C_TEAL)
tb12_r = s12.shapes.add_textbox(Inches(6.95), Inches(1.95), Inches(5.4), Inches(2.0))
tf12_r = tb12_r.text_frame
tf12_r.word_wrap = True

kpi_rules = [
    ("1. 重複提單排查率 100%：", "300 筆精準縮減為 285 筆（剔除 15 筆重複）。"),
    ("2. 提單與櫃號代碼 100% 正規化：", "全大寫、無首尾空格、長度正確符合 ISO。"),
    ("3. 裝卸港 UN/LOCODE 標準化：", "VLOOKUP 成功轉換，代碼符合率達 100%。"),
    ("4. 數值與單位 100% 剝離：", "毛重與運費純數值化，成功計算全體加總。"),
    ("5. 日期格式統一且邏輯正確：", "YYYY-MM-DD 序號，航程天數（ETA-ETD）無負值。")
]
for t, d in kpi_rules:
    p = tf12_r.add_paragraph()
    r_t = p.add_run()
    r_t.text = t
    r_t.font.name = FONT_MAIN
    r_t.font.size = Pt(9.5)
    r_t.font.bold = True
    r_t.font.color.rgb = C_TEAL
    r_d = p.add_run()
    r_d.text = d
    r_d.font.name = FONT_MAIN
    r_d.font.size = Pt(9.0)
    r_d.font.color.rgb = C_LIGHT_GRY

# Bottom Image: 300 rows defect overview
img_path_300 = os.path.join(IMG_DIR, "fig_6_1_practice_300_overview.png")
if os.path.exists(img_path_300):
    add_card(s12, Inches(0.8), Inches(4.25), Inches(11.733), Inches(2.8), "300 筆情境練習檔之業務結構與資料缺陷分佈藍圖", C_CYAN)
    s12.shapes.add_picture(img_path_300, Inches(0.95), Inches(4.7), width=Inches(11.433))

# -------------------------------------------------------------
# SLIDE 13: 第七章 300筆參考解題步驟與驗收對帳表
# -------------------------------------------------------------
s13 = add_base_slide(prs)
add_header(s13, "第七章：解答對帳", "300 筆情境練習題關鍵指標查核對帳表 (KPIs Audit)", 13)

# Full Table for 300 rows audit
add_card(s13, Inches(0.8), Inches(1.5), Inches(11.733), Inches(5.5), "清洗成果完整對帳表 (加總數據完全吻合即通過檢定)", C_EMERALD)

tbl_audit_shape = s13.shapes.add_table(9, 4, Inches(1.0), Inches(2.0), Inches(11.333), Inches(4.5))
t_audit = tbl_audit_shape.table
t_audit.columns[0].width = Inches(2.4)
t_audit.columns[1].width = Inches(2.7)
t_audit.columns[2].width = Inches(2.9)
t_audit.columns[3].width = Inches(3.333)

audit_headers = ["檢核對帳項目", "未清洗原始數據 (Raw Data)", "清洗完成驗收數據 (Cleaned)", "品質評定與查核說明"]
for c_i, h in enumerate(audit_headers):
    cell = t_audit.cell(0, c_i)
    cell.fill.solid()
    cell.fill.fore_color.rgb = RGBColor(16, 50, 40)
    p = cell.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = h
    r.font.name = FONT_MAIN
    r.font.size = Pt(10)
    r.font.bold = True
    r.font.color.rgb = C_WHITE

audit_data = [
    ("總資料筆數 (Rows)", "300 筆記錄", "285 筆獨立記錄", "成功移除 15 筆重複拋轉資料 (合格 ✓)"),
    ("總毛重加總 (Total GW)", "無法計算 (SUM=0)", "4,749,200.0 KG", "單位清除完成，空值依規格平均補正 (合格 ✓)"),
    ("總海運費營收 (Total Frt)", "無法計算 (夾帶$與USD)", "$ 735,450.00 USD", "乘 1 消除綠三角，純數值加總精準 (合格 ✓)"),
    ("總計費材積 (Total CBM)", "無法計算 (帶有CBM字串)", "10,958.4 CBM", "純數值化，小數點維持 1 位 (合格 ✓)"),
    ("裝貨港代碼 (POL) 標準化", "38.2% (中文俗稱混雜)", "100.0% (UN/LOCODE)", "VLOOKUP 成功跨表自動轉換 (合格 ✓)"),
    ("卸貨港代碼 (POD) 標準化", "41.5% (別名混雜)", "100.0% (UN/LOCODE)", "VLOOKUP 成功精準匹配 (合格 ✓)"),
    ("日期格式標準化率", "26.0% (四種格式混雜)", "100.0% (YYYY-MM-DD)", "資料剖析修復為真日期序號 (合格 ✓)"),
    ("報關狀態資料完整率", "91.3% (含空值)", "100.0% (補齊待補單)", "遺漏值填補符合通關業務規範 (合格 ✓)")
]

for r_i, rvals in enumerate(audit_data, start=1):
    bg_c = RGBColor(16, 32, 45) if r_i % 2 == 0 else RGBColor(12, 24, 35)
    for c_i, val in enumerate(rvals):
        cell = t_audit.cell(r_i, c_i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg_c
        p = cell.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = val
        r.font.name = FONT_MAIN
        r.font.size = Pt(9.0)
        r.font.bold = (c_i in [0, 2])
        r.font.color.rgb = C_EMERALD if c_i == 2 else (C_ROSE if c_i == 1 and "無法" in val else C_LIGHT_GRY)

# -------------------------------------------------------------
# SLIDE 14: 結語與實戰資源清單
# -------------------------------------------------------------
s14 = add_base_slide(prs)
add_header(s14, "結語與資源", "航運資料前處理學習總結與實踐教材資源清單", 14)

# Left: 5 Core Mindsets
add_card(s14, Inches(0.8), Inches(1.5), Inches(5.7), Inches(5.5), "航運行政人員前處理五大核心心法", C_CYAN)
tb14_l = s14.shapes.add_textbox(Inches(0.95), Inches(2.0), Inches(5.4), Inches(4.8))
tf14_l = tb14_l.text_frame
tf14_l.word_wrap = True

mindsets = [
    ("1. 備份第一，自保無憂：", "拿到檔案先右鍵【建立複本】，永不在原始資料上直接施工。"),
    ("2. 堅守二維，消滅合併：", "嚴格禁止跨欄置中合併儲存格，一列一筆、一行一屬性。"),
    ("3. 熟記快捷，效率翻倍：", "Ctrl+H 取代、Ctrl+Shift+L 篩選、F4 鎖定、雙擊右下角黑十字閃電填滿。"),
    ("4. 建立管線，自動重複：", "文字用 =UPPER(TRIM(CLEAN()))，港名用 VLOOKUP 跨表自動化關聯。"),
    ("5. 對帳核驗，守護合規：", "清洗完成必核對筆數、毛重總和與運費加總，對帳達標才封存！")
]
for t, d in mindsets:
    p = tf14_l.add_paragraph()
    r_t = p.add_run()
    r_t.text = t
    r_t.font.name = FONT_MAIN
    r_t.font.size = Pt(10)
    r_t.font.bold = True
    r_t.font.color.rgb = C_CYAN
    r_d = p.add_run()
    r_d.text = d + "\n"
    r_d.font.name = FONT_MAIN
    r_d.font.size = Pt(9.5)
    r_d.font.color.rgb = C_LIGHT_GRY

# Right: Package Files
add_card(s14, Inches(6.8), Inches(1.5), Inches(5.733), Inches(5.5), "完整配套學習資源清單 (全套實體電子檔)", C_TEAL)
tb14_r = s14.shapes.add_textbox(Inches(6.95), Inches(2.0), Inches(5.4), Inches(4.8))
tf14_r = tb14_r.text_frame
tf14_r.word_wrap = True

res_files = [
    ("📄 完整 Word 教學手冊 (.docx)：", "《航運行政人員_Excel基礎資料前處理操作教學手冊.docx》\n約 2.9 MB，具備封面、目錄、圖表目錄、頁碼與 14 張操作截圖。"),
    ("📊 30 筆示範活頁簿 (.xlsx)：", "《航運資料前處理教學_30筆示範檔.xlsx》\n含原始髒資料、港口對照表與清洗完成標準資料。"),
    ("🎯 300 筆多情境練習題庫 (.xlsx)：", "《航運資料前處理實務_300筆情境練習檔.xlsx》\n含 FCL/LCL/冷藏/危險品多情境練習題、解答庫與 KPI 對帳表。"),
    ("🌐 互動式 PPT 網頁簡報 (.html)：", "《shipping_excel_presentation.html》\n全功能 14 頁互動簡報，支援即時公式模擬器與互動小測驗！")
]
for t, d in res_files:
    p = tf14_r.add_paragraph()
    r_t = p.add_run()
    r_t.text = t
    r_t.font.name = FONT_MAIN
    r_t.font.size = Pt(10)
    r_t.font.bold = True
    r_t.font.color.rgb = C_TEAL
    r_d = p.add_run()
    r_d.text = d + "\n"
    r_d.font.name = FONT_MAIN
    r_d.font.size = Pt(9.0)
    r_d.font.color.rgb = C_LIGHT_GRY

out_path = r"c:\shipping_tutorial\航運行政人員_Excel基礎資料前處理實戰簡報.pptx"
prs.save(out_path)
print(f"PowerPoint Presentation generated successfully at {out_path} ({os.path.getsize(out_path)/1024:.1f} KB).")
