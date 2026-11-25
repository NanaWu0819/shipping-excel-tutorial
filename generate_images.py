import matplotlib.pyplot as plt
import matplotlib.patches as patches
import os

# Configure matplotlib for Traditional Chinese
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

out_dir = "c:/shipping_tutorial/images"
os.makedirs(out_dir, exist_ok=True)

# -------------------------------------------------------------
# 1. 圖 1-1：MS Excel 操作介面全景導覽圖解
# -------------------------------------------------------------
def make_fig_1_1():
    fig, ax = plt.subplots(figsize=(12, 7), dpi=200)
    ax.set_facecolor("#F8F9FA")
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 70)
    ax.axis("off")

    ax.text(50, 67, "Microsoft Excel 2021 / 365 核心操作介面區域導覽", ha="center", va="center", fontsize=15, fontweight="bold", color="#1B365D")

    window = patches.Rectangle((2, 2), 96, 62, linewidth=1.5, edgecolor="#1B365D", facecolor="#FFFFFF", zorder=1)
    ax.add_patch(window)

    title_bar = patches.Rectangle((2, 59), 96, 5, facecolor="#107C41", edgecolor="none", zorder=2)
    ax.add_patch(title_bar)
    ax.text(5, 61.5, "Autosave [On]  |  航運出口提單總表.xlsx - Excel", color="#FFFFFF", fontsize=10, fontweight="bold", va="center", zorder=3)
    ax.text(88, 61.5, "--   [ ]   X", color="#FFFFFF", fontsize=11, fontweight="bold", va="center", zorder=3)

    ribbon_tabs = patches.Rectangle((2, 55.5), 96, 3.5, facecolor="#E1DFDD", edgecolor="#C8C6C4", zorder=2)
    ax.add_patch(ribbon_tabs)
    tabs = ["檔案", "常用 (Home)", "插入", "繪圖", "版面配置", "公式 (Formulas)", "資料 (Data)", "校閱", "檢視 (View)", "說明"]
    x_pos = 4
    for i, t in enumerate(tabs):
        bg = "#FFFFFF" if i == 0 or t.startswith("常用") else "#E1DFDD"
        tc = "#107C41" if t.startswith("常用") else "#323130"
        tab_box = patches.Rectangle((x_pos, 55.5), len(t)*1.8 + 2.5, 3.5, facecolor=bg, edgecolor="none", zorder=3)
        ax.add_patch(tab_box)
        ax.text(x_pos + (len(t)*1.8 + 2.5)/2, 57.2, t, color=tc, fontsize=8.5, fontweight="bold" if t.startswith("常用") else "normal", ha="center", va="center", zorder=4)
        x_pos += len(t)*1.8 + 3.5

    ribbon_body = patches.Rectangle((2, 43), 96, 12.5, facecolor="#F3F2F1", edgecolor="#EDEBE9", zorder=2)
    ax.add_patch(ribbon_body)
    groups = [
        ("剪貼簿", 4, 18, ["貼上", "複製", "剪下", "格式轉換"]),
        ("字型 (Font)", 23, 40, ["新細明體/微軟正黑體", "字級 11", "粗體 B / 填色 / 框線"]),
        ("對齊方式 (Align)", 41, 58, ["靠左 / 置中 / 靠右", "自動換欄", "跨欄置中"]),
        ("數值 (Number)", 59, 77, ["通用格式 (選單)", "$ 貨幣", "% 百分比", "小數增減"]),
        ("樣式與編輯", 78, 96, ["設定格式化的條件", "表格樣式", "自動加總", "排序與篩選"])
    ]
    for gname, x1, x2, items in groups:
        box = patches.Rectangle((x1, 44), x2 - x1, 11, facecolor="#FFFFFF", edgecolor="#D2D0CE", linewidth=0.8, zorder=3)
        ax.add_patch(box)
        ax.text((x1 + x2)/2, 45.2, gname, color="#605E5C", fontsize=7.5, fontweight="bold", ha="center", zorder=4)
        for idx_it, it in enumerate(items):
            ax.text((x1 + x2)/2, 52.5 - idx_it*2.3, it, color="#323130", fontsize=7.2, ha="center", zorder=4)

    ax.add_patch(patches.Rectangle((2, 38.5), 15, 4.2, facecolor="#FFFFFF", edgecolor="#C8C6C4", linewidth=1, zorder=3))
    ax.text(9.5, 40.6, "C5", color="#107C41", fontsize=11, fontweight="bold", ha="center", va="center", zorder=4)
    
    ax.add_patch(patches.Rectangle((18, 38.5), 5, 4.2, facecolor="#F3F2F1", edgecolor="#C8C6C4", linewidth=1, zorder=3))
    ax.text(20.5, 40.6, "fx", color="#605E5C", fontsize=10, fontstyle="italic", fontweight="bold", ha="center", va="center", zorder=4)
    
    ax.add_patch(patches.Rectangle((23.5, 38.5), 74.5, 4.2, facecolor="#FFFFFF", edgecolor="#C8C6C4", linewidth=1, zorder=3))
    ax.text(25, 40.6, "=TRIM(CLEAN(B5))", color="#004578", fontsize=10, fontfamily="Consolas", va="center", zorder=4)

    col_w = 11.5
    ax.add_patch(patches.Rectangle((2, 33.5), 5, 4.5, facecolor="#E1DFDD", edgecolor="#C8C6C4", zorder=3))
    cols = ["A (項次)", "B (提單號)", "C (船名航次)", "D (起運港)", "E (目的港)", "F (開航日)", "G (毛重_KG)", "H (海運費)"]
    for i, c in enumerate(cols):
        ax.add_patch(patches.Rectangle((7 + i*col_w, 33.5), col_w, 4.5, facecolor="#F3F2F1", edgecolor="#C8C6C4", zorder=3))
        ax.text(7 + i*col_w + col_w/2, 35.7, c, color="#323130", fontsize=7.5, fontweight="bold", ha="center", va="center", zorder=4)

    rows_data = [
        ("1", ["1", "OOCLU9823101", "EVER GIVEN V.0123W", "TWKHH", "USLAX", "2026/11/05", "18,500 KGS", "$2,850"]),
        ("2", ["2", "EMCE8901234", "YM UNICORN / 045E", "高雄", "LAX", "2026.11.08", "9,800 kgs", "$1,600"]),
        ("3", ["3", "MSK7654321", "OOCL POLAND - 102N", "KHH", "Los Angeles", "12-Nov-2026", "14,200", "2400"]),
        ("4", ["4", "whl9876543", "WAN HAI 315 V.S088", "TWKEL", "Tokyo", "2026/11/15", "21500 KGS", "$3,100"]),
        ("5", ["5", "ONE8877665", "ONE HARBOR V.022W", "基隆", "DEHAM", "2026.11.18", "16,800 KGS", "$3,850"])
    ]
    for r_idx, (r_no, vals) in enumerate(rows_data):
        y = 29.5 - r_idx*4
        ax.add_patch(patches.Rectangle((2, y), 5, 4, facecolor="#F3F2F1", edgecolor="#C8C6C4", zorder=3))
        ax.text(4.5, y + 2, r_no, color="#323130", fontsize=8, fontweight="bold", ha="center", va="center", zorder=4)
        for c_idx, val in enumerate(vals):
            is_active = (r_no == "5" and c_idx == 2)
            bg = "#D0E7F9" if is_active else ("#FFFFFF" if r_idx % 2 == 0 else "#FAF9F8")
            edge = "#0078D4" if is_active else "#E1DFDD"
            lw = 2 if is_active else 0.8
            ax.add_patch(patches.Rectangle((7 + c_idx*col_w, y), col_w, 4, facecolor=bg, edgecolor=edge, linewidth=lw, zorder=3))
            ax.text(7 + c_idx*col_w + 1, y + 2, val, color="#0078D4" if is_active else "#201F1E", fontsize=7.2, va="center", zorder=4)

    ax.add_patch(patches.Rectangle((2, 5.5), 96, 3.5, facecolor="#F3F2F1", edgecolor="#C8C6C4", zorder=2))
    tabs_ws = [("01_原始資料(未清理)", True), ("02_港口標準代碼表", False), ("03_清洗完成標準資料", False)]
    tx = 5
    for tname, active in tabs_ws:
        w = len(tname)*1.7 + 4
        bg = "#FFFFFF" if active else "#EDEBE9"
        tc = "#107C41" if active else "#605E5C"
        ax.add_patch(patches.Rectangle((tx, 5.5), w, 3.5, facecolor=bg, edgecolor="#C8C6C4", zorder=3))
        if active:
            ax.add_patch(patches.Rectangle((tx, 5.5), w, 0.6, facecolor="#107C41", edgecolor="none", zorder=4))
        ax.text(tx + w/2, 7.2, tname, color=tc, fontsize=7.5, fontweight="bold" if active else "normal", ha="center", va="center", zorder=4)
        tx += w + 2
    ax.text(tx + 2, 7.2, "+ [新增工作表]", color="#605E5C", fontsize=8, va="center", zorder=3)

    ax.add_patch(patches.Rectangle((2, 2), 96, 3.5, facecolor="#107C41", edgecolor="none", zorder=2))
    ax.text(5, 3.7, "就緒 (Ready)   |   篩選中: 30 筆中已選取 1 筆儲存格", color="#FFFFFF", fontsize=8, va="center", zorder=3)
    ax.text(78, 3.7, "平均值: 16,800  計數: 1  加總: 16,800   |  100% [- | +]", color="#FFFFFF", fontsize=7.8, va="center", zorder=3)

    callouts = [
        ((10, 60), (10, 65), "1. 標題列與快速存取工具列", "#1B365D"),
        ((35, 57), (35, 63), "2. 功能區索引標籤 (Ribbon Tabs)", "#008080"),
        ((9, 41), (9, 47), "3. 名稱方塊 (顯示目前格位代碼)", "#C2410C"),
        ((50, 41), (50, 47), "4. 資料編輯列 (輸入與檢視公式)", "#004578"),
        ((3, 35), (3, 26), "5. 欄號與列號座標系統", "#4B5563"),
        ((40, 15), (40, 11), "6. 二維儲存格區域 (選取格為 C5)", "#2563EB"),
        ((20, 6), (20, 1), "7. 工作表標籤 (Sheet Tabs 導覽)", "#15803D"),
        ((85, 3), (85, -2), "8. 狀態列 (即時統計、檢視比例縮放)", "#047857")
    ]
    for (px, py), (tx, ty), text, col in callouts:
        ax.annotate(text, xy=(px, py), xytext=(tx, ty),
                    arrowprops=dict(facecolor=col, edgecolor=col, arrowstyle="->", lw=1.5),
                    fontsize=8.5, fontweight="bold", color=col, bbox=dict(boxstyle="round,pad=0.3", fc="#FEF3C7" if "格位" in text else "#FFFFFF", ec=col, lw=1.2), zorder=6)

    plt.tight_layout()
    plt.savefig(f"{out_dir}/fig_1_1_excel_ui.png", dpi=300)
    plt.close()

# -------------------------------------------------------------
# 2. 圖 2-1：航運資料前處理核心五步 SOP 流程圖
# -------------------------------------------------------------
def make_fig_2_1():
    fig, ax = plt.subplots(figsize=(12, 4.5), dpi=200)
    ax.set_facecolor("#F8F9FA")
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 40)
    ax.axis("off")

    ax.text(50, 36.5, "航運業務資料前處理黃金標準作業流程 (Standard Preprocessing Pipeline)", ha="center", fontsize=14, fontweight="bold", color="#1B365D")

    steps = [
        ("第 1 步：防護備份", "原始工作表複製\n保留 Raw Data\n嚴禁覆蓋原檔", "#1E3A8A"),
        ("第 2 步：結構與唯一性", "凍結窗格與自動篩選\n依 B/L 提單唯一值\n【移除重複項】", "#0D9488"),
        ("第 3 步：字串與代碼清洗", "TRIM 清除首尾空格\nUPPER 大寫正規化\n代碼對照表標準化", "#D97706"),
        ("第 4 步：數值與日期修復", "【尋找取代】剔除單位\n文字數字乘1轉數值\n【資料剖析】修復日期", "#2563EB"),
        ("第 5 步：遺漏填補與封存", "IF / ISBLANK 缺漏填補\n設定格式化條件標記\n另存分析標準庫", "#16A34A")
    ]

    box_w = 16.5
    box_h = 24
    start_x = 4.5
    gap = 2.8

    for i, (title, content, color) in enumerate(steps):
        bx = start_x + i * (box_w + gap)
        by = 7
        ax.add_patch(patches.FancyBboxPatch((bx + 0.4, by - 0.4), box_w, box_h, boxstyle="round,pad=0.5", facecolor="#E5E7EB", edgecolor="none", zorder=1))
        ax.add_patch(patches.FancyBboxPatch((bx, by), box_w, box_h, boxstyle="round,pad=0.5", facecolor="#FFFFFF", edgecolor=color, linewidth=2, zorder=2))
        ax.add_patch(patches.FancyBboxPatch((bx, by + box_h - 6), box_w, 6, boxstyle="round,pad=0.2", facecolor=color, edgecolor="none", zorder=3))
        ax.text(bx + box_w/2, by + box_h - 3, title, color="#FFFFFF", fontsize=9.5, fontweight="bold", ha="center", va="center", zorder=4)
        lines = content.split("\n")
        for l_idx, line in enumerate(lines):
            ax.text(bx + box_w/2, by + box_h - 9.5 - l_idx*4.2, line, color="#374151", fontsize=8.2, ha="center", va="center", zorder=4)

        if i < 4:
            ax.annotate("", xy=(bx + box_w + gap - 0.3, by + box_h/2), xytext=(bx + box_w + 0.3, by + box_h/2),
                        arrowprops=dict(facecolor="#9CA3AF", edgecolor="#9CA3AF", width=3, headwidth=8, headlength=6), zorder=5)

    plt.tight_layout()
    plt.savefig(f"{out_dir}/fig_2_1_preprocessing_sop.png", dpi=300)
    plt.close()

# -------------------------------------------------------------
# 3. 圖 2-2：標準二維表格 (Tidy Data) vs. 錯誤結構對照圖
# -------------------------------------------------------------
def make_fig_2_2():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 5), dpi=200)
    for ax in [ax1, ax2]:
        ax.set_facecolor("#F9FAFB")
        ax.set_xlim(0, 50)
        ax.set_ylim(0, 45)
        ax.axis("off")

    ax1.set_title("[X] 航運行政常見錯誤結構 (非標準化報表)", fontsize=11, fontweight="bold", color="#DC2626", pad=12)
    ax1.add_patch(patches.Rectangle((2, 35), 46, 5, facecolor="#FCA5A5", edgecolor="#B91C1C", linewidth=1.5))
    ax1.text(25, 37.5, "2026年11月份 出口航次裝船彙總表 (跨欄合併儲存格)", fontsize=9, fontweight="bold", color="#7F1D1D", ha="center", va="center")
    
    ax1.add_patch(patches.Rectangle((2, 30), 12, 5, facecolor="#FECACA", edgecolor="#B91C1C", linewidth=1))
    ax1.text(8, 32.5, "船名航次", fontsize=8, fontweight="bold", color="#7F1D1D", ha="center", va="center")
    ax1.add_patch(patches.Rectangle((14, 30), 18, 5, facecolor="#FECACA", edgecolor="#B91C1C", linewidth=1))
    ax1.text(23, 32.5, "貨櫃與重量 (混同欄位)", fontsize=8, fontweight="bold", color="#7F1D1D", ha="center", va="center")
    ax1.add_patch(patches.Rectangle((32, 30), 16, 5, facecolor="#FECACA", edgecolor="#B91C1C", linewidth=1))
    ax1.text(40, 32.5, "運費合計 (含幣別)", fontsize=8, fontweight="bold", color="#7F1D1D", ha="center", va="center")

    ax1.add_patch(patches.Rectangle((2, 18), 12, 12, facecolor="#FFFFFF", edgecolor="#D1D5DB", linewidth=1))
    ax1.text(8, 24, "EVER GIVEN\n(垂直合併兩列)", fontsize=7.5, color="#DC2626", ha="center", va="center")
    
    ax1.add_patch(patches.Rectangle((14, 24), 18, 6, facecolor="#FFFFFF", edgecolor="#D1D5DB", linewidth=1))
    ax1.text(23, 27, "40HQ: 18,500 KGS", fontsize=7.5, color="#374151", ha="center", va="center")
    ax1.add_patch(patches.Rectangle((32, 24), 16, 6, facecolor="#FFFFFF", edgecolor="#D1D5DB", linewidth=1))
    ax1.text(40, 27, "$2,850 USD", fontsize=7.5, color="#374151", ha="center", va="center")

    ax1.add_patch(patches.Rectangle((14, 18), 18, 6, facecolor="#FFFFFF", edgecolor="#D1D5DB", linewidth=1))
    ax1.text(23, 21, "20GP: 9,800 KGS", fontsize=7.5, color="#374151", ha="center", va="center")
    ax1.add_patch(patches.Rectangle((32, 18), 16, 6, facecolor="#FFFFFF", edgecolor="#D1D5DB", linewidth=1))
    ax1.text(40, 21, "$1,600 USD", fontsize=7.5, color="#374151", ha="center", va="center")

    ax1.add_patch(patches.Rectangle((2, 12), 30, 6, facecolor="#FEF3C7", edgecolor="#F59E0B", linewidth=1))
    ax1.text(17, 15, "小計 (Subtotal: 混入資料列中)", fontsize=8, fontweight="bold", color="#B45309", ha="center", va="center")
    ax1.add_patch(patches.Rectangle((32, 12), 16, 6, facecolor="#FEF3C7", edgecolor="#F59E0B", linewidth=1))
    ax1.text(40, 15, "$4,450", fontsize=8, fontweight="bold", color="#B45309", ha="center", va="center")

    ax1.text(25, 4, "致命缺點：合併儲存格導致無法排序、無法篩選、\n公式參照錯位，樞紐分析表直接失效！", fontsize=8.2, color="#B91C1C", ha="center", bbox=dict(boxstyle="round,pad=0.4", fc="#FEE2E2", ec="#EF4444"))

    # Right: 標準 Tidy Data 結構
    ax2.set_title("[V] 國際標準二維資料表 (Tidy Data 規範結構)", fontsize=11, fontweight="bold", color="#15803D", pad=12)
    headers = [("提單號碼", 2, 11), ("船名", 11, 20), ("櫃型", 20, 28), ("毛重_KG", 28, 39), ("運費_USD", 39, 48)]
    for h, x1, x2 in headers:
        ax2.add_patch(patches.Rectangle((x1, 32), x2 - x1, 6, facecolor="#16A34A", edgecolor="#14532D", linewidth=1))
        ax2.text((x1 + x2)/2, 35, h, fontsize=8, fontweight="bold", color="#FFFFFF", ha="center", va="center")

    tidy_rows = [
        ["OOCLU9823101", "EVER GIVEN", "40HQ", "18500.0", "2850.00"],
        ["EMCE8901234", "EVER GIVEN", "20GP", "9800.0", "1600.00"],
        ["MSK7654321", "YM UNICORN", "40GP", "14200.0", "2400.00"],
        ["WHL9876543", "WAN HAI 315", "40HQ", "21500.0", "3100.00"]
    ]
    for r_idx, rvals in enumerate(tidy_rows):
        y = 26 - r_idx*5.5
        bg = "#F0FDF4" if r_idx % 2 == 0 else "#FFFFFF"
        for c_idx, (h, x1, x2) in enumerate(headers):
            ax2.add_patch(patches.Rectangle((x1, y), x2 - x1, 5.5, facecolor=bg, edgecolor="#D1D5DB", linewidth=0.8))
            align = "right" if c_idx >= 3 else "left"
            tx = (x2 - 1.5) if align == "right" else (x1 + 1.5)
            ax2.text(tx, y + 2.7, rvals[c_idx], fontsize=7.5, color="#1F2937", ha=align, va="center")

    ax2.text(25, 4, "核心原則：單一標題列、一列一筆記錄、一行一屬性、\n數值無文字單位，完美對接公式、樞紐分析與 BI！", fontsize=8.2, color="#15803D", ha="center", bbox=dict(boxstyle="round,pad=0.4", fc="#DCFCE7", ec="#22C55E"))

    plt.tight_layout()
    plt.savefig(f"{out_dir}/fig_2_2_tidy_data.png", dpi=300)
    plt.close()

# -------------------------------------------------------------
# 4. 圖 4-1：工作表複製與備份操作對話框
# -------------------------------------------------------------
def make_fig_4_1():
    fig, ax = plt.subplots(figsize=(8, 5.5), dpi=200)
    ax.set_facecolor("#F3F4F6")
    ax.set_xlim(0, 80)
    ax.set_ylim(0, 55)
    ax.axis("off")

    ax.text(40, 52, "Excel 核心防護技能：工作表完整備份操作", fontsize=12, fontweight="bold", color="#1B365D", ha="center")

    dlg = patches.FancyBboxPatch((12, 6), 56, 42, boxstyle="round,pad=0.4", facecolor="#FFFFFF", edgecolor="#0078D4", linewidth=2)
    ax.add_patch(dlg)

    ax.add_patch(patches.Rectangle((12, 42), 56, 6, facecolor="#0078D4", edgecolor="none"))
    ax.text(15, 45, "移動或複製 (Move or Copy)", color="#FFFFFF", fontsize=10, fontweight="bold", va="center")
    ax.text(65, 45, "X", color="#FFFFFF", fontsize=10, fontweight="bold", va="center")

    ax.text(16, 38, "將選取工作表移至活頁簿 (To book):", fontsize=8.5, color="#374151")
    ax.add_patch(patches.Rectangle((16, 32), 48, 4.5, facecolor="#F9FAFB", edgecolor="#D1D5DB"))
    ax.text(18, 34.2, "航運資料前處理教學_30筆示範檔.xlsx  (選單)", fontsize=8, color="#111827", va="center")

    ax.text(16, 28, "在下列工作表之前 (Before sheet):", fontsize=8.5, color="#374151")
    ax.add_patch(patches.Rectangle((16, 15), 48, 12, facecolor="#FFFFFF", edgecolor="#D1D5DB"))
    sheets = ["01_原始資料(未清理)", "02_港口標準代碼表", "03_清洗完成標準資料", "(移至最後)"]
    for i, s in enumerate(sheets):
        bg = "#0078D4" if i == 0 else "#FFFFFF"
        tc = "#FFFFFF" if i == 0 else "#374151"
        if i == 0:
            ax.add_patch(patches.Rectangle((16.2, 23.5), 47.6, 3.2, facecolor=bg, edgecolor="none"))
        ax.text(18, 25 - i*2.8, s, fontsize=7.8, color=tc, va="center")

    ax.add_patch(patches.Rectangle((16, 9.5), 3, 3, facecolor="#0078D4", edgecolor="#0078D4"))
    ax.text(17.5, 11, "V", color="#FFFFFF", fontsize=9, fontweight="bold", ha="center", va="center")
    ax.text(21, 11, "建立複本 (Create a copy)  [極度重要操作]", fontsize=9, fontweight="bold", color="#DC2626", va="center")

    ax.add_patch(patches.Rectangle((44, 7), 10, 3.5, facecolor="#0078D4", edgecolor="none"))
    ax.text(49, 8.7, "確定", color="#FFFFFF", fontsize=8.5, fontweight="bold", ha="center", va="center")
    ax.add_patch(patches.Rectangle((55, 7), 9, 3.5, facecolor="#E5E7EB", edgecolor="#D1D5DB"))
    ax.text(59.5, 8.7, "取消", color="#374151", fontsize=8.5, ha="center", va="center")

    ax.annotate("務必勾選！確保清洗失誤時，\n可立即退回原始數據重新比對", xy=(35, 9.5), xytext=(45, 1.5),
                arrowprops=dict(facecolor="#DC2626", edgecolor="#DC2626", arrowstyle="->", lw=1.5),
                fontsize=8.5, fontweight="bold", color="#DC2626", bbox=dict(boxstyle="round,pad=0.3", fc="#FEF2F2", ec="#DC2626"))

    plt.tight_layout()
    plt.savefig(f"{out_dir}/fig_4_1_sheet_backup.png", dpi=300)
    plt.close()

# -------------------------------------------------------------
# 5. 圖 4-2：凍結窗格與自動篩選
# -------------------------------------------------------------
def make_fig_4_2():
    fig, ax = plt.subplots(figsize=(10, 5), dpi=200)
    ax.set_facecolor("#F9FAFB")
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 50)
    ax.axis("off")

    ax.text(50, 46.5, "常用檢視與探索功能：凍結首列與啟用自動篩選", fontsize=12, fontweight="bold", color="#1B365D", ha="center")

    ax.add_patch(patches.Rectangle((5, 34), 42, 9, facecolor="#EFF6FF", edgecolor="#3B82F6", linewidth=1.2))
    ax.text(26, 40.5, "步驟 A：啟用首列凍結 (Freeze Top Row)", fontsize=8.5, fontweight="bold", color="#1E40AF", ha="center")
    ax.text(26, 36.5, "選單路徑：【檢視】->【凍結窗格】->【凍結首列】\n滾動資料時標題欄位永不消失！", fontsize=7.5, color="#1F2937", ha="center")

    ax.add_patch(patches.Rectangle((53, 34), 42, 9, facecolor="#F0FDF4", edgecolor="#22C55E", linewidth=1.2))
    ax.text(74, 40.5, "步驟 B：啟用自動篩選 (Filter)", fontsize=8.5, fontweight="bold", color="#166534", ha="center")
    ax.text(74, 36.5, "選單路徑：【資料】->【篩選】(快速鍵 Ctrl+Shift+L)\n各欄標題自動出現下拉按鈕 (v)", fontsize=7.5, color="#1F2937", ha="center")

    ax.add_patch(patches.Rectangle((5, 6), 90, 24, facecolor="#FFFFFF", edgecolor="#CBD5E1", linewidth=1.5))
    ax.axhline(y=24, xmin=0.05, xmax=0.95, color="#2563EB", linewidth=2.5, linestyle="-")
    ax.text(96, 24, "<- 凍結線", color="#2563EB", fontsize=8, fontweight="bold", va="center")

    headers = [("項次", 5, 15), ("提單號碼", 15, 33), ("船名航次", 33, 56), ("起運港 (POL)", 56, 75), ("卸貨港 (POD)", 75, 95)]
    for h, x1, x2 in headers:
        ax.add_patch(patches.Rectangle((x1, 24), x2 - x1, 6, facecolor="#1E3A8A", edgecolor="#CBD5E1"))
        ax.text(x1 + 2, 27, h, color="#FFFFFF", fontsize=8, fontweight="bold", va="center")
        ax.add_patch(patches.Rectangle((x2 - 5, 25), 4, 4, facecolor="#DBEAFE", edgecolor="#3B82F6"))
        ax.text(x2 - 3, 27, "v", color="#1E40AF", fontsize=7, fontweight="bold", ha="center", va="center")

    rows = [
        ["1", "OOCLU9823101", "EVER GIVEN V.0123W", "TWKHH", "USLAX"],
        ["2", "EMCE8901234", "YM UNICORN / 045E", "高雄 (混雜)", "LAX (簡寫)"],
        ["3", "MSK7654321", "OOCL POLAND - 102N", "KHH (縮寫)", "Los Angeles"]
    ]
    for r_idx, r in enumerate(rows):
        y = 18 - r_idx*5.5
        for c_idx, (h, x1, x2) in enumerate(headers):
            ax.add_patch(patches.Rectangle((x1, y), x2 - x1, 5.5, facecolor="#FFFFFF" if r_idx % 2 == 0 else "#F8FAFC", edgecolor="#E2E8F0"))
            ax.text(x1 + 2, y + 2.7, r[c_idx], color="#1E293B", fontsize=7.5, va="center")

    plt.tight_layout()
    plt.savefig(f"{out_dir}/fig_4_2_freeze_panes_filter.png", dpi=300)
    plt.close()

# -------------------------------------------------------------
# 6. 圖 4-3：尋找與取代 (Find & Replace)
# -------------------------------------------------------------
def make_fig_4_3():
    fig, ax = plt.subplots(figsize=(8.5, 5), dpi=200)
    ax.set_facecolor("#F9FAFB")
    ax.set_xlim(0, 85)
    ax.set_ylim(0, 50)
    ax.axis("off")

    ax.text(42.5, 47, "批次清理利器：【尋找與取代】剔除單位字尾與符號", fontsize=12, fontweight="bold", color="#1B365D", ha="center")

    ax.add_patch(patches.FancyBboxPatch((10, 6), 65, 38, boxstyle="round,pad=0.4", facecolor="#FFFFFF", edgecolor="#0078D4", linewidth=1.8))
    ax.add_patch(patches.Rectangle((10, 38), 65, 6, facecolor="#0078D4", edgecolor="none"))
    ax.text(14, 41, "尋找與取代 (Find and Replace)  - [快速鍵 Ctrl+H]", color="#FFFFFF", fontsize=9.5, fontweight="bold", va="center")
    ax.text(72, 41, "X", color="#FFFFFF", fontsize=10, va="center")

    ax.add_patch(patches.Rectangle((14, 34), 16, 4, facecolor="#FFFFFF", edgecolor="#0078D4", linewidth=1))
    ax.text(22, 36, "取代 (Replace)", color="#0078D4", fontsize=8.5, fontweight="bold", ha="center", va="center")
    ax.add_patch(patches.Rectangle((30, 34), 16, 4, facecolor="#F3F4F6", edgecolor="#D1D5DB", linewidth=1))
    ax.text(38, 36, "尋找 (Find)", color="#6B7280", fontsize=8.5, ha="center", va="center")

    ax.text(14, 29, "尋找目標 (Find what):", fontsize=8.5, color="#374151")
    ax.add_patch(patches.Rectangle((35, 27), 36, 4.5, facecolor="#FFFFFF", edgecolor="#2563EB", linewidth=1.5))
    ax.text(37, 29.2, " KGS", color="#DC2626", fontsize=9, fontweight="bold", fontfamily="Consolas", va="center")

    ax.text(14, 22, "取代為 (Replace with):", fontsize=8.5, color="#374151")
    ax.add_patch(patches.Rectangle((35, 20), 36, 4.5, facecolor="#FFFFFF", edgecolor="#D1D5DB", linewidth=1))
    ax.text(37, 22.2, "(完全保持空白，不輸入任何字元)", color="#9CA3AF", fontsize=7.5, fontstyle="italic", va="center")

    ax.text(14, 15, "說明：先選取『毛重』欄位，將 ' KGS' 與 ' kgs' 批次取代為空白，\n即可將文字型態字串還原為純數值！海運費欄位同理剔除 '$' 與 ' USD'。", fontsize=7.5, color="#4B5563")

    ax.add_patch(patches.Rectangle((14, 8), 16, 4.5, facecolor="#2563EB", edgecolor="none"))
    ax.text(22, 10.2, "全部取代 (A)", color="#FFFFFF", fontsize=8, fontweight="bold", ha="center", va="center")
    ax.add_patch(patches.Rectangle((32, 8), 12, 4.5, facecolor="#F3F4F6", edgecolor="#D1D5DB"))
    ax.text(38, 10.2, "取代 (R)", color="#374151", fontsize=8, ha="center", va="center")
    ax.add_patch(patches.Rectangle((46, 8), 14, 4.5, facecolor="#F3F4F6", edgecolor="#D1D5DB"))
    ax.text(53, 10.2, "尋找全部 (I)", color="#374151", fontsize=8, ha="center", va="center")
    ax.add_patch(patches.Rectangle((62, 8), 10, 4.5, facecolor="#F3F4F6", edgecolor="#D1D5DB"))
    ax.text(67, 10.2, "關閉", color="#374151", fontsize=8, ha="center", va="center")

    plt.tight_layout()
    plt.savefig(f"{out_dir}/fig_4_3_find_replace.png", dpi=300)
    plt.close()

# -------------------------------------------------------------
# 7. 圖 4-4：資料剖析精靈 (Text to Columns) 拆分船名航次
# -------------------------------------------------------------
def make_fig_4_4():
    fig, ax = plt.subplots(figsize=(12, 5.5), dpi=200)
    ax.set_facecolor("#F8FAFC")
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 50)
    ax.axis("off")

    ax.text(50, 47, "Excel 核心技能：【資料剖析精靈】三步驟拆分『船名與航次』", fontsize=13, fontweight="bold", color="#1B365D", ha="center")

    steps = [
        ("步驟 1/3：選擇分割類型", "(*) 分隔符號 (Delimited)\n  例如逗號、空格或特定字元\n( ) 固定寬度 (Fixed width)", "#1E40AF"),
        ("步驟 2/3：設定分隔符號", "[ ] 標籤  [ ] 分號  [ ] 逗號\n[V] 空格 (Space)\n[V] 其他: [ V. / - ] 依序拆分", "#0D9488"),
        ("步驟 3/3：指定目標儲存格", "欄位的資料格式：一般\n目標儲存格 (Destination):\n指定拆分寫入新欄位 C5:D5", "#16A34A")
    ]

    for i, (stitle, sdesc, scolor) in enumerate(steps):
        x = 5 + i*32
        ax.add_patch(patches.FancyBboxPatch((x, 10), 28, 33, boxstyle="round,pad=0.4", facecolor="#FFFFFF", edgecolor=scolor, linewidth=1.8))
        ax.add_patch(patches.Rectangle((x, 37), 28, 6, facecolor=scolor, edgecolor="none"))
        ax.text(x + 14, 40, stitle, color="#FFFFFF", fontsize=9, fontweight="bold", ha="center", va="center")
        lines = sdesc.split("\n")
        for l_idx, line in enumerate(lines):
            ax.text(x + 3, 32 - l_idx*5, line, color="#1F2937", fontsize=8.2, va="center")

        if i < 2:
            ax.annotate("", xy=(x + 31, 26), xytext=(x + 28.5, 26),
                        arrowprops=dict(facecolor="#94A3B8", edgecolor="#94A3B8", width=2.5, headwidth=7), zorder=5)

    ax.add_patch(patches.Rectangle((5, 2), 90, 6.5, facecolor="#FEF3C7", edgecolor="#F59E0B", linewidth=1.2))
    ax.text(50, 5.2, "剖析成果：原始欄位『EVER GIVEN V.0123W』 -> 成功拆解為『船名：EVER GIVEN』與『航次：0123W』獨立兩欄！", color="#92400E", fontsize=8.8, fontweight="bold", ha="center", va="center")

    plt.tight_layout()
    plt.savefig(f"{out_dir}/fig_4_4_text_to_columns.png", dpi=300)
    plt.close()

# -------------------------------------------------------------
# 8. 圖 4-5：資料剖析修復文字型日期
# -------------------------------------------------------------
def make_fig_4_5():
    fig, ax = plt.subplots(figsize=(9, 5), dpi=200)
    ax.set_facecolor("#F9FAFB")
    ax.set_xlim(0, 90)
    ax.set_ylim(0, 50)
    ax.axis("off")

    ax.text(45, 47, "日期大修復：利用【資料剖析精靈】一鍵將雜亂文字日期轉為標準日期", fontsize=12, fontweight="bold", color="#1B365D", ha="center")

    ax.add_patch(patches.FancyBboxPatch((10, 8), 70, 36, boxstyle="round,pad=0.4", facecolor="#FFFFFF", edgecolor="#0078D4", linewidth=1.8))
    ax.add_patch(patches.Rectangle((10, 38), 70, 6, facecolor="#0078D4", edgecolor="none"))
    ax.text(14, 41, "資料剖析精靈 - 步驟 3 之 3 (修復日期關鍵步驟)", color="#FFFFFF", fontsize=9.5, fontweight="bold", va="center")

    ax.text(14, 34, "欄位的資料格式 (Column data format):", fontsize=8.5, color="#374151")
    ax.text(16, 29, "( ) 一般 (G)   ( ) 文字 (T)   (*) 日期 (D):", fontsize=8.5, color="#1F2937")
    
    ax.add_patch(patches.Rectangle((46, 27), 16, 4.5, facecolor="#EFF6FF", edgecolor="#2563EB", linewidth=1.5))
    ax.text(54, 29.2, "YMD  (選單)", color="#1E40AF", fontsize=9, fontweight="bold", ha="center", va="center")

    ax.add_patch(patches.Rectangle((14, 12), 62, 12, facecolor="#F8FAFC", edgecolor="#CBD5E1"))
    ax.text(16, 20.5, "待修復雜亂文字: 2026.11.08 / 12-Nov-2026 / 20261115", fontsize=7.8, color="#DC2626")
    ax.text(16, 15.5, "剖析轉換後結果: 2026-11-08 / 2026-11-12 / 2026-11-15 (真日期序號)", fontsize=7.8, color="#16A34A", fontweight="bold")

    ax.annotate("直接選取 YMD (年-月-日)，\nExcel 會自動解析點號、斜線或緊湊文字！", xy=(62, 29), xytext=(65, 36),
                arrowprops=dict(facecolor="#2563EB", edgecolor="#2563EB", arrowstyle="->", lw=1.5),
                fontsize=8, fontweight="bold", color="#2563EB")

    plt.tight_layout()
    plt.savefig(f"{out_dir}/fig_4_5_date_parsing.png", dpi=300)
    plt.close()

# -------------------------------------------------------------
# 9. 圖 4-6：移除重複項 (Remove Duplicates)
# -------------------------------------------------------------
def make_fig_4_6():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5), dpi=200)
    for ax in [ax1, ax2]:
        ax.set_facecolor("#F9FAFB")
        ax.set_xlim(0, 50)
        ax.set_ylim(0, 45)
        ax.axis("off")

    ax1.set_title("步驟一：【移除重複項】核取方塊設定", fontsize=11, fontweight="bold", color="#1E40AF", pad=10)
    ax1.add_patch(patches.FancyBboxPatch((4, 4), 42, 38, boxstyle="round,pad=0.3", facecolor="#FFFFFF", edgecolor="#0078D4", linewidth=1.5))
    ax1.add_patch(patches.Rectangle((4, 36), 42, 6, facecolor="#0078D4", edgecolor="none"))
    ax1.text(7, 39, "移除重複項 (Remove Duplicates)", color="#FFFFFF", fontsize=8.5, fontweight="bold", va="center")

    ax1.text(6, 33, "若要刪除重複值，請選取一個或多個包含重複項的欄：", fontsize=7.5, color="#4B5563")
    
    ax1.add_patch(patches.Rectangle((6, 12), 38, 19, facecolor="#FFFFFF", edgecolor="#D1D5DB"))
    col_checks = [
        ("[V] 提單號碼 (B/L No.)", True),
        ("[V] 貨櫃號碼 (Container No.)", True),
        ("[ ] 船名航次", False),
        ("[ ] 裝貨港 / 卸貨港", False),
        ("[ ] 毛重 / 體積 / 運費", False)
    ]
    for i, (txt, checked) in enumerate(col_checks):
        col_c = "#1E40AF" if checked else "#4B5563"
        font_w = "bold" if checked else "normal"
        ax1.text(8, 28 - i*3.5, txt, fontsize=7.8, color=col_c, fontweight=font_w, va="center")

    ax1.text(6, 7, "航運實務準則：僅勾選提單號與櫃號作為唯一鍵值！", fontsize=7.5, color="#B91C1C", fontweight="bold")

    ax2.set_title("步驟二：執行完成統計彈窗", fontsize=11, fontweight="bold", color="#15803D", pad=10)
    ax2.add_patch(patches.FancyBboxPatch((4, 12), 42, 24, boxstyle="round,pad=0.3", facecolor="#FFFFFF", edgecolor="#107C41", linewidth=1.5))
    ax2.add_patch(patches.Rectangle((4, 30), 42, 6, facecolor="#107C41", edgecolor="none"))
    ax2.text(7, 33, "Microsoft Excel 提示", color="#FFFFFF", fontsize=8.5, fontweight="bold", va="center")
    
    ax2.text(8, 23, "[i]", color="#0078D4", fontsize=14, fontweight="bold", va="center")
    ax2.text(13, 23, "已找到並移除 2 個重複的值；\n現在還保留 28 個唯一的值。", color="#1F2937", fontsize=8.2, fontweight="bold", va="center")

    ax2.add_patch(patches.Rectangle((32, 14), 10, 4, facecolor="#107C41", edgecolor="none"))
    ax2.text(37, 16, "確定", color="#FFFFFF", fontsize=8, fontweight="bold", ha="center", va="center")

    plt.tight_layout()
    plt.savefig(f"{out_dir}/fig_4_6_remove_duplicates.png", dpi=300)
    plt.close()

# -------------------------------------------------------------
# 10. 圖 4-7：選擇性貼上乘運算消除文字數字綠色三角
# -------------------------------------------------------------
def make_fig_4_7():
    fig, ax = plt.subplots(figsize=(10, 5), dpi=200)
    ax.set_facecolor("#F9FAFB")
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 50)
    ax.axis("off")

    ax.text(50, 47, "神級技巧：【選擇性貼上】乘運算瞬間消除綠色錯誤三角提示", fontsize=12, fontweight="bold", color="#1B365D", ha="center")

    ax.add_patch(patches.FancyBboxPatch((5, 10), 38, 33, boxstyle="round,pad=0.3", facecolor="#FFFFFF", edgecolor="#DC2626", linewidth=1.5))
    ax.add_patch(patches.Rectangle((5, 37), 38, 6, facecolor="#FEE2E2", edgecolor="none"))
    ax.text(24, 40, "常見困擾：以文字儲存的數字", color="#991B1B", fontsize=8.5, fontweight="bold", ha="center", va="center")
    
    ax.add_patch(patches.Rectangle((10, 24), 28, 8, facecolor="#FFFFFF", edgecolor="#D1D5DB", linewidth=1))
    ax.add_patch(patches.Polygon([[10, 32], [14, 32], [10, 28]], facecolor="#16A34A", edgecolor="none"))
    ax.text(16, 28, "2400", color="#1F2937", fontsize=12, fontfamily="Consolas", va="center")
    ax.text(24, 18, "[!] 儲存格靠左對齊，無法用 SUM 加總！\n提示訊息：『此儲存格中的數字已格式化為文字』", fontsize=7.2, color="#B91C1C", ha="center")

    ax.add_patch(patches.FancyBboxPatch((48, 8), 47, 35, boxstyle="round,pad=0.3", facecolor="#FFFFFF", edgecolor="#0078D4", linewidth=1.5))
    ax.add_patch(patches.Rectangle((48, 37), 47, 6, facecolor="#0078D4", edgecolor="none"))
    ax.text(52, 40, "選擇性貼上 (Paste Special) - 運算選【乘】", color="#FFFFFF", fontsize=8.5, fontweight="bold", va="center")

    ax.text(52, 33, "貼上 (Paste):  (*) 全部 (A)   ( ) 數值 (V)   ( ) 格式 (T)", fontsize=7.5, color="#4B5563")
    ax.text(52, 28, "運算 (Operation):", fontsize=8, fontweight="bold", color="#1E40AF")
    ax.text(54, 24, "( ) 無 (N)        ( ) 加 (D)", fontsize=7.8, color="#374151")
    ax.text(54, 20, "( ) 減 (S)        (*) 乘 (M)  [點選此項]", fontsize=8, fontweight="bold", color="#DC2626")
    ax.text(54, 16, "( ) 除 (I)", fontsize=7.8, color="#374151")

    ax.annotate("", xy=(47, 26), xytext=(44, 26),
                arrowprops=dict(facecolor="#2563EB", edgecolor="#2563EB", width=2.5, headwidth=7))
    ax.text(45.5, 29, "在空白格輸入 1 複製\n再對目標範圍選【乘】", fontsize=7.2, color="#2563EB", ha="center")

    plt.tight_layout()
    plt.savefig(f"{out_dir}/fig_4_7_paste_special_multiply.png", dpi=300)
    plt.close()

# -------------------------------------------------------------
# 11. 圖 4-8：設定格式化的條件標示重複與缺漏
# -------------------------------------------------------------
def make_fig_4_8():
    fig, ax = plt.subplots(figsize=(9, 4.8), dpi=200)
    ax.set_facecolor("#F9FAFB")
    ax.set_xlim(0, 90)
    ax.set_ylim(0, 45)
    ax.axis("off")

    ax.text(45, 42, "視覺化查核：【設定格式化的條件】標記重複提單與遺漏值", fontsize=12, fontweight="bold", color="#1B365D", ha="center")

    headers = [("提單號碼", 5, 25), ("船名航次", 25, 48), ("毛重 (KG)", 48, 68), ("報關放行狀態", 68, 88)]
    for h, x1, x2 in headers:
        ax.add_patch(patches.Rectangle((x1, 30), x2 - x1, 6, facecolor="#1E3A8A", edgecolor="#CBD5E1"))
        ax.text((x1 + x2)/2, 33, h, color="#FFFFFF", fontsize=8.5, fontweight="bold", ha="center", va="center")

    sample_rows = [
        (["ONE8877665", "ONE HARBOR V.022W", "16,800 KGS", "已放行"], True, False, False),
        (["OOCLU9823103", "OOCL SHANGHAI V.078N", "", "待補單"], False, True, False),
        (["ONE8877665", "ONE HARBOR V.022W", "16,800 KGS", "已放行"], True, False, False),
        (["OOCLU9823104", "OOCL GERMANY V.003E", "20,100 KGS", ""], False, False, True)
    ]

    for r_idx, (rvals, is_dup, is_weight_blank, is_status_blank) in enumerate(sample_rows):
        y = 24 - r_idx*5.5
        for c_idx, (h, x1, x2) in enumerate(headers):
            bg = "#FFFFFF"
            border = "#E2E8F0"
            if c_idx == 0 and is_dup:
                bg = "#FEE2E2"
                border = "#EF4444"
            elif c_idx == 2 and is_weight_blank:
                bg = "#FEF3C7"
                border = "#F59E0B"
            elif c_idx == 3 and is_status_blank:
                bg = "#FEF3C7"
                border = "#F59E0B"

            ax.add_patch(patches.Rectangle((x1, y), x2 - x1, 5.5, facecolor=bg, edgecolor=border, linewidth=1.2))
            val = rvals[c_idx]
            if val == "":
                val = "[空白儲存格 !]"
            tc = "#DC2626" if (c_idx == 0 and is_dup) else ("#D97706" if "[空白" in val else "#1F2937")
            ax.text((x1 + x2)/2, y + 2.7, val, color=tc, fontsize=7.8, fontweight="bold" if is_dup or "[空白" in val else "normal", ha="center", va="center")

    ax.add_patch(patches.Rectangle((15, 0.5), 60, 4.5, facecolor="#F1F5F9", edgecolor="#CBD5E1"))
    ax.add_patch(patches.Rectangle((18, 1.8), 3, 2, facecolor="#FEE2E2", edgecolor="#EF4444"))
    ax.text(22, 2.8, "紅色高亮：重複值規則標記", fontsize=7.5, color="#991B1B", va="center")
    ax.add_patch(patches.Rectangle((50, 1.8), 3, 2, facecolor="#FEF3C7", edgecolor="#F59E0B"))
    ax.text(54, 2.8, "黃色高亮：遺漏值 / 空白格標記", fontsize=7.5, color="#92400E", va="center")

    plt.tight_layout()
    plt.savefig(f"{out_dir}/fig_4_8_conditional_formatting.png", dpi=300)
    plt.close()

# -------------------------------------------------------------
# 12. 圖 5-1：字串清洗函數 TRIM()、UPPER()、CLEAN() 原理
# -------------------------------------------------------------
def make_fig_5_1():
    fig, ax = plt.subplots(figsize=(11, 4.5), dpi=200)
    ax.set_facecolor("#F8FAFC")
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 45)
    ax.axis("off")

    ax.text(50, 41, "文字清洗核心管線：TRIM()、CLEAN() 與 UPPER() 串聯運算", fontsize=12, fontweight="bold", color="#1B365D", ha="center")

    boxes = [
        ("原始髒字串", "  \"  ooclu9823101 \\n  \"  \n* 前後包夾多餘空白\n* 夾帶不可見換行符號\n* 英文字母為小寫", 5, 25, "#EF4444"),
        ("CLEAN() 清洗", "CLEAN(A1)\n剔除 ASCII 0~31 不可列印字元\n消除系統匯出換行符", 28, 22, "#D97706"),
        ("TRIM() 去空格", "TRIM(...)\n自動刪除首尾所有半形全形空格\n單字間若有多個空格只保留一個", 53, 22, "#2563EB"),
        ("UPPER() 標準化", "UPPER(...)\n所有英文代碼強制轉為大寫\n產出標準唯一鍵值", 78, 19, "#16A34A")
    ]

    for title, desc, bx, bw, color in boxes:
        ax.add_patch(patches.FancyBboxPatch((bx, 8), bw, 26, boxstyle="round,pad=0.3", facecolor="#FFFFFF", edgecolor=color, linewidth=1.8))
        ax.add_patch(patches.Rectangle((bx, 28), bw, 6, facecolor=color, edgecolor="none"))
        ax.text(bx + bw/2, 31, title, color="#FFFFFF", fontsize=8.8, fontweight="bold", ha="center", va="center")
        lines = desc.split("\n")
        for l_idx, line in enumerate(lines):
            ax.text(bx + bw/2, 24 - l_idx*4.2, line, color="#374151", fontsize=7.5, ha="center", va="center")

    ax.annotate("", xy=(27.5, 21), xytext=(25.5, 21), arrowprops=dict(facecolor="#64748B", edgecolor="#64748B", width=2, headwidth=6))
    ax.annotate("", xy=(52.5, 21), xytext=(50.5, 21), arrowprops=dict(facecolor="#64748B", edgecolor="#64748B", width=2, headwidth=6))
    ax.annotate("", xy=(77.5, 21), xytext=(75.5, 21), arrowprops=dict(facecolor="#64748B", edgecolor="#64748B", width=2, headwidth=6))

    ax.add_patch(patches.Rectangle((5, 1.5), 90, 4.5, facecolor="#EFF6FF", edgecolor="#3B82F6"))
    ax.text(50, 3.7, "終極單行整合公式： =UPPER(TRIM(CLEAN(B2)))  ->  瞬間產出 100% 潔淨的標準提單代碼！", color="#1E40AF", fontsize=8.5, fontweight="bold", ha="center", va="center")

    plt.tight_layout()
    plt.savefig(f"{out_dir}/fig_5_1_trim_upper_clean.png", dpi=300)
    plt.close()

# -------------------------------------------------------------
# 13. 圖 5-2：VLOOKUP / XLOOKUP 跨表關聯查詢機制
# -------------------------------------------------------------
def make_fig_5_2():
    fig, ax = plt.subplots(figsize=(11, 5), dpi=200)
    ax.set_facecolor("#F8FAFC")
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 50)
    ax.axis("off")

    ax.text(50, 46.5, "代碼標準化大師：VLOOKUP / XLOOKUP 跨表自動精準關聯", fontsize=12, fontweight="bold", color="#1B365D", ha="center")

    ax.add_patch(patches.FancyBboxPatch((5, 8), 38, 33, boxstyle="round,pad=0.3", facecolor="#FFFFFF", edgecolor="#2563EB", linewidth=1.5))
    ax.add_patch(patches.Rectangle((5, 35), 38, 6, facecolor="#2563EB", edgecolor="none"))
    ax.text(24, 38, "交易工作表：01_原始資料", color="#FFFFFF", fontsize=8.5, fontweight="bold", ha="center", va="center")
    
    ax.text(7, 30, "提單號碼       裝貨港 (POL 查閱值)", fontsize=7.8, fontweight="bold", color="#4B5563")
    ax.text(7, 25, "OOCLU9823101   TWKHH", fontsize=7.8, color="#1F2937")
    ax.text(7, 20, "EMCE8901234   高雄   <- 【查閱目標】", fontsize=7.8, color="#DC2626", fontweight="bold")
    ax.text(7, 15, "MSK7654321    KHH", fontsize=7.8, color="#1F2937")
    ax.text(7, 10, "WHL9876543    基隆", fontsize=7.8, color="#1F2937")

    ax.add_patch(patches.FancyBboxPatch((55, 8), 40, 33, boxstyle="round,pad=0.3", facecolor="#FFFFFF", edgecolor="#0D9488", linewidth=1.5))
    ax.add_patch(patches.Rectangle((55, 35), 40, 6, facecolor="#0D9488", edgecolor="none"))
    ax.text(75, 38, "代碼主檔：02_港口標準代碼表", color="#FFFFFF", fontsize=8.5, fontweight="bold", ha="center", va="center")

    ax.text(57, 30, "港口別稱 (比對欄)   標準UN代碼   中文港名", fontsize=7.8, fontweight="bold", color="#4B5563")
    ax.text(57, 25, "TWKHH             TWKHH       高雄港", fontsize=7.5, color="#1F2937")
    ax.text(57, 20, "高雄              TWKHH       高雄港 <-", fontsize=7.8, color="#0D9488", fontweight="bold")
    ax.text(57, 15, "KHH               TWKHH       高雄港", fontsize=7.5, color="#1F2937")
    ax.text(57, 10, "基隆              TWKEL       基隆港", fontsize=7.5, color="#1F2937")

    ax.annotate("VLOOKUP(F2, 代碼表!$A$2:$C$12, 2, FALSE)", xy=(73, 20), xytext=(35, 12),
                arrowprops=dict(facecolor="#DC2626", edgecolor="#DC2626", arrowstyle="->", lw=2, connectionstyle="arc3,rad=-0.2"),
                fontsize=8.5, fontweight="bold", color="#DC2626", bbox=dict(boxstyle="round,pad=0.3", fc="#FEF2F2", ec="#DC2626"))

    plt.tight_layout()
    plt.savefig(f"{out_dir}/fig_5_2_vlookup_port_mapping.png", dpi=300)
    plt.close()

# -------------------------------------------------------------
# 14. 圖 6-1：300筆航運情境練習檔概覽
# -------------------------------------------------------------
def make_fig_6_1():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5), dpi=200)

    trades = ['整櫃海運 (FCL)', '併櫃海運 (LCL)', '冷藏保鮮 (Reefer)', '危險品櫃 (DG)']
    counts = [150, 60, 45, 45]
    colors = ['#1B365D', '#008080', '#0284C7', '#EA580C']
    ax1.pie(counts, labels=trades, autopct='%1.1f%%', startangle=140, colors=colors,
            textprops={'fontsize': 8.5, 'color': '#1F2937'})
    ax1.set_title("300筆情境練習檔：業務型態分佈結構", fontsize=10.5, fontweight="bold", color="#1B365D")

    defects = ['重複提單記錄', '港口俗稱/簡寫', '首尾多餘空格', '雜亂日期格式', '數值黏雜單位', '英文小寫混雜', '關鍵毛重缺漏']
    d_counts = [15, 78, 85, 120, 145, 68, 14]
    y_pos = range(len(defects))
    ax2.barh(y_pos, d_counts, color="#2563EB", edgecolor="#1D4ED8", height=0.6)
    ax2.set_yticks(y_pos)
    ax2.set_yticklabels(defects, fontsize=8)
    ax2.invert_yaxis()
    ax2.set_xlabel("潛藏筆數 (含交集重疊)", fontsize=8.5)
    ax2.set_title("300筆原始資料：資料缺陷分佈診斷", fontsize=10.5, fontweight="bold", color="#1B365D")
    for i, v in enumerate(d_counts):
        ax2.text(v + 3, i, str(v), color='#1E40AF', fontweight='bold', fontsize=8, va='center')
    ax2.set_xlim(0, 170)

    plt.tight_layout()
    plt.savefig(f"{out_dir}/fig_6_1_practice_300_overview.png", dpi=300)
    plt.close()

make_fig_1_1()
make_fig_2_1()
make_fig_2_2()
make_fig_4_1()
make_fig_4_2()
make_fig_4_3()
make_fig_4_4()
make_fig_4_5()
make_fig_4_6()
make_fig_4_7()
make_fig_4_8()
make_fig_5_1()
make_fig_5_2()
make_fig_6_1()
print("Clean regenerated 14 figures successfully.")
