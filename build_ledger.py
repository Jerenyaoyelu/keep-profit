# -*- coding: utf-8 -*-
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border
from openpyxl.styles.borders import Side
from openpyxl.utils import get_column_letter

OUT = "/home/user/Doubao/chats/38454771784079874/中兴通讯监控/中兴通讯交易台账.xlsx"

HEAD_FILL = PatternFill("solid", fgColor="1F2933")
HEAD_FONT = Font(name="微软雅黑", size=10.5, bold=True, color="FFFFFF")
BODY_FONT = Font(name="微软雅黑", size=10.5)
TITLE_FONT = Font(name="微软雅黑", size=13, bold=True, color="1F2933")
NOTE_FONT = Font(name="微软雅黑", size=9, color="6B7785")
RED = Font(name="微软雅黑", size=10.5, bold=True, color="C0392B")
GREEN = Font(name="微软雅黑", size=10.5, bold=True, color="1E8449")
thin = Side(style="thin", color="D5DBE1")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def setup(ws, title, headers, widths, note=""):
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws.row_dimensions[1].height = 26
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=2, column=i, value=h)
        c.fill = HEAD_FILL
        c.font = HEAD_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
    ws.row_dimensions[2].height = 30
    ws.freeze_panes = "A3"
    if note:
        ws.cell(row=2 + 40, column=1, value=note).font = NOTE_FONT


def write_rows(ws, start_row, rows, num_cols=(), pct_cols=(), center_cols=()):
    for r, row in enumerate(rows, start=start_row):
        for c, val in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=True,
                                       horizontal="center" if c in center_cols else "left")
            if c in num_cols:
                cell.number_format = "0.000"
            if c in pct_cols:
                cell.number_format = "0.00%"
        ws.row_dimensions[r].height = 24


wb = Workbook()

# ---------- 1. 参数与持仓 ----------
ws = wb.active
ws.title = "参数与持仓"
setup(ws, "中兴通讯（000063.SZ）持仓参数 · 规则执行基准", ["项目", "数值", "说明与口径"], [22, 26, 58])
params = [
    ["标的", "中兴通讯", "深市 A 股"],
    ["代码", "000063.SZ", "H 股为 00763.HK"],
    ["持仓成本", 27.716, "保本需 +2.08%（现价 26.43）"],
    ["当前价", 26.43, "2026-11-20 14:03 盘中快照，收盘以 15:00 价为准"],
    ["浮动盈亏", -0.0446, "（26.43 − 27.716）/ 27.716，未计费用"],
    ["仓位档位", "重仓档（>30% 总资产）", "减仓区因此前移至 26.90–27.00，不等反抽回成本"],
    ["保本线", 27.72, "≈ MA5 27.68，属第一技术阻力"],
    ["第一减仓区", "26.90–27.00", "重仓档触发即减 1/3"],
    ["加仓区", "24.60–25.20", "需 ≥2 个触底信号成立，单次 30%"],
    ["硬止损", 24.40, "收盘价确认，盘中插针不算"],
    ["绿灯确认位", 28.03, "MA20；需配合成交额>15亿且主力连续2日净流入"],
    ["板块破位线", 5031, "通信设备（申万）指数 11-17 低点"],
    ["数据更新时间", "2026-11-20 14:03", "由定时任务每个交易日 10:30 / 13:30 / 15:10 更新"],
]
write_rows(ws, 3, params, num_cols=(2,))
ws["B5"].number_format = "0.000"
ws["B6"].number_format = "0.000"
ws["B7"].number_format = "0.00%"
ws["B7"].font = RED
ws["B11"].number_format = "0.000"

# ---------- 2. 规则表 ----------
ws2 = wb.create_sheet("规则表")
setup(ws2, "规则状态机：写死，不临场改口", ["编号", "规则名称", "触发条件（收盘口径确认）", "状态", "动作", "当前是否触发"],
      [8, 18, 42, 10, 40, 14])
rules = [
    ["R1", "硬止损", "收盘价 < 24.40", "红灯", "无条件清仓，不参与任何反弹", "未触发"],
    ["R2", "破位降仓", "收盘价 < 26.55 或 板块指数 < 5031", "红灯", "反弹至 26.90–27.00 减 1/3；不反弹则持到 R5", "已触发"],
    ["R3", "禁止买入", "主力净流出 > 1 亿 且 流出强度 > 15%", "红灯", "当日禁止任何买入操作", "已触发（20.6%）"],
    ["R4", "缩量观望", "量比 < 1 但 主力净流出 > 1 亿", "黄灯", "观望，不补不追，不加仓", "已触发"],
    ["R5", "底部回补", "价格进入 24.60–25.20 且 触底信号 ≥2 个", "绿灯", "补仓 30%，测算成本降至约 27.09", "未触发"],
    ["R6", "趋势回补", "收盘 > 28.03 且 成交额 > 15 亿 且 主力连续 2 日净流入", "绿灯", "回补至总仓位 ≤25%，单次 ≤20%", "未触发"],
    ["", "触底信号清单", "① 成交额>35亿且收长下影 ② 次日缩量<15亿不创新低 ③ 主力连续2日净流入 ④ 板块收复5031 ⑤ 日线MACD底背离", "—", "任意 2 个成立才允许 R5", "—"],
]
write_rows(ws2, 3, rules, center_cols=(1, 4, 6))
for r in (3, 4, 5, 6):
    ws2.cell(row=r, column=4).font = RED
ws2.cell(row=7, column=4).font = Font(name="微软雅黑", size=10.5, bold=True, color="B9770E")
ws2.cell(row=8, column=4).font = GREEN
ws2.cell(row=9, column=4).font = GREEN

# ---------- 3. 交易记录 ----------
ws3 = wb.create_sheet("交易记录")
setup(ws3, "交易记录：每次操作必须填写触发规则编号，否则次日不生成建议",
      ["日期", "方向", "成交价", "数量(股)", "金额(元)", "触发规则编号", "决策理由（必填）", "事后归因（14 日后填）", "是否按规则"],
      [12, 10, 11, 12, 14, 14, 40, 32, 12])
write_rows(ws3, 3, [], center_cols=(1, 2, 3, 4, 6, 9))
# 预填一条建仓记录占位（成本已知，数量未知，留空待用户补）
row = ["2026-11（待补）", "买入（建仓）", 27.716, None, None, "—", "建仓（历史，未记录规则依据）", "", "待补"]
for c, v in enumerate(row, start=1):
    cell = ws3.cell(row=3, column=c, value=v)
    cell.font = BODY_FONT
    cell.border = BORDER
    cell.alignment = Alignment(vertical="center", wrap_text=True,
                               horizontal="center" if c in (1, 2, 3, 4, 6, 9) else "left")
ws3["C3"].number_format = "0.000"
ws3["E3"].number_format = "#,##0"

# ---------- 4. 信号历史 ----------
ws4 = wb.create_sheet("信号历史")
setup(ws4, "每日信号：由定时任务自动追加，执行情况需人工回填",
      ["日期", "时间", "状态", "触发规则", "最新价", "主力净流入(亿)", "流出强度", "关键盘面", "建议动作", "是否执行"],
      [12, 9, 9, 16, 10, 16, 12, 46, 34, 12])
signals = [
    ["2026-11-19", "收盘", "黄灯", "R3 临界", 27.14, -2.227, 0.117, "放量下跌，成交额 19.08 亿，换手 1.75%，板块 +0.15% 但个股跑输 3.1pct", "跌破 26.55 即减仓", "未执行"],
    ["2026-11-20", "14:03", "红灯", "R2 + R3 + R4", 26.43, -2.193, 0.206, "低见 26.26 创两年新低；缩量（量比 0.92）但主力流出强度翻倍；板块 4958 未收复 5031", "反弹 26.90–27.00 减 1/3", "待记录"],
]
write_rows(ws4, 3, signals, num_cols=(5,), pct_cols=(7,), center_cols=(1, 2, 3, 4, 5, 6, 7, 10))
ws4.cell(row=3, column=3).font = Font(name="微软雅黑", size=10.5, bold=True, color="B9770E")
ws4.cell(row=4, column=3).font = RED
ws4.cell(row=4, column=10).font = Font(name="微软雅黑", size=10.5, bold=True, color="B9770E")

# ---------- 5. 复盘统计 ----------
ws5 = wb.create_sheet("复盘统计")
setup(ws5, "复盘统计：用同一套规则比较「遵守」与「违背」的胜率", ["指标", "数值", "说明"], [26, 22, 64])
stats = [
    ["交易笔数（已记录）", '=COUNTA(交易记录!A3:A500)', "仅统计填了日期的行"],
    ["按规则执行笔数", '=COUNTIF(交易记录!I3:I500,"是")', "触发规则与动作一致"],
    ["违背规则笔数", '=COUNTIF(交易记录!I3:I500,"否")', "主观干预"],
    ["规则遵守率", '=IFERROR(COUNTIF(交易记录!I3:I500,"是")/COUNTA(交易记录!A3:A500),0)', "目标 ≥80%；低于 60% 说明工具无效，问题在执行"],
    ["红灯期间买入次数", '=COUNTIFS(信号历史!C3:C500,"红灯",交易记录!B3:B500,"买入")', "理想值 0"],
    ["信号触发次数", '=COUNTA(信号历史!A3:A500)', "自动生成"],
    ["未执行信号次数", '=COUNTIF(信号历史!J3:J500,"未执行")', "每次未执行都要在归因列写原因"],
    ["平均浮亏（当前）", -0.0446, "按 26.43 / 27.716 计算"],
]
write_rows(ws5, 3, stats)
ws5["B10"].number_format = "0.00%"
ws5["B9"].font = RED
for r in range(3, 10):
    ws5.cell(row=r, column=2).number_format = "0.00%" if r in (6,) else "0"

note = "使用说明：① 每次买卖后当天在「交易记录」补一行，触发规则编号必填；② 「信号历史」由定时任务自动追加，是否执行需回填；③ 14 天后在「事后归因」写清当时判断与结果的差异；④ 复盘统计会自动算出遵守率，两个月后这份数据比任何指标都更能说明你的问题。"
ws5.cell(row=13, column=1, value=note).font = NOTE_FONT
ws5.merge_cells(start_row=13, start_column=1, end_row=13, end_column=3)
ws5.cell(row=13, column=1).alignment = Alignment(wrap_text=True, vertical="top")
ws5.row_dimensions[13].height = 60

wb.save(OUT)
print("saved", OUT, os.path.getsize(OUT))
for s in wb.sheetnames:
    print(s)
