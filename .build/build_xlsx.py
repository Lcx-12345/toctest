import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import DataBarRule
from openpyxl.utils import get_column_letter

OUT = "/workspace/ai-trends-2026/2026 AI 应用趋势数据.xlsx"

INK = "1A1F2B"
DARK = "1F2A3D"
CYAN = "0E8F88"
LIGHT = "F2F7FA"
BAND = "F7FAFC"
BORDER_C = "DCE3EC"

thin = Side(style="thin", color=BORDER_C)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
FONT_CN = "Noto Sans CJK SC"

wb = Workbook()


def style_sheet(ws, title, subtitle, headers, rows, widths, start_row=4):
    # 标题
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    c = ws.cell(row=1, column=1, value=title)
    c.font = Font(name=FONT_CN, size=16, bold=True, color=INK)
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 30

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
    c = ws.cell(row=2, column=1, value=subtitle)
    c.font = Font(name=FONT_CN, size=9.5, color="6B7488")
    c.alignment = Alignment(vertical="center")
    ws.row_dimensions[2].height = 20

    # 表头
    hr = start_row
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(row=hr, column=i, value=h)
        cell.font = Font(name=FONT_CN, size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=DARK)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[hr].height = 28

    # 数据
    for r, row in enumerate(rows, start=hr + 1):
        for i, v in enumerate(row, start=1):
            cell = ws.cell(row=r, column=i, value=v)
            cell.font = Font(name=FONT_CN, size=10, color=INK)
            cell.border = BORDER
            cell.alignment = Alignment(
                horizontal="center" if i > 1 else "left",
                vertical="center", wrap_text=True,
            )
            if (r - hr) % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=BAND)
        ws.row_dimensions[r].height = 24

    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = ws.cell(row=hr + 1, column=1)
    ws.sheet_view.showGridLines = False
    return hr + len(rows)


# ================= Sheet 1 市场概览 =================
ws1 = wb.active
ws1.title = "市场概览"
rows1 = [
    ["2022", 400, None, "技术验证"],
    ["2023", 670, 0.68, "能力竞赛"],
    ["2024", 1370, 1.05, "试点扩散"],
    ["2025", 2420, 0.77, "场景收敛"],
    ["2026", 3980, 0.64, "规模落地"],
    ["2027（预测）", 5960, 0.50, "深度整合"],
    ["2028（预测）", 8220, 0.38, "常态渗透"],
]
last1 = style_sheet(
    ws1,
    "全球生成式 AI 市场规模",
    "单位：亿美元｜2026 年及以后为示意性预测，仅用于趋势判断",
    ["年份", "市场规模（亿美元）", "同比增速", "阶段特征"],
    rows1,
    [14, 20, 12, 18],
)
for r in range(5, last1 + 1):
    ws1.cell(row=r, column=3).number_format = "0%"

# 折线图
ch1 = LineChart()
ch1.title = "市场规模走势（亿美元）"
ch1.style = 12
ch1.height = 8.5
ch1.width = 17
data = Reference(ws1, min_col=2, min_row=4, max_row=last1)
cats = Reference(ws1, min_col=1, min_row=5, max_row=last1)
ch1.add_data(data, titles_from_data=True)
ch1.set_categories(cats)
ch1.y_axis.title = "亿美元"
ch1.legend = None
ch1.series[0].graphicalProperties.line.solidFill = "0E8F88"
ch1.series[0].graphicalProperties.line.width = 28000
ch1.series[0].marker.symbol = "circle"
ch1.series[0].marker.size = 6
ch1.series[0].marker.graphicalProperties.solidFill = "0E8F88"
ws1.add_chart(ch1, "F4")

# ================= Sheet 2 行业渗透 =================
ws2 = wb.create_sheet("行业渗透")
rows2 = [
    ["软件与互联网", 0.78, "研发提效、内容生成", "质量与安全"],
    ["金融", 0.62, "风控、投研、客服", "合规与审计"],
    ["专业服务", 0.58, "文档、检索、交付", "知识资产化"],
    ["零售电商", 0.52, "素材生产、推荐、客服", "品牌一致性"],
    ["制造", 0.47, "质检、排产、知识库", "数据分散"],
    ["教育", 0.44, "备课、个性化辅导", "内容准确性"],
    ["医疗健康", 0.39, "辅助诊断、病历整理", "数据与责任"],
    ["政务", 0.35, "政务问答、材料审核", "安全与可控"],
]
last2 = style_sheet(
    ws2,
    "2026 年各行业企业采用率",
    "口径：已进入生产环境的企业占比｜公开研究综合整理",
    ["行业", "企业采用率", "典型场景", "主要障碍"],
    rows2,
    [16, 14, 26, 16],
)
for r in range(5, last2 + 1):
    ws2.cell(row=r, column=2).number_format = "0%"

ch2 = BarChart()
ch2.type = "bar"
ch2.title = "行业采用率对比"
ch2.style = 12
ch2.height = 9.5
ch2.width = 17
data = Reference(ws2, min_col=2, min_row=4, max_row=last2)
cats = Reference(ws2, min_col=1, min_row=5, max_row=last2)
ch2.add_data(data, titles_from_data=True)
ch2.set_categories(cats)
ch2.legend = None
ch2.x_axis.title = "采用率"
ch2.x_axis.number_format = "0%"
ch2.series[0].graphicalProperties.solidFill = "5A4BC8"
ws2.add_chart(ch2, "F4")

# ================= Sheet 3 区域对比 =================
ws3 = wb.create_sheet("区域对比")
rows3 = [
    ["北美", 0.71, 1860, 0.38],
    ["中国", 0.58, 980, 0.46],
    ["欧洲", 0.54, 720, 0.33],
    ["亚太其他", 0.49, 520, 0.41],
    ["拉美", 0.41, 180, 0.35],
    ["中东非", 0.33, 140, 0.39],
]
last3 = style_sheet(
    ws3,
    "区域市场对比",
    "单位：亿美元｜采用率为已进入生产环境的企业占比",
    ["区域", "企业采用率", "2026 年投入（亿美元）", "同比增速"],
    rows3,
    [14, 14, 22, 12],
)
for r in range(5, last3 + 1):
    ws3.cell(row=r, column=2).number_format = "0%"
    ws3.cell(row=r, column=4).number_format = "0%"

ch3 = BarChart()
ch3.title = "区域 AI 投入（亿美元）"
ch3.style = 12
ch3.height = 8.5
ch3.width = 17
data = Reference(ws3, min_col=3, min_row=4, max_row=last3)
cats = Reference(ws3, min_col=1, min_row=5, max_row=last3)
ch3.add_data(data, titles_from_data=True)
ch3.set_categories(cats)
ch3.legend = None
ch3.series[0].graphicalProperties.solidFill = "0E8F88"
ch3.dataLabels = DataLabelList()
ch3.dataLabels.showVal = True
ws3.add_chart(ch3, "F4")

# ================= Sheet 4 优先级矩阵 =================
ws4 = wb.create_sheet("优先级矩阵")
rows4 = [
    ["智能体工作流", 3.8, 4.7, 4.7, "P0", "选定 1–2 个高频流程端到端改造"],
    ["多模态内容生成", 4.2, 4.3, 4.3, "P0", "建立品牌风格库与批量生产管线"],
    ["AI 原生软件工程", 3.6, 4.5, 4.5, "P0", "在测试与评审环节先行落地"],
    ["企业 AI 治理", 3.1, 4.0, 4.0, "P1", "建立模型清单与审计留痕机制"],
    ["端侧小模型与硬件", 2.9, 3.8, 3.8, "P1", "评估隐私敏感场景本地化方案"],
    ["AI 科学发现", 2.4, 3.9, 3.9, "P2", "与科研机构合作试点"],
    ["具身智能与机器人", 2.1, 3.4, 3.4, "P2", "保持技术跟踪，暂不重投入"],
]
last4 = style_sheet(
    ws4,
    "趋势落地优先级矩阵",
    "综合得分 = 成熟度 × 0.4 + 业务价值 × 0.6（满分 5 分）",
    ["趋势方向", "技术成熟度", "业务价值", "综合得分", "优先级", "建议动作"],
    rows4,
    [20, 14, 12, 12, 10, 34],
)
for r in range(5, last4 + 1):
    ws4.cell(row=r, column=2).number_format = "0.0"
    ws4.cell(row=r, column=3).number_format = "0.0"
    ws4.cell(row=r, column=4).number_format = "0.0"
    p = ws4.cell(row=r, column=5)
    p.font = Font(name=FONT_CN, size=10, bold=True,
                  color={"P0": "C2185B", "P1": "2F5FD0", "P2": "5B6779"}[p.value])
    p.alignment = Alignment(horizontal="center", vertical="center")

ws4.conditional_formatting.add(
    f"D5:D{last4}",
    DataBarRule(start_type="num", start_value=0, end_type="num", end_value=5,
                color="0E8F88", showValue=True),
)

wb.save(OUT)
print("saved:", OUT, os.path.getsize(OUT))
