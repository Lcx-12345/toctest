import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = "/workspace/ai-trends-2026/2026 AI 应用趋势报告.docx"
FONT = "Noto Sans CJK SC"
INK = RGBColor(0x1A, 0x1F, 0x2B)
GRAY = RGBColor(0x6B, 0x74, 0x88)
CYAN = RGBColor(0x0E, 0x8F, 0x88)
VIOLET = RGBColor(0x5A, 0x4B, 0xC8)

doc = Document()

# ---------- 页面 ----------
sec = doc.sections[0]
sec.page_width = Cm(21.0)
sec.page_height = Cm(29.7)
sec.top_margin = Cm(2.4)
sec.bottom_margin = Cm(2.2)
sec.left_margin = Cm(2.6)
sec.right_margin = Cm(2.6)

# ---------- 默认样式 ----------
normal = doc.styles["Normal"]
normal.font.name = FONT
normal.font.size = Pt(10.5)
normal.font.color.rgb = INK
normal._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
normal.paragraph_format.line_spacing = 1.65
normal.paragraph_format.space_after = Pt(6)


def set_cn(run, size=None, bold=None, color=None, font=FONT):
    run.font.name = font
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


def para(text="", size=10.5, bold=False, color=INK, align=None,
         space_before=0, space_after=6, line=1.65, indent=0):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line
    if indent:
        pf.first_line_indent = Pt(indent)
    if align is not None:
        p.alignment = align
    if text:
        set_cn(p.add_run(text), size=size, bold=bold, color=color)
    return p


def h1(text, num=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(22)
    p.paragraph_format.space_after = Pt(10)
    p.paragraph_format.keep_with_next = True
    if num:
        set_cn(p.add_run(num + "  "), size=19, bold=True, color=CYAN)
    set_cn(p.add_run(text), size=19, bold=True, color=INK)
    # 分隔线
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "6")
    bottom.set(qn("w:color"), "DCE3EC")
    pbdr.append(bottom)
    pPr.append(pbdr)
    return p


def h2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    set_cn(p.add_run(text), size=12.5, bold=True, color=VIOLET)
    return p


def bullet(text, bold_prefix=None):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.55
    p.paragraph_format.left_indent = Cm(0.9)
    p.paragraph_format.first_line_indent = Cm(-0.4)
    if bold_prefix:
        set_cn(p.add_run(bold_prefix), size=10.5, bold=True)
    set_cn(p.add_run(text), size=10.5)
    return p


def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)


def cell_text(cell, text, bold=False, color=INK, size=9.5, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.25
    set_cn(p.add_run(text), size=size, bold=bold, color=color)


def make_table(headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        cell_text(hdr[i], h, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), size=9.5)
        shade(hdr[i], "1F2A3D")
    for r in rows:
        cells = t.add_row().cells
        for i, v in enumerate(r):
            cell_text(cells[i], v, size=9.5)
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    # 表后间距
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def callout(title, text):
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = t.rows[0].cells[0]
    shade(cell, "F2F7FA")
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    set_cn(p.add_run(title), size=10.5, bold=True, color=CYAN)
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(8)
    p2.paragraph_format.line_spacing = 1.55
    set_cn(p2.add_run(text), size=10, color=RGBColor(0x33, 0x3B, 0x4D))
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


# ---------- 页脚页码 ----------
footer_p = sec.footer.paragraphs[0]
footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = footer_p.add_run()
fld1 = OxmlElement("w:fldChar"); fld1.set(qn("w:fldCharType"), "begin")
instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = " PAGE "
fld2 = OxmlElement("w:fldChar"); fld2.set(qn("w:fldCharType"), "end")
run._r.append(fld1); run._r.append(instr); run._r.append(fld2)
set_cn(run, size=9, color=GRAY)

# ---------- 封面 ----------
para("INSIGHT REPORT", size=10, bold=True, color=CYAN, space_before=40, space_after=4)
para("2026 AI 应用趋势报告", size=30, bold=True, space_before=6, space_after=2, line=1.25)
para("从「能用」到「常用」：智能体、多模态与 AI 原研工程", size=13, color=GRAY, space_after=26)
callout(
    "核心判断",
    "模型能力仍在快速抬升，但真正的分水岭已经出现：谁能把 AI 嵌入业务流程、沉淀为可复用的智能体资产，"
    "谁就能把技术红利转化为持续的经营效率。2026 年的关键词不是「能力验证」，而是「规模化落地」。",
)
para()
meta = doc.add_table(rows=3, cols=2)
meta.autofit = False
items = [
    ("发布机构", "趋势研究组"),
    ("发布日期", "2026 年 11 月"),
    ("数据口径", "公开研究综合整理与示意性估算"),
]
for i, (k, v) in enumerate(items):
    cell_text(meta.rows[i].cells[0], k, bold=True, size=10, color=GRAY)
    cell_text(meta.rows[i].cells[1], v, size=10)
    for c in meta.rows[i].cells:
        c.width = Cm(8.0)

doc.add_page_break()

# ---------- 目录 ----------
para("目 录", size=16, bold=True, space_before=4, space_after=14)
toc = [
    ("执行摘要", "01"),
    ("市场概览：增速放缓，深度提升", "02"),
    ("五大趋势", "03"),
    ("行业渗透与区域格局", "04"),
    ("落地优先级建议", "05"),
    ("风险与挑战", "06"),
    ("行动路线", "07"),
    ("附录：数据说明", "08"),
]
for name, no in toc:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(5)
    set_cn(p.add_run(f"{no}    "), size=10.5, bold=True, color=VIOLET)
    set_cn(p.add_run(name), size=11)
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "dotted"); bottom.set(qn("w:sz"), "4")
    bottom.set(qn("w:space"), "2"); bottom.set(qn("w:color"), "D5DCE6")
    pbdr.append(bottom); pPr.append(pbdr)

doc.add_page_break()

# ---------- 01 执行摘要 ----------
h1("执行摘要", "01")
para(
    "2023 年是生成式 AI 的能力验证期，2024 年进入试点扩散，2026 年则进入规模化落地。"
    "全球生成式 AI 市场规模预计达到约 3,980 亿美元，2023 年至 2028 年的年复合增速约为 65%；"
    "与此同时，增速逐步回归理性，企业之间的差距从「是否使用 AI」转向「使用深度与组织适配度」。",
    indent=21,
)
para(
    "本报告基于公开研究与企业实践的综合梳理，提出五个值得优先关注的方向：智能体工作流、多模态内容生产、"
    "AI 原生软件工程、端侧小模型与个人硬件、以及企业 AI 治理。三者属于应当在本年度启动规模化投入的 P0 方向，"
    "后两者建议以试点与机制建设为主。",
    indent=21,
)

h2("三个关键结论")
bullet("智能体正在取代对话框，成为企业软件的新交互层与执行层。", "结论一：")
bullet("投入产出能否被度量，决定 AI 项目在第二年能否继续获得预算。", "结论二：")
bullet("治理不是成本项，而是进入大客户供应链的准入条件。", "结论三：")

# ---------- 02 市场概览 ----------
h1("市场概览：增速放缓，深度提升", "02")
para(
    "从规模看，全球生成式 AI 市场仍处于高速扩张区间，但曲线斜率已从 2024 年的爆发式增长转为更可持续的复合增长。"
    "从结构看，收入重心正从「模型调用」向「应用与解决方案」迁移，行业 Know-how 与流程整合能力成为溢价来源。",
    indent=21,
)
make_table(
    ["年份", "市场规模（亿美元）", "同比增速", "阶段特征"],
    [
        ["2022", "400", "—", "技术验证"],
        ["2023", "670", "+68%", "能力竞赛"],
        ["2024", "1,370", "+105%", "试点扩散"],
        ["2025", "2,420", "+77%", "场景收敛"],
        ["2026", "3,980", "+64%", "规模落地"],
        ["2027（预测）", "5,960", "+50%", "深度整合"],
        ["2028（预测）", "8,220", "+38%", "常态渗透"],
    ],
    widths=[3.2, 4.6, 3.2, 4.4],
)
para("数据来源：公开研究综合整理，2026 年及以后为示意性预测，仅用于趋势判断。", size=9, color=GRAY)

# ---------- 03 五大趋势 ----------
h1("五大趋势", "03")

h2("趋势一 · 智能体工作流成为新的软件形态")
para(
    "单一对话框正在被「可编排、可追踪、可回滚」的智能体工作流取代。AI 不再只回答问题，而是调用工具、读写系统、"
    "执行多步骤任务，并对结果负责。客服、财务对账、供应链调度等高重复、规则相对清晰的场景率先跑通 ROI。",
    indent=21,
)
bullet("从「逐步人工确认」转向「人在环上」的例外管理。")
bullet("评估集、运行监控与回滚能力成为选型的硬性门槛。")
bullet("头部企业单任务处理成本下降约 40%，流程周期缩短 3 倍以上。")

h2("趋势二 · 多模态内容生产进入工业化")
para(
    "视频、语音、图像生成成本持续下探，营销素材、培训课件、电商详情页开始由「拍一次」转向「按需生成、批量迭代」。"
    "内容团队的组织方式随之重构：从拍摄与剪辑，转向风格管理与质量把关。",
    indent=21,
)
bullet("短视频与配音生成成本较 2023 年下降一个数量级。")
bullet("品牌一致性控制（风格锚点、版权溯源）成为新刚需。")
bullet("内容审核与合规链路前置到生成环节，而非事后补救。")

h2("趋势三 · AI 原生软件工程重塑研发流程")
para(
    "代码补全只是起点。需求拆解、测试生成、代码评审、线上问题定位正在被端到端接管。"
    "在头部公司，研发团队的人均交付量已出现数十个百分点的提升，遗留系统迁移成为确定性最高的价值场景。",
    indent=21,
)
bullet("「规格即代码」压缩需求文档与实现之间的中间层。")
bullet("测试与评审由 AI 承担第一轮，人工聚焦架构设计与风险判断。")
bullet("回归测试耗时平均下降约 50%。")

h2("趋势四 · 端侧小模型与个人 AI 硬件起量")
para(
    "蒸馏与量化让小模型在手机、PC、可穿戴设备上本地运行。隐私、时延与离线可用性，使端侧成为云端推理之外的第二条主线；"
    "AI 眼镜、吊坠类设备开始探索「常驻感知」的交互形态。",
    indent=21,
)
bullet("端侧推理边际成本趋近于零，隐私敏感场景优先迁移。")
bullet("端云协同的调度策略决定用户体验上限。")
bullet("硬件形态尚未收敛，生态与续航仍是主要瓶颈。")

h2("趋势五 · AI 治理从合规成本变为信任资产")
para(
    "监管框架逐步落地，企业开始建立模型清单、数据血缘、输出审计与红队测试机制。治理能力不再只是法务议题，"
    "而是进入采购与合作的评估清单，直接影响大客户中标率。",
    indent=21,
)
bullet("可解释性记录与人工复核留痕成为审计刚需。")
bullet("供应商安全问卷从「有无机制」转向「证据强度」。")
bullet("治理成熟度与客户信任度呈明显正相关。")

# ---------- 04 行业渗透 ----------
h1("行业渗透与区域格局", "04")
para(
    "行业侧的差异主要来自数据基础与流程标准化程度。软件与互联网仍是渗透率最高的行业，"
    "金融与专业服务紧随其后；制造、医疗等行业的渗透受制于数据分散与合规约束，但增量空间更大。",
    indent=21,
)
make_table(
    ["行业", "2026 企业采用率", "典型场景", "主要障碍"],
    [
        ["软件与互联网", "78%", "研发提效、内容生成", "质量与安全"],
        ["金融", "62%", "风控、投研、客服", "合规与审计"],
        ["专业服务", "58%", "文档、检索、交付", "知识资产化"],
        ["零售电商", "52%", "素材生产、推荐、客服", "品牌一致性"],
        ["制造", "47%", "质检、排产、知识库", "数据分散"],
        ["教育", "44%", "备课、个性化辅导", "内容准确性"],
        ["医疗健康", "39%", "辅助诊断、病历整理", "数据与责任"],
        ["政务", "35%", "政务问答、材料审核", "安全与可控"],
    ],
    widths=[3.4, 4.0, 5.0, 3.6],
)
h2("区域对比")
make_table(
    ["区域", "企业采用率", "2026 年投入（亿美元）", "同比增速"],
    [
        ["北美", "71%", "1,860", "38%"],
        ["中国", "58%", "980", "46%"],
        ["欧洲", "54%", "720", "33%"],
        ["亚太其他", "49%", "520", "41%"],
        ["拉美", "41%", "180", "35%"],
        ["中东非", "33%", "140", "39%"],
    ],
    widths=[4.0, 4.0, 5.2, 3.6],
)

# ---------- 05 优先级 ----------
h1("落地优先级建议", "05")
para(
    "按「技术成熟度 × 业务价值」两个维度评估，P0 表示建议在本年度内启动规模化投入，P1 以试点验证为主，P2 保持技术跟踪。",
    indent=21,
)
make_table(
    ["趋势方向", "成熟度", "业务价值", "优先级", "建议动作"],
    [
        ["智能体工作流", "3.8 / 5", "4.7 / 5", "P0", "选定 1–2 个高频流程端到端改造"],
        ["多模态内容生成", "4.2 / 5", "4.3 / 5", "P0", "建立品牌风格库与批量生产管线"],
        ["AI 原生软件工程", "3.6 / 5", "4.5 / 5", "P0", "在测试与评审环节先行落地"],
        ["企业 AI 治理", "3.1 / 5", "4.0 / 5", "P1", "建立模型清单与审计留痕机制"],
        ["端侧小模型与硬件", "2.9 / 5", "3.8 / 5", "P1", "评估隐私敏感场景本地化方案"],
        ["AI 科学发现", "2.4 / 5", "3.9 / 5", "P2", "与科研机构合作试点"],
        ["具身智能与机器人", "2.1 / 5", "3.4 / 5", "P2", "保持技术跟踪，暂不重投入"],
    ],
    widths=[3.8, 2.6, 2.6, 2.2, 5.4],
)

# ---------- 06 风险 ----------
h1("风险与挑战", "06")
h2("数据质量与可用性")
para("AI 的效果上限由数据决定。分散在多个系统、口径不一致的数据，会让智能体在真实业务中频繁失效。", indent=21)
h2("投入产出难以量化")
para("缺少基线指标与归因方法，导致项目难以证明价值，预算往往在第二年收缩。", indent=21)
h2("组织与流程适配不足")
para("把 AI 当作插件而非重构流程，人机分工不清，效率提升被流程摩擦抵消。", indent=21)
h2("合规与安全边界")
para("数据出境、生成内容版权、模型输出责任划分，仍需在项目立项阶段前置设计。", indent=21)

# ---------- 07 行动 ----------
h1("行动路线", "07")
para("建议以「单点闭环—能力沉淀—治理规模化」三步推进，避免一次性大规模铺开。", indent=21)
h2("第一步：选对第一个场景")
para("优先选择高频、规则清晰、可度量的流程，用 8–12 周完成闭环，建立内部信心与基线数据。", indent=21)
h2("第二步：沉淀可复用资产")
para("把提示词、工具接口、评估集、审核规则沉淀为平台能力，避免每个项目重复造轮子。", indent=21)
h2("第三步：治理与规模化同步")
para("在扩展阶段同步建立模型清单、审计日志与红队机制，让治理能力成为对外合作的信任背书。", indent=21)

# ---------- 08 附录 ----------
h1("附录：数据说明", "08")
para(
    "本报告中的市场规模、增长率、采用率与投入产出数据，系基于公开研究资料的综合整理与示意性估算，"
    "用于呈现趋势方向与相对关系，不代表任何单一机构的口径，亦不构成投资建议。"
    "企业在做具体决策时，应结合自身行业数据与内部基线进行校准。",
    indent=21, size=9.5, color=GRAY,
)

os.makedirs(os.path.dirname(OUT), exist_ok=True)
doc.save(OUT)
print("saved:", OUT, os.path.getsize(OUT))
