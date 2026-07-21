from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


OUTPUT = "/workspace/储能光伏产品认证工程师简历.docx"
NAVY = "17365D"
TEAL = "0F766E"
GRAY = "5B6573"
LIGHT = "EAF1F8"


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill)
    tc_pr.append(shading)


def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right"):
        if edge in kwargs:
            tag = "w:{}".format(edge)
            element = borders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                borders.append(element)
            for key, value in kwargs[edge].items():
                element.set(qn("w:{}".format(key)), str(value))


def font(run, size=10.2, bold=False, color=None, name="等线"):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def add_text(paragraph, text, size=10.2, bold=False, color=None, name="等线"):
    run = paragraph.add_run(text)
    font(run, size, bold, color, name)
    return run


def add_bullet(doc, text):
    p = doc.add_paragraph(style="Resume Bullet")
    p.paragraph_format.space_after = Pt(2.5)
    p.paragraph_format.line_spacing = 1.16
    add_text(p, text)
    return p


def section_title(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(9)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    p_pr = p._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "14")
    bottom.set(qn("w:space"), "3")
    bottom.set(qn("w:color"), TEAL)
    borders.append(bottom)
    p_pr.append(borders)
    add_text(p, text, 13, True, NAVY)
    return p


def job_heading(doc, period, company, title):
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    table.columns[0].width = Cm(3.25)
    table.columns[1].width = Cm(13.75)
    row = table.rows[0]
    row.cells[0].width = Cm(3.25)
    row.cells[1].width = Cm(13.75)
    for cell in row.cells:
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_border(cell, bottom={"val": "single", "sz": "4", "color": "D7E3EF"})
    set_cell_shading(row.cells[0], LIGHT)
    p = row.cells[0].paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    add_text(p, period, 9.2, True, NAVY)
    p = row.cells[1].paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    add_text(p, company, 11, True, NAVY)
    add_text(p, "  |  " + title, 10.2, True, GRAY)


def add_publication(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.2)
    p.paragraph_format.first_line_indent = Cm(-0.2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.1
    add_text(p, text, 9.2)


doc = Document()
section = doc.sections[0]
section.top_margin = Cm(1.35)
section.bottom_margin = Cm(1.35)
section.left_margin = Cm(1.55)
section.right_margin = Cm(1.55)

styles = doc.styles
normal = styles["Normal"]
normal.font.name = "等线"
normal._element.rPr.rFonts.set(qn("w:eastAsia"), "等线")
normal.font.size = Pt(10.2)

bullet = styles.add_style("Resume Bullet", WD_STYLE_TYPE.PARAGRAPH)
bullet.base_style = styles["Normal"]
bullet.paragraph_format.left_indent = Cm(0.42)
bullet.paragraph_format.first_line_indent = Cm(-0.42)
bullet.paragraph_format.space_after = Pt(2)
bullet.paragraph_format.line_spacing = 1.15
bullet._element.get_or_add_pPr().insert(
    0, OxmlElement("w:numPr")
)
num_pr = bullet._element.pPr.numPr
ilvl = OxmlElement("w:ilvl")
ilvl.set(qn("w:val"), "0")
num_id = OxmlElement("w:numId")
num_id.set(qn("w:val"), "1")
num_pr.append(ilvl)
num_pr.append(num_id)

# Header
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(3)
add_text(p, "个人简历", 21, True, NAVY, "微软雅黑")
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(9)
add_text(p, "产品认证工程师  |  光伏储能合规  |  国际技术支持", 10.5, True, TEAL)

contact = doc.add_table(rows=1, cols=3)
contact.autofit = False
for i, label in enumerate(("姓名：________", "手机：________", "邮箱：________")):
    cell = contact.rows[0].cells[i]
    set_cell_shading(cell, LIGHT)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.space_before = Pt(3)
    add_text(p, label, 9.5, False, GRAY)

section_title(doc, "个人优势与总结")
summary = (
    "环境工程本科、硕士学历，具备光伏、储能及海外技术服务相关经验。英语可作为工作语言，"
    "长期对接欧洲客户及海外第三方运维服务商，熟悉跨文化商务沟通、技术谈判与外贸项目协同。"
    "能够将 AI 大模型作为辅助工具，提升法规研读、客诉数据整理及中英文技术文档产出的效率。"
)
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(3)
p.paragraph_format.line_spacing = 1.16
add_text(p, summary)
add_bullet(doc, "合规文件处理：辅助检索、翻译及对比各国并网标准与认证法规，提炼合规差异点并缩短标准研读周期。")
add_bullet(doc, "客诉数据研判：辅助拆解历史海外客诉数据，识别失效共性问题，形成趋势分析和改善建议。")
add_bullet(doc, "双语文档输出：辅助起草和校对中英文售后报告、客诉偏差应答、技术标书及合规说明，提升海外业务响应效率。")

section_title(doc, "工作经历")
job_heading(doc, "2023.04 – 至今", "无锡天青元储智能科技有限公司", "项目支持部｜产品认证工程师")
for item in [
    "AI 辅助合规研判：借助大语言模型辅助完成光伏、储能法规的翻译与条款拆解，梳理标准更新差异；搭建并持续迭代逆变器全球并网参数库，适配不同国家市场准入要求。",
    "第三方机构资源管理：开发、筛选莱茵、南德、SGS、KIWA、BV 等国内外认证检测机构，搭建标准化供应商资源库并维护长期合作渠道。",
    "认证项目全流程管控：负责逆变器、储能电池产品认证方案招标议价、合同拟定审核、测试进度跟进、付款对账及票据归档；测算认证预算并优化测试方案，控制认证成本。",
    "认证实操落地：统筹样机国内外寄送、认证资料准备、实验室测试问题整改、证书及报告审核归档。",
    "多国准入合规：熟悉逆变器安规、EMC、EN 50549、VDE 4105/4110、TOR、NRS 等并网要求；熟悉 UN38.3、IEC 62619/63056/62477/62040、UL 9540A、EU 2023/1542 等储能认证规范，以及配套设备的安规、EMC、RF、RoHS、REACH、EN 18031/CRA 网络安全要求。",
    "OEM/ODM 客户支持：统筹逆变器、储能电池派生副证申请，梳理产品差异、优化测试及列名注册流程，输出标准化 SOP，加速贴牌产品上市。",
    "多项目统筹：对接销售需求，动态调整认证优先级，协调研发和实验室资源，并行推进多个跨国认证项目，保障新品上市节奏。",
    "证书与法规管理：建立证书报告全生命周期台账，持续跟踪全球标准、法规迭代，推动产品持续合规。",
]:
    add_bullet(doc, item)

job_heading(doc, "2017.04 – 2023.04", "无锡尚德太阳能电力有限公司", "国际客服部｜高级工程师")
for item in [
    "海外售后支持：利用 AI 工具辅助处理海外业务邮件及历史客诉数据；独立处理 1,000+ 起欧洲客户组件投诉，客户满意度达 99%。",
    "失效分析与改善：负责组件失效判定、根因分析及制造类客诉改善推进，协调工厂端落实改善方案。",
    "客诉成本追偿：确认产品失效责任，推动向代工厂、原材料供应商、物流商及保险公司进行客损费用转嫁。",
    "海外运维体系建设：开拓欧洲本地合作资源，落地 25 家运维服务商，搭建欧洲售后运维网络。",
    "客户与备件业务管理：建立大客户服务档案并统筹分销客户技术服务；独立负责组件及维修备件海外销售、备货、订单及报关出口，累计销售额 48 万元、净利润 12 万元。",
    "项目技术支持：编制国内外项目中英文技术标书、技术偏差应答，参与合同评审与商务技术条款谈判，支持销售订单落地。",
    "内部协同与产品资料：将合同特殊质量、BOM、认证及测试要求传导至工厂；维护 PCN、安装手册、运输包装规范、质保文件，参与区域 Roadmap 维护及技术培训。",
    "认证规划与技术答疑：跟踪销售区域认证维护和标准更新，评估特殊安装工况并解答客户技术问题。",
]:
    add_bullet(doc, item)

job_heading(doc, "2016.07 – 2017.04", "无锡尚德太阳能电力有限公司", "EHS 部｜工程师")
for item in [
    "环保监测与排污管理：负责企业排污申报；完成废水（COD、氨氮、总氮、总磷、氟离子、pH、SS）及废气（HF、HCl、NOx、Cl₂、NH₃、TVOC）采样和实验室检测。",
    "环保技改项目：跟进废水、废气处理设施改造，开展现场监测并验证改造达标效果。",
    "安全与合规管理：编制岗位安全中英文培训材料并组织培训；开展安全隐患和环保设施巡检，跟踪整改闭环及职业健康管理。",
    "环评与体系工作：负责改扩建项目环评立项、光伏行业环保规范化申报；维护 EHS 管理体系并配合客户审厂和第三方体系审核。",
]:
    add_bullet(doc, item)

section_title(doc, "教育背景")
education = doc.add_table(rows=2, cols=3)
education.autofit = False
for row, values in zip(education.rows, [
    ("2009.09 – 2013.06", "河海大学", "环境工程｜本科"),
    ("2013.09 – 2016.06", "南京大学", "环境工程｜硕士"),
]):
    for i, value in enumerate(values):
        cell = row.cells[i]
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_border(cell, bottom={"val": "single", "sz": "4", "color": "D7E3EF"})
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.space_before = Pt(2)
        add_text(p, value, 9.8, i == 1, NAVY if i == 1 else None)

section_title(doc, "专业技能")
skills = [
    ("语言能力", "英语可作为工作语言，可独立完成英文技术文档、海外客户沟通及商务技术谈判。"),
    ("AI 工具应用", "ChatGPT、Claude 等大模型辅助法规翻译、标准对比、数据复盘及双语技术文件起草。"),
    ("认证体系", "光伏逆变器、储能电池、充电桩及采集终端的安规、EMC、并网、运输、环保与网络安全认证全流程。"),
    ("业务能力", "跨国认证项目统筹、供应商管理、OEM 派生认证、海外客诉处理、外贸备件销售及商务技术谈判。"),
    ("检测与体系", "液质联用仪、离子色谱、原子吸收等实验室仪器操作维护；EHS、环评申报、排污监测及安全环保合规管理。"),
]
for name, detail in skills:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.12
    add_text(p, name + "：", 10, True, NAVY)
    add_text(p, detail, 10)

section_title(doc, "科研成果")
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(3)
add_text(p, "研究方向：", 10, True, NAVY)
add_text(p, "饮用水中新型污染物的检测鉴定、生成机理与控制技术；新型饮用水处理技术。", 10)
for item in [
    "2015.01 – 2017.02｜国家自然科学基金青年基金项目（51408296，25 万元）｜主要负责人",
    "2014.07 – 2017.06｜江苏省自然科学基金青年基金项目（BK20140607，20 万元）｜主要负责人",
    "研究成果：3 篇 SCI 论文、1 项专利。",
]:
    add_bullet(doc, item)

for item in [
    "Pan, Y.*; Li, W.; An, H.; Cui, H.; Wang, Y. Formation and occurrence of new polar iodinated disinfection byproducts in drinking water. Chemosphere, 2016, 144, 2312–2320.（影响因子：7.1）",
    "Pan, Y.*; Li, W.; Li, A. M.*; Zhou, Q.; Shi, P.; Wang, Y. A new group of disinfection byproducts trihalo-hydroxy-cyclopentene-diones in drinking water. Environmental Science & Technology, 2016, 50(14), 7344–7352.（影响因子：9.03）",
    "Zhang, D.; Wu, Y.; Zhang, X.; Li, W.; Li, Y.; Li, A.; Pan, Y.* Identification, formation and control of polar brominated disinfection byproducts during cooking with edible salt, organic matter and simulated tap water. Water Research, 2020, 172, 115526.（影响因子：11.23）",
    "潘旸，李雯彬，李爱民，周庆，施鹏，王莹，李欢．一种新型消毒副产物三溴代-羟基环戊烯二酮的制备和结构鉴定方法．",
]:
    add_publication(doc, item)

# Footer
footer = section.footer.paragraphs[0]
footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_text(footer, "个人简历  ·  产品认证 / 光伏储能合规 / 国际技术支持", 8.2, False, GRAY)

doc.save(OUTPUT)
print(OUTPUT)
