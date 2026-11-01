# -*- coding: utf-8 -*-
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT = r"D:\deepseekagent\dsh\fullstack-blog"
SHOTS = os.path.join(ROOT, "shots")
OUT = os.path.join(ROOT, "全栈应用开发实践报告_王佳骏.docx")

doc = Document()
style = doc.styles["Normal"]
style.font.name = "宋体"; style.font.size = Pt(12)
style.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
style.paragraph_format.line_spacing = 1.5

for name, size in (("Heading 1", 16), ("Heading 2", 14), ("Heading 3", 12.5)):
    st = doc.styles[name]
    st.font.name = "黑体"; st.font.size = Pt(size); st.font.bold = True
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")

for section in doc.sections:
    section.top_margin = Cm(2.4); section.bottom_margin = Cm(2.4)
    section.left_margin = Cm(2.6); section.right_margin = Cm(2.6)

def _cn(run, font="宋体"):
    run.font.name = font
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font)

def para(text, size=12, bold=False, align=None, color=None, italic=False, after=6):
    p = doc.add_paragraph(); run = p.add_run(text)
    run.font.size = Pt(size); run.bold = bold; run.italic = italic
    _cn(run)
    if color: run.font.color.rgb = color
    if align is not None: p.alignment = align
    p.paragraph_format.space_after = Pt(after)
    return p

def caption(text):
    p = doc.add_paragraph(); run = p.add_run(text)
    run.font.size = Pt(10); run.font.color.rgb = RGBColor(0x60,0x60,0x60); _cn(run)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(12)

def figure(path, cap, width=15.5):
    doc.add_picture(path, width=Cm(width))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption(cap)

def shade(p, color="F2F4F7"):
    pPr = p._p.get_or_add_pPr()
    sh = OxmlElement("w:shd"); sh.set(qn("w:val"), "clear"); sh.set(qn("w:fill"), color)
    pPr.append(sh)

def code(lines):
    for ln in lines:
        p = doc.add_paragraph()
        run = p.add_run(ln if ln else " ")
        run.font.size = Pt(9.5)
        _cn(run, "Consolas")
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.1
        p.paragraph_format.left_indent = Cm(0.6)
        shade(p)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def table(headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]; c.text = ""
        r = c.paragraphs[0].add_run(h); r.bold = True; r.font.size = Pt(10.5); _cn(r)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            r = cells[i].paragraphs[0].add_run(str(v)); r.font.size = Pt(10.5); _cn(r)
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

# 封面
for _ in range(3): doc.add_paragraph()
para("全栈应用开发实践报告", size=26, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=10)
para("——个人作品集博客系统", size=16, align=WD_ALIGN_PARAGRAPH.CENTER, after=40)
for _ in range(4): doc.add_paragraph()
for label, value in (("姓　　名","王佳骏"),("课程名称","数字化技术基础及应用"),("项目类型","个人独立完成"),("完成日期","2026 年 11 月 1 日")):
    p = doc.add_paragraph(); r = p.add_run(f"{label}：{value}")
    r.font.size = Pt(13); _cn(r)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(10)
doc.add_page_break()

# 一
doc.add_heading("一、项目概述", level=1)
para("本次实践依据《Vibe Coding 全栈应用开发》文档中“个人博客系统”的示例思路，结合个人已有的作品集网站，完成了一个主题为"
     "「个人作品集博客」的全栈应用。项目完整实现了前端页面、后端接口与数据库存储三层结构，覆盖鉴权、文章增删改查、"
     "分页、搜索、分类筛选、浏览量统计等核心功能，并全程使用 Git 进行版本管理。")
para("选题理由：文档提供了博客、问答社区、在线商城三个示例。博客系统涉及 CRUD 与用户认证，是理解全栈链路最典型的场景；"
     "同时我此前已用 React + Vite 完成过一个静态作品集网站（已部署至 Netlify），把博客主题定为「作品集博客」，"
     "两个项目可以形成呼应，也符合任务要求的“内容融入自己的设计”。")

doc.add_heading("二、技术选型", level=1)
para("文档推荐 React + TypeScript + Vite + Tailwind CSS 搭配 Node.js + Express 与 MySQL。结合本次实践的时间与本地环境"
     "（本机未安装 MySQL 与 Docker），我在保证“前端 + 后端 + 数据库”三层结构不变的前提下做了两处等价替换：")
caption("表 1　技术选型对照")
table(["层次","文档推荐","本项目实现","说明"], [
    ["前端","React + TS + Vite + Tailwind","原生 HTML + CSS + JavaScript","省去构建链，零依赖启动；交互逻辑等价"],
    ["后端","Node.js + Express","Node.js 内置 http 模块","路由、请求体解析、鉴权中间件手写，原理与 Express 一致"],
    ["数据库","MySQL","SQLite（Node 24 内置 node:sqlite）","本机零安装；表结构与 SQL 可直接迁移"],
    ["认证","JWT","管理口令 + 请求头鉴权","保留“校验通过才可写”的鉴权模型"],
], widths=[2.2,4.4,5.4,5.0])
para("替换的原因是环境成本：MySQL 需安装服务、配置账号与端口，在有限时间内风险较高；Node.js 24 已内置 SQLite，"
     "一条 node server.js 即可同时启动服务与数据库。表结构、REST 接口与鉴权流程均按文档思路实现，后续换回"
     "React + Express + MySQL 只需替换对应层，业务逻辑不变。")

# 三
doc.add_heading("三、系统设计", level=1)
doc.add_heading("3.1 数据库设计", level=2)
para("项目采用单表存储文章，字段参考文档中的文章表，并补充标签与浏览量：")
caption("表 2　posts 表结构")
table(["字段","类型","说明"], [
    ["id","INTEGER","主键，自增"],
    ["title","TEXT","文章标题"],
    ["content","TEXT","正文内容"],
    ["category","TEXT","分类（项目复盘 / 学习笔记 / 前端）"],
    ["tags","TEXT","标签，逗号分隔"],
    ["views","INTEGER","浏览次数，进入详情自动 +1"],
    ["created_at","TEXT","创建时间"],
    ["updated_at","TEXT","最后更新时间"],
], widths=[3.2,3.0,9.5])

doc.add_heading("3.2 接口设计", level=2)
para("后端提供 6 组 REST 接口，统一返回 JSON：")
caption("表 3　接口清单")
table(["方法","路径","功能","鉴权"], [
    ["GET","/api/posts","列表，支持 keyword / category / page / size","否"],
    ["GET","/api/posts/:id","详情，浏览量 +1","否"],
    ["POST","/api/posts","发布文章","是"],
    ["PUT","/api/posts/:id","编辑指定文章","是"],
    ["DELETE","/api/posts/:id","删除指定文章","是"],
    ["GET","/api/stats","统计：文章数、总浏览、分类数","否"],
], widths=[2.0,4.6,7.6,1.8])

doc.add_heading("3.3 前端页面", level=2)
para("前端为单页应用，通过 fetch 调用后端接口，包含：首页（统计卡片 + 搜索栏 + 分类标签 + 文章卡片 + 分页）、"
     "文章详情弹窗、写文章 / 编辑弹窗。深色主题与个人作品集网站保持一致，布局为响应式。")

# 四
doc.add_heading("四、功能演示", level=1)
para("以下为系统运行时的实际界面截图，覆盖列表、搜索、详情、发布与分类筛选五个关键场景。")

doc.add_heading("4.1 首页与文章列表", level=2)
figure(os.path.join(SHOTS,"fig1_home.png"), "图 1　首页：统计信息、搜索栏、分类标签与文章卡片列表")
para("顶部展示文章总数、总浏览量、分类数三项统计；中部为搜索框与分类下拉；下方为文章卡片列表，每张卡片显示分类、标题、"
     "摘要、发布日期与浏览量；底部为分页控件。数据全部由后端从 SQLite 读取。")

doc.add_heading("4.2 搜索与筛选", level=2)
figure(os.path.join(SHOTS,"fig2_search.png"), "图 2　输入关键词「第一」后的搜索结果")
para("搜索为输入即搜（防抖 250 毫秒），后端对标题与正文同时做 LIKE 匹配；分类标签与下拉框联动，筛选后页码重置为第 1 页。"
     "图中搜索“第一”后仅返回命中的《Vibe Coding 的第一周》一篇。")

doc.add_heading("4.3 文章详情", level=2)
figure(os.path.join(SHOTS,"fig3_detail.png"), "图 3　文章详情弹窗（含分类、时间、浏览量、标签）")
para("点击卡片打开详情弹窗，展示完整正文、分类、发布时间、标签与浏览量。打开详情时后端将 views 字段 +1，"
     "图中浏览量已由 128 增至 129，说明读写链路正常。")

doc.add_heading("4.4 发布文章", level=2)
figure(os.path.join(SHOTS,"fig4_editor.png"), "图 4　写文章弹窗：标题、分类、标签、正文与管理口令")
para("点击“写文章”打开编辑弹窗。后端校验标题不少于 2 字、正文不少于 5 字，口令错误返回 401；"
     "提交成功后写入数据库并刷新列表。")

doc.add_heading("4.5 分类筛选", level=2)
figure(os.path.join(SHOTS,"fig5_category.png"), "图 5　按「前端」分类筛选后的列表")
para("分类标签与下拉框共用同一份统计数据，点击任一分类即按 category 条件查询，总数与分页随之更新。")

# 五 版本管理
doc.add_heading("五、Git 版本管理实践", level=1)
para("项目从第一天起就纳入 Git 管理，遵循“小步提交、语义化信息”的规范。完整提交历史如下：")
code([
 "$ git log --oneline --graph --all",
 "*   7da8af5 merge: 合并 README 分支",
 "|\\",
 "| * 14e3fd2 docs: 添加 README 说明接口与启动方式",
 "|/",
 "* 3e3ede0 docs: 补充启动脚本与项目说明",
 "* 39bca16 chore: 加入 SQLite 数据库与示例文章数据",
 "* 0d1ed16 feat: 初始化全栈博客骨架（Node 内置 http + SQLite + 单页前端）",
])
para("实验内容包括：init 初始化仓库、配置 .gitignore 排除 node_modules 与临时文件、按功能拆分提交、"
     "新建 feature/readme 分支开发、使用 --no-ff 合并回主干，以及通过 git status / git diff 检查工作区状态。")

doc.add_heading("5.1 分支与合并", level=2)
para("分支命名为 feature/readme，在其上完成 README 编写并提交，随后切回 master 以 --no-ff 方式合并，"
     "保留分支历史便于回溯。合并后历史上出现清晰的菱形结构。")

# 六 问题
doc.add_heading("六、遇到的问题与解决", level=1)
caption("表 4　典型问题与处理")
table(["问题","原因","解决方式"], [
    ["本机无 MySQL 与 Docker","环境不满足文档推荐栈","改用 Node 内置 SQLite，保持表结构与 SQL 通用"],
    ["原生 Node 无路由与解析","未使用 Express","手写 URL 解析、JSON 请求体读取与 404/401 响应"],
    ["中文搜索无结果","关键词未编码直接拼接 URL","前端使用 URLSearchParams 统一编码参数"],
    ["写接口易被误操作","缺少鉴权","写操作统一校验请求头口令，未通过返回 401"],
    ["切换分类后页码错乱","分页状态未重置","筛选变化时统一将 page 重置为 1"],
], widths=[4.2,4.4,7.0])

# 七
doc.add_heading("七、总结与后续扩展", level=1)
para("通过本次实践，我完整走通了“前端页面 → 后端接口 → 数据库”的全栈链路，理解了 REST 接口设计、参数校验、鉴权、"
     "分页与搜索的实现方式，也熟悉了 Git 分支开发与合并的流程。Vibe Coding 的关键在于先把需求写清楚，再让 AI 生成代码，"
     "最后用可执行的命令逐项验证。")
para("后续可扩展：接入 React + Express + MySQL 完整还原文档技术栈；增加用户注册与 JWT 登录；支持 Markdown 渲染与封面上传；"
     "增加标签云与文章置顶；部署到云服务器并与作品集网站统一域名。")

doc.save(OUT)
print("saved", OUT, os.path.getsize(OUT))