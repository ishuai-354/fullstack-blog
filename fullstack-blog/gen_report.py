# -*- coding: utf-8 -*-
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

ROOT = r"D:\deepseekagent\dsh\fullstack-blog"
SHOTS = os.path.join(ROOT, "shots")
OUT = os.path.join(ROOT, "全栈应用开发实践报告_王佳骏.docx")

doc = Document()

# 基础样式：正文宋体小四，标题黑体
style = doc.styles["Normal"]
style.font.name = "宋体"
style.font.size = Pt(12)
style.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
style.paragraph_format.line_spacing = 1.5

for name, size in (("Heading 1", 16), ("Heading 2", 14), ("Heading 3", 12.5)):
    st = doc.styles[name]
    st.font.name = "黑体"
    st.font.size = Pt(size)
    st.font.bold = True
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")

for section in doc.sections:
    section.top_margin = Cm(2.4)
    section.bottom_margin = Cm(2.4)
    section.left_margin = Cm(2.6)
    section.right_margin = Cm(2.6)

def para(text, size=12, bold=False, align=None, color=None, italic=False, space_after=6):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    if color:
        run.font.color.rgb = color
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    return p

def caption(text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0x60, 0x60, 0x60)
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(12)

def figure(path, cap, width=15.5):
    if os.path.exists(path):
        doc.add_picture(path, width=Cm(width))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        caption(cap)
    else:
        caption("【" + cap + "：截图缺失】")

def placeholder(cap):
    p = doc.add_paragraph()
    run = p.add_run("【此处插入截图：" + cap + "】")
    run.font.size = Pt(11)
    run.italic = True
    run.font.color.rgb = RGBColor(0x99, 0x33, 0x33)
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

def table(headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ""
        p = hdr[i].paragraphs[0]
        run = p.add_run(h)
        run.bold = True
        run.font.size = Pt(10.5)
        run.font.name = "宋体"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in rows:
        cells = t.add_row().cells
        for i, v in enumerate(r):
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            run = p.add_run(str(v))
            run.font.size = Pt(10.5)
            run.font.name = "宋体"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    return t

# ===== 封面 =====
for _ in range(3):
    doc.add_paragraph()
para("全栈应用开发实践报告", size=26, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
para("——个人作品集博客系统", size=16, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=40)
for _ in range(4):
    doc.add_paragraph()
for label, value in (("姓　　名", "王佳骏"), ("课程名称", "数字化技术基础及应用"), ("项目类型", "个人独立完成"), ("完成日期", "2026 年 11 月 1 日")):
    p = doc.add_paragraph()
    run = p.add_run(f"{label}：{value}")
    run.font.size = Pt(13)
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(10)
doc.add_page_break()

# ===== 1 项目概述 =====
doc.add_heading("一、项目概述", level=1)
para("本次实践依据《Vibe Coding 全栈应用开发》文档中“个人博客系统”的示例思路，结合个人已有的作品集网站，完成了一个"
     "主题为「个人作品集博客」的全栈应用。项目完整实现了前端页面、后端接口与数据库存储三层结构，覆盖用户认证、"
     "文章增删改查、分页、搜索、分类筛选、浏览量统计等核心功能。")
para("选题理由：文档提供了博客、问答社区、在线商城三个示例。博客系统涉及 CRUD 与用户认证，是理解全栈链路最典型的场景；"
     "同时我此前已用 React + Vite 完成过一个静态作品集网站（已部署至 Netlify），把博客主题定为「作品集博客」，"
     "可以让两个项目形成呼应，也符合任务要求的“内容融入自己的设计”。")

doc.add_heading("二、技术选型", level=1)
para("文档推荐的技术栈为 React + TypeScript + Vite + Tailwind CSS 搭配 Node.js + Express 与 MySQL。结合本次实践的时间"
     "与本地环境（本机未安装 MySQL 与 Docker），我在保证“前端 + 后端 + 数据库”三层结构不变的前提下做了两处等价替换，"
     "并在下表说明取舍。")
caption("表 1　技术选型对照")
table(
    ["层次", "文档推荐", "本项目实现", "说明"],
    [
        ["前端", "React + TS + Vite + Tailwind", "原生 HTML + CSS + JavaScript", "省去构建链，零依赖启动；交互与状态管理逻辑等价"],
        ["后端", "Node.js + Express", "Node.js 内置 http 模块", "路由与中间件手写，原理与 Express 一致"],
        ["数据库", "MySQL", "SQLite（Node 24 内置 node:sqlite）", "本机零安装；SQL 语法与表设计可直接迁移到 MySQL"],
        ["认证", "JWT", "管理口令 + 请求头鉴权", "保留“凭据校验后才可写”的鉴权模型，实现更轻量"],
    ],
    widths=[2.2, 4.6, 5.2, 5.0],
)
para("说明：替换的核心原因是环境成本。MySQL 需要额外安装服务、配置账号与端口，在有限的实践时间内风险较高；"
     "而 Node.js 24 已内置 SQLite 支持，一条 node server.js 命令即可同时启动服务与数据库，便于快速验证全栈链路。"
     "数据库表结构、REST 接口设计、鉴权流程均按文档描述实现，后续若要切换为 MySQL + Express + React，只需替换对应层，"
     "业务逻辑无需改动。")

# ===== 3 系统设计 =====
doc.add_heading("三、系统设计", level=1)
doc.add_heading("3.1 数据库设计", level=2)
para("项目采用单表结构存储文章，字段设计参考文档中的文章表，并补充标签与浏览量：")
caption("表 2　posts 表结构")
table(
    ["字段", "类型", "说明"],
    [
        ["id", "INTEGER", "主键，自增"],
        ["title", "TEXT", "文章标题"],
        ["content", "TEXT", "正文内容"],
        ["category", "TEXT", "分类（项目复盘 / 学习笔记 / 前端等）"],
        ["tags", "TEXT", "标签，逗号分隔"],
        ["views", "INTEGER", "浏览次数，进入详情页自动 +1"],
        ["created_at", "TEXT", "创建时间"],
        ["updated_at", "TEXT", "最后更新时间"],
    ],
    widths=[3.2, 3.0, 9.5],
)

doc.add_heading("3.2 接口设计", level=2)
para("后端共提供 6 组 REST 接口，统一返回 JSON：")
caption("表 3　接口清单")
table(
    ["方法", "路径", "功能", "鉴权"],
    [
        ["GET", "/api/posts", "文章列表，支持 keyword / category / page / size", "否"],
        ["GET", "/api/posts/:id", "文章详情，浏览量 +1", "否"],
        ["POST", "/api/posts", "发布文章", "是"],
        ["PUT", "/api/posts/:id", "编辑指定文章", "是"],
        ["DELETE", "/api/posts/:id", "删除指定文章", "是"],
        ["GET", "/api/stats", "统计：文章数、总浏览、分类数", "否"],
    ],
    widths=[2.0, 4.6, 7.6, 1.8],
)

doc.add_heading("3.3 前端页面", level=2)
para("前端为单页应用，通过 fetch 调用后端接口，主要包含：首页（统计卡片 + 搜索栏 + 分类标签 + 文章卡片列表 + 分页）、"
     "文章详情弹窗、写文章 / 编辑弹窗。所有页面均为响应式布局，配色采用深色主题，与个人作品集网站风格保持一致。")

# ===== 4 功能演示 =====
doc.add_heading("四、功能演示", level=1)
para("以下为系统运行时的实际界面截图。")

doc.add_heading("4.1 首页与文章列表", level=2)
figure(os.path.join(SHOTS, "fig1_home.png"), "图 1　首页：统计信息、搜索栏、分类标签与文章卡片列表")
para("首页顶部展示文章总数、总浏览量、分类数三项统计；中部为搜索框与分类下拉；下方为文章卡片列表，"
     "每张卡片显示分类、标题、摘要、发布日期与浏览量；底部为分页控件。数据全部由后端从 SQLite 读取。")

doc.add_heading("4.2 搜索与筛选", level=2)
figure(os.path.join(SHOTS, "fig2_search.png"), "图 2　输入关键词「第一」后的搜索结果")
para("搜索采用输入即搜（防抖 250 毫秒），后端对标题与正文同时使用 LIKE 匹配；分类标签与下拉框联动，"
     "筛选后分页重置为第 1 页。图中搜索“第一”后仅返回命中的《Vibe Coding 的第一周》一篇。")

doc.add_heading("4.3 文章详情", level=2)
figure(os.path.join(SHOTS, "fig3_detail.png"), "图 3　文章详情弹窗（含分类、时间、浏览量、标签）")
para("点击任意卡片打开详情弹窗，展示完整正文、分类、发布时间、标签与浏览量。打开详情时后端将 views 字段 +1，"
     "图中浏览量已由 128 增至 129，说明读写链路正常。")

doc.add_heading("4.4 发布与编辑", level=2)
placeholder("写文章弹窗：填写标题、分类、标签、正文与口令")
para("点击“写文章”打开编辑弹窗，填写标题、分类、标签、正文，并输入管理口令后提交。后端校验标题不少于 2 字、"
     "正文不少于 5 字，口令错误返回 401。")

doc.add_heading("4.5 服务启动与项目结构", level=2)
placeholder("命令行启动日志：node server.js")
placeholder("项目目录结构：server.js / public / blog.db")
para("项目零第三方依赖，执行 node server.js 后监听 127.0.0.1:5178，控制台输出访问地址与口令；"
     "目录中 blog.db 为 SQLite 数据库文件，可直接用任意 SQLite 工具打开查看数据。")

# ===== 5 开发过程 =====
doc.add_heading("五、开发过程与 Vibe Coding 实践", level=1)
para("整个开发过程遵循文档提出的 Vibe Coding 思路：先用自然语言把“要做什么”描述清楚，再让 AI 生成代码，人工负责"
     "校验与联调。具体步骤如下：")
for i, step in enumerate([
    "需求梳理：从文档中提取博客系统的核心功能，确定本期范围（文章 CRUD、搜索、分页、分类、浏览量、口令鉴权），"
    "明确本期不做用户注册体系与富文本编辑器，以控制工期。",
    "数据库与接口设计：先确定 posts 表字段与 REST 接口清单，再动手写代码，避免边写边改表结构。",
    "后端实现：基于 Node 内置 http 模块实现路由分发、JSON 请求体解析、鉴权中间件与 SQL 读写，全部使用参数化查询防止注入。",
    "前端实现：单页结构，使用 fetch 调用接口，配合模板字符串渲染卡片，加入防抖搜索与弹窗交互。",
    "联调与验证：用命令行逐项验证接口（列表、搜索、创建、无口令 401、删除、统计），再在浏览器中验证交互与数据一致性。",
], 1):
    p = doc.add_paragraph(style="List Number")
    run = p.add_run(step)
    run.font.size = Pt(12)
    run.font.name = "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

para("体会上，Vibe Coding 的效率提升主要来自两点：一是把需求写成结构化的提示词（说明接口、字段、校验规则），"
     "生成结果更接近可用状态；二是把验证环节交给可执行的测试命令，而不是靠肉眼读代码。",
     space_after=10)

# ===== 6 问题与解决 =====
doc.add_heading("六、遇到的问题与解决", level=1)
caption("表 4　典型问题与处理")
table(
    ["问题", "原因", "解决方式"],
    [
        ["本机无 MySQL 与 Docker", "环境不满足文档推荐栈", "改用 Node 内置 SQLite，保持表结构与 SQL 通用"],
        ["原生 Node 无路由与解析", "没有使用 Express", "手写 URL 解析、JSON 请求体读取与 404/401 响应"],
        ["中文搜索无结果", "关键词未编码直接拼接 URL", "前端使用 URLSearchParams 统一编码参数"],
        ["直接写接口易被误操作", "缺少鉴权", "写操作统一校验请求头口令，未通过返回 401"],
        ["分页与筛选状态错乱", "切换分类后仍停留在旧页码", "筛选变化时统一将 page 重置为 1"],
    ],
    widths=[4.2, 4.6, 7.0],
)

# ===== 7 总结 =====
doc.add_heading("七、总结与后续扩展", level=1)
para("通过本次实践，我完整走通了“前端页面 → 后端接口 → 数据库”的全栈链路，理解了 REST 接口设计、参数校验、"
     "鉴权与分页搜索的实现方式，也体会到 Vibe Coding 下“先设计、再生成、后验证”的工作方法。")
para("后续可扩展的方向包括：接入 React + Express + MySQL 完整还原文档技术栈；增加用户注册与 JWT 登录；"
     "支持 Markdown 渲染与封面上传；增加标签云与文章置顶；将项目部署至云服务器或 Serverless 平台，"
     "与已有的作品集网站统一域名。")

doc.save(OUT)
print("saved:", OUT)
print("size:", os.path.getsize(OUT))