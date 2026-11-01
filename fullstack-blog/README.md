# 作品集博客（全栈练习）

一个零依赖的全栈小项目：Node 内置 http 提供接口，Node 24 内置 SQLite 存数据，前端是单页 HTML。

## 启动

    node server.js

打开 http://127.0.0.1:5178 ，管理口令 shuai2026。

## 接口

- GET  /api/posts        列表（keyword / category / page / size）
- GET  /api/posts/:id    详情（浏览量 +1）
- POST /api/posts        新建（需 x-token）
- PUT  /api/posts/:id    修改（需 x-token）
- DELETE /api/posts/:id  删除（需 x-token）
- GET  /api/stats        统计
