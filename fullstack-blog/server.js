const http = require("node:http");
const fs = require("node:fs");
const path = require("node:path");
const { DatabaseSync } = require("node:sqlite");

const PORT = 5178;
const ADMIN_TOKEN = "shuai2026";
const ROOT = __dirname;
const PUBLIC = path.join(ROOT, "public");

const db = new DatabaseSync(path.join(ROOT, "blog.db"));
db.exec(`
CREATE TABLE IF NOT EXISTS posts (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  title TEXT NOT NULL,
  content TEXT NOT NULL,
  category TEXT NOT NULL DEFAULT '随笔',
  tags TEXT NOT NULL DEFAULT '',
  views INTEGER NOT NULL DEFAULT 0,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);
`);

const now = () => new Date().toISOString().replace("T", " ").slice(0, 19);

const seed = db.prepare("SELECT COUNT(*) AS n FROM posts").get();
if (!seed || seed.n === 0) {
  const ins = db.prepare(
    "INSERT INTO posts (title, content, category, tags, views, created_at, updated_at) VALUES (?,?,?,?,?,?,?)"
  );
  const t = now();
  ins.run("从零搭一个作品集网站", "用 React + Vite + Tailwind 做静态作品集，部署到 Netlify。\n\n关键点：路由、响应式、SEO 描述、构建产物。", "项目复盘", "React,Vite,Netlify", 128, t, t);
  ins.run("Vibe Coding 的第一周", "和 AI 结对写代码的体会：把需求讲清楚比会写语法更重要。\n\n小技巧：先让它写需求文档，再让它写代码。", "学习笔记", "AI,效率", 76, t, t);
  ins.run("CSS 也能很讲究", "同一个页面，间距、字重、留白变一下，观感完全不一样。", "前端", "CSS,设计", 41, t, t);
}

function send(res, code, body, type = "application/json; charset=utf-8") {
  const data = typeof body === "string" || Buffer.isBuffer(body) ? body : JSON.stringify(body);
  res.writeHead(code, { "Content-Type": type, "Cache-Control": "no-store" });
  res.end(data);
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let raw = "";
    req.on("data", (c) => {
      raw += c;
      if (raw.length > 2 * 1024 * 1024) reject(new Error("body too large"));
    });
    req.on("end", () => {
      try { resolve(raw ? JSON.parse(raw) : {}); } catch { reject(new Error("invalid json")); }
    });
    req.on("error", reject);
  });
}

function auth(req) {
  return (req.headers["x-token"] || "") === ADMIN_TOKEN;
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://localhost:${PORT}`);
  const p = url.pathname;

  try {
    if (p === "/api/login" && req.method === "POST") {
      const b = await readBody(req);
      return b.token === ADMIN_TOKEN
        ? send(res, 200, { ok: true })
        : send(res, 401, { ok: false, error: "口令错误" });
    }

    if (p === "/api/posts" && req.method === "GET") {
      const keyword = (url.searchParams.get("keyword") || "").trim();
      const category = (url.searchParams.get("category") || "").trim();
      const page = Math.max(1, parseInt(url.searchParams.get("page") || "1", 10));
      const size = Math.min(20, Math.max(1, parseInt(url.searchParams.get("size") || "6", 10)));
      const where = [];
      const args = [];
      if (keyword) { where.push("(title LIKE ? OR content LIKE ?)"); args.push(`%${keyword}%`, `%${keyword}%`); }
      if (category) { where.push("category = ?"); args.push(category); }
      const whereSql = where.length ? "WHERE " + where.join(" AND ") : "";
      const total = db.prepare(`SELECT COUNT(*) AS n FROM posts ${whereSql}`).get(...args).n;
      const rows = db.prepare(
        `SELECT id,title,category,tags,views,created_at,substr(content,1,80) AS excerpt
         FROM posts ${whereSql} ORDER BY id DESC LIMIT ? OFFSET ?`
      ).all(...args, size, (page - 1) * size);
      const cats = db.prepare("SELECT category, COUNT(*) AS n FROM posts GROUP BY category ORDER BY n DESC").all();
      return send(res, 200, { total, page, size, posts: rows, categories: cats });
    }

    const detail = p.match(/^\/api\/posts\/(\d+)$/);
    if (detail) {
      const id = Number(detail[1]);
      if (req.method === "GET") {
        db.prepare("UPDATE posts SET views = views + 1 WHERE id = ?").run(id);
        const row = db.prepare("SELECT * FROM posts WHERE id = ?").get(id);
        return row ? send(res, 200, row) : send(res, 404, { error: "文章不存在" });
      }
      if (!auth(req)) return send(res, 401, { error: "需要管理口令" });
      if (req.method === "PUT") {
        const b = await readBody(req);
        const old = db.prepare("SELECT * FROM posts WHERE id = ?").get(id);
        if (!old) return send(res, 404, { error: "文章不存在" });
        db.prepare("UPDATE posts SET title=?,content=?,category=?,tags=?,updated_at=? WHERE id=?")
          .run(b.title ?? old.title, b.content ?? old.content, b.category ?? old.category,
               b.tags ?? old.tags, now(), id);
        return send(res, 200, { ok: true });
      }
      if (req.method === "DELETE") {
        db.prepare("DELETE FROM posts WHERE id = ?").run(id);
        return send(res, 200, { ok: true });
      }
    }

    if (p === "/api/posts" && req.method === "POST") {
      if (!auth(req)) return send(res, 401, { error: "需要管理口令" });
      const b = await readBody(req);
      if (!b.title || String(b.title).trim().length < 2) return send(res, 400, { error: "标题至少 2 个字" });
      if (!b.content || String(b.content).trim().length < 5) return send(res, 400, { error: "内容太短" });
      const t = now();
      const info = db.prepare(
        "INSERT INTO posts (title,content,category,tags,views,created_at,updated_at) VALUES (?,?,?,?,0,?,?)"
      ).run(String(b.title).trim(), String(b.content).trim(), (b.category || "随笔").trim(), (b.tags || "").trim(), t, t);
      return send(res, 200, { ok: true, id: Number(info.lastInsertRowid) });
    }

    if (p === "/api/stats") {
      const total = db.prepare("SELECT COUNT(*) AS n FROM posts").get().n;
      const views = db.prepare("SELECT COALESCE(SUM(views),0) AS n FROM posts").get().n;
      const cats = db.prepare("SELECT COUNT(DISTINCT category) AS n FROM posts").get().n;
      return send(res, 200, { total, views, cats });
    }

    // 静态文件
    let file = p === "/" ? "/index.html" : p;
    file = path.join(PUBLIC, path.normalize(file).replace(/^(\.\.[/\\])+/, ""));
    if (!file.startsWith(PUBLIC)) return send(res, 403, "forbidden");
    if (fs.existsSync(file) && fs.statSync(file).isFile()) {
      const ext = path.extname(file);
      const type = { ".html": "text/html; charset=utf-8", ".css": "text/css; charset=utf-8",
                     ".js": "text/javascript; charset=utf-8", ".svg": "image/svg+xml" }[ext] || "application/octet-stream";
      return send(res, 200, fs.readFileSync(file), type);
    }
    return send(res, 404, "Not Found", "text/plain; charset=utf-8");
  } catch (e) {
    return send(res, 500, { error: String(e && e.message || e) });
  }
});

server.listen(PORT, "127.0.0.1", () => {
  console.log(`博客已启动: http://127.0.0.1:${PORT}`);
  console.log(`管理口令: ${ADMIN_TOKEN}`);
});