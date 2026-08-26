/**
 * 快捷指令 — 本地 HTTPS 静态服务器（仅使用 Node.js 标准库，无需安装任何依赖）。
 *
 * Office 要求加载项内容通过 HTTPS 提供（即使是 localhost），且证书必须被系统信任。
 * 启动前请在项目根目录准备好 localhost.crt 与 localhost.key（见 README）。
 *
 * 用法：  node server.js
 * 默认端口 3000，可用环境变量 PORT 覆盖。
 */
const https = require("https");
const fs = require("fs");
const path = require("path");

const ROOT = __dirname;
const PORT = process.env.PORT || 3000;
const CERT = path.join(ROOT, "localhost.crt");
const KEY = path.join(ROOT, "localhost.key");

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".ico": "image/x-icon",
  ".svg": "image/svg+xml",
};

if (!fs.existsSync(CERT) || !fs.existsSync(KEY)) {
  console.error(
    "缺少证书文件：需要 localhost.crt 与 localhost.key。\n" +
      "请先运行： npx --yes office-addin-dev-certs install\n" +
      "（或见 README 中的 openssl 备选方案）"
  );
  process.exit(1);
}

const options = {
  cert: fs.readFileSync(CERT),
  key: fs.readFileSync(KEY),
};

const server = https.createServer(options, (req, res) => {
  // 去掉查询字符串并解码
  let urlPath = decodeURIComponent(req.url.split("?")[0]);
  if (urlPath === "/") urlPath = "/src/taskpane.html";

  // 规整并防止路径穿越
  const safePath = path
    .normalize(urlPath)
    .replace(/^(\.\.[/\\])+/, "")
    .replace(/^[/\\]+/, "");
  const filePath = path.join(ROOT, safePath);

  if (!filePath.startsWith(ROOT)) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      res.end("404 Not Found: " + urlPath);
      return;
    }
    const ext = path.extname(filePath).toLowerCase();
    res.writeHead(200, {
      "Content-Type": MIME[ext] || "application/octet-stream",
      "Access-Control-Allow-Origin": "*",
      "Cache-Control": "no-cache",
    });
    res.end(data);
  });
});

server.listen(PORT, () => {
  console.log(`快捷指令开发服务器已启动: https://localhost:${PORT}`);
  console.log(`任务窗格: https://localhost:${PORT}/src/taskpane.html`);
});
