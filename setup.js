/**
 * 快捷指令 — 一键发布到 GitHub Pages 的配置脚本（仅用 Node.js 标准库）。
 *
 * 作用：把 manifest.xml 中占位的 GitHub Pages 地址
 *   https://YOUR-GITHUB-USERNAME.github.io/office-quick-commands
 * 替换为你的真实地址，并打印后续的上传 / 启用 Pages 步骤。
 *
 * 用法：  node setup.js
 */
const fs = require("fs");
const path = require("path");
const readline = require("readline");

const ROOT = __dirname;
const MANIFEST = path.join(ROOT, "manifest.xml");
const PLACEHOLDER = "YOUR-GITHUB-USERNAME";
const REPO_NAME = "office-quick-commands";

const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

rl.question("请输入你的 GitHub 用户名：", (raw) => {
  const username = (raw || "").trim();
  if (!username) {
    console.error("未输入用户名，已取消。");
    rl.close();
    process.exit(1);
  }

  let s = fs.readFileSync(MANIFEST, "utf8");
  if (!s.includes(PLACEHOLDER)) {
    console.log("提示：manifest.xml 中未找到占位符（可能已配置过，未做修改）。");
  }
  // 替换资源地址（带仓库路径）与 AppDomain（仅域名）
  s = s.split(`https://${PLACEHOLDER}.github.io/${REPO_NAME}`).join(`https://${username}.github.io/${REPO_NAME}`);
  s = s.split(`https://${PLACEHOLDER}.github.io`).join(`https://${username}.github.io`);
  fs.writeFileSync(MANIFEST, s);

  console.log(`\n已更新 manifest.xml，资源地址指向：https://${username}.github.io/${REPO_NAME}/\n`);
  console.log("接下来，在 office-quick-commands 目录下依次执行：\n");
  console.log("  git init");
  console.log("  git add -A");
  console.log('  git commit -m "快捷指令 Office 加载项"');
  console.log("  git branch -M main");
  console.log(`  git remote add origin https://github.com/${username}/${REPO_NAME}.git`);
  console.log("  git push -u origin main");
  console.log("\n然后在 GitHub 仓库页面：Settings → Pages → Source 选择「main / (root)」，保存。");
  console.log("约 1 分钟后，用 Office「上传我的加载项」选择本目录的 manifest.xml 即可（详见 README）。");
  console.log("\n（如需换用本地服务器调试，请使用 manifest-local.xml，并运行 node server.js）");

  rl.close();
});
