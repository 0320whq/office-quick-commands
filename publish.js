#!/usr/bin/env node
/**
 * publish.js — 一键校验并发布「快捷指令」Office 加载项。
 *
 * 用法：
 *   node publish.js           校验通过后自动 git add / commit / push
 *   node publish.js --check  仅做完整性校验（不触碰 git）
 *   node publish.js --dry     校验通过后打印将要执行的 git 命令（不真正推送）
 *   node publish.js --help    显示本帮助
 *
 * 校验项：
 *   1. 三个 JS 文件语法（node --check）
 *   2. shortcuts.json 可解析
 *   3. manifest.xml / manifest-local.xml XML 格式良好（借助 python）
 *   4. 一致性交叉核对：KB / RIBBON 动作 id ↔ shortcuts.json 动作 ↔ 两清单 FunctionName
 *
 * 仅使用 Node 标准库；git 与 python 为可选依赖（缺失时相应步骤降级提示）。
 */

const fs = require("fs");
const path = require("path");
const os = require("os");
const { execSync } = require("child_process");

const ROOT = __dirname;

function section(t) { console.log("\n== " + t + " =="); }
function ok(m) { console.log("  \u2713 " + m); }
function bad(m) { console.log("  \u2717 " + m); }

// --- 校验：JS 语法 ---
function jsSyntax(file) {
  const p = path.join(ROOT, file);
  if (!fs.existsSync(p)) return false;
  try {
    execSync('"' + process.execPath + '" --check "' + p + '"', { stdio: "pipe" });
    return true;
  } catch (e) {
    return false;
  }
}

// --- 校验：XML 格式（借助 python；没有就降级） ---
function xmlWellFormed(file) {
  const p = path.join(ROOT, file);
  if (!fs.existsSync(p)) return false;
  const tmp = path.join(os.tmpdir(), "oqc_xml_check.py");
  fs.writeFileSync(tmp, 'import xml.dom.minidom,sys\ntry:\n  xml.dom.minidom.parse(r"""' + p + '""")\nexcept Exception as e:\n  print(e); sys.exit(1)\n');
  for (const py of ["python", "python3"]) {
    try {
      execSync('"' + py + '" "' + tmp + '"', { stdio: "pipe" });
      return true;
    } catch (e) { /* try next */ }
  }
  return null; // 未知（无 python）
}

// --- 解析 taskpane.js 中的 KB / RIBBON 数组 id ---
function grabArray(name) {
  const js = fs.readFileSync(path.join(ROOT, "src/taskpane.js"), "utf8");
  const s = js.indexOf("const " + name + " = [");
  if (s < 0) return [];
  const o = js.indexOf("[", s);
  let depth = 0, i = o;
  for (; i < js.length; i++) {
    if (js[i] === "[") depth++;
    else if (js[i] === "]") { depth--; if (depth === 0) break; }
  }
  const body = js.slice(o + 1, i);
  const ids = [];
  const re = /\[\s*["']([A-Za-z0-9_]+)["']/g;
  let m;
  while ((m = re.exec(body))) ids.push(m[1]);
  return ids;
}

function functionNames(file) {
  const xml = fs.readFileSync(path.join(ROOT, file), "utf8");
  return [...new Set([...xml.matchAll(/<FunctionName>([^<]+)<\/FunctionName>/g)].map((m) => m[1]))];
}

function main() {
  const args = process.argv.slice(2);
  if (args.includes("--help") || args.includes("-h")) {
    console.log("用法：node publish.js [--check|--dry|--help]");
    process.exit(0);
  }
  const onlyCheck = args.includes("--check");
  const dry = args.includes("--dry");

  let pass = true;
  section("校验加载项（publish.js）");

  // 1. JS 语法
  const jsFiles = ["src/taskpane.js", "setup.js", "server.js"];
  let jsOk = true;
  jsFiles.forEach((f) => { if (!jsSyntax(f)) { jsOk = false; bad("JS 语法错误：" + f); } });
  if (jsOk) ok("JS 语法（" + jsFiles.join(" / ") + "）");
  else pass = false;

  // 2. shortcuts.json
  let sc = null;
  try {
    sc = JSON.parse(fs.readFileSync(path.join(ROOT, "shortcuts.json"), "utf8"));
    ok("shortcuts.json 解析（" + sc.actions.length + " 动作 / " + sc.shortcuts.length + " 快捷键）");
  } catch (e) {
    bad("shortcuts.json 解析失败：" + e.message);
    pass = false;
  }

  // 3. XML
  const xmlFiles = ["manifest.xml", "local-only/manifest-local.xml"];
  let xmlUnknown = false;
  xmlFiles.forEach((f) => {
    const r = xmlWellFormed(f);
    if (r === true) ok(f + " XML 格式良好");
    else if (r === null) { xmlUnknown = true; console.log("  ? " + f + " 跳过 XML 校验（未找到 python）"); }
    else { bad(f + " XML 格式错误"); pass = false; }
  });
  if (xmlUnknown) console.log("  （缺少 python，XML 校验已跳过；可在有 python 的环境重跑）");

  // 4. 一致性交叉核对
  if (sc) {
    const KB = grabArray("KB");
    const RIBBON = grabArray("RIBBON");
    const scIds = sc.actions.map((a) => a.id);
    const manFn = functionNames("manifest.xml");
    const manFnLocal = functionNames("local-only/manifest-local.xml");

    const scNotKb = scIds.filter((x) => !KB.includes(x) && x !== "KbShow" && x !== "KbHide");
    const kbNoSc = KB.filter((x) => !scIds.includes(x));
    const manNotRib = manFn.filter((x) => !RIBBON.includes(x));
    const ribNoBtn = RIBBON.filter((x) => !manFn.includes(x));
    const fnSetsEqual = JSON.stringify(manFn.sort()) === JSON.stringify(manFnLocal.sort());

    let crossOk = true;
    if (scNotKb.length) { crossOk = false; bad("快捷键动作不在 KB 中：" + scNotKb.join(", ")); }
    if (kbNoSc.length) { crossOk = false; bad("KB 动作缺少快捷键：" + kbNoSc.join(", ")); }
    if (manNotRib.length) { crossOk = false; bad("功能区 FunctionName 未注册：" + manNotRib.join(", ")); }
    if (ribNoBtn.length) { crossOk = false; bad("RIBBON 动作缺少功能区按钮：" + ribNoBtn.join(", ")); }
    if (!fnSetsEqual) { crossOk = false; bad("两清单 FunctionName 集合不一致"); }
    if (crossOk) ok("一致性：KB(" + KB.length + ") / RIBBON(" + RIBBON.length + ") / 快捷键(" + scIds.length + ") / 功能区(" + manFn.length + ") 全部对齐");
    else pass = false;
  }

  if (!pass) {
    console.log("\n\u2717 校验未通过，已停止（未推送任何内容）。请修复上方问题后重试。");
    process.exit(1);
  }
  ok("全部校验通过");

  if (onlyCheck) {
    console.log("\n--check 模式：仅校验，未执行 git 操作。");
    process.exit(0);
  }

  // --- git 发布 ---
  section("发布到 git");
  function git(args) {
    try {
      return execSync("git " + args, { cwd: ROOT, stdio: "pipe" }).toString().trim();
    } catch (e) {
      return null;
    }
  }
  if (git("rev-parse --is-inside-work-tree") === null) {
    console.log("  当前不是 git 仓库，跳过推送。请先 `git init` 并配置 origin 远程。");
    process.exit(0);
  }
  const branch = git("rev-parse --abbrev-ref HEAD") || "main";
  const status = git("status --porcelain");
  if (status) {
    git("add -A");
    const msg = "chore: update office-quick-commands (" + new Date().toISOString().slice(0, 10) + ")";
    git('commit -m "' + msg + '"');
    ok("已提交变更：" + msg);
  } else {
    ok("没有需要提交的变更");
  }
  const remote = git("remote get-url origin");
  if (!remote) {
    console.log("  未配置 origin 远程，跳过推送。");
    process.exit(0);
  }
  if (dry) {
    console.log("\n--dry 模式：将执行（未真正推送）：");
    console.log("  git push -u origin " + branch);
    process.exit(0);
  }
  const push = git("push -u origin " + branch);
  if (push === null) {
    bad("推送失败（请检查网络 / 凭据 / 远程配置）");
    process.exit(1);
  }
  ok("已推送至 origin/" + branch);
  console.log("\n完成。若已开启 GitHub Pages，约 1 分钟后即可在 Office 中通过 manifest.xml 使用。");
}

main();
