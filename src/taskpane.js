/* global Office, Excel, Word, document */

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
function nowStamp() {
  const d = new Date();
  const p = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}

function nowDate() {
  const d = new Date();
  const p = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

function host() {
  return Office.context.host;
}

function isExcel() { return host() === Office.HostType.Excel; }
function isWord() { return host() === Office.HostType.Word; }

// ---------------------------------------------------------------------------
// Command implementations
//   Each function is host-aware: it performs the operation in the host(s) that
//   support it and is a quiet no-op elsewhere (the command panel only shows
//   commands relevant to the current host).
// ---------------------------------------------------------------------------

// --- Shared: insert current date/time ---
async function doInsertDate() {
  if (isExcel()) {
    await Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load("rowCount, columnCount");
      await ctx.sync();
      const rows = range.rowCount, cols = range.columnCount;
      const vals = Array.from({ length: rows }, () => Array.from({ length: cols }, () => nowDate()));
      range.values = vals;
      await ctx.sync();
    });
  } else if (isWord()) {
    await Word.run(async (ctx) => {
      ctx.document.getSelection().insertText(nowStamp(), Word.InsertLocation.replace);
      await ctx.sync();
    });
  }
}

// --- Shared: bold toggle ---
async function doBold() {
  if (isWord()) {
    await Word.run(async (ctx) => {
      const range = ctx.document.getSelection();
      range.font.load("bold");
      await ctx.sync();
      range.font.bold = !range.font.bold;
      await ctx.sync();
    });
  } else if (isExcel()) {
    await Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load("format/font/bold");
      await ctx.sync();
      range.format.font.bold = !range.format.font.bold;
      await ctx.sync();
    });
  }
}

// --- Shared: italic toggle ---
async function doItalic() {
  if (isWord()) {
    await Word.run(async (ctx) => {
      const range = ctx.document.getSelection();
      range.font.load("italic");
      await ctx.sync();
      range.font.italic = !range.font.italic;
      await ctx.sync();
    });
  } else if (isExcel()) {
    await Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load("format/font/italic");
      await ctx.sync();
      range.format.font.italic = !range.format.font.italic;
      await ctx.sync();
    });
  }
}

// --- Shared: underline toggle ---
async function doUnderline() {
  if (isWord()) {
    await Word.run(async (ctx) => {
      const range = ctx.document.getSelection();
      range.font.load("underline");
      await ctx.sync();
      range.font.underline =
        (range.font.underline === Word.UnderlineType.none || !range.font.underline)
          ? Word.UnderlineType.single
          : Word.UnderlineType.none;
      await ctx.sync();
    });
  } else if (isExcel()) {
    await Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load("format/font/underline");
      await ctx.sync();
      range.format.font.underline = !range.format.font.underline;
      await ctx.sync();
    });
  }
}

// --- Word: built-in styles ---
async function doHeading1() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    ctx.document.getSelection().styleBuiltIn = Word.BuiltInStyle.heading1;
    await ctx.sync();
  });
}
async function doHeading2() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    ctx.document.getSelection().styleBuiltIn = Word.BuiltInStyle.heading2;
    await ctx.sync();
  });
}
async function doBodyText() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    ctx.document.getSelection().styleBuiltIn = Word.BuiltInStyle.normal;
    await ctx.sync();
  });
}

// --- Word: lists ---
async function doBulletList() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    const paras = ctx.document.getSelection().paragraphs;
    paras.load("list");
    await ctx.sync();
    paras.items.forEach((p) => {
      if (p.list) p.list.delete();
      else p.startNewList(Word.ListType.bulleted);
    });
    await ctx.sync();
  });
}
async function doNumberList() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    const paras = ctx.document.getSelection().paragraphs;
    paras.load("list");
    await ctx.sync();
    paras.items.forEach((p) => {
      if (p.list) p.list.delete();
      else p.startNewList(Word.ListType.numbered);
    });
    await ctx.sync();
  });
}
async function doRemoveList() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    const paras = ctx.document.getSelection().paragraphs;
    paras.load("list");
    await ctx.sync();
    paras.items.forEach((p) => { if (p.list) p.list.delete(); });
    await ctx.sync();
  });
}

// --- Shared: paragraph / cell alignment ---
async function doAlignLeft() {
  if (isWord()) {
    await Word.run(async (ctx) => {
      ctx.document.getSelection().paragraphFormat.alignment = Word.Alignment.left;
      await ctx.sync();
    });
  } else if (isExcel()) {
    await Excel.run(async (ctx) => {
      ctx.workbook.getSelectedRange().format.horizontalAlignment = Excel.HorizontalAlignment.left;
      await ctx.sync();
    });
  }
}
async function doAlignCenter() {
  if (isWord()) {
    await Word.run(async (ctx) => {
      ctx.document.getSelection().paragraphFormat.alignment = Word.Alignment.centered;
      await ctx.sync();
    });
  } else if (isExcel()) {
    await Excel.run(async (ctx) => {
      ctx.workbook.getSelectedRange().format.horizontalAlignment = Excel.HorizontalAlignment.center;
      await ctx.sync();
    });
  }
}
async function doAlignRight() {
  if (isWord()) {
    await Word.run(async (ctx) => {
      ctx.document.getSelection().paragraphFormat.alignment = Word.Alignment.right;
      await ctx.sync();
    });
  } else if (isExcel()) {
    await Excel.run(async (ctx) => {
      ctx.workbook.getSelectedRange().format.horizontalAlignment = Excel.HorizontalAlignment.right;
      await ctx.sync();
    });
  }
}
async function doAlignJustify() {
  if (isWord()) {
    await Word.run(async (ctx) => {
      ctx.document.getSelection().paragraphFormat.alignment = Word.Alignment.justified;
      await ctx.sync();
    });
  } else if (isExcel()) {
    await Excel.run(async (ctx) => {
      ctx.workbook.getSelectedRange().format.horizontalAlignment = Excel.HorizontalAlignment.justify;
      await ctx.sync();
    });
  }
}

// --- Word: line spacing ---
async function doLineSpacing15() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    ctx.document.getSelection().paragraphFormat.lineSpacingRule = Word.LineSpacingRule.onePointFive;
    await ctx.sync();
  });
}
async function doLineSpacingDouble() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    ctx.document.getSelection().paragraphFormat.lineSpacingRule = Word.LineSpacingRule.double;
    await ctx.sync();
  });
}

// --- Word: indents ---
async function doFirstLineIndent() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    ctx.document.getSelection().paragraphFormat.firstLineIndent = 24; // ~2 字符
    await ctx.sync();
  });
}
async function doNoIndent() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    ctx.document.getSelection().paragraphFormat.firstLineIndent = 0;
    await ctx.sync();
  });
}

// --- Word: page break before paragraph (useful for headings) ---
async function doPageBreakBefore() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    ctx.document.getSelection().paragraphFormat.pageBreakBefore = true;
    await ctx.sync();
  });
}

// --- Shared: highlight / clear format ---
async function doHighlight() {
  if (isWord()) {
    await Word.run(async (ctx) => {
      ctx.document.getSelection().font.highlightColor = "Yellow";
      await ctx.sync();
    });
  } else if (isExcel()) {
    await Excel.run(async (ctx) => {
      ctx.workbook.getSelectedRange().format.fill.color = "#FFFF00";
      await ctx.sync();
    });
  }
}
async function doClearFormat() {
  if (isWord()) {
    await Word.run(async (ctx) => {
      const range = ctx.document.getSelection();
      range.font.clear();
      range.paragraphFormat.alignment = Word.Alignment.left;
      await ctx.sync();
    });
  } else if (isExcel()) {
    await Excel.run(async (ctx) => {
      ctx.workbook.getSelectedRange().clear(Excel.ClearApplyments.formats);
      await ctx.sync();
    });
  }
}

// --- Word: case transform ---
async function doUppercase() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    const range = ctx.document.getSelection();
    range.load("text");
    await ctx.sync();
    range.insertText(range.text.toUpperCase(), Word.InsertLocation.replace);
    await ctx.sync();
  });
}
async function doLowercase() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    const range = ctx.document.getSelection();
    range.load("text");
    await ctx.sync();
    range.insertText(range.text.toLowerCase(), Word.InsertLocation.replace);
    await ctx.sync();
  });
}

// --- Word: insert table / page break ---
async function doInsertTable() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    ctx.document.getSelection().insertTable(3, 3, Word.InsertLocation.after);
    await ctx.sync();
  });
}
async function doInsertPageBreak() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    ctx.document.getSelection().insertBreak(Word.BreakType.page, Word.InsertLocation.after);
    await ctx.sync();
  });
}

// --- Word: clear highlight color ---
async function doRemoveHighlight() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    const range = ctx.document.getSelection();
    range.font.load("highlightColor");
    await ctx.sync();
    range.font.highlightColor = "No color";
    await ctx.sync();
  });
}

// --- Word: remove empty paragraphs (selection, else whole document) ---
function getTargetParagraphs(ctx) {
  let paras = ctx.document.getSelection().paragraphs;
  paras.load("items");
  return ctx.sync().then(() => {
    if (paras.items.length === 0) {
      paras = ctx.document.body.paragraphs;
      paras.load("items");
      return ctx.sync().then(() => paras);
    }
    return paras;
  });
}
async function doRemoveEmptyLines() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    const paras = await getTargetParagraphs(ctx);
    paras.items.forEach((p) => p.load("text"));
    await ctx.sync();
    let removed = 0;
    paras.items.forEach((p) => {
      if (p.text.trim() === "") { p.delete(); removed++; }
    });
    await ctx.sync();
    setStatus(`已删除 ${removed} 个空段落`, "ok");
  });
}

// --- Word: collapse multiple spaces into one + trim per paragraph ---
async function doNormalizeSpaces() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    const paras = await getTargetParagraphs(ctx);
    paras.items.forEach((p) => p.load("text"));
    await ctx.sync();
    let changed = 0;
    paras.items.forEach((p) => {
      const norm = p.text.replace(/[ \t　]+/g, " ").replace(/^\s+|\s+$/g, "");
      if (norm !== p.text) {
        const range = p.getRange();
        range.insertText(norm, Word.InsertLocation.replace);
        changed++;
      }
    });
    await ctx.sync();
    setStatus(`已清理 ${changed} 个段落的空格`, "ok");
  });
}

// --- Word: word / character count (selection or whole document) ---
async function doWordCount() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    sel.load("text");
    const body = ctx.document.body;
    body.load("wordCount");
    await ctx.sync();
    const selText = sel.text || "";
    const selChars = selText.replace(/\s/g, "").length;
    const total = body.wordCount;
    setStats(`全文约 ${total} 词 ｜ 选区 ${selChars} 字（${selText.length} 含空格）`);
  });
}

// --- Word: insert a special symbol at the cursor ---
const SYMBOLS = [
  "§", "¶", "©", "®", "™", "°", "×", "÷", "±", "≠",
  "≤", "≥", "≈", "∞", "∑", "√", "π", "α", "β", "μ",
  "€", "£", "¥", "¢", "•", "…", "→", "←", "↑", "↓",
  "◆", "■", "●", "★", "☆", "▲", "▼", "♠", "♥", "♦",
];
function insertSymbol(ch) {
  if (!isWord()) { setStatus("插入符号仅支持 Word", "err"); return; }
  Word.run(async (ctx) => {
    ctx.document.getSelection().insertText(ch, Word.InsertLocation.replace);
    await ctx.sync();
    setStatus("已插入：" + ch, "ok");
  }).catch((e) => setStatus("插入失败：" + (e && e.message ? e.message : e), "err"));
}

// --- Shared: export current document to PDF and download ---
function getFileAsBase64(fileType) {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(fileType, { sliceSize: 65536 }, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) { reject(result.error); return; }
      const file = result.value;
      const sliceCount = file.sliceCount;
      const slices = new Array(sliceCount);
      let received = 0;
      const getSlice = (i) => {
        file.getSliceAsync(i, (res) => {
          if (res.status !== Office.AsyncResultStatus.Succeeded) { file.closeAsync(); reject(res.error); return; }
          slices[i] = res.value.data;
          received++;
          if (received === sliceCount) {
            file.closeAsync();
            resolve(slices.join(""));
          } else {
            getSlice(i + 1);
          }
        });
      };
      getSlice(0);
    });
  });
}
async function doExportPdf() {
  if (!isWord() && !isExcel()) return;
  try {
    const base64 = await getFileAsBase64(Office.FileType.Pdf);
    const byteChars = atob(base64);
    const bytes = new Uint8Array(byteChars.length);
    for (let i = 0; i < byteChars.length; i++) bytes[i] = byteChars.charCodeAt(i);
    const blob = new Blob([bytes], { type: "application/pdf" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    const srcName = (Office.context.document.url || "document").split(/[\\/]/).pop().replace(/\.[^.]+$/, "") || "document";
    a.download = srcName + ".pdf";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 2000);
    setStatus("已导出 PDF：" + a.download, "ok");
  } catch (e) {
    setStatus("导出 PDF 失败：" + (e && e.message ? e.message : e) + "（部分版本/网页版可能不支持）", "err", true);
  }
}

// --- Excel: cycle fill color ---
async function doCycleFill() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    const fmt = range.format;
    fmt.load("fill/color");
    const colors = ["#FFFFFF", "#C7CC7A", "#7560BA", "#9DD9D2", "#FFE1A8", "#E26D5C", "#A9D18E", "#F4B183"];
    await ctx.sync();
    let i = colors.indexOf((fmt.fill.color || "").toUpperCase());
    i = (i + 1) % colors.length;
    range.format.fill.color = colors[i];
    await ctx.sync();
  });
}

// --- Excel: clear contents (values) ---
async function doClearContents() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    ctx.workbook.getSelectedRange().clear(Excel.ClearApplyments.contents);
    await ctx.sync();
  });
}

// --- Excel: trim extra spaces in selection ---
async function doTrim() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    ctx.workbook.getSelectedRange().trim();
    await ctx.sync();
  });
}

// --- Excel: auto-fit ---
async function doAutoFit() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    range.getColumns().format.autofitColumns();
    range.getRows().format.autofitRows();
    await ctx.sync();
  });
}

// --- Excel: freeze panes ---
async function doFreezeTopRow() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    ctx.workbook.getSelectedRange().worksheet.freezePanes.freezeRows(1);
    await ctx.sync();
  });
}
async function doFreezeFirstColumn() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    ctx.workbook.getSelectedRange().worksheet.freezePanes.freezeColumns(1);
    await ctx.sync();
  });
}
async function doUnfreeze() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    ctx.workbook.getSelectedRange().worksheet.freezePanes.unfreeze();
    await ctx.sync();
  });
}

// --- Excel: insert row / column ---
async function doInsertRow() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    ctx.workbook.getSelectedRange().getCell(0, 0).getEntireRow().insert(Excel.InsertShiftDirection.down);
    await ctx.sync();
  });
}
async function doInsertColumn() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    ctx.workbook.getSelectedRange().getCell(0, 0).getEntireColumn().insert(Excel.InsertShiftDirection.right);
    await ctx.sync();
  });
}

// --- Excel: aggregate (SUM / AVERAGE) below each column ---
async function applyAggregate(fnName) {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    range.load("rowCount, columnCount");
    await ctx.sync();
    const rc = range.rowCount, cc = range.columnCount;
    const cols = [];
    for (let c = 0; c < cc; c++) {
      const col = range.getColumn(c);
      col.load("address");
      cols.push(col);
    }
    await ctx.sync();
    const target = range.getOffset(rc, 0);
    for (let c = 0; c < cc; c++) {
      target.getColumn(c).formula = `=${fnName}(${cols[c].address})`;
    }
    await ctx.sync();
  });
}
async function doSum() { await applyAggregate("SUM"); }
async function doAverage() { await applyAggregate("AVERAGE"); }

// --- Excel: sort ---
async function doSortAsc() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    range.getSort().apply([{ key: 0, ascending: true }]);
    await ctx.sync();
  });
}
async function doSortDesc() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    range.getSort().apply([{ key: 0, ascending: false }]);
    await ctx.sync();
  });
}

// --- Excel: chart ---
async function doChart() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    const sheet = range.worksheet;
    const chart = sheet.charts.add(Excel.ChartType.columnClustered, range, Excel.ChartSeriesBy.auto);
    chart.title.text = "图表";
    await ctx.sync();
  });
}

// --- Excel: conditional data bar ---
async function doConditionalBar() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
    cf.dataBar.barColor = "638EC6";
    await ctx.sync();
  });
}

// --- Excel: convert range to table ---
async function doAddTable() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    const table = range.worksheet.tables.add(range, true);
    table.style = "TableStyleMedium2";
    await ctx.sync();
  });
}

// --- Excel: remove duplicate rows ---
async function doRemoveDuplicates() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    ctx.workbook.getSelectedRange().removeDuplicates([0], true);
    await ctx.sync();
  });
}

// --- Excel: delete entirely blank rows within the selection ---
async function doDeleteBlankRows() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    range.load("rowCount, columnCount, values");
    await ctx.sync();
    const rc = range.rowCount, cc = range.columnCount;
    const vals = range.values;
    const blankRows = [];
    for (let r = 0; r < rc; r++) {
      let blank = true;
      for (let c = 0; c < cc; c++) {
        const v = vals[r][c];
        if (v !== null && v !== undefined && v !== "") { blank = false; break; }
      }
      if (blank) blankRows.push(r);
    }
    for (let i = blankRows.length - 1; i >= 0; i--) {
      range.getRow(blankRows[i]).delete(Excel.DeleteShiftDirection.up);
    }
    await ctx.sync();
    setStatus(`已删除 ${blankRows.length} 个空白行`, "ok");
  });
}

// --- Excel: merge / unmerge ---
async function doMergeCells() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    ctx.workbook.getSelectedRange().merge();
    await ctx.sync();
  });
}
async function doUnmergeCells() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    ctx.workbook.getSelectedRange().unmerge();
    await ctx.sync();
  });
}

// --- Excel: select the worksheet's used range ---
async function doSelectUsedRange() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.getSelectedRange().worksheet;
    sheet.getUsedRange().select();
    await ctx.sync();
  });
}

// --- Excel: sum numeric cells whose fill color matches the active cell ---
async function doSumByColor() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    const sample = ctx.workbook.getSelectedRange().getCell(0, 0);
    sample.load("format/fill/color, rowIndex, columnIndex");
    await ctx.sync();
    const target = (sample.format.fill.color || "").toUpperCase();
    if (!target || target === "#FFFFFF" || target === "WHITE" || target === "00000000") {
      setStatus("请先选中一个带填充色的单元格作为基准", "err");
      return;
    }
    const sheet = sample.worksheet;
    const used = sheet.getUsedRange();
    used.load("rowCount, columnCount, values");
    await ctx.sync();
    const m = used.rowCount, n = used.columnCount;
    const MAX = 20000;
    const lim = Math.min(m * n, MAX);
    const cells = [];
    for (let r = 0; r < m && r * n < lim; r++) {
      for (let c = 0; c < n && r * n + c < lim; c++) {
        const cell = used.getCell(r, c);
        cell.format.fill.load("color");
        cells.push(cell);
      }
    }
    await ctx.sync();
    const values = used.values;
    let sum = 0, count = 0, idx = 0;
    for (let r = 0; r < m && r * n < lim; r++) {
      for (let c = 0; c < n && r * n + c < lim; c++) {
        const fc = (cells[idx].format.fill.color || "").toUpperCase();
        idx++;
        if (fc === target) {
          const v = values[r][c];
          const num = typeof v === "number" ? v : parseFloat(String(v).replace(/,/g, ""));
          if (!isNaN(num)) { sum += num; count++; }
        }
      }
    }
    const outCell = sheet.getCell(sample.rowIndex, sample.columnIndex + 1);
    outCell.values = [[sum]];
    outCell.format.font.bold = true;
    await ctx.sync();
    setStatus(`按 ${target} 色求和：共 ${count} 个单元格，合计 ${sum}`, "ok");
  });
}

// ---------------------------------------------------------------------------
// New command batch (Excel + Word enhancements + shared eye-comfort)
// ---------------------------------------------------------------------------

// --- Excel: find & highlight duplicate values in the selection (keep first) ---
async function doFindDuplicates() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    range.load("rowCount, columnCount, values");
    await ctx.sync();
    const rc = range.rowCount, cc = range.columnCount;
    const vals = range.values;
    const seen = new Set();
    const dup = [];
    for (let r = 0; r < rc; r++) {
      for (let c = 0; c < cc; c++) {
        const v = vals[r][c];
        const key = (v === null || v === undefined) ? "" : String(v);
        if (key === "") continue;
        if (seen.has(key)) dup.push([r, c]);
        else seen.add(key);
      }
    }
    dup.forEach(([r, c]) => { range.getCell(r, c).format.fill.color = "#FFC7CE"; });
    await ctx.sync();
    setStatus(`找到 ${dup.length} 个重复单元格（已标红，保留首个）`, dup.length ? "ok" : "ok");
  });
}

// --- Excel: apply a number format string to the selection ---
async function applyNumberFormat(fmt) {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    range.numberFormat = fmt;
    await ctx.sync();
  });
}
async function doNumCurrency() { await applyNumberFormat('"¥"#,##0.00'); setStatus("已设为货币格式", "ok"); }
async function doNumPercent() { await applyNumberFormat("0.00%"); setStatus("已设为百分比格式", "ok"); }
async function doNumThousands() { await applyNumberFormat("#,##0"); setStatus("已设为千分位格式", "ok"); }
async function doNumDate() { await applyNumberFormat("yyyy-mm-dd"); setStatus("已设为日期格式", "ok"); }

// --- Excel: borders ---
const BORDER_IDX = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"];
async function setBorders(style) {
  await Excel.run(async (ctx) => {
    const b = ctx.workbook.getSelectedRange().format.borders;
    BORDER_IDX.forEach((k) => { b.getItem(Excel.BorderIndex[k]).style = style; });
    await ctx.sync();
  });
}
async function doAddBorders() { if (!isExcel()) return; await setBorders(Excel.BorderLineStyle.continuous); setStatus("已添加边框", "ok"); }
async function doRemoveBorders() { if (!isExcel()) return; await setBorders(Excel.BorderLineStyle.none); setStatus("已去除边框", "ok"); }

// --- Excel: hide / unhide rows & columns ---
async function doHideRows() { if (!isExcel()) return; await Excel.run(async (ctx) => { ctx.workbook.getSelectedRange().rowHidden = true; await ctx.sync(); }); setStatus("已隐藏选中行", "ok"); }
async function doUnhideRows() { if (!isExcel()) return; await Excel.run(async (ctx) => { ctx.workbook.getSelectedRange().rowHidden = false; await ctx.sync(); }); setStatus("已取消隐藏行", "ok"); }
async function doHideCols() { if (!isExcel()) return; await Excel.run(async (ctx) => { ctx.workbook.getSelectedRange().columnHidden = true; await ctx.sync(); }); setStatus("已隐藏选中列", "ok"); }
async function doUnhideCols() { if (!isExcel()) return; await Excel.run(async (ctx) => { ctx.workbook.getSelectedRange().columnHidden = false; await ctx.sync(); }); setStatus("已取消隐藏列", "ok"); }

// --- Excel: transpose selection to its right ---
async function doTranspose() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    range.load("rowCount, columnCount, rowIndex, columnIndex, values, numberFormat");
    await ctx.sync();
    const rc = range.rowCount, cc = range.columnCount;
    if (rc === 0 || cc === 0) { setStatus("请先选择要转置的区域", "err"); return; }
    const vals = range.values, fmts = range.numberFormat;
    const startCol = range.columnIndex + cc;
    const outRange = range.worksheet.getCell(range.rowIndex, startCol).getResizedRange(rc - 1, cc - 1);
    const newVals = [], newFmts = [];
    for (let r = 0; r < rc; r++) {
      newVals.push([]); newFmts.push([]);
      for (let c = 0; c < cc; c++) { newVals[r].push(vals[c][r]); newFmts[r].push(fmts[c][r]); }
    }
    outRange.values = newVals;
    outRange.numberFormat = newFmts;
    await ctx.sync();
    setStatus(`已转置 ${rc}×${cc} 区域至右侧`, "ok");
  });
}

// --- Excel: convert formulas in selection to their computed values ---
async function doValuesOnly() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    range.load("values, numberFormat");
    await ctx.sync();
    range.values = range.values;
    range.numberFormat = range.numberFormat;
    await ctx.sync();
    setStatus("已将选区公式转为值", "ok");
  });
}

// --- Word: delete blank pages (empty / page-break-only paragraphs) ---
async function doDeleteBlankPages() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    const paras = ctx.document.body.paragraphs;
    paras.load("items");
    await ctx.sync();
    paras.items.forEach((p) => p.load("text"));
    await ctx.sync();
    let removed = 0;
    paras.items.forEach((p) => {
      if (p.text.replace(/[\f\n\r\t ]/g, "") === "") { p.delete(); removed++; }
    });
    await ctx.sync();
    setStatus(`已清理 ${removed} 个空白段落/分页占位（多数空白页由它们导致）`, "ok");
  });
}

// --- Word: insert a section break at the cursor ---
async function doInsertSectionBreak() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    ctx.document.getSelection().insertBreak(Word.BreakType.sectionNext, Word.InsertLocation.after);
    await ctx.sync();
    setStatus("已插入分节符", "ok");
  });
}

// --- Word: case transforms ---
function applyCase(kind) {
  return async function () {
    if (!isWord()) return;
    await Word.run(async (ctx) => {
      const range = ctx.document.getSelection();
      range.load("text");
      await ctx.sync();
      const text = range.text;
      let out = text;
      if (kind === "title") out = text.replace(/\w\S*/g, (w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase());
      else if (kind === "sentence") out = text.charAt(0).toUpperCase() + text.slice(1);
      else if (kind === "toggle") {
        const up = (text.match(/[A-Z]/g) || []).length;
        const lo = (text.match(/[a-z]/g) || []).length;
        out = up > lo ? text.toLowerCase() : text.toUpperCase();
      }
      range.insertText(out, Word.InsertLocation.replace);
      await ctx.sync();
      setStatus("已应用大小写转换", "ok");
    });
  };
}
const doCaseTitle = applyCase("title");
const doCaseSentence = applyCase("sentence");
const doCaseToggle = applyCase("toggle");

// --- Word: table <-> text ---
async function doTableToText() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    const tables = ctx.document.getSelection().tables;
    tables.load("items");
    await ctx.sync();
    if (tables.items.length === 0) { setStatus("光标未在表格内", "err"); return; }
    tables.items.forEach((t) => t.convertToText(Word.CellSeparator.tab, true));
    await ctx.sync();
    setStatus("已将表格转为文本（制表符分隔）", "ok");
  });
}
async function doTextToTable() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    const range = ctx.document.getSelection();
    range.load("text");
    await ctx.sync();
    const table = range.convertToTable(Word.CellSeparator.paragraph, false, 1);
    await ctx.sync();
    setStatus("已将所选文本转为表格", "ok");
  });
}

// --- Shared: eye-comfort toggle for the task pane (pane-only UI, persists) ---
async function doToggleEyeComfort() {
  const body = document.body;
  body.classList.toggle("qc-eye-comfort");
  const on = body.classList.contains("qc-eye-comfort");
  try { localStorage.setItem("qc-eye-comfort", on ? "1" : "0"); } catch (e) {}
  setStatus(on ? "已开启护眼模式" : "已关闭护眼模式", "ok", true);
}

// ---------------------------------------------------------------------------
// Remaining high-frequency batch: Excel data tools + Word document structure
// ---------------------------------------------------------------------------

// Small prompt helper (works in desktop task panes; falls back to default).
function askInput(message, fallback) {
  try {
    if (typeof window !== "undefined" && typeof window.prompt === "function") {
      const r = window.prompt(message, fallback);
      if (r === null) return null; // user cancelled
      return r;
    }
  } catch (e) {}
  return fallback;
}

// --- Excel: split the first selected column into multiple columns by a delimiter ---
async function doSplitText() {
  if (!isExcel()) return;
  const input = askInput("请输入分列的分隔符（逗号、空格、Tab 用 \\t、竖线 | 等）：", ",");
  if (input === null) { setStatus("已取消分列", "ok"); return; }
  const delim = input === "\\t" ? "\t" : (input || ",");
  await Excel.run(async (ctx) => {
    const sel = ctx.workbook.getSelectedRange();
    const col = sel.getColumn(0);
    col.load("rowCount, rowIndex, columnIndex, values");
    await ctx.sync();
    const vals = col.values;
    const rc = col.rowCount, baseRow = col.rowIndex, baseCol = col.columnIndex;
    let maxParts = 1;
    const parts = vals.map((row) => {
      const s = row[0] == null ? "" : String(row[0]);
      const p = delim ? s.split(delim) : [s];
      if (p.length > maxParts) maxParts = p.length;
      return p;
    });
    const out = ctx.workbook.getSelectedRange().worksheet.getCell(baseRow, baseCol + 1).getResizedRange(rc - 1, maxParts - 1);
    const outVals = parts.map((p) => {
      const arr = p.slice();
      while (arr.length < maxParts) arr.push("");
      return arr;
    });
    out.values = outVals;
    await ctx.sync();
    setStatus(`已对首列分列（共 ${maxParts} 列，结果写到右侧、原数据保留）`, "ok");
  });
}

// --- Excel: select blank cells within the selection (Go To Special → Blanks) ---
async function doGoToBlanks() {
  if (!isExcel()) return;
  await Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    const blanks = range.getSpecialCellsOrNullObject(Excel.SpecialCellType.blanks);
    await ctx.sync();
    if (blanks.isNullObject) { setStatus("选区内没有空白单元格", "ok"); return; }
    blanks.select();
    await ctx.sync();
    setStatus("已选中全部空白单元格（可直接输入后按 Ctrl+Enter 批量填充）", "ok");
  });
}

// --- Excel: add a dropdown (list) data validation to the selection ---
async function doDataValidation() {
  if (!isExcel()) return;
  const input = askInput("请输入下拉选项（用逗号分隔，如 合格,不合格,待复检）：", "合格,不合格,待复检");
  if (input === null) { setStatus("已取消", "ok"); return; }
  const list = input.split(",").map((s) => s.trim()).filter((s) => s);
  if (list.length === 0) { setStatus("未输入有效选项", "err"); return; }
  await Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    range.dataValidation.rule = { list: { source: list.join(",") } };
    await ctx.sync();
    setStatus(`已为选区添加下拉：${list.join(" / ")}`, "ok");
  });
}

// --- Word: insert page numbers (centered PAGE field) into every section's footer ---
async function doInsertPageNumbers() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    const sections = ctx.document.sections;
    sections.load("items");
    await ctx.sync();
    let count = 0;
    sections.items.forEach((sec) => {
      const footer = sec.getFooter("primary");
      const para = footer.insertParagraph("", "end");
      para.alignment = Word.Alignment.centered;
      para.insertField(Word.FieldType.page, "", "");
      count++;
    });
    await ctx.sync();
    setStatus(`已在 ${count} 个节的页脚插入居中的页码`, "ok");
  });
}

// --- Word: insert header text into every section's header ---
async function doInsertHeader() {
  if (!isWord()) return;
  const input = askInput("请输入页眉文字：", "公司名称 / 文档标题");
  if (input === null) { setStatus("已取消", "ok"); return; }
  await Word.run(async (ctx) => {
    const sections = ctx.document.sections;
    sections.load("items");
    await ctx.sync();
    let count = 0;
    sections.items.forEach((sec) => {
      sec.getHeader("primary").insertParagraph(input, "end");
      count++;
    });
    await ctx.sync();
    setStatus(`已在 ${count} 个节的页眉插入文字`, "ok");
  });
}

// --- Word: insert footer text into every section's footer ---
async function doInsertFooter() {
  if (!isWord()) return;
  const input = askInput("请输入页脚文字：", "第 1 版 / 内部资料");
  if (input === null) { setStatus("已取消", "ok"); return; }
  await Word.run(async (ctx) => {
    const sections = ctx.document.sections;
    sections.load("items");
    await ctx.sync();
    let count = 0;
    sections.items.forEach((sec) => {
      sec.getFooter("primary").insertParagraph(input, "end");
      count++;
    });
    await ctx.sync();
    setStatus(`已在 ${count} 个节的页脚插入文字`, "ok");
  });
}

// --- Word: turn selected paragraphs into a multi-level numbered list (by indent) ---
async function doMultiLevelList() {
  if (!isWord()) return;
  await Word.run(async (ctx) => {
    const paras = ctx.document.getSelection().paragraphs;
    paras.load("items");
    await ctx.sync();
    const leads = paras.items.map((p) => {
      p.load("text");
      return p;
    });
    await ctx.sync();
    const levels = leads.map((p) => Math.min(8, Math.floor((p.text.match(/^[ \t]+/)[0] || "").length / 2)));
    paras.items.forEach((p) => p.startNewList(Word.ListType.numbered));
    await ctx.sync();
    paras.items.forEach((p, i) => { p.listItem.level = levels[i]; });
    await ctx.sync();
    setStatus("已将所选段落设为多级编号列表（按前导空格分级）", "ok");
  });
}

// --- Word: Simplified <-> Traditional conversion (via opencc-js, full dictionary) ---
// opencc-js is bundled statically as src/lib/opencc-bundle.js and exposes window.OpenCC.
// It is phrase-aware (correctly resolves context-sensitive chars such as 干->乾/幹,
// 里->裡/裏), unlike a flat character map, and needs no runtime network access.
let _qcConv = { s2t: null, t2s: null };
function getOpenCcConverter(toTraditional) {
  const key = toTraditional ? "s2t" : "t2s";
  if (_qcConv[key]) return _qcConv[key];
  const O = (typeof window !== "undefined" && window.OpenCC) || (typeof OpenCC !== "undefined" ? OpenCC : null);
  if (!O || typeof O.Converter !== "function") {
    throw new Error("简繁转换组件未加载（请确认已引入 lib/opencc-bundle.js）");
  }
  _qcConv[key] = O.Converter({ from: toTraditional ? "cn" : "tw", to: toTraditional ? "tw" : "cn" });
  return _qcConv[key];
}

async function doConvertTrad() {
  if (!isWord()) return;
  let conv;
  try { conv = getOpenCcConverter(true); } catch (e) { setStatus(e.message, "err", true); return; }
  await Word.run(async (ctx) => {
    const range = ctx.document.getSelection();
    range.load("text");
    await ctx.sync();
    range.insertText(conv(range.text), Word.InsertLocation.replace);
    await ctx.sync();
    setStatus("已将选区转为繁体（opencc 全量词典）", "ok");
  });
}
async function doConvertSimp() {
  if (!isWord()) return;
  let conv;
  try { conv = getOpenCcConverter(false); } catch (e) { setStatus(e.message, "err", true); return; }
  await Word.run(async (ctx) => {
    const range = ctx.document.getSelection();
    range.load("text");
    await ctx.sync();
    range.insertText(conv(range.text), Word.InsertLocation.replace);
    await ctx.sync();
    setStatus("已将选区转为简体（opencc 全量词典）", "ok");
  });
}

// --- Settings persistence (localStorage, best-effort) ---
function getSetting(key, def) {
  try { const v = localStorage.getItem("qc-" + key); return v === null ? def : v; } catch (e) { return def; }
}
function setSetting(key, val) {
  try { localStorage.setItem("qc-" + key, val); } catch (e) {}
}

// --- Shared: advanced find & replace (Word + Excel) with options ---
function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
// Replace within a single string, honoring match-case / whole-word. Returns count.
function replaceInText(text, find, replace, matchCase, wholeWord) {
  if (!text) return { text: text, count: 0 };
  let count = 0;
  let out;
  if (wholeWord) {
    // whole-word: only when not preceded/followed by a word char (letters/digits/_)
    const re = new RegExp("(^|[^\\w])(" + escapeRegExp(find) + ")(?=$|[^\\w])", matchCase ? "g" : "gi");
    out = text.replace(re, (m, p1) => { count++; return p1 + replace; });
  } else {
    const re = new RegExp(escapeRegExp(find), matchCase ? "g" : "gi");
    out = text.replace(re, () => { count++; return replace; });
  }
  return { text: out, count };
}

async function doAdvancedReplace() {
  if (!isWord() && !isExcel()) return;
  const find = askInput("查找内容：", getSetting("lastFind", ""));
  if (find === null) { setStatus("已取消替换", "ok"); return; }
  if (find === "") { setStatus("查找内容为空", "err"); return; }
  const replace = askInput("替换为：", getSetting("lastReplace", ""));
  if (replace === null) { setStatus("已取消替换", "ok"); return; }
  const optsInput = askInput("选项（留空=普通；case=区分大小写；word=全字匹配；case,word=两者）：", getSetting("lastReplaceOpts", ""));
  if (optsInput === null) { setStatus("已取消替换", "ok"); return; }
  const opts = (optsInput || "").split(",").map((s) => s.trim()).filter(Boolean);
  const matchCase = opts.includes("case");
  const wholeWord = opts.includes("word");
  setSetting("lastFind", find);
  setSetting("lastReplace", replace);
  setSetting("lastReplaceOpts", optsInput || "");

  if (isWord()) {
    await Word.run(async (ctx) => {
      const range = ctx.document.getSelection();
      range.load("text");
      await ctx.sync();
      const { text: out, count } = replaceInText(range.text, find, replace, matchCase, wholeWord);
      range.insertText(out, Word.InsertLocation.replace);
      await ctx.sync();
      setStatus(`已在 Word 选区替换 ${count} 处`, "ok");
    });
  } else if (isExcel()) {
    await Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load("rowCount, columnCount, values");
      await ctx.sync();
      const rc = range.rowCount, cc = range.columnCount;
      const vals = range.values;
      let total = 0;
      const out = vals.map((row) => row.map((v) => {
        if (typeof v !== "string") return v;
        const r = replaceInText(v, find, replace, matchCase, wholeWord);
        total += r.count;
        return r.text;
      }));
      range.values = out;
      await ctx.sync();
      setStatus(`已在 Excel 选区替换 ${total} 处`, "ok");
    });
  }
}

// ---------------------------------------------------------------------------
// Wrappers
//   - Keyboard shortcuts: function returns a Promise.
//   - Ribbon ExecuteFunction: function receives `event` and MUST call event.completed().
// ---------------------------------------------------------------------------
function wrapKb(fn) {
  return () => Promise.resolve().then(fn);
}

function wrapRibbon(fn) {
  return (event) => {
    Promise.resolve()
      .then(fn)
      .then(() => event.completed())
      .catch((e) => {
        console.error("快捷指令命令执行失败:", e);
        event.completed();
      });
  };
}

// ---------------------------------------------------------------------------
// Action registry: every id used in shortcuts.json or manifest.xml is registered.
//   KB    -> keyboard shortcuts (returns a promise)
//   RIBBON-> ribbon ExecuteFunction buttons (receives event, calls completed)
// ---------------------------------------------------------------------------
const KB = [
  ["KbDate", doInsertDate],
  ["KbBold", doBold],
  ["KbItalic", doItalic],
  ["KbUnderline", doUnderline],
  ["KbHeading1", doHeading1],
  ["KbBulletList", doBulletList],
  ["KbNumberList", doNumberList],
  ["KbAlignCenter", doAlignCenter],
  ["KbHighlight", doHighlight],
  ["KbClear", doClearFormat],
  ["KbUppercase", doUppercase],
  ["KbCycleFill", doCycleFill],
  ["KbAutoFit", doAutoFit],
  ["KbFreezeTop", doFreezeTopRow],
  ["KbUnfreeze", doUnfreeze],
  ["KbInsertRow", doInsertRow],
  ["KbSum", doSum],
  ["KbSortAsc", doSortAsc],
  ["KbChart", doChart],
  ["KbClearContents", doClearContents],
  // New
  ["KbExportPdf", doExportPdf],
  ["KbWordCount", doWordCount],
  ["KbTrim", doTrim],
  ["KbSumByColor", doSumByColor],
  ["KbMerge", doMergeCells],
  // New keyboard shortcuts
  ["KbFindDup", doFindDuplicates],
  ["KbAddBorders", doAddBorders],
  ["KbTranspose", doTranspose],
  ["KbValuesOnly", doValuesOnly],
  // Remaining batch
  ["KbSplitText", doSplitText],
  ["KbGoToBlanks", doGoToBlanks],
  ["KbDataValidation", doDataValidation],
  ["KbPageNumbers", doInsertPageNumbers],
  ["KbInsertHeader", doInsertHeader],
  ["KbInsertFooter", doInsertFooter],
  ["KbMultiList", doMultiLevelList],
  ["KbConvertTrad", doConvertTrad],
  ["KbConvertSimp", doConvertSimp],
  ["KbAdvReplace", doAdvancedReplace],
];

const RIBBON = [
  // Word
  ["RibbonDate", doInsertDate],
  ["RibbonBold", doBold],
  ["RibbonItalic", doItalic],
  ["RibbonUnderline", doUnderline],
  ["RibbonHeading1", doHeading1],
  ["RibbonBulletList", doBulletList],
  ["RibbonNumberList", doNumberList],
  ["RibbonAlignCenter", doAlignCenter],
  ["RibbonHighlight", doHighlight],
  ["RibbonClear", doClearFormat],
  ["RibbonRemoveEmptyLines", doRemoveEmptyLines],
  ["RibbonRemoveHighlight", doRemoveHighlight],
  ["RibbonExportPdf", doExportPdf],
  ["RibbonWordCount", doWordCount],
  // Excel
  ["RibbonCycleFill", doCycleFill],
  ["RibbonAutoFit", doAutoFit],
  ["RibbonFreezeTopRow", doFreezeTopRow],
  ["RibbonUnfreeze", doUnfreeze],
  ["RibbonInsertRow", doInsertRow],
  ["RibbonSum", doSum],
  ["RibbonSortAsc", doSortAsc],
  ["RibbonChart", doChart],
  ["RibbonTrim", doTrim],
  ["RibbonSumByColor", doSumByColor],
  ["RibbonMergeCells", doMergeCells],
  ["RibbonExportPdf", doExportPdf],
  // New Word ribbon
  ["RibbonDeleteBlankPages", doDeleteBlankPages],
  ["RibbonPageBreak", doInsertPageBreak],
  ["RibbonSectionBreak", doInsertSectionBreak],
  ["RibbonCaseLower", doLowercase],
  ["RibbonCaseTitle", doCaseTitle],
  ["RibbonCaseSentence", doCaseSentence],
  ["RibbonCaseToggle", doCaseToggle],
  ["RibbonTableToText", doTableToText],
  ["RibbonTextToTable", doTextToTable],
  ["RibbonEyeComfort", doToggleEyeComfort],
  // New Excel ribbon
  ["RibbonFindDup", doFindDuplicates],
  ["RibbonNumCurrency", doNumCurrency],
  ["RibbonNumPercent", doNumPercent],
  ["RibbonNumThousands", doNumThousands],
  ["RibbonNumDate", doNumDate],
  ["RibbonAddBorders", doAddBorders],
  ["RibbonRemoveBorders", doRemoveBorders],
  ["RibbonHideRows", doHideRows],
  ["RibbonUnhideRows", doUnhideRows],
  ["RibbonHideCols", doHideCols],
  ["RibbonUnhideCols", doUnhideCols],
  ["RibbonTranspose", doTranspose],
  ["RibbonValuesOnly", doValuesOnly],
  // Remaining batch
  ["RibbonSplitText", doSplitText],
  ["RibbonGoToBlanks", doGoToBlanks],
  ["RibbonDataValidation", doDataValidation],
  ["RibbonPageNumbers", doInsertPageNumbers],
  ["RibbonInsertHeader", doInsertHeader],
  ["RibbonInsertFooter", doInsertFooter],
  ["RibbonMultiLevelList", doMultiLevelList],
  ["RibbonConvertTrad", doConvertTrad],
  ["RibbonConvertSimp", doConvertSimp],
  ["RibbonAdvReplace", doAdvancedReplace],
];

function registerActions() {
  KB.forEach(([id, fn]) => Office.actions.associate(id, wrapKb(fn)));
  RIBBON.forEach(([id, fn]) => Office.actions.associate(id, wrapRibbon(fn)));
  // Panel show / hide (keyboard only)
  Office.actions.associate("KbShow", () => Office.addin.showAsTaskpane().catch(() => {}));
  Office.actions.associate("KbHide", () => Office.addin.hide().catch(() => {}));
}

// ---------------------------------------------------------------------------
// Task pane UI — full command catalogue (filtered by current host)
// ---------------------------------------------------------------------------
const COMMANDS = [
  // Shared
  { name: "插入日期", desc: "Excel：选区填充当前日期；Word：光标处插入日期时间。", kb: "Ctrl+Alt+D", hosts: ["word", "excel"], run: doInsertDate },
  { name: "加粗", desc: "切换选区加粗（Word 字体 / Excel 单元格）。", kb: "Ctrl+Alt+B", hosts: ["word", "excel"], run: doBold },
  { name: "斜体", desc: "切换选区斜体。", kb: "Ctrl+Alt+I", hosts: ["word", "excel"], run: doItalic },
  { name: "下划线", desc: "切换选区下划线。", kb: "Ctrl+Alt+U", hosts: ["word", "excel"], run: doUnderline },
  { name: "突出显示", desc: "Excel 填充黄色 / Word 高亮选区。", kb: "Ctrl+Alt+H", hosts: ["word", "excel"], run: doHighlight },
  { name: "清除格式", desc: "清除选区格式（Excel 仅清除格式 / Word 清除字体与对齐）。", kb: "Ctrl+Alt+K", hosts: ["word", "excel"], run: doClearFormat },
  { name: "左对齐", desc: "段落 / 单元格左对齐。", kb: "", hosts: ["word", "excel"], run: doAlignLeft },
  { name: "居中", desc: "段落 / 单元格居中对齐。", kb: "Ctrl+Alt+C", hosts: ["word", "excel"], run: doAlignCenter },
  { name: "右对齐", desc: "段落 / 单元格右对齐。", kb: "", hosts: ["word", "excel"], run: doAlignRight },
  { name: "两端对齐", desc: "段落 / 单元格两端对齐。", kb: "", hosts: ["word", "excel"], run: doAlignJustify },
  { name: "导出 PDF", desc: "将当前文档导出为 PDF 并触发下载。", kb: "Ctrl+Alt+P", hosts: ["word", "excel"], run: doExportPdf },
  { name: "护眼模式", desc: "切换任务窗格的柔和配色（仅影响本面板，状态记忆）。", kb: "", hosts: ["word", "excel"], run: doToggleEyeComfort },

  // Word only
  { name: "标题 1", desc: "将选区设为“标题 1”样式。", kb: "Ctrl+Alt+J", hosts: ["word"], run: doHeading1 },
  { name: "标题 2", desc: "将选区设为“标题 2”样式。", kb: "", hosts: ["word"], run: doHeading2 },
  { name: "正文", desc: "将选区设为正文（Normal）样式。", kb: "", hosts: ["word"], run: doBodyText },
  { name: "项目符号列表", desc: "为选区段落添加/取消项目符号。", kb: "Ctrl+Alt+L", hosts: ["word"], run: doBulletList },
  { name: "编号列表", desc: "为选区段落添加/取消编号。", kb: "Ctrl+Alt+N", hosts: ["word"], run: doNumberList },
  { name: "取消列表", desc: "移除选区段落的列表格式。", kb: "", hosts: ["word"], run: doRemoveList },
  { name: "1.5 倍行距", desc: "将选区段落设为 1.5 倍行距。", kb: "", hosts: ["word"], run: doLineSpacing15 },
  { name: "2 倍行距", desc: "将选区段落设为 2 倍行距。", kb: "", hosts: ["word"], run: doLineSpacingDouble },
  { name: "首行缩进", desc: "选区段落首行缩进约 2 字符。", kb: "", hosts: ["word"], run: doFirstLineIndent },
  { name: "取消缩进", desc: "选区段落首行缩进归零。", kb: "", hosts: ["word"], run: doNoIndent },
  { name: "段前分页", desc: "为所选段落设置“段前分页”。", kb: "", hosts: ["word"], run: doPageBreakBefore },
  { name: "转为大写", desc: "将选区文本转为大写。", kb: "Ctrl+Alt+Shift+U", hosts: ["word"], run: doUppercase },
  { name: "转为小写", desc: "将选区文本转为小写。", kb: "", hosts: ["word"], run: doLowercase },
  { name: "插入 3×3 表格", desc: "在光标后插入一个 3 行 3 列表格。", kb: "", hosts: ["word"], run: doInsertTable },
  { name: "插入分页符", desc: "在光标后插入分页符。", kb: "", hosts: ["word"], run: doInsertPageBreak },
  { name: "删除空段落", desc: "删除所选（或全文）中的空段落。", kb: "", hosts: ["word"], run: doRemoveEmptyLines },
  { name: "清理多余空格", desc: "将所选段落内连续空格合并为单个并去首尾空格。", kb: "", hosts: ["word"], run: doNormalizeSpaces },
  { name: "清除高亮", desc: "取消所选内容的高亮颜色。", kb: "", hosts: ["word"], run: doRemoveHighlight },
  { name: "字数统计", desc: "统计全文词数与选区字数/字符数。", kb: "Ctrl+Alt+W", hosts: ["word"], run: doWordCount },
  { name: "删除空白页", desc: "删除由空段落/分页占位导致的空白页。", kb: "", hosts: ["word"], run: doDeleteBlankPages },
  { name: "插入分节符", desc: "在光标后插入分节符。", kb: "", hosts: ["word"], run: doInsertSectionBreak },
  { name: "转为标题大小写", desc: "将选区每个单词首字母大写。", kb: "", hosts: ["word"], run: doCaseTitle },
  { name: "转为句首大写", desc: "将选区句首字母大写。", kb: "", hosts: ["word"], run: doCaseSentence },
  { name: "大小写切换", desc: "依据当前大小写在大写/小写间切换。", kb: "", hosts: ["word"], run: doCaseToggle },
  { name: "表格转文本", desc: "将光标所在表格转为制表符分隔的文本。", kb: "", hosts: ["word"], run: doTableToText },
  { name: "文本转表格", desc: "将所选段落文本转为单列表格。", kb: "", hosts: ["word"], run: doTextToTable },
  { name: "插入页码", desc: "在每个节的页脚居中插入页码（PAGE 域）。", kb: "Ctrl+Alt+Shift+P", hosts: ["word"], run: doInsertPageNumbers },
  { name: "插入页眉", desc: "在每个节的页眉插入指定文字。", kb: "Ctrl+Alt+Shift+H", hosts: ["word"], run: doInsertHeader },
  { name: "插入页脚", desc: "在每个节的页脚插入指定文字。", kb: "Ctrl+Alt+Shift+J", hosts: ["word"], run: doInsertFooter },
  { name: "多级列表", desc: "将所选段落设为多级编号（按前导空格分级）。", kb: "Ctrl+Alt+Shift+M", hosts: ["word"], run: doMultiLevelList },
  { name: "转为繁体", desc: "将选区简体转为繁体（opencc 全量词典，组件已随加载项打包）。", kb: "Ctrl+Alt+Shift+T", hosts: ["word"], run: doConvertTrad },
  { name: "转为简体", desc: "将选区繁体转为简体（opencc 全量词典，组件已随加载项打包）。", kb: "Ctrl+Alt+Shift+S", hosts: ["word"], run: doConvertSimp },
  { name: "高级替换", desc: "在选区中查找并替换，可选区分大小写 / 全字匹配（记忆上次输入）。", kb: "Ctrl+Alt+Shift+R", hosts: ["word"], run: doAdvancedReplace },

  // Excel only
  { name: "循环填充色", desc: "在 8 色调色板中循环切换选区填充色。", kb: "Ctrl+Alt+Q", hosts: ["excel"], run: doCycleFill },
  { name: "清除内容", desc: "仅清除选区的值（保留格式）。", kb: "Ctrl+Alt+X", hosts: ["excel"], run: doClearContents },
  { name: "去空格（Trim）", desc: "去除选区单元格内首尾及多余空格。", kb: "Ctrl+Alt+T", hosts: ["excel"], run: doTrim },
  { name: "自动调整", desc: "自动调整选区列宽与行高。", kb: "Ctrl+Alt+A", hosts: ["excel"], run: doAutoFit },
  { name: "冻结首行", desc: "冻结当前工作表的首行。", kb: "Ctrl+Alt+F", hosts: ["excel"], run: doFreezeTopRow },
  { name: "冻结首列", desc: "冻结当前工作表的首列。", kb: "", hosts: ["excel"], run: doFreezeFirstColumn },
  { name: "取消冻结", desc: "取消所有冻结窗格。", kb: "Ctrl+Alt+Shift+F", hosts: ["excel"], run: doUnfreeze },
  { name: "插入行", desc: "在当前选区上方插入一行。", kb: "Ctrl+Alt+R", hosts: ["excel"], run: doInsertRow },
  { name: "插入列", desc: "在当前选区左侧插入一列。", kb: "", hosts: ["excel"], run: doInsertColumn },
  { name: "求和", desc: "在选区每列下方插入 SUM 公式。", kb: "Ctrl+Alt+S", hosts: ["excel"], run: doSum },
  { name: "平均值", desc: "在选区每列下方插入 AVERAGE 公式。", kb: "", hosts: ["excel"], run: doAverage },
  { name: "升序排序", desc: "按第一列对选区升序排序。", kb: "Ctrl+Alt+E", hosts: ["excel"], run: doSortAsc },
  { name: "降序排序", desc: "按第一列对选区降序排序。", kb: "", hosts: ["excel"], run: doSortDesc },
  { name: "生成图表", desc: "用选区数据生成聚类柱形图。", kb: "Ctrl+Alt+G", hosts: ["excel"], run: doChart },
  { name: "数据条", desc: "为选区添加数据条条件格式。", kb: "", hosts: ["excel"], run: doConditionalBar },
  { name: "转为表格", desc: "将选区转换为带样式的表格。", kb: "", hosts: ["excel"], run: doAddTable },
  { name: "删除重复行", desc: "按首列删除重复行（保留表头）。", kb: "", hosts: ["excel"], run: doRemoveDuplicates },
  { name: "删除空白行", desc: "删除选区中的整行空白行。", kb: "", hosts: ["excel"], run: doDeleteBlankRows },
  { name: "合并单元格", desc: "合并所选单元格。", kb: "Ctrl+Alt+M", hosts: ["excel"], run: doMergeCells },
  { name: "拆分单元格", desc: "取消所选区域的合并。", kb: "", hosts: ["excel"], run: doUnmergeCells },
  { name: "选中数据区", desc: "选中当前工作表的已用区域。", kb: "", hosts: ["excel"], run: doSelectUsedRange },
  { name: "按颜色求和", desc: "以活动单元格填充色为基准，对全表同色数值求和并写入其右侧。", kb: "Ctrl+Alt+Y", hosts: ["excel"], run: doSumByColor },
  { name: "查找重复值", desc: "高亮选区中的重复值（保留首个，标红）。", kb: "Ctrl+Alt+O", hosts: ["excel"], run: doFindDuplicates },
  { name: "设为货币格式", desc: "选区设为货币格式（¥ 千分位两位小数）。", kb: "", hosts: ["excel"], run: doNumCurrency },
  { name: "设为百分比", desc: "选区设为百分比格式。", kb: "", hosts: ["excel"], run: doNumPercent },
  { name: "设为千分位", desc: "选区设为千分位整数格式。", kb: "", hosts: ["excel"], run: doNumThousands },
  { name: "设为日期格式", desc: "选区设为 yyyy-mm-dd 日期格式。", kb: "", hosts: ["excel"], run: doNumDate },
  { name: "加边框", desc: "为选区添加全边框。", kb: "Ctrl+Alt+Shift+B", hosts: ["excel"], run: doAddBorders },
  { name: "去边框", desc: "去除选区所有边框。", kb: "", hosts: ["excel"], run: doRemoveBorders },
  { name: "隐藏选中行", desc: "隐藏当前选中的行。", kb: "", hosts: ["excel"], run: doHideRows },
  { name: "取消隐藏行", desc: "取消隐藏当前选中的行。", kb: "", hosts: ["excel"], run: doUnhideRows },
  { name: "隐藏选中列", desc: "隐藏当前选中的列。", kb: "", hosts: ["excel"], run: doHideCols },
  { name: "取消隐藏列", desc: "取消隐藏当前选中的列。", kb: "", hosts: ["excel"], run: doUnhideCols },
  { name: "转置", desc: "将选区转置到其右侧。", kb: "Ctrl+Alt+Z", hosts: ["excel"], run: doTranspose },
  { name: "公式转值", desc: "将选区公式替换为计算所得的值。", kb: "Ctrl+Alt+V", hosts: ["excel"], run: doValuesOnly },
  { name: "文本分列", desc: "按指定分隔符将选区首列拆分到右侧多列（原数据保留）。", kb: "Ctrl+Alt+Shift+D", hosts: ["excel"], run: doSplitText },
  { name: "定位空值", desc: "选中选区内的所有空白单元格（类似 Go To Special → 空值）。", kb: "Ctrl+Alt+Shift+K", hosts: ["excel"], run: doGoToBlanks },
  { name: "数据有效性", desc: "为选区添加下拉列表数据有效性（逗号分隔选项）。", kb: "Ctrl+Alt+Shift+L", hosts: ["excel"], run: doDataValidation },
  { name: "高级替换", desc: "在选区中查找并替换，可选区分大小写 / 全字匹配（记忆上次输入）。", kb: "Ctrl+Alt+Shift+R", hosts: ["excel"], run: doAdvancedReplace },
];

let statusTimer = null;
function setStatus(msg, type, persist) {
  const el = document.getElementById("qc-status");
  if (!el) return;
  el.textContent = msg;
  el.className = "qc-status" + (type ? " " + type : "");
  if (statusTimer) clearTimeout(statusTimer);
  if (type && !persist) {
    statusTimer = setTimeout(() => {
      el.textContent = "";
      el.className = "qc-status";
    }, 4000);
  }
}

function setStats(msg) {
  const el = document.getElementById("qc-stats");
  if (!el) return;
  el.textContent = msg;
  el.className = "qc-stats";
}

function buildUI() {
  const hostName = isExcel() ? "Excel" : isWord() ? "Word" : "Office";
  const info = document.getElementById("host-info");
  if (info) info.textContent = `当前主机：${hostName}`;

  const list = document.getElementById("qc-list");
  list.innerHTML = "";

  const visible = COMMANDS.filter((c) => c.hosts.includes(hostName.toLowerCase()));

  visible.forEach((cmd) => {
    const btn = document.createElement("button");
    btn.className = "qc-btn";
    btn.type = "button";
    btn.innerHTML =
      `<span class="qc-name">${cmd.name}</span>` +
      `<span class="qc-desc">${cmd.desc}</span>` +
      (cmd.kb ? `<span class="qc-kb">快捷键 ${cmd.kb}</span>` : "");
    btn.addEventListener("click", async () => {
      try {
        await cmd.run();
        if (!document.getElementById("qc-status").textContent.startsWith("已")) {
          setStatus(`已执行：${cmd.name}`, "ok");
        }
      } catch (e) {
        console.error(e);
        setStatus(`执行失败：${cmd.name}（${e && e.message ? e.message : e}）`, "err");
      }
    });
    list.appendChild(btn);
  });

  // Word-only: special symbols grid
  const box = document.getElementById("qc-symbols-box");
  const symWrap = document.getElementById("qc-symbols");
  if (box && symWrap) {
    if (isWord()) {
      box.hidden = false;
      symWrap.innerHTML = "";
      SYMBOLS.forEach((ch) => {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "qc-sym";
        b.textContent = ch;
        b.title = "插入 " + ch;
        b.addEventListener("click", () => insertSymbol(ch));
        symWrap.appendChild(b);
      });
    } else {
      box.hidden = true;
    }
  }

  const search = document.getElementById("qc-search");
  search.addEventListener("input", () => {
    const q = search.value.trim().toLowerCase();
    Array.from(list.children).forEach((child, i) => {
      const text = (visible[i].name + visible[i].desc).toLowerCase();
      child.style.display = !q || text.includes(q) ? "" : "none";
    });
  });
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------
Office.onReady(() => {
  try { if (localStorage.getItem("qc-eye-comfort") === "1") document.body.classList.add("qc-eye-comfort"); } catch (e) {}
  buildUI();
  registerActions();
});
