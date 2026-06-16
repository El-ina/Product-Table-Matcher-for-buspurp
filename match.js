"use strict";

const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");


const CYR2LAT = { а:"a", в:"b", е:"e", к:"k", м:"m", н:"h", о:"o", р:"p", с:"c", т:"t", у:"y", х:"x" };
const foldConfusables = (s) => s.replace(/[авекмнорстух]/g, (ch) => CYR2LAT[ch] || ch);

const normBG = (s) => String(s).toLowerCase().replace(/\s+/g, " ").trim();
const normDE = (s) => foldConfusables(String(s).toLowerCase()).replace(/\s+/g, " ").trim();

function volumeToLiters(numStr, unit) {
  const n = parseFloat(numStr.replace(/\s/g, "").replace(",", "."));
  return /м/.test(unit) || /ml/i.test(unit) ? n / 1000 : n;
}

function parseEntry(rawInput) {
  let raw = String(rawInput).replace(/^\s*lebosol\s*-\s*/i, "").trim();
  let m, base, liters;

  if ((m = raw.match(/\(([\d.,\s]+)\s*(л|l)\.?\)\s*$/i))) {
    liters = volumeToLiters(m[1], m[2]); base = raw.slice(0, m.index);
  } else if ((m = raw.match(/[-–]\s*([\d][\d.,\s]*?)\s*(мл|ml|л|l)\.?\s*$/i))) {
    liters = volumeToLiters(m[1], m[2]); base = raw.slice(0, m.index);
  } else if ((m = raw.match(/\s([\d.,]+)\s*(л|l)\.?\s*$/i))) {
    liters = volumeToLiters(m[1], m[2]); base = raw.slice(0, m.index);
  } else {
    return null;
  }

  base = base.replace(/[-–\s]+$/, "").trim();
  return { base, liters };
}

function cellVal(cell) {
  const v = cell.value;
  if (v === null || v === undefined) return null;
  if (typeof v === "object") {
    if ("result" in v) return v.result;
    if ("text" in v) return v.text;
    if ("richText" in v) return v.richText.map((t) => t.text).join("");
    return v;
  }
  return v;
}

async function loadTranslations(translationsPath) {
  if (!fs.existsSync(translationsPath)) {
    throw new Error(
      `Translations file not found: ${translationsPath}\n` +
      `Create a two-column Excel file (Bulgarian base name | German base name) and place it there.`
    );
  }
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(translationsPath);
  const ws = wb.worksheets[0];
  const map = {};
  let count = 0;
  for (let r = 1; r <= ws.rowCount; r++) {
    const bg = cellVal(ws.getCell(r, 1));
    const de = cellVal(ws.getCell(r, 2));
    if (bg === null || de === null) continue;
    if (/^(bulgarian|български)/i.test(String(bg).trim())) continue; // skip header
    map[normBG(bg)] = String(de).trim();
    count++;
  }
  console.log(`Loaded ${count} translation(s) from ${translationsPath}.`);
  return map;
}

async function loadSheet(filePath, sheetName) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);

  const ws = wb.getWorksheet(sheetName);
  if (!ws) {
    const available = wb.worksheets.map((w) => `"${w.name}"`).join(", ");
    throw new Error(`Sheet "${sheetName}" not found. Available sheets: ${available}`);
  }

  let headerRow = null;
  for (let r = 1; r <= ws.rowCount; r++) {
    const a = cellVal(ws.getCell(r, 1));
    if (a && String(a).trim().toLowerCase() === "artikel") { headerRow = r; break; }
  }
  if (!headerRow) throw new Error('Could not find the German header (cell "Artikel" in column A).');

  const germanTable = [];
  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const name = cellVal(ws.getCell(r, 1));
    if (name === null || String(name).trim() === "") continue;
    germanTable.push({ name: String(name), row: r });
  }

  const bulgarianTable = [];
  for (let r = 1; r <= ws.rowCount; r++) {
    const name = cellVal(ws.getCell(r, 8));
    if (name === null || String(name).trim() === "") continue;
    bulgarianTable.push({
      name: String(name),
      menge: cellVal(ws.getCell(r, 9)),
      ek: cellVal(ws.getCell(r, 10)),
      vk: cellVal(ws.getCell(r, 11)),
      row: r,
    });
  }

  return { wb, ws, germanTable, bulgarianTable };
}

function buildGermanIndex(germanTable) {
  const idx = new Map();
  for (const g of germanTable) {
    const p = parseEntry(g.name);
    if (!p) continue;
    const key = normDE(p.base) + "|" + p.liters;
    if (!idx.has(key)) idx.set(key, []);
    idx.get(key).push(g);
  }
  return idx;
}

function findGermanMatch(bgRow, germanIndex, translations) {
  const p = parseEntry(bgRow.name);
  if (!p) return { ok: false, reason: "no-volume-parsed" };

  const deBase = translations[normBG(p.base)];
  if (!deBase) return { ok: false, reason: "no-translation", base: p.base, liters: p.liters };

  const hits = germanIndex.get(normDE(deBase) + "|" + p.liters);
  if (!hits || hits.length === 0)
    return { ok: false, reason: "no-german-row", base: p.base, liters: p.liters, deBase };

  return { ok: true, match: hits[0], extra: hits.slice(1) };
}

const COL = { Menge: 3, EK: 4, VK: 5, FLAG: 12 };

function transferAndFlag(ws, bulgarianTable, germanIndex, translations) {
  const unmatched = [];
  const warnings = [];
  const usedGermanRows = new Map();
  let transferred = 0;

  for (const bg of bulgarianTable) {
    const r = findGermanMatch(bg, germanIndex, translations);

    if (!r.ok) {
      const cell = ws.getCell(bg.row, COL.FLAG);
      cell.value = "<";
      cell.font = { color: { argb: "FFFF0000" }, bold: true };
      unmatched.push({ row: bg.row, name: bg.name, reason: r.reason, base: r.base, liters: r.liters });
      continue;
    }

    const g = r.match;
    if (usedGermanRows.has(g.row)) {
      warnings.push(`German "${g.name}" (row ${g.row}) matched by both "${usedGermanRows.get(g.row)}" and "${bg.name}".`);
    }
    usedGermanRows.set(g.row, bg.name);
    if (r.extra.length) {
      warnings.push(`"${bg.name}" matched ${r.extra.length + 1} German rows; wrote to the first ("${g.name}").`);
    }

    ws.getCell(g.row, COL.Menge).value = bg.menge;
    ws.getCell(g.row, COL.EK).value = bg.ek;
    ws.getCell(g.row, COL.VK).value = bg.vk;
    transferred++;
  }

  return { transferred, unmatched, warnings };
}

async function main() {
  const argv = process.argv.slice(2);

  let translationsPath = null;
  const mi = argv.indexOf("--map");
  if (mi !== -1) { translationsPath = argv[mi + 1]; argv.splice(mi, 2); }

  const [inputFile, sheetName, outputArg] = argv;
  if (!inputFile || !sheetName) {
    console.error("Usage: node match.js <inputFile> <sheetName> [outputFile] [--map <file>]");
    process.exit(1);
  }
  const outputFile = outputArg || inputFile; // in place by default
  if (!translationsPath) translationsPath = path.join(path.dirname(inputFile) || ".", "translations.xlsx");

  const translations = await loadTranslations(translationsPath);
  const { wb, ws, germanTable, bulgarianTable } = await loadSheet(inputFile, sheetName);
  console.log(`Sheet "${sheetName}": ${germanTable.length} German rows, ${bulgarianTable.length} Bulgarian rows.`);

  const germanIndex = buildGermanIndex(germanTable);
  const { transferred, unmatched, warnings } = transferAndFlag(ws, bulgarianTable, germanIndex, translations);

  const tmp = path.join(path.dirname(outputFile) || ".", "~$match_tmp_" + Date.now() + ".xlsx");
  await wb.xlsx.writeFile(tmp);
  await fs.promises.copyFile(tmp, outputFile);
  await fs.promises.unlink(tmp);

  console.log(`\nTransferred ${transferred} product(s) into German C/D/E.`);

  const newProducts = [];
  const seen = new Set();
  for (const u of unmatched) {
    if (u.reason !== "no-translation" || !u.base) continue;
    const k = normBG(u.base);
    if (seen.has(k)) continue;
    seen.add(k);
    newProducts.push(u);
  }
  const sizeGaps = unmatched.filter((u) => u.reason === "no-german-row");
  const unparsed = unmatched.filter((u) => u.reason === "no-volume-parsed");

  if (newProducts.length) {
    console.log(`\n>> ${newProducts.length} NEW product(s) with no known translation — add a row to ${path.basename(translationsPath)}:`);
    for (const u of newProducts) console.log(`     "${u.base}"   ->   ???        (seen at ${u.liters} l)`);
  }
  if (sizeGaps.length) {
    console.log(`\n   ${sizeGaps.length} translated but no matching German size (no slot to fill — skipped):`);
    for (const u of sizeGaps) console.log(`     ${u.name}  (${u.base} @ ${u.liters} l)`);
  }
  if (unparsed.length) {
    console.log(`\n   ${unparsed.length} row(s) whose volume couldn't be parsed:`);
    for (const u of unparsed) console.log(`     ${u.name}`);
  }
  if (warnings.length) {
    console.log("\n   Warnings:");
    for (const w of warnings) console.log("     ! " + w);
  }
  console.log(`\nAll unmatched rows flagged with a red "<" in column L.`);
  console.log(`Saved -> ${outputFile}${outputArg ? "" : "  (in place)"}`);
}

main().catch((err) => { console.error("ERROR:", err.message); process.exit(1); });