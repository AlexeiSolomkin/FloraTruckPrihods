"use strict";

// ═══════════════════════════════════════════════════════════════════════════
//  XML HELPERS
// ═══════════════════════════════════════════════════════════════════════════

// workbook.xml + rels → { sheetName: 'xl/worksheets/sheetN.xml' }
function parseWorkbook(wbXml, wbRels) {
  const p = new DOMParser();
  const wb = p.parseFromString(wbXml, "application/xml");
  const rel = p.parseFromString(wbRels, "application/xml");
  const ridMap = {};
  for (const r of Array.from(rel.getElementsByTagName("Relationship")))
    ridMap[r.getAttribute("Id")] = r.getAttribute("Target");
  const REL_NS =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const result = {};
  for (const s of Array.from(wb.getElementsByTagName("sheet"))) {
    const name = s.getAttribute("name");
    const rid = s.getAttributeNS(REL_NS, "id") || s.getAttribute("r:id");
    if (rid && ridMap[rid]) result[name] = "xl/" + ridMap[rid];
  }
  return result;
}

// Записать значение в ячейку:
//   число  → t убираем, <v>число</v>
//   строка → t="inlineStr", <is><t>строка</t></is>  (для AWB с ведущим нулём)
function setCellValue(doc, ns, colLetter, rowNum, value) {
  const addr = colLetter + rowNum;
  const isStr = typeof value === "string";
  const sdEl = doc.getElementsByTagName("sheetData")[0];

  // Найти / создать строку
  const allRows = Array.from(doc.getElementsByTagName("row"));
  let row =
    allRows.find((r) => parseInt(r.getAttribute("r"), 10) === rowNum) ||
    null;
  if (!row) {
    row = doc.createElementNS(ns, "row");
    row.setAttribute("r", String(rowNum));
    const after = allRows.find(
      (r) => parseInt(r.getAttribute("r"), 10) > rowNum,
    );
    after ? sdEl.insertBefore(row, after) : sdEl.appendChild(row);
  }

  // Найти / создать ячейку
  const allCells = Array.from(row.getElementsByTagName("c"));
  let cell = allCells.find((c) => c.getAttribute("r") === addr) || null;
  if (!cell) {
    cell = doc.createElementNS(ns, "c");
    cell.setAttribute("r", addr);
    const myIdx = colToIdx(colLetter);
    const after = allCells.find(
      (c) => colToIdx(c.getAttribute("r").replace(/\d+$/, "")) > myIdx,
    );
    after ? row.insertBefore(cell, after) : row.appendChild(cell);
  }

  // Убрать формулу
  Array.from(cell.getElementsByTagName("f")).forEach((f) =>
    cell.removeChild(f),
  );

  if (isStr) {
    cell.setAttribute("t", "inlineStr");
    // Убрать <v> если есть
    Array.from(cell.getElementsByTagName("v")).forEach((v) =>
      cell.removeChild(v),
    );
    // Создать <is><t>...</t></is>
    let is = cell.getElementsByTagName("is")[0];
    if (!is) {
      is = doc.createElementNS(ns, "is");
      cell.appendChild(is);
    }
    let t = is.getElementsByTagName("t")[0];
    if (!t) {
      t = doc.createElementNS(ns, "t");
      is.appendChild(t);
    }
    t.textContent = value;
  } else {
    cell.removeAttribute("t");
    Array.from(cell.getElementsByTagName("is")).forEach((s) =>
      cell.removeChild(s),
    );
    let v = cell.getElementsByTagName("v")[0];
    if (!v) {
      v = doc.createElementNS(ns, "v");
      cell.appendChild(v);
    }
    v.textContent = String(value);
  }
}

// Принудительный полный пересчёт всех формул при открытии файла.
// Без этого Excel на Windows показывает старые закэшированные значения формул,
// пока вручную не кликнешь по ячейке.
function forceFullCalcOnLoad(wbXml) {
  // Уже есть <calcPr ...> — добавляем/заменяем нужные атрибуты
  if (/<calcPr\b/.test(wbXml)) {
    return wbXml.replace(/<calcPr\b[^>]*\/?>/, (tag) => {
      let t = tag.replace(/\s*\/?>$/, ""); // отрезаем закрытие
      // выкидываем старые версии этих атрибутов, чтобы не задвоить
      t = t
        .replace(/\s+fullCalcOnLoad="[^"]*"/g, "")
        .replace(/\s+calcMode="[^"]*"/g, "")
        .replace(/\s+calcId="[^"]*"/g, "");
      return t + ' calcId="0" calcMode="auto" fullCalcOnLoad="1"/>';
    });
  }
  // <calcPr> нет — вставляем его после </sheets> (валидное место по схеме)
  return wbXml.replace(
    /<\/sheets>/,
    '</sheets><calcPr calcId="0" calcMode="auto" fullCalcOnLoad="1"/>',
  );
}

// AWB: последние 4 цифры (строка)
function awb4digits(awb) {
  return String(awb).replace(/\D/g, "").slice(-4);
}
// Как записать в ячейку: если ведущий ноль → строка с точкой, иначе число
function awbDisplayVal(digits4) {
  return digits4.startsWith("0") ? "." + digits4 : parseInt(digits4, 10);
}

// Индекс → буква (0→A, 4→E, ...)
function idxToCol(idx) {
  let n = idx + 1,
    col = "";
  while (n > 0) {
    col = String.fromCharCode(((n - 1) % 26) + 65) + col;
    n = Math.floor((n - 1) / 26);
  }
  return col;
}
// Буква → индекс
function colToIdx(col) {
  let r = 0;
  for (const ch of col) r = r * 26 + (ch.charCodeAt(0) - 64);
  return r - 1;
}
