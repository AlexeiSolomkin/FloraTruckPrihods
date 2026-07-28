"use strict";

// Парсер голландского табеля

// Ищем строку с заголовком "№ Клиента", начиная с fromIdx (обе таблицы —
// срезка/зелень сверху и горшки снизу — используют одинаковый заголовок,
// поэтому за одну таблицу отвечает один вызов этой функции)
function findHollandHeaderIdx(rows, fromIdx) {
  for (let i = fromIdx; i < rows.length; i++) {
    if (rows[i][5] === "№ Клиента") return i;
  }
  return -1;
}

// Парсим одну таблицу (строки с данными клиентов до итоговой строки).
// Возвращает распарсенные строки и индекс, на котором остановились
// (сама итоговая строка, либо конец листа) — с него продолжаем поиск
// заголовка следующей таблицы.
function parseHollandTable(rows, startIdx) {
  const tableRows = [];
  let i = startIdx;
  while (i < rows.length) {
    const row = rows[i];
    const tabelClient = row[4], // Клиент
      tabelNumClient = row[5], // № Клиента
      tabelMark = row[6], // Маркировка
      tabelBoxes = row[7], // Кол-во коробок
      tabelCarts = row[8], // Кол-во телег
      tabelComment = row[13]; // Комментарий (Колонка N — индекс 13)

    // итоговая строка (есть только сумма коробок в col[7] и сумма телег в col[8]) —
    // проверяем ДО skip по пустому № клиента, иначе цикл проскочит итоговую
    // строку и дойдёт до заголовка следующей таблицы, приняв его за данные
    if (!tabelClient && (tabelBoxes || tabelCarts)) {
      break;
    }
    if (!tabelNumClient) {
      i++;
      continue;
    }

    tableRows.push({
      tabelClient,
      tabelNumClient,
      tabelMark,
      tabelBoxes,
      tabelCarts,
      tabelComment,
    });
    i++;
  }
  return { tableRows, endIdx: i };
}

// Определяем тип груза по умолчанию для таблицы + переопределение по
// ключевым словам в комментарии (независимо от регистра)
function addHollandRows(clientData, tableRows, defaultType, overrideKeywords, overrideType) {
  for (const row of tableRows) {
    const comment = String(row.tabelComment || "").toLowerCase();
    const type = overrideKeywords.some((kw) => comment.includes(kw))
      ? overrideType
      : defaultType;

    const clientNum = row.tabelNumClient;
    if (!clientData[clientNum]) clientData[clientNum] = [];
    clientData[clientNum].push({
      mark: row.tabelMark,
      boxes: row.tabelBoxes || 0,
      carts: row.tabelCarts || 0,
      comment: row.tabelComment,
      type,
    });
  }
}

function parseTabelHolland(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  // номер машины
  let machineNum = "???";
  for (const row of rows) {
    if (row[0] === "Номер машины" && row[1]) {
      const m = String(row[1]).match(/\d+/);
      machineNum = m ? m[0] : String(row[1]);
      break;
    }
  }

  // ── Таблица №1 (верхняя: срезка/зелень) ──
  const hdr1Idx = findHollandHeaderIdx(rows, 0);
  if (hdr1Idx < 0) throw new Error("Не найден заголовок «№ Клиента» в Табеле");
  const { tableRows: table1Rows, endIdx: end1 } = parseHollandTable(
    rows,
    hdr1Idx + 1,
  );

  // ── Таблица №2 (нижняя: горшки), если есть ──
  const hdr2Idx = findHollandHeaderIdx(rows, end1);
  let table2Rows = [];
  if (hdr2Idx >= 0) {
    ({ tableRows: table2Rows } = parseHollandTable(rows, hdr2Idx + 1));
  }

  const clientData = {};
  // Таблица 1: по умолчанию СРЕЗКА, но если в комментарии есть "зелень" — ЗЕЛЕНЬ
  addHollandRows(clientData, table1Rows, "СРЕЗКА", ["зелень"], "ЗЕЛЕНЬ");
  // Таблица 2: по умолчанию ГОРШКИ, но если в комментарии "срез"/"срезка" — СРЕЗКА
  // (comment.includes("срез") ловит и "срезка", т.к. это подстрока)
  addHollandRows(clientData, table2Rows, "ГОРШКИ", ["срез"], "СРЕЗКА");

  return { machineNum, clientData };
}
