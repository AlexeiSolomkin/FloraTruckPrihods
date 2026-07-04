"use strict";

// ═══════════════════════════════════════════════════════════════════════════
//  ПАРСЕР ТАБЕЛЯ
// ═══════════════════════════════════════════════════════════════════════════
/*
  Столбцы (0-indexed):
    [3]  Дата загрузки
    [5]  № Клиента
    [6]  Клиент
    [8]  Кол-во коробок
   [10]  AWB №
   [11]  Страна
   [12]  Нетто вес
   [13]  Брутто вес

  Жёлтые строки = col[5] === null  (консолидации / миксы)
  Каждой жёлтой с нижней таблицей — своя суб-таблица (по порядку).
*/
function parseTabel(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

  // Номер машины
  let machineNum = "???";
  for (const row of rows) {
    if (row[0] === "Номер машины" && row[1]) {
      const m = String(row[1]).match(/\d+/);
      machineNum = m ? m[0] : String(row[1]);
      break;
    }
  }

  // Заголовок главной таблицы
  let hdrIdx = -1;
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][5] === "№ Клиента") {
      hdrIdx = i;
      break;
    }
  }
  if (hdrIdx < 0)
    throw new Error("Не найден заголовок «№ Клиента» в Табеле");

  // ── Главная таблица ──
  const mainRows = [];
  let i = hdrIdx + 1;
  while (i < rows.length) {
    const row = rows[i];
    const colF = row[5],
      colG = row[6],
      colI = row[8],
      awb = row[10],
      date = row[3];

    if (!awb && !date && colF === null && colG === null) {
      i++;
      continue;
    }
    // Итоговая строка (нет AWB, нет даты, есть сумма коробок в col[8])
    if (
      !awb &&
      !date &&
      typeof colI === "number" &&
      colI > 0 &&
      colF === null
    ) {
      i++;
      break;
    }
    // Начало суб-таблицы
    if (colF === "№ Клиента") break;

    if (awb || date) {
      const cs = String(colG || "")
        .trim()
        .toLowerCase();
      const isKons = cs === "консолидация";
      const isMix = cs.startsWith("микс");
      const isSingle = cs.includes("консолидация") && !isKons;
      const hasSub = (isKons || isMix) && !isSingle;
      mainRows.push({
        clientNum: typeof colF === "number" ? colF : null,
        clientName: String(colG || "").trim(),
        awb: String(awb || "").trim(),
        country: String(row[11] || "")
          .trim()
          .toUpperCase(),
        boxes: typeof colI === "number" ? colI : 0,
        brutto: typeof row[13] === "number" ? row[13] : 0,
        netto: typeof row[12] === "number" ? row[12] : 0, // нетто из главной таблицы
        isYellow: colF === null,
        isKons,
        isMix,
        isSingle,
        hasSub,
      });
    }
    i++;
  }

  // ── Суб-таблицы ──
  // Каждая суб-таблица: { rows: [...], amNum: string|null }
  // amNum берётся из маркера «В ПРИХОДЕ АМ X», стоящего непосредственно
  // ПЕРЕД заголовком «№ Клиента» суб-таблицы (т.е. над ней).
  const subTables = [];
  let curSub = null;
  let curSubAmNum = null; // amNum для текущей собираемой суб-таблицы
  let nextSubAmNum = null; // amNum, прочитанный из маркера, ожидает следующую суб-таблицу
  while (i < rows.length) {
    const row = rows[i];
    const colF = row[5],
      colI = row[8];

    // Маркер «В ПРИХОДЕ АМ X» (жёлтая строка между суб-таблицами)
    if (
      typeof colF === "string" &&
      colF !== "№ Клиента" &&
      /В ПРИХОДЕ АМ/i.test(colF)
    ) {
      const m2 = String(colF).match(/(\d+)/);
      nextSubAmNum = m2 ? m2[1] : null;
      i++;
      continue;
    }

    if (colF === "№ Клиента") {
      if (curSub !== null)
        subTables.push({ rows: curSub, amNum: curSubAmNum });
      curSub = [];
      curSubAmNum = nextSubAmNum; // маркер, стоявший перед этим заголовком
      nextSubAmNum = null;
      i++;
      continue;
    }
    if (curSub !== null && typeof colF === "number" && colF > 0) {
      curSub.push({
        clientNum: colF,
        awb: String(row[10] || "").trim(),
        boxes: typeof colI === "number" ? colI : 0,
        brutto: typeof row[13] === "number" ? row[13] : 0,
        netto: typeof row[12] === "number" ? row[12] : 0, // может быть 0 если пустое
      });
      i++;
      continue;
    }
    if (curSub !== null && colF === null && typeof colI === "number") {
      subTables.push({ rows: curSub, amNum: curSubAmNum });
      curSub = null;
      curSubAmNum = null;
      i++;
      continue;
    }
    i++;
  }
  if (curSub && curSub.length)
    subTables.push({ rows: curSub, amNum: curSubAmNum });

  // ── Карта клиентов ──
  const clientData = {};
  const skippedKons = [];
  const splitLog = []; // записи о разбитых на машины АВБ

  function getType(r) {
    // isKons  = точно "консолидация" (жёлтая строка, есть суб-таблица)
    // isSingle = "консолидация Имя" — одноклиентская, без суб-таблицы
    const isKonsLike = r.isKons || r.isSingle;
    if (isKonsLike && r.country === "UIO") return "ЭКВАДОР";
    if (isKonsLike && r.country === "BOG") return "КОЛУМБИЯ";
    return "ИМПОРТ";
  }

  function addEntry(cn, awb, country, brutto, netto, boxes, type) {
    if (!clientData[cn]) clientData[cn] = [];
    const ex = clientData[cn].find((e) => e.awb === awb);
    if (ex) {
      ex.boxes += boxes;
      ex.brutto += brutto;
      ex.netto += netto;
    } else
      clientData[cn].push({
        awb,
        digits4: awb4digits(awb),
        country,
        brutto,
        netto,
        boxes,
        type,
      });
  }

  // Одиночные клиенты (не жёлтые)
  // Используем getType(r): "консолидация Имя" с UIO/BOG → ЭКВАДОР/КОЛУМБИЯ,
  // остальные → ИМПОРТ
  for (const r of mainRows) {
    if (!r.isYellow && r.clientNum !== null)
      addEntry(
        r.clientNum,
        r.awb,
        r.country,
        r.brutto,
        r.netto,
        r.boxes,
        getType(r),
      );
  }

  // Жёлтые строки с суб-таблицами
  const yellowSub = mainRows.filter((r) => r.isYellow && r.hasSub);
  for (let j = 0; j < yellowSub.length; j++) {
    const mr = yellowSub[j];
    const tt = getType(mr);
    if (j >= subTables.length) continue;
    const sub = subTables[j];

    // Проверяем маркер «В ПРИХОДЕ АМ X»
    if (sub.amNum !== null) {
      if (sub.amNum !== machineNum) {
        // Поставка идёт в другую машину — пропускаем
        splitLog.push({ awb: mr.awb, amNum: sub.amNum, included: false });
        continue;
      } else {
        // Поставка разбита, но числится в ЭТОЙ машине
        splitLog.push({ awb: mr.awb, amNum: sub.amNum, included: true });
      }
    }

    // Группируем по клиенту
    const grouped = {};
    for (const sr of sub.rows) {
      const cn = sr.clientNum;
      if (!grouped[cn])
        grouped[cn] = {
          awb: sr.awb || mr.awb,
          boxes: 0,
          brutto: 0,
          netto: 0,
        };
      grouped[cn].boxes += sr.boxes;
      grouped[cn].brutto += sr.brutto;
      grouped[cn].netto += sr.netto; // суммируем нетто из суб-таблицы (может быть 0)
    }

    for (const [cn, g] of Object.entries(grouped))
      addEntry(
        Number(cn),
        g.awb,
        mr.country,
        g.brutto,
        g.netto,
        g.boxes,
        tt,
      );
  }

  // Одноклиентские консолидации — пропускаем
  for (const r of mainRows.filter((r) => r.isYellow && r.isSingle))
    skippedKons.push({ awb: r.awb, clientName: r.clientName });

  return { machineNum, clientData, skippedKons, splitLog };
}
