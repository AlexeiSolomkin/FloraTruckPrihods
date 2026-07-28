"use strict";

const SECTIONS_HOLLAND = {
  СРЕЗКА: { markR: 66, bxR: 67, cartsR: 206, label: "ГОЛЛАНДИЯ СТАНДАРТ" },
  ЗЕЛЕНЬ: { markR: 133, bxR: 145, cartsR: 208, label: "ЗЕЛЕНЬ" },
  ГОРШКИ: { markR: 193, bxR: 194, cartsR: 207, label: "ГОРШКИ" }, // данных пока нет, просто задел
};

// Строка "Маркировка" таблицы ТЕЛЕГИ / СС — общая на все три типа груза
// (206/207/208 — это уже строки с числом телег конкретного типа)
const CARTS_MARK_ROW = 205;

// Парсим номер машины из имени файла Прихода: ищем "ам <число>"
function parseMachineFromPrikhod(filename) {
  const m = filename.match(/ам\s+(\d+)/i);
  return m ? m[1] : null;
}

// ═══════════════════════════════════════════════════════════════════════════
//  ЗАПОЛНЕНИЕ ПРИХОДА ДЛЯ ГОЛЛАНДИИ (JSZip + DOMParser — стили не трогаем)
// ═══════════════════════════════════════════════════════════════════════════
async function fillPrikhodHolland(prikhodBytes, clientData, machineNum) {
  const zip = await JSZip.loadAsync(prikhodBytes);

  const wbXml = await zip.file("xl/workbook.xml").async("string");
  const wbRels = await zip.file("xl/_rels/workbook.xml.rels").async("string");
  const sheetMap = parseWorkbook(wbXml, wbRels);

  const fillLog = [],
    warnings = [];
  const parser = new DOMParser(),
    serial = new XMLSerializer();
  let processed = 0;

  // ── Записываем номер машины в C2 листа '!!!!!' ──
  // Все клиентские листы ссылаются на C2 через формулу ='!!!!!'!C2
  const bangPath = sheetMap["!!!!!"];
  if (bangPath) {
    const bangFile = zip.file(bangPath);
    if (bangFile) {
      const bangDoc = parser.parseFromString(
        await bangFile.async("string"),
        "application/xml",
      );
      const bangSd = bangDoc.getElementsByTagName("sheetData")[0];
      if (bangSd) {
        const bangNs =
          bangSd.namespaceURI ||
          "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        setCellValue(bangDoc, bangNs, "C", 2, parseInt(machineNum, 10) || 0);
        let bangXml = serial.serializeToString(bangDoc);
        bangXml = bangXml.replace(
          /^<\?xml[^?]*\?>/,
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        );
        if (!bangXml.startsWith("<?xml"))
          bangXml =
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
            bangXml;
        zip.file(bangPath, bangXml);
      }
    }
  }

  for (const [sheetName, xmlPath] of Object.entries(sheetMap)) {
    const trimmed = sheetName.trim();
    const m = trimmed.match(/^(\d+)/);
    if (!m) continue;
    const clientNum = parseInt(m[1]);
    const clientName = trimmed.replace(/^\d+\)\s*/, "").trim();
    const entries = clientData[clientNum];
    if (!entries || !entries.length) continue;

    const xmlFile = zip.file(xmlPath);
    if (!xmlFile) {
      warnings.push(`Файл листа не найден: ${xmlPath}`);
      continue;
    }

    const doc = parser.parseFromString(
      await xmlFile.async("string"),
      "application/xml",
    );
    if (doc.querySelector("parsererror")) {
      warnings.push(`XML ошибка на листе «${trimmed}»`);
      continue;
    }

    const sdEl = doc.getElementsByTagName("sheetData")[0];
    if (!sdEl) continue;
    const ns =
      sdEl.namespaceURI ||
      "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    // Группируем по типу
    const byType = { СРЕЗКА: [], ЗЕЛЕНЬ: [], ГОРШКИ: [] };
    for (const e of entries) {
      if (byType[e.type]) byType[e.type].push(e);
    }

    const descParts = [];
    let cartsCol = 0; // сквозной счётчик колонок таблицы ТЕЛЕГИ / СС (общий на все типы)

    for (const [typeName, sec] of Object.entries(SECTIONS_HOLLAND)) {
      const ents = byType[typeName];
      if (!ents.length) continue;

      for (let idx = 0; idx < ents.length; idx++) {
        const e = ents[idx];
        const col = idxToCol(4 + idx); // E, F, G, ...

        // маркировка — пишем всегда, колонка занята записью
        setCellValue(doc, ns, col, sec.markR, e.mark);
        // коробки — только если реально есть, иначе оставляем пусто вместо 0
        if (e.boxes) setCellValue(doc, ns, col, sec.bxR, e.boxes);

        // телеги идут в отдельную компактную таблицу ТЕЛЕГИ / СС —
        // маркировка (строка 205, общая) + число (строка типа) в одну и ту же
        // колонку; колонка расходуется, только если телеги реально есть
        if (e.carts) {
          const cartsCode = idxToCol(4 + cartsCol);
          setCellValue(doc, ns, cartsCode, CARTS_MARK_ROW, e.mark);
          setCellValue(doc, ns, cartsCode, sec.cartsR, e.carts);
          cartsCol++;
        }
      }
      descParts.push(`${sec.label}×${ents.length}`);
      if (ents.length > 10)
        warnings.push(
          `Клиент ${clientNum} (${typeName}): кол-во маркировок ${ents.length} — проверьте диапазон формулы`,
        );
    }
    if (cartsCol > 10)
      warnings.push(
        `Клиент ${clientNum}: кол-во телег в таблице ТЕЛЕГИ / СС ${cartsCol} — проверьте диапазон формулы`,
      );

    // Сериализуем XML
    let newXml = serial.serializeToString(doc);
    newXml = newXml.replace(
      /^<\?xml[^?]*\?>/,
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    );
    if (!newXml.startsWith("<?xml"))
      newXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + newXml;
    zip.file(xmlPath, newXml);

    fillLog.push(`✓ «${trimmed}»: ${descParts.join(", ")}`);
    processed++;
  }

  fillLog.unshift(`Обработано листов: ${processed}`);

  // Удаляем calcChain.xml (будет пересоздан Excel'ем)
  if (zip.file("xl/calcChain.xml")) zip.remove("xl/calcChain.xml");

  // Заставляем Excel пересчитать все формулы при открытии
  const patchedWbXml = forceFullCalcOnLoad(wbXml);
  zip.file("xl/workbook.xml", patchedWbXml);

  const outBytes = await zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
  return { bytes: outBytes, fillLog, warnings };
}
