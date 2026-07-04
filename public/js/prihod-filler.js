"use strict";

// ═══════════════════════════════════════════════════════════════════════════
//  ЗАПОЛНЕНИЕ ПРИХОДА (JSZip + DOMParser — стили не трогаем)
// ═══════════════════════════════════════════════════════════════════════════
const SKIP_SHEETS = new Set(["тарифы", "!!!!!", "TOTAL", "образец"]);
const SECTIONS = {
  ЭКВАДОР: { awbR: 35, wtR: 36, bxR: 37, label: "КОНСОЛЬ ЭКВАДОР" },
  КОЛУМБИЯ: { awbR: 42, wtR: 43, bxR: 45, label: "КОНСОЛЬ КОЛУМБИЯ" },
  ИМПОРТ: { awbR: 50, wtR: 51, bxR: 52, label: "ИМПОРТ" },
};

// Парсим номер машины из имени файла Прихода: ищем "ам <число>"
function parseMachineFromPrikhod(filename) {
  const m = filename.match(/ам\s+(\d+)/i);
  return m ? m[1] : null;
}

async function fillPrikhod(prikhodBytes, clientData, machineNum) {
  const zip = await JSZip.loadAsync(prikhodBytes);

  const wbXml = await zip.file("xl/workbook.xml").async("string");
  const wbRels = await zip
    .file("xl/_rels/workbook.xml.rels")
    .async("string");
  const sheetMap = parseWorkbook(wbXml, wbRels);

  const fillLog = [],
    warnings = [];
  const nettoExcLog = [],
    certExcLog = [];
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
        setCellValue(
          bangDoc,
          bangNs,
          "C",
          2,
          parseInt(machineNum, 10) || 0,
        );
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
    if (SKIP_SHEETS.has(trimmed)) continue;
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
    const byType = { ЭКВАДОР: [], КОЛУМБИЯ: [], ИМПОРТ: [] };
    for (const e of entries) {
      if (byType[e.type]) byType[e.type].push(e);
    }

    const descParts = [];

    for (const [typeName, sec] of Object.entries(SECTIONS)) {
      const ents = byType[typeName];
      if (!ents.length) continue;

      for (let idx = 0; idx < ents.length; idx++) {
        const e = ents[idx];
        const col = idxToCol(4 + idx); // E, F, G, ...

        // AWB (с учётом ведущего нуля)
        setCellValue(doc, ns, col, sec.awbR, awbDisplayVal(e.digits4));

        // Вес: для ИМПОРТ-исключений пробуем нетто
        let weight;
        if (NETTO_SET.has(clientNum) && typeName === "ИМПОРТ") {
          if (e.netto > 0) {
            weight = e.netto;
            nettoExcLog.push({
              num: clientNum,
              name: clientName,
              awb: e.awb,
              weight,
              usedBrutto: false,
            });
          } else {
            weight = e.brutto;
            nettoExcLog.push({
              num: clientNum,
              name: clientName,
              awb: e.awb,
              weight,
              usedBrutto: true,
            });
          }
        } else {
          weight = e.brutto;
        }
        setCellValue(doc, ns, col, sec.wtR, weight);
        setCellValue(doc, ns, col, sec.bxR, e.boxes);
      }
      descParts.push(`${sec.label}×${ents.length}`);
      if (ents.length > 10)
        warnings.push(
          `Клиент ${clientNum} (${typeName}): ${ents.length} AWB — проверьте диапазон формулы`,
        );
    }

    // Сертификат → B29
    const certCount = calcCertCount(clientNum, byType, clientData);
    setCellValue(doc, ns, "B", 29, certCount);

    // Логируем серт-исключения
    if (byType["ИМПОРТ"].length > 0 && certCount === 0) {
      const awbs = byType["ИМПОРТ"]
        .map((e) => "…" + e.digits4)
        .join(", ");
      if (NO_CERT_SET.has(clientNum)) {
        certExcLog.push({
          num: clientNum,
          name: clientName,
          reason: null,
          awbs,
        });
      } else if (clientNum === ALBERT_NUM) {
        certExcLog.push({
          num: clientNum,
          name: clientName,
          reason: `в одном авб с Олегом №${OLEG_NUM}`,
          awbs,
        });
      }
    }

    // Сериализуем XML
    let newXml = serial.serializeToString(doc);
    newXml = newXml.replace(
      /^<\?xml[^?]*\?>/,
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    );
    if (!newXml.startsWith("<?xml"))
      newXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
        newXml;
    zip.file(xmlPath, newXml);

    fillLog.push(
      `✓ «${trimmed}»: ${descParts.join(", ")}, серт=${certCount}`,
    );
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
  return { bytes: outBytes, fillLog, warnings, nettoExcLog, certExcLog };
}

// Расчёт сертификата с учётом исключений
function calcCertCount(clientNum, byType, clientData) {
  if (NO_CERT_SET.has(clientNum)) return 0;

  const importEnts = byType["ИМПОРТ"];
  if (!importEnts.length) return 0;

  // Спецправило: Альберт (86) в одном авб с Олегом (62) → 0 сертов у Альберта
  if (clientNum === ALBERT_NUM) {
    const olegEnts = clientData[OLEG_NUM];
    if (olegEnts) {
      const olegAwbs = new Set(
        olegEnts.filter((e) => e.type === "ИМПОРТ").map((e) => e.awb),
      );
      if (importEnts.some((e) => olegAwbs.has(e.awb))) return 0;
    }
  }

  return importEnts.length;
}
