"use strict";

// ═══════════════════════════════════════════════════════════════════════════
//  КНОПКИ ОБРАБОТКИ / СКАЧИВАНИЯ
// ═══════════════════════════════════════════════════════════════════════════
document.getElementById("btnProcess").addEventListener("click", async () => {
  const btn = document.getElementById("btnProcess");
  btn.disabled = true;
  btn.classList.add("btn-spin");
  btn.textContent = "";
  resultBytes = null;
  logClear();
  await new Promise((r) => setTimeout(r, 30));

  try {
    log("=== Парсинг Табелей ===", "l-info");

    let tabelMachineNum = null,
      clientData = null,
      skippedKons = [],
      splitLog = [];
    let hollandMachineNum = null,
      hollandClientData = null;

    if (tabelBytes) {
      const wb = XLSX.read(tabelBytes, {
        type: "array",
        cellFormula: false,
        cellStyles: false,
      });
      ({ tabelMachineNum, clientData, skippedKons, splitLog } = parseTabel(wb));
    }
    if (tabelHollandBytes) {
      const wb = XLSX.read(tabelHollandBytes, {
        type: "array",
        cellFormula: false,
        cellStyles: false,
      });
      ({ machineNum: hollandMachineNum, clientData: hollandClientData } =
        parseTabelHolland(wb));
    }

    // ── Сверка номера машины между двумя табелями (если загружены оба) ──
    if (
      tabelMachineNum &&
      hollandMachineNum &&
      tabelMachineNum !== hollandMachineNum
    ) {
      log(
        `⚠ НОМЕР МАШИНЫ НЕ СОВПАДАЕТ МЕЖДУ ТАБЕЛЯМИ: Импорт = ${tabelMachineNum}, Голландия = ${hollandMachineNum}`,
        "l-warn",
      );
      const ok = confirm(
        `⚠ Номера машины в Табеле импорта и Табеле Голландии не совпадают!\n\n` +
          `Табель импорта:    ${tabelMachineNum}\n` +
          `Табель Голландии:  ${hollandMachineNum}\n\n` +
          `Продолжить заполнение?`,
      );
      if (!ok) {
        log("Операция отменена.", "l-warn");
        return;
      }
    }

    const machineNum = tabelMachineNum || hollandMachineNum;

    // ── Проверка номера машины vs имя файла Прихода ──
    const machineFromPrikhod = parseMachineFromPrikhod(prikhodName);
    if (!machineFromPrikhod) {
      log(
        `⚠ Не удалось определить номер машины из имени файла Прихода`,
        "l-warn",
      );
    } else if (machineFromPrikhod !== machineNum) {
      log(
        `⚠ НОМЕР МАШИНЫ НЕ СОВПАДАЕТ: Табель = ${machineNum}, Приход = ${machineFromPrikhod}`,
        "l-warn",
      );
      const ok = confirm(
        `⚠ Номер машины НЕ СОВПАДАЕТ!\n\n` +
          `В Табеле:  ${machineNum}\n` +
          `В Приходе: ${machineFromPrikhod}\n\n` +
          `Продолжить заполнение?`,
      );
      if (!ok) {
        log("Операция отменена.", "l-warn");
        return;
      }
    } else {
      log(`✓ Номер машины совпадает: ${machineNum}`, "l-ok");
    }

    log(`Машина: ${machineNum}`, "l-info");

    // ── Сводка по обычному Табелю ──
    if (clientData) {
      log(
        `Клиентов в Табеле импорта: ${Object.keys(clientData).length}`,
        "l-info",
      );
      for (const cn of Object.keys(clientData).sort((a, b) => +a - +b)) {
        const ek = clientData[cn].filter((e) => e.type === "ЭКВАДОР").length;
        const kol = clientData[cn].filter((e) => e.type === "КОЛУМБИЯ").length;
        const imp = clientData[cn].filter((e) => e.type === "ИМПОРТ").length;
        const p = [
          ek && `Эквадор×${ek}`,
          kol && `Колумбия×${kol}`,
          imp && `Импорт×${imp}`,
        ].filter(Boolean);
        log(`  Клиент ${cn}: ${p.join(", ")}`, "l-ok");
      }

      if (skippedKons.length) {
        for (const s of skippedKons)
          log(
            `⚠ Пропущена одноклиентская консол. AWB ${s.awb} («${s.clientName}»)`,
            "l-warn",
          );
      }
    }

    // ── Сводка по Табелю Голландии ──
    if (hollandClientData) {
      log(
        `Клиентов в Табеле Голландии: ${Object.keys(hollandClientData).length}`,
        "l-info",
      );
      for (const cn of Object.keys(hollandClientData).sort((a, b) => +a - +b)) {
        const srez = hollandClientData[cn].filter(
          (e) => e.type === "СРЕЗКА",
        ).length;
        const zel = hollandClientData[cn].filter(
          (e) => e.type === "ЗЕЛЕНЬ",
        ).length;
        const gor = hollandClientData[cn].filter(
          (e) => e.type === "ГОРШКИ",
        ).length;
        const p = [
          srez && `Срезка×${srez}`,
          zel && `Зелень×${zel}`,
          gor && `Горшки×${gor}`,
        ].filter(Boolean);
        log(`  Клиент ${cn}: ${p.join(", ")}`, "l-ok");
      }
    }

    log("", "");
    log("=== Заполнение Прихода ===", "l-info");

    let bytes = prikhodBytes;

    // ── Заполнение по обычному Табелю ──
    if (clientData) {
      const res = await fillPrikhod(bytes, clientData, machineNum);
      bytes = res.bytes;

      log("", "");
      log("─── Импорт ────────────────────────────────────────", "l-head");
      for (const m of res.fillLog) log(m, "l-ok");
      for (const w of res.warnings) log("⚠ " + w, "l-warn");

      log("", "");
      log("─── Исключения по весу ─────────────────────────────", "l-head");
      if (res.nettoExcLog.length) {
        for (const e of res.nettoExcLog) {
          if (e.usedBrutto) {
            log(
              `  №${e.num} ${e.name}  AWB …${e.awb.slice(-4)}: нетто отсутствует в табеле, выставлен брутто ${e.weight} кг — ПРОВЕРЬТЕ!`,
              "l-exc-warn",
            );
          } else {
            log(
              `  №${e.num} ${e.name}  AWB …${e.awb.slice(-4)}: нетто ${e.weight} кг ✓`,
              "l-ok",
            );
          }
        }
      } else {
        log("  Нет клиентов с нетто-исключением в этом табеле", "l-info");
      }

      log("─── Исключения по сертификату ───────────────────────", "l-head");
      if (res.certExcLog.length) {
        for (const e of res.certExcLog) {
          const reason = e.reason ? ` (${e.reason})` : "";
          log(
            `  №${e.num} ${e.name}  AWB ${e.awbs}${reason}: серт = 0`,
            "l-ok",
          );
        }
      } else {
        log("  Нет клиентов с исключением по серту в этом табеле", "l-info");
      }

      if (splitLog.length) {
        log("", "");
        log("─── Груз разбит на машины ───────────────────────────", "l-head");
        for (const s of splitLog) {
          if (s.included) {
            log(
              `  ⚠ АВБ ${s.awb}: груз разбит на машины — стоит в этом Приходе (АМ ${s.amNum}) ✓`,
              "l-ok",
            );
          } else {
            log(
              `  ℹ АВБ ${s.awb}: груз разбит на машины — в другой машине (АМ ${s.amNum}), в Приход НЕ включается`,
              "l-warn",
            );
          }
        }
      }
    }

    // ── Заполнение по Табелю Голландии ──
    if (hollandClientData) {
      const resHolland = await fillPrikhodHolland(
        bytes,
        hollandClientData,
        machineNum,
      );
      bytes = resHolland.bytes;

      log("", "");
      log("─── Голландия ────────────────────────────────────────", "l-head");
      for (const m of resHolland.fillLog) log(m, "l-ok");
      for (const w of resHolland.warnings) log("⚠ " + w, "l-warn");
    }

    log("", "");
    log("✅ Готово!", "l-ok");
    resultBytes = bytes;
    document.getElementById("downloadWrap").style.display = "block";
  } catch (e) {
    log("❌ Ошибка: " + e.message, "l-err");
    console.error(e);
  } finally {
    btn.classList.remove("btn-spin");
    btn.textContent = "Сформировать Приход";
    btn.disabled = false;
  }
});

document.getElementById("btnDownload").addEventListener("click", () => {
  if (!resultBytes) return;
  const name = prikhodName.replace(/\.xlsx?$/i, "") + "_ЗАПОЛНЕННЫЙ.xlsx";
  const blob = new Blob([resultBytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
});
