"use strict";

// ═══════════════════════════════════════════════════════════════════════════
//  КНОПКИ ОБРАБОТКИ / СКАЧИВАНИЯ
// ═══════════════════════════════════════════════════════════════════════════
document
  .getElementById("btnProcess")
  .addEventListener("click", async () => {
    const btn = document.getElementById("btnProcess");
    btn.disabled = true;
    btn.classList.add("btn-spin");
    btn.textContent = "";
    resultBytes = null;
    logClear();
    await new Promise((r) => setTimeout(r, 30));

    try {
      log("=== Парсинг Табеля ===", "l-info");
      const wb = XLSX.read(tabelBytes, {
        type: "array",
        cellFormula: false,
        cellStyles: false,
      });
      const { machineNum, clientData, skippedKons, splitLog } =
        parseTabel(wb);

      // ── Проверка номера машины ──
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
      log(
        `Клиентов в Табеле: ${Object.keys(clientData).length}`,
        "l-info",
      );

      for (const cn of Object.keys(clientData).sort((a, b) => +a - +b)) {
        const ek = clientData[cn].filter(
          (e) => e.type === "ЭКВАДОР",
        ).length;
        const kol = clientData[cn].filter(
          (e) => e.type === "КОЛУМБИЯ",
        ).length;
        const imp = clientData[cn].filter(
          (e) => e.type === "ИМПОРТ",
        ).length;
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

      log("", "");
      log("=== Заполнение Прихода ===", "l-info");

      const { bytes, fillLog, warnings, nettoExcLog, certExcLog } =
        await fillPrikhod(prikhodBytes, clientData, machineNum);

      for (const m of fillLog) log(m, "l-ok");
      for (const w of warnings) log("⚠ " + w, "l-warn");

      // ── Итоговый блок исключений ──
      log("", "");
      log(
        "─── Исключения по весу ─────────────────────────────",
        "l-head",
      );
      if (nettoExcLog.length) {
        for (const e of nettoExcLog) {
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

      log(
        "─── Исключения по сертификату ───────────────────────",
        "l-head",
      );
      if (certExcLog.length) {
        for (const e of certExcLog) {
          const reason = e.reason ? ` (${e.reason})` : "";
          log(
            `  №${e.num} ${e.name}  AWB ${e.awbs}${reason}: серт = 0`,
            "l-ok",
          );
        }
      } else {
        log(
          "  Нет клиентов с исключением по серту в этом табеле",
          "l-info",
        );
      }

      // ── Разбитые на машины АВБ ──
      if (splitLog.length) {
        log("", "");
        log(
          "─── Груз разбит на машины ───────────────────────────",
          "l-head",
        );
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
