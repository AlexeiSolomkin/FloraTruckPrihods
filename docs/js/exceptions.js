"use strict";

// ═══════════════════════════════════════════════════════════════════════════
//  СПИСКИ ИСКЛЮЧЕНИЙ
// ═══════════════════════════════════════════════════════════════════════════

// Клиенты с нетто весом (вместо брутто для ИМПОРТ)
let NETTO_LIST = [
  { num: 48, name: "Алексей Флаворит" },
  { num: 66, name: "Андрей ZNS" },
  { num: 85, name: "Jaaz Flowers" },
  { num: 102, name: "M-FLOWERS MF" },
  { num: 103, name: "Долина Роз" },
  { num: 162, name: "Росцветторг" },
  { num: 194, name: "Юрий (Надежда) FMGZ" },
  { num: 217, name: "Сергей Саранск" },
  { num: 310, name: "Орленко Василий" },
  { num: 330, name: "Миракс Трак" },
];

// Клиенты без сертификата на импорт
let NO_CERT_LIST = [
  { num: 13, name: "Слепов Николай" },
  { num: 22, name: "Механиков Захар" },
  { num: 33, name: "Светлана Тверь" },
  { num: 71, name: "Карен Rodjer" },
  { num: 223, name: "Флорентина" },
  { num: 226, name: "Константин Север" },
  { num: 267, name: "Волков Максим" },
  { num: 310, name: "Орленко Василий" },
  { num: 319, name: "Москалев Андрей" },
  { num: 322, name: "Юрьев Сергей" },
  { num: 388, name: "Хасанова Замиля" },
  { num: 397, name: "Любимов Владимир" },
];

// Спецправило: Альберт (86) в одном авб с Олегом (62) → 86 без серта
const ALBERT_NUM = 86;
const OLEG_NUM = 62;

// Быстрые Set'ы для lookup
let NETTO_SET = new Set(NETTO_LIST.map((e) => e.num));
let NO_CERT_SET = new Set(NO_CERT_LIST.map((e) => e.num));

function getExcName(list, num) {
  return (list.find((e) => e.num === num) || {}).name || `Клиент №${num}`;
}

// ═══════════════════════════════════════════════════════════════════════════
//  UI ИСКЛЮЧЕНИЙ
// ═══════════════════════════════════════════════════════════════════════════
function renderExcLists() {
  // Нетто
  document.getElementById("nettoCount").textContent = NETTO_LIST.length;
  const nl = document.getElementById("nettoList");
  nl.innerHTML = NETTO_LIST.map(
    (e) =>
      `<div class="exc-item"><span class="exc-num">№${e.num}</span><span class="exc-name">${e.name}</span></div>`,
  ).join("");

  // Сертификат
  document.getElementById("certCount").textContent =
    NO_CERT_LIST.length + " + спецправило";
  const cl = document.getElementById("certList");
  const special = cl.querySelector(".exc-special");
  cl.innerHTML = NO_CERT_LIST.map(
    (e) =>
      `<div class="exc-item"><span class="exc-num">№${e.num}</span><span class="exc-name">${e.name}</span></div>`,
  ).join("");
  if (special) cl.appendChild(special);
}

function toggleExcList(type) {
  const el = document.getElementById(type + "List");
  const btn = event.currentTarget;
  const open = el.style.display === "block";
  el.style.display = open ? "none" : "block";
  btn.classList.toggle("active", !open);
}

function addExc(type) {
  const num = parseInt(prompt(`Введите номер клиента:`), 10);
  if (!num || isNaN(num)) return;
  const name =
    prompt(`Введите название клиента (для отображения):`) || `Клиент №${num}`;
  if (type === "netto") {
    if (NETTO_SET.has(num)) {
      alert(`Клиент №${num} уже есть в списке нетто.`);
      return;
    }
    NETTO_LIST.push({ num, name });
    NETTO_SET.add(num);
  } else {
    if (NO_CERT_SET.has(num)) {
      alert(`Клиент №${num} уже есть в списке без сертификата.`);
      return;
    }
    NO_CERT_LIST.push({ num, name });
    NO_CERT_SET.add(num);
  }
  NETTO_LIST.sort((a, b) => a.num - b.num);
  NO_CERT_LIST.sort((a, b) => a.num - b.num);
  renderExcLists();
}

renderExcLists();
