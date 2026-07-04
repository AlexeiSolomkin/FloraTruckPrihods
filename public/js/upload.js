"use strict";

// ═══════════════════════════════════════════════════════════════════════════
//  ЗАГРУЗКА ФАЙЛОВ
// ═══════════════════════════════════════════════════════════════════════════
let tabelBytes = null;
let prikhodBytes = null;
let prikhodName = "";
let resultBytes = null;

function setupUpload(inputId, areaId, fnameId, onSuccess) {
  const input = document.getElementById(inputId);
  const area = document.getElementById(areaId);
  const fnEl = document.getElementById(fnameId);
  async function handle(file) {
    if (!file) return;
    try {
      const b = await readBytes(file);
      fnEl.textContent = file.name;
      area.classList.add("loaded");
      onSuccess(b, file.name);
    } catch (e) {
      alert("Ошибка чтения: " + e.message);
    }
  }
  input.addEventListener("change", (e) => handle(e.target.files[0]));
  area.addEventListener("dragover", (e) => {
    e.preventDefault();
    area.classList.add("drag-over");
  });
  area.addEventListener("dragleave", () =>
    area.classList.remove("drag-over"),
  );
  area.addEventListener("drop", (e) => {
    e.preventDefault();
    area.classList.remove("drag-over");
    handle(e.dataTransfer.files[0]);
  });
}

function readBytes(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = (e) => res(new Uint8Array(e.target.result));
    r.onerror = () => rej(new Error("Ошибка FileReader"));
    r.readAsArrayBuffer(file);
  });
}

setupUpload("fileT", "areaT", "fnameT", (b) => {
  tabelBytes = b;
  checkReady();
});
setupUpload("fileP", "areaP", "fnameP", (b, n) => {
  prikhodBytes = b;
  prikhodName = n;
  checkReady();
});

function checkReady() {
  document.getElementById("btnProcess").disabled = !(
    tabelBytes && prikhodBytes
  );
}

// ═══════════════════════════════════════════════════════════════════════════
//  ЛОГ
// ═══════════════════════════════════════════════════════════════════════════
function logClear() {
  document.getElementById("log").innerHTML = "";
  document.getElementById("logWrap").style.display = "block";
  document.getElementById("downloadWrap").style.display = "none";
}
function log(text, cls = "") {
  const el = document.getElementById("log");
  const d = document.createElement("div");
  if (cls) d.className = cls;
  d.textContent = text;
  el.appendChild(d);
  el.scrollTop = el.scrollHeight;
}
