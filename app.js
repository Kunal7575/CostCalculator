/* =======================
   UofG Cost Calculator - app.js
   - Reads tuition_export_with_compulsory.xlsx in-browser (SheetJS)
   - Uses UG_Tuition for Undergraduate, GR_Tuition for Graduate
   - Matches by: Level, Residency, Province, Load, CohortYear, Program
   - Living: On_campus_Living_Costs, Off_campus_Living_Costs, Meal_Plan
   - Fallback: data.json (optional)

   Primary tuition source:
      FallTuition_Compulsory / WinterTuition_Compulsory / SummerTuition_Compulsory
   Summer toggle works for BOTH UG + GR
   Fallbacks supported:
      (FallTuition + FallCompulsoryFees) etc
      OR FallTotal/WinterTotal/SummerTotal etc
======================= */

let DATA = {
  UG_Tuition: [],
  GR_Tuition: [],
  On_campus_Living_Costs: [],
  Off_campus_Living_Costs: [],
  Meal_Plan: [],
};

let LAST_ESTIMATE = {
  firstName: "",
  email: "",
  level: "",
  residency: "",
  province: "",
  load: "",
  cohort: "",
  program: "",
  includeSummer: false,
  housing: "",
  mealplan: "",
  tuitionFall: 0,
  tuitionWinter: 0,
  tuitionSummer: 0,
  tuitionTotal: 0,
  livingFall: 0,
  livingWinter: 0,
  livingTotal: 0,
  mealTotal: 0,
  grandTotal: 0
};

const $ = (id) => document.getElementById(id);

/* -----------------------
   Helpers
------------------------ */
function cleanText(s) {
  return String(s ?? "").trim();
}

function normalizeToken(s) {
  return cleanText(s).toLowerCase().replace(/[\s\-_]+/g, "");
}

function unique(arr) {
  return [...new Set(arr.filter(Boolean))];
}

function moneyToNumber(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === "number" && Number.isFinite(v)) return v;

  const s = String(v).trim();
  if (!s || s.toUpperCase() === "N/A") return null;

  const normalized = s.replace(/\$/g, "").replace(/,/g, "");
  const n = Number(normalized);
  return Number.isFinite(n) ? n : null;
}

function fmt(n) {
  if (n === null || Number.isNaN(n)) return "N/A";
  return n.toLocaleString("en-CA", { style: "currency", currency: "CAD" });
}

function todayDisplay() {
  return new Date().toLocaleString("en-CA", {
    year: "numeric",
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit"
  });
}

function safeFileNamePart(value, fallback = "Estimate") {
  const cleaned = cleanText(value || fallback).replace(/[^\w\- ]+/g, "").trim();
  return cleaned ? cleaned.replace(/\s+/g, "_") : fallback;
}

function setOptions(selectEl, values, placeholder = null) {
  if (!selectEl) return;
  selectEl.innerHTML = "";

  if (placeholder) {
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = placeholder;
    selectEl.appendChild(opt);
  }

  values.forEach(v => {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = v;
    selectEl.appendChild(opt);
  });
}

/* Robust column getter:
   - tries exact match
   - then tries normalized match
*/
function getCol(row, candidates) {
  if (!row || typeof row !== "object") return undefined;

  const keys = Object.keys(row);
  const keyMap = new Map(keys.map(k => [normalizeToken(k), k]));

  for (const c of candidates) {
    if (c in row) return row[c];
    const nk = normalizeToken(c);
    if (keyMap.has(nk)) return row[keyMap.get(nk)];
  }
  return undefined;
}

/* -----------------------
   Current filters
------------------------ */
function currentFilters() {
  return {
    Level: cleanText($("level")?.value || "Undergraduate"),
    Residency: cleanText($("residency")?.value || "Domestic"),
    Province: cleanText($("province")?.value || "ON"),
    Load: cleanText($("load")?.value || "Full-time"),
    CohortYear: cleanText($("cohort")?.value || ""),
    Program: cleanText($("program")?.value || ""),
  };
}

/* -----------------------
   Tuition source selection
------------------------ */
function getTuitionRowsForLevel(level) {
  return level === "Graduate" ? (DATA.GR_Tuition || []) : (DATA.UG_Tuition || []);
}

/* -----------------------
   Province rules
   - If International => Province forced to INT
------------------------ */
function updateProvinceOptions() {
  const residency = cleanText($("residency")?.value);
  const provinceEl = $("province");
  if (!provinceEl) return;

  const isInternational = residency === "International";
  if (isInternational) {
    setOptions(provinceEl, ["INT"]);
    provinceEl.value = "INT";
    provinceEl.disabled = true;
    return;
  }

  setOptions(provinceEl, ["ON", "Non-ON"]);
  if (provinceEl.value !== "ON" && provinceEl.value !== "Non-ON") provinceEl.value = "ON";
  provinceEl.disabled = false;
}

/* -----------------------
   Summer UI (both UG + GR)
------------------------ */
function setSummerUI() {
  const wrap = $("summerWrap");
  const cb = $("includeSummer");
  if (!wrap || !cb) return;
  wrap.style.display = "block";
}

/* -----------------------
   Cohort options
------------------------ */
function updateCohortOptions() {
  const cohortEl = $("cohort");
  if (!cohortEl) return;

  const f = currentFilters();
  const rows = getTuitionRowsForLevel(f.Level);

  const wantRes = normalizeToken(f.Residency);
  const wantLoad = normalizeToken(f.Load);
  const wantProv = normalizeToken(f.Province);

  const cohorts = unique(
    rows
      .filter(r => normalizeToken(getCol(r, ["Residency"])) === wantRes)
      .filter(r => normalizeToken(getCol(r, ["Load"])) === wantLoad)
      .filter(r => {
        if (f.Level === "Graduate") return true;
        return normalizeToken(getCol(r, ["Province"])) === wantProv;
      })
      .map(r => cleanText(getCol(r, ["CohortYear", "Cohort Year", "Cohort"])))
      .filter(Boolean)
  ).sort();

  const prev = cleanText(cohortEl.value);
  setOptions(cohortEl, cohorts, cohorts.length ? null : "No cohorts found");

  if (prev && cohorts.includes(prev)) cohortEl.value = prev;
  else if (cohorts.length) cohortEl.value = cohorts[0];
  else cohortEl.value = "";
}

/* -----------------------
   Program options
------------------------ */
function updateProgramOptions() {
  const programEl = $("program");
  if (!programEl) return;

  const rows = getFilteredTuitionRows();

  const programs = unique(
    rows.map(r => cleanText(getCol(r, ["Program", "ProgramName", "Program Name"])))
  ).sort();

  const prev = cleanText(programEl.value);
  setOptions(programEl, programs, programs.length ? "Select program" : "No programs found");

  if (prev && programs.includes(prev)) programEl.value = prev;
  else if (programs.length) programEl.value = programs[0];
  else programEl.value = "";
}

/* -----------------------
   Filter tuition rows
------------------------ */
function getFilteredTuitionRows() {
  const f = currentFilters();
  const rows = getTuitionRowsForLevel(f.Level);

  const wantRes = normalizeToken(f.Residency);
  const wantLoad = normalizeToken(f.Load);
  const wantProv = normalizeToken(f.Province);
  const wantCoh = normalizeToken(f.CohortYear);

  return rows.filter(r => {
    const rRes = normalizeToken(getCol(r, ["Residency"]));
    const rLoad = normalizeToken(getCol(r, ["Load"]));
    const rProv = normalizeToken(getCol(r, ["Province"]));
    const rCoh = normalizeToken(getCol(r, ["CohortYear", "Cohort Year", "Cohort"]));

    if (rRes !== wantRes) return false;
    if (rLoad !== wantLoad) return false;

    if (f.Level !== "Graduate") {
      if (rProv !== wantProv) return false;
    } else {
      const hasProvince = cleanText(getCol(r, ["Province"])) !== "";
      if (hasProvince) {
        if (rProv !== wantProv && wantProv !== "nonon") return false;
      }
    }

    if (wantCoh && rCoh !== wantCoh) return false;
    return true;
  });
}

/* -----------------------
   Tuition per term (primary + fallbacks)
------------------------ */
function termTotal(row, term /* "Fall"|"Winter"|"Summer" */) {
  const precomputed = moneyToNumber(getCol(row, [
    `${term}Tuition_Compulsory`,
    `${term}Tuition Compulsory`,
    `${term}TuitionCompulsory`,
    `${term}Tuition_CompulsoryFees`,
  ]));
  if (precomputed !== null) return precomputed;

  const tuition = moneyToNumber(getCol(row, [
    `${term}Tuition`, `${term} Tuition`, `${term}_Tuition`
  ]));
  const compulsory = moneyToNumber(getCol(row, [
    `${term}CompulsoryFees`, `${term} Compulsory Fees`, `${term}_CompulsoryFees`,
    `${term}Compulsory`, `${term} Compulsory`
  ]));
  if (tuition !== null || compulsory !== null) {
    return (tuition ?? 0) + (compulsory ?? 0);
  }

  const total = moneyToNumber(getCol(row, [
    `${term}Total`, `${term} Total`, `${term}TermTotal`,
    `${term}TermTotalCost`, `${term}`
  ]));
  return total ?? 0;
}

/* -----------------------
   Info tooltip system
------------------------ */
const INFO_TEXT = {
  "Tuition & Compulsory fees (Fall)": "Tuition and compulsory fees for the Fall term based on your selections.",
  "Tuition & Compulsory fees (Winter)": "Tuition and compulsory fees for the Winter term based on your selections.",
  "Tuition & Compulsory fees (Summer)": "Only included if you toggle Summer.",
  "Tuition & Compulsory fees (Total)": "Total tuition and compulsory fees for the year.",
  "Living Cost (Fall)": "Estimated housing/living cost for the Fall term.",
  "Living Cost (Winter)": "Estimated housing/living cost for the Winter term.",
  "Living Cost (Total)": "Total living cost for Fall + Winter.",
  "Meal plan (Total)": "Estimated meal plan cost for the year."
};

let __activeTooltip = null;

function closeTooltip() {
  if (__activeTooltip) {
    __activeTooltip.remove();
    __activeTooltip = null;
  }
}

function showTooltip(anchorEl, title, body) {
  closeTooltip();

  const tip = document.createElement("div");
  tip.className = "tooltip";
  tip.innerHTML = `
    <div class="tooltip-title">${title}</div>
    <div class="tooltip-body">${body}</div>
  `;
  document.body.appendChild(tip);

  const r = anchorEl.getBoundingClientRect();
  const padding = 10;

  let left = r.left + 18;
  let top = r.top + 18;

  const tipRect = tip.getBoundingClientRect();
  if (left + tipRect.width > window.innerWidth - padding) {
    left = window.innerWidth - tipRect.width - padding;
  }
  if (top + tipRect.height > window.innerHeight - padding) {
    top = r.top - tipRect.height - 10;
  }
  if (top < padding) top = padding;

  tip.style.left = `${left}px`;
  tip.style.top = `${top}px`;

  __activeTooltip = tip;
}

function infoIconSVG() {
  return `
    <svg viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
      <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 17c-.55 0-1-.45-1-1v-6c0-.55.45-1 1-1s1 .45 1 1v6c0 .55-.45 1-1 1zm0-10c-.83 0-1.5-.67-1.5-1.5S11.17 6 12 6s1.5.67 1.5 1.5S12.83 9 12 9z"/>
    </svg>
  `;
}

function renderBreakdownWithInfo(lines) {
  const table = $("breakdown");
  if (!table) return;

  table.innerHTML = lines.map(([label, val]) => {
    const info = INFO_TEXT[label] || "";
    const leftCell = info
      ? `
        <span class="fee-label">
          <span>${label}</span>
          <button type="button" class="info-btn" data-info="${encodeURIComponent(info)}" data-title="${encodeURIComponent(label)}">
            ${infoIconSVG()}
          </button>
        </span>
      `
      : `<span>${label}</span>`;

    return `
      <tr>
        <td>${leftCell}</td>
        <td>${fmt(val)}</td>
      </tr>
    `;
  }).join("");

  table.querySelectorAll("[data-info]").forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      const title = decodeURIComponent(btn.getAttribute("data-title") || "Info");
      const body = decodeURIComponent(btn.getAttribute("data-info") || "");
      showTooltip(btn, title, body);
    };
  });
}

document.addEventListener("click", (e) => {
  const isInfo = e.target.closest?.("[data-info]");
  const isTooltip = e.target.closest?.(".tooltip");
  if (!isInfo && !isTooltip) closeTooltip();
});

document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") closeTooltip();
});

/* -----------------------
   Living dropdowns
------------------------ */
function updateLivingDropdowns() {
  const onCampusRows = DATA?.On_campus_Living_Costs ?? [];
  const ocRows = onCampusRows
    .map(r => ({
      room: cleanText(getCol(r, ["RoomType", "Room Type"])),
      res: cleanText(getCol(r, ["ResidenceArea", "Residence Area", "Residence"])).toUpperCase(),
      fall: moneyToNumber(getCol(r, ["Fall Term", "FallTerm", "Fall"])),
      winter: moneyToNumber(getCol(r, ["Winter Term", "WinterTerm", "Winter"])),
      total: moneyToNumber(getCol(r, ["Cost", "Total", "TotalCost"])),
    }))
    .filter(x => x.room && x.res);

  const roomTypes = unique(ocRows.map(x => x.room)).sort();
  setOptions($("oncampusRoom"), roomTypes, "Select room type");

  function updateOnCampusResidences() {
    const selectedRoom = $("oncampusRoom")?.value || "";
    const residences = unique(ocRows.filter(x => x.room === selectedRoom).map(x => x.res)).sort();
    setOptions($("oncampusRes"), residences, "Select residence");
    if (residences.length) $("oncampusRes").value = residences[0];
  }

  window.__OC_ROWS__ = ocRows;

  if ($("oncampusRoom")) $("oncampusRoom").onchange = () => {
    updateOnCampusResidences();
    compute();
  };

  if ($("oncampusRes")) $("oncampusRes").onchange = () => compute();

  if (roomTypes.length && $("oncampusRoom")) {
    $("oncampusRoom").value = roomTypes[0];
    updateOnCampusResidences();
  }

  const offCampusRows = DATA?.Off_campus_Living_Costs ?? [];
  const off = offCampusRows
    .map(r => ({
      room: cleanText(getCol(r, ["RoomType", "Room Type", "OffCampusOption", "Off-campus option"])),
      term: cleanText(getCol(r, ["Term"])),
      total: moneyToNumber(getCol(r, ["TotalTermCost", "Total Term Cost", "Total"])),
    }))
    .filter(x => x.room && x.term);

  const offTypes = unique(off.map(x => x.room)).sort();
  setOptions($("offcampus"), offTypes, "Select off-campus option");

  const mealRows = DATA?.Meal_Plan ?? [];
  const mp = mealRows
    .map(r => ({
      name: cleanText(getCol(r, ["Meal Plan Size", "MealPlan", "Meal Plan"])),
      total: moneyToNumber(getCol(r, ["Total cost per year", "TotalCostPerYear", "Total"])),
    }))
    .filter(x => x.name);

  const mealSelect = $("mealplan");
  if (mealSelect) {
    mealSelect.innerHTML = `<option value="None">None</option>`;
    mp.forEach(x => {
      const opt = document.createElement("option");
      opt.value = x.name;
      opt.textContent = x.total !== null ? `${x.name} (${fmt(x.total)}/yr)` : x.name;
      mealSelect.appendChild(opt);
    });
  }

  compute();
}

function toggleLivingInputs() {
  const housing = cleanText($("housing")?.value);

  const ocRoomWrap = $("onCampusRoomWrap");
  const ocResWrap = $("onCampusResWrap");
  const offWrap = $("offCampusWrap");

  const showOnCampus = housing === "OnCampus";
  const showOffCampus = housing === "OffCampus";

  if (ocRoomWrap) ocRoomWrap.style.display = showOnCampus ? "block" : "none";
  if (ocResWrap) ocResWrap.style.display = showOnCampus ? "block" : "none";
  if (offWrap) offWrap.style.display = showOffCampus ? "block" : "none";

  const ocRoom = $("oncampusRoom");
  const ocRes = $("oncampusRes");
  const offSel = $("offcampus");

  if (ocRoom) ocRoom.disabled = !showOnCampus;
  if (ocRes) ocRes.disabled = !showOnCampus;
  if (offSel) offSel.disabled = !showOffCampus;

  compute();
}

/* -----------------------
   Compute totals
------------------------ */
function compute() {
  const f = currentFilters();
  const rows = getFilteredTuitionRows();
  const program = cleanText($("program")?.value);

  const match = rows.find(
    r => cleanText(getCol(r, ["Program", "ProgramName", "Program Name"])) === program
  );

  if (!match) {
    LAST_ESTIMATE = {
      firstName: "",
      email: "",
      level: f.Level,
      residency: f.Residency,
      province: f.Province,
      load: f.Load,
      cohort: f.CohortYear,
      program,
      includeSummer: $("includeSummer")?.checked === true,
      housing: cleanText($("housing")?.value),
      mealplan: cleanText($("mealplan")?.value),
      tuitionFall: 0,
      tuitionWinter: 0,
      tuitionSummer: 0,
      tuitionTotal: 0,
      livingFall: 0,
      livingWinter: 0,
      livingTotal: 0,
      mealTotal: 0,
      grandTotal: 0
    };

    if ($("grandTotal")) $("grandTotal").textContent = "N/A";
    if ($("breakdown")) {
      $("breakdown").innerHTML = `
        <tr><td>Tuition & fees</td><td>N/A</td></tr>
        <tr><td colspan="2" class="muted">No matching tuition row for the selected filters.</td></tr>
      `;
    }
    return;
  }

  const fallTCF = termTotal(match, "Fall");
  const winterTCF = termTotal(match, "Winter");
  const summerTCF = termTotal(match, "Summer");

  const includeSummer = $("includeSummer")?.checked === true;
  const tuitionTotal = fallTCF + winterTCF + (includeSummer ? summerTCF : 0);

  const housing = cleanText($("housing")?.value);
  let livingFall = 0;
  let livingWinter = 0;
  let livingYear = 0;

  if (housing === "OnCampus") {
    const room = $("oncampusRoom")?.value || "";
    const res = (cleanText($("oncampusRes")?.value) || "").toUpperCase();
    const ocRows = window.__OC_ROWS__ || [];
    const matchOC = ocRows.find(x => x.room === room && x.res === res);
    if (matchOC) {
      livingFall = matchOC.fall ?? 0;
      livingWinter = matchOC.winter ?? 0;
      livingYear = matchOC.total ?? (livingFall + livingWinter);
    }
  }

  if (housing === "OffCampus") {
    const selectedType = $("offcampus")?.value || "";
    const offRows = (DATA?.Off_campus_Living_Costs ?? [])
      .map(r => ({
        room: cleanText(getCol(r, ["RoomType", "Room Type", "OffCampusOption", "Off-campus option"])),
        term: cleanText(getCol(r, ["Term"])),
        total: moneyToNumber(getCol(r, ["TotalTermCost", "Total Term Cost", "Total"])),
      }));

    livingFall = offRows.find(x => x.room === selectedType && normalizeToken(x.term) === "fall")?.total ?? 0;
    livingWinter = offRows.find(x => x.room === selectedType && normalizeToken(x.term) === "winter")?.total ?? 0;
    livingYear = livingFall + livingWinter;
  }

  const meal = cleanText($("mealplan")?.value);
  let mealYear = 0;
  if (meal && meal !== "None") {
    const mp = (DATA?.Meal_Plan ?? [])
      .map(r => ({
        name: cleanText(getCol(r, ["Meal Plan Size", "MealPlan", "Meal Plan"])),
        total: moneyToNumber(getCol(r, ["Total cost per year", "TotalCostPerYear", "Total"])),
      }))
      .find(x => x.name === meal);

    if (mp) mealYear = mp.total ?? 0;
  }

  const grand = tuitionTotal + livingYear + mealYear;

  LAST_ESTIMATE = {
    firstName: "",
    email: "",
    level: f.Level,
    residency: f.Residency,
    province: f.Province,
    load: f.Load,
    cohort: f.CohortYear,
    program,
    includeSummer,
    housing,
    mealplan: meal,
    tuitionFall: fallTCF,
    tuitionWinter: winterTCF,
    tuitionSummer: includeSummer ? summerTCF : 0,
    tuitionTotal,
    livingFall,
    livingWinter,
    livingTotal: livingYear,
    mealTotal: mealYear,
    grandTotal: grand
  };

  if ($("grandTotal")) $("grandTotal").textContent = fmt(grand);

  const lines = [
    ["Tuition & Compulsory fees (Fall)", fallTCF],
    ["Tuition & Compulsory fees (Winter)", winterTCF],
    ["Tuition & Compulsory fees (Summer)", includeSummer ? summerTCF : 0],
    ["Tuition & Compulsory fees (Total)", tuitionTotal],
    ["Living Cost (Fall)", livingFall],
    ["Living Cost (Winter)", livingWinter],
    ["Living Cost (Total)", livingYear],
    ["Meal plan (Total)", mealYear],
  ];

  renderBreakdownWithInfo(lines);
}

/* -----------------------
   PDF generation
------------------------ */
/* -----------------------
   PDF generation
------------------------ */
function drawLabelValue(doc, label, value, xLabel, xValue, y) {
  doc.setTextColor(44, 52, 64);
  doc.setFont("helvetica", "bold");
  doc.text(String(label), xLabel, y);
  doc.setFont("helvetica", "normal");
  doc.text(String(value), xValue, y);
}

function drawSectionBox(doc, x, y, w, h, title) {
  doc.setFillColor(250, 250, 250);
  doc.setDrawColor(214, 214, 214);
  doc.roundedRect(x, y, w, h, 4, 4, "FD");

  doc.setTextColor(44, 52, 64);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.text(title, x + 5, y + 8);
}

function finishEstimatePDF(doc) {
  const pageWidth = doc.internal.pageSize.getWidth();
  const left = 16;
  const right = pageWidth - 16;
  const fullWidth = right - left;

  let y = 36;

  // Title row
  doc.setTextColor(44, 52, 64);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(15);
  doc.text("Student Estimate", left, y);

  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);
  doc.setTextColor(90, 90, 90);
  doc.text(`Generated: ${todayDisplay()}`, right, y, { align: "right" });

  y += 12;

  // Grand total
  doc.setFillColor(252, 246, 224);
  doc.setDrawColor(255, 196, 41);
  doc.roundedRect(left, y, fullWidth, 20, 4, 4, "FD");

  doc.setTextColor(44, 52, 64);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text("Estimated Grand Total", left + 5, y + 8);

  doc.setTextColor(0, 0, 0);
  doc.setFontSize(18);
  doc.text(fmt(LAST_ESTIMATE.grandTotal), right - 5, y + 13, { align: "right" });

  y += 30;

  // Selections
  drawSectionBox(doc, left, y, fullWidth, 52, "Selections");
  doc.setFontSize(10);

  drawLabelValue(doc, "Level:", LAST_ESTIMATE.level || "N/A", left + 5, left + 30, y + 18);
  drawLabelValue(doc, "Residency:", LAST_ESTIMATE.residency || "N/A", 106, 134, y + 18);

  drawLabelValue(doc, "Province:", LAST_ESTIMATE.province || "N/A", left + 5, left + 30, y + 28);
  drawLabelValue(doc, "Load:", LAST_ESTIMATE.load || "N/A", 106, 134, y + 28);

  drawLabelValue(doc, "Cohort:", LAST_ESTIMATE.cohort || "N/A", left + 5, left + 30, y + 38);
  drawLabelValue(doc, "Summer:", LAST_ESTIMATE.includeSummer ? "Included" : "Not included", 106, 134, y + 38);

  doc.setTextColor(44, 52, 64);
  doc.setFont("helvetica", "bold");
  doc.text("Program:", left + 5, y + 48);
  doc.setFont("helvetica", "normal");
  const wrappedProgram = doc.splitTextToSize(LAST_ESTIMATE.program || "N/A", fullWidth - 36);
  doc.text(wrappedProgram, left + 30, y + 48);

  y += 66;

  // Breakdown
  const rows = [
    ["Tuition (Fall)", fmt(LAST_ESTIMATE.tuitionFall)],
    ["Tuition (Winter)", fmt(LAST_ESTIMATE.tuitionWinter)],
    ["Tuition (Summer)", fmt(LAST_ESTIMATE.tuitionSummer)],
    ["Tuition Total", fmt(LAST_ESTIMATE.tuitionTotal)],
    ["Living (Fall)", fmt(LAST_ESTIMATE.livingFall)],
    ["Living (Winter)", fmt(LAST_ESTIMATE.livingWinter)],
    ["Living Total", fmt(LAST_ESTIMATE.livingTotal)],
    ["Meal Plan", fmt(LAST_ESTIMATE.mealTotal)]
  ];

  const breakdownHeight = 18 + rows.length * 9;
  drawSectionBox(doc, left, y, fullWidth, breakdownHeight, "Estimate Breakdown");

  let rowY = y + 18;
  rows.forEach((row, index) => {
    if (index > 0) {
      doc.setDrawColor(230, 230, 230);
      doc.line(left + 5, rowY - 5, right - 5, rowY - 5);
    }

    doc.setTextColor(70, 70, 70);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    doc.text(row[0], left + 5, rowY);

    doc.setTextColor(20, 20, 20);
    doc.setFont("helvetica", "bold");
    doc.text(row[1], right - 5, rowY, { align: "right" });

    rowY += 9;
  });

  y += breakdownHeight + 10;

  // Optional selections
  drawSectionBox(doc, left, y, fullWidth, 24, "Optional Selections");
  doc.setFontSize(10);
  drawLabelValue(doc, "Housing:", LAST_ESTIMATE.housing || "None", left + 5, left + 32, y + 16);
  drawLabelValue(doc, "Meal Plan:", LAST_ESTIMATE.mealplan || "None", 106, 138, y + 16);

  y += 34;

  // Footer
  doc.setTextColor(110, 110, 110);
  doc.setFont("helvetica", "italic");
  doc.setFontSize(9);
  const note = "This document is an estimate only and does not replace official tuition, fee, housing, or meal plan information published by the University of Guelph.";
  doc.text(doc.splitTextToSize(note, fullWidth), left, y);

  const fileName = `UofG_Cost_Estimate_${safeFileNamePart(LAST_ESTIMATE.program, "Student")}.pdf`;
  doc.save(fileName);
}

function generateEstimatePDF() {
  compute();

  const { jsPDF } = window.jspdf || {};
  if (!jsPDF) {
    throw new Error("jsPDF is not loaded.");
  }

  const doc = new jsPDF({
    orientation: "portrait",
    unit: "mm",
    format: "a4"
  });

  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const left = 16;

  // Page background
  doc.setFillColor(255, 255, 255);
  doc.rect(0, 0, pageWidth, pageHeight, "F");

  // Header
  doc.setFillColor(229, 25, 55);
  doc.rect(0, 0, pageWidth, 26, "F");

  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(16);
  doc.text("University of Guelph", left, 12);

  doc.setFont("helvetica", "normal");
  doc.setFontSize(11);
  doc.text("Cost Estimate Summary", left, 19);

  // Logo badge area so it looks intentional even if image has white background
  const badgeX = pageWidth - 48;
  const badgeY = 5;
  const badgeW = 30;
  const badgeH = 14;

  doc.setFillColor(255, 255, 255);
  doc.roundedRect(badgeX, badgeY, badgeW, badgeH, 1.5, 1.5, "F");

  const logo = new Image();
  logo.src = "./image.png";

  logo.onload = function () {
    try {
      const imgW = logo.naturalWidth || 1;
      const imgH = logo.naturalHeight || 1;
      const ratio = Math.min((badgeW - 2) / imgW, (badgeH - 2) / imgH);

      const finalW = imgW * ratio;
      const finalH = imgH * ratio;

      const x = badgeX + (badgeW - finalW) / 2;
      const y = badgeY + (badgeH - finalH) / 2;

      doc.addImage(logo, "PNG", x, y, finalW, finalH);
    } catch (e) {
      console.warn("Logo could not be added:", e);
    }

    finishEstimatePDF(doc);
  };

  logo.onerror = function () {
    console.warn("image.png not found or could not load.");
    finishEstimatePDF(doc);
  };
}

/* -----------------------
   Download UI
------------------------ */
function initDownloadUI() {
  const downloadBtn = $("downloadEstimateBtn");

  downloadBtn?.addEventListener("click", () => {
    try {
      generateEstimatePDF();
    } catch (error) {
      console.error("PDF generation failed:", error);
      alert("PDF download failed. Please refresh and try again.");
    }
  });
}

/* -----------------------
   Excel loading (SheetJS)
------------------------ */
function sheetToJson(workbook, sheetName) {
  const ws = workbook.Sheets[sheetName];
  if (!ws) return [];
  return XLSX.utils.sheet_to_json(ws, { defval: "" });
}

function setStatus(text, ok = true) {
  const el = $("dataStatus");
  if (!el) return;
  el.textContent = text;
  el.style.background = ok ? "#f5f5f5" : "#f5f5f5";
  el.style.borderColor = ok ? "rgba(46,125,50,0.25)" : "rgba(198,40,40,0.25)";
  el.style.color = ok ? "#2e7f35" : "#b3142c";
}

async function loadFromJsonFallback() {
  try {
    const res = await fetch("./data.json", { cache: "no-store" });
    const json = await res.json();

    DATA.UG_Tuition = json.UG_Tuition || [];
    DATA.GR_Tuition = json.GR_Tuition || [];

    DATA.On_campus_Living_Costs = json.On_campus_Living_Costs || [];
    DATA.Off_campus_Living_Costs = json.Off_campus_Living_Costs || [];
    DATA.Meal_Plan = json.Meal_Plan || [];

    setStatus("Loaded data.json fallback.", true);
    return true;
  } catch (e) {
    setStatus("Could not load Excel or data.json. Check file paths and run Live Server.", false);
    console.error(e);
    return false;
  }
}

async function loadExcelFile(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });

  const required = ["UG_Tuition", "GR_Tuition", "On_campus_Living_Costs", "Off_campus_Living_Costs", "Meal_Plan"];

  DATA.UG_Tuition = sheetToJson(wb, "UG_Tuition");
  DATA.GR_Tuition = sheetToJson(wb, "GR_Tuition");
  DATA.On_campus_Living_Costs = sheetToJson(wb, "On_campus_Living_Costs");
  DATA.Off_campus_Living_Costs = sheetToJson(wb, "Off_campus_Living_Costs");
  DATA.Meal_Plan = sheetToJson(wb, "Meal_Plan");

  const missing = required.filter(n => !wb.Sheets[n]);
  if (missing.length) {
    setStatus(`Loaded Excel, but missing sheets: ${missing.join(", ")}. Tuition may still work if UG_Tuition/GR_Tuition exist.`, false);
  } else {
    setStatus(`Loaded Excel: ${file.name}`, true);
  }
}

/* -----------------------
   Rebuild UI from loaded DATA
------------------------ */
function rebuildUI() {
  updateProvinceOptions();
  setSummerUI();
  updateCohortOptions();
  updateProgramOptions();
  updateLivingDropdowns();
  toggleLivingInputs();
  compute();
}

/* -----------------------
   Init + Events
------------------------ */
async function init() {
  initDownloadUI();
  initBackToTop();
  await loadFromJsonFallback();
  rebuildUI();

  $("excelFile")?.addEventListener("change", async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      await loadExcelFile(file);
      rebuildUI();
    } catch (err) {
      console.error(err);
      setStatus("Failed to read Excel file. Try saving as .xlsx and re-upload.", false);
    }
  });

  $("reloadBtn")?.addEventListener("click", async () => {
    const file = $("excelFile")?.files?.[0];
    if (file) {
      await loadExcelFile(file);
      rebuildUI();
    } else {
      await loadFromJsonFallback();
      rebuildUI();
    }
  });

  $("level")?.addEventListener("change", () => {
    setSummerUI();
    updateCohortOptions();
    updateProgramOptions();
    compute();
  });

  $("residency")?.addEventListener("change", () => {
    updateProvinceOptions();
    updateCohortOptions();
    updateProgramOptions();
    compute();
  });

  $("province")?.addEventListener("change", () => {
    updateCohortOptions();
    updateProgramOptions();
    compute();
  });

  $("load")?.addEventListener("change", () => {
    updateCohortOptions();
    updateProgramOptions();
    compute();
  });

  $("cohort")?.addEventListener("change", () => {
    updateProgramOptions();
    compute();
  });

  $("program")?.addEventListener("change", compute);
  $("includeSummer")?.addEventListener("change", compute);

  $("housing")?.addEventListener("change", toggleLivingInputs);
  $("offcampus")?.addEventListener("change", compute);
  $("mealplan")?.addEventListener("change", compute);
}

init().catch(err => {
  console.error(err);
  alert("Init failed. Run with Live Server and open DevTools Console for details.");
});

/* -----------------------
   Back to top button
------------------------ */
function initBackToTop() {
  const scrollTopBtn = $("scrollTopBtn");
  if (!scrollTopBtn) return;

  function toggleBackToTop() {
    if (window.scrollY > 200) {
      scrollTopBtn.classList.add("show");
    } else {
      scrollTopBtn.classList.remove("show");
    }
  }

  scrollTopBtn.addEventListener("click", () => {
    window.scrollTo({
      top: 0,
      behavior: "smooth"
    });
  });

  window.addEventListener("scroll", toggleBackToTop);
  toggleBackToTop();
}
