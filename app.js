/* =======================
   UofG Cost Calculator - app.js
   Accessible + cleaned version
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

const DOWNLOAD_FLOW_URL = "https://defaultbe62a12b2cad49a1a5fa85f4f3156a.7d.environment.api.powerplatform.com/powerautomate/automations/direct/workflows/5d659c8475fd4b61a77ccaae4e6cc35f/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=2x7e1hQTmjfH9_Qh5nAD7GCNv0FBdAl3v_aZ0M6HPCc";

const EMAIL_FLOW_URL = "https://defaultbe62a12b2cad49a1a5fa85f4f3156a.7d.environment.api.powerplatform.com/powerautomate/automations/direct/workflows/e9e2e002a99646c1b0f01ff7542b17b4/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=XSSTo_9jUavHCg_UTlvYFUn7QqCWcz6VzUc8oP-Lg20";

/* -----------------------
   Helpers
------------------------ */
function cleanText(value) {
  return String(value ?? "").trim();
}

function normalizeToken(value) {
  return cleanText(value).toLowerCase().replace(/[\s\-_]+/g, "");
}

function unique(arr) {
  return [...new Set(arr.filter(Boolean))];
}

function moneyToNumber(value) {
  if (value === null || value === undefined) return null;
  if (typeof value === "number" && Number.isFinite(value)) return value;

  const str = String(value).trim();
  if (!str || str.toUpperCase() === "N/A") return null;

  const normalized = str.replace(/\$/g, "").replace(/,/g, "");
  const num = Number(normalized);

  return Number.isFinite(num) ? num : null;
}

function fmt(value) {
  if (value === null || value === undefined || Number.isNaN(value)) return "N/A";
  return Number(value).toLocaleString("en-CA", {
    style: "currency",
    currency: "CAD"
  });
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
function syncSummerToggleWithProgram() {
  const programEl = $("program");
  const summerEl = $("includeSummer");

  if (!programEl || !summerEl) return;

  const selectedProgram = cleanText(programEl.value);
  const isSummerOnly = selectedProgram.toLowerCase().includes("summer only");

  summerEl.checked = isSummerOnly;
  summerEl.disabled = isSummerOnly;
  summerEl.setAttribute("aria-disabled", isSummerOnly ? "true" : "false");
}
function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function setOptions(selectEl, values, placeholder = null) {
  if (!selectEl) return;

  const previousValue = cleanText(selectEl.value);
  selectEl.innerHTML = "";

  if (placeholder) {
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = placeholder;
    selectEl.appendChild(opt);
  }

  values.forEach((value) => {
    const opt = document.createElement("option");
    opt.value = value;
    opt.textContent = value;
    selectEl.appendChild(opt);
  });

  if (previousValue && values.includes(previousValue)) {
    selectEl.value = previousValue;
  }
}
function hasValidProgramSelection() {
  const program = cleanText($("program")?.value);
  return !!program && program !== "Select program";
}

function validateEstimateRequirements(showMessage = true) {
  const firstNameEl = $("firstName");
  const emailEl = $("email");
  const programEl = $("program");

  const firstName = cleanText(firstNameEl?.value);
  const email = cleanText(emailEl?.value);
  const hasProgram = hasValidProgramSelection();

  syncEstimateContactFields();

  setFieldErrorState(firstNameEl, false);
  setFieldErrorState(emailEl, false);
  if (programEl) programEl.setAttribute("aria-invalid", "false");

  if (!hasProgram) {
    if (programEl) programEl.setAttribute("aria-invalid", "true");

    if (showMessage) {
      showEmailMessage(
        "Please select a program before downloading or emailing your estimate.",
        true
      );
      programEl?.focus({ preventScroll: true });
    }
    return false;
  }

  const firstNameInvalid = !firstName;
  const emailInvalid = !email || !isValidEmail(email);

  setFieldErrorState(firstNameEl, firstNameInvalid);
  setFieldErrorState(emailEl, emailInvalid);

  if (firstNameInvalid || emailInvalid) {
    if (showMessage) {
      showEmailMessage(
        "Please enter a valid first name and email address before downloading or emailing your estimate.",
        true
      );

      if (firstNameInvalid) firstNameEl?.focus({ preventScroll: true });
      else emailEl?.focus({ preventScroll: true });
    }
    return false;
  }

  hideEmailMessage();
  return true;
}
function getCol(row, candidates) {
  if (!row || typeof row !== "object") return undefined;

  const keys = Object.keys(row);
  const keyMap = new Map(keys.map((k) => [normalizeToken(k), k]));

  for (const candidate of candidates) {
    if (candidate in row) return row[candidate];

    const normalizedCandidate = normalizeToken(candidate);
    if (keyMap.has(normalizedCandidate)) {
      return row[keyMap.get(normalizedCandidate)];
    }
  }

  return undefined;
}

function isValidEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(cleanText(value));
}

function setHiddenState(el, hidden) {
  if (!el) return;

  el.hidden = hidden;
  el.setAttribute("aria-hidden", hidden ? "true" : "false");
  el.style.display = hidden ? "none" : "";
}

function setDisabledState(el, disabled) {
  if (!el) return;

  el.disabled = disabled;
  el.setAttribute("aria-disabled", disabled ? "true" : "false");
}

/* -----------------------
   Contact helpers
------------------------ */
function syncEstimateContactFields() {
  LAST_ESTIMATE.firstName = cleanText($("firstName")?.value);
  LAST_ESTIMATE.email = cleanText($("email")?.value);
}

function setFieldErrorState(inputEl, hasError) {
  if (!inputEl) return;
  inputEl.setAttribute("aria-invalid", hasError ? "true" : "false");
}

function showEmailMessage(text, isError = false) {
  const msg = $("emailMsg");
  if (!msg) return;

  msg.hidden = false;
  msg.style.display = "block";
  msg.textContent = text;

  if (isError) {
    msg.setAttribute("role", "alert");
    msg.setAttribute("aria-live", "assertive");
    msg.style.color = "#b3142c";
  } else {
    msg.setAttribute("role", "status");
    msg.setAttribute("aria-live", "polite");
    msg.style.color = "#2e7f35";
  }
}

function hideEmailMessage() {
  const msg = $("emailMsg");
  if (!msg) return;

  msg.textContent = "";
  msg.hidden = true;
  msg.style.display = "none";
  msg.setAttribute("role", "status");
  msg.setAttribute("aria-live", "polite");
}

function validateContactFields(showMessage = true) {
  const firstNameEl = $("firstName");
  const emailEl = $("email");

  const firstName = cleanText(firstNameEl?.value);
  const email = cleanText(emailEl?.value);

  syncEstimateContactFields();

  const firstNameInvalid = !firstName;
  const emailInvalid = !email || !isValidEmail(email);

  setFieldErrorState(firstNameEl, firstNameInvalid);
  setFieldErrorState(emailEl, emailInvalid);

  if (firstNameInvalid || emailInvalid) {
    if (showMessage) {
      showEmailMessage(
        "Please enter a valid first name and email address before downloading or emailing your estimate.",
        true
      );
    }
    return false;
  }

  hideEmailMessage();
  return true;
}

function saveContactFieldsToLocalStorage() {
  try {
    localStorage.setItem("cc_firstName", cleanText($("firstName")?.value));
    localStorage.setItem("cc_email", cleanText($("email")?.value));
  } catch (error) {
    console.warn("Could not save contact fields to localStorage:", error);
  }
}

function restoreContactFieldsFromLocalStorage() {
  try {
    const savedFirstName = localStorage.getItem("cc_firstName");
    const savedEmail = localStorage.getItem("cc_email");

    if ($("firstName") && savedFirstName) $("firstName").value = savedFirstName;
    if ($("email") && savedEmail) $("email").value = savedEmail;

    syncEstimateContactFields();
  } catch (error) {
    console.warn("Could not restore contact fields from localStorage:", error);
  }
}

function clearContactFields() {
  const firstNameEl = $("firstName");
  const emailEl = $("email");

  if (firstNameEl) firstNameEl.value = "";
  if (emailEl) emailEl.value = "";

  setFieldErrorState(firstNameEl, false);
  setFieldErrorState(emailEl, false);

  try {
    localStorage.removeItem("cc_firstName");
    localStorage.removeItem("cc_email");
  } catch (error) {
    console.warn("Could not clear localStorage:", error);
  }

  syncEstimateContactFields();
  hideEmailMessage();
}

/* -----------------------
   Payload / Power Automate
------------------------ */
function buildEstimatePayload() {
  syncEstimateContactFields();

  return {
    studentName: cleanText(LAST_ESTIMATE.firstName),
    studentEmail: cleanText(LAST_ESTIMATE.email),
    level: cleanText(LAST_ESTIMATE.level),
    residency: cleanText(LAST_ESTIMATE.residency),
    province: cleanText(LAST_ESTIMATE.province),
    load: cleanText(LAST_ESTIMATE.load),
    cohortYear: cleanText(LAST_ESTIMATE.cohort),
    program: cleanText(LAST_ESTIMATE.program),
    housing: cleanText(LAST_ESTIMATE.housing),
    mealPlanName: cleanText(LAST_ESTIMATE.mealplan),

    tuitionFall: Number(LAST_ESTIMATE.tuitionFall || 0),
    tuitionWinter: Number(LAST_ESTIMATE.tuitionWinter || 0),
    tuitionSummer: Number(LAST_ESTIMATE.tuitionSummer || 0),
    tuitionTotal: Number(LAST_ESTIMATE.tuitionTotal || 0),

    livingFall: Number(LAST_ESTIMATE.livingFall || 0),
    livingWinter: Number(LAST_ESTIMATE.livingWinter || 0),
    livingTotal: Number(LAST_ESTIMATE.livingTotal || 0),

    mealPlanTotal: Number(LAST_ESTIMATE.mealTotal || 0),
    grandTotal: Number(LAST_ESTIMATE.grandTotal || 0)
  };
}

async function callPowerAutomateFlow(url, payload) {
  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });

  if (!response.ok) {
    const errorText = await response.text().catch(() => "");
    throw new Error(`Flow call failed: ${response.status} ${response.statusText} ${errorText}`);
  }

  return response;
}

async function logDownloadEstimate() {
  return callPowerAutomateFlow(DOWNLOAD_FLOW_URL, buildEstimatePayload());
}

async function emailEstimateFlow() {
  return callPowerAutomateFlow(EMAIL_FLOW_URL, buildEstimatePayload());
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
------------------------ */
function updateProvinceOptions() {
  const residency = cleanText($("residency")?.value);
  const provinceEl = $("province");
  if (!provinceEl) return;

  const previousValue = cleanText(provinceEl.value);
  const isInternational = residency === "International";

  if (isInternational) {
    provinceEl.innerHTML = `
      <option value="INT">International</option>
    `;
    provinceEl.value = "INT";
    setDisabledState(provinceEl, true);
    return;
  }

  provinceEl.innerHTML = `
    <option value="ON">Ontario</option>
    <option value="Non-ON">Non-Ontario</option>
  `;
  provinceEl.value = (previousValue === "ON" || previousValue === "Non-ON") ? previousValue : "ON";
  setDisabledState(provinceEl, false);
}

/* -----------------------
   Summer UI
------------------------ */
function setSummerUI() {
  const wrap = $("summerWrap");
  const cb = $("includeSummer");
  if (!wrap || !cb) return;

  setHiddenState(wrap, false);
  cb.setAttribute("aria-controls", "breakdown");
}

/* -----------------------
   Cohort options
------------------------ */
function updateCohortOptions() {
  const cohortEl = $("cohort");
  if (!cohortEl) return;

  const filters = currentFilters();
  const rows = getTuitionRowsForLevel(filters.Level);

  const wantRes = normalizeToken(filters.Residency);
  const wantLoad = normalizeToken(filters.Load);
  const wantProv = normalizeToken(filters.Province);

  const cohorts = unique(
    rows
      .filter((row) => normalizeToken(getCol(row, ["Residency"])) === wantRes)
      .filter((row) => normalizeToken(getCol(row, ["Load"])) === wantLoad)
      .filter((row) => {
        if (filters.Level === "Graduate") return true;
        return normalizeToken(getCol(row, ["Province"])) === wantProv;
      })
      .map((row) => cleanText(getCol(row, ["CohortYear", "Cohort Year", "Cohort"])))
      .filter(Boolean)
  ).sort((a, b) => {
    const aEnd = parseInt(String(a).split("-")[1], 10) || 0;
    const bEnd = parseInt(String(b).split("-")[1], 10) || 0;
    return bEnd - aEnd;
  });

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
    rows.map((row) => cleanText(getCol(row, ["Program", "ProgramName", "Program Name"])))
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
  const filters = currentFilters();
  const rows = getTuitionRowsForLevel(filters.Level);

  const wantRes = normalizeToken(filters.Residency);
  const wantLoad = normalizeToken(filters.Load);
  const wantProv = normalizeToken(filters.Province);
  const wantCoh = normalizeToken(filters.CohortYear);

  return rows.filter((row) => {
    const rowRes = normalizeToken(getCol(row, ["Residency"]));
    const rowLoad = normalizeToken(getCol(row, ["Load"]));
    const rowProv = normalizeToken(getCol(row, ["Province"]));
    const rowCoh = normalizeToken(getCol(row, ["CohortYear", "Cohort Year", "Cohort"]));

    if (rowRes !== wantRes) return false;
    if (rowLoad !== wantLoad) return false;

    if (filters.Level !== "Graduate") {
      if (rowProv !== wantProv) return false;
    } else {
      const hasProvince = cleanText(getCol(row, ["Province"])) !== "";
      if (hasProvince) {
        if (rowProv !== wantProv && wantProv !== "nonon") return false;
      }
    }

    if (wantCoh && rowCoh !== wantCoh) return false;
    return true;
  });
}

/* -----------------------
   Tuition per term
------------------------ */
function termTotal(row, term) {
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
   Tooltip system
------------------------ */
const INFO_TEXT = {
  "Tuition & Compulsory fees (Fall)": "Tuition and compulsory fees for the Fall term based on your selections.",
  "Tuition & Compulsory fees (Winter)": "Tuition and compulsory fees for the Winter term based on your selections.",
  "Tuition & Compulsory fees (Summer)": "Only included if you toggle Summer.",
  "Tuition & Compulsory fees (Total)": "Total tuition and compulsory fees for the year.",
  "Living Cost (Fall)": "Estimated housing or living cost for the Fall term.",
  "Living Cost (Winter)": "Estimated housing or living cost for the Winter term.",
  "Living Cost (Total)": "Total living cost for Fall and Winter.",
  "Meal plan (Total)": "Estimated meal plan cost for the year."
};

let activeTooltip = null;
let activeTooltipButton = null;
let tooltipCounter = 0;

function closeTooltip() {
  if (activeTooltipButton) {
    activeTooltipButton.setAttribute("aria-expanded", "false");
    activeTooltipButton.removeAttribute("aria-describedby");
    activeTooltipButton.removeAttribute("aria-controls");
  }

  if (activeTooltip) {
    activeTooltip.remove();
  }

  const announcer = $("sr-tooltip-announcer");
  if (announcer) {
    announcer.textContent = "";
  }

  activeTooltip = null;
  activeTooltipButton = null;
}

function showTooltip(anchorEl, title, body) {
  if (!anchorEl) return;

  if (activeTooltipButton === anchorEl) {
    closeTooltip();
    return;
  }

  closeTooltip();

  const announcer = $("sr-tooltip-announcer");
  if (announcer) {
    announcer.textContent = `${title}. ${body}`;
  }

  tooltipCounter += 1;
  const tooltipId = `breakdown-tooltip-${tooltipCounter}`;

  const tip = document.createElement("div");
  tip.className = "tooltip";
  tip.id = tooltipId;
  tip.setAttribute("role", "tooltip");
  tip.setAttribute("tabindex", "-1");
  tip.innerHTML = `
    <div class="tooltip-title">${escapeHtml(title)}</div>
    <div class="tooltip-body">${escapeHtml(body)}</div>
  `;

  document.body.appendChild(tip);

  const rect = anchorEl.getBoundingClientRect();
  const padding = 10;

  let left = rect.left + 18;
  let top = rect.bottom + 10;

  const tipRect = tip.getBoundingClientRect();

  if (left + tipRect.width > window.innerWidth - padding) {
    left = window.innerWidth - tipRect.width - padding;
  }

  if (top + tipRect.height > window.innerHeight - padding) {
    top = rect.top - tipRect.height - 10;
  }

  if (top < padding) top = padding;
  if (left < padding) left = padding;

  tip.style.left = `${left}px`;
  tip.style.top = `${top}px`;

  anchorEl.setAttribute("aria-expanded", "true");
  anchorEl.setAttribute("aria-describedby", tooltipId);
  anchorEl.setAttribute("aria-controls", tooltipId);

  activeTooltip = tip;
  activeTooltipButton = anchorEl;
}

function infoIconSVG() {
  return `
    <svg viewBox="0 0 24 24" fill="currentColor" aria-hidden="true" focusable="false">
      <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 17c-.55 0-1-.45-1-1v-6c0-.55.45-1 1-1s1 .45 1 1v6c0 .55-.45 1-1 1zm0-10c-.83 0-1.5-.67-1.5-1.5S11.17 6 12 6s1.5.67 1.5 1.5S12.83 9 12 9z"></path>
    </svg>
  `;
}

function renderBreakdownWithInfo(lines) {
  const table = $("breakdown");
  if (!table) return;

  const tbody = document.createElement("tbody");

  lines.forEach(([label, value]) => {
    const tr = document.createElement("tr");

    const th = document.createElement("th");
    th.setAttribute("scope", "row");

    const info = INFO_TEXT[label] || "";
    if (info) {
      const wrap = document.createElement("span");
      wrap.className = "fee-label";

      const labelSpan = document.createElement("span");
      labelSpan.textContent = label;

      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "info-btn";
      btn.setAttribute("data-info", info);
      btn.setAttribute("data-title", label);
      btn.setAttribute("aria-label", `More information about ${label}`);
      btn.setAttribute("title", `More information about ${label}`);
      btn.setAttribute("aria-expanded", "false");
      btn.innerHTML = infoIconSVG();

      wrap.appendChild(labelSpan);
      wrap.appendChild(btn);
      th.appendChild(wrap);
    } else {
      th.textContent = label;
    }

    const td = document.createElement("td");
    td.textContent = fmt(value);

    tr.appendChild(th);
    tr.appendChild(td);
    tbody.appendChild(tr);
  });

  table.innerHTML = "";
  table.appendChild(tbody);

  table.querySelectorAll(".info-btn").forEach((btn) => {
    btn.addEventListener("click", (event) => {
      event.stopPropagation();
      const title = btn.getAttribute("data-title") || "Info";
      const body = btn.getAttribute("data-info") || "";
      showTooltip(btn, title, body);
    });

    btn.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        const title = btn.getAttribute("data-title") || "Info";
        const body = btn.getAttribute("data-info") || "";
        showTooltip(btn, title, body);
      }
    });
  });
}

document.addEventListener("click", (event) => {
  const clickedInfo = event.target.closest?.(".info-btn");
  const clickedTooltip = event.target.closest?.(".tooltip");

  if (!clickedInfo && !clickedTooltip) {
    closeTooltip();
  }
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    closeTooltip();
  }
});

window.addEventListener("resize", closeTooltip);
window.addEventListener("scroll", closeTooltip, true);

/* -----------------------
   Living dropdowns
------------------------ */
function updateLivingDropdowns() {
  const onCampusRows = DATA?.On_campus_Living_Costs ?? [];
  const normalizedOnCampusRows = onCampusRows
    .map((row) => ({
      room: cleanText(getCol(row, ["RoomType", "Room Type"])),
      res: cleanText(getCol(row, ["ResidenceArea", "Residence Area", "Residence"])).toUpperCase(),
      fall: moneyToNumber(getCol(row, ["Fall Term", "FallTerm", "Fall"])),
      winter: moneyToNumber(getCol(row, ["Winter Term", "WinterTerm", "Winter"])),
      total: moneyToNumber(getCol(row, ["Cost", "Total", "TotalCost"])),
    }))
    .filter((item) => item.room && item.res);

  const roomTypes = unique(normalizedOnCampusRows.map((item) => item.room)).sort();
  setOptions($("oncampusRoom"), roomTypes, "Select room type");

  function updateOnCampusResidences() {
    const selectedRoom = cleanText($("oncampusRoom")?.value);
    const residences = unique(
      normalizedOnCampusRows
        .filter((item) => item.room === selectedRoom)
        .map((item) => item.res)
    ).sort();

    setOptions($("oncampusRes"), residences, "Select residence");

    if (residences.length) {
      $("oncampusRes").value = residences[0];
    }
  }

  window.__OC_ROWS__ = normalizedOnCampusRows;

  if ($("oncampusRoom")) {
    $("oncampusRoom").addEventListener("change", () => {
      updateOnCampusResidences();
      compute();
    });
  }

  if ($("oncampusRes")) {
    $("oncampusRes").addEventListener("change", compute);
  }

  if (roomTypes.length && $("oncampusRoom")) {
    $("oncampusRoom").value = roomTypes[0];
    updateOnCampusResidences();
  }

  const offCampusRows = DATA?.Off_campus_Living_Costs ?? [];
  const offCampusOptions = offCampusRows
    .map((row) => ({
      room: cleanText(getCol(row, ["RoomType", "Room Type", "OffCampusOption", "Off-campus option"])),
      term: cleanText(getCol(row, ["Term"])),
      total: moneyToNumber(getCol(row, ["TotalTermCost", "Total Term Cost", "Total"])),
    }))
    .filter((item) => item.room && item.term);

  const offTypes = unique(offCampusOptions.map((item) => item.room)).sort();
  setOptions($("offcampus"), offTypes, "Select off-campus option");

  const mealRows = DATA?.Meal_Plan ?? [];
  const mealPlans = mealRows
    .map((row) => ({
      name: cleanText(getCol(row, ["Meal Plan Size", "MealPlan", "Meal Plan"])),
      total: moneyToNumber(getCol(row, ["Total cost per year", "TotalCostPerYear", "Total"])),
    }))
    .filter((item) => item.name);

  const mealSelect = $("mealplan");
  if (mealSelect) {
    mealSelect.innerHTML = `<option value="None">None</option>`;
    mealPlans.forEach((item) => {
      const opt = document.createElement("option");
      opt.value = item.name;
      opt.textContent = item.total !== null ? `${item.name} (${fmt(item.total)}/yr)` : item.name;
      mealSelect.appendChild(opt);
    });
  }

  compute();
}

function toggleLivingInputs() {
  const housing = cleanText($("housing")?.value);

  const onCampusRoomWrap = $("onCampusRoomWrap");
  const onCampusResWrap = $("onCampusResWrap");
  const offCampusWrap = $("offCampusWrap");

  const showOnCampus = housing === "OnCampus";
  const showOffCampus = housing === "OffCampus";

  setHiddenState(onCampusRoomWrap, !showOnCampus);
  setHiddenState(onCampusResWrap, !showOnCampus);
  setHiddenState(offCampusWrap, !showOffCampus);

  setDisabledState($("oncampusRoom"), !showOnCampus);
  setDisabledState($("oncampusRes"), !showOnCampus);
  setDisabledState($("offcampus"), !showOffCampus);

  compute();
}

/* -----------------------
   Compute totals
------------------------ */
function renderNoMatchBreakdown() {
  const table = $("breakdown");
  if (!table) return;

  table.innerHTML = `
    <tbody>
      <tr>
        <th scope="row">Tuition and fees</th>
        <td>N/A</td>
      </tr>
      <tr>
        <td colspan="2" class="muted">No matching tuition row for the selected filters.</td>
      </tr>
    </tbody>
  `;
}

// function compute() {
//   const filters = currentFilters();
//   const rows = getFilteredTuitionRows();
//   const selectedProgram = cleanText($("program")?.value);

//   const previousFirstName = LAST_ESTIMATE.firstName;
//   const previousEmail = LAST_ESTIMATE.email;

//   const match = rows.find(
//     (row) => cleanText(getCol(row, ["Program", "ProgramName", "Program Name"])) === selectedProgram
//   );

//   if (!match) {
//     LAST_ESTIMATE = {
//       firstName: previousFirstName,
//       email: previousEmail,
//       level: filters.Level,
//       residency: filters.Residency,
//       province: filters.Province,
//       load: filters.Load,
//       cohort: filters.CohortYear,
//       program: selectedProgram,
//       includeSummer: $("includeSummer")?.checked === true,
//       housing: cleanText($("housing")?.value),
//       mealplan: cleanText($("mealplan")?.value),
//       tuitionFall: 0,
//       tuitionWinter: 0,
//       tuitionSummer: 0,
//       tuitionTotal: 0,
//       livingFall: 0,
//       livingWinter: 0,
//       livingTotal: 0,
//       mealTotal: 0,
//       grandTotal: 0
//     };

//     if ($("grandTotal")) $("grandTotal").textContent = "N/A";
//     renderNoMatchBreakdown();
//     return;
//   }

//   const fallTCF = termTotal(match, "Fall");
//   const winterTCF = termTotal(match, "Winter");
//   const summerTCF = termTotal(match, "Summer");

//   const includeSummer = $("includeSummer")?.checked === true;
//   const tuitionTotal = fallTCF + winterTCF + (includeSummer ? summerTCF : 0);

//   const housing = cleanText($("housing")?.value);
//   let livingFall = 0;
//   let livingWinter = 0;
//   let livingYear = 0;

//   if (housing === "OnCampus") {
//     const room = cleanText($("oncampusRoom")?.value);
//     const res = cleanText($("oncampusRes")?.value).toUpperCase();
//     const onCampusRows = window.__OC_ROWS__ || [];

//     const matchOC = onCampusRows.find((item) => item.room === room && item.res === res);
//     if (matchOC) {
//       livingFall = matchOC.fall ?? 0;
//       livingWinter = matchOC.winter ?? 0;
//       livingYear = matchOC.total ?? (livingFall + livingWinter);
//     }
//   }

//   if (housing === "OffCampus") {
//     const selectedType = cleanText($("offcampus")?.value);

//     const offRows = (DATA?.Off_campus_Living_Costs ?? []).map((row) => ({
//       room: cleanText(getCol(row, ["RoomType", "Room Type", "OffCampusOption", "Off-campus option"])),
//       term: cleanText(getCol(row, ["Term"])),
//       total: moneyToNumber(getCol(row, ["TotalTermCost", "Total Term Cost", "Total"])),
//     }));

//     livingFall = offRows.find(
//       (item) => item.room === selectedType && normalizeToken(item.term) === "fall"
//     )?.total ?? 0;

//     livingWinter = offRows.find(
//       (item) => item.room === selectedType && normalizeToken(item.term) === "winter"
//     )?.total ?? 0;

//     livingYear = livingFall + livingWinter;
//   }

//   const meal = cleanText($("mealplan")?.value);
//   let mealYear = 0;

//   if (meal && meal !== "None") {
//     const matchMealPlan = (DATA?.Meal_Plan ?? [])
//       .map((row) => ({
//         name: cleanText(getCol(row, ["Meal Plan Size", "MealPlan", "Meal Plan"])),
//         total: moneyToNumber(getCol(row, ["Total cost per year", "TotalCostPerYear", "Total"])),
//       }))
//       .find((item) => item.name === meal);

//     if (matchMealPlan) {
//       mealYear = matchMealPlan.total ?? 0;
//     }
//   }

//   const grandTotal = tuitionTotal + livingYear + mealYear;

//   LAST_ESTIMATE = {
//     firstName: previousFirstName,
//     email: previousEmail,
//     level: filters.Level,
//     residency: filters.Residency,
//     province: filters.Province,
//     load: filters.Load,
//     cohort: filters.CohortYear,
//     program: selectedProgram,
//     includeSummer,
//     housing,
//     mealplan: meal,
//     tuitionFall: fallTCF,
//     tuitionWinter: winterTCF,
//     tuitionSummer: includeSummer ? summerTCF : 0,
//     tuitionTotal,
//     livingFall,
//     livingWinter,
//     livingTotal: livingYear,
//     mealTotal: mealYear,
//     grandTotal
//   };

//   if ($("grandTotal")) {
//     $("grandTotal").textContent = fmt(grandTotal);
//   }

//   const lines = [
//     ["Tuition & Compulsory fees (Fall)", fallTCF],
//     ["Tuition & Compulsory fees (Winter)", winterTCF],
//     ["Tuition & Compulsory fees (Summer)", includeSummer ? summerTCF : 0],
//     ["Tuition & Compulsory fees (Total)", tuitionTotal],
//     ["Living Cost (Fall)", livingFall],
//     ["Living Cost (Winter)", livingWinter],
//     ["Living Cost (Total)", livingYear],
//     ["Meal plan (Total)", mealYear]
//   ];

//   renderBreakdownWithInfo(lines);
// }
function compute() {
  const filters = currentFilters();
  const rows = getFilteredTuitionRows();
  const selectedProgram = cleanText($("program")?.value);

  if (!selectedProgram || selectedProgram === "Select program") {
    if ($("grandTotal")) $("grandTotal").textContent = "--";

    const table = $("breakdown");
    if (table) {
      table.innerHTML = `
        <tbody>
          <tr>
            <th scope="row">Estimate status</th>
            <td>Select a program to calculate your estimate.</td>
          </tr>
        </tbody>
      `;
    }
    return;
  }

  const previousFirstName = LAST_ESTIMATE.firstName;
  const previousEmail = LAST_ESTIMATE.email;

  const match = rows.find(
    (row) => cleanText(getCol(row, ["Program", "ProgramName", "Program Name"])) === selectedProgram
  );

  if (!match) {
    LAST_ESTIMATE = {
      firstName: previousFirstName,
      email: previousEmail,
      level: filters.Level,
      residency: filters.Residency,
      province: filters.Province,
      load: filters.Load,
      cohort: filters.CohortYear,
      program: selectedProgram,
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
    renderNoMatchBreakdown();
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
    const room = cleanText($("oncampusRoom")?.value);
    const res = cleanText($("oncampusRes")?.value).toUpperCase();
    const onCampusRows = window.__OC_ROWS__ || [];

    const matchOC = onCampusRows.find((item) => item.room === room && item.res === res);
    if (matchOC) {
      livingFall = matchOC.fall ?? 0;
      livingWinter = matchOC.winter ?? 0;
      livingYear = matchOC.total ?? (livingFall + livingWinter);
    }
  }

  if (housing === "OffCampus") {
    const selectedType = cleanText($("offcampus")?.value);

    const offRows = (DATA?.Off_campus_Living_Costs ?? []).map((row) => ({
      room: cleanText(getCol(row, ["RoomType", "Room Type", "OffCampusOption", "Off-campus option"])),
      term: cleanText(getCol(row, ["Term"])),
      total: moneyToNumber(getCol(row, ["TotalTermCost", "Total Term Cost", "Total"])),
    }));

    livingFall = offRows.find(
      (item) => item.room === selectedType && normalizeToken(item.term) === "fall"
    )?.total ?? 0;

    livingWinter = offRows.find(
      (item) => item.room === selectedType && normalizeToken(item.term) === "winter"
    )?.total ?? 0;

    livingYear = livingFall + livingWinter;
  }

  const meal = cleanText($("mealplan")?.value);
  let mealYear = 0;

  if (meal && meal !== "None") {
    const matchMealPlan = (DATA?.Meal_Plan ?? [])
      .map((row) => ({
        name: cleanText(getCol(row, ["Meal Plan Size", "MealPlan", "Meal Plan"])),
        total: moneyToNumber(getCol(row, ["Total cost per year", "TotalCostPerYear", "Total"])),
      }))
      .find((item) => item.name === meal);

    if (matchMealPlan) {
      mealYear = matchMealPlan.total ?? 0;
    }
  }

  const grandTotal = tuitionTotal + livingYear + mealYear;

  LAST_ESTIMATE = {
    firstName: previousFirstName,
    email: previousEmail,
    level: filters.Level,
    residency: filters.Residency,
    province: filters.Province,
    load: filters.Load,
    cohort: filters.CohortYear,
    program: selectedProgram,
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
    grandTotal
  };

  if ($("grandTotal")) {
    $("grandTotal").textContent = fmt(grandTotal);
  }

  const lines = [
    ["Tuition & Compulsory fees (Fall)", fallTCF],
    ["Tuition & Compulsory fees (Winter)", winterTCF],
    ["Tuition & Compulsory fees (Summer)", includeSummer ? summerTCF : 0],
    ["Tuition & Compulsory fees (Total)", tuitionTotal],
    ["Living Cost (Fall)", livingFall],
    ["Living Cost (Winter)", livingWinter],
    ["Living Cost (Total)", livingYear],
    ["Meal plan (Total)", mealYear]
  ];

  renderBreakdownWithInfo(lines);
}
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

  doc.setTextColor(44, 52, 64);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(15);
  doc.text("Student Estimate", left, y);

  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);
  doc.setTextColor(90, 90, 90);
  doc.text(`Generated: ${todayDisplay()}`, right, y, { align: "right" });

  y += 12;

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

  drawSectionBox(doc, left, y, fullWidth, 24, "Optional Selections");
  doc.setFontSize(10);
  drawLabelValue(doc, "Housing:", LAST_ESTIMATE.housing || "None", left + 5, left + 32, y + 16);
  drawLabelValue(doc, "Meal Plan:", LAST_ESTIMATE.mealplan || "None", 106, 138, y + 16);

  y += 34;

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

  doc.setFillColor(255, 255, 255);
  doc.rect(0, 0, pageWidth, pageHeight, "F");

  doc.setFillColor(229, 25, 55);
  doc.rect(0, 0, pageWidth, 26, "F");

  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(16);
  doc.text("University of Guelph", left, 12);

  doc.setFont("helvetica", "normal");
  doc.setFontSize(11);
  doc.text("Cost Estimate Summary", left, 19);

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
    } catch (error) {
      console.warn("Logo could not be added:", error);
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
  if (!downloadBtn) return;

  downloadBtn.addEventListener("click", async () => {
    try {
      compute();

      if (!validateEstimateRequirements(true)) return;

      saveContactFieldsToLocalStorage();
      syncEstimateContactFields();

      try {
        await logDownloadEstimate();
      } catch (flowError) {
        console.error("Download logging failed:", flowError);
      }

      generateEstimatePDF();
    } catch (error) {
      console.error("PDF generation failed:", error);
      showEmailMessage("PDF download failed. Please refresh and try again.", true);
    }
  });
}
/* -----------------------
   Email estimate UI
------------------------ */
function initEmailPrototypeUI() {
  const emailBtn = $("emailEstimateBtn");
  const clearBtn = $("clearEmailBtn");
  const firstNameEl = $("firstName");
  const emailEl = $("email");

  firstNameEl?.addEventListener("input", () => {
    syncEstimateContactFields();
    setFieldErrorState(firstNameEl, false);
    hideEmailMessage();
  });

  emailEl?.addEventListener("input", () => {
    syncEstimateContactFields();
    setFieldErrorState(emailEl, false);
    hideEmailMessage();
  });

  emailBtn?.addEventListener("click", async () => {
    compute();

    if (!validateEstimateRequirements(true)) return;

    saveContactFieldsToLocalStorage();
    syncEstimateContactFields();

    try {
      await emailEstimateFlow();
      showEmailMessage(`Your estimate has been emailed to ${LAST_ESTIMATE.email}.`);
      console.log("Email flow payload:", buildEstimatePayload());
    } catch (error) {
      console.error("Email flow failed:", error);
      showEmailMessage(
        "Sorry, something went wrong while sending your estimate email. Please try again.",
        true
      );
    }
  });

  clearBtn?.addEventListener("click", clearContactFields);

  restoreContactFieldsFromLocalStorage();
}

/* -----------------------
   Excel loading
------------------------ */
function sheetToJson(workbook, sheetName) {
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) return [];
  return XLSX.utils.sheet_to_json(worksheet, { defval: "" });
}

function setStatus(text, ok = true) {
  const el = $("dataStatus");
  if (!el) return;

  el.hidden = false;
  el.textContent = text;
  el.style.background = "#f5f5f5";
  el.style.borderColor = ok ? "rgba(46,125,50,0.25)" : "rgba(198,40,40,0.25)";
  el.style.color = ok ? "#2e7f35" : "#b3142c";
  el.setAttribute("role", ok ? "status" : "alert");
  el.setAttribute("aria-live", ok ? "polite" : "assertive");
}

async function loadFromJsonFallback() {
  try {
    const response = await fetch("./data.json", { cache: "no-store" });
    const json = await response.json();

    DATA.UG_Tuition = json.UG_Tuition || [];
    DATA.GR_Tuition = json.GR_Tuition || [];
    DATA.On_campus_Living_Costs = json.On_campus_Living_Costs || [];
    DATA.Off_campus_Living_Costs = json.Off_campus_Living_Costs || [];
    DATA.Meal_Plan = json.Meal_Plan || [];

    setStatus("Loaded data.json fallback.", true);
    return true;
  } catch (error) {
    setStatus("Could not load Excel or data.json. Check file paths and run Live Server.", false);
    console.error(error);
    return false;
  }
}

async function loadExcelFile(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });

  const requiredSheets = [
    "UG_Tuition",
    "GR_Tuition",
    "On_campus_Living_Costs",
    "Off_campus_Living_Costs",
    "Meal_Plan"
  ];

  DATA.UG_Tuition = sheetToJson(workbook, "UG_Tuition");
  DATA.GR_Tuition = sheetToJson(workbook, "GR_Tuition");
  DATA.On_campus_Living_Costs = sheetToJson(workbook, "On_campus_Living_Costs");
  DATA.Off_campus_Living_Costs = sheetToJson(workbook, "Off_campus_Living_Costs");
  DATA.Meal_Plan = sheetToJson(workbook, "Meal_Plan");

  const missing = requiredSheets.filter((name) => !workbook.Sheets[name]);

  if (missing.length) {
    setStatus(
      `Loaded Excel, but missing sheets: ${missing.join(", ")}. Tuition may still work if UG_Tuition or GR_Tuition exist.`,
      false
    );
  } else {
    setStatus(`Loaded Excel: ${file.name}`, true);
  }
}

/* -----------------------
   Rebuild UI
------------------------ */
function rebuildUI() {
  updateProvinceOptions();
  setSummerUI();
  updateCohortOptions();
  updateProgramOptions();
  syncSummerToggleWithProgram();
  updateLivingDropdowns();
  toggleLivingInputs();
  compute();
}

/* -----------------------
   Back to top
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

/* -----------------------
   Event binding
------------------------ */
function bindFilterEvents() {
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

  $("program")?.addEventListener("change", () => {
    syncSummerToggleWithProgram();
    compute();
  });

  $("includeSummer")?.addEventListener("change", compute);

  $("housing")?.addEventListener("change", toggleLivingInputs);
  $("offcampus")?.addEventListener("change", compute);
  $("mealplan")?.addEventListener("change", compute);
}

function bindOptionalDataControls() {
  $("excelFile")?.addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      await loadExcelFile(file);
      rebuildUI();
    } catch (error) {
      console.error(error);
      setStatus("Failed to read Excel file. Try saving as .xlsx and re-upload.", false);
    }
  });

  $("reloadBtn")?.addEventListener("click", async () => {
    const file = $("excelFile")?.files?.[0];

    if (file) {
      await loadExcelFile(file);
      rebuildUI();
      return;
    }

    await loadFromJsonFallback();
    rebuildUI();
  });
}

/* -----------------------
   Init
------------------------ */
async function init() {
  initDownloadUI();
  initEmailPrototypeUI();
  initBackToTop();
  bindFilterEvents();
  bindOptionalDataControls();

  await loadFromJsonFallback();
  rebuildUI();
}

init().catch((error) => {
  console.error(error);
  showEmailMessage("Initialization failed. Run with Live Server and check the console for details.", true);
});
