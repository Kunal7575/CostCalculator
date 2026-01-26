let DATA = null;

const $ = (id) => document.getElementById(id);

/* -----------------------
   Helpers
------------------------ */

function cleanText(s) {
  return String(s ?? "").trim();
}

function unique(arr) {
  return [...new Set(arr)];
}

// Finds a value even if the key in JSON has extra spaces like " Rent    "
function getVal(obj, wantedKey) {
  if (!obj || typeof obj !== "object") return undefined;
  const foundKey = Object.keys(obj).find(k => k.trim() === wantedKey);
  return foundKey ? obj[foundKey] : undefined;
}

function moneyToNumber(v) {
  // Handles "3,045.46", "$750", "N/A", null
  if (v === null || v === undefined) return null;
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

function setOptions(selectEl, values, placeholder = null) {
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

/* -----------------------
   Tuition filtering
------------------------ */

function currentFilters() {
  return {
    Residency: $("residency").value,
    Province: $("province").value,
    Load: $("load").value,
    CohortYear: $("cohort").value
  };
}

function getFilteredTuitionRows() {
  const f = currentFilters();
  const source = DATA?.Tuition_Fees ?? [];

  const rows = source.map(r => ({
    ...r,
    Program: cleanText(r.Program),
    Major: cleanText(r.Major),
    Residency: cleanText(r.Residency),
    Province: cleanText(r.Province),
    Load: cleanText(r.Load),
    CohortYear: cleanText(r.CohortYear),
    FallTotalN: moneyToNumber(r.FallTotal),
    WinterTotalN: moneyToNumber(r.WinterTotal),
    FallWinterTotalN: moneyToNumber(r.FallWinterTotal),
  }));

  return rows.filter(r =>
    r.Residency === f.Residency &&
    r.Province === f.Province &&
    r.Load === f.Load &&
    r.CohortYear === f.CohortYear
  );
}

/* -----------------------
   Program/Major UI
------------------------ */

function updateProgramMajor() {
  const rows = getFilteredTuitionRows();
  const programs = unique(rows.map(r => r.Program)).filter(Boolean).sort();

  setOptions($("program"), programs, "Select program");

  // auto-select first program if available
  if (programs.length) $("program").value = programs[0];

  updateMajors();
}

function updateMajors() {
  const rows = getFilteredTuitionRows();
  const program = $("program").value;

  const majors = unique(
    rows.filter(r => r.Program === program).map(r => r.Major)
  ).filter(Boolean).sort();

  setOptions($("major"), majors, "Select major");

  if (majors.length) $("major").value = majors[0];

  compute();
}

/* -----------------------
   Living / Meal dropdowns
------------------------ */

function updateLivingDropdowns() {
  const onCampusRows = DATA?.On_campus_Living_Costs ?? [];
  const offCampusRows = DATA?.Off_campus_Living_Costs ?? [];
  const mealRows = DATA?.Meal_Plan ?? [];

  // ---- On-campus dropdown: "ResidenceArea — RoomType"
  const oc = onCampusRows.map(r => ({
    label: `${cleanText(r.ResidenceArea)} - ${cleanText(r.RoomType)}`,
    fall: moneyToNumber(r["Fall Term"]),
    winter: moneyToNumber(r["Winter Term"]),
    total: moneyToNumber(r.Cost),
    deposit: moneyToNumber(r.Deposit),
  })).filter(x => x.label && x.total !== null);

  setOptions($("oncampus"), oc.map(x => x.label), "Select on-campus option");

  // ---- Off-campus dropdown: unique RoomType (keys might have spaces)
  const off = offCampusRows.map(r => ({
    room: cleanText(getVal(r, "RoomType")),
    term: cleanText(getVal(r, "Term")),
    total: moneyToNumber(getVal(r, "TotalTermCost")),
  }))
  .map(x => ({
    ...x,
    // Trim values like "Fall " -> "Fall"
    term: cleanText(x.term),
    room: cleanText(x.room),
  }))
  .filter(x => x.room && x.term && x.total !== null);

  const offTypes = unique(off.map(x => x.room)).sort();
  setOptions($("offcampus"), offTypes, "Select off-campus option");

  // ---- Meal plan dropdown (Meal_Plan)
  const mp = mealRows.map(r => ({
    name: cleanText(r["Meal Plan Size"]),
    fall: moneyToNumber(r["Semesterly cost (Fall)"]),
    winter: moneyToNumber(r["Semesterly cost (Winter)"]),
    total: moneyToNumber(r["Total cost per year"]),
  })).filter(x => x.name);

  const mealSelect = $("mealplan");
  mealSelect.innerHTML = `<option value="None">None</option>`;
  mp.forEach(x => {
    const opt = document.createElement("option");
    opt.value = x.name;
    opt.textContent = `${x.name} (${fmt(x.total)}/yr)`;
    mealSelect.appendChild(opt);
  });

  compute();
}

/* -----------------------
   Compute totals
------------------------ */

function compute() {
  const rows = getFilteredTuitionRows();
  const program = $("program").value;
  const major = $("major").value;

  const match = rows.find(r => r.Program === program && r.Major === major);

  // If no tuition match, show N/A clearly
  if (!match) {
    $("grandTotal").textContent = "N/A";
    $("breakdown").innerHTML = `
      <tr><td>Tuition & fees (Fall+Winter)</td><td>N/A</td></tr>
      <tr><td colspan="2" class="muted">No matching tuition row for the selected filters.</td></tr>
    `;
    return;
  }

  // Tuition totals
  const fallTuition = match.FallTotalN ?? null;
  const winterTuition = match.WinterTotalN ?? null;

  const tuitionYear = match.FallWinterTotalN ?? (
    (fallTuition !== null && winterTuition !== null) ? (fallTuition + winterTuition) : null
  );

  // Living totals
  const housing = $("housing").value;
  let livingFall = 0, livingWinter = 0, livingYear = 0;

  if (housing === "OnCampus") {
    const selected = $("oncampus").value;

    const oc = (DATA?.On_campus_Living_Costs ?? []).map(r => ({
      label: `${cleanText(r.ResidenceArea)} — ${cleanText(r.RoomType)}`,
      fall: moneyToNumber(r["Fall Term"]),
      winter: moneyToNumber(r["Winter Term"]),
      total: moneyToNumber(r.Cost),
    })).find(x => x.label === selected);

    if (oc) {
      livingFall = oc.fall ?? 0;
      livingWinter = oc.winter ?? 0;
      livingYear = oc.total ?? (livingFall + livingWinter);
    }
  }

  if (housing === "OffCampus") {
    const selectedType = $("offcampus").value;

    const off = (DATA?.Off_campus_Living_Costs ?? []).map(r => ({
      room: cleanText(getVal(r, "RoomType")),
      term: cleanText(getVal(r, "Term")),
      total: moneyToNumber(getVal(r, "TotalTermCost")),
    }))
    .map(x => ({
      ...x,
      term: cleanText(x.term),
      room: cleanText(x.room),
    }));

    const fall = off.find(x => x.room === selectedType && x.term === "Fall")?.total ?? 0;
    const winter = off.find(x => x.room === selectedType && x.term === "Winter")?.total ?? 0;

    livingFall = fall;
    livingWinter = winter;
    livingYear = fall + winter;
  }

  // Meal plan totals
  const meal = $("mealplan").value;
  let mealFall = 0, mealWinter = 0, mealYear = 0;

  if (meal !== "None") {
    const mp = (DATA?.Meal_Plan ?? []).map(r => ({
      name: cleanText(r["Meal Plan Size"]),
      fall: moneyToNumber(r["Semesterly cost (Fall)"]),
      winter: moneyToNumber(r["Semesterly cost (Winter)"]),
      total: moneyToNumber(r["Total cost per year"]),
    })).find(x => x.name === meal);

    if (mp) {
      mealFall = mp.fall ?? 0;
      mealWinter = mp.winter ?? 0;
      mealYear = mp.total ?? (mealFall + mealWinter);
    }
  }

  // Grand total
  const grand = (tuitionYear ?? 0) + livingYear + mealYear;
  $("grandTotal").textContent = fmt(grand);

  // Breakdown table (keep 0 as $0.00, not N/A)
  const lines = [
    ["Tuition & fees (Fall+Winter)", tuitionYear],
    ["Living (Fall)", livingFall],
    ["Living (Winter)", livingWinter],
    ["Living (Total)", livingYear],
    ["Meal plan (Total)", mealYear],
  ];

  $("breakdown").innerHTML = lines.map(([label, val]) => `
    <tr>
      <td>${label}</td>
      <td>${fmt(val)}</td>
    </tr>
  `).join("");
}

/* -----------------------
   UI toggles / syncing
------------------------ */

function toggleLivingInputs() {
  const housing = $("housing").value;
  $("oncampus").disabled = housing !== "OnCampus";
  $("offcampus").disabled = housing !== "OffCampus";
  compute();
}

// Prevent impossible filter combos
function syncProvinceToResidency() {
  const residency = $("residency").value;

  if (residency === "International") {
    $("province").value = "INT";
  } else {
    // Domestic
    if ($("province").value === "INT") $("province").value = "ON";
  }
}

/* -----------------------
   Init
------------------------ */

async function init() {
  const res = await fetch("./data.json");
  DATA = await res.json();

  // Living/meal dropdowns
  updateLivingDropdowns();

  // Ensure residency/province are compatible before populating tuition-driven dropdowns
  syncProvinceToResidency();

  // Populate programs/majors
  updateProgramMajor();

  // Event listeners
  ["residency", "province", "load", "cohort"].forEach(id =>
    $(id).addEventListener("change", () => {
      syncProvinceToResidency();
      updateProgramMajor();
    })
  );

  $("program").addEventListener("change", updateMajors);
  $("major").addEventListener("change", compute);

  $("housing").addEventListener("change", toggleLivingInputs);
  $("oncampus").addEventListener("change", compute);
  $("offcampus").addEventListener("change", compute);
  $("mealplan").addEventListener("change", compute);

  toggleLivingInputs();
}

init().catch(err => {
  console.error(err);
  alert("Failed to load data.json. Run a local server (Live Server) and try again.");
});
