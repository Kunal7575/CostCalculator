let DATA = null;

const $ = (id) => document.getElementById(id);

/* -----------------------
   Helpers
------------------------ */
function cleanText(s) {
  return String(s ?? "").trim();
}

function normalizeResidence(s) {
  return cleanText(s).toUpperCase(); // keeps matching consistent
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

/* -----------------------
   Label builders (must match everywhere)
------------------------ */
function onCampusLabel(row) {
  // Keeping this in case you still use it elsewhere, but we no longer build a giant on-campus dropdown.
  return `${cleanText(row.ResidenceArea)} - ${cleanText(row.RoomType)}`;
}

/* -----------------------
   Current selections / filters
------------------------ */
function getLevel() {
  // Requires <select id="level"> in HTML
  return cleanText($("level")?.value || "Undergraduate");
}

function currentFilters() {
  return {
    Level: getLevel(),
    Residency: $("residency").value,
    Province: $("province").value,
    Load: $("load").value,
    CohortYear: $("cohort").value
  };
}

/* -----------------------
   Tuition sources (UG + Grad)
------------------------ */
function getTuitionRows() {
  const f = currentFilters();

  if (f.Level === "Graduate") {
    const src = DATA?.Tuition_Fees_Graduate ?? [];
    return src.map(r => ({
      Level: "Graduate",
      Program: cleanText(r.Program),
      // Grad sheet appears to have no Major and no Province
      Major: "",
      Residency: cleanText(r.Residency),
      Province: "",
      Load: cleanText(r.Load),
      CohortYear: cleanText(r.CohortYear),

      FallTotalN: moneyToNumber(r.FallTotal),
      WinterTotalN: moneyToNumber(r.WinterTotal),
      SummerTotalN: moneyToNumber(r.SummerTotal),

      FallWinterTotalN: moneyToNumber(r.FallWinterTotal),
    }));
  }

  // Undergraduate
  const src = DATA?.Tuition_Fees ?? [];
  return src.map(r => ({
    Level: "Undergraduate",
    Program: cleanText(r.Program),
    Major: cleanText(r.Major),
    Residency: cleanText(r.Residency),
    Province: cleanText(r.Province),
    Load: cleanText(r.Load),
    CohortYear: cleanText(r.CohortYear),

    FallTotalN: moneyToNumber(r.FallTotal),
    WinterTotalN: moneyToNumber(r.WinterTotal),
    SummerTotalN: null,
    FallWinterTotalN: moneyToNumber(r.FallWinterTotal),
  }));
}

function getFilteredTuitionRows() {
  const f = currentFilters();
  const rows = getTuitionRows();

  if (f.Level === "Graduate") {
    // Grad: no Province, no Major. Filter by Residency + Load + CohortYear.
    return rows.filter(r =>
      r.Level === "Graduate" &&
      r.Residency === f.Residency &&
      r.Load === f.Load &&
      r.CohortYear === f.CohortYear
    );
  }

  // Undergrad: includes province
  return rows.filter(r =>
    r.Level === "Undergraduate" &&
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
  const level = getLevel();
  const rows = getFilteredTuitionRows();

  // Programs
  const programs = unique(rows.map(r => r.Program)).filter(Boolean).sort();
  setOptions($("program"), programs, "Select program");
  if (programs.length) $("program").value = programs[0];

  // Major behaviour depends on level
  if (level === "Graduate") {
    // Disable major in grad mode (since data has no Major)
    if ($("major")) {
      setOptions($("major"), [], "N/A for Graduate");
      $("major").value = "";
      $("major").disabled = true;
    }
    compute();
    return;
  }

  // Undergrad: enable + populate majors
  if ($("major")) $("major").disabled = false;
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
  const offCampusRows = DATA?.Off_campus_Living_Costs ?? [];
  const mealRows = DATA?.Meal_Plan ?? [];

  // ----- On-campus: RoomType -> ResidenceArea -----
  const onCampusRows = DATA?.On_campus_Living_Costs ?? [];

  // Build clean rows
  const ocRows = onCampusRows.map(r => ({
    room: cleanText(r.RoomType),
    res: normalizeResidence(r.ResidenceArea),
    fall: moneyToNumber(r["Fall Term"]),
    winter: moneyToNumber(r["Winter Term"]),
    total: moneyToNumber(r.Cost),
  })).filter(x => x.room && x.res && x.total !== null);

  // Populate RoomType dropdown
  const roomTypes = unique(ocRows.map(x => x.room)).sort();
  setOptions($("oncampusRoom"), roomTypes, "Select room type");

  // When room type changes, populate residences
  function updateOnCampusResidences() {
    const selectedRoom = $("oncampusRoom")?.value || "";
    const residences = unique(
      ocRows.filter(x => x.room === selectedRoom).map(x => x.res)
    ).sort();

    setOptions($("oncampusRes"), residences, "Select residence");
    if (residences.length) $("oncampusRes").value = residences[0];
  }

  // Store for compute()
  window.__OC_ROWS__ = ocRows;

  // Hook change events once (overwrite any existing handler)
  if ($("oncampusRoom")) {
    $("oncampusRoom").onchange = () => { updateOnCampusResidences(); compute(); };
  }
  if ($("oncampusRes")) {
    $("oncampusRes").onchange = () => compute();
  }

  // Initialize residences if any
  if (roomTypes.length && $("oncampusRoom")) {
    $("oncampusRoom").value = roomTypes[0];
    updateOnCampusResidences();
  }

  // ----- Off-campus options (keys may have spaces) -----
  const off = offCampusRows.map(r => ({
    room: cleanText(getVal(r, "RoomType")),
    term: cleanText(getVal(r, "Term")),
    total: moneyToNumber(getVal(r, "TotalTermCost")),
  }))
  .map(x => ({ ...x, room: cleanText(x.room), term: cleanText(x.term) }))
  .filter(x => x.room && x.term && x.total !== null);

  const offTypes = unique(off.map(x => x.room)).sort();
  setOptions($("offcampus"), offTypes, "Select off-campus option");

  // ----- Meal plans -----
  const mp = mealRows.map(r => ({
    name: cleanText(r["Meal Plan Size"]),
    total: moneyToNumber(r["Total cost per year"]),
  })).filter(x => x.name);

  const mealSelect = $("mealplan");
  if (mealSelect) {
    mealSelect.innerHTML = `<option value="None">None</option>`;
    mp.forEach(x => {
      const opt = document.createElement("option");
      opt.value = x.name;
      opt.textContent = `${x.name} (${fmt(x.total)}/yr)`;
      mealSelect.appendChild(opt);
    });
  }

  compute();
}

/* -----------------------
   Compute totals
------------------------ */
function compute() {
  const level = getLevel();
  const rows = getFilteredTuitionRows();
  const program = $("program").value;
  const major = $("major") ? $("major").value : "";

  // Find tuition match
  let match = null;
  if (level === "Graduate") {
    match = rows.find(r => r.Program === program);
  } else {
    match = rows.find(r => r.Program === program && r.Major === major);
  }

  if (!match) {
    $("grandTotal").textContent = "N/A";
    $("breakdown").innerHTML = `
      <tr><td>Tuition & fees</td><td>N/A</td></tr>
      <tr><td colspan="2" class="muted">No matching tuition row for the selected filters.</td></tr>
    `;
    return;
  }

  // Tuition totals
  const fall = match.FallTotalN ?? 0;
  const winter = match.WinterTotalN ?? 0;

  let tuitionTotal = 0;

  if (level === "Graduate") {
    // Include Summer if present
    const summer = match.SummerTotalN ?? 0;

    // Prefer explicit total if present, otherwise compute (Fall+Winter) and add Summer
    const fallWinter = match.FallWinterTotalN ?? (fall + winter);
    tuitionTotal = fallWinter + summer;
  } else {
    tuitionTotal = match.FallWinterTotalN ?? (fall + winter);
  }

  // Living totals
  const housing = $("housing").value;
  let livingFall = 0, livingWinter = 0, livingYear = 0;

  if (housing === "OnCampus") {
    const room = $("oncampusRoom")?.value || "";
    const res = normalizeResidence($("oncampusRes")?.value || "");

    const ocRows = window.__OC_ROWS__ || [];
    const matchOC = ocRows.find(x => x.room === room && x.res === res);

    if (matchOC) {
      livingFall = matchOC.fall ?? 0;
      livingWinter = matchOC.winter ?? 0;
      livingYear = matchOC.total ?? (livingFall + livingWinter);
    }
  }

  if (housing === "OffCampus") {
    const selectedType = $("offcampus").value;

    const off = (DATA?.Off_campus_Living_Costs ?? []).map(r => ({
      room: cleanText(getVal(r, "RoomType")),
      term: cleanText(getVal(r, "Term")),
      total: moneyToNumber(getVal(r, "TotalTermCost")),
    }))
    .map(x => ({ ...x, room: cleanText(x.room), term: cleanText(x.term) }));

    const fallT = off.find(x => x.room === selectedType && x.term === "Fall")?.total ?? 0;
    const winterT = off.find(x => x.room === selectedType && x.term === "Winter")?.total ?? 0;

    livingFall = fallT;
    livingWinter = winterT;
    livingYear = fallT + winterT;
  }

  // Meal plan totals (applies regardless of level for now)
  const meal = $("mealplan").value;
  let mealYear = 0;

  if (meal !== "None") {
    const mp = (DATA?.Meal_Plan ?? []).map(r => ({
      name: cleanText(r["Meal Plan Size"]),
      total: moneyToNumber(r["Total cost per year"]),
    })).find(x => x.name === meal);

    if (mp) mealYear = mp.total ?? 0;
  }

  // Grand total
  const grand = tuitionTotal + livingYear + mealYear;
  $("grandTotal").textContent = fmt(grand);

  // Breakdown table
  const lines = [];

  if (level === "Graduate") {
    lines.push(["Tuition & fees (Fall+Winter+Summer)", tuitionTotal]);
  } else {
    lines.push(["Tuition & fees (Fall+Winter)", tuitionTotal]);
  }

  lines.push(["Living (Fall)", livingFall]);
  lines.push(["Living (Winter)", livingWinter]);
  lines.push(["Living (Total)", livingYear]);
  lines.push(["Meal plan (Total)", mealYear]);

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

  const ocRoom = $("oncampusRoom");
  const ocRes = $("oncampusRes");

  if (ocRoom && ocRes) {
    ocRoom.disabled = housing !== "OnCampus";
    ocRes.disabled = housing !== "OnCampus";
  }

  $("offcampus").disabled = housing !== "OffCampus";

  compute();
}

function syncProvinceToResidency() {
  // Undergrad uses Province. Grad does not.
  const level = getLevel();
  const residency = $("residency").value;

  if (level === "Graduate") {
    // Disable province for graduate (not used in data)
    $("province").disabled = true;
    return;
  }

  $("province").disabled = false;

  if (residency === "International") {
    $("province").value = "INT";
  } else {
    if ($("province").value === "INT") $("province").value = "ON";
  }
}

function syncUIToLevel() {
  const level = getLevel();

  // Province only meaningful for undergrad
  if (level === "Graduate") {
    $("province").disabled = true;
    // Don't force province values in graduate mode
  } else {
    $("province").disabled = false;
    syncProvinceToResidency();
  }

  // Major disabled in grad mode
  if ($("major")) {
    $("major").disabled = (level === "Graduate");
  }
}

/* -----------------------
   Init
------------------------ */
async function init() {
  const res = await fetch("./data.json");
  DATA = await res.json();

  updateLivingDropdowns();
  syncUIToLevel();
  updateProgramMajor();

  // Events
  if ($("level")) {
    $("level").addEventListener("change", () => {
      syncUIToLevel();
      // Rebuild program/major choices based on level
      updateProgramMajor();
    });
  }

  ["residency", "province", "load", "cohort"].forEach(id =>
    $(id).addEventListener("change", () => {
      syncUIToLevel();
      updateProgramMajor();
    })
  );

  $("program").addEventListener("change", () => {
    if (getLevel() === "Graduate") compute();
    else updateMajors();
  });

  if ($("major")) $("major").addEventListener("change", compute);

  $("housing").addEventListener("change", toggleLivingInputs);
  // Removed: $("oncampus").addEventListener("change", compute);
  $("offcampus").addEventListener("change", compute);
  $("mealplan").addEventListener("change", compute);

  toggleLivingInputs();
}

init().catch(err => {
  console.error(err);
  alert("Failed to load data.json. Run a local server (Live Server) and try again.");
});
