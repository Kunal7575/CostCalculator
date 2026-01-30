/* =======================
   UofG Cost Calculator - app.js (fixed + prototype email UI)

   Fixes:
   1) Graduate Part-time reads from Tuition_Fees_Grad_Part_tim (not Tuition_Fees_Graduate)
   2) Load comparisons are normalized everywhere (Part-time vs Part-Time vs part time)
   3) Province dropdown for Graduate:
        - Disabled (as intended)
        - Shows INT when Residency = International
        - Shows ON/Non-ON when Residency = Domestic (but disabled)
   4) Prototype-only "Email my estimate" UI logic (localStorage save + validation)
======================= */

let DATA = null;
const $ = (id) => document.getElementById(id);

/* -----------------------
   Helpers
------------------------ */
function cleanText(s) { return String(s ?? "").trim(); }
function normalizeResidence(s) { return cleanText(s).toUpperCase(); }
function unique(arr) { return [...new Set(arr)]; }

function normalizeLoadToken(s) {
  return cleanText(s).toLowerCase().replace(/\s+/g, "").replace(/-/g, "");
}
function isPartTime(load) { return normalizeLoadToken(load) === "parttime"; }
function isFullTime(load) { return normalizeLoadToken(load) === "fulltime"; }

function getVal(obj, wantedKey) {
  if (!obj || typeof obj !== "object") return undefined;
  const foundKey = Object.keys(obj).find(k => k.trim() === wantedKey);
  return foundKey ? obj[foundKey] : undefined;
}

function moneyToNumber(v) {
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

function normalizeCohortToken(s) {
  return cleanText(s)
    .replace(/[–—]/g, "-")
    .replace(/\s+/g, "")
    .toLowerCase();
}

/* -----------------------
   Province options
   - Undergrad:
       Domestic -> ON, Non-ON
       International -> INT
   - Graduate:
       Province disabled, but display:
         International -> INT
         Domestic -> ON/Non-ON (value kept if already ON/Non-ON else ON)
------------------------ */
function updateProvinceOptions() {
  const level = cleanText($("level")?.value);
  const residency = cleanText($("residency")?.value);
  const provinceEl = $("province");
  if (!provinceEl) return;

  const isInternational = residency === "International";

  if (isInternational) {
    setOptions(provinceEl, ["INT"]);
    provinceEl.value = "INT";
  } else {
    setOptions(provinceEl, ["ON", "Non-ON"]);
    if (provinceEl.value !== "ON" && provinceEl.value !== "Non-ON") {
      provinceEl.value = "ON";
    }
  }

  // disable only for Graduate, but don't skip setting the correct value
  provinceEl.disabled = (level === "Graduate");
}

/* -----------------------
   Summer UI (Graduate only)
------------------------ */
function setSummerUI() {
  const level = cleanText($("level")?.value);
  const wrap = $("summerWrap");
  const cb = $("includeSummer");
  if (!wrap || !cb) return;

  if (level === "Graduate") {
    wrap.style.display = "block";
  } else {
    wrap.style.display = "none";
    cb.checked = false;
  }
}

/* -----------------------
   Credits UI
   - UG Part-time: 0.25..1.75
   - Grad Part-time: 0.25..1.00 (Summer uses fixed 1.25 row)
------------------------ */
function setCreditsUI() {
  const level = cleanText($("level")?.value);
  const load = cleanText($("load")?.value);

  const wrap = $("creditsWrap");
  const creditsSel = $("credits");
  const hint = $("creditsHint");

  if (!wrap || !creditsSel || !hint) return;

  const ugPT = (level === "Undergraduate" && isPartTime(load));
  const gradPT = (level === "Graduate" && isPartTime(load));
  const show = ugPT || gradPT;

  if (show) {
    wrap.style.display = "block";

    const creditOptions = gradPT
      ? ["0.25", "0.5", "0.75", "1"]
      : ["0.25", "0.5", "0.75", "1", "1.25", "1.5", "1.75"];

    setOptions(creditsSel, creditOptions, "Select credits");

    // Keep current selection if possible
    if (!creditOptions.includes(cleanText(creditsSel.value))) {
      creditsSel.value = creditOptions[0];
    }

    hint.textContent = gradPT
      ? "Graduate part-time credits apply to Fall/Winter. Summer uses the 1.25-credit rate (if included)."
      : "";

    // Major UI handling
    if ($("major") && ugPT) {
      $("major").disabled = true;
      $("major").innerHTML = `<option value="">N/A for part-time</option>`;
      $("major").value = "";
    }
  } else {
    wrap.style.display = "none";
    hint.textContent = "Full-time is 2.00+ credits (no selection needed).";

    if ($("major")) {
      if (level === "Undergraduate" && isFullTime(load)) {
        $("major").disabled = false;
      }
      if (level === "Graduate") {
        $("major").disabled = true;
        $("major").innerHTML = `<option value="">N/A for Graduate</option>`;
        $("major").value = "";
      }
    }
  }
}

/* -----------------------
   Filters
------------------------ */
function getLevel() {
  return cleanText($("level")?.value || "Undergraduate");
}

function currentFilters() {
  return {
    Level: getLevel(),
    Residency: cleanText($("residency")?.value),
    Province: cleanText($("province")?.value),
    Load: cleanText($("load")?.value),
    CohortYear: cleanText($("cohort")?.value),
    Credits: cleanText($("credits")?.value),
  };
}

/* -----------------------
   Tuition sources
------------------------ */
function getUndergradPartTimeRowsRaw() {
  // Your JSON: Tuition_Fees_Undergrad_Part_tim
  return DATA?.Tuition_Fees_Undergrad_Part_tim ?? [];
}

function getGradPartTimeRowsRaw() {
  // Your JSON: Tuition_Fees_Grad_Part_tim
  return DATA?.Tuition_Fees_Grad_Part_tim ?? [];
}

function getTuitionRows() {
  const f = currentFilters();

  // Graduate (choose correct table based on load)
  if (f.Level === "Graduate") {
    const src = isPartTime(f.Load)
      ? getGradPartTimeRowsRaw()
      : (DATA?.Tuition_Fees_Graduate ?? []);

    return (src || []).map(r => ({
      Level: "Graduate",
      Program: cleanText(r.Program),
      Major: "",
      Credits: cleanText(r.Credits), // present for grad PT; may be blank for grad FT
      Residency: cleanText(r.Residency),
      Province: cleanText(r.Province),
      Load: cleanText(r.Load),
      CohortYear: cleanText(r.CohortYear),

      FallTotalN: moneyToNumber(r.FallTotal),
      WinterTotalN: moneyToNumber(r.WinterTotal),
      SummerTotalN: moneyToNumber(r.SummerTotal),
      FallWinterTotalN: moneyToNumber(r.FallWinterTotal),
    }));
  }

  // Undergrad Part-time
  if (isPartTime(f.Load)) {
    const src = getUndergradPartTimeRowsRaw();
    return (src || []).map(r => ({
      Level: "Undergraduate",
      Program: cleanText(r.Program),
      Major: "",
      Credits: cleanText(r.Credits),
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

  // Undergrad Full-time
  const src = DATA?.Tuition_Fees ?? [];
  return (src || []).map(r => ({
    Level: "Undergraduate",
    Program: cleanText(r.Program),
    Major: cleanText(r.Major),
    Credits: "",
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

  // Graduate: Residency + Load + CohortYear (normalized load + cohort)
  if (f.Level === "Graduate") {
    const wantLoad = normalizeLoadToken(f.Load);
    const wantCohort = normalizeCohortToken(f.CohortYear);

    return rows.filter(r =>
      r.Level === "Graduate" &&
      r.Residency === f.Residency &&
      normalizeLoadToken(r.Load) === wantLoad &&
      normalizeCohortToken(r.CohortYear) === wantCohort
    );
  }

  // UG Part-time: Residency + Province + CohortYear + Load token
  if (isPartTime(f.Load)) {
    const wantLoad = "parttime";
    const wantCohort = normalizeCohortToken(f.CohortYear);

    return rows.filter(r =>
      r.Level === "Undergraduate" &&
      r.Residency === f.Residency &&
      r.Province === f.Province &&
      normalizeCohortToken(r.CohortYear) === wantCohort &&
      normalizeLoadToken(r.Load) === wantLoad
    );
  }

  // UG Full-time: Residency + Province + Load + CohortYear (normalized cohort)
  const wantCohort = normalizeCohortToken(f.CohortYear);
  return rows.filter(r =>
    r.Level === "Undergraduate" &&
    r.Residency === f.Residency &&
    r.Province === f.Province &&
    r.Load === f.Load &&
    normalizeCohortToken(r.CohortYear) === wantCohort
  );
}

/* -----------------------
   Program/Major UI
------------------------ */
function updateProgramMajor() {
  const f = currentFilters();
  const rows = getFilteredTuitionRows();

  const programs = unique(rows.map(r => r.Program)).filter(Boolean).sort();
  setOptions($("program"), programs, "Select program");

  // Only auto-select if we actually have options
  if (programs.length) {
    $("program").value = programs[0];
  }

  // Graduate
  if (f.Level === "Graduate") {
    if ($("major")) {
      setOptions($("major"), [], "N/A for Graduate");
      $("major").value = "";
      $("major").disabled = true;
    }
    compute();
    return;
  }

  // UG Part-time
  if (isPartTime(f.Load)) {
    if ($("major")) {
      $("major").disabled = true;
      $("major").innerHTML = `<option value="">N/A for part-time</option>`;
      $("major").value = "";
    }
    compute();
    return;
  }

  // UG Full-time
  if ($("major")) $("major").disabled = false;
  updateMajors();
}

function updateMajors() {
  const rows = getFilteredTuitionRows();
  const program = cleanText($("program")?.value);

  const majors = unique(rows.filter(r => r.Program === program).map(r => r.Major))
    .filter(Boolean)
    .sort();

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
  const onCampusRows = DATA?.On_campus_Living_Costs ?? [];

  const ocRows = onCampusRows
    .map(r => ({
      room: cleanText(r.RoomType),
      res: normalizeResidence(r.ResidenceArea),
      fall: moneyToNumber(r["Fall Term"]),
      winter: moneyToNumber(r["Winter Term"]),
      total: moneyToNumber(r.Cost),
    }))
    .filter(x => x.room && x.res && x.total !== null);

  const roomTypes = unique(ocRows.map(x => x.room)).sort();
  setOptions($("oncampusRoom"), roomTypes, "Select room type");

  function updateOnCampusResidences() {
    const selectedRoom = $("oncampusRoom")?.value || "";
    const residences = unique(ocRows.filter(x => x.room === selectedRoom).map(x => x.res)).sort();
    setOptions($("oncampusRes"), residences, "Select residence");
    if (residences.length) $("oncampusRes").value = residences[0];
  }

  window.__OC_ROWS__ = ocRows;

  if ($("oncampusRoom")) {
    $("oncampusRoom").onchange = () => { updateOnCampusResidences(); compute(); };
  }
  if ($("oncampusRes")) {
    $("oncampusRes").onchange = () => compute();
  }

  if (roomTypes.length && $("oncampusRoom")) {
    $("oncampusRoom").value = roomTypes[0];
    updateOnCampusResidences();
  }

  const off = offCampusRows
    .map(r => ({
      room: cleanText(getVal(r, "RoomType")),
      term: cleanText(getVal(r, "Term")),
      total: moneyToNumber(getVal(r, "TotalTermCost")),
    }))
    .map(x => ({ ...x, room: cleanText(x.room), term: cleanText(x.term) }))
    .filter(x => x.room && x.term && x.total !== null);

  const offTypes = unique(off.map(x => x.room)).sort();
  setOptions($("offcampus"), offTypes, "Select off-campus option");

  const mp = mealRows
    .map(r => ({
      name: cleanText(r["Meal Plan Size"]),
      total: moneyToNumber(r["Total cost per year"]),
    }))
    .filter(x => x.name);

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
  const f = currentFilters();
  const rows = getFilteredTuitionRows();
  const program = cleanText($("program")?.value);

  let match = null;
  let matchSummerRow = null;

  const gradPT = (f.Level === "Graduate" && isPartTime(f.Load));
  const ugPT = (f.Level === "Undergraduate" && isPartTime(f.Load));

  if (gradPT) {
    const creditsFW = cleanText($("credits")?.value); // 0.25..1
    match = rows.find(r => r.Program === program && r.Credits === creditsFW);

    // Summer stored/priced as 1.25 credits in dataset
    matchSummerRow = rows.find(r => r.Program === program && r.Credits === "1.25");
  } else if (f.Level === "Graduate") {
    match = rows.find(r => r.Program === program);
  } else if (ugPT) {
    const credits = cleanText($("credits")?.value);
    match = rows.find(r => r.Program === program && r.Credits === credits);
  } else {
    const major = cleanText($("major")?.value);
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

  // Tuition
  const includeSummer = $("includeSummer")?.checked === true;

  const fall = match?.FallTotalN ?? 0;
  const winter = match?.WinterTotalN ?? 0;

  const tuitionFallWinter = match?.FallWinterTotalN ?? (fall + winter);

  let summer = 0;
  if (gradPT) {
    summer = matchSummerRow?.SummerTotalN ?? 0;
  } else {
    summer = match?.SummerTotalN ?? 0;
  }

  const summerUsed = (f.Level === "Graduate" && includeSummer) ? summer : 0;
  const tuitionTotal = tuitionFallWinter + summerUsed;

  // Living
  const housing = cleanText($("housing")?.value);
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
    const selectedType = $("offcampus")?.value;

    const off = (DATA?.Off_campus_Living_Costs ?? [])
      .map(r => ({
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

  // Meal plan
  const meal = cleanText($("mealplan")?.value);
  let mealYear = 0;

  if (meal && meal !== "None") {
    const mp = (DATA?.Meal_Plan ?? [])
      .map(r => ({
        name: cleanText(r["Meal Plan Size"]),
        total: moneyToNumber(r["Total cost per year"]),
      }))
      .find(x => x.name === meal);

    if (mp) mealYear = mp.total ?? 0;
  }

  // Totals
  const grand = tuitionTotal + livingYear + mealYear;
  $("grandTotal").textContent = fmt(grand);

  // Breakdown
  const lines = [];

  if (f.Level === "Graduate") {
    lines.push(["Tuition & fees (Fall)", fall]);
    lines.push(["Tuition & fees (Winter)", winter]);
    lines.push(["Tuition & fees (Summer)", includeSummer ? summer : 0]);
    lines.push(["Tuition & fees (Total)", tuitionTotal]);
  } else if (ugPT) {
    lines.push([`Tuition & fees (Fall+Winter) - ${cleanText($("credits")?.value)} credits`, tuitionTotal]);
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
   Living toggles (show/hide blocks)
------------------------ */
function toggleLivingInputs() {
  const housing = cleanText($("housing")?.value);

  const ocRoom = $("oncampusRoom");
  const ocRes = $("oncampusRes");
  const offSel = $("offcampus");

  const ocRoomWrap = $("onCampusRoomWrap");
  const ocResWrap = $("onCampusResWrap");
  const offWrap = $("offCampusWrap");

  const showOnCampus = housing === "OnCampus";
  const showOffCampus = housing === "OffCampus";

  if (ocRoomWrap) ocRoomWrap.style.display = showOnCampus ? "block" : "none";
  if (ocResWrap) ocResWrap.style.display = showOnCampus ? "block" : "none";
  if (offWrap) offWrap.style.display = showOffCampus ? "block" : "none";

  if (ocRoom) ocRoom.disabled = !showOnCampus;
  if (ocRes) ocRes.disabled = !showOnCampus;
  if (offSel) offSel.disabled = !showOffCampus;

  if (!showOnCampus) {
    if (ocRoom) ocRoom.selectedIndex = 0;
    if (ocRes) ocRes.selectedIndex = 0;
  }
  if (!showOffCampus) {
    if (offSel) offSel.selectedIndex = 0;
  }

  compute();
}

/* -----------------------
   Prototype-only: Email estimate UI (local only)
   Requires HTML ids:
   firstName, email, emailEstimateBtn, clearEmailBtn, emailMsg
------------------------ */
function initEmailPrototypeUI() {
  const emailBtn = $("emailEstimateBtn");
  const clearBtn = $("clearEmailBtn");
  const msg = $("emailMsg");

  function isValidEmail(v) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(v || "").trim());
  }

  emailBtn?.addEventListener("click", () => {
    const firstName = cleanText($("firstName")?.value);
    const email = cleanText($("email")?.value);

    if (!msg) return;

    if (!firstName || !email || !isValidEmail(email)) {
      msg.style.display = "block";
      msg.textContent = "Please enter a valid first name and email address.";
      return;
    }

    localStorage.setItem("cc_firstName", firstName);
    localStorage.setItem("cc_email", email);

    msg.style.display = "block";
    msg.textContent =
      `Saved for prototype!! We will email the current estimate to ${email} (coming soon (Just a prototype😒)).`;
  });

  clearBtn?.addEventListener("click", () => {
    if ($("firstName")) $("firstName").value = "";
    if ($("email")) $("email").value = "";
    localStorage.removeItem("cc_firstName");
    localStorage.removeItem("cc_email");
    if (msg) msg.style.display = "none";
  });

  // Prefill if saved
  window.addEventListener("load", () => {
    const fn = localStorage.getItem("cc_firstName");
    const em = localStorage.getItem("cc_email");
    if (fn && $("firstName")) $("firstName").value = fn;
    if (em && $("email")) $("email").value = em;
  });
}

/* -----------------------
   Init
------------------------ */
async function init() {
  const res = await fetch("./data.json");
  DATA = await res.json();

  // Optional UI-only feature
  initEmailPrototypeUI();

  updateLivingDropdowns();

  updateProvinceOptions();
  setSummerUI();
  setCreditsUI();
  updateProgramMajor();

  // EVENTS
  if ($("level")) {
    $("level").addEventListener("change", () => {
      updateProvinceOptions();
      setSummerUI();
      setCreditsUI();
      updateProgramMajor();
    });
  }

  if ($("includeSummer")) $("includeSummer").addEventListener("change", compute);

  if ($("residency")) {
    $("residency").addEventListener("change", () => {
      updateProvinceOptions();
      setSummerUI();
      setCreditsUI();
      updateProgramMajor();
    });
  }

  if ($("province")) {
    $("province").addEventListener("change", () => {
      setCreditsUI();
      updateProgramMajor();
    });
  }

  if ($("load")) {
    $("load").addEventListener("change", () => {
      updateProvinceOptions();
      setSummerUI();
      setCreditsUI();
      updateProgramMajor();
    });
  }

  if ($("credits")) {
    $("credits").addEventListener("change", () => {
      updateProgramMajor();
      compute();
    });
  }

  if ($("cohort")) {
    $("cohort").addEventListener("change", () => {
      setCreditsUI();
      updateProgramMajor();
    });
  }

  if ($("program")) {
    $("program").addEventListener("change", () => {
      const f = currentFilters();
      if (f.Level === "Undergraduate" && isFullTime(f.Load)) updateMajors();
      else compute();
    });
  }

  if ($("major")) $("major").addEventListener("change", compute);

  if ($("housing")) $("housing").addEventListener("change", toggleLivingInputs);
  if ($("offcampus")) $("offcampus").addEventListener("change", compute);
  if ($("mealplan")) $("mealplan").addEventListener("change", compute);

  toggleLivingInputs();
}

init().catch(err => {
  console.error(err);
  alert("Failed to load data.json. Run a local server (Live Server) and try again.");
});
