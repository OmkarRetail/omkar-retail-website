(function () {
  const dayValues = { P: 1, WO: 1, "A-R": 1, "F-R": 1, HD: 0.5, "HD-R": 0.5, A: 0, F: 0, L: 0, PENDING: 0 };
  const files = { attendance: null, master: null, structure: null };
  const data = { calculations: [], activeCalculation: null };
  const arrearsByEmployee = new Map();
  const $ = (id) => document.getElementById(id);
  const normal = (value) => String(value || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  const text = (value) => String(value ?? "").trim();
  const firstValue = (...values) => values.find((value) => text(value)) || "";
  const money = (value) => new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: 2 }).format(Number(value || 0));
  const amountInWords = (value) => { const n = Math.round(Number(value || 0)); const ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"]; const tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]; const underThousand = (number) => { const parts = []; if (number >= 100) { parts.push(`${ones[Math.floor(number / 100)]} Hundred`); number %= 100; } if (number >= 20) { parts.push(tens[Math.floor(number / 10)]); if (number % 10) parts.push(ones[number % 10]); } else if (number) parts.push(ones[number]); return parts.join(" "); }; if (!n) return "Rupees Zero Only"; const parts = []; let remaining = n; [[10000000, "Crore"], [100000, "Lakh"], [1000, "Thousand"]].forEach(([unit, label]) => { if (remaining >= unit) { parts.push(`${underThousand(Math.floor(remaining / unit))} ${label}`); remaining %= unit; } }); if (remaining) parts.push(underThousand(remaining)); return `Rupees ${parts.join(" ")} Only`; };
  const escape = (value) => text(value).replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" })[char]);
  const status = (message, kind = "info") => { const el = $("status"); el.textContent = message; el.className = `status ${kind}`; };
  const calendarDate = (value) => { if (typeof value === "number") { const parts = XLSX.SSF.parse_date_code(Math.floor(value)); return parts ? `${parts.y}-${String(parts.m).padStart(2, "0")}-${String(parts.d).padStart(2, "0")}` : ""; } if (value instanceof Date) return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, "0")}-${String(value.getDate()).padStart(2, "0")}`; const match = text(value).match(/^(\d{4}-\d{2}-\d{2})/); return match ? match[1] : ""; };
  const displayDate = (value) => { const date = calendarDate(value); if (!date) return "Not available"; const [year, month, day] = date.split("-"); return `${day} ${["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][Number(month) - 1]} ${year}`; };
  const daysBetween = (start, end) => Math.round((end - start) / 86400000) + 1;
  const parseDoublePayDates = (value) => {
    const enteredDates = text(value).split(/[;,\n]+/).map((item) => item.trim()).filter(Boolean);
    const invalid = []; const dates = [];
    enteredDates.forEach((date) => {
      const match = date.match(/^(\d{2})-(\d{2})-(\d{4})$/);
      if (!match) { invalid.push(date); return; }
      const [, dayText, monthText, yearText] = match;
      const year = Number(yearText), month = Number(monthText), day = Number(dayText);
      const parsed = new Date(Date.UTC(year, month - 1, day));
      if (parsed.getUTCFullYear() !== year || parsed.getUTCMonth() !== month - 1 || parsed.getUTCDate() !== day) { invalid.push(date); return; }
      dates.push(`${yearText}-${monthText}-${dayText}`);
    });
    return { dates: new Set(dates), invalid };
  };

  function workbookRows(book, preferredSheet, requiredHeaders) {
    const preferred = book.SheetNames.find((sheet) => normal(sheet) === normal(preferredSheet));
    const sheets = preferred ? [preferred] : book.SheetNames;
    for (const sheetName of sheets) {
      const rows = XLSX.utils.sheet_to_json(book.Sheets[sheetName], { header: 1, defval: "", raw: true });
      for (let rowIndex = 0; rowIndex < Math.min(rows.length, 15); rowIndex++) {
        const headers = rows[rowIndex].map(normal);
        if (!requiredHeaders.every((header) => Array.isArray(header) ? header.some((option) => headers.includes(normal(option))) : headers.includes(normal(header)))) continue;
        return rows.slice(rowIndex + 1).map((row) => Object.fromEntries(rows[rowIndex].map((header, index) => [normal(header), row[index]]))).filter((row) => Object.values(row).some((value) => text(value)));
      }
    }
    return [];
  }

  async function readFile(file) { return XLSX.read(await file.arrayBuffer(), { type: "array", cellFormula: true, cellDates: false }); }

  function structureMap(book) {
    const map = new Map();
    const add = (key, item) => { if (key) map.set(normal(key), item); };
    book.SheetNames.forEach((sheetName) => {
      const sheet = book.Sheets[sheetName];
      const endRow = XLSX.utils.decode_range(sheet["!ref"] || "A1").e.r + 1;
      const markers = [];
      for (let row = 1; row <= endRow; row++) {
        const title = text(sheet[`A${row}`]?.v);
        const key = normal(title);
        if (key.startsWith("salarystucture") || key.startsWith("salarystructure") || key === "parttimesalary") markers.push({ row, title });
      }
      markers.forEach((marker, index) => {
        const nextRow = markers[index + 1]?.row || endRow + 1;
        const components = {}; let hasPf = false; let hasEsi = false;
        for (let row = marker.row + 1; row < nextRow; row++) {
          const label = normal(sheet[`A${row}`]?.v); const value = Number(sheet[`B${row}`]?.v || 0);
          if (label) components[label] = value;
          if (label === "pf") hasPf = true;
          if (label === "esi") hasEsi = true;
        }
        add(marker.title, { name: marker.title, basic: components.basicsalary || 0, hra: components.hra || 0, special: components.specialallowance || 0, conveyance: components.conveyanceallowance || 0, professionalTax: components.professionaltax || 0, hasPf, hasEsi, directFixed: false });
      });
      for (let row = 1; row <= endRow; row++) {
      const title = text(sheet[`A${row}`]?.v); if (!/^Sti(?:fund|pend)\s+\d+$/i.test(title)) continue;
        const monthlyGross = Number(sheet[`B${row}`]?.v || 0);
        add(title, { name: title, basic: monthlyGross, hra: 0, special: 0, conveyance: 0, professionalTax: 0, hasPf: false, hasEsi: false, directFixed: true });
      }
    });
    return map;
  }

  function renderList() {
    const list = $("employeeList");
    list.innerHTML = data.calculations.map((calc, index) => `<button type="button" data-index="${index}" class="${calc.employee.id === data.activeCalculation?.employee?.id ? "active" : ""}"><strong>${escape(calc.employee.name)}</strong><small>${escape(calc.employee.id)} · Net ${money(calc.net)}</small></button>`).join("");
    list.querySelectorAll("button").forEach((button) => button.addEventListener("click", () => { const calc = data.calculations[Number(button.dataset.index)]; list.querySelectorAll("button").forEach((item) => item.classList.remove("active")); button.classList.add("active"); renderSlip(calc); loadArrearsForm(calc.employee.id); }));
  }

  const row = (label, value) => `<div class="slip-row"><span>${label}</span><strong>${money(value)}</strong></div>`;
  const total = (label, value) => `<div class="slip-total"><span>${label}</span><strong>${money(value)}</strong></div>`;

  function updateArrearsForCalculation(calc, adjustment = arrearsByEmployee.get(calc.employee.id)) {
    const amount = Number(adjustment?.amount || 0);
    calc.arrears = Number.isFinite(amount) ? amount : 0;
    calc.arrearsReason = text(adjustment?.reason);
    calc.arrearsEarning = Math.max(calc.arrears, 0);
    calc.arrearsRecovery = Math.max(-calc.arrears, 0);
    calc.gross = calc.baseGross + calc.arrearsEarning;
    calc.deductions = calc.baseDeductions + calc.arrearsRecovery;
    calc.net = calc.gross - calc.deductions;
  }

  function populateArrearsEmployees(calculations) {
    const select = $("arrearsEmployee");
    select.innerHTML = calculations.map((calc) => `<option value="${escape(calc.employee.id)}">${escape(calc.employee.name)} (${escape(calc.employee.id)})</option>`).join("");
    $("arrearsControls").hidden = !calculations.length;
  }

  function loadArrearsForm(employeeId) {
    const select = $("arrearsEmployee");
    if (!employeeId || ![...select.options].some((option) => option.value === employeeId)) return;
    select.value = employeeId;
    const adjustment = arrearsByEmployee.get(employeeId);
    $("arrearsAmount").value = adjustment?.amount || "";
    $("arrearsReason").value = adjustment?.reason || "";
  }

  function renderSlip(calc) {
    data.activeCalculation = calc;
    const c = calc; const employee = c.employee;
    const detail = (label, value, blankWhenMissing = false) => `<div class="employee-detail"><small>${label}</small><strong>: ${escape(value || (blankWhenMissing ? "" : "Not available"))}</strong></div>`;
    const validUan = /^\d{12}$/.test(text(employee.uan)) ? text(employee.uan) : "";
    const pfNumber = ["", "notavailable", "na", "nan"].includes(normal(employee.pfNumber)) ? "" : text(employee.pfNumber);
    const isStipend = c.directFixed || /sti(?:pend|fund)/i.test(text(employee.structure));
    const statutoryDetails = isStipend ? "" : `${detail("UAN", validUan, true)}${detail("PF NUMBER", pfNumber, true)}`;
    const arrearsLabel = `Arrears${c.arrearsReason ? ` – ${escape(c.arrearsReason)}` : ""}`;
    const recoveryLabel = `Arrears Recovery${c.arrearsReason ? ` – ${escape(c.arrearsReason)}` : ""}`;
    const earnings = [[c.directFixed ? "Stipend Pay" : "Basic Salary", c.basic], ["HRA", c.hra], ["Special Allowance", c.special], ["Conveyance Allowance", c.conveyance], ["Double Pay", c.doublePay], [arrearsLabel, c.arrearsEarning], ["Attendance Bonus", c.bonus]].filter(([, value]) => value > 0).map(([label, value]) => row(label, value)).join("");
    const deductions = [["Provident Fund", c.pf], ["ESI", c.esi], ["Professional Tax", c.pt], [recoveryLabel, c.arrearsRecovery]].filter(([, value]) => value > 0).map(([label, value]) => row(label, value)).join("") || `<div class="slip-row"><span>No employee deductions in this structure</span><strong>${money(0)}</strong></div>`;
    $("payslipPreview").innerHTML = `<div class="slip-head"><div class="slip-brand">OMKAR RETAIL VENTURES</div><div class="statement-period">Salary Statement for ${escape(c.period)}</div></div><div class="slip-person"><div class="employee-column">${detail("EMPLOYEE NAME", employee.name)}${detail("EMPLOYEE ID", employee.id)}${detail("LOCATION", c.location)}${detail("DESIGNATION", employee.role)}${detail("DAYS WORKED", `${c.paidDays} / ${c.cycleDays}`)}${c.doublePayDays ? detail("DOUBLE-PAY DAYS", c.doublePayDays) : ""}</div><div class="employee-column">${detail("PAN", employee.pan)}${statutoryDetails}${detail("BANK NAME", employee.bank)}${detail("BANK ACCOUNT NUMBER", employee.accountNumber)}${detail("DATE OF JOINING", displayDate(employee.doj))}</div></div><div class="slip-tables"><div class="pay-table"><div class="table-heading"><span>PARTICULARS</span><span>EARNINGS</span></div>${earnings}${total("GROSS EARNINGS", c.gross)}</div><div class="pay-table"><div class="table-heading"><span>PARTICULARS</span><span>DEDUCTIONS</span></div>${deductions}${total("TOTAL DEDUCTIONS", c.deductions)}</div></div><div class="net-pay"><span>NET PAY</span><strong>${money(c.net)}</strong></div><div class="net-words">(${escape(amountInWords(c.net))})</div><p class="note">* This is a system-generated payslip and is confidential; therefore no signature is required.</p>`;
  }

  function buildMasterMap(shiftRows, masterRows) {
    const employees = new Map();
    [...shiftRows, ...masterRows].forEach((row) => {
      const id = firstValue(row[normal("Employee ID")], row[normal("Z ID")]); if (!id) return;
      const prior = employees.get(id) || {};
      employees.set(id, { id, name: firstValue(row.name, prior.name), email: firstValue(row.email, prior.email), role: firstValue(row.role, prior.role), structure: firstValue(row[normal("Salary Structure")], prior.structure), doj: firstValue(row.doj, prior.doj), pan: firstValue(row.pan, prior.pan), uan: firstValue(row.uan, prior.uan), bank: firstValue(row.bank, prior.bank), accountNumber: firstValue(row[normal("Account number")], prior.accountNumber), pfNumber: firstValue(row[normal("PF number")], prior.pfNumber), location: firstValue(row.location, prior.location) });
  });
  return employees;
}

  async function getOwnerDeliveryToken() {
    const config = window.OMKAR_SITE_CONFIG || {};
    const firebaseConfig = config.firebase;
    if (!firebaseConfig) throw new Error("Firebase login is not configured.");
    const [{ initializeApp, getApps }, authMod] = await Promise.all([
      import("https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js"),
      import("https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js")
    ]);
    const app = getApps()[0] || initializeApp(firebaseConfig);
    const auth = authMod.getAuth(app);
    const user = auth.currentUser || await new Promise((resolve) => {
      const unsubscribe = authMod.onAuthStateChanged(auth, (signedInUser) => { unsubscribe(); resolve(signedInUser); });
    });
    const email = text(user?.email).toLowerCase();
    if (!user || config.adminRoles?.[email] !== "owner") throw new Error("Sign in as the Omkar Admin account before saving or emailing payslips.");
    return user.getIdToken();
  }

  function payslipFileName(calc) {
    const employeeName = text(calc?.employee?.name).replace(/[^A-Za-z0-9]+/g, "_").replace(/^_+|_+$/g, "") || "Employee";
    const employeeId = text(calc?.employee?.id).replace(/[^A-Za-z0-9_-]/g, "_") || "employee";
    const salaryCycle = text($("cycleMonth")?.value).replace(/[^0-9-]/g, "") || "cycle";
    return `${employeeName}-${employeeId}-${salaryCycle}.pdf`;
  }

  async function saveToDriveAndEmail() {
    const config = window.OMKAR_SITE_CONFIG || {};
    const endpoint = text(config.payslipDeliveryWebAppUrl);
    const calc = data.activeCalculation;
    if (!endpoint) return status("Drive and email delivery is not connected yet. Add the deployed Payslip Delivery Apps Script URL in config.js when Google Drive is available.", "error");
    if (!calc || !$("payslipPreview").innerHTML) return status("Generate and select a payslip first.", "error");
    if (typeof window.html2pdf !== "function") return status("PDF delivery support could not be loaded. Check your internet connection and try again.", "error");
    try {
      status("Creating the payslip PDF and sending the secure delivery request...");
      const idToken = await getOwnerDeliveryToken();
      const fileName = payslipFileName(calc);
      const pdfDataUri = await window.html2pdf().set({ margin: 6, filename: fileName, html2canvas: { scale: 2 }, jsPDF: { unit: "mm", format: "a4", orientation: "landscape" }, pagebreak: { mode: ["avoid-all", "css", "legacy"] } }).from($("payslipPreview")).outputPdf("datauristring");
      const payload = {
        idToken,
        fileName,
        employeeName: calc.employee.name,
        employeeId: calc.employee.id,
        employeeEmail: calc.employee.email || "",
        salaryCycle: calc.period,
        salaryMonth: $("cycleMonth").value,
        pdfBase64: pdfDataUri.split(",")[1]
      };
      await fetch(endpoint, { method: "POST", mode: "no-cors", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify(payload) });
      status("Payslip delivery request sent. It will be saved to Drive and emailed when the employee has a valid email address.", "success");
    } catch (error) {
      console.error("Payslip delivery failed", error);
      status(error.message || "Unable to send the payslip delivery request.", "error");
    }
  }

  function calculate() {
    if (!files.attendance || !files.master || !files.structure) return status("Please select all three Excel files first.", "error");
    Promise.all([readFile(files.attendance), readFile(files.master), readFile(files.structure)]).then(([attendanceBook, masterBook, structureBook]) => {
      const attendance = workbookRows(attendanceBook, "Attendance", ["employee_code", "scheduled_date", "muster_status"]);
      const master = buildMasterMap(workbookRows(masterBook, "Employee_Shift", ["Name", ["Employee ID", "Z ID"], "Salary Structure"]), workbookRows(masterBook, "Master", ["Name", ["Employee ID", "Z ID"], "Salary Structure"]));
      const filter = $("employeeFilter");
      if (!filter.dataset.loaded) { [...master.values()].sort((a, b) => a.name.localeCompare(b.name)).forEach((employee) => filter.insertAdjacentHTML("beforeend", `<option value="${escape(employee.id)}">${escape(employee.name)} (${escape(employee.id)})</option>`)); filter.dataset.loaded = "1"; }
      const structures = structureMap(structureBook); const month = $("cycleMonth").value;
      if (!month) return status("Choose the salary-cycle month (the cycle runs from the previous 21st to this month’s 20th).", "error");
      const [year, monthNumber] = month.split("-").map(Number); const start = new Date(year, monthNumber - 2, 21); const end = new Date(year, monthNumber - 1, 20); const cycleDays = daysBetween(start, end);
      const requested = filter.value;
      const doublePayInput = parseDoublePayDates($("doublePayDates").value);
      if (doublePayInput.invalid.length) return status("Enter double-pay dates in DD-MM-YYYY format, separated by commas.", "error");
      const doublePayDates = doublePayInput.dates;
      const grouped = new Map();
      const cycleStart = calendarDate(start); const cycleEnd = calendarDate(end);
      const outsideCycleDates = [...doublePayDates].filter((date) => date < cycleStart || date > cycleEnd);
      if (outsideCycleDates.length) return status(`Double-pay date${outsideCycleDates.length === 1 ? "" : "s"} must fall within this salary cycle: ${outsideCycleDates.join(", ")}.`, "error");
      attendance.forEach((record) => {
        const dateKey = calendarDate(record.scheduleddate); const id = text(record.employeecode);
        if (!dateKey || dateKey < cycleStart || dateKey > cycleEnd || !master.has(id)) return;
        if (!grouped.has(id)) grouped.set(id, new Map()); const dates = grouped.get(id);
        if (!dates.has(dateKey)) dates.set(dateKey, { statuses: [], locations: [] });
        const entry = dates.get(dateKey); entry.statuses.push(text(record.musterstatus).toUpperCase()); if (text(record.storename)) entry.locations.push(text(record.storename));
      });
      const results = []; const missingStructures = new Set(); const conflicts = [];
      grouped.forEach((dates, id) => {
        if (requested && requested !== id) return;
        const employee = master.get(id); const structure = structures.get(normal(employee.structure));
        if (!structure) { missingStructures.add(employee.structure || "(blank)"); return; }
        const dayEntries = [...dates.entries()].sort(([a], [b]) => a.localeCompare(b)); const conflictingDates = dayEntries.filter(([, entry]) => new Set(entry.statuses).size > 1);
        if (conflictingDates.length) { conflicts.push(`${employee.name} (${conflictingDates.map(([date]) => date).join(", ")})`); return; }
        const statuses = dayEntries.map(([, entry]) => entry.statuses[0]); const paidDays = statuses.reduce((sum, value) => sum + (dayValues[value] ?? 0), 0);
        const employeeDoublePayDays = dayEntries.filter(([date, entry]) => doublePayDates.has(date) && entry.statuses[0] === "P").length; const factor = paidDays / cycleDays; const doublePayFactor = employeeDoublePayDays / cycleDays;
        const basic = structure.basic * factor, hra = structure.hra * factor, special = structure.special * factor, conveyance = structure.conveyance * factor;
        const doublePay = (structure.basic + structure.hra + structure.special + structure.conveyance) * doublePayFactor; const fixedGross = basic + hra + special + conveyance + doublePay;
    const bonus = paidDays === cycleDays ? 500 : 0; const pfBasic = structure.basic * (factor + doublePayFactor); const pf = structure.hasPf ? pfBasic * 0.12 : 0, esi = structure.hasEsi ? fixedGross * 0.0075 : 0, pt = structure.professionalTax * (factor + doublePayFactor); const baseGross = fixedGross + bonus, baseDeductions = pf + esi + pt;
        const locations = [...new Set(dayEntries.flatMap(([, entry]) => entry.locations))];
        const calculation = { employee, location: firstValue(locations.join(", "), employee.location), period: `${start.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" })} – ${end.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" })}`, cycleDays, paidDays, doublePayDays: employeeDoublePayDays, basic, hra, special, conveyance, doublePay, bonus, baseGross, baseDeductions, gross: baseGross, pf, esi, pt, deductions: baseDeductions, net: baseGross - baseDeductions, directFixed: structure.directFixed, statusSummary: [...new Set(statuses)].join(", ") || "No records", countedDates: dayEntries.map(([date, entry]) => `${date} (${entry.statuses[0]})`) };
        updateArrearsForCalculation(calculation);
        results.push(calculation);
      });
      data.calculations = results.sort((a, b) => a.employee.name.localeCompare(b.employee.name));
      if (!results.length) { $("resultArea").hidden = true; return status("No matched employees were found for this cycle. Check employee IDs, attendance dates, and salary-structure names.", "error"); }
      $("resultArea").hidden = false; populateArrearsEmployees(data.calculations); renderSlip(results[0]); renderList(); loadArrearsForm(results[0].employee.id);
      const skipped = missingStructures.size ? ` ${missingStructures.size} salary-structure reference${missingStructures.size === 1 ? " is" : "s are"} not present in the uploaded salary workbook and were skipped: ${[...missingStructures].join(", ")}.` : "";
      const conflictNote = conflicts.length ? ` ${conflicts.length} employee${conflicts.length === 1 ? " has" : "s have"} conflicting attendance records and ${conflicts.length === 1 ? "was" : "were"} blocked for review: ${conflicts.join("; ")}.` : "";
      const doublePayNote = doublePayDates.size ? ` Double pay was added for employees marked P on: ${[...doublePayDates].map(displayDate).join(", ")}.` : "";
      status(`${results.length} payslip${results.length === 1 ? "" : "s"} generated for the selected salary cycle.${doublePayNote}${skipped}${conflictNote}`);
    }).catch((error) => status(`Unable to read the files: ${error.message}`, "error"));
  }

  function applyArrears() {
    const employeeId = text($("arrearsEmployee").value);
    const amountText = text($("arrearsAmount").value);
    const amount = Number(amountText);
    const reason = text($("arrearsReason").value);
    const calc = data.calculations.find((item) => item.employee.id === employeeId);
    if (!calc) return status("Generate the employee payslip before applying arrears.", "error");
    if (!amountText || !Number.isFinite(amount) || amount === 0) return status("Enter a non-zero arrears amount.", "error");
    if (!reason) return status("Enter the reason for the arrears.", "error");
    arrearsByEmployee.set(employeeId, { amount, reason });
    updateArrearsForCalculation(calc);
    renderSlip(calc); renderList(); loadArrearsForm(employeeId);
    status(`${amount > 0 ? "Arrears" : "Arrears recovery"} of ${money(Math.abs(amount))} applied for ${calc.employee.name}.`, "success");
  }

  function clearArrears() {
    const employeeId = text($("arrearsEmployee").value);
    const calc = data.calculations.find((item) => item.employee.id === employeeId);
    if (!calc) return status("Generate the employee payslip before removing arrears.", "error");
    arrearsByEmployee.delete(employeeId);
    updateArrearsForCalculation(calc);
    renderSlip(calc); renderList(); loadArrearsForm(employeeId);
    status(`Arrears removed for ${calc.employee.name}.`, "success");
  }

  ["attendanceFile", "masterFile", "structureFile"].forEach((id) => $(id).addEventListener("change", (event) => { files[id.replace("File", "")] = event.target.files[0] || null; }));
  $("arrearsEmployee").addEventListener("change", (event) => { const calc = data.calculations.find((item) => item.employee.id === event.target.value); loadArrearsForm(event.target.value); if (calc) { renderSlip(calc); renderList(); } });
  $("cycleMonth").value = new Date().toISOString().slice(0, 7); $("generateButton").addEventListener("click", calculate); $("applyArrearsButton").addEventListener("click", applyArrears); $("clearArrearsButton").addEventListener("click", clearArrears); $("saveToDriveButton").addEventListener("click", saveToDriveAndEmail); $("printButton").addEventListener("click", () => window.print());
})();
