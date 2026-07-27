(function () {
  const allowedRoles = new Set(["fribassociate","frccassociate","frshiftincharge","frloader","frpacker","frassociateparttime"]);
  const dayValues = { P: 1, WO: 1, "A-R": 1, "F-R": 1, HD: .5, "HD-R": .5, A: 0, F: 0, L: 0, PENDING: 0 };
  const files = { attendance: null, master: null, structure: null };
  const data = { employees: [], calculations: [] };
  const $ = (id) => document.getElementById(id);
  const normal = (value) => String(value || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  const text = (value) => String(value ?? "").trim();
  const money = (value) => new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: 2 }).format(Number(value || 0));
  const escape = (value) => text(value).replace(/[&<>"']/g, (char) => ({ "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#039;" })[char]);
  const status = (message, kind = "info") => { const el = $("status"); el.textContent = message; el.className = `status ${kind}`; };
  const excelDate = (value) => { if (value instanceof Date) return new Date(value.getFullYear(), value.getMonth(), value.getDate()); if (typeof value === "number") return new Date(Date.UTC(1899, 11, 30) + value * 86400000); const d = new Date(value); return Number.isNaN(d.getTime()) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate()); };
  const daysBetween = (start, end) => Math.round((end - start) / 86400000) + 1;
  function workbookRows(book, preferredSheet, requiredHeaders) {
    const name = preferredSheet && book.SheetNames.find((s) => normal(s) === normal(preferredSheet));
    const sheetNames = name ? [name] : book.SheetNames;
    for (const sheetName of sheetNames) { const rows = XLSX.utils.sheet_to_json(book.Sheets[sheetName], { header: 1, defval: "", raw: true }); for (let i = 0; i < Math.min(rows.length, 15); i++) { const headers = rows[i].map(normal); if (requiredHeaders.every((header) => headers.includes(normal(header)))) return rows.slice(i + 1).map((row) => Object.fromEntries(rows[i].map((header, index) => [normal(header), row[index]]))).filter((row) => Object.values(row).some((v) => text(v))); } } return [];
  }
  async function readFile(file) { return XLSX.read(await file.arrayBuffer(), { type: "array", cellFormula: true, cellDates: true }); }
  function structureMap(book) {
    const map = new Map();
    const add = (key, item) => { if (key) map.set(normal(key), item); };
    book.SheetNames.forEach((sheetName) => {
      const sheet = book.Sheets[sheetName]; const endRow = XLSX.utils.decode_range(sheet["!ref"] || "A1").e.r + 1; const markers = [];
      for (let r = 1; r <= endRow; r++) { const label = text(sheet[`A${r}`]?.v); const key = normal(label); if (key.startsWith("salarystucture") || key.startsWith("salarystructure") || key === "parttimesalary") markers.push({ row: r, title: label }); }
      markers.forEach((marker, index) => {
        const next = markers[index + 1]?.row || endRow + 1; const components = {}; let hasPf = false, hasEsi = false;
        for (let r = marker.row + 1; r < next; r++) { const label = normal(sheet[`A${r}`]?.v); const value = Number(sheet[`B${r}`]?.v || 0); if (label) components[label] = value; if (label === "pf") hasPf = true; if (label === "esi") hasEsi = true; }
        const item = { name: marker.title, basic: components.basicsalary || 0, hra: components.hra || 0, special: components.specialallowance || 0, conveyance: components.conveyanceallowance || 0, professionalTax: components.professionaltax || 0, hasPf, hasEsi, directFixed: false };
        item.monthlyGross = item.basic + item.hra + item.special + item.conveyance; add(marker.title, item);
      });
      for (let r = 1; r <= endRow; r++) { const title = text(sheet[`A${r}`]?.v); if (!/^Stifund\s+\d+$/i.test(title)) continue; const monthlyGross = Number(sheet[`B${r}`]?.v || 0); add(title, { name: title, basic: monthlyGross, hra: 0, special: 0, conveyance: 0, professionalTax: 0, hasPf: false, hasEsi: false, directFixed: true, monthlyGross }); }
    }); return map;
  }
  function renderList() { const list = $("employeeList"); list.innerHTML = data.calculations.map((calc, index) => `<button type="button" data-index="${index}" class="${index === 0 ? "active" : ""}"><strong>${escape(calc.employee.name)}</strong><small>${escape(calc.employee.id)} · Net ${money(calc.net)}</small></button>`).join(""); list.querySelectorAll("button").forEach((button) => button.addEventListener("click", () => { list.querySelectorAll("button").forEach((b) => b.classList.remove("active")); button.classList.add("active"); renderSlip(data.calculations[Number(button.dataset.index)]); })); }
  function renderSlip(calc) { const c = calc; const earningRows = [[c.directFixed ? "Stifund Pay" : "Basic Salary", c.basic], ["HRA", c.hra], ["Special Allowance", c.special], ["Conveyance Allowance", c.conveyance], ["Attendance Bonus", c.bonus]].filter(([, value]) => value > 0).map(([label, value]) => row(label, value)).join(""); const deductionRows = [["Provident Fund", c.pf], ["ESI", c.esi], ["Professional Tax", c.pt]].filter(([, value]) => value > 0).map(([label, value]) => row(label, value)).join("") || `<div class="slip-row"><span>No employee deductions in this structure</span><strong>${money(0)}</strong></div>`; $("payslipPreview").innerHTML = `<div class="slip-head"><div><div class="slip-brand">OMKAR RETAIL VENTURES</div><small>Employee salary statement</small></div><div class="slip-label">PAYSLIP<small>${escape(c.period)}</small></div></div><div class="slip-person"><div><small>EMPLOYEE</small><strong>${escape(c.employee.name)}</strong></div><div><small>EMPLOYEE CODE</small><strong>${escape(c.employee.id)}</strong></div><div><small>ROLE</small><strong>${escape(c.employee.role)}</strong></div><div><small>PAID DAYS</small><strong>${c.paidDays} / ${c.cycleDays}</strong></div></div><div class="slip-tables"><div><span class="earning-head">EARNINGS</span>${earningRows}${total("Gross Earnings", c.gross)}</div><div><span class="earning-head">DEDUCTIONS</span>${deductionRows}${total("Total Deductions", c.deductions)}</div></div><div class="net-pay"><div><small>NET PAYABLE</small><strong>${money(c.net)}</strong></div><div><small>SALARY CYCLE</small><strong>${escape(c.period)}</strong></div></div><p class="note">Attendance: ${c.statusSummary}. Attendance bonus is ₹500 because ${c.leaveCount === 0 ? "there are no L statuses" : "leave was recorded"}. PF and ESI are calculated from prorated fixed salary components where the selected salary structure specifies them.</p>`; }
  const row = (label, value) => `<div class="slip-row"><span>${label}</span><strong>${money(value)}</strong></div>`;
  const total = (label, value) => `<div class="slip-total"><span>${label}</span><strong>${money(value)}</strong></div>`;
  function calculate() {
    if (!files.attendance || !files.master || !files.structure) return status("Please select all three Excel files first.", "error");
    Promise.all([readFile(files.attendance), readFile(files.master), readFile(files.structure)]).then(([attendanceBook, masterBook, structureBook]) => {
      const attendance = workbookRows(attendanceBook, "Attendance", ["employee_code", "scheduled_date", "current_role_name", "muster_status"]);
      const master = workbookRows(masterBook, "Employee_Shift", ["Name", "Z ID", "Salary Structure"]);
      const structures = structureMap(structureBook);
      const month = $("cycleMonth").value; if (!month) return status("Choose the salary-cycle month (the cycle runs from the previous 21st to this month’s 20th).", "error");
      const [year, monthNumber] = month.split("-").map(Number); const start = new Date(year, monthNumber - 2, 21); const end = new Date(year, monthNumber - 1, 20); const cycleDays = daysBetween(start, end);
      const masters = new Map(master.map((r) => [text(r[normal("Z ID")]), { id: text(r[normal("Z ID")]), name: text(r.name), role: text(r.role), structure: text(r[normal("Salary Structure")]) }]));
      const grouped = new Map(); attendance.forEach((r) => { const date = excelDate(r.scheduleddate); const id = text(r.employeecode); if (!date || date < start || date > end || !allowedRoles.has(normal(r.currentrolename)) || !masters.has(id)) return; if (!grouped.has(id)) grouped.set(id, []); grouped.get(id).push(text(r.musterstatus).toUpperCase()); });
      const requested = $("employeeFilter").value; const results = []; const missingStructures = new Set(); grouped.forEach((statuses, id) => { if (requested && requested !== id) return; const employee = masters.get(id); const structure = structures.get(normal(employee.structure)); if (!structure) { missingStructures.add(employee.structure || "(blank)"); return; } const paidDays = statuses.reduce((sum, s) => sum + (dayValues[s] ?? 0), 0); const factor = paidDays / cycleDays; const basic = structure.basic * factor, hra = structure.hra * factor, special = structure.special * factor, conveyance = structure.conveyance * factor; const fixedGross = basic + hra + special + conveyance; const leaveCount = statuses.filter((s) => s === "L").length; const bonus = leaveCount === 0 ? 500 : 0; const pf = structure.hasPf ? basic * .12 : 0, esi = structure.hasEsi ? fixedGross * .0075 : 0, pt = structure.professionalTax * factor; const gross = fixedGross + bonus, deductions = pf + esi + pt; results.push({ employee, period: `${start.toLocaleDateString("en-IN", { day:"2-digit", month:"short", year:"numeric" })} – ${end.toLocaleDateString("en-IN", { day:"2-digit", month:"short", year:"numeric" })}`, cycleDays, paidDays, basic, hra, special, conveyance, bonus, gross, pf, esi, pt, deductions, net: gross - deductions, leaveCount, directFixed: structure.directFixed, statusSummary: [...new Set(statuses)].join(", ") || "No records" }); });
      data.calculations = results.sort((a, b) => a.employee.name.localeCompare(b.employee.name)); const filter = $("employeeFilter"); if (!filter.dataset.loaded) { [...masters.values()].sort((a,b) => a.name.localeCompare(b.name)).forEach((e) => filter.insertAdjacentHTML("beforeend", `<option value="${escape(e.id)}">${escape(e.name)} (${escape(e.id)})</option>`)); filter.dataset.loaded = "1"; }
      if (!results.length) { $("resultArea").hidden = true; return status("No eligible employees were found for this cycle. Check the role names, employee codes, cycle, and salary-structure names.", "error"); }
      $("resultArea").hidden = false; renderList(); renderSlip(results[0]); const skipped = missingStructures.size ? ` ${missingStructures.size} salary-structure reference${missingStructures.size === 1 ? " is" : "s are"} not present in the uploaded salary workbook and were skipped: ${[...missingStructures].join(", ")}.` : ""; status(`${results.length} payslip${results.length === 1 ? "" : "s"} generated for the selected salary cycle.${skipped}`);
    }).catch((error) => status(`Unable to read the files: ${error.message}`, "error"));
  }
  ["attendanceFile", "masterFile", "structureFile"].forEach((id) => $(id).addEventListener("change", (event) => { files[id.replace("File", "")] = event.target.files[0] || null; }));
  $("cycleMonth").value = new Date().toISOString().slice(0, 7); $("generateButton").addEventListener("click", calculate); $("printButton").addEventListener("click", () => window.print());
})();
