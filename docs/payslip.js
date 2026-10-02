(function () {
  const dayValues = { P: 1, WO: 1, "A-R": 1, "F-R": 1, HD: 0.5, "HD-R": 0.5, A: 0, F: 0, L: 0, PENDING: 0 };
  const files = { attendance: null, master: null, structure: null, salaryAdvance: null, incentive: null, fullFinalAttendance: null };
  const data = { calculations: [], activeCalculation: null };
  const arrearsByEmployee = new Map();
  const salaryAdvanceByEmployee = new Map();
  const incentiveByEmployee = new Map();
  const fullFinalByEmployee = new Map();
  const $ = (id) => document.getElementById(id);
  const normal = (value) => String(value || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  const text = (value) => String(value ?? "").trim();
  const firstValue = (...values) => values.find((value) => text(value)) || "";
  const money = (value) => new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: 2 }).format(Number(value || 0));
  const signedMoney = (value) => { const amount = Number(value || 0); return `${amount > 0 ? "+" : ""}${money(amount)}`; };
  const wholeRupees = (value) => Math.round(Number(value || 0));
  const roundedDays = (value) => Math.round(Number(value || 0) * 100) / 100;
  const amountInWords = (value) => { const n = Math.round(Number(value || 0)); const ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"]; const tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]; const underThousand = (number) => { const parts = []; if (number >= 100) { parts.push(`${ones[Math.floor(number / 100)]} Hundred`); number %= 100; } if (number >= 20) { parts.push(tens[Math.floor(number / 10)]); if (number % 10) parts.push(ones[number % 10]); } else if (number) parts.push(ones[number]); return parts.join(" "); }; if (!n) return "Rupees Zero Only"; const parts = []; let remaining = n; [[10000000, "Crore"], [100000, "Lakh"], [1000, "Thousand"]].forEach(([unit, label]) => { if (remaining >= unit) { parts.push(`${underThousand(Math.floor(remaining / unit))} ${label}`); remaining %= unit; } }); if (remaining) parts.push(underThousand(remaining)); return `Rupees ${parts.join(" ")} Only`; };
  const escape = (value) => text(value).replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" })[char]);
  const status = (message, kind = "info") => { const el = $("status"); el.textContent = message; el.className = `status ${kind}`; };
  const calendarDate = (value) => { if (typeof value === "number") { const parts = XLSX.SSF.parse_date_code(Math.floor(value)); return parts ? `${parts.y}-${String(parts.m).padStart(2, "0")}-${String(parts.d).padStart(2, "0")}` : ""; } if (value instanceof Date) return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, "0")}-${String(value.getDate()).padStart(2, "0")}`; const iso = text(value).match(/^(\d{4}-\d{2}-\d{2})/); if (iso) return iso[1]; const slash = text(value).match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})(?:\s|$)/); if (!slash) return ""; const month = Number(slash[1]), day = Number(slash[2]), year = Number(slash[3]); const date = new Date(Date.UTC(year, month - 1, day)); return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day ? `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}` : ""; };
  const displayDate = (value) => { const date = calendarDate(value); if (!date) return "Not available"; const [year, month, day] = date.split("-"); return `${day} ${["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][Number(month) - 1]} ${year}`; };
  const daysBetween = (start, end) => Math.round((end - start) / 86400000) + 1;
  const inclusiveCalendarDays = (startDate, endDate) => { const [startYear, startMonth, startDay] = startDate.split("-").map(Number); const [endYear, endMonth, endDay] = endDate.split("-").map(Number); return Math.round((Date.UTC(endYear, endMonth - 1, endDay) - Date.UTC(startYear, startMonth - 1, startDay)) / 86400000) + 1; };
  const isWorkedAttendanceStatus = (status) => ["P", "HD", "HD-R", "A-R", "F-R"].includes(text(status).toUpperCase());
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

  function salaryAdvanceRows(book) {
    const rows = workbookRows(book, "Salary Advance", [["Employee ID", "Employee_code", "Z ID"], ["Salary Advance", "Advance Amount", "Amount"]]);
    if (!rows.length) throw new Error("The salary advance sheet must contain Employee ID / Employee_code / Z ID and Salary Advance / Advance Amount / Amount columns.");
    const entries = [];
    const skipped = [];
    rows.forEach((row, index) => {
      const employeeId = text(firstValue(row.employeeid, row.employeecode, row.zid));
      const amount = Number(text(firstValue(row.salaryadvance, row.advanceamount, row.amount)).replace(/,/g, ""));
      const sourceDate = text(firstValue(row.date, row.advancedate, row.disbursementdate));
      const date = calendarDate(sourceDate);
      const reason = text(firstValue(row.reason, row.remarks, row.note, "Salary advance recovery"));
      if (!employeeId || !Number.isFinite(amount) || amount <= 0) {
        skipped.push(index + 2);
        return;
      }
      entries.push({ employeeId, amount, date, dateLabel: date ? displayDate(date) : sourceDate, reason });
    });
    return { entries, skipped };
  }

  function incentiveRows(book) {
    const amountFor = (value) => Number(text(value).replace(/,/g, ""));
    const entries = []; const skipped = [];
    for (const sheetName of book.SheetNames) {
      const rows = XLSX.utils.sheet_to_json(book.Sheets[sheetName], { header: 1, defval: "", raw: true });
      for (let headerRow = 0; headerRow < Math.min(rows.length, 15); headerRow += 1) {
        const headers = (rows[headerRow] || []).map(normal);
        const amountColumns = headers.map((header, index) => header === "incentiveamount" || header === "amount" ? index : -1).filter((index) => index >= 0);
        if (!amountColumns.length) continue;
        const employeeIdColumn = headers.findIndex((header) => ["employeeid", "employeecode", "zid"].includes(header));
        const nameColumn = headers.findIndex((header) => ["name", "employeename"].includes(header));
        const typeColumn = headers.findIndex((header) => ["type", "incentivetype", "category"].includes(header));
        const standardLayout = employeeIdColumn >= 0 || nameColumn >= 0;
        const columns = standardLayout ? amountColumns.map((amountColumn) => ({ amountColumn, employeeIdColumn, nameColumn, typeColumn, category: "" })) : amountColumns.map((amountColumn) => {
          const personColumn = amountColumn - 1;
          let category = "";
          for (let row = headerRow - 1; row >= 0; row -= 1) {
            const label = text(rows[row]?.[personColumn] || rows[row]?.[amountColumn]);
            if (label) { category = label; break; }
          }
          return { amountColumn, employeeIdColumn: -1, nameColumn: personColumn, typeColumn: -1, category };
        });
        rows.slice(headerRow + 1).forEach((row, index) => columns.forEach((column) => {
          const employeeId = column.employeeIdColumn >= 0 ? text(row[column.employeeIdColumn]) : "";
          const name = column.nameColumn >= 0 ? text(row[column.nameColumn]) : "";
          const rawAmount = row[column.amountColumn];
          const amount = amountFor(rawAmount);
          if (!employeeId && !name && !text(rawAmount)) return;
          if (!Number.isFinite(amount) || amount <= 0 || (!employeeId && !name)) {
            skipped.push(`${sheetName} row ${headerRow + index + 2}`);
            return;
          }
          entries.push({ employeeId, name, amount, category: text(column.typeColumn >= 0 ? row[column.typeColumn] : column.category) });
        }));
        return { entries, skipped };
      }
    }
    throw new Error("The incentive sheet must contain an Incentive Amount column with an Employee ID or Name column, or use the shared Inbound / Outbound / Outbound PT layout.");
  }

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

  function ecrMonthName(value) {
    const match = text(value).match(/^(\d{4})-(\d{2})$/);
    if (!match) return "Payroll";
    return `${["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"][Number(match[2]) - 1]} ${match[1]}`;
  }

  function ecrRow(calc) {
    const uan = text(calc.employee.uan).replace(/\s/g, "");
    const pfWages = wholeRupees(calc.pfWages);
    const employeePf = wholeRupees(calc.pf);
    const employerEps = Math.min(1250, wholeRupees(Math.min(pfWages, 15000) * 0.0833));
    return [
      uan,
      text(calc.employee.name).toUpperCase(),
      wholeRupees(calc.gross),
      pfWages,
      pfWages,
      employeePf,
      employerEps,
      Math.max(0, employeePf - employerEps),
      roundedDays(Math.max(0, calc.cycleDays - calc.paidDays)),
      0
    ];
  }

  function downloadEcr() {
    const calculations = data.calculations.filter((calc) => calc.pf > 0);
    const withUan = calculations.filter((calc) => /^\d{12}$/.test(text(calc.employee.uan).replace(/\s/g, "")));
    const withoutUan = calculations.filter((calc) => !withUan.includes(calc));
    if (!withUan.length) {
      return status("No PF employee with a valid 12-digit UAN is available for ECR export.", "error");
    }

    const headers = [
      "UAN", "MEMBER NAME AS PER UAN", "GROSS WAGES", "EPF WAGES", "EDLI WAGES",
      "EMPLOYEE PF CONTRIBUTION", "EMPLOYER EPS CONTRIBUTION", "EMPLOYER PF CONTRIBUTION",
      "NCP DAYS (NON-CONTRIBUTORY PERIOD)", "REFUND OF ADVANCE"
    ];
    const sheet = XLSX.utils.aoa_to_sheet([headers, ...withUan.map(ecrRow)]);
    sheet["!cols"] = [
      { wch: 17 }, { wch: 31 }, { wch: 14 }, { wch: 14 }, { wch: 14 },
      { wch: 25 }, { wch: 25 }, { wch: 24 }, { wch: 34 }, { wch: 19 }
    ];
    sheet["!rows"] = [{ hpt: 24 }, ...withUan.map(() => ({ hpt: 15 }))];
    sheet["!autofilter"] = { ref: `A1:J${withUan.length + 1}` };

    const headerStyle = {
      font: { name: "Arial", sz: 10, bold: true, color: { rgb: "FFFFFF" } },
      fill: { patternType: "solid", fgColor: { rgb: "0000FF" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true }
    };
    const textStyle = { font: { name: "Arial", sz: 10 }, alignment: { vertical: "center" } };
    const numberStyle = { font: { name: "Arial", sz: 10 }, alignment: { horizontal: "right", vertical: "center" }, numFmt: "0.##" };
    const wholeNumberStyle = { font: { name: "Arial", sz: 10 }, alignment: { horizontal: "right", vertical: "center" }, numFmt: "0" };
    headers.forEach((_, column) => { sheet[XLSX.utils.encode_cell({ r: 0, c: column })].s = headerStyle; });
    withUan.forEach((_, row) => {
      for (let column = 0; column < headers.length; column += 1) {
        const cell = sheet[XLSX.utils.encode_cell({ r: row + 1, c: column })];
        if (column < 2) cell.s = textStyle;
        else if (column === 8) cell.s = numberStyle;
        else cell.s = wholeNumberStyle;
      }
    });

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, "ReturnSheet_1");
    XLSX.writeFile(workbook, `ECR ${ecrMonthName($("cycleMonth").value)}.xlsx`, { bookType: "xlsx", cellStyles: true });
    const omitted = withoutUan.length ? ` ${withoutUan.length} PF employee${withoutUan.length === 1 ? " was" : "s were"} omitted because a valid UAN is not available.` : "";
    status(`ECR Excel downloaded with ${withUan.length} PF employee${withUan.length === 1 ? "" : "s"}.${omitted}`, "success");
  }

  const row = (label, value) => `<div class="slip-row"><span>${label}</span><strong>${money(value)}</strong></div>`;
  const total = (label, value) => `<div class="slip-total"><span>${label}</span><strong>${money(value)}</strong></div>`;

  function detectedLastWorkingDate(calc) {
    const postCycleWorked = (calc.sourceExtensionDayEntries || []).filter((entry) => isWorkedAttendanceStatus(entry.status));
    const cycleWorked = (calc.sourceDayEntries || []).filter((entry) => isWorkedAttendanceStatus(entry.status));
    const candidates = postCycleWorked.length ? postCycleWorked : cycleWorked;
    return candidates.map((entry) => entry.date).sort().at(-1) || "";
  }

  function recalculateEarnedAmounts(calc, lastWorkingDate = "") {
    const structure = calc.sourceStructure;
    const mainCycleDays = calc.baseCycleDays || calc.cycleDays;
    const mainEntries = calc.sourceDayEntries.filter((entry) => !lastWorkingDate || entry.date <= lastWorkingDate);
    const extensionEntries = lastWorkingDate ? (calc.sourceExtensionDayEntries || []).filter((entry) => entry.date <= lastWorkingDate) : [];
    const mainPaidDays = mainEntries.reduce((sum, entry) => sum + (dayValues[entry.status] ?? 0), 0);
    const extensionPaidDays = extensionEntries.reduce((sum, entry) => sum + (dayValues[entry.status] ?? 0), 0);
    const doublePayDays = mainEntries.filter((entry) => calc.doublePayDateKeys.includes(entry.date) && entry.status === "P").length;
    const mainFactor = mainPaidDays / mainCycleDays;
    const extensionFactor = extensionEntries.length ? extensionPaidDays / calc.extensionCycleDays : 0;
    const paidDays = mainPaidDays + extensionPaidDays;
    calc.paidDays = paidDays;
    calc.cycleDays = mainCycleDays + (extensionEntries.length ? calc.extensionCycleDays : 0);
    calc.period = lastWorkingDate ? `${displayDate(calc.cycleStart)} – ${displayDate(lastWorkingDate)}` : calc.basePeriod;
    calc.doublePayDays = doublePayDays;
    calc.basic = structure.basic * (mainFactor + extensionFactor);
    calc.hra = structure.hra * (mainFactor + extensionFactor);
    calc.special = structure.special * (mainFactor + extensionFactor);
    calc.conveyance = structure.conveyance * (mainFactor + extensionFactor);
    calc.doublePay = calc.monthlyGross * (doublePayDays / mainCycleDays);
    const isPartTime = /part[\s_-]*time/i.test(`${calc.employee.role} ${calc.employee.structure}`);
    calc.bonus = !lastWorkingDate && mainPaidDays === mainCycleDays ? (isPartTime ? 250 : 500) : 0;
    calc.pfWages = structure.basic * (mainFactor + extensionFactor);
    calc.pf = structure.hasPf ? calc.pfWages * 0.12 : 0;
    calc.basePfWages = calc.pfWages;
    calc.basePf = calc.pf;
    const fixedGross = calc.basic + calc.hra + calc.special + calc.conveyance + calc.doublePay;
    calc.esi = structure.hasEsi ? fixedGross * 0.0075 : 0;
    calc.baseGross = fixedGross + calc.bonus;
    calc.baseDeductions = calc.pf + calc.esi;
    const allEntries = [...mainEntries, ...extensionEntries];
    calc.statusSummary = [...new Set(allEntries.map((entry) => entry.status))].join(", ") || "No records";
    calc.countedDates = allEntries.map((entry) => `${entry.date} (${entry.status})`);
    calc.location = firstValue([...new Set(allEntries.flatMap((entry) => entry.locations))].join(", "), calc.employee.location);
  }

  function updateArrearsForCalculation(calc, adjustment = arrearsByEmployee.get(calc.employee.id)) {
    const amount = Number(adjustment?.amount || 0);
    calc.arrears = Number.isFinite(amount) ? amount : 0;
    calc.arrearsDays = Number(adjustment?.days || 0);
    calc.arrearsEarning = Math.max(calc.arrears, 0);
    calc.arrearsRecovery = 0;
    const arrearsCycleDays = Number(calc.baseCycleDays || calc.cycleDays || 0);
    const basePf = Number(calc.basePf ?? calc.pf ?? 0);
    const basePfWages = Number(calc.basePfWages ?? calc.pfWages ?? 0);
    // Arrears are paid on gross, but PF applies only to their Basic-salary portion.
    calc.arrearsPfWages = calc.sourceStructure?.hasPf && arrearsCycleDays > 0
      ? (Number(calc.sourceStructure.basic || 0) / arrearsCycleDays) * Math.max(0, calc.arrearsDays)
      : 0;
    calc.arrearsPf = calc.arrearsPfWages * 0.12;
    calc.pfWages = basePfWages + calc.arrearsPfWages;
    calc.pf = basePf + calc.arrearsPf;
    const savedSalaryAdvance = salaryAdvanceByEmployee.get(calc.employee.id);
    const salaryAdvance = Number(savedSalaryAdvance?.amount || 0);
    calc.salaryAdvance = Number.isFinite(salaryAdvance) && salaryAdvance > 0 ? salaryAdvance : 0;
    calc.salaryAdvanceReason = text(savedSalaryAdvance?.reason);
    calc.salaryAdvanceEntries = Array.isArray(savedSalaryAdvance?.entries) ? savedSalaryAdvance.entries : (calc.salaryAdvance ? [{ amount: calc.salaryAdvance, reason: calc.salaryAdvanceReason, date: "" }] : []);
    const savedIncentive = incentiveByEmployee.get(calc.employee.id);
    const incentive = Number(savedIncentive?.amount || 0);
    calc.incentive = Number.isFinite(incentive) && incentive > 0 ? incentive : 0;
    calc.incentiveEntries = Array.isArray(savedIncentive?.entries) ? savedIncentive.entries : (calc.incentive ? [{ amount: calc.incentive, category: "" }] : []);
    calc.gross = calc.baseGross + calc.arrearsEarning + calc.incentive;
    // Arrears affect PF on their Basic portion only.
    // Professional Tax uses ordinary earnings plus incentive only.
    const deductionEligibleGross = calc.baseGross + calc.incentive;
    calc.pt = deductionEligibleGross > 25000 ? (calc.salaryCycleMonth === 2 ? 300 : 200) : 0;
    calc.baseDeductions = calc.pf + calc.esi + calc.pt;
    const fullFinal = fullFinalByEmployee.get(calc.employee.id);
    calc.lastWorkingDate = text(fullFinal?.lastWorkingDate);
    calc.noticeSubmittedDate = text(fullFinal?.noticeSubmittedDate);
    calc.noticeServedDays = Number(fullFinal?.noticeServedDays || 0);
    calc.noticeRecoveryDays = Number(fullFinal?.noticeRecoveryDays || 0);
    calc.fullFinalNote = text(fullFinal?.note);
    calc.isFullFinal = Boolean(fullFinal);
    const requestedNoticeRecovery = calc.noticeRecoveryDays > 0 ? (calc.monthlyGross / 30) * calc.noticeRecoveryDays : 0;
    const payableBeforeNoticeRecovery = calc.gross - calc.baseDeductions - calc.salaryAdvance;
    calc.noticeRecovery = Math.min(requestedNoticeRecovery, Math.max(0, payableBeforeNoticeRecovery));
    calc.noticeRecoveryBalance = Math.max(0, requestedNoticeRecovery - calc.noticeRecovery);
    calc.deductions = calc.baseDeductions + calc.salaryAdvance + calc.noticeRecovery;
    calc.netBeforeRoundOff = calc.gross - calc.deductions;
    calc.roundOff = Math.round(calc.netBeforeRoundOff) - calc.netBeforeRoundOff;
    if (Math.abs(calc.roundOff) < 0.005) calc.roundOff = 0;
    calc.net = Math.round(calc.netBeforeRoundOff);
  }

  function populateArrearsEmployees(calculations) {
    const select = $("arrearsEmployee");
    select.innerHTML = calculations.map((calc) => `<option value="${escape(calc.employee.id)}">${escape(calc.employee.name)} (${escape(calc.employee.id)})</option>`).join("");
    const salaryAdvanceSelect = $("salaryAdvanceEmployee");
    salaryAdvanceSelect.innerHTML = select.innerHTML;
    const fullFinalSelect = $("fullFinalEmployee");
    fullFinalSelect.innerHTML = select.innerHTML;
    $("arrearsControls").hidden = !calculations.length;
    $("salaryAdvanceControls").hidden = !calculations.length;
    $("fullFinalControls").hidden = !calculations.length;
  }

  function loadArrearsForm(employeeId) {
    const select = $("arrearsEmployee");
    if (!employeeId || ![...select.options].some((option) => option.value === employeeId)) return;
    const calc = data.calculations.find((item) => item.employee.id === employeeId);
    select.value = employeeId;
    $("salaryAdvanceEmployee").value = employeeId;
    const adjustment = arrearsByEmployee.get(employeeId);
    $("arrearsDays").value = adjustment?.days || "";
    const salaryAdvance = salaryAdvanceByEmployee.get(employeeId);
    $("salaryAdvanceAmount").value = salaryAdvance?.amount || "";
    $("salaryAdvanceReason").value = salaryAdvance?.reason || "";
    $("fullFinalEmployee").value = employeeId;
    const fullFinal = fullFinalByEmployee.get(employeeId);
    $("detectedLastWorkingDate").textContent = displayDate(fullFinal?.lastWorkingDate || detectedLastWorkingDate(calc));
    $("noticeSubmittedDate").value = fullFinal?.noticeSubmittedDate || "";
    $("fullFinalNote").value = fullFinal?.note || "";
  }

  function renderSlip(calc) {
    data.activeCalculation = calc;
    const c = calc; const employee = c.employee;
    const detail = (label, value, blankWhenMissing = false) => `<div class="employee-detail"><small>${label}</small><strong>: ${escape(value || (blankWhenMissing ? "" : "Not available"))}</strong></div>`;
    const roleKey = normal(employee.role);
    const designation = ["frparttime", "frwarehouseintern"].includes(roleKey) ? "Warehouse Intern" : employee.role;
    const validUan = /^\d{12}$/.test(text(employee.uan)) ? text(employee.uan) : "";
    const isStipend = c.directFixed || /sti(?:pend|fund)/i.test(text(employee.structure));
    const statutoryDetails = isStipend ? "" : detail("UAN", validUan, true);
    const arrearsLabel = "Arrears";
    const salaryAdvanceLabel = `Salary Advance${c.salaryAdvanceReason && c.salaryAdvanceReason !== "Salary advance recovery" ? ` – ${escape(c.salaryAdvanceReason)}` : ""}`;
    const noticeRecoveryLabel = `Notice Period Recovery – ${c.noticeRecoveryDays} Day${c.noticeRecoveryDays === 1 ? "" : "s"}`;
    const incentiveEarnings = c.incentiveEntries.length ? c.incentiveEntries.map((entry) => [`Incentive${entry.category ? ` – ${escape(entry.category)}` : ""}`, entry.amount]) : [["Incentive", c.incentive]];
    const earnings = [[c.directFixed ? "Stipend Pay" : "Basic Salary", c.basic], ["HRA", c.hra], ["Special Allowance", c.special], ["Conveyance Allowance", c.conveyance], ["Double Pay", c.doublePay], [arrearsLabel, c.arrearsEarning], ...incentiveEarnings, ["Attendance Bonus", c.bonus]].filter(([, value]) => value > 0).map(([label, value]) => row(label, value)).join("");
    const deductions = [["Provident Fund", c.pf], ["ESI", c.esi], ["Professional Tax", c.pt], [salaryAdvanceLabel, c.salaryAdvance], [noticeRecoveryLabel, c.noticeRecovery]].filter(([, value]) => value > 0).map(([label, value]) => row(label, value)).join("");
    const fullFinalDetails = c.isFullFinal ? `${detail("LAST WORKING DATE", displayDate(c.lastWorkingDate))}${c.noticeSubmittedDate ? detail("NOTICE SUBMITTED", displayDate(c.noticeSubmittedDate)) : ""}${detail("NOTICE SERVED", `${c.noticeServedDays} of 15 days`)}${detail("NOTICE RECOVERY", `${c.noticeRecoveryDays} day${c.noticeRecoveryDays === 1 ? "" : "s"}`)}${c.fullFinalNote ? detail("F&F REFERENCE", c.fullFinalNote) : ""}` : "";
    const balanceRecovery = c.noticeRecoveryBalance > 0 ? `<p class="note"><strong>Balance recoverable:</strong> ${money(c.noticeRecoveryBalance)}. This amount is not deducted from this payslip.</p>` : "";
    $("payslipPreview").innerHTML = `<div class="slip-head"><div class="slip-brand">OMKAR RETAIL VENTURES</div><div class="statement-period">${c.isFullFinal ? "Full &amp; Final Settlement" : "Salary Statement"} for ${escape(c.period)}</div></div><div class="slip-person"><div class="employee-column">${detail("EMPLOYEE NAME", employee.name)}${detail("EMPLOYEE ID", employee.id)}${detail("LOCATION", c.location)}${detail("DESIGNATION", designation)}${detail("DAYS WORKED", `${c.paidDays} / ${c.cycleDays}`)}${c.doublePayDays ? detail("DOUBLE-PAY DAYS", c.doublePayDays) : ""}${fullFinalDetails}</div><div class="employee-column">${detail("PAN", employee.pan)}${statutoryDetails}${detail("BANK NAME", employee.bank)}${detail("BANK ACCOUNT NUMBER", employee.accountNumber)}${detail("DATE OF JOINING", displayDate(employee.doj))}</div></div><div class="slip-tables"><div class="pay-table"><div class="table-heading"><span>PARTICULARS</span><span>EARNINGS</span></div>${earnings}${total("GROSS EARNINGS", c.gross)}</div><div class="pay-table"><div class="table-heading"><span>PARTICULARS</span><span>DEDUCTIONS</span></div>${deductions}${total("TOTAL DEDUCTIONS", c.deductions)}</div></div><div class="net-pay"><span>NET PAY</span><strong>${money(c.net)}</strong></div><div class="round-off"><span>Round Off</span><strong>${signedMoney(c.roundOff)}</strong></div><div class="net-words">(${escape(amountInWords(c.net))})</div>${balanceRecovery}<p class="note">* This is a system-generated payslip and is confidential; therefore no signature is required.</p>`;
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

  async function createPayslipPdf(calc) {
    renderSlip(calc);
    await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    await new Promise((resolve) => window.setTimeout(resolve, 150));
    const preview = $("payslipPreview");
    if (!preview?.textContent.trim()) throw new Error("The payslip preview is empty. Generate the payslip again before saving.");
    const pdfDataUri = await window.html2pdf().set({
      margin: 6,
      filename: payslipFileName(calc),
      image: { type: "jpeg", quality: 0.98 },
      html2canvas: { scale: 2, useCORS: true, backgroundColor: "#ffffff", logging: false },
      jsPDF: { unit: "mm", format: "a4", orientation: "landscape" },
      pagebreak: { mode: ["avoid-all", "css", "legacy"] }
    }).from(preview).outputPdf("datauristring");
    if (!pdfDataUri.includes(",")) throw new Error("The payslip PDF could not be created. Please try again.");
    return pdfDataUri;
  }

  async function deliverPayslip(deliveryMode) {
    const config = window.OMKAR_SITE_CONFIG || {};
    const endpoint = text(config.payslipDeliveryWebAppUrl);
    const calc = data.activeCalculation;
    const sendingEmail = deliveryMode === "email";
    if (!endpoint) return status(`${sendingEmail ? "Email" : "Drive"} delivery is not connected yet. Add the deployed Payslip Delivery Apps Script URL in config.js.`, "error");
    if (!calc || !$("payslipPreview").innerHTML) return status("Generate and select a payslip first.", "error");
    if (typeof window.html2pdf !== "function") return status("PDF delivery support could not be loaded. Check your internet connection and try again.", "error");
    if (sendingEmail && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(text(calc.employee.email))) return status("This employee does not have a valid email address in the uploaded master sheet.", "error");
    try {
      status(`Creating the payslip PDF and sending the ${sendingEmail ? "email" : "Drive"} request...`);
      const idToken = await getOwnerDeliveryToken();
      const fileName = payslipFileName(calc);
      const pdfDataUri = await createPayslipPdf(calc);
      const payload = {
        idToken,
        fileName,
        employeeName: calc.employee.name,
        employeeId: calc.employee.id,
        employeeEmail: calc.employee.email || "",
        salaryCycle: calc.period,
        salaryMonth: $("cycleMonth").value,
        deliveryMode,
        pdfBase64: pdfDataUri.split(",")[1]
      };
      await fetch(endpoint, { method: "POST", mode: "no-cors", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify(payload) });
      status(sendingEmail ? "Payslip email request sent." : "Payslip save-to-Drive request sent.", "success");
    } catch (error) {
      console.error("Payslip delivery failed", error);
      status(error.message || `Unable to ${sendingEmail ? "email" : "save"} the payslip.`, "error");
    }
  }

  async function saveAllPayslipsToDrive() {
    const config = window.OMKAR_SITE_CONFIG || {};
    const endpoint = text(config.payslipDeliveryWebAppUrl);
    const calculations = [...data.calculations];
    const saveButton = $("saveToDriveButton");
    if (!endpoint) return status("Drive delivery is not connected yet. Add the deployed Payslip Delivery Apps Script URL in config.js.", "error");
    if (!calculations.length) return status("Generate the payslips first.", "error");
    if (typeof window.html2pdf !== "function") return status("PDF delivery support could not be loaded. Check your internet connection and try again.", "error");

    const originalCalculation = data.activeCalculation;
    try {
      saveButton.disabled = true;
      const idToken = await getOwnerDeliveryToken();
      // Apps Script accepts limited request sizes. Send small batches instead of
      // sending every PDF in one large request, which can fail silently.
      const batchSize = 5;
      let sent = 0;
      for (let offset = 0; offset < calculations.length; offset += batchSize) {
        const batchCalculations = calculations.slice(offset, offset + batchSize);
        const payslips = [];
        for (let index = 0; index < batchCalculations.length; index += 1) {
          const calc = batchCalculations[index];
          status(`Preparing payslip ${offset + index + 1} of ${calculations.length} for Drive...`);
          const fileName = payslipFileName(calc);
          const pdfDataUri = await createPayslipPdf(calc);
          payslips.push({ fileName, employeeName: calc.employee.name, employeeId: calc.employee.id, pdfBase64: pdfDataUri.split(",")[1] });
        }
        status(`Sending payslips ${offset + 1}–${offset + payslips.length} of ${calculations.length} to Drive...`);
        await fetch(endpoint, { method: "POST", mode: "no-cors", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify({ idToken, salaryMonth: $("cycleMonth").value, deliveryMode: "drive-batch", payslips }) });
        sent += payslips.length;
      }
      status(`${sent} payslip save request${sent === 1 ? " was" : "s were"} sent to Drive in small batches. Open the month folder in a few moments to confirm the files are there.`, "success");
    } catch (error) {
      console.error("Batch Drive save failed", error);
      status(error.message || "Unable to prepare the payslips for Drive.", "error");
    } finally {
      if (originalCalculation) renderSlip(originalCalculation);
      renderList();
      saveButton.disabled = false;
    }
  }

  function calculate() {
    if (!files.attendance || !files.master || !files.structure) return status("Please select all three Excel files first.", "error");
    const inputBooks = [readFile(files.attendance), readFile(files.master), readFile(files.structure), files.salaryAdvance ? readFile(files.salaryAdvance) : Promise.resolve(null), files.incentive ? readFile(files.incentive) : Promise.resolve(null), files.fullFinalAttendance ? readFile(files.fullFinalAttendance) : Promise.resolve(null)];
    Promise.all(inputBooks).then(([attendanceBook, masterBook, structureBook, salaryAdvanceBook, incentiveBook, fullFinalAttendanceBook]) => {
      const attendance = workbookRows(attendanceBook, "Attendance", ["employee_code", "scheduled_date", "current_role_name", "muster_status"]);
      const fullFinalAttendance = fullFinalAttendanceBook ? workbookRows(fullFinalAttendanceBook, "Attendance", ["employee_code", "scheduled_date", "current_role_name", "muster_status"]) : [];
      const master = buildMasterMap(workbookRows(masterBook, "Employee_Shift", ["Name", ["Employee ID", "Z ID"], "Salary Structure"]), workbookRows(masterBook, "Master", ["Name", ["Employee ID", "Z ID"], "Salary Structure"]));
      let uploadedAdvanceEntries = [];
      let skippedAdvanceRows = [];
      if (salaryAdvanceBook) {
        const uploadedAdvances = salaryAdvanceRows(salaryAdvanceBook);
        skippedAdvanceRows = uploadedAdvances.skipped;
        uploadedAdvanceEntries = uploadedAdvances.entries;
      }
      let uploadedIncentiveEntries = [];
      let skippedIncentiveRows = [];
      if (incentiveBook) {
        const uploadedIncentives = incentiveRows(incentiveBook);
        skippedIncentiveRows = uploadedIncentives.skipped;
        uploadedIncentiveEntries = uploadedIncentives.entries;
      }
      const filter = $("employeeFilter");
      if (!filter.dataset.loaded) { [...master.values()].sort((a, b) => a.name.localeCompare(b.name)).forEach((employee) => filter.insertAdjacentHTML("beforeend", `<option value="${escape(employee.id)}">${escape(employee.name)} (${escape(employee.id)})</option>`)); filter.dataset.loaded = "1"; }
      const structures = structureMap(structureBook); const month = $("cycleMonth").value;
      if (!month) return status("Choose the salary-cycle month (the cycle runs from the previous 21st to this month’s 20th).", "error");
      const [year, monthNumber] = month.split("-").map(Number); const start = new Date(year, monthNumber - 2, 21); const end = new Date(year, monthNumber - 1, 20); const cycleDays = daysBetween(start, end);
      const extensionStart = new Date(year, monthNumber - 1, 21); const extensionEnd = new Date(year, monthNumber, 20); const extensionCycleDays = daysBetween(extensionStart, extensionEnd);
      const requested = filter.value;
      const doublePayInput = parseDoublePayDates($("doublePayDates").value);
      if (doublePayInput.invalid.length) return status("Enter double-pay dates in DD-MM-YYYY format, separated by commas.", "error");
      const doublePayDates = doublePayInput.dates;
      const grouped = new Map(); const extensionGrouped = new Map();
      const cycleStart = calendarDate(start); const cycleEnd = calendarDate(end);
      const extensionCycleStart = calendarDate(extensionStart); const extensionCycleEnd = calendarDate(extensionEnd);
      let uploadedAdvanceCount = 0;
      let unmatchedAdvanceCount = 0;
      let uploadedIncentiveCount = 0;
      let unmatchedIncentiveCount = 0;
      const unmatchedIncentiveEmployees = [];
      if (salaryAdvanceBook) {
        salaryAdvanceByEmployee.clear();
        uploadedAdvanceEntries.forEach((advance) => {
          if (!master.has(advance.employeeId)) {
            unmatchedAdvanceCount += 1;
            return;
          }
          const existing = salaryAdvanceByEmployee.get(advance.employeeId);
          salaryAdvanceByEmployee.set(advance.employeeId, {
            amount: (existing?.amount || 0) + advance.amount,
            reason: "Salary advance recovery",
            entries: [...(existing?.entries || []), advance]
          });
        });
        uploadedAdvanceCount = salaryAdvanceByEmployee.size;
      }
      if (incentiveBook) {
        const employeesByName = new Map();
        master.forEach((employee) => {
          const key = normal(employee.name);
          if (!key) return;
          employeesByName.set(key, employeesByName.has(key) ? null : employee.id);
        });
        incentiveByEmployee.clear();
        uploadedIncentiveEntries.forEach((incentive) => {
          const employeeId = incentive.employeeId && master.has(incentive.employeeId) ? incentive.employeeId : employeesByName.get(normal(incentive.name));
          if (!employeeId || !master.has(employeeId)) {
            unmatchedIncentiveCount += 1;
            unmatchedIncentiveEmployees.push(incentive.employeeId || incentive.name || "unknown employee");
            return;
          }
          const existing = incentiveByEmployee.get(employeeId);
          incentiveByEmployee.set(employeeId, { amount: (existing?.amount || 0) + incentive.amount, entries: [...(existing?.entries || []), incentive] });
        });
        uploadedIncentiveCount = incentiveByEmployee.size;
      }
      const outsideCycleDates = [...doublePayDates].filter((date) => date < cycleStart || date > cycleEnd);
      if (outsideCycleDates.length) return status(`Double-pay date${outsideCycleDates.length === 1 ? "" : "s"} must fall within this salary cycle: ${outsideCycleDates.join(", ")}.`, "error");
      const addAttendance = (records, target, startDate, endDate) => records.forEach((record) => {
        const dateKey = calendarDate(record.scheduleddate); const id = text(record.employeecode);
        const attendanceRole = text(record.currentrolename);
        if (!dateKey || dateKey < startDate || dateKey > endDate || !master.has(id) || normal(attendanceRole) === "flexcity" || !/^FR_/i.test(attendanceRole)) return;
        if (!target.has(id)) target.set(id, new Map()); const dates = target.get(id);
        if (!dates.has(dateKey)) dates.set(dateKey, { statuses: [], locations: [] });
        const entry = dates.get(dateKey); entry.statuses.push(text(record.musterstatus).toUpperCase()); if (text(record.storename)) entry.locations.push(text(record.storename));
      });
      addAttendance(attendance, grouped, cycleStart, cycleEnd);
      if (fullFinalAttendanceBook) addAttendance(fullFinalAttendance, extensionGrouped, extensionCycleStart, extensionCycleEnd);
      const results = []; const missingStructures = new Set(); const conflicts = [];
      grouped.forEach((dates, id) => {
        if (requested && requested !== id) return;
        const employee = master.get(id); const structure = structures.get(normal(employee.structure));
        if (!structure) { missingStructures.add(employee.structure || "(blank)"); return; }
        const dayEntries = [...dates.entries()].sort(([a], [b]) => a.localeCompare(b)); const conflictingDates = dayEntries.filter(([, entry]) => new Set(entry.statuses).size > 1);
        const extensionDates = extensionGrouped.get(id) || new Map();
        const extensionDayEntries = [...extensionDates.entries()].sort(([a], [b]) => a.localeCompare(b)); const extensionConflicts = extensionDayEntries.filter(([, entry]) => new Set(entry.statuses).size > 1);
        if (conflictingDates.length || extensionConflicts.length) { conflicts.push(`${employee.name} (${[...conflictingDates, ...extensionConflicts].map(([date]) => date).join(", ")})`); return; }
        const statuses = dayEntries.map(([, entry]) => entry.statuses[0]); const paidDays = statuses.reduce((sum, value) => sum + (dayValues[value] ?? 0), 0);
        const employeeDoublePayDays = dayEntries.filter(([date, entry]) => doublePayDates.has(date) && entry.statuses[0] === "P").length; const factor = paidDays / cycleDays; const doublePayFactor = employeeDoublePayDays / cycleDays;
        const basic = structure.basic * factor, hra = structure.hra * factor, special = structure.special * factor, conveyance = structure.conveyance * factor;
        const doublePay = (structure.basic + structure.hra + structure.special + structure.conveyance) * doublePayFactor; const fixedGross = basic + hra + special + conveyance + doublePay;
        const isPartTime = /part[\s_-]*time/i.test(`${employee.role} ${employee.structure}`);
        const bonus = paidDays === cycleDays ? (isPartTime ? 250 : 500) : 0;
        // PF applies only to earned Basic Salary. Double Pay remains an earning, not a PF wage.
        const pfBasic = structure.basic * factor;
        const pf = structure.hasPf ? pfBasic * 0.12 : 0;
        const esi = structure.hasEsi ? fixedGross * 0.0075 : 0;
        const baseGross = fixedGross + bonus;
        const pt = 0;
        const baseDeductions = pf + esi;
        const locations = [...new Set(dayEntries.flatMap(([, entry]) => entry.locations))];
        const monthlyGross = structure.basic + structure.hra + structure.special + structure.conveyance;
        const basePeriod = `${start.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" })} – ${end.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" })}`;
        const calculation = { employee, location: firstValue(locations.join(", "), employee.location), period: basePeriod, basePeriod, cycleStart, cycleEnd, cycleDays, baseCycleDays: cycleDays, extensionCycleStart, extensionCycleEnd, extensionCycleDays, paidDays, doublePayDays: employeeDoublePayDays, basic, hra, special, conveyance, doublePay, bonus, baseGross, baseDeductions, gross: baseGross, pf, basePf: pf, pfWages: pfBasic, basePfWages: pfBasic, esi, pt, salaryCycleMonth: monthNumber, monthlyGross, deductions: baseDeductions, net: baseGross - baseDeductions, directFixed: structure.directFixed, statusSummary: [...new Set(statuses)].join(", ") || "No records", countedDates: dayEntries.map(([date, entry]) => `${date} (${entry.statuses[0]})`), sourceStructure: structure, sourceDayEntries: dayEntries.map(([date, entry]) => ({ date, status: entry.statuses[0], locations: entry.locations })), sourceExtensionDayEntries: extensionDayEntries.map(([date, entry]) => ({ date, status: entry.statuses[0], locations: entry.locations })), doublePayDateKeys: [...doublePayDates] };
        const savedFullFinal = fullFinalByEmployee.get(employee.id);
        if (savedFullFinal?.lastWorkingDate) recalculateEarnedAmounts(calculation, savedFullFinal.lastWorkingDate);
        updateArrearsForCalculation(calculation);
        results.push(calculation);
      });
      data.calculations = results.sort((a, b) => a.employee.name.localeCompare(b.employee.name));
      if (!results.length) { $("resultArea").hidden = true; return status("No matched employees were found for this cycle. Check employee IDs, attendance dates, and salary-structure names.", "error"); }
      $("resultArea").hidden = false; populateArrearsEmployees(data.calculations); renderSlip(results[0]); renderList(); loadArrearsForm(results[0].employee.id);
      const skipped = missingStructures.size ? ` ${missingStructures.size} salary-structure reference${missingStructures.size === 1 ? " is" : "s are"} not present in the uploaded salary workbook and were skipped: ${[...missingStructures].join(", ")}.` : "";
      const conflictNote = conflicts.length ? ` ${conflicts.length} employee${conflicts.length === 1 ? " has" : "s have"} conflicting attendance records and ${conflicts.length === 1 ? "was" : "were"} blocked for review: ${conflicts.join("; ")}.` : "";
      const doublePayNote = doublePayDates.size ? ` Double pay was added for employees marked P on: ${[...doublePayDates].map(displayDate).join(", ")}.` : "";
      const advanceNote = salaryAdvanceBook ? ` Salary advances applied for ${uploadedAdvanceCount} employee${uploadedAdvanceCount === 1 ? "" : "s"}. All uploaded advances are recovered in this selected salary cycle.${unmatchedAdvanceCount ? ` ${unmatchedAdvanceCount} advance record${unmatchedAdvanceCount === 1 ? "" : "s"} did not match an Employee ID in the master sheet.` : ""}${skippedAdvanceRows.length ? ` Rows ${skippedAdvanceRows.join(", ")} were skipped because the Employee ID or advance amount is missing/invalid.` : ""}` : "";
      const incentiveNote = incentiveBook ? ` Incentives applied for ${uploadedIncentiveCount} employee${uploadedIncentiveCount === 1 ? "" : "s"}.${unmatchedIncentiveCount ? ` ${unmatchedIncentiveCount} incentive record${unmatchedIncentiveCount === 1 ? "" : "s"} could not be matched to the master sheet: ${[...new Set(unmatchedIncentiveEmployees)].join(", ")}. Add the Employee ID to the incentive sheet or correct the name.` : ""}${skippedIncentiveRows.length ? ` ${skippedIncentiveRows.join(", ")} ${skippedIncentiveRows.length === 1 ? "was" : "were"} skipped because the employee or incentive amount is missing/invalid.` : ""}` : "";
      status(`${results.length} payslip${results.length === 1 ? "" : "s"} generated for the selected salary cycle.${doublePayNote}${advanceNote}${incentiveNote}${skipped}${conflictNote}`);
    }).catch((error) => status(`Unable to read the files: ${error.message}`, "error"));
  }

  function applyArrears() {
    const employeeId = text($("arrearsEmployee").value);
    const daysText = text($("arrearsDays").value);
    const days = Number(daysText);
    const calc = data.calculations.find((item) => item.employee.id === employeeId);
    if (!calc) return status("Generate the employee payslip before applying arrears.", "error");
    if (!daysText || !Number.isFinite(days) || days <= 0) return status("Enter arrears days greater than zero.", "error");
    const cycleDays = Number(calc.baseCycleDays || calc.cycleDays || 0);
    if (!Number.isFinite(cycleDays) || cycleDays <= 0) return status("Unable to determine the salary-cycle days for arrears.", "error");
    const amount = (Number(calc.monthlyGross || 0) / cycleDays) * days;
    if (!Number.isFinite(amount) || amount <= 0) return status("Unable to calculate arrears from this employee's fixed gross salary.", "error");
    arrearsByEmployee.set(employeeId, { days, amount });
    updateArrearsForCalculation(calc);
    renderSlip(calc); renderList(); loadArrearsForm(employeeId);
    status(`Arrears of ${money(amount)} for ${days} day${days === 1 ? "" : "s"} applied for ${calc.employee.name}. PF was calculated only on the Basic portion of the arrears.`, "success");
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

  function applySalaryAdvance() {
    const employeeId = text($("salaryAdvanceEmployee").value);
    const amountText = text($("salaryAdvanceAmount").value);
    const amount = Number(amountText);
    const reason = text($("salaryAdvanceReason").value);
    const calc = data.calculations.find((item) => item.employee.id === employeeId);
    if (!calc) return status("Generate the employee payslip before applying a salary advance deduction.", "error");
    if (!amountText || !Number.isFinite(amount) || amount <= 0) return status("Enter a salary advance amount greater than zero.", "error");
    if (!reason) return status("Enter the reason for the salary advance deduction.", "error");
    salaryAdvanceByEmployee.set(employeeId, { amount, reason, entries: [{ amount, reason, date: "" }] });
    updateArrearsForCalculation(calc);
    renderSlip(calc); renderList(); loadArrearsForm(employeeId);
    status(`Salary advance deduction of ${money(amount)} applied for ${calc.employee.name}.`, "success");
  }

  function clearSalaryAdvance() {
    const employeeId = text($("salaryAdvanceEmployee").value);
    const calc = data.calculations.find((item) => item.employee.id === employeeId);
    if (!calc) return status("Generate the employee payslip before removing a salary advance deduction.", "error");
    salaryAdvanceByEmployee.delete(employeeId);
    updateArrearsForCalculation(calc);
    renderSlip(calc); renderList(); loadArrearsForm(employeeId);
    status(`Salary advance deduction removed for ${calc.employee.name}.`, "success");
  }

  function applyFullFinal() {
    const employeeId = text($("fullFinalEmployee").value);
    const calc = data.calculations.find((item) => item.employee.id === employeeId);
    const lastWorkingDate = calc ? detectedLastWorkingDate(calc) : "";
    const noticeSubmittedDate = text($("noticeSubmittedDate").value);
    const note = text($("fullFinalNote").value);
    if (!calc) return status("Generate the employee payslip before applying Full & Final settlement.", "error");
    if (!lastWorkingDate) return status("No paid attendance record was found to determine the last working date.", "error");
    if (noticeSubmittedDate && noticeSubmittedDate > lastWorkingDate) return status("Notice submitted date cannot be after the last working date.", "error");
    const noticeServedDays = noticeSubmittedDate ? Math.max(0, inclusiveCalendarDays(noticeSubmittedDate, lastWorkingDate)) : 0;
    const noticeRecoveryDays = Math.max(0, 15 - Math.min(15, noticeServedDays));
    fullFinalByEmployee.set(employeeId, { lastWorkingDate, noticeSubmittedDate, noticeServedDays, noticeRecoveryDays, note });
    recalculateEarnedAmounts(calc, lastWorkingDate);
    updateArrearsForCalculation(calc);
    renderSlip(calc); renderList(); loadArrearsForm(employeeId);
    const requested = (calc.monthlyGross / 30) * noticeRecoveryDays;
    status(`Full & Final applied for ${calc.employee.name}. Last working date: ${displayDate(lastWorkingDate)}. Notice served: ${noticeServedDays} of 15 calendar days; recovery: ${noticeRecoveryDays} day${noticeRecoveryDays === 1 ? "" : "s"}. Requested: ${money(requested)}; deducted from this payslip: ${money(calc.noticeRecovery)}.`, "success");
  }

  function clearFullFinal() {
    const employeeId = text($("fullFinalEmployee").value);
    const calc = data.calculations.find((item) => item.employee.id === employeeId);
    if (!calc) return status("Generate the employee payslip before removing Full & Final settlement.", "error");
    fullFinalByEmployee.delete(employeeId);
    recalculateEarnedAmounts(calc);
    updateArrearsForCalculation(calc);
    renderSlip(calc); renderList(); loadArrearsForm(employeeId);
    status(`Full & Final settlement removed for ${calc.employee.name}.`, "success");
  }

  ["attendanceFile", "masterFile", "structureFile", "salaryAdvanceFile", "incentiveFile", "fullFinalAttendanceFile"].forEach((id) => $(id).addEventListener("change", (event) => { files[id.replace("File", "")] = event.target.files[0] || null; }));
  $("arrearsEmployee").addEventListener("change", (event) => { const calc = data.calculations.find((item) => item.employee.id === event.target.value); loadArrearsForm(event.target.value); if (calc) { renderSlip(calc); renderList(); } });
  $("salaryAdvanceEmployee").addEventListener("change", (event) => { const calc = data.calculations.find((item) => item.employee.id === event.target.value); loadArrearsForm(event.target.value); if (calc) { renderSlip(calc); renderList(); } });
  $("fullFinalEmployee").addEventListener("change", (event) => { const calc = data.calculations.find((item) => item.employee.id === event.target.value); loadArrearsForm(event.target.value); if (calc) { renderSlip(calc); renderList(); } });
  $("cycleMonth").value = new Date().toISOString().slice(0, 7); $("generateButton").addEventListener("click", calculate); $("applyArrearsButton").addEventListener("click", applyArrears); $("clearArrearsButton").addEventListener("click", clearArrears); $("applySalaryAdvanceButton").addEventListener("click", applySalaryAdvance); $("clearSalaryAdvanceButton").addEventListener("click", clearSalaryAdvance); $("applyFullFinalButton").addEventListener("click", applyFullFinal); $("clearFullFinalButton").addEventListener("click", clearFullFinal); $("downloadEcrButton").addEventListener("click", downloadEcr); $("saveToDriveButton").addEventListener("click", saveAllPayslipsToDrive); $("emailPayslipButton").addEventListener("click", () => deliverPayslip("email")); $("printButton").addEventListener("click", () => window.print());
})();
