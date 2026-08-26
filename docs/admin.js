(function () {
  const config = window.OMKAR_CONFIG || {};
  const keys = {
    applications: "omkar_applications",
    employerLeads: "omkar_employer_leads",
    contacts: "omkar_contacts",
    supports: "omkar_supports"
  };

  let firebaseLoaded = false;
  let firebaseApp = null;
  let firebaseAuth = null;
  let dashboardRendered = false;
  let onboardingProfiles = [];
  let currentAdminRole = "";
  let currentOnboardingFilter = "active";

  const els = {
    accessForm: document.getElementById("admin-access-form"),
    emailInput: document.getElementById("admin-email"),
    pinInput: document.getElementById("admin-pin"),
    accessNote: document.getElementById("admin-access-note"),
    resetPassword: document.getElementById("admin-reset-password"),
    logout: document.getElementById("admin-logout"),
    dashboard: document.getElementById("admin-dashboard"),
    totalApplications: document.getElementById("kpi-applications"),
    totalEmployers: document.getElementById("kpi-employers"),
    totalContacts: document.getElementById("kpi-contacts"),
    totalSupports: document.getElementById("kpi-supports"),
    applicationsSheetLink: document.getElementById("open-applications-sheet"),
    resumeFolderLink: document.getElementById("open-resume-folder"),
    onboardingBody: document.getElementById("onboarding-body"),
    onboardingNote: document.getElementById("onboarding-admin-note"),
    refreshOnboarding: document.getElementById("refresh-onboarding"),
    onboardingReportFields: document.getElementById("onboarding-report-fields"),
    showOnboardingReport: document.getElementById("show-onboarding-report"),
    downloadOnboardingReport: document.getElementById("download-onboarding-report"),
    onboardingReportSearch: document.getElementById("onboarding-report-search"),
    onboardingReportStatus: document.getElementById("onboarding-report-status"),
    onboardingReportRole: document.getElementById("onboarding-report-role"),
    onboardingReportShift: document.getElementById("onboarding-report-shift"),
    onboardingReportJoinFrom: document.getElementById("onboarding-report-join-from"),
    onboardingReportJoinTo: document.getElementById("onboarding-report-join-to"),
    onboardingReportCount: document.getElementById("onboarding-report-count"),
    onboardingReportPreviewHead: document.getElementById("onboarding-report-preview-head"),
    onboardingReportPreviewBody: document.getElementById("onboarding-report-preview-body"),
    applicationsBody: document.getElementById("applications-body"),
    employerBody: document.getElementById("employers-body"),
    contactsBody: document.getElementById("contacts-body"),
    supportBody: document.getElementById("supports-body")
  };

  function getFirebaseConfig() {
    const fb = config.firebase || {};
    const required = ["apiKey", "authDomain", "projectId", "storageBucket", "appId"];
    if (!required.every((key) => Boolean(fb[key]))) {
      throw new Error("Firebase is not configured.");
    }
    return fb;
  }

  async function getFirebaseApp() {
    if (firebaseApp) return firebaseApp;
    const fb = getFirebaseConfig();
    const { initializeApp, getApps } = await import("https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js");
    firebaseApp = getApps()[0] || initializeApp(fb);
    return firebaseApp;
  }

  async function getFirebaseAuth() {
    if (firebaseAuth) return firebaseAuth;
    const [app, authMod] = await Promise.all([
      getFirebaseApp(),
      import("https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js")
    ]);
    firebaseAuth = authMod.getAuth(app);
    await authMod.setPersistence(firebaseAuth, authMod.browserSessionPersistence);
    if (localStorage.getItem("omkar_force_fresh_login") === "1") {
      await authMod.signOut(firebaseAuth);
      localStorage.removeItem("omkar_force_fresh_login");
    }
    return firebaseAuth;
  }

  function showAccessNote(message, type) {
    if (!els.accessNote) return;
    els.accessNote.textContent = message;
    els.accessNote.className = `note ${type || ""}`.trim();
  }

  function isAuthorizedAdmin(user) {
    const email = String(user?.email || "").trim().toLowerCase();
    return Boolean(email) && Boolean(config.adminRoles?.[email]);
  }

  function getAdminRole(user) {
    return String(config.adminRoles?.[String(user?.email || "").trim().toLowerCase()] || "");
  }

  function canManageOnboarding() {
    return currentAdminRole === "owner";
  }

  const employeeRoleOptions = [
    "FR_Packer",
    "FR_IB_Associate",
    "FR_CC",
    "FR_Shift_Incharge",
    "Security",
    "FR_Loader",
    "FR_Part_Time"
  ];
  const employeeShiftOptions = {
    "Full time": ["7-4", "8-5", "9-6", "10-7", "11-8", "4-1"],
    "Part Time": ["6-10", "4-8", "5-9", "7-11"]
  };

  function normaliseEmployeeRole(value) {
    const raw = String(value || "").trim();
    const simplified = raw.toLowerCase().replace(/[\s-]+/g, "_");
    const aliases = { "fr_cc_associate": "FR_CC" };
    if (aliases[simplified]) return aliases[simplified];
    const matchedRole = employeeRoleOptions.find((role) => role.toLowerCase() === simplified);
    return matchedRole || "";
  }

  function employeeRoleLabel(row) {
    return normaliseEmployeeRole(row?.assignedRole || row?.designation) || "-";
  }

  function normaliseEmployeeShift(value) {
    const shift = String(value || "").trim();
    return Object.entries(employeeShiftOptions).some(([employmentType, shifts]) =>
      shifts.some((time) => `${employmentType} | ${time}` === shift)
    ) ? shift : "";
  }

  function showOnboardingNote(message, type) {
    if (!els.onboardingNote) return;
    els.onboardingNote.textContent = message || "";
    els.onboardingNote.className = `note ${type || ""}`.trim();
  }

  function getLoginErrorMessage(error) {
    switch (error?.code) {
      case "auth/operation-not-allowed":
        return "Email/password sign-in is not enabled in Firebase yet.";
      case "auth/unauthorized-domain":
        return "This website domain has not been authorised in Firebase yet.";
      case "auth/network-request-failed":
        return "Network connection failed. Please check your internet connection and try again.";
      case "auth/invalid-credential":
      case "auth/user-not-found":
      case "auth/wrong-password":
        return "This email address or password is not recognised. Please check it or use Forgot Password.";
      case "auth/not-authorised-admin":
        return "This account is not authorised to access the dashboard.";
      default:
        return "Unable to sign in. Please check your Firebase settings and try again.";
    }
  }

  function readLocal(key) {
    try {
      return JSON.parse(localStorage.getItem(key) || "[]");
    } catch {
      return [];
    }
  }

  function setCount(el, value) {
    if (el) el.textContent = String(value);
  }

  function createCellRow(values) {
    const tr = document.createElement("tr");
    values.forEach((value) => {
      const td = document.createElement("td");
      td.textContent = value || "-";
      tr.appendChild(td);
    });
    return tr;
  }

  function fillTable(bodyEl, rows, mapper) {
    if (!bodyEl) return;
    bodyEl.innerHTML = "";

    if (!rows.length) {
      bodyEl.appendChild(createCellRow(["No records yet"]));
      return;
    }

    rows.forEach((row) => bodyEl.appendChild(createCellRow(mapper(row))));
  }

  function formatDate(value) {
    if (!value) return "-";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return String(value);
    return date.toLocaleString();
  }

  function setupGoogleAdminLinks() {
    const sheetId = String(config.googleSheetId || "").trim();
    const folderId = String(config.googleDriveResumeFolderId || "").trim();

    if (els.applicationsSheetLink) {
      if (sheetId) {
        els.applicationsSheetLink.href = `https://docs.google.com/spreadsheets/d/${sheetId}/edit`;
      } else {
        els.applicationsSheetLink.hidden = true;
      }
    }

    if (els.resumeFolderLink) {
      if (folderId) {
        els.resumeFolderLink.href = `https://drive.google.com/drive/folders/${folderId}`;
      } else {
        els.resumeFolderLink.hidden = true;
      }
    }
  }

  function toCsv(rows, headers, mapper) {
    const lines = [headers.join(",")];
    rows.forEach((row) => {
      const values = mapper(row).map((v) => `"${String(v || "").replace(/"/g, '""')}"`);
      lines.push(values.join(","));
    });
    return lines.join("\n");
  }

  const onboardingReportColumns = [
    { key: "employeeId", label: "Employee ID", value: (row) => row.employeeId || row.submittedEmployeeId },
    { key: "fullName", label: "Full Name", value: (row) => row.fullName },
    { key: "status", label: "Status", value: (row) => row.status || "pending" },
    { key: "assignedRole", label: "Role", value: (row) => employeeRoleLabel(row) },
    { key: "assignedShift", label: "Shift", value: (row) => normaliseEmployeeShift(row.assignedShift) || "-" },
    { key: "mobile", label: "Mobile Number", value: (row) => row.mobile },
    { key: "personalEmail", label: "Personal Email", value: (row) => row.personalEmail },
    { key: "loginEmail", label: "Login Email", value: (row) => row.email },
    { key: "dateOfJoining", label: "Date of Joining", value: (row) => row.dateOfJoining },
    { key: "dateOfBirth", label: "Date of Birth", value: (row) => row.dateOfBirth },
    { key: "maritalStatus", label: "Marital Status", value: (row) => row.maritalStatus },
    { key: "pan", label: "PAN", value: (row) => row.pan },
    { key: "uan", label: "UAN", value: (row) => row.uan },
    { key: "nomineeName", label: "PF Nominee Name", value: (row) => row.nomineeName },
    { key: "nomineeRelationship", label: "Nominee Relationship", value: (row) => row.nomineeRelationship },
    { key: "bankName", label: "Bank Name", value: (row) => row.bankName },
    { key: "bankAccountNumber", label: "Bank Account Number", value: (row) => row.bankAccountNumber },
    { key: "ifsc", label: "IFSC Code", value: (row) => row.ifsc },
    { key: "currentAddress", label: "Current Address", value: (row) => row.currentAddress },
    { key: "emergencyContactName", label: "Emergency Contact Name", value: (row) => row.emergencyContactName },
    { key: "emergencyContactMobile", label: "Emergency Contact Number", value: (row) => row.emergencyContactMobile },
    { key: "submittedAt", label: "Submitted On", value: (row) => formatDate(row.submittedAt) },
    { key: "approvedAt", label: "Approved On", value: (row) => formatDate(row.approvedAt) },
    { key: "inactiveUntil", label: "Inactive Retention Until", value: (row) => formatDate(row.inactiveUntil) }
  ];

  function renderOnboardingReportFields() {
    if (!els.onboardingReportFields || els.onboardingReportFields.childElementCount) return;
    const initiallySelected = new Set(["employeeId", "fullName", "mobile", "personalEmail", "dateOfJoining", "status", "assignedRole", "assignedShift"]);
    onboardingReportColumns.forEach((column) => {
      const label = document.createElement("label");
      label.style.margin = "0";
      label.style.fontWeight = "600";
      const input = document.createElement("input");
      input.type = "checkbox";
      input.value = column.key;
      input.checked = initiallySelected.has(column.key);
      input.style.cssText = "width:auto; margin:0 7px 0 0; vertical-align:middle;";
      label.append(input, document.createTextNode(column.label));
      els.onboardingReportFields.appendChild(label);
    });
  }

  function setupOnboardingReportFilters() {
    if (els.onboardingReportRole && els.onboardingReportRole.options.length === 1) {
      employeeRoleOptions.forEach((role) => {
        const option = document.createElement("option");
        option.value = role;
        option.textContent = role;
        els.onboardingReportRole.appendChild(option);
      });
    }
    if (els.onboardingReportShift && els.onboardingReportShift.options.length === 0) {
      Object.entries(employeeShiftOptions).forEach(([employmentType, shifts]) => {
        shifts.forEach((time) => {
          const option = document.createElement("option");
          option.value = `${employmentType} | ${time}`;
          option.textContent = `${employmentType} — ${time}`;
          els.onboardingReportShift.appendChild(option);
        });
      });
    }
  }

  function getSelectedOnboardingReportColumns() {
    const selectedKeys = Array.from(els.onboardingReportFields?.querySelectorAll('input[type="checkbox"]:checked') || []).map((input) => input.value);
    return onboardingReportColumns.filter((column) => selectedKeys.includes(column.key));
  }

  function getFilteredOnboardingProfiles() {
    const search = String(els.onboardingReportSearch?.value || "").trim().toLowerCase();
    const status = String(els.onboardingReportStatus?.value || "").trim().toLowerCase();
    const role = String(els.onboardingReportRole?.value || "").trim();
    const selectedShifts = Array.from(els.onboardingReportShift?.selectedOptions || []).map((option) => option.value).filter(Boolean);
    const joiningFrom = String(els.onboardingReportJoinFrom?.value || "").trim();
    const joiningTo = String(els.onboardingReportJoinTo?.value || "").trim();
    return onboardingProfiles.filter((row) => {
      const joiningDate = String(row.dateOfJoining || "").slice(0, 10);
      const searchable = [row.fullName, row.employeeId, row.submittedEmployeeId, row.mobile, row.personalEmail, row.email].join(" ").toLowerCase();
      return (!search || searchable.includes(search))
        && (!status || String(row.status || "pending").toLowerCase() === status)
        && (!role || employeeRoleLabel(row) === role)
        && (!selectedShifts.length || selectedShifts.includes(normaliseEmployeeShift(row.assignedShift)))
        && (!joiningFrom || (joiningDate && joiningDate >= joiningFrom))
        && (!joiningTo || (joiningDate && joiningDate <= joiningTo));
    });
  }

  function renderOnboardingReportPreview() {
    const selectedColumns = getSelectedOnboardingReportColumns();
    const rows = getFilteredOnboardingProfiles();
    const summary = rows.reduce((counts, row) => {
      const status = String(row.status || "pending").toLowerCase();
      counts[status] = (counts[status] || 0) + 1;
      return counts;
    }, { active: 0, pending: 0, inactive: 0 });
    if (els.onboardingReportCount) {
      els.onboardingReportCount.textContent = `${rows.length} selected · ${summary.active} Active · ${summary.pending} Pending · ${summary.inactive} Inactive`;
    }
    if (!els.onboardingReportPreviewHead || !els.onboardingReportPreviewBody) return;
    els.onboardingReportPreviewHead.innerHTML = "";
    els.onboardingReportPreviewBody.innerHTML = "";
    if (!selectedColumns.length) {
      els.onboardingReportPreviewBody.appendChild(createCellRow(["Select at least one report field to view the report."]));
      return;
    }
    selectedColumns.forEach((column) => {
      const th = document.createElement("th");
      th.textContent = column.label;
      els.onboardingReportPreviewHead.appendChild(th);
    });
    if (!rows.length) {
      const tr = document.createElement("tr");
      const td = document.createElement("td");
      td.colSpan = selectedColumns.length;
      td.textContent = "No employees match the selected filters.";
      tr.appendChild(td);
      els.onboardingReportPreviewBody.appendChild(tr);
      return;
    }
    rows.forEach((row) => {
      const tr = document.createElement("tr");
      selectedColumns.forEach((column) => {
        const td = document.createElement("td");
        td.textContent = column.value(row) || "-";
        tr.appendChild(td);
      });
      els.onboardingReportPreviewBody.appendChild(tr);
    });
  }

  function downloadOnboardingReport() {
    const selectedColumns = getSelectedOnboardingReportColumns();
    if (!selectedColumns.length) {
      showOnboardingNote("Select at least one report field.", "error");
      return;
    }
    const filteredRows = getFilteredOnboardingProfiles();
    if (!filteredRows.length) {
      showOnboardingNote("No employees match the selected report filters.", "error");
      return;
    }
    const csv = toCsv(filteredRows, selectedColumns.map((column) => column.label), (row) => selectedColumns.map((column) => column.value(row)));
    downloadCsv("onboarding-report.csv", csv);
    renderOnboardingReportPreview();
    showOnboardingNote("Your filtered employee report has been downloaded.", "success");
  }

  function downloadCsv(filename, text) {
    const blob = new Blob([text], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.click();
    URL.revokeObjectURL(url);
  }

  function bindExportButtons(data) {
    const appBtn = document.getElementById("export-applications");
    const empBtn = document.getElementById("export-employers");

    if (appBtn) {
      appBtn.addEventListener("click", () => {
        const csv = toCsv(
          data.applications,
          ["Name", "Mobile", "Email", "Role", "Location", "Resume", "Submitted At"],
          (row) => [
            row.fullName,
            row.mobile,
            row.email,
            row.role,
            row.currentLocation,
            row.resumeUrl || row.resumeFileName,
            formatDate(row.submittedAt)
          ]
        );
        downloadCsv("applications.csv", csv);
      });
    }

    if (empBtn) {
      empBtn.addEventListener("click", () => {
        const csv = toCsv(
          data.employerLeads,
          ["Company", "Contact Person", "Employees Needed", "Location", "Contact Number", "Submitted At"],
          (row) => [
            row.companyName,
            row.contactPerson,
            row.employeesNeeded,
            row.location,
            row.contactNumber,
            formatDate(row.submittedAt)
          ]
        );
        downloadCsv("employer-leads.csv", csv);
      });
    }
  }

  async function loadFirebaseData() {
    try {
      const [app, firestoreMod] = await Promise.all([
        getFirebaseApp(),
        import("https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js")
      ]);
      const db = firestoreMod.getFirestore(app);

      async function getCollection(name) {
        const snapshot = await firestoreMod.getDocs(firestoreMod.collection(db, name));
        return snapshot.docs.map((doc) => {
          const value = doc.data();
          let submittedAt = value.submittedAt;
          if (submittedAt && typeof submittedAt.toDate === "function") {
            submittedAt = submittedAt.toDate().toISOString();
          }
          return { ...value, submittedAt };
        });
      }

      const [applications, employerLeads, contacts, supports] = await Promise.all([
        getCollection("applications"),
        getCollection("employerLeads"),
        getCollection("contacts"),
        getCollection("supports")
      ]);

      firebaseLoaded = true;
      return { applications, employerLeads, contacts, supports };
    } catch (error) {
      console.error("Admin firebase read failed", error);
      return null;
    }
  }

  async function renderDashboard() {
    const localData = {
      applications: readLocal(keys.applications),
      employerLeads: readLocal(keys.employerLeads),
      contacts: readLocal(keys.contacts),
      supports: readLocal(keys.supports)
    };

    const firebaseData = await loadFirebaseData();
    const data = firebaseData || localData;

    setCount(els.totalApplications, data.applications.length);
    setCount(els.totalEmployers, data.employerLeads.length);
    setCount(els.totalContacts, data.contacts.length);
    setCount(els.totalSupports, data.supports.length);

    fillTable(els.applicationsBody, data.applications, (row) => [
      row.fullName,
      row.mobile,
      row.role,
      row.currentLocation,
      formatDate(row.submittedAt)
    ]);

    fillTable(els.employerBody, data.employerLeads, (row) => [
      row.companyName,
      row.contactPerson,
      row.employeesNeeded,
      row.location,
      formatDate(row.submittedAt)
    ]);

    fillTable(els.contactsBody, data.contacts, (row) => [
      row.name,
      row.email,
      row.phone,
      formatDate(row.submittedAt)
    ]);

    fillTable(els.supportBody, data.supports, (row) => [
      row.accountEmail,
      row.issueType,
      row.priority,
      formatDate(row.submittedAt)
    ]);

    bindExportButtons(data);

    const modeNode = document.getElementById("admin-source-mode");
    if (modeNode) {
      modeNode.textContent = firebaseLoaded
        ? "Live Firebase data mode"
        : "Local browser data mode";
    }

    loadOnboardingProfiles();
  }

  async function loadOnboardingProfiles() {
    if (!els.onboardingBody) return;
    showOnboardingNote("Loading onboarding profiles...");
    try {
      const [app, firestoreMod] = await Promise.all([
        getFirebaseApp(),
        import("https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js")
      ]);
      const db = firestoreMod.getFirestore(app);
      const snapshot = await firestoreMod.getDocs(firestoreMod.collection(db, "employeeProfiles"));
      onboardingProfiles = snapshot.docs.map((entry) => {
        const value = entry.data();
        const submittedAt = value.submittedAt?.toDate ? value.submittedAt.toDate().toISOString() : value.submittedAt;
        const approvedAt = value.approvedAt?.toDate ? value.approvedAt.toDate().toISOString() : value.approvedAt;
        return { id: entry.id, ...value, submittedAt, approvedAt };
      }).sort((a, b) => String(b.submittedAt || "").localeCompare(String(a.submittedAt || "")));
      const deletedCount = await deleteExpiredInactiveProfiles(db, firestoreMod);
      renderOnboardingTable();
      renderOnboardingReportPreview();
      if (deletedCount) {
        showOnboardingNote(`${deletedCount} inactive employee record${deletedCount === 1 ? "" : "s"} expired after 90 days and ${deletedCount === 1 ? "was" : "were"} deleted.`, "success");
      } else {
        showOnboardingNote(onboardingProfiles.length ? "" : "No onboarding profiles have been submitted yet.");
      }
    } catch (error) {
      console.error("Onboarding profile load failed", error);
      els.onboardingBody.innerHTML = "";
      els.onboardingBody.appendChild(createCellRow(["Onboarding data is not available yet."]));
      showOnboardingNote("Enable Cloud Firestore and publish the updated security rules to use onboarding approvals.", "error");
    }
  }

  async function deleteExpiredInactiveProfiles(db, firestoreMod) {
    if (!canManageOnboarding()) return 0;
    const now = Date.now();
    const expiredProfiles = onboardingProfiles.filter((profile) => {
      if (String(profile.status || "").toLowerCase() !== "inactive") return false;
      const expiryTime = new Date(profile.inactiveUntil || "").getTime();
      return Number.isFinite(expiryTime) && expiryTime <= now;
    });
    if (!expiredProfiles.length) return 0;

    const batch = firestoreMod.writeBatch(db);
    expiredProfiles.forEach((profile) => {
      batch.delete(firestoreMod.doc(db, "employeeProfiles", profile.id));
      batch.delete(firestoreMod.doc(db, "disabledAccounts", profile.id));
    });
    await batch.commit();
    onboardingProfiles = onboardingProfiles.filter((profile) => !expiredProfiles.some((expired) => expired.id === profile.id));
    return expiredProfiles.length;
  }

  function renderOnboardingTable() {
    if (!els.onboardingBody) return;
    els.onboardingBody.innerHTML = "";
    const visibleProfiles = onboardingProfiles.filter((row) => String(row.status || "pending").toLowerCase() === currentOnboardingFilter);
    if (!visibleProfiles.length) {
      els.onboardingBody.appendChild(createCellRow([`No ${currentOnboardingFilter} onboarding profiles yet`]));
      return;
    }
    visibleProfiles.forEach((row) => {
      const tr = document.createElement("tr");
      [row.fullName, row.mobile, row.dateOfJoining || "-"].forEach((value) => {
        const td = document.createElement("td");
        td.textContent = value || "-";
        tr.appendChild(td);
      });
      const employeeIdCell = document.createElement("td");
      if (canManageOnboarding()) {
        const employeeIdInput = document.createElement("input");
        employeeIdInput.type = "text";
        employeeIdInput.value = row.employeeId || row.submittedEmployeeId || "";
        employeeIdInput.placeholder = "Employee ID";
        employeeIdInput.dataset.employeeIdFor = row.id;
        employeeIdInput.disabled = String(row.status || "").toLowerCase() === "active";
        employeeIdCell.appendChild(employeeIdInput);
      } else {
        employeeIdCell.textContent = row.employeeId || row.submittedEmployeeId || "-";
      }
      tr.appendChild(employeeIdCell);
      const roleCell = document.createElement("td");
      if (canManageOnboarding()) {
        const roleSelect = document.createElement("select");
        roleSelect.dataset.employeeRoleFor = row.id;
        const placeholder = document.createElement("option");
        placeholder.value = "";
        placeholder.textContent = "Select role";
        roleSelect.appendChild(placeholder);
        employeeRoleOptions.forEach((role) => {
          const roleOption = document.createElement("option");
          roleOption.value = role;
          roleOption.textContent = role;
          roleOption.selected = role === normaliseEmployeeRole(row.assignedRole || row.designation);
          roleSelect.appendChild(roleOption);
        });
        roleCell.appendChild(roleSelect);
      } else {
        roleCell.textContent = employeeRoleLabel(row);
      }
      tr.appendChild(roleCell);
      const shiftCell = document.createElement("td");
      if (canManageOnboarding()) {
        const shiftSelect = document.createElement("select");
        shiftSelect.dataset.employeeShiftFor = row.id;
        const placeholder = document.createElement("option");
        placeholder.value = "";
        placeholder.textContent = "Select shift";
        shiftSelect.appendChild(placeholder);
        Object.entries(employeeShiftOptions).forEach(([employmentType, shifts]) => {
          const group = document.createElement("optgroup");
          group.label = employmentType;
          shifts.forEach((time) => {
            const shiftOption = document.createElement("option");
            shiftOption.value = `${employmentType} | ${time}`;
            shiftOption.textContent = time;
            shiftOption.selected = shiftOption.value === normaliseEmployeeShift(row.assignedShift);
            group.appendChild(shiftOption);
          });
          shiftSelect.appendChild(group);
        });
        shiftCell.appendChild(shiftSelect);
      } else {
        shiftCell.textContent = normaliseEmployeeShift(row.assignedShift) || "-";
      }
      tr.appendChild(shiftCell);
      const dateCell = document.createElement("td");
      dateCell.textContent = formatDate(row.submittedAt);
      tr.appendChild(dateCell);
      const actionCell = document.createElement("td");
      if (canManageOnboarding()) {
        const status = String(row.status || "pending").toLowerCase();
        const action = document.createElement("button");
        action.type = "button";
        if (status === "active") {
          action.className = "btn btn-outline";
          action.textContent = "Make inactive";
          action.dataset.deactivateProfile = row.id;
        } else {
          action.className = "btn btn-primary";
          action.textContent = status === "inactive" ? "Reactivate" : "Activate";
          action.dataset.activateProfile = row.id;
        }
        actionCell.appendChild(action);
      } else {
        actionCell.textContent = "Read only";
      }
      tr.appendChild(actionCell);
      els.onboardingBody.appendChild(tr);
    });
  }

  function switchOnboardingFilter(filter) {
    currentOnboardingFilter = filter;
    document.querySelectorAll("[data-onboarding-filter]").forEach((button) => {
      const isSelected = button.dataset.onboardingFilter === filter;
      button.classList.toggle("active", isSelected);
      button.setAttribute("aria-selected", String(isSelected));
    });
    renderOnboardingTable();
  }

  async function activateProfile(profileId) {
    if (!canManageOnboarding()) {
      showOnboardingNote("Your account has report-only access.", "error");
      return;
    }
    const idField = document.querySelector(`[data-employee-id-for="${profileId}"]`);
    const employeeId = String(idField?.value || "").trim();
    const roleField = document.querySelector(`[data-employee-role-for="${profileId}"]`);
    const assignedRole = normaliseEmployeeRole(roleField?.value);
    const shiftField = document.querySelector(`[data-employee-shift-for="${profileId}"]`);
    const assignedShift = normaliseEmployeeShift(shiftField?.value);
    if (!employeeId) {
      showOnboardingNote("Enter an Employee ID before activating the account.", "error");
      return;
    }
    if (!assignedRole) {
      showOnboardingNote("Select a role before activating the account.", "error");
      return;
    }
    showOnboardingNote("Activating employee account...");
    try {
      const [app, firestoreMod] = await Promise.all([
        getFirebaseApp(),
        import("https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js")
      ]);
      const db = firestoreMod.getFirestore(app);
      const batch = firestoreMod.writeBatch(db);
      batch.update(firestoreMod.doc(db, "employeeProfiles", profileId), {
        status: "active",
        employeeId,
        assignedRole,
        assignedShift,
        roleUpdatedBy: firebaseAuth?.currentUser?.email || "",
        roleUpdatedAt: firestoreMod.serverTimestamp(),
        approvedBy: firebaseAuth?.currentUser?.email || "",
        approvedAt: firestoreMod.serverTimestamp(),
        inactiveAt: firestoreMod.deleteField(),
        inactiveBy: firestoreMod.deleteField(),
        inactiveUntil: firestoreMod.deleteField(),
        inactiveExpiresAt: firestoreMod.deleteField()
      });
      batch.delete(firestoreMod.doc(db, "disabledAccounts", profileId));
      await batch.commit();
      await loadOnboardingProfiles();
      showOnboardingNote("Employee account activated.", "success");
    } catch (error) {
      console.error("Employee activation failed", error);
      showOnboardingNote("Unable to activate this account. Check the Firebase rules and try again.", "error");
    }
  }

  async function updateEmployeeShift(profileId, selectedShift) {
    if (!canManageOnboarding()) {
      showOnboardingNote("Your account has report-only access.", "error");
      return;
    }
    const assignedShift = normaliseEmployeeShift(selectedShift);
    if (!assignedShift) {
      showOnboardingNote("Select a shift before saving.", "error");
      return;
    }
    showOnboardingNote("Saving employee shift...");
    try {
      const [app, firestoreMod] = await Promise.all([
        getFirebaseApp(),
        import("https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js")
      ]);
      const db = firestoreMod.getFirestore(app);
      await firestoreMod.updateDoc(firestoreMod.doc(db, "employeeProfiles", profileId), {
        assignedShift,
        shiftUpdatedBy: firebaseAuth?.currentUser?.email || "",
        shiftUpdatedAt: firestoreMod.serverTimestamp()
      });
      await loadOnboardingProfiles();
      showOnboardingNote("Employee shift updated.", "success");
    } catch (error) {
      console.error("Employee shift update failed", error);
      showOnboardingNote("Unable to update the shift. Check the Firebase rules and try again.", "error");
    }
  }

  async function updateEmployeeRole(profileId, selectedRole) {
    if (!canManageOnboarding()) {
      showOnboardingNote("Your account has report-only access.", "error");
      return;
    }
    const assignedRole = normaliseEmployeeRole(selectedRole);
    if (!assignedRole) {
      showOnboardingNote("Select a role before saving.", "error");
      return;
    }
    showOnboardingNote("Saving employee role...");
    try {
      const [app, firestoreMod] = await Promise.all([
        getFirebaseApp(),
        import("https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js")
      ]);
      const db = firestoreMod.getFirestore(app);
      await firestoreMod.updateDoc(firestoreMod.doc(db, "employeeProfiles", profileId), {
        assignedRole,
        roleUpdatedBy: firebaseAuth?.currentUser?.email || "",
        roleUpdatedAt: firestoreMod.serverTimestamp()
      });
      await loadOnboardingProfiles();
      showOnboardingNote("Employee role updated.", "success");
    } catch (error) {
      console.error("Employee role update failed", error);
      showOnboardingNote("Unable to update the role. Check the Firebase rules and try again.", "error");
    }
  }

  async function deactivateProfile(profileId) {
    if (!canManageOnboarding()) {
      showOnboardingNote("Your account has report-only access.", "error");
      return;
    }
    showOnboardingNote("Making employee inactive...");
    try {
      const [app, firestoreMod] = await Promise.all([
        getFirebaseApp(),
        import("https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js")
      ]);
      const db = firestoreMod.getFirestore(app);
      const inactiveUntil = new Date();
      inactiveUntil.setDate(inactiveUntil.getDate() + 90);
      const batch = firestoreMod.writeBatch(db);
      batch.update(firestoreMod.doc(db, "employeeProfiles", profileId), {
        status: "inactive",
        inactiveBy: firebaseAuth?.currentUser?.email || "",
        inactiveAt: firestoreMod.serverTimestamp(),
        inactiveUntil: inactiveUntil.toISOString(),
        inactiveExpiresAt: firestoreMod.Timestamp.fromDate(inactiveUntil)
      });
      batch.set(firestoreMod.doc(db, "disabledAccounts", profileId), {
        uid: profileId,
        status: "inactive",
        disabledAt: firestoreMod.serverTimestamp(),
        inactiveExpiresAt: firestoreMod.Timestamp.fromDate(inactiveUntil)
      });
      await batch.commit();
      await loadOnboardingProfiles();
      showOnboardingNote("Employee has been moved to the Inactive list and will be deleted after 90 days.", "success");
    } catch (error) {
      console.error("Employee deactivation failed", error);
      showOnboardingNote("Unable to make this employee inactive. Check the Firebase rules and try again.", "error");
    }
  }

  function unlockDashboard() {
    if (dashboardRendered) return;
    dashboardRendered = true;
    if (els.dashboard) {
      els.dashboard.hidden = false;
    }
    if (els.accessForm) {
      els.accessForm.closest(".form-shell").hidden = true;
    }
    setupGoogleAdminLinks();
    renderDashboard();
  }

  if (!els.accessForm) {
    return;
  }

  // Reuse the secure Firebase session created from the main Login page.
  // Only allow the explicitly configured admin accounts into this dashboard.
  getFirebaseAuth().then(async (auth) => {
    const authMod = await import("https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js");
    authMod.onAuthStateChanged(auth, (user) => {
      if (!isAuthorizedAdmin(user)) return;
      currentAdminRole = getAdminRole(user);
      showAccessNote("");
      unlockDashboard();
    });
  }).catch((error) => {
    console.error("Admin session check failed", error);
  });

  els.accessForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const email = (els.emailInput?.value || "").trim();
    const password = els.pinInput?.value || "";
    showAccessNote("Signing in...");

    try {
      const [auth, authMod] = await Promise.all([
        getFirebaseAuth(),
        import("https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js")
      ]);
      await authMod.signInWithEmailAndPassword(auth, email, password);
      if (!isAuthorizedAdmin(auth.currentUser)) {
        await authMod.signOut(auth);
        throw { code: "auth/not-authorised-admin" };
      }
      currentAdminRole = getAdminRole(auth.currentUser);
      showAccessNote("");
      unlockDashboard();
    } catch (error) {
      console.error("Admin sign-in failed", error);
      showAccessNote(getLoginErrorMessage(error), "error");
    }
  });

  if (els.resetPassword) {
    els.resetPassword.addEventListener("click", async () => {
      const email = (els.emailInput?.value || "").trim();
      if (!email) {
        showAccessNote("Enter your email address first, then select Forgot Password.", "error");
        return;
      }
      try {
        const [auth, authMod] = await Promise.all([
          getFirebaseAuth(),
          import("https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js")
        ]);
        await authMod.sendPasswordResetEmail(auth, email);
        showAccessNote("A password reset email has been sent if this address has an account.", "success");
      } catch (error) {
        console.error("Password reset failed", error);
        showAccessNote("Unable to send a password reset email. Please try again.", "error");
      }
    });
  }

  if (els.logout) {
    els.logout.addEventListener("click", async () => {
      try {
        const [auth, authMod] = await Promise.all([
          getFirebaseAuth(),
          import("https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js")
        ]);
        await authMod.signOut(auth);
      } finally {
        window.location.reload();
      }
    });
  }

  if (els.refreshOnboarding) {
    els.refreshOnboarding.addEventListener("click", loadOnboardingProfiles);
  }

  renderOnboardingReportFields();
  setupOnboardingReportFilters();
  renderOnboardingReportPreview();

  if (els.showOnboardingReport) {
    els.showOnboardingReport.addEventListener("click", renderOnboardingReportPreview);
  }

  [
    els.onboardingReportSearch,
    els.onboardingReportStatus,
    els.onboardingReportRole,
    els.onboardingReportShift,
    els.onboardingReportJoinFrom,
    els.onboardingReportJoinTo,
    els.onboardingReportFields
  ].filter(Boolean).forEach((control) => {
    control.addEventListener("input", renderOnboardingReportPreview);
    control.addEventListener("change", renderOnboardingReportPreview);
  });

  if (els.downloadOnboardingReport) {
    els.downloadOnboardingReport.addEventListener("click", downloadOnboardingReport);
  }

  if (els.onboardingBody) {
    els.onboardingBody.addEventListener("click", (event) => {
      const button = event.target.closest("[data-activate-profile]");
      if (button) activateProfile(button.dataset.activateProfile);
      const deactivateButton = event.target.closest("[data-deactivate-profile]");
      if (deactivateButton) deactivateProfile(deactivateButton.dataset.deactivateProfile);
    });
    els.onboardingBody.addEventListener("change", (event) => {
      const employeeRole = event.target.closest("[data-employee-role-for]");
      if (employeeRole) updateEmployeeRole(employeeRole.dataset.employeeRoleFor, employeeRole.value);
      const employeeShift = event.target.closest("[data-employee-shift-for]");
      if (employeeShift) updateEmployeeShift(employeeShift.dataset.employeeShiftFor, employeeShift.value);
    });
  }

  document.querySelectorAll("[data-onboarding-filter]").forEach((button) => {
    button.addEventListener("click", () => switchOnboardingFilter(button.dataset.onboardingFilter));
  });
})();
