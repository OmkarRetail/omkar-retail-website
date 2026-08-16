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
    return firebaseAuth;
  }

  function showAccessNote(message, type) {
    if (!els.accessNote) return;
    els.accessNote.textContent = message;
    els.accessNote.className = `note ${type || ""}`.trim();
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
      showAccessNote("");
      unlockDashboard();
    } catch (error) {
      console.error("Admin sign-in failed", error);
      showAccessNote("Incorrect email address or password. Try again.", "error");
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
})();
