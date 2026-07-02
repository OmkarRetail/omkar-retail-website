(function () {
  const config = window.OMKAR_CONFIG || {};
  const storageKeys = {
    applications: "omkar_applications",
    employerLeads: "omkar_employer_leads",
    contacts: "omkar_contacts",
    supports: "omkar_supports"
  };

  let firebaseClients = null;

  function showNote(noteEl, text, type) {
    if (!noteEl) return;
    noteEl.textContent = text;
    noteEl.className = `note ${type}`;
  }

  function readLocal(key) {
    try {
      return JSON.parse(localStorage.getItem(key) || "[]");
    } catch {
      return [];
    }
  }

  function writeLocal(key, payload) {
    const current = readLocal(key);
    current.unshift(payload);
    localStorage.setItem(key, JSON.stringify(current));
  }

  function fieldValue(form, fieldName) {
    const field = form.elements.namedItem(fieldName);
    return field && "value" in field ? String(field.value).trim() : "";
  }

  function isSpam(form) {
    const honey = form.querySelector("[name='website']");
    if (honey && honey.value.trim()) {
      return true;
    }

    const startedAt = Number(form.dataset.startedAt || "0");
    const elapsed = Date.now() - startedAt;
    if (startedAt && elapsed < 800) {
      return true;
    }

    return false;
  }

  async function loadFirebaseClients() {
    if (firebaseClients) {
      return firebaseClients;
    }

    const fb = config.firebase || {};
    const required = ["apiKey", "authDomain", "projectId", "storageBucket", "appId"];
    const hasFirebase = required.every((key) => Boolean(fb[key]));

    if (!hasFirebase) {
      return null;
    }

    try {
      const [{ initializeApp }, firestoreMod, storageMod] = await Promise.all([
        import("https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js"),
        import("https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js"),
        import("https://www.gstatic.com/firebasejs/10.12.2/firebase-storage.js")
      ]);

      const app = initializeApp(fb);
      const db = firestoreMod.getFirestore(app);
      const storage = storageMod.getStorage(app);

      firebaseClients = {
        db,
        storage,
        fs: firestoreMod,
        st: storageMod
      };
      return firebaseClients;
    } catch (error) {
      console.error("Firebase init failed", error);
      return null;
    }
  }

  async function notifyAdmin(subject, payload) {
    if (!config.web3formsAccessKey) {
      return;
    }

    try {
      const formData = new FormData();
      formData.append("access_key", config.web3formsAccessKey);
      formData.append("subject", subject);
      formData.append("from_name", config.companyName || "OMKAR RETAIL VENTURES");
      formData.append("email", config.email || "omkarretailventure@gmail.com");
      formData.append("message", JSON.stringify(payload, null, 2));

      await fetch("https://api.web3forms.com/submit", {
        method: "POST",
        body: formData
      });
    } catch (error) {
      console.error("Email notification failed", error);
    }
  }

  async function uploadResumeIfAvailable(file) {
    if (!file) {
      return { fileName: "", fileUrl: "" };
    }

    const firebase = await loadFirebaseClients();
    if (!firebase) {
      return { fileName: file.name, fileUrl: "" };
    }

    const safeName = `${Date.now()}-${file.name.replace(/\s+/g, "-")}`;
    const path = `resumes/${safeName}`;
    const resumeRef = firebase.st.ref(firebase.storage, path);
    await firebase.st.uploadBytes(resumeRef, file);
    const fileUrl = await firebase.st.getDownloadURL(resumeRef);
    return { fileName: file.name, fileUrl };
  }

  function readResumeForGoogle(file) {
    if (!file) {
      return Promise.resolve({ fileName: "", mimeType: "", base64: "" });
    }

    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const result = String(reader.result || "");
        const commaIndex = result.indexOf(",");
        resolve({
          fileName: file.name,
          mimeType: file.type || "application/octet-stream",
          base64: commaIndex >= 0 ? result.slice(commaIndex + 1) : result
        });
      };
      reader.onerror = () => reject(new Error("Could not read resume file."));
      reader.readAsDataURL(file);
    });
  }

  async function submitCandidateApplicationToGoogle(payload, resumeFile) {
    const webAppUrl = String(config.googleAppsScriptWebAppUrl || "").trim();
    if (!webAppUrl) {
      return false;
    }

    const body = new URLSearchParams();
    body.append(
      "payload",
      JSON.stringify({
        formType: "candidateApplication",
        submitToken: config.googleSubmitToken || "",
        ...payload,
        resume: await readResumeForGoogle(resumeFile)
      })
    );

    await fetch(webAppUrl, {
      method: "POST",
      mode: "no-cors",
      body
    });

    return true;
  }

  async function saveRecord(collectionName, localKey, payload) {
    const firebase = await loadFirebaseClients();

    if (firebase) {
      await firebase.fs.addDoc(firebase.fs.collection(firebase.db, collectionName), {
        ...payload,
        createdAt: firebase.fs.serverTimestamp()
      });
    }

    writeLocal(localKey, payload);
  }

  function setFormStartTimes() {
    document.querySelectorAll("form[data-track-form]").forEach((form) => {
      form.dataset.startedAt = String(Date.now());
    });
  }

  function prefillApplyFormFromQuery() {
    const applyForm = document.getElementById("candidate-form");
    if (!applyForm) {
      return;
    }

    const params = new URLSearchParams(window.location.search);
    const role = params.get("role") || "";
    const location = params.get("location") || "";

    const roleField = document.getElementById("candidate-role");
    const locationField = document.getElementById("candidate-location");

    if (roleField && role) {
      roleField.value = role;
    }

    if (locationField && location) {
      locationField.value = location;
    }
  }

  function bindCandidateForm() {
    const form = document.getElementById("candidate-form");
    const note = document.getElementById("candidate-note");
    if (!form) {
      return;
    }

    form.addEventListener("submit", async (event) => {
      event.preventDefault();

      if (isSpam(form)) {
        showNote(note, "Submission blocked by anti-spam protection.", "error");
        return;
      }

      const resumeField = form.elements.namedItem("resume");
      const resumeFile = resumeField && resumeField.files ? resumeField.files[0] : null;

      try {
        const resume = await uploadResumeIfAvailable(resumeFile);
        const payload = {
          fullName: fieldValue(form, "fullName"),
          mobile: fieldValue(form, "mobile"),
          email: fieldValue(form, "email"),
          age: fieldValue(form, "age"),
          gender: fieldValue(form, "gender"),
          education: fieldValue(form, "education"),
          experience: fieldValue(form, "experience"),
          role: fieldValue(form, "role"),
          currentLocation: fieldValue(form, "currentLocation"),
          aadhaar: fieldValue(form, "aadhaar"),
          pan: fieldValue(form, "pan"),
          joinAvailability: fieldValue(form, "joinAvailability"),
          resumeFileName: resume.fileName,
          resumeUrl: resume.fileUrl,
          source: "website",
          submittedAt: new Date().toISOString()
        };

        try {
          await submitCandidateApplicationToGoogle(payload, resumeFile);
        } catch (googleError) {
          console.error("Google Apps Script submit failed", googleError);
        }

        await saveRecord("applications", storageKeys.applications, payload);
        await notifyAdmin("New Candidate Application", payload);

        form.reset();
        showNote(note, "Application submitted successfully. Our team will contact you soon.", "success");
      } catch (error) {
        console.error(error);
        showNote(note, "Could not submit now. Please try again.", "error");
      }
    });
  }

  function bindEmployerForm() {
    const form = document.getElementById("employer-form");
    const note = document.getElementById("employer-note");
    if (!form) {
      return;
    }

    form.addEventListener("submit", async (event) => {
      event.preventDefault();

      if (isSpam(form)) {
        showNote(note, "Submission blocked by anti-spam protection.", "error");
        return;
      }

      const payload = {
        companyName: fieldValue(form, "companyName"),
        contactPerson: fieldValue(form, "contactPerson"),
        requirementDetails: fieldValue(form, "requirementDetails"),
        employeesNeeded: fieldValue(form, "employeesNeeded"),
        location: fieldValue(form, "location"),
        contactNumber: fieldValue(form, "contactNumber"),
        email: fieldValue(form, "email"),
        submittedAt: new Date().toISOString()
      };

      try {
        await saveRecord("employerLeads", storageKeys.employerLeads, payload);
        await notifyAdmin("New Zepto Hiring", payload);
        form.reset();
        showNote(note, "Enquiry submitted. Our hiring team will contact your company shortly.", "success");
      } catch (error) {
        console.error(error);
        showNote(note, "Could not submit now. Please try again.", "error");
      }
    });
  }

  function bindContactForm() {
    const form = document.getElementById("contact-form");
    const note = document.getElementById("contact-note");
    if (!form) {
      return;
    }

    form.addEventListener("submit", async (event) => {
      event.preventDefault();

      if (isSpam(form)) {
        showNote(note, "Submission blocked by anti-spam protection.", "error");
        return;
      }

      const payload = {
        name: fieldValue(form, "name"),
        email: fieldValue(form, "email"),
        phone: fieldValue(form, "phone"),
        message: fieldValue(form, "message"),
        submittedAt: new Date().toISOString()
      };

      try {
        await saveRecord("contacts", storageKeys.contacts, payload);
        await notifyAdmin("New Contact Message", payload);
        form.reset();
        showNote(note, "Message sent successfully. We will get back to you shortly.", "success");
      } catch (error) {
        console.error(error);
        showNote(note, "Could not submit now. Please try again.", "error");
      }
    });
  }

  function bindSupportForm() {
    const form = document.getElementById("support-form");
    const note = document.getElementById("support-note");
    if (!form) {
      return;
    }

    form.addEventListener("submit", async (event) => {
      event.preventDefault();

      if (isSpam(form)) {
        showNote(note, "Submission blocked by anti-spam protection.", "error");
        return;
      }

      const payload = {
        accountEmail: fieldValue(form, "accountEmail"),
        issueType: fieldValue(form, "issueType"),
        priority: fieldValue(form, "priority"),
        details: fieldValue(form, "details"),
        submittedAt: new Date().toISOString()
      };

      try {
        await saveRecord("supports", storageKeys.supports, payload);
        await notifyAdmin("New Support Ticket", payload);
        form.reset();
        showNote(note, "Support request received. Ticket ID will be shared on email.", "success");
      } catch (error) {
        console.error(error);
        showNote(note, "Could not submit now. Please try again.", "error");
      }
    });
  }

  setFormStartTimes();
  prefillApplyFormFromQuery();
  bindCandidateForm();
  bindEmployerForm();
  bindContactForm();
  bindSupportForm();
})();

