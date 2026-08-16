(function () {
  const config = window.OMKAR_CONFIG || {};
  const firebaseVersion = "10.12.2";
  let firebaseApp;
  let firebaseAuth;
  let firebaseServices;
  let signedInUser = null;
  let profile = null;

  const els = {
    authPanel: document.getElementById("auth-panel"),
    onboardingPanel: document.getElementById("onboarding-panel"),
    showSignup: document.getElementById("show-signup"),
    showSignin: document.getElementById("show-signin"),
    signupForm: document.getElementById("signup-form"),
    signinForm: document.getElementById("signin-form"),
    authNote: document.getElementById("auth-note"),
    resetPassword: document.getElementById("reset-password"),
    employeeLogout: document.getElementById("employee-logout"),
    onboardingForm: document.getElementById("onboarding-form"),
    onboardingNote: document.getElementById("onboarding-note"),
    approvalStatus: document.getElementById("approval-status")
  };

  function showNote(element, message, type) {
    if (!element) return;
    element.textContent = message || "";
    element.className = `note ${type || ""}`.trim();
  }

  function getFirebaseConfig() {
    const fb = config.firebase || {};
    const required = ["apiKey", "authDomain", "projectId", "storageBucket", "appId"];
    if (!required.every((key) => Boolean(fb[key]))) {
      throw new Error("Firebase is not configured for this website yet.");
    }
    return fb;
  }

  async function getServices() {
    if (firebaseServices) return firebaseServices;
    const fb = getFirebaseConfig();
    const [appMod, authMod, firestoreMod, storageMod] = await Promise.all([
      import(`https://www.gstatic.com/firebasejs/${firebaseVersion}/firebase-app.js`),
      import(`https://www.gstatic.com/firebasejs/${firebaseVersion}/firebase-auth.js`),
      import(`https://www.gstatic.com/firebasejs/${firebaseVersion}/firebase-firestore.js`),
      import(`https://www.gstatic.com/firebasejs/${firebaseVersion}/firebase-storage.js`)
    ]);
    firebaseApp = appMod.getApps()[0] || appMod.initializeApp(fb);
    firebaseAuth = authMod.getAuth(firebaseApp);
    firebaseServices = {
      auth: firebaseAuth,
      authMod,
      db: firestoreMod.getFirestore(firebaseApp),
      fs: firestoreMod,
      storage: storageMod.getStorage(firebaseApp),
      st: storageMod
    };
    return firebaseServices;
  }

  function switchAuth(mode) {
    const signup = mode === "signup";
    els.signupForm.hidden = !signup;
    els.signinForm.hidden = signup;
    els.showSignup.classList.toggle("active", signup);
    els.showSignin.classList.toggle("active", !signup);
    els.showSignup.setAttribute("aria-selected", String(signup));
    els.showSignin.setAttribute("aria-selected", String(!signup));
    showNote(els.authNote, "");
  }

  function cleanValue(value) {
    return String(value || "").trim();
  }

  function getFormValue(name) {
    return cleanValue(els.onboardingForm.elements.namedItem(name)?.value);
  }

  function setFormValue(name, value) {
    const field = els.onboardingForm.elements.namedItem(name);
    if (field) field.value = value || "";
  }

  function safeFileName(name) {
    return String(name || "document").replace(/[^a-zA-Z0-9._-]/g, "-");
  }

  function validateFile(file, type) {
    if (!file) return "";
    if (file.size > 5 * 1024 * 1024) return "Each document must be 5 MB or smaller.";
    const allowed = type === "photo"
      ? ["image/jpeg", "image/png"]
      : ["application/pdf", "image/jpeg", "image/png"];
    if (!allowed.includes(file.type)) return "Use only the file types shown beside each upload field.";
    return "";
  }

  async function uploadDocument(services, userId, key, file) {
    if (!file) return profile?.documents?.[key] || null;
    const validation = validateFile(file, key);
    if (validation) throw new Error(validation);
    const path = `onboarding/${userId}/${key}-${Date.now()}-${safeFileName(file.name)}`;
    const fileRef = services.st.ref(services.storage, path);
    await services.st.uploadBytes(fileRef, file, { contentType: file.type });
    return { path, name: file.name, type: file.type, updatedAt: new Date().toISOString() };
  }

  function renderStatus() {
    if (!els.approvalStatus) return;
    const status = String(profile?.status || "pending").toLowerCase();
    const employeeId = cleanValue(profile?.employeeId);
    els.approvalStatus.className = `approval-status ${status}`;
    if (status === "active") {
      els.approvalStatus.textContent = employeeId
        ? `Your profile is active. Your employee ID is ${employeeId}.`
        : "Your profile is active.";
    } else {
      els.approvalStatus.textContent = "Your details are pending HR review. You can update them until approval is completed.";
    }
  }

  function populateProfile() {
    if (!signedInUser || !els.onboardingForm) return;
    setFormValue("fullName", profile?.fullName);
    setFormValue("mobile", profile?.mobile);
    setFormValue("email", signedInUser.email || profile?.email);
    setFormValue("dateOfBirth", profile?.dateOfBirth);
    setFormValue("maritalStatus", profile?.maritalStatus);
    setFormValue("emergencyContactName", profile?.emergencyContactName);
    setFormValue("emergencyContactMobile", profile?.emergencyContactMobile);
    setFormValue("pan", profile?.pan);
    setFormValue("uan", profile?.uan);
    setFormValue("nomineeName", profile?.nomineeName);
    setFormValue("nomineeRelationship", profile?.nomineeRelationship);
    setFormValue("bankName", profile?.bankName);
    setFormValue("bankAccountNumber", profile?.bankAccountNumber);
    setFormValue("ifsc", profile?.ifsc);
    setFormValue("currentAddress", profile?.currentAddress);
    renderStatus();
  }

  async function loadProfile() {
    const services = await getServices();
    const ref = services.fs.doc(services.db, "employeeProfiles", signedInUser.uid);
    const snapshot = await services.fs.getDoc(ref);
    profile = snapshot.exists() ? snapshot.data() : null;
    populateProfile();
  }

  async function showEmployeePortal(user) {
    signedInUser = user;
    els.authPanel.hidden = true;
    els.onboardingPanel.hidden = false;
    showNote(els.onboardingNote, "Loading your profile...");
    try {
      await loadProfile();
      showNote(els.onboardingNote, "");
    } catch (error) {
      console.error("Employee profile load failed", error);
      showNote(els.onboardingNote, "Your account was created, but onboarding storage is not ready yet. Please contact HR.", "error");
    }
  }

  function getAuthError(error) {
    switch (error?.code) {
      case "auth/email-already-in-use": return "This email already has an account. Please sign in instead.";
      case "auth/invalid-email": return "Enter a valid email address.";
      case "auth/weak-password": return "Use a password with at least 8 characters.";
      case "auth/invalid-credential": return "Incorrect email address or password.";
      case "auth/too-many-requests": return "Too many attempts. Please wait and try again later.";
      default: return "Unable to continue. Please try again or contact HR.";
    }
  }

  els.showSignup?.addEventListener("click", () => switchAuth("signup"));
  els.showSignin?.addEventListener("click", () => switchAuth("signin"));

  els.signupForm?.addEventListener("submit", async (event) => {
    event.preventDefault();
    const email = cleanValue(document.getElementById("signup-email").value);
    const password = document.getElementById("signup-password").value;
    const confirmPassword = document.getElementById("signup-password-confirm").value;
    if (password !== confirmPassword) {
      showNote(els.authNote, "The passwords do not match.", "error");
      return;
    }
    showNote(els.authNote, "Creating your account...");
    try {
      const services = await getServices();
      const credential = await services.authMod.createUserWithEmailAndPassword(services.auth, email, password);
      try { await services.authMod.sendEmailVerification(credential.user); } catch (_) { /* Email verification is optional. */ }
      showNote(els.authNote, "");
      await showEmployeePortal(credential.user);
    } catch (error) {
      console.error("Employee sign-up failed", error);
      showNote(els.authNote, getAuthError(error), "error");
    }
  });

  els.signinForm?.addEventListener("submit", async (event) => {
    event.preventDefault();
    showNote(els.authNote, "Signing in...");
    try {
      const services = await getServices();
      const credential = await services.authMod.signInWithEmailAndPassword(
        services.auth,
        cleanValue(document.getElementById("signin-email").value),
        document.getElementById("signin-password").value
      );
      showNote(els.authNote, "");
      await showEmployeePortal(credential.user);
    } catch (error) {
      console.error("Employee sign-in failed", error);
      showNote(els.authNote, getAuthError(error), "error");
    }
  });

  els.resetPassword?.addEventListener("click", async () => {
    const email = cleanValue(document.getElementById("signin-email").value);
    if (!email) {
      showNote(els.authNote, "Enter your email address first, then select Forgot password.", "error");
      return;
    }
    try {
      const services = await getServices();
      await services.authMod.sendPasswordResetEmail(services.auth, email);
      showNote(els.authNote, "If this email has an account, a password reset link has been sent.", "success");
    } catch (error) {
      showNote(els.authNote, "Unable to send the password reset email. Please try again.", "error");
    }
  });

  els.onboardingForm?.addEventListener("submit", async (event) => {
    event.preventDefault();
    if (!signedInUser) return;
    const mobile = getFormValue("mobile");
    const emergencyMobile = getFormValue("emergencyContactMobile");
    if (!/^\d{10}$/.test(mobile) || !/^\d{10}$/.test(emergencyMobile)) {
      showNote(els.onboardingNote, "Enter valid 10-digit mobile numbers.", "error");
      return;
    }
    const files = {
      aadhaar: els.onboardingForm.elements.namedItem("aadhaarFile").files[0],
      pan: els.onboardingForm.elements.namedItem("panFile").files[0],
      photo: els.onboardingForm.elements.namedItem("photoFile").files[0]
    };
    if (!profile?.documents?.aadhaar && !files.aadhaar || !profile?.documents?.pan && !files.pan || !profile?.documents?.photo && !files.photo) {
      showNote(els.onboardingNote, "Please upload Aadhaar, PAN, and a recent photo.", "error");
      return;
    }
    showNote(els.onboardingNote, "Saving your details securely...");
    try {
      const services = await getServices();
      const documents = {
        aadhaar: await uploadDocument(services, signedInUser.uid, "aadhaar", files.aadhaar),
        pan: await uploadDocument(services, signedInUser.uid, "pan", files.pan),
        photo: await uploadDocument(services, signedInUser.uid, "photo", files.photo)
      };
      const payload = {
        ownerUid: signedInUser.uid,
        email: signedInUser.email || "",
        fullName: getFormValue("fullName"),
        mobile,
        dateOfBirth: getFormValue("dateOfBirth"),
        maritalStatus: getFormValue("maritalStatus"),
        emergencyContactName: getFormValue("emergencyContactName"),
        emergencyContactMobile: emergencyMobile,
        pan: getFormValue("pan").toUpperCase(),
        uan: getFormValue("uan").toUpperCase(),
        nomineeName: getFormValue("nomineeName"),
        nomineeRelationship: getFormValue("nomineeRelationship"),
        bankName: getFormValue("bankName"),
        bankAccountNumber: getFormValue("bankAccountNumber"),
        ifsc: getFormValue("ifsc").toUpperCase(),
        currentAddress: getFormValue("currentAddress"),
        documents,
        lastUpdated: services.fs.serverTimestamp()
      };
      if (!profile) {
        payload.status = "pending";
        payload.employeeId = "";
        payload.approvedBy = "";
        payload.approvedAt = null;
        payload.submittedAt = services.fs.serverTimestamp();
      }
      await services.fs.setDoc(services.fs.doc(services.db, "employeeProfiles", signedInUser.uid), payload, { merge: true });
      await loadProfile();
      showNote(els.onboardingNote, "Your onboarding details have been saved and sent for HR review.", "success");
    } catch (error) {
      console.error("Onboarding save failed", error);
      showNote(els.onboardingNote, "Unable to save your onboarding details. Please try again or contact HR.", "error");
    }
  });

  els.employeeLogout?.addEventListener("click", async () => {
    const services = await getServices();
    await services.authMod.signOut(services.auth);
  });

  getServices().then((services) => {
    services.authMod.onAuthStateChanged(services.auth, (user) => {
      if (user) {
        showEmployeePortal(user);
      } else {
        signedInUser = null;
        profile = null;
        els.authPanel.hidden = false;
        els.onboardingPanel.hidden = true;
      }
    });
  }).catch((error) => {
    console.error("Firebase setup failed", error);
    showNote(els.authNote, "Employee onboarding is being set up. Please contact HR for assistance.", "error");
  });
})();
