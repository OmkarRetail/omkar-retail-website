(function () {
  const config = window.OMKAR_CONFIG || {};
  const firebaseVersion = "10.12.2";
  let firebaseApp;
  let firebaseAuth;
  let firebaseServices;
  let signedInUser = null;
  let profile = null;
  let signupInProgress = false;
  let signupRequiresLogin = false;
  let inactiveAccountBlocked = false;

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
    approvalStatus: document.getElementById("approval-status"),
    onboardingSaveAction: document.getElementById("onboarding-save-action")
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

  function isAdminUser(user) {
    const email = cleanValue(user?.email).toLowerCase();
    return Boolean(config.adminRoles?.[email]);
  }

  function openCorrectPortal(user) {
    if (signupRequiresLogin) return;
    if (isAdminUser(user)) {
      window.location.replace("admin.html");
      return;
    }
    showEmployeePortal(user);
  }

  async function getServices() {
    if (firebaseServices) return firebaseServices;
    const fb = getFirebaseConfig();
    const [appMod, authMod, firestoreMod] = await Promise.all([
      import(`https://www.gstatic.com/firebasejs/${firebaseVersion}/firebase-app.js`),
      import(`https://www.gstatic.com/firebasejs/${firebaseVersion}/firebase-auth.js`),
      import(`https://www.gstatic.com/firebasejs/${firebaseVersion}/firebase-firestore.js`)
    ]);
    firebaseApp = appMod.getApps()[0] || appMod.initializeApp(fb);
    firebaseAuth = authMod.getAuth(firebaseApp);
    firebaseServices = {
      auth: firebaseAuth,
      authMod,
      db: firestoreMod.getFirestore(firebaseApp),
      fs: firestoreMod
    };
    return firebaseServices;
  }

  function switchAuth(mode) {
    const signup = mode === "signup";
    els.signupForm.hidden = !signup;
    els.signinForm.hidden = signup;
    els.signupForm.style.display = signup ? "grid" : "none";
    els.signinForm.style.display = signup ? "none" : "grid";
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

  function renderStatus() {
    if (!els.approvalStatus) return;
    const status = String(profile?.status || "pending").toLowerCase();
    const employeeId = cleanValue(profile?.employeeId);
    els.approvalStatus.className = `approval-status ${status}`;
    if (status === "active") {
      els.approvalStatus.textContent = employeeId
        ? `Your onboarding is complete. Your employee ID is ${employeeId}. Your saved details are shown below.`
        : "Your onboarding is complete. Your saved details are shown below.";
    } else if (status === "inactive") {
      els.approvalStatus.textContent = "Your employee profile is inactive. Your saved details are shown below. Please contact HR for assistance.";
    } else {
      els.approvalStatus.textContent = "Your details are pending HR review. You can update them until approval is completed.";
    }
  }

  function setOnboardingEditable() {
    if (!els.onboardingForm) return;
    const status = String(profile?.status || "").toLowerCase();
    const isComplete = status === "active" || status === "inactive";
    Array.from(els.onboardingForm.elements).forEach((field) => {
      field.disabled = isComplete;
    });
    if (els.onboardingSaveAction) {
      els.onboardingSaveAction.hidden = isComplete;
      els.onboardingSaveAction.style.display = isComplete ? "none" : "";
    }
  }

  function populateProfile() {
    if (!signedInUser || !els.onboardingForm) return;
    setFormValue("fullName", profile?.fullName);
    setFormValue("mobile", profile?.mobile);
    setFormValue("submittedEmployeeId", profile?.submittedEmployeeId || profile?.employeeId);
    setFormValue("personalEmail", profile?.personalEmail || signedInUser.email || profile?.email);
    setFormValue("dateOfBirth", profile?.dateOfBirth);
    setFormValue("dateOfJoining", profile?.dateOfJoining);
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
    setOnboardingEditable();
  }

  async function loadProfile() {
    const services = await getServices();
    const disabledRef = services.fs.doc(services.db, "disabledAccounts", signedInUser.uid);
    const disabledSnapshot = await services.fs.getDoc(disabledRef);
    if (disabledSnapshot.exists()) {
      throw { code: "account-inactive" };
    }
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
      // Only a real disabled-account record can make an employee inactive.
      // A permissions or connection issue must never sign out a new joiner.
      if (error?.code === "account-inactive") {
        if (inactiveAccountBlocked) return;
        inactiveAccountBlocked = true;
        try {
          const services = await getServices();
          await services.authMod.signOut(services.auth);
        } finally {
          signedInUser = null;
          profile = null;
          els.onboardingPanel.hidden = true;
          els.authPanel.hidden = false;
          switchAuth("signin");
          showNote(els.authNote, "This employee account is inactive. Please contact HR for assistance.", "error");
          inactiveAccountBlocked = false;
        }
        return;
      }
      showNote(els.onboardingNote, "We could not load your saved details just now. You can complete the form below and submit it for HR review.", "error");
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
    if (signupInProgress) return;
    const email = cleanValue(document.getElementById("signup-email").value);
    const password = document.getElementById("signup-password").value.trim();
    const confirmPassword = document.getElementById("signup-password-confirm").value.trim();
    if (password !== confirmPassword) {
      showNote(els.authNote, "The passwords do not match. Please enter the same password in both fields.", "error");
      document.getElementById("signup-password-confirm").focus();
      return;
    }
    signupInProgress = true;
    signupRequiresLogin = true;
    const submitButton = els.signupForm.querySelector('button[type="submit"]');
    if (submitButton) submitButton.disabled = true;
    showNote(els.authNote, "Creating your account...");
    try {
      const services = await getServices();
      const credential = await services.authMod.createUserWithEmailAndPassword(services.auth, email, password);
      try { await services.authMod.sendEmailVerification(credential.user); } catch (_) { /* Email verification is optional. */ }
      await services.authMod.signOut(services.auth);
      signupRequiresLogin = false;
      switchAuth("signin");
      document.getElementById("signin-email").value = email;
      document.getElementById("signin-password").value = "";
      showNote(els.authNote, "Account created successfully. Please sign in to complete your onboarding details.", "success");
    } catch (error) {
      console.error("Employee sign-up failed", error);
      showNote(els.authNote, getAuthError(error), "error");
    } finally {
      signupInProgress = false;
      signupRequiresLogin = false;
      if (submitButton) submitButton.disabled = false;
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
      await openCorrectPortal(credential.user);
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
      await services.authMod.sendPasswordResetEmail(services.auth, email, {
        url: "https://omkarretailventures.in/onboarding.html",
        handleCodeInApp: false
      });
      showNote(els.authNote, "Password reset link sent. Please check your inbox and spam folder.", "success");
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
    showNote(els.onboardingNote, "Saving your details securely...");
    try {
      const services = await getServices();
      const payload = {
        ownerUid: signedInUser.uid,
        email: signedInUser.email || "",
        submittedEmployeeId: getFormValue("submittedEmployeeId"),
        personalEmail: getFormValue("personalEmail"),
        fullName: getFormValue("fullName"),
        mobile,
        dateOfBirth: getFormValue("dateOfBirth"),
        dateOfJoining: getFormValue("dateOfJoining"),
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
      showNote(els.onboardingNote, "Your onboarding details have been submitted successfully for HR review.", "success");
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
        openCorrectPortal(user);
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
