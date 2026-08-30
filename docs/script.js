(function () {
  const defaults = {
    companyName: "OMKAR RETAIL VENTURES",
    domain: "omkarretailventures.in",
    industry: "Zepto Dark Store Hiring",
    location: "Zepto Store, 1256, Gnanabharathi, Stage II, Kengeri Satellite Town, Bengaluru, Karnataka 560059",
    phone: "+91 9986362446",
    whatsappNumber: "919986362446",
    email: "omkarretailventure@gmail.com",
    officeAddress: "OMKAR RETAIL VENTURES - Zepto Store, 1256, Gnanabharathi, Stage II, Kengeri Satellite Town, Bengaluru, Karnataka 560059",
    social: {
      instagram: "https://instagram.com/",
      facebook: "https://facebook.com/",
      linkedin: "https://linkedin.com/"
    },
    analyticsMeasurementId: "",
    firebase: {},
    googleAppsScriptWebAppUrl: "",
    googleSheetId: "",
    googleDriveResumeFolderId: "",
    googleSubmitToken: "",
    web3formsAccessKey: ""
  };

  const userConfig = window.OMKAR_SITE_CONFIG || {};
  const config = {
    ...defaults,
    ...userConfig,
    social: {
      ...defaults.social,
      ...(userConfig.social || {})
    }
  };
  window.OMKAR_CONFIG = config;

  const setText = (selector, value) => {
    document.querySelectorAll(selector).forEach((node) => {
      node.textContent = value;
    });
  };

  const setHref = (selector, href) => {
    document.querySelectorAll(selector).forEach((node) => {
      node.setAttribute("href", href);
    });
  };

  setText("[data-company-name]", config.companyName);
  setText("[data-company-domain]", config.domain);
  setText("[data-company-location]", config.location);
  setText("[data-company-phone]", config.phone);
  setText("[data-company-email]", config.email);
  setText(".current-year", String(new Date().getFullYear()));
  const normalizedPhone = config.phone.replace(/[^\d+]/g, "");
  setHref("[data-company-phone]", `tel:${normalizedPhone}`);
  setHref("[data-company-email]", `mailto:${config.email}`);

  const whatsappLink = `https://wa.me/${config.whatsappNumber}`;
  setHref("[data-whatsapp-link]", whatsappLink);

  const social = config.social || {};
  if (social.instagram) setHref("[data-social-instagram]", social.instagram);
  if (social.facebook) setHref("[data-social-facebook]", social.facebook);
  if (social.linkedin) setHref("[data-social-linkedin]", social.linkedin);

  const mobileToggle = document.querySelector("[data-nav-toggle]");
  const nav = document.querySelector("[data-site-nav]");
  if (mobileToggle && nav) {
    mobileToggle.addEventListener("click", () => {
      const isOpen = nav.classList.toggle("open");
      mobileToggle.setAttribute("aria-expanded", String(isOpen));
    });

    nav.querySelectorAll("a").forEach((link) => {
      link.addEventListener("click", () => {
        nav.classList.remove("open");
        mobileToggle.setAttribute("aria-expanded", "false");
      });
    });
  }

  const current = location.pathname.split("/").pop() || "index.html";
  document.querySelectorAll("[data-nav-link]").forEach((link) => {
    const href = link.getAttribute("href");
    if (href === current) {
      link.classList.add("active");
    }
  });

  const forceFreshLoginKey = "omkar_force_fresh_login";
  const internalNavigationKey = "omkar_internal_auth_navigation";
  // These pages manage their own sign-in flow.  Letting both this shared
  // script and the page script clear a just-created session can cancel the
  // first sign-in attempt on slower connections.
  const isAuthPortal = Boolean(
    document.getElementById("auth-panel") || document.getElementById("admin-access-form")
  );

  function safeStorage(action) {
    try { return action(); } catch (_) { return null; }
  }

  function preserveAuthForInternalNavigation(event) {
    const link = event.target.closest("a[href]");
    if (!link || event.defaultPrevented || link.target === "_blank" || event.button > 0) return;
    const destination = new URL(link.href, window.location.href);
    if (destination.origin === window.location.origin && (destination.pathname !== window.location.pathname || destination.search !== window.location.search)) {
      safeStorage(() => sessionStorage.setItem(internalNavigationKey, "1"));
    }
  }

  function installBrowserCloseLogout(auth, authMod) {
    if (window.__omkarCloseLogoutInstalled) return;
    window.__omkarCloseLogoutInstalled = true;
    document.addEventListener("click", preserveAuthForInternalNavigation);
    window.addEventListener("pagehide", () => {
      const isInternalNavigation = safeStorage(() => sessionStorage.getItem(internalNavigationKey) === "1");
      if (isInternalNavigation) {
        safeStorage(() => sessionStorage.removeItem(internalNavigationKey));
        return;
      }
      // This synchronous marker also covers mobile browsers that restore a closed tab.
      safeStorage(() => localStorage.setItem(forceFreshLoginKey, "1"));
      authMod.signOut(auth).catch(() => {});
    });
  }

  // On public pages, use the same navigation control for login and logout.
  // The dashboard and employee portal also retain their own Logout buttons.
  const loginLinks = document.querySelectorAll(".login-link");
  if (config.firebase && config.firebase.apiKey) {
    Promise.all([
      import("https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js"),
      import("https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js")
    ])
      .then(async ([appMod, authMod]) => {
        const app = appMod.getApps()[0] || appMod.initializeApp(config.firebase);
        const auth = authMod.getAuth(app);
        await authMod.setPersistence(auth, authMod.browserSessionPersistence);
        if (!isAuthPortal && safeStorage(() => localStorage.getItem(forceFreshLoginKey) === "1")) {
          await authMod.signOut(auth);
          safeStorage(() => localStorage.removeItem(forceFreshLoginKey));
        }
        installBrowserCloseLogout(auth, authMod);

        authMod.onAuthStateChanged(auth, (user) => {
          loginLinks.forEach((link) => {
            link.onclick = null;
            if (!user) {
              link.textContent = "Login";
              link.href = "onboarding.html";
              link.removeAttribute("aria-label");
              return;
            }

            link.textContent = "Logout";
            link.href = "#logout";
            link.setAttribute("aria-label", "Log out of your account");
            link.onclick = async (event) => {
              event.preventDefault();
              try {
                await authMod.signOut(auth);
              } finally {
                window.location.href = "onboarding.html";
              }
            };
          });
        });
      })
      .catch(() => {
        // Leave the normal Login link available if Firebase is temporarily unavailable.
      });
  }

  const revealItems = document.querySelectorAll(".reveal");
  if ("IntersectionObserver" in window && revealItems.length) {
    const observer = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            entry.target.classList.add("visible");
            observer.unobserve(entry.target);
          }
        });
      },
      { threshold: 0.1 }
    );

    revealItems.forEach((item, index) => {
      item.style.transitionDelay = `${Math.min(index * 45, 260)}ms`;
      observer.observe(item);
    });
  } else {
    revealItems.forEach((item) => item.classList.add("visible"));
  }

  const motionItems = document.querySelectorAll(".card, .partner-chip, .form-shell, .map-frame, .table-wrap");
  if ("IntersectionObserver" in window && motionItems.length) {
    const motionObserver = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            entry.target.classList.add("motion-in");
            motionObserver.unobserve(entry.target);
          }
        });
      },
      { threshold: 0.12 }
    );

    motionItems.forEach((item, index) => {
      item.classList.add("motion-item");
      item.style.setProperty("--motion-delay", `${(index % 5) * 70}ms`);
      motionObserver.observe(item);
    });
  }

  if (config.analyticsMeasurementId) {
    const script = document.createElement("script");
    script.async = true;
    script.src = `https://www.googletagmanager.com/gtag/js?id=${config.analyticsMeasurementId}`;
    document.head.appendChild(script);

    window.dataLayer = window.dataLayer || [];
    window.gtag = function () {
      dataLayer.push(arguments);
    };
    window.gtag("js", new Date());
    window.gtag("config", config.analyticsMeasurementId);
  }
})();


