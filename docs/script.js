(function () {
  const defaults = {
    companyName: "OMKAR RETAIL VENTURES",
    domain: "omkarretailventures.in",
    industry: "Zepto Dark Store Hiring",
    location: "Bangalore, Karnataka, India",
    phone: "9986362446",
    whatsappNumber: "919986362446",
    email: "omkarretailventure@gmail.com",
    officeAddress: "OMKAR RETAIL VENTURES - Zepto Dark Store Hiring Desk, Bangalore, Karnataka, India",
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


