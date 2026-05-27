(function () {
  const jobs = window.OMKAR_JOBS || [];
  const jobsHost = document.getElementById("jobs-list");
  const locationFilter = document.getElementById("filter-location");
  const roleFilter = document.getElementById("filter-role");

  if (!jobsHost) {
    return;
  }

  const fillRoleFilter = () => {
    if (!roleFilter) {
      return;
    }
    const uniqueRoles = [...new Set(jobs.map((job) => job.role))];
    uniqueRoles.forEach((role) => {
      const option = document.createElement("option");
      option.value = role;
      option.textContent = role;
      roleFilter.appendChild(option);
    });
  };

  const createCard = (job) => {
    const card = document.createElement("article");
    card.className = "card reveal";
    card.innerHTML = `
      <h3>${job.role}</h3>
      <p>${job.summary}</p>
      <div class="meta">
        <span>${job.location}</span>
        <span>${job.type}</span>
        <span>${job.shift}</span>
      </div>
      <p style="margin-top:10px;"><strong>${job.salary}</strong></p>
      <div style="margin-top:14px; display:flex; gap:10px; flex-wrap:wrap;">
        <a class="btn btn-primary" href="apply.html?role=${encodeURIComponent(job.role)}&location=${encodeURIComponent(job.location)}">Apply Now</a>
      </div>
    `;
    return card;
  };

  const render = () => {
    const locationValue = locationFilter ? locationFilter.value : "All";
    const roleValue = roleFilter ? roleFilter.value : "All";

    const filtered = jobs.filter((job) => {
      const locationMatch = locationValue === "All" || job.location === locationValue;
      const roleMatch = roleValue === "All" || job.role === roleValue;
      return locationMatch && roleMatch;
    });

    jobsHost.innerHTML = "";
    if (!filtered.length) {
      jobsHost.innerHTML =
        '<article class="card"><h3>No jobs found</h3><p>Try a different location or role filter.</p></article>';
      return;
    }

    filtered.forEach((job) => jobsHost.appendChild(createCard(job)));
  };

  fillRoleFilter();
  render();

  if (locationFilter) locationFilter.addEventListener("change", render);
  if (roleFilter) roleFilter.addEventListener("change", render);
})();
