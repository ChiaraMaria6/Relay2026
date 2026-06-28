const navButtons = document.querySelectorAll(".nav-btn");
const panels = document.querySelectorAll(".tab-panel");
const pageTitle = document.getElementById("pageTitle");
const adminMode = document.getElementById("adminMode");
const lightbox = document.getElementById("lightbox");
const lightboxImg = document.getElementById("lightboxImg");
const setupSearch = document.getElementById("setupSearch");

navButtons.forEach(button => {
  button.addEventListener("click", () => {
    const target = button.dataset.tab;
    navButtons.forEach(btn => btn.classList.remove("active"));
    panels.forEach(panel => panel.classList.remove("active"));
    button.classList.add("active");
    document.getElementById(target).classList.add("active");
    pageTitle.textContent = button.textContent.trim();
    window.scrollTo({ top: 0, behavior: "smooth" });
  });
});

adminMode.addEventListener("change", () => {
  document.body.classList.toggle("admin-on", adminMode.checked);
});

document.querySelectorAll(".reference-img").forEach(img => {
  img.addEventListener("click", () => {
    lightboxImg.src = img.src;
    lightbox.classList.add("show");
  });
});

lightbox.addEventListener("click", () => {
  lightbox.classList.remove("show");
  lightboxImg.src = "";
});

if (setupSearch) {
  setupSearch.addEventListener("input", () => {
    const query = setupSearch.value.trim().toLowerCase();
    document.querySelectorAll(".searchable").forEach(card => {
      const text = card.dataset.search || "";
      card.style.display = text.includes(query) ? "" : "none";
    });
  });
}
