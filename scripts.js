// =========================================================
// Fatou Mané Ndiaye — Portfolio scripts
// =========================================================

document.addEventListener('DOMContentLoaded', () => {

  // ----- Footer year -----
  const yearEl = document.getElementById('year');
  if (yearEl) yearEl.textContent = new Date().getFullYear();

  // ----- Mobile menu toggle -----
  const menuToggle = document.getElementById('menuToggle');
  const navbar = document.getElementById('navbar');

  if (menuToggle && navbar) {
    menuToggle.addEventListener('click', () => {
      menuToggle.classList.toggle('active');
      navbar.classList.toggle('active');
    });

    // Close menu when a link is clicked
    navbar.querySelectorAll('.nav-link').forEach(link => {
      link.addEventListener('click', () => {
        menuToggle.classList.remove('active');
        navbar.classList.remove('active');
      });
    });
  }

  // ----- Sticky header shadow on scroll -----
  const header = document.querySelector('.site-header');
  const backToTop = document.getElementById('backToTop');

  const onScroll = () => {
    const scrolled = window.scrollY > 20;
    if (header) header.classList.toggle('scrolled', scrolled);
    if (backToTop) backToTop.classList.toggle('visible', window.scrollY > 400);
  };
  document.addEventListener('scroll', onScroll, { passive: true });
  onScroll();

  // ----- Active nav link on scroll -----
  const sections = document.querySelectorAll('section[id]');
  const navLinks = document.querySelectorAll('.nav-link');

  const setActiveLink = () => {
    let current = '';
    sections.forEach(section => {
      const sectionTop = section.offsetTop - 120;
      if (window.scrollY >= sectionTop) {
        current = section.getAttribute('id');
      }
    });
    navLinks.forEach(link => {
      link.classList.toggle('active', link.getAttribute('href') === `#${current}`);
    });
  };
  document.addEventListener('scroll', setActiveLink, { passive: true });
  setActiveLink();

  // ----- Reveal on scroll (Intersection Observer) -----
  const revealTargets = document.querySelectorAll(
    '.about-text, .about-facts .fact, .skill-card, .timeline-item, .education-item, .project-card, .contact-card'
  );
  revealTargets.forEach(el => el.classList.add('reveal'));

  const observer = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        entry.target.classList.add('visible');
        observer.unobserve(entry.target);
      }
    });
  }, { threshold: 0.15 });

  revealTargets.forEach(el => observer.observe(el));
});
