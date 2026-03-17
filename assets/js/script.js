/* ============================================
   PORTFOLIO - script.js
   Efeitos visuais e interatividade
   ============================================ */

document.addEventListener('DOMContentLoaded', () => {

  /* --- Navbar: trocar de light para solid ao rolar --- */
  const navbar = document.querySelector('.navbar');
  const hero = document.querySelector('.hero');

  if (navbar && (navbar.classList.contains('navbar--light') || navbar.classList.contains('navbar--dark')) && hero) {
    const handleScroll = () => {
      const heroBottom = hero.offsetTop + hero.offsetHeight;
      if (window.scrollY > heroBottom - 80) {
        navbar.classList.add('scrolled');
        navbar.classList.remove('navbar--light', 'navbar--dark');
      } else {
        navbar.classList.remove('scrolled');
        if (document.querySelector('.hero--dark')) {
          navbar.classList.add('navbar--dark');
        } else {
          navbar.classList.add('navbar--light');
        }
      }
    };

    window.addEventListener('scroll', handleScroll, { passive: true });
    handleScroll(); // estado inicial
  }

  /* --- Mobile menu toggle --- */
  const toggle = document.querySelector('.navbar__toggle');
  const links = document.querySelector('.navbar__links');

  if (toggle && links) {
    toggle.addEventListener('click', () => {
      links.classList.toggle('open');
      const icon = toggle.querySelector('i');
      if (icon) {
        icon.classList.toggle('fa-bars');
        icon.classList.toggle('fa-xmark');
      }
    });

    links.querySelectorAll('a').forEach(link => {
      link.addEventListener('click', () => {
        links.classList.remove('open');
        const icon = toggle.querySelector('i');
        if (icon) {
          icon.classList.remove('fa-xmark');
          icon.classList.add('fa-bars');
        }
      });
    });
  }

  /* --- Active nav link baseado na página atual --- */
  const path = window.location.pathname || window.location.href;
  const currentPage = path.split('/').pop() || 'index.html';
  const isProjectDetail = path.includes('/projects/') || path.includes('projects\\');

  document.querySelectorAll('.navbar__links a').forEach(link => {
    link.classList.remove('active');
    const href = link.getAttribute('href');
    const hrefPage = (href || '').replace(/^\.\.\//, '').split('/').pop();
    const matches = hrefPage === currentPage || (currentPage === '' && href.includes('index.html'));
    const projectsMatch = isProjectDetail && href.includes('projects');
    if (matches || projectsMatch) {
      link.classList.add('active');
    }
  });

  /* --- Smooth scroll para âncoras --- */
  document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', (e) => {
      const targetId = anchor.getAttribute('href');
      if (targetId === '#') return;
      const target = document.querySelector(targetId);
      if (target) {
        e.preventDefault();
        target.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    });
  });

  /* --- Reveal animation ao rolar (Intersection Observer) --- */
  const revealElements = document.querySelectorAll('.reveal, .project-meta-card, .project-section');

  if (revealElements.length > 0) {
    const revealObserver = new IntersectionObserver((entries) => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          entry.target.classList.add('visible');
        }
      });
    }, {
      threshold: 0.1,
      rootMargin: '0px 0px -30px 0px'
    });

    revealElements.forEach((el, i) => {
      el.style.opacity = '0';
      el.style.transform = 'translateY(24px)';
      el.style.transition = `opacity 0.6s ease ${i * 0.05}s, transform 0.6s ease ${i * 0.05}s`;
      revealObserver.observe(el);
    });
  }

  /* --- Efeito hover suave nos cards (opcional - já no CSS) --- */
  document.querySelectorAll('.project-card, .skill-card').forEach(card => {
    card.addEventListener('mouseenter', function() {
      this.style.transition = 'all 0.3s cubic-bezier(0.4, 0, 0.2, 1)';
    });
  });

});
