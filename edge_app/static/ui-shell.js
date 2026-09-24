/* Navegação e tema compartilhados pelas telas do portal. */
(() => {
  function init() {
    const sidebar = document.getElementById('edgeAppSidebar');
    const menuButton = document.getElementById('edgeNavToggle');
    const mainNav = document.getElementById('edgeMainNav');

    function closeMenu() {
      if (!sidebar || !menuButton) return;
      sidebar.classList.remove('is-open');
      menuButton.setAttribute('aria-expanded', 'false');
    }

    if (sidebar && menuButton && mainNav) {
      menuButton.addEventListener('click', () => {
        const isOpen = sidebar.classList.toggle('is-open');
        menuButton.setAttribute('aria-expanded', String(isOpen));
      });
      document.addEventListener('keydown', event => {
        if (event.key === 'Escape' && sidebar.classList.contains('is-open')) {
          closeMenu();
          menuButton.focus();
        }
      });
      window.matchMedia('(min-width: 901px)').addEventListener('change', closeMenu);
    }

    const themeButton = document.getElementById('edgeThemeToggle');
    function setTheme(theme) {
      const next = theme === 'dark' ? 'dark' : 'light';
      document.documentElement.setAttribute('data-theme', next);
      try { localStorage.setItem('edge-theme', next); } catch (_) { /* storage opcional */ }
      if (themeButton) {
        const label = next === 'dark' ? 'Tema claro' : 'Tema escuro';
        const icon = themeButton.querySelector('.edge-theme-icon');
        const text = themeButton.querySelector('.edge-theme-text');
        if (icon) icon.textContent = next === 'dark' ? '☀' : '☾';
        if (text) text.textContent = label;
        themeButton.setAttribute('aria-label', next === 'dark' ? 'Ativar tema claro' : 'Ativar tema escuro');
      }
      const meta = document.querySelector('meta[name="theme-color"]');
      if (meta) meta.setAttribute('content', next === 'dark' ? '#10161a' : '#f3f7f5');
    }
    if (themeButton) {
      setTheme(document.documentElement.getAttribute('data-theme'));
      themeButton.addEventListener('click', () => {
        setTheme(document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark');
      });
    }

    const main = document.querySelector('.app-layout > .main-area');
    const heading = main?.querySelector('h1');
    if (main && heading && location.pathname !== '/' && !main.querySelector('.edge-breadcrumb')) {
      const breadcrumb = document.createElement('nav');
      breadcrumb.className = 'edge-breadcrumb';
      breadcrumb.setAttribute('aria-label', 'Caminho da página');
      const home = document.createElement('a');
      home.href = '/';
      home.textContent = 'Início';
      breadcrumb.append(home);
      const group = sidebar?.querySelector('.edge-nav-group[open] summary');
      if (group) {
        const separator = document.createElement('span');
        separator.textContent = '/';
        const groupName = document.createElement('span');
        groupName.textContent = group.textContent.trim();
        breadcrumb.append(separator, groupName);
      }
      const separator = document.createElement('span');
      separator.textContent = '/';
      const current = document.createElement('strong');
      current.textContent = heading.textContent.trim();
      breadcrumb.append(separator, current);
      main.prepend(breadcrumb);
    }
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
