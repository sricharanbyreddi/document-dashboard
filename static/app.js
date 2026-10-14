const fileInput = document.getElementById('fileInput');
const addButton = document.getElementById('addButton');
const uploadForm = document.getElementById('uploadForm');
addButton?.addEventListener('click', () => fileInput.click());
fileInput?.addEventListener('change', () => { if (fileInput.files.length) uploadForm.submit(); });

const search = document.getElementById('docSearch');
const links = [...document.querySelectorAll('.document-link')];
const noResults = document.getElementById('noResults');
search?.addEventListener('input', () => {
  const query = search.value.trim().toLowerCase();
  let visible = 0;
  links.forEach(link => {
    const match = link.dataset.name.includes(query);
    link.hidden = !match;
    if (match) visible++;
  });
  if (noResults) noResults.hidden = visible !== 0;
});

const themeToggle = document.getElementById('themeToggle');
const themeIcon = document.getElementById('themeIcon');
const setTheme = light => {
  document.body.classList.toggle('light-theme', light);
  if (themeIcon) themeIcon.textContent = light ? '☀' : '☾';
  try { localStorage.setItem('document-dashboard-theme', light ? 'light' : 'dark'); } catch (_) {}
};
let savedTheme = 'dark';
try { savedTheme = localStorage.getItem('document-dashboard-theme') || 'dark'; } catch (_) {}
setTheme(savedTheme === 'light');
themeToggle?.addEventListener('click', () => setTheme(!document.body.classList.contains('light-theme')));

if (window.mermaid) {
  mermaid.initialize({ startOnLoad: true, theme: 'dark', securityLevel: 'strict' });
}
