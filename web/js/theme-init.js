/* Runs before first paint to avoid a flash of the wrong theme. */
(function () {
  var t = 'light';
  try {
    t = localStorage.getItem('webypdf-theme') || 'light';
  } catch (e) { /* storage blocked */ }
  document.documentElement.setAttribute('data-theme', t);
})();
