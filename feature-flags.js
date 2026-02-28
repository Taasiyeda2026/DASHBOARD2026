const ENABLE_SCHEDULING = false;
window.ENABLE_SCHEDULING = ENABLE_SCHEDULING;

if (!ENABLE_SCHEDULING && window.location.hash === '#scheduling') {
  window.location.hash = '#dashboard';
}
