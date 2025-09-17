Office.onReady(() => {
  console.log('here we are!!!')
  window.noop = function (event) {
    try { event.completed && event.completed(); } catch (_) {}
  };
});
