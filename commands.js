const makeToken = (key) => `{{${key}}}`;

function insertVariable(event) {
  const controlId = event.source?.id ?? '';
  const key = controlId.startsWith('mf.') ? controlId.slice(3) : controlId || 'Unknown';

  Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertText(makeToken(key), Word.InsertLocation.replace);

    const cc = selection.insertContentControl();
    cc.tag = key;
    cc.title = key;
    cc.appearance = 'BoundingBox';

    await context.sync();
    event.completed();
  }).catch((error) => {
    console.error('Insert error:', error);
    try {
      event.completed();
    } catch (_) {}
  });
}

Office.onReady(() => {
  console.log('here we are!!!');
  window.insertVariable = insertVariable; // keep available for future commands
  window.noop = (event) => {
    try {
      event.completed?.();
    } catch (_) {}
  };
});
