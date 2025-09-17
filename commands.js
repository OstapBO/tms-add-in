// const makeToken = (key) => `{{${key}}}`;
// function insertVariable(event) {
//   const controlId = event.source && event.source.id ? event.source.id : '';
//   const key = controlId.startsWith('mf.') ? controlId.slice(3) : controlId || 'Unknown';

//   Word.run(async (context) => {
//     const selection = context.document.getSelection();
//     const token = makeToken(key);
//     selection.insertText(token, Word.InsertLocation.replace);
//     const cc = selection.insertContentControl();
//     cc.tag = key;
//     cc.title = key;
//     cc.appearance = 'BoundingBox';

//     await context.sync();
//     event.completed();
//   }).catch((error) => {
//     console.error('Insert error:', error);
//     try {
//       event.completed();
//     } catch {}
//   });
// }

// // Must be globally visible for Office to find it
// Office.onReady(() => {});

// window.insertVariable = insertVariable;

Office.onReady(() => {
  console.log('here we are!!!')
  window.noop = function (event) {
    try { event.completed && event.completed(); } catch (_) {}
  };
});
