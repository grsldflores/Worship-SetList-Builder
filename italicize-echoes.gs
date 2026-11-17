/*
  ECHOES IN BOLD ITALIC: add-on for the Google connection script

  What it does: after the app creates or updates a setlist Google Doc, it makes
  the echo parts of the lyrics bold and italic:
    (oh-oh)                  → keeps its parentheses
    [se moverá, se moverá]   → the brackets are removed
  Notes like (x2), (3x), (solo), (start high), (para terminar) stay normal.
  The song summary list at the top and the song titles are not touched.

  HOW TO ADD IT:
  1. Open the Apps Script.
  2. In doPost, add this line right under the 'updateSetlistDoc' line
     (skip this if you already added it):
       case 'italicizeEchoes': return out_(italicizeEchoes_(req));
  3. Paste everything below this comment at the very bottom (replacing any
     earlier italicizeEchoes_ you pasted) and click Save.
  4. Deploy → Manage deployments → pencil icon → Version: New version → Deploy.
     The URL stays the same.
*/

function italicizeEchoes_(req) {
  var m = String(req.docUrl || '').match(/[-\w]{25,}/);
  if (!m) throw new Error('bad doc link');
  var doc = DocumentApp.openById(m[0]);
  var body = doc.getBody();

  var titles = (req.titles || []).map(function (t) { return String(t).trim().toLowerCase(); });
  var NOTE = /^\(\s*(x\s*\d+|\d+\s*x|solo\b.*|start\b.*|high|low\b.*|full run-through|to end\b.*|para terminar|repe\w*.*|\d+(st|nd|rd|th)\b.*)\s*\)$/i;
  var count = 0;

  for (var i = 0; i < body.getNumChildren(); i++) {
    var el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;   // skips the bullet summary
    var p = el.asParagraph();
    if (p.getHeading() !== DocumentApp.ParagraphHeading.NORMAL) continue; // skips headings
    var text = p.getText();
    if (!text || (text.indexOf('(') < 0 && text.indexOf('[') < 0)) continue;
    var low = text.trim().toLowerCase();
    if (titles.some(function (t) { return low === t || low.indexOf(t + ' (') === 0; })) continue; // song titles / keys

    var edit = p.editAsText();
    var re = /\([^()]*\)|\[[^\[\]]*\]/g, x, hits = [];
    while ((x = re.exec(text))) {
      var br = x[0].charAt(0) === '[';
      if (!br && NOTE.test(x[0])) continue;
      hits.push([x.index, x.index + x[0].length - 1, br]);
    }
    // work from the end so earlier positions don't shift
    for (var h = hits.length - 1; h >= 0; h--) {
      var a = hits[h][0], b = hits[h][1];
      edit.setBold(a, b, true);
      edit.setItalic(a, b, true);
      if (hits[h][2]) { edit.deleteText(b, b); edit.deleteText(a, a); } // [ ] are hidden, ( ) stay
      count++;
    }
  }
  doc.saveAndClose();
  return { ok: true, count: count };
}
