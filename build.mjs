// Construit dist/Couteau-suisse.html : un seul fichier autonome (hors ligne) contenant
// l'interface, SheetJS, JSZip et le modèle Word extrait de modele.docx.
// Usage : node build.mjs
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const dir = path.dirname(fileURLToPath(import.meta.url));
const MODELE = path.join(dir, 'modele.docx');

// Découpe le contenu de <w:body> en éléments de premier niveau (w:p, w:tbl, w:sectPr).
function topLevel(body) {
  const tok = /<(\/?)(w:p|w:tbl|w:sectPr)(?=[\s>\/])[^>]*?(\/?)>/g;
  const out = [];
  let depth = 0, start = 0, tag = null, m;
  while ((m = tok.exec(body))) {
    const [, close, name, self] = m;
    if (depth === 0 && !close) {
      start = m.index; tag = name;
      if (self) { out.push(body.slice(start, tok.lastIndex)); continue; }
    }
    if (name !== tag) continue;
    if (close) depth--; else if (!self) depth++;
    if (depth === 0) out.push(body.slice(start, tok.lastIndex));
  }
  return out;
}

const firstMatch = (s, re, what) => {
  const m = s.match(re);
  if (!m) throw new Error(`Modèle : ${what} introuvable`);
  return m[0];
};
const pPrOf = (p) => firstMatch(p, /<w:pPr>[\s\S]*?<\/w:pPr>/, 'w:pPr');
const runRPrOf = (p) => firstMatch(p, /<w:r[ >](?:(?!<\/w:r>)[\s\S])*?<w:rPr>[\s\S]*?<\/w:rPr>/, 'w:rPr').replace(/^[\s\S]*(<w:rPr>)/, '$1');
const openTag = (p) => firstMatch(p, /^<w:p[^>]*>/, 'w:p');
const paragraphs = (s) => s.match(/<w:p[ >][\s\S]*?<\/w:p>|<w:p\/>/g) || [];

async function extractTemplate(buf) {
  const zip = await JSZip.loadAsync(buf);
  const xml = await zip.file('word/document.xml').async('string');
  const b = xml.indexOf('<w:body>') + '<w:body>'.length;
  const els = topLevel(xml.slice(b, xml.indexOf('</w:body>')));
  const sectPr = els.find((e) => e.startsWith('<w:sectPr'));
  // La première page s'arrête au paragraphe « Merci de déposer… ».
  const end = els.findIndex((e) => e.includes('secrétariat'));
  const tblIdx = els.findIndex((e) => e.startsWith('<w:tbl') && e.includes('>7P<'));
  if (end < 0 || tblIdx < 0 || tblIdx > end) throw new Error('Modèle : structure de page inattendue');
  const page = els.slice(0, end + 1);

  // Titre
  const title = page[0];
  page[0] = openTag(title) + pPrOf(title).replace('<w:pPr>', '<w:pPr>{{PAGEBREAK}}') +
    `<w:r>${runRPrOf(title)}<w:t xml:space="preserve">{{TITRE}}</w:t></w:r></w:p>`;

  // Ligne « Enseignant – salle … »
  const hIdx = page.findIndex((e) => e.includes('salle'));
  const header = page[hIdx];
  const headerRPr = runRPrOf(header).replace(/<w:highlight [^>]*\/>/, '');
  page[hIdx] = openTag(header) + pPrOf(header) + '{{ENTETE}}</w:p>';

  // Tableau des élèves : la 1re ligne (« 7P | 8P ») et la 2e (les cases) servent de modèles,
  // les lignes sont ensuite générées selon les niveaux présents ({{ROWS}}).
  const tbl = page[tblIdx];
  const trs = tbl.match(/<w:tr[ >][\s\S]*?<\/w:tr>/g);
  if (!trs || trs.length < 2) throw new Error('Modèle : tableau des élèves inattendu');
  let cellNo = 0;
  const headerRow = trs[0].replace(/<w:t>[^<]*<\/w:t>/g, () => `<w:t>{{H${++cellNo}}}</w:t>`);
  if (cellNo !== 2) throw new Error('Modèle : en-têtes 7P / 8P introuvables');
  let rowHeight = 2084, spacerXml = null, studentPPr = null, studentRPr = null;
  cellNo = 0;
  const contentRow = trs[1]
    .replace(/<w:trHeight w:val="(\d+)"/, (m, h) => { rowHeight = +h; return '<w:trHeight w:val="{{HEIGHT}}"'; })
    .replace(/<w:tc>([\s\S]*?)<\/w:tc>/g, (tc, inner) => {
      const c = cellNo++;
      if (c === 0) return tc;
      const ps = paragraphs(inner);
      if (!spacerXml) spacerXml = ps[0];
      const sp = ps.find((p) => /<w:t[ >]/.test(p) && /\(\d/.test(p));
      if (sp && !studentPPr) { studentPPr = pPrOf(sp); studentRPr = runRPrOf(sp); }
      const tcPr = (inner.match(/<w:tcPr>[\s\S]*?<\/w:tcPr>/) || [''])[0];
      return `<w:tc>${tcPr}{{C${c}}}</w:tc>`;
    });
  if (!spacerXml || !studentPPr) throw new Error('Modèle : ligne élève introuvable dans le tableau');
  page[tblIdx] = tbl.slice(0, tbl.indexOf(trs[0])) + '{{ROWS}}' +
    tbl.slice(tbl.indexOf(trs[trs.length - 1]) + trs[trs.length - 1].length);

  const stripIds = (s) => s.replace(/ w14:(paraId|textId)="[^"]*"/g, '');
  return {
    docStart: xml.slice(0, b),
    pageXml: stripIds(page.join('')),
    sectPr,
    headerRPr,
    headerRow: stripIds(headerRow),
    contentRow: stripIds(contentRow),
    rowHeight,
    headerHeight: +((trs[0].match(/<w:trHeight w:val="(\d+)"/) || [0, 0])[1]),
    spacerXml: stripIds(spacerXml),
    studentPPr,
    studentRPr: studentRPr.replace(/<w:rStyle [^>]*\/>/, ''),
  };
}

const inlineJs = (s) => s.replace(/<\/script/gi, '<\\/script');

const buf = fs.readFileSync(MODELE);
const tpl = await extractTemplate(buf);
if (process.argv.includes('--debug')) fs.writeFileSync(path.join(dir, 'dist-template.json'), JSON.stringify(tpl, null, 1));
tpl.docxBase64 = buf.toString('base64');

const read = (p) => fs.readFileSync(path.join(dir, p), 'utf8');
const html = read('src/index.html')
  .replace('/*__LIBS__*/', () => inlineJs(read('node_modules/xlsx/dist/xlsx.mini.min.js') + '\n' +
    read('node_modules/jszip/dist/jszip.min.js')))
  .replace('/*__CORE__*/', () => inlineJs(read('src/core.js')))
  .replace('/*__TEMPLATE__*/', () => 'window.TEMPLATE = ' + inlineJs(JSON.stringify(tpl)) + ';')
  .replace('/*__APP__*/', () => inlineJs(read('src/app.js')))
  .replace('__ACCUEIL_IMG__', () => 'data:image/jpeg;base64,' + fs.readFileSync(path.join(dir, 'src/accueil.jpg')).toString('base64'));

fs.mkdirSync(path.join(dir, 'dist'), { recursive: true });
fs.writeFileSync(path.join(dir, 'dist/Couteau-suisse.html'), html);
console.log(`dist/Couteau-suisse.html (${Math.round(html.length / 1024)} Ko)`);
