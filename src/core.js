// Logique sans interface : lecture des lignes Excel, répartition, génération du document.xml.
// Utilisable dans le navigateur (window.Core) et dans Node (require) pour les tests.
(function (root) {
  'use strict';

  const JOURS = ['Dimanche', 'Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi'];
  const MOIS = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août',
    'septembre', 'octobre', 'novembre', 'décembre'];

  const clean = (v) => (v == null ? '' : String(v)).replace(/\s+/g, ' ').trim();
  const key = (v) => clean(v).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');

  function normalizeClass(raw) {
    return clean(raw).toUpperCase().replace(/\//g, '').replace(/\s+/g, ' ');
  }

  const AUTRES = 'Autres';

  // Rubrique du tableau Word (6P, 7P, 9VG…). « 7-8 EI » va en 7P comme dans le modèle.
  function categoryOf(classe) {
    const c = classe.replace(/\s/g, '');
    let m = c.match(/^(\d{1,2})(P|VG|VP)/);
    if (m) {
      const n = +m[1];
      if (m[2] === 'P' ? n >= 1 && n <= 8 : n >= 9 && n <= 11) return { cat: n + m[2], sure: true };
    }
    m = c.match(/^(\d{1,2})/);
    const n = m ? parseInt(m[1], 10) : NaN;
    if (n >= 1 && n <= 8) return { cat: n + 'P', sure: /^\d-\d/.test(c) };
    if (n >= 9 && n <= 11) return { cat: n + 'VG', sure: false };
    return { cat: AUTRES, sure: true };
  }

  // Ordre d'affichage : 1P…8P, 9VG, 9VP, …, 11VP, puis Autres.
  function categoryRank(cat) {
    const m = cat.match(/^(\d+)(P|VG|VP)$/);
    if (!m) return 1000;
    return +m[1] * 10 + ['P', 'VG', 'VP'].indexOf(m[2]);
  }

  // Lignes du tableau Word (paires de cases). Les 4 lignes du modèle sont toujours là ;
  // 1P|2P, 3P|4P, 5P|6P et Autres seulement si un élève en a besoin.
  function tableLayout(students) {
    const used = new Set(students.map((s) => s.cat));
    const rows = [];
    for (const n of [1, 3, 5]) {
      if (used.has(n + 'P') || used.has(n + 1 + 'P')) rows.push([n + 'P', n + 1 + 'P']);
    }
    rows.push(['7P', '8P'], ['9VG', '9VP'], ['10VG', '10VP'], ['11VG', '11VP']);
    if (used.has(AUTRES)) rows.push([AUTRES, null]);
    return rows;
  }

  function excelDate(v) {
    if (typeof v === 'number' && v > 20000) return new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
    if (v instanceof Date) return new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate()));
    const m = clean(v).match(/(\d{1,2})[./](\d{1,2})[./](\d{2,4})/);
    if (m) {
      const y = m[3].length === 2 ? 2000 + +m[3] : +m[3];
      return new Date(Date.UTC(y, +m[2] - 1, +m[1]));
    }
    return null;
  }

  function formatDateLong(d) {
    return `${JOURS[d.getUTCDay()]} ${d.getUTCDate()} ${MOIS[d.getUTCMonth()]} ${d.getUTCFullYear()}`;
  }

  function findCol(header, tests, fallback) {
    const h = header.map(key);
    for (const t of tests) {
      const i = h.findIndex((x) => x.startsWith(t));
      if (i >= 0) return i;
    }
    return fallback;
  }

  // rows : tableau de lignes (sheet_to_json avec header:1), ligne 0 = en-têtes.
  function parseRows(rows, sheetName) {
    const header = (rows[0] || []).map(clean);
    const col = {
      nom: findCol(header, ["nom de l'eleve", 'nom de l’eleve'], 2),
      prenom: findCol(header, ['prenom de'], 3),
      classe: findCol(header, ['classe'], 4),
      heures: findCol(header, ["nombre d'heures", 'nombre d’heures', 'nombre'], 5),
      date: findCol(header, ['date des ha'], 7),
      adresse: findCol(header, ['adresse des parents'], 13),
    };
    const students = [];
    const dates = {};
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const nom = clean(row[col.nom]);
      const classeRaw = clean(row[col.classe]);
      if (!nom || !classeRaw) continue;
      const classe = normalizeClass(classeRaw);
      const { cat, sure } = categoryOf(classe);
      const pm = clean(row[col.heures]).match(/\d+/);
      const periodes = pm ? parseInt(pm[0], 10) : 0;
      const d = excelDate(row[col.date]);
      if (d) dates[d.getTime()] = (dates[d.getTime()] || 0) + 1;
      students.push({
        id: students.length,
        nom,
        prenom: clean(row[col.prenom]),
        classe,
        cat,
        catSure: sure,
        periodes,
        adresse: key(row[col.adresse]),
      });
    }
    // Date : la plus fréquente dans la colonne « Date des HA », sinon le nom de la feuille.
    let date = null;
    const best = Object.entries(dates).sort((a, b) => b[1] - a[1])[0];
    if (best) date = new Date(+best[0]);
    else date = excelDate(sheetName);
    return { students, date };
  }

  function periodLabel(n) {
    if (!n) return '';
    return `${n} période${n > 1 ? 's' : ''}`;
  }

  function studentLine(s) {
    const p = periodLabel(s.periodes);
    return `${s.nom} ${s.prenom} – ${s.classe}${p ? ` (${p})` : ''}`;
  }

  // Deux élèves de la même famille : même nom de famille ou même adresse des parents.
  function sameFamily(a, b) {
    if (key(a.nom) === key(b.nom)) return true;
    return !!a.adresse && a.adresse === b.adresse;
  }

  function conflicts(students, assign, teacherCount) {
    const byT = Array.from({ length: teacherCount }, () => []);
    students.forEach((s) => { if (assign[s.id] >= 0) byT[assign[s.id]].push(s); });
    const out = {}; // id -> {famille:[ids], classe:[ids]}
    let famille = 0, classe = 0;
    for (const list of byT) {
      for (let i = 0; i < list.length; i++) {
        for (let j = i + 1; j < list.length; j++) {
          const a = list[i], b = list[j];
          const f = sameFamily(a, b), c = a.classe === b.classe;
          if (!f && !c) continue;
          for (const [x, y] of [[a, b], [b, a]]) {
            out[x.id] = out[x.id] || { famille: [], classe: [] };
            if (f) out[x.id].famille.push(y.id);
            if (c) out[x.id].classe.push(y.id);
          }
          if (f) famille++;
          if (c) classe++;
        }
      }
    }
    return { byStudent: out, famille, classe, score: famille * 100 + classe };
  }

  function lexLess(a, b) {
    for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return a[i] < b[i];
    return false;
  }

  // Répartition équilibrée : glouton puis échanges tant que le nombre de conflits baisse.
  function distribute(students, teacherCount) {
    const assign = {};
    if (teacherCount < 1) return assign;
    const cap = Math.ceil(students.length / teacherCount);
    const famSize = (s) => students.filter((o) => sameFamily(s, o)).length;
    const clsSize = (s) => students.filter((o) => o.classe === s.classe).length;
    const order = [...students].sort((a, b) =>
      famSize(b) - famSize(a) || clsSize(b) - clsSize(a) ||
      a.classe.localeCompare(b.classe) || a.nom.localeCompare(b.nom));
    const lists = Array.from({ length: teacherCount }, () => []);
    for (const s of order) {
      let best = -1, bestKey = null;
      for (let t = 0; t < teacherCount; t++) {
        if (lists[t].length >= cap) continue;
        const k = [
          lists[t].filter((o) => sameFamily(s, o)).length,
          lists[t].filter((o) => o.classe === s.classe).length,
          lists[t].length,
          lists[t].reduce((n, o) => n + o.periodes, 0),
        ];
        if (!bestKey || lexLess(k, bestKey)) {
          best = t; bestKey = k;
        }
      }
      lists[best].push(s);
      assign[s.id] = best;
    }
    let score = conflicts(students, assign, teacherCount).score;
    for (let improved = true, guard = 0; improved && guard < 50; guard++) {
      improved = false;
      for (const a of students) {
        for (const b of students) {
          if (assign[a.id] === assign[b.id]) continue;
          [assign[a.id], assign[b.id]] = [assign[b.id], assign[a.id]];
          const s2 = conflicts(students, assign, teacherCount).score;
          if (s2 < score) { score = s2; improved = true; }
          else [assign[a.id], assign[b.id]] = [assign[b.id], assign[a.id]];
        }
      }
    }
    return assign;
  }

  // ---------- Génération du document Word ----------

  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  function run(rPr, text) {
    return `<w:r>${rPr}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
  }

  // tpl : objet produit par build.mjs (voir extractTemplate) ; pages : [{enseignant, salle, students}]
  function buildDocumentXml(tpl, { date, salles, pages }) {
    const titre = `Listing des élèves aux heures d'arrêts - ${formatDateLong(date)}`;
    const hl = tpl.headerRPr.replace('</w:rPr>', '<w:highlight w:val="yellow"/></w:rPr>');
    // Même tableau sur toutes les pages ; hauteur des cases réduite s'il y a plus de 4 lignes.
    const layout = tableLayout(pages.flatMap((p) => p.students));
    const total = 4 * (tpl.rowHeight + tpl.headerHeight);
    const height = Math.floor(total / Math.max(4, layout.length) - tpl.headerHeight);
    const cell = (students, cat) => {
      if (!cat) return tpl.spacerXml;
      const list = students.filter((s) => s.cat === cat)
        .sort((a, b) => a.classe.localeCompare(b.classe, 'fr', { numeric: true }) ||
          a.nom.localeCompare(b.nom, 'fr') || a.prenom.localeCompare(b.prenom, 'fr'));
      const paras = list.map((s) => `<w:p>${tpl.studentPPr}${run(tpl.studentRPr, studentLine(s))}</w:p>`);
      return tpl.spacerXml + paras.join('') + (paras.length ? tpl.spacerXml : '');
    };
    const body = pages.map((p, i) => {
      const entete = [run(tpl.headerRPr, `${p.enseignant} – salle `)];
      salles.forEach((salle, j) => {
        if (j > 0) entete.push(run(tpl.headerRPr, ' / '));
        entete.push(run(salle === p.salle ? hl : tpl.headerRPr, salle));
      });
      const rows = layout.map(([a, b]) =>
        tpl.headerRow.replace('{{H1}}', esc(a)).replace('{{H2}}', esc(b || '')) +
        tpl.contentRow.replace('{{HEIGHT}}', height)
          .replace('{{C1}}', () => cell(p.students, a)).replace('{{C2}}', () => cell(p.students, b))).join('');
      return tpl.pageXml
        .replace('{{PAGEBREAK}}', i > 0 ? '<w:pageBreakBefore/>' : '')
        .replace('{{TITRE}}', () => esc(titre))
        .replace('{{ENTETE}}', () => entete.join(''))
        .replace('{{ROWS}}', () => rows);
    }).join('');
    return tpl.docStart + body + tpl.sectPr + '</w:body></w:document>';
  }

  const api = {
    normalizeClass, categoryOf, categoryRank, tableLayout, excelDate, formatDateLong, parseRows,
    studentLine, sameFamily, conflicts, distribute, buildDocumentXml,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.Core = api;
})(typeof window !== 'undefined' ? window : globalThis);
