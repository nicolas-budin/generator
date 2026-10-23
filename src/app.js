// Interface : chargement de l'Excel, saisie des surveillant·es, répartition, téléchargement du Word.
(function () {
  'use strict';
  const $ = (sel) => document.querySelector(sel);
  const EXCLU = -1;

  const store = {
    get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* navigation privée */ } },
  };

  const state = {
    wb: null,
    sheet: null,
    students: [],
    date: null,
    salles: store.get('ha.salles', 'B121 / B214 / B220'),
    teachers: store.get('ha.teachers', [{ name: '', salle: 'B121' }, { name: '', salle: 'B214' }]),
    assign: {},
  };

  const sallesList = () => state.salles.split('/').map((s) => s.trim()).filter(Boolean);
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  function message(text, kind = 'info') {
    const el = $('#message');
    el.textContent = text;
    el.className = 'message ' + kind;
    el.hidden = !text;
  }

  // ---------- Fichier Excel ----------

  function loadFile(file) {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        state.wb = XLSX.read(new Uint8Array(reader.result), { type: 'array' });
      } catch (e) {
        message('Impossible de lire ce fichier. Est-ce bien un fichier Excel (.xlsx) ?', 'error');
        return;
      }
      $('#filename').textContent = file.name;
      fillSheets();
    };
    reader.readAsArrayBuffer(file);
  }

  function fillSheets() {
    const names = state.wb.SheetNames;
    const today = new Date(); today.setHours(0, 0, 0, 0);
    // Présélection : la prochaine date à venir (nom de feuille « JJ.MM.AAAA »), sinon la dernière.
    let pick = names[names.length - 1], bestDiff = Infinity;
    for (const n of names) {
      const d = Core.excelDate(n);
      if (!d) continue;
      const diff = d.getTime() - Date.UTC(today.getFullYear(), today.getMonth(), today.getDate());
      if (diff >= 0 && diff < bestDiff) { bestDiff = diff; pick = n; }
    }
    $('#sheet').innerHTML = names.map((n) => `<option${n === pick ? ' selected' : ''}>${esc(n)}</option>`).join('');
    $('#step-sheet').hidden = false;
    selectSheet(pick);
  }

  function selectSheet(name) {
    state.sheet = name;
    const rows = XLSX.utils.sheet_to_json(state.wb.Sheets[name], { header: 1, raw: true, defval: '' });
    const { students, date } = Core.parseRows(rows, name);
    state.students = students;
    state.date = date;
    if (!students.length) {
      message(`La feuille « ${name} » ne contient aucun élève.`, 'warn');
      $('#step-teachers').hidden = true;
      $('#step-result').hidden = true;
      return;
    }
    message('');
    $('#sheet-info').textContent = `${students.length} élève${students.length > 1 ? 's' : ''}` +
      (date ? ` · ${Core.formatDateLong(date)}` : ' · date introuvable');
    $('#step-teachers').hidden = false;
    $('#step-result').hidden = false;
    redistribute();
  }

  // ---------- Surveillant·es ----------

  function renderTeachers() {
    const salles = sallesList();
    $('#count').value = state.teachers.length;
    if (document.activeElement !== $('#salles')) $('#salles').value = state.salles;
    $('#teachers').innerHTML = state.teachers.map((t, i) => `
      <div class="teacher-row">
        <span class="num">${i + 1}</span>
        <input type="text" data-i="${i}" class="t-name" placeholder="Prénom Nom" value="${esc(t.name)}">
        <select data-i="${i}" class="t-salle">
          ${salles.map((s) => `<option${s === t.salle ? ' selected' : ''}>${esc(s)}</option>`).join('')}
          ${t.salle && !salles.includes(t.salle) ? `<option selected>${esc(t.salle)}</option>` : ''}
        </select>
      </div>`).join('');
  }

  function setCount(n) {
    n = Math.max(1, Math.min(8, n | 0));
    const salles = sallesList();
    while (state.teachers.length < n) {
      state.teachers.push({ name: '', salle: salles[state.teachers.length] || salles[0] || '' });
    }
    state.teachers.length = n;
    saveTeachers();
    renderTeachers();
    redistribute();
  }

  function saveTeachers() {
    store.set('ha.teachers', state.teachers);
    store.set('ha.salles', state.salles);
  }

  // ---------- Répartition ----------

  function redistribute() {
    if (!state.students.length) return;
    state.assign = Core.distribute(state.students, state.teachers.length);
    renderResult();
  }

  function renderResult() {
    const n = state.teachers.length;
    const active = state.students.filter((s) => state.assign[s.id] !== EXCLU);
    const conf = Core.conflicts(active, state.assign, n);
    const cols = state.teachers.map((t, i) => ({ title: t.name || `Surveillant·e ${i + 1}`, idx: i }));
    if (state.students.some((s) => state.assign[s.id] === EXCLU)) cols.push({ title: 'Exclus du listing', idx: EXCLU });

    const options = (cur) => state.teachers.map((t, i) =>
      `<option value="${i}"${i === cur ? ' selected' : ''}>→ ${esc(t.name || 'Surveillant·e ' + (i + 1))}</option>`).join('') +
      `<option value="${EXCLU}"${cur === EXCLU ? ' selected' : ''}>Exclure</option>`;

    $('#columns').innerHTML = cols.map((col) => {
      const list = state.students.filter((s) => state.assign[s.id] === col.idx)
        .sort((a, b) => Core.categoryRank(a.cat) - Core.categoryRank(b.cat) ||
          a.classe.localeCompare(b.classe, 'fr', { numeric: true }) || a.nom.localeCompare(b.nom, 'fr'));
      const periodes = list.reduce((x, s) => x + s.periodes, 0);
      return `<section class="col${col.idx === EXCLU ? ' excl' : ''}">
        <header><h3>${esc(col.title)}</h3><span>${list.length} élève${list.length > 1 ? 's' : ''} · ${periodes} pér.</span></header>
        <ul>${list.map((s) => {
          const c = conf.byStudent[s.id];
          const tags = [];
          if (c && c.famille.length) tags.push('<span class="tag fam" title="Même nom ou même adresse qu\'un·e autre élève de cette liste">famille</span>');
          if (c && c.classe.length) tags.push('<span class="tag cls" title="Même classe qu\'un·e autre élève de cette liste">même classe</span>');
          if (!s.catSure) tags.push(`<span class="tag warn" title="Classe non reconnue, placée dans la case ${s.cat}">case ${s.cat} ?</span>`);
          return `<li><div class="line">${esc(Core.studentLine(s))} ${tags.join('')}</div>
            <select data-id="${s.id}" class="move">${options(col.idx)}</select></li>`;
        }).join('')}</ul></section>`;
    }).join('');

    const parts = [];
    if (conf.famille) parts.push(`${conf.famille} conflit${conf.famille > 1 ? 's' : ''} de famille`);
    if (conf.classe) parts.push(`${conf.classe} paire${conf.classe > 1 ? 's' : ''} d'élèves de la même classe`);
    $('#summary').textContent = parts.length
      ? `Inévitable ou à corriger à la main : ${parts.join(', ')}.`
      : 'Aucun frère/sœur ni élève de la même classe chez la même personne.';
    $('#summary').className = parts.length ? 'summary warn' : 'summary ok';
  }

  // ---------- Word ----------

  async function generate() {
    const missing = state.teachers.findIndex((t) => !t.name.trim());
    if (missing >= 0) {
      message(`Indiquer le nom de la personne n° ${missing + 1}.`, 'error');
      $(`.t-name[data-i="${missing}"]`).focus();
      return;
    }
    if (!state.date) { message('Date des heures d\'arrêts introuvable dans la feuille.', 'error'); return; }
    message('');
    const pages = state.teachers.map((t, i) => ({
      enseignant: t.name.trim(),
      salle: t.salle,
      students: state.students.filter((s) => state.assign[s.id] === i),
    }));
    const xml = Core.buildDocumentXml(window.TEMPLATE, { date: state.date, salles: sallesList(), pages });
    const zip = await JSZip.loadAsync(window.TEMPLATE.docxBase64, { base64: true });
    zip.file('word/document.xml', xml);
    const blob = await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      compression: 'DEFLATE',
    });
    const d = state.date;
    const jjmm = String(d.getUTCDate()).padStart(2, '0') + String(d.getUTCMonth() + 1).padStart(2, '0');
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `Liste des élèves pour les enseignants-${jjmm}.docx`;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
    message(`Fichier « ${a.download} » enregistré dans Téléchargements.`, 'ok');
  }

  // ---------- Événements ----------

  const drop = $('#drop');
  drop.addEventListener('click', () => $('#file').click());
  $('#file').addEventListener('change', (e) => {
    if (e.target.files[0]) loadFile(e.target.files[0]);
    e.target.value = ''; // permet de recharger le même fichier après modification
  });
  ['dragenter', 'dragover'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', (e) => e.dataTransfer.files[0] && loadFile(e.dataTransfer.files[0]));

  $('#sheet').addEventListener('change', (e) => selectSheet(e.target.value));
  $('#count').addEventListener('change', (e) => setCount(+e.target.value));
  $('#minus').addEventListener('click', () => setCount(state.teachers.length - 1));
  $('#plus').addEventListener('click', () => setCount(state.teachers.length + 1));
  $('#salles').addEventListener('input', (e) => {
    state.salles = e.target.value;
    const salles = sallesList();
    // Une salle retirée de la liste est remplacée par celle de même position.
    state.teachers.forEach((t, i) => { if (!salles.includes(t.salle)) t.salle = salles[i] || salles[0] || ''; });
    saveTeachers();
    renderTeachers();
  });
  $('#teachers').addEventListener('input', (e) => {
    const i = +e.target.dataset.i;
    if (e.target.classList.contains('t-name')) state.teachers[i].name = e.target.value;
    if (e.target.classList.contains('t-salle')) state.teachers[i].salle = e.target.value;
    saveTeachers();
    if (state.students.length) renderResult();
  });
  $('#teachers').addEventListener('change', (e) => {
    if (e.target.classList.contains('t-salle')) { state.teachers[+e.target.dataset.i].salle = e.target.value; saveTeachers(); }
  });
  $('#columns').addEventListener('change', (e) => {
    if (!e.target.classList.contains('move')) return;
    state.assign[+e.target.dataset.id] = +e.target.value;
    renderResult();
  });
  $('#auto').addEventListener('click', redistribute);
  $('#generate').addEventListener('click', () => generate().catch((err) => {
    console.error(err);
    message('Erreur pendant la création du Word : ' + err.message, 'error');
  }));

  // Onglets : un par document à générer.
  function showTab(name) {
    const tabs = [...document.querySelectorAll('.tabs [data-tab]')];
    if (!tabs.some((t) => t.dataset.tab === name)) name = tabs[0].dataset.tab;
    tabs.forEach((t) => t.setAttribute('aria-selected', String(t.dataset.tab === name)));
    document.querySelectorAll('.tab-panel').forEach((p) => { p.hidden = p.dataset.tab !== name; });
    document.body.dataset.tab = name;
    window.scrollTo(0, 0);
  }
  $('.tabs').addEventListener('click', (e) => { if (e.target.dataset.tab) showTab(e.target.dataset.tab); });
  document.addEventListener('click', (e) => {
    const goto = e.target.closest('[data-goto]');
    if (goto) showTab(goto.dataset.goto);
  });
  showTab('accueil');

  renderTeachers();
})();
