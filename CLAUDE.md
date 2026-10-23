# Couteau suisse de la Doyenne


Outils de génération de documents pour le décanat d'un établissement scolaire vaudois
(utilisatrice : Florence, la Doyenne). Une seule page HTML autonome, organisée en onglets,
chaque onglet étant un outil qui lit son propre fichier Excel et produit un document Word.

## Contraintes

- **Mac / Safari**, ouverture par double-clic en `file://`. Rien à installer, aucun accès réseau.
- **Un seul fichier livré** : `dist/Couteau-suisse.html`. Librairies, modèles Word et image y sont intégrés.
- **Données personnelles d'élèves** (noms, adresses des parents, motifs) : elles ne doivent jamais quitter le
  navigateur, donc pas de CDN, pas de télémétrie, pas d'upload.
- Interface et messages **en français**, écriture inclusive avec le point médian (surveillant·es).
- JS vanilla sans framework ni bundler. Node ne sert qu'au build et aux tests.

## Structure
```
src/index.html   HTML + CSS ; marqueurs /*__LIBS__*/ /*__CORE__*/ /*__TEMPLATE__*/ /*__APP__*/ et __ACCUEIL_IMG__
src/core.js      logique pure, sans DOM (window.Core dans le navigateur, module.exports dans Node)
src/app.js       interface : état, rendu par innerHTML, délégation d'événements, onglets
src/accueil.jpg  photo de la page d'accueil (JPEG ~1400 px, qualité ~50)
modele.docx      Word d'origine servant de modèle pour l'onglet « Heures d'arrêts »
build.mjs        extrait le modèle Word et assemble dist/Couteau-suisse.html
test/run.mjs     test de bout en bout sans navigateur
```

## Commandes
```sh
npm install
node build.mjs                                    # → dist/Couteau-suisse.html
node build.mjs --debug                            # écrit aussi dist-template.json (requis par test/run.mjs)
node test/run.mjs 28.10.2026 2 test/out.docx      # <onglet Excel> <nb surveillant·es> <sortie>
```
Toujours relancer `node build.mjs` après une modification de `src/` ou de `modele.docx`.
Pas de navigateur headless : vérifier la syntaxe des scripts intégrés avec `new Function(...)` et
l'aperçu du .docx avec `qlmanage -t -s 1400 -o <dir> fichier.docx`. Quick Look n'affiche ni les bordures,
ni le surlignage, ni les sauts de page.

## Génération Word (principe)

Pas de librairie docx. On réutilise `modele.docx` et on remplace seulement `word/document.xml` :

- au build, `extractTemplate()` garde la 1re page, puis place des marqueurs (`{{TITRE}}`, `{{PAGEBREAK}}`,
  `{{ENTETE}}`, `{{ROWS}}`) et extrait les lignes modèles du tableau (`headerRow` avec `{{H1}}`/`{{H2}}`,
  `contentRow` avec `{{C1}}`/`{{C2}}`/`{{HEIGHT}}`) ainsi que les `pPr`/`rPr` d'une ligne élève ;
- dans le navigateur, `Core.buildDocumentXml()` produit une page par surveillant·e, puis JSZip réinjecte
  le XML dans le .docx encodé en base64 et le fichier est téléchargé via `<a download>` ;
- le XML est manipulé par chaînes et regex : toujours échapper le texte (`esc`) et retirer les
  `w14:paraId` des parties dupliquées.

## Onglet « Heures d'arrêts »

- Source : classeur avec **un onglet Excel par date** (« 28.10.2026 »). En-têtes en ligne 1. Colonnes
  repérées par leur nom, avec repli sur les lettres C nom, D prénom, E classe, F périodes, H date (numéro de
  série Excel) et N adresse des parents.
- Classes normalisées (`10VP/2` → `10VP2`). Une case par niveau : 7P|8P … 11VG|11VP toujours présentes,
  1P–6P et « Autres » ajoutées si besoin, `7-8 EI` → 7P.
- Répartition (`Core.distribute`) : glouton équilibré, puis échanges. On sépare d'abord les fratries
  (même nom **ou** même adresse des parents), puis les élèves d'une même classe.
- `localStorage` (clés `ha.*`) : surveillant·es et salles uniquement, accès protégés par try/catch.

## Ajouter un outil (onglet)

Un bouton dans `.tabs`, un `.tab-panel[data-tab]` et une carte sur l'accueil (`data-goto`).
`app.js` cible encore directement les identifiants de l'onglet « Heures d'arrêts ». Pour le prochain outil,
découper d'abord en `src/tabs/<outil>.js` (chacun limité à son panneau, avec ses propres identifiants
et son propre modèle Word) et ajouter les marqueurs correspondants dans `build.mjs`.

## Fichiers d'exemple (racine du projet)

Ce sont de **vraies données d'élèves**. Ne pas les diffuser ni les intégrer au HTML.

- `Heures d'arrêts 2026-2027.xlsx` : le classeur source (un onglet par date) ; le `.pdf` du même nom en est un export.
- `Liste des élèves pour les enseignants-2810.docx` / `.pdf` : le Word réalisé à la main pour le 28.10,
  **référence visuelle** à reproduire. `modele.docx` en est une copie.

## Particularités de l'Excel

- Colonne A : en général la date de convocation, parfois « Reconvoc du 07.10.2026/ motif: pas venu ». Elle n'est pas utilisée.
- Les onglets des dates futures existent mais sont vides : afficher un message, ne pas générer de Word.
- Cellules mal nettoyées : espaces en début ou en fin (`' Yermani '`), `\n` et `\t` dans les localités et
  les noms, casse variable des classes (`7p2`, `7-8 ei`).

## Résultats de référence (non-régression)

- `node test/run.mjs 28.10.2026 2` : 26 élèves, **13 / 13**, 0 conflit de famille, **4 paires de même classe**
  (minimum possible : 4× 10VP2, 3× 9VG1, 3× 9VG2). Nduakulu et Islamagic séparés, ainsi que les élèves
  d'un même foyer (Feuillère, Clochatte).
- `node test/run.mjs 07.10.2026 2` : Aloise Lorenzo (6P2) → une ligne **5P | 6P** est ajoutée au tableau.

## Décisions validées (ne pas remettre en cause sans demander)

- Nom exact : « Couteau suisse de la Doyenne » (D majuscule). Onglets : Accueil, Heures d'arrêts, Prochainement :).
- L'outil s'ouvre **toujours sur l'accueil** (photo et une carte par outil). Le dernier onglet n'est pas mémorisé.
- **Chaque onglet charge son propre fichier** : rien n'est partagé entre onglets (Excel, messages compris).
- Dans le Word, le tableau est **le même sur toutes les pages** (lignes calculées sur l'ensemble des élèves
  de l'onglet), et la hauteur totale du tableau est conservée quand des lignes sont ajoutées.
- Dans chaque case, les élèves sont triés par classe puis par nom ; seule la salle de la personne est surlignée.
- Après chaque modification : `node build.mjs`, puis l'utilisateur recharge la page avec Cmd+R. Si le fichier a été copié ailleurs, il faut le recopier.

## Points ouverts

- `docxBase64` contient le `document.xml` d'origine du Word du 28.10, avec de vrais noms d'élèves.
  Il faudrait le vider au build avant de diffuser le HTML (en attente de validation).
