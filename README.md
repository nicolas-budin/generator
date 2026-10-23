# Couteau suisse de la Doyenne

Onglet « Heures d'arrêts » : génère le listing des élèves pour les enseignants.

## Utilisation
Double-cliquer sur `dist/Couteau-suisse.html` (s'ouvre dans Safari, fonctionne hors ligne) :
1. glisser le classeur `Heures d'arrêts ….xlsx` ;
2. choisir l'onglet (la date), le nombre de surveillant·es, leurs noms et salles ;
3. vérifier ou ajuster la répartition ;
4. « Générer le Word » : le .docx arrive dans Téléchargements.

Le fichier HTML peut être copié seul n'importe où (bureau, clé USB, OneDrive).

## Modifier la mise en page
Le Word est produit à partir de `modele.docx` (seule la première page est utilisée : titre,
ligne « Nom – salle … », tableau 7P…11VP, tableau absents/retards, « Merci de déposer … »).
Après modification du modèle :

```sh
npm install
node build.mjs            # régénère dist/Couteau-suisse.html
```

## Test sans navigateur
```sh
node build.mjs --debug    # écrit aussi dist-template.json
node test/run.mjs 28.10.2026 2 test/out.docx
```
