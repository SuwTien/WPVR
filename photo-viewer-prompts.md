# Prompts pour la session Claude Code — Viewer photo

## Comment utiliser ce fichier

1. Ouvre une nouvelle session Claude Code dans VS Code, dans un dossier de projet vide.
2. Donne d'abord `photo-viewer-spec.md` (colle-le entièrement ou mets-le dans le dossier et référence-le) pour que Claude Code ait tout le contexte.
3. Envoie ensuite les prompts ci-dessous **un par un, dans l'ordre**. Teste sur ta machine réelle (tactile, écran partagé) après chaque étape avant de passer à la suivante — ne pas enchaîner à l'aveugle.
4. Chaque prompt rappelle le contexte essentiel pour rester utilisable même isolément si besoin.

---

## Prompt 0 — Cadrage initial du projet

```
Je veux construire une application Windows de visualisation de photos, en PowerShell + WPF (XAML), sans aucun binaire compilé, sans installation. Voici le fichier de spec complet : [colle photo-viewer-spec.md ici, ou réfère-toi au fichier dans le dossier].

Avant de coder quoi que ce soit, propose-moi une structure de fichiers/dossiers pour le projet (scripts .ps1, fichiers .xaml séparés ou XAML inline, organisation du code C#-like en PowerShell si besoin de classes). Explique brièvement tes choix. N'écris pas encore de code fonctionnel, juste la structure et un plan de fichiers.

Contraintes à garder en tête pour tout le projet :
- Aucun exe compilé, aucune installation, tout doit rester des fichiers texte (.ps1, .xaml).
- Usage tactile prioritaire (peu ou pas de clavier physique), écran partagé en demi-largeur avec Excel.
- Fluidité maximale : chargement de miniatures asynchrone, pas de blocage de l'UI.
- On avance étape par étape (voir plan de dev dans la spec) — ne développe que ce qui est demandé dans chaque prompt, sans anticiper les étapes suivantes.
```

---

## Prompt 1 — Squelette WPF + grille virtualisée + tap basique

```
Étape 1 du plan : squelette de l'application.

Objectif :
- Une fenêtre WPF qui s'ouvre sur un seul répertoire de photos (chemin en dur pour l'instant, on gérera le switch de répertoire à l'étape suivante).
- Une grille de vignettes virtualisée (WrapPanel virtualisé, ou ItemsControl avec VirtualizingPanel) qui affiche les photos du dossier.
- Chargement des miniatures de façon asynchrone (pas de gel de l'UI pendant le chargement, même avec beaucoup de photos).
- Tap simple sur une vignette → pour l'instant, ouvre juste une popup vide (placeholder) avec le nom du fichier affiché. On implémentera le vrai contenu du renommage à l'étape 3.
- Le layout doit s'adapter à une fenêtre étroite (simulateur d'écran partagé en demi-largeur), portrait et paysage.

Pas encore : switch de répertoire, renommage réel, suppression, plein écran, caméra, détection auto, intégration Excel. On garde uniquement ce périmètre pour cette étape.

Point technique à valider dès cette étape : pour la gestion des événements tactiles (clic simple pour l'instant, mais aussi en préparation des gestes plus avancés des étapes suivantes comme le swipe et le pincer-zoomer), utilise si besoin `Add-Type` pour compiler à la volée un bloc de code C# embarqué dans le script plutôt que de tout gérer en PowerShell pur — ça reste un seul fichier .ps1 livré, sans aucun binaire distribué, mais ça facilite grandement le câblage des événements `Manipulation`/`Touch`. Fais un petit test de geste tactile basique (tap) dès cette étape pour valider que l'approche choisie fonctionne bien sur la machine cible avant de construire le reste dessus.

Teste avec un dossier contenant au moins 50-100 photos pour valider que le scroll reste fluide.
```

---

## Prompt 2 — Switch entre les deux répertoires + lecture seule

```
Étape 2 du plan : gestion des deux répertoires.

Objectif :
- Deux chemins de dossiers configurables : "existant" (lecture seule) et "fraîches" (lecture + écriture, où le renommage/suppression seront permis plus tard).
- Un bouton/switch dans l'interface pour basculer l'affichage de la grille entre les deux répertoires.
- Indication visuelle claire de quel répertoire est actuellement affiché.
- Prépare (sans forcément l'implémenter en dur maintenant) un mécanisme simple pour savoir, dans le reste du code, si le répertoire actif autorise le renommage/suppression ou non (ex. une propriété booléenne exposée quelque part) — les actions de renommage/suppression réelles seront branchées dans les étapes suivantes, mais elles devront pouvoir vérifier cette info facilement.

Garde le reste du périmètre de l'étape 1 fonctionnel (grille virtualisée, tap → popup placeholder).
```

---

## Prompt 3 — Renommage unitaire rapide

```
Étape 3 du plan : renommage unitaire, la fonction cœur de l'appli.

Objectif :
- Sur tap d'une vignette dans le répertoire "fraîches" uniquement (pas d'action de renommage possible sur "existant") : ouverture d'une popup de renommage.
- La popup affiche le nom actuel du fichier, pré-rempli et éditable.
- Pavé numérique custom intégré (gros boutons tactiles 0 à 9, backspace, validation) — ne pas utiliser le clavier virtuel Windows.
- Un bouton dans la popup permet de basculer vers un clavier texte complet (custom ou système) pour saisir des lettres si besoin.
- Un petit bouton corbeille rouge en haut à droite de la popup — pour l'instant, il peut juste déclencher une popup de confirmation vide (placeholder), la vraie suppression sera codée à l'étape 4.
- Validation du renommage → renomme réellement le fichier sur le disque, ferme la popup, la grille se met à jour.
- Sur le répertoire "existant" : le tap ne doit PAS ouvrir cette popup de renommage (comportement à définir — par exemple juste afficher l'image en grand, ou ne rien faire de spécial, à toi de proposer quelque chose de cohérent).

Priorité : cet enchaînement doit être ultra rapide et fluide au tactile — pas de latence perceptible entre le tap et l'ouverture de la popup, ni entre la validation et la mise à jour de la grille.
```

---

## Prompt 4 — Suppression unitaire avec confirmation

```
Étape 4 du plan : suppression réelle.

Objectif :
- Le bouton corbeille de la popup de renommage (répertoire "fraîches" uniquement) déclenche une vraie confirmation ("Supprimer [nom du fichier] ?"), puis supprime réellement le fichier du disque si confirmé.
- Après suppression, la popup se ferme et la grille se met à jour (la photo disparaît).
- Pas de suppression multiple à prévoir — uniquement photo par photo, depuis cette popup.
```

---

## Prompt 5 — Mode plein écran

```
Étape 5 du plan : visualisation plein écran.

Objectif :
- Appui long sur une vignette de la grille → ouvre la photo en plein écran.
- Navigation entre les photos du répertoire actuellement affiché par swipe gauche/droite (tactile) et flèches clavier si un clavier est présent.
- Zoom par pincement (tactile) et molette/raccourci (clavier), avec déplacement (pan) possible dans l'image zoomée.
- Le nom du fichier est affiché en bas de l'écran plein écran, et il est cliquable : un tap dessus ouvre la même popup de renommage que dans la grille (celle de l'étape 3), sans quitter le mode plein écran une fois la popup fermée.
- Un moyen clair de sortir du plein écran (geste ou bouton dédié, + touche Échap si clavier disponible).

Réutilise la popup de renommage déjà construite à l'étape 3 plutôt que d'en recréer une nouvelle.
```

---

## Prompt 6 — Détection automatique des nouvelles photos

```
Étape 6 du plan : détection des nouvelles photos.

Objectif :
- Mets en place une surveillance continue du répertoire "fraîches" (FileSystemWatcher ou équivalent natif .NET) : toute nouvelle photo ajoutée dans ce dossier doit apparaître automatiquement dans la grille, sans action de l'utilisateur, et sans geler l'interface.
- Ajoute également un bouton "rafraîchir" manuel visible dans l'interface, en secours, qui recharge le contenu du répertoire actif à la demande.
- Vérifie que ça fonctionne correctement même si l'application reste ouverte longtemps et que plusieurs photos arrivent d'affilée.
```

---

## Prompt 7 — Bouton caméra

```
Étape 7 du plan : intégration de la caméra.

Objectif :
- Ajoute un bouton "PHOTO" dans la barre d'outils de l'application.
- Au clic : si l'appli caméra Windows est déjà en cours d'exécution, ramène sa fenêtre au premier plan ; sinon, lance l'appli caméra Windows.
- Pas besoin de gérer le retour vers le viewer explicitement — la détection automatique de l'étape 6 doit suffire pour que la nouvelle photo apparaisse quand l'utilisateur revient sur le viewer.
```

---

## Prompt 8 — Intégration avec Excel (lancement)

```
Étape 8 du plan : lancement depuis Excel.

Objectif :
- Rédige le code VBA à ajouter dans le classeur Excel existant de l'utilisateur (qui a déjà un bandeau personnalisé avec 3 boutons) pour ajouter un 4e bouton.
- Ce bouton doit lancer le script PowerShell de l'application en tâche de fond, sans afficher de fenêtre de console PowerShell (Shell caché / vbHide ou équivalent), en pointant vers le chemin du script principal (.ps1) créé dans les étapes précédentes.
- Documente clairement, en commentaire dans le code VBA, les étapes pour rattacher ce code à un bouton du bandeau personnalisé existant dans Excel.
- Précise aussi, en commentaire ou dans une note à part, les points à vérifier si le lancement ne fonctionne pas (politique de sécurité macro, chemin du script, restrictions PowerShell type ExecutionPolicy) — pense à proposer une commande d'exécution qui contourne une ExecutionPolicy restrictive sans nécessiter de droits admin (ex. `-ExecutionPolicy Bypass` au lancement du process, qui ne change la policy que pour ce process et pas globalement sur la machine).
```

---

## Prompt 9 — Chantier séparé : export planche contact Excel

```
Chantier indépendant du viewer WPF (pas de dépendance avec les étapes précédentes) : macro VBA d'export planche contact.

Objectif :
- Bouton dans le bandeau Excel qui scanne le répertoire "photos fraîches".
- Crée un **nouveau classeur Excel vierge et autonome** (Workbooks.Add), sans modifier le classeur macro existant, puis l'enregistre (SaveAs) sous un nom à définir (ex. horodaté).
- Dans ce nouveau classeur, pour chaque photo trouvée, une ligne :
  - Colonne A : nom du fichier
  - Colonne B : l'image elle-même, insérée via Shapes.AddPicture, redimensionnée pour rester lisible sans être trop lourde en mémoire
  - Colonne C : horodatage (date de prise de vue EXIF si disponible, sinon date de modification du fichier)
- Doit gérer proprement le cas où le fichier de sortie existe déjà (ex. proposer un nom horodaté unique, ou demander confirmation avant d'écraser).
- Le but final est que l'utilisateur puisse ensuite copier manuellement cette feuille dans un autre classeur si besoin (pas à automatiser ici) — pas besoin de gérer la mise en page/impression dans ce chantier, juste la génération du fichier et des données/images dedans.
```
