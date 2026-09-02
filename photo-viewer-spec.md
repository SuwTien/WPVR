# Spec — Viewer photo portable (Excel → PowerShell → WPF)

## Contexte

Application Windows de visualisation/tri de photos, à usage personnel (extensible plus tard à d'autres utilisateurs). Utilisée sur un **PC tablette**, en **écran partagé** avec Excel (peu d'espace horizontal), la plupart du temps **sans clavier** (usage tactile prioritaire).

Contrainte forte : impossible d'installer quoi que ce soit sur la machine, et les `.exe` non signés/sans réputation risquent d'être supprimés par l'antivirus d'entreprise. → Aucune installation, aucun binaire compilé "maison".

## Architecture retenue

```
Excel (classeur macro déjà utilisé, bandeau personnalisé)
   └─ Bouton 4 (nouveau) → macro VBA
         └─ Shell("powershell.exe -File PhotoViewer.ps1", vbHide)
               └─ Script PowerShell charge une UI WPF (XAML)
                     └─ Fenêtre applicative (le "viewer")
```

**Pourquoi ce choix :** toute la chaîne repose sur des outils déjà présents et déjà autorisés sur la machine (Excel avec macros actives, PowerShell, .NET Framework/WPF natif à Windows). Aucun `.exe` n'est copié ou téléchargé — uniquement des fichiers texte (`.ps1`, `.xaml`, macro VBA dans le classeur). Cela contourne le problème de réputation/scan AV qui touche les exécutables compilés non signés.

WPF a été choisi plutôt qu'une interface web locale (PowerShell + serveur HTTP + navigateur) car :
- perfs natives équivalentes à un exe compilé (JIT .NET, rendu GPU)
- virtualisation de listes/grilles native (`VirtualizingStackPanel` / `WrapPanel` virtualisé) → scroll fluide même avec beaucoup de photos
- support tactile natif via les événements `Manipulation`/`Touch`
- moins de code à écrire pour un résultat équivalent

**Note technique — gestes tactiles avancés :** le câblage d'événements simples (clic) est direct en PowerShell pur, mais les événements `Manipulation`/`Touch` (swipe, pincer-zoomer) sont plus fastidieux à gérer en script PowerShell brut que dans du C# classique. Stratégie retenue : garder `PhotoViewer.ps1` comme unique fichier livré (toujours du texte, toujours lancé via `powershell.exe`), mais compiler à la volée, en mémoire, un bloc de code C# via `Add-Type` pour la partie gestion des gestes tactiles — ce qui permet d'écrire cette logique normalement (comme dans une vraie appli WPF) sans sacrifier l'architecture "aucun binaire distribué". À valider techniquement dès l'étape 1 (squelette), en testant les gestes de base avant de construire le reste dessus.

## Contraintes non-fonctionnelles

- **Tactile en priorité** : cibles de clic larges, gestes (tap, appui long, swipe, pincer-zoomer), le clavier physique reste optionnel partout.
- **Peu d'espace écran** : l'appli tourne en demi-écran (à côté d'Excel). Layout pensé pour une largeur réduite, adaptable portrait/paysage.
- **Fluidité maximale** : chargement de miniatures asynchrone + mise en cache, virtualisation de la grille, pas de blocage UI pendant le scroll ou le chargement d'image.
- **Pas d'installation, pas de binaire compilé** : uniquement scripts PowerShell + XAML + macro VBA existante.
- **Pas d'édition d'horodatage / métadonnées de prise de vue** : fonctionnalité explicitement exclue, à ne jamais développer (question d'intégrité des données).

## Fonctionnalités — v1 (viewer)

### Deux répertoires
- **Répertoire "existant"** : lecture seule (pas de renommage, pas de suppression).
- **Répertoire "fraîches"** : lecture + **renommage** + **suppression** autorisés.
- Pas d'affichage côte à côte des deux dossiers (espace écran trop limité) : **un bouton/switch pour basculer la vue entre les deux répertoires**. L'appli indique clairement quel répertoire est affiché, et grise/masque renommage + suppression quand on est sur "existant".

### Vue principale : grille de photos
- Grille de vignettes virtualisée (pas de liste verticale), remplit l'espace disponible en demi-écran, s'adapte portrait/paysage.
- Barre d'outils avec un bouton **"PHOTO"** (lance/ramène l'appli caméra au premier plan — voir section dédiée plus bas).
- **Tap simple sur une vignette → ouvre directement la popup de renommage** (le renommage est l'action prioritaire).
- **Appui long sur une vignette → mode plein écran** sur cette photo précise.

### Mode plein écran
- Ouverture sur la photo concernée (via appui long depuis la grille).
- Swipe gauche/droite (tactile) ou flèches clavier (si dispo) pour naviguer entre les photos du répertoire courant.
- Zoom : pincer pour zoomer (tactile) ou molette/raccourci (clavier), + déplacement (pan) dans l'image zoomée.
- **Nom de la photo affiché en bas de l'écran, cliquable** → ouvre la même popup de renommage que depuis la grille, sans quitter le plein écran.
- Sortie du plein écran par geste/bouton dédié (et touche Échap si clavier présent).

### Renommage — répertoire "fraîches" uniquement

**v1 — renommage unitaire rapide (fonction cœur, inspirée de l'appli Android "FPR" déjà utilisée et éprouvée par l'utilisateur) :**
- Tap sur une photo → popup avec le nom actuel pré-rempli, ou accès identique depuis le nom cliquable en plein écran.
- **Pavé numérique custom** intégré à l'appli (gros boutons tactiles 0-9, backspace, validation) — pas le clavier virtuel Windows.
- Bouton pour basculer vers un clavier texte complet (custom ou système) si des lettres sont nécessaires.
- **Petit bouton corbeille rouge en haut à droite de la popup** → supprime directement la photo depuis cet écran (avec confirmation, voir section Suppression).
- Objectif : renommage en quelques taps, sans lag, un enchaînement rapide photo par photo.

**Pas de sélection multiple ni de renommage par lot en v1** — Windows Explorer couvre déjà le renommage groupé basique (sélection multiple + F2) si besoin ponctuel. Simple recommandation d'implémentation pour rester compatible si l'idée revient un jour : représenter chaque photo de la grille par son propre petit objet (plutôt qu'un simple chemin de fichier) et centraliser la gestion du tap dans un seul point de code. Ce n'est pas une fonctionnalité à construire, juste une habitude de code raisonnable qui coûte rien et évite une réécriture si le besoin apparaît plus tard.

### Suppression — répertoire "fraîches" uniquement
- **Suppression unitaire uniquement**, via le bouton corbeille de la popup de renommage (photo par photo).
- Confirmation obligatoire avant suppression (éviter les suppressions accidentelles au doigt).
- Pas de suppression multiple dans l'appli : pour un nettoyage en masse, l'utilisateur passe par l'explorateur Windows (le dossier est un dossier normal, accessible sans contrainte).

### Détection des nouvelles photos
- **Détection automatique** : surveillance continue du dossier "fraîches" via `FileSystemWatcher` (natif .NET) — toute nouvelle photo (prise avec l'appareil, transférée depuis un téléphone, etc.) apparaît automatiquement dans la grille, sans action de l'utilisateur.
- **Bouton "rafraîchir" manuel en secours** : au cas où la détection automatique manquerait un cas (ex. dossier réseau, latence du système de fichiers).

### Intégration caméra
- Bouton dédié dans l'appli : si l'appli caméra Windows est déjà lancée, la ramène au premier plan ; sinon, la lance. Permet de rester dans le flux sans chercher l'icône ailleurs, tout en gardant le viewer ouvert (écran partagé) ou en y revenant facilement ensuite.
- Combiné à la détection automatique ci-dessus : la photo prise apparaît dans la grille dès le retour sur le viewer, sans manipulation supplémentaire.

## Fonctionnalité — v1, chantier séparé (macro VBA, indépendante du viewer)

### Export "planche contact" Excel
- Nouveau bouton dans le bandeau Excel (à côté de celui qui lance le viewer).
- Scanne le répertoire "photos fraîches".
- Crée un nouvel onglet Excel avec, par photo, une ligne :
  - Colonne A : nom du fichier
  - Colonne B : image insérée/ancrée dans la cellule (redimensionnée)
  - Colonne C : horodatage (date de prise de vue EXIF si dispo, sinon date de fichier)
- Usage réel : constitution de rapports illustrés (photos + remarques/commentaires), mis en page ensuite pour impression PDF (hors périmètre de ce projet).
- Remplace le processus actuel manuel "photo par photo".
- Techniquement indépendant du viewer WPF (pas de dépendance croisée) — peut être développé et testé séparément. Réutilise la logique VBA déjà existante dans le classeur de l'utilisateur.

## Hors scope v1 / Backlog
- Affichage côte à côte des deux répertoires (abandonné au profit du switch, faute d'espace écran).
- **Édition d'horodatage / métadonnées EXIF de prise de vue : exclu définitivement, pas seulement reporté.**

## Notes pour la session de développement (Claude Code / VS Code)

- Livrables attendus : un script `.ps1` (point d'entrée), un ou plusieurs `.xaml` (UI), et un extrait de macro VBA à intégrer dans le classeur existant de l'utilisateur (bouton de bandeau).
- Priorité de développement suggérée :
  1. Squelette WPF + grille virtualisée sur un seul répertoire + tap → popup de renommage de base.
  2. Switch entre les deux répertoires + règle "lecture seule" sur le répertoire existant.
  3. Renommage unitaire rapide (popup + pavé numérique custom + bascule clavier texte + bouton corbeille).
  4. Suppression unitaire avec confirmation (depuis la popup de renommage).
  5. Mode plein écran (appui long, swipe, zoom/pan, nom cliquable en bas).
  6. Détection auto des nouvelles photos (`FileSystemWatcher`) + bouton refresh manuel.
  7. Bouton caméra (premier plan / lancement).
  8. Intégration du lancement depuis Excel (macro VBA + bouton bandeau).
  9. (Chantier séparé) macro d'export planche contact Excel.
- Tester dès le début le lancement via Excel → PowerShell → WPF sur la machine cible réelle (contexte tactile, écran partagé, restrictions antivirus), pas seulement en environnement de dev classique.
- Garder la fluidité comme critère de validation à chaque étape (pas d'attente ni de saccade, en particulier lors du scroll de la grille et du chargement des grandes images).
