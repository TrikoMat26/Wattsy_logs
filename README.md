# Boîte à Outils d'Analyse de Logs ELISA & Traçabilité Production

Ce dossier contient les outils d'analyse pour les logs de production ELISA (`ELISA_Prod_Log_*.txt`) ainsi qu'un outil de traçabilité par lots.
Plusieurs outils PowerShell sont disponibles, ainsi que les anciens scripts individuels en **VBScript** (conservés pour référence).

---

## 🚀 _Gestion_Logs_Wattsy_Auto.ps1 (Automatisation & SFTP)

**Outil de synchronisation et de gestion.** Ce script permet de gérer les logs directement sur le Raspberry Pi de test et de les rapatrier sur le OneDrive local.

### Fonctionnalités Clés
- **Synchronisation SFTP** : Utilise WinSCP Portable pour télécharger les nouveaux logs du Pi vers le PC.
- **Double Vérification (MD5)** : Garantit l'intégrité des fichiers après téléchargement.
- **Gestion RTC** : Vérifie l'horloge du Pi par rapport au PC et permet de la synchroniser (mise à l'heure du module RTC DS3231).
- **Archivage Distant** : Permet de déplacer les logs traités dans des sous-dossiers sur le Pi pour garder le dossier principal propre.
- **Zéro Admin** : Conçu pour s'exécuter sans droits administrateur (WinSCP portable inclus).

### Utilisation rapide
1. Lancer le script (clic droit -> Exécuter avec PowerShell).
2. Utiliser le menu interactif (1 à 4) pour scanner, télécharger ou archiver.
3. Les journaux d'exécution sont stockés dans le dossier `_logs_exec`.

---

## 🏆 MasterLogTool.ps1 (Version PowerShell Recommandée)

**C'est l'outil principal à utiliser.** Il regroupe toutes les fonctionnalités des anciens scripts VBScript dans une interface graphique unique, avec des performances nettement supérieures.

### Avantages
- **Interface Graphique :** Plus besoin de lancer des scripts en ligne de commande, tout se fait via une fenêtre simple.
- **Performance :** Traitement ultra-rapide grâce au moteur .NET (10x à 50x plus rapide que VBS).
- **Compatibilité :** Supporte automatiquement les deux formats de logs rencontrés (Ancien format avec `#` et Nouveau format 2026 avec `:`).
- **Robustesse :** Écriture en UTF-8 (accents préservés), filtrage automatique des URLs parasites et nettoyage des caractères spéciaux.

### Fonctionnalités
L'interface propose 4 actions :
1.  **Extraction par Liste** : Extrait les logs complets pour les produits listés dans `NumSerieKO.txt`. Le fichier généré est nommé `SN_défaut1_défaut2.txt` pour identifier les pannes en un coup d'œil.
2.  **Inventaire Global** : Liste tous les numéros de série trouvés dans tous les logs.
3.  **Inventaire Validés (OK)** : Liste les succès (`[PROD_OK]`), analyse les doublons, génère une **liste globale confondus** de tous les SN OK uniques, puis effectue une **analyse par segments (lots/OF)** avec détection des numéros manquants. Propose d'ajouter les manquants dans `NumSerieKO.txt`.
4.  **Historique Complet** : Trace tout l'historique de chaque produit (OK, Erreur précise, ou Incomplet).
5.  **Gestion des OF** : Interface dédiée pour associer des numéros de série à des OF (7 chiffres). Permet de déclarer des plages de SN et sauvegarde les données pour usage ultérieur (`OF_Registry.json`).

### Utilisation-
1.  Faire un clic droit sur `MasterLogTool.ps1`.
2.  Choisir **"Exécuter avec PowerShell"**.
3.  Sélectionner l'action désirée et cliquer sur "EXÉCUTER LE SCRIPT".
4.  Les fichiers générés apparaissent dans la liste de droite et peuvent être ouverts directement.

---

## 📋 Liste_OF.ps1 (Segmentation par Lots — Script Autonome)

**Version standalone** de l'analyse de segmentation. Peut être utilisé indépendamment de MasterLogTool pour analyser une liste brute de numéros de série depuis un fichier texte.

> [!NOTE]
> Cette logique de segmentation est **aussi intégrée dans MasterLogTool.ps1** (option 3 — Inventaire Validés). Pour un usage normal, préférer MasterLogTool.

### Utilisation autonome
1. Placer `Liste_OF.txt` (liste brute de numéros) dans le même dossier que le script.
2. Lancer : `powershell.exe -ExecutionPolicy Bypass -File .\Liste_OF.ps1`
3. Le résultat est écrit dans `Liste_OF_traité.txt`.

### Configuration
- `$GapThreshold` (ligne 13) : Seuil d'écart pour couper un segment (défaut : 5).
- `$Exclude` (ligne 14) : Tableau de numéros à exclure (ex : `@("043246")`).

---

## 📂 Anciens Scripts (VBScript) - *Obsolètes mais fonctionnels*

Ces scripts individuels réalisent les mêmes tâches mais sont plus lents et moins pratiques. Ils sont conservés pour référence.

### 1. RechercheSerie.vbs
Extrait les blocs de logs pour les numéros de série présents dans `NumSerieKO.txt`.
*Sortie : Un fichier .txt par numéro de série.*

### 2. InventaireSeries.vbs
Liste tous les produits uniques trouvés fichier par fichier.
*Sortie : `Inventaire_Series.txt`*

### 3. InventaireSeriesOK.vbs
Liste les produits ayant réussi le test (`[PROD_OK]`) avec tri et détection des doublons.
*Sortie : `Inventaire_Series_OK.txt`*

### 4. HistoriqueTests.vbs
Génère un rapport complet de l'état de chaque test pour chaque produit.
*Sortie : `Historique_Tests_Complet.txt`*

---

## ⚙️ Standards Techniques (Pour Développeurs)

Pour toute maintenance ou modification, se référer impérativement au fichier : **`LLM_Instructions.md`**.

**Points Critiques :**
- **Encodage :** Lecture en ANSI, Écriture en **UTF-8 sans BOM** (accents préservés).
- **Nettoyage :** Suppression des octets nuls (`Chr(0)`) et trim strict des numéros de série.
- **Filtrage SN :** Seuls les numéros **purement numériques** sont acceptés (les URLs sont ignorées).
- **Format Datamatrix :** Les scripts d'analyse supportent les deux formats :
    - Ancien : `Datamatrix: #2025#SN`
    - Nouveau (2026+) : `Datamatrix: SN`
- **Compatibilité PS 5.1 :** Les scripts PowerShell doivent rester compatibles Windows PowerShell 5.1 (pas de `??`, conversions explicites `[int]`).
- **Segmentation OF :** Seuil d'écart configurable (`$GapThreshold`), préservation des zéros initiaux.
