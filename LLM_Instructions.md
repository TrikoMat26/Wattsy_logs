# Directives de Projet : Analyseur de Logs ELISA & Outils de Traçabilité

Ce fichier définit les standards techniques et l'historique des résolutions pour garantir la performance et la stabilité des outils de traitement de logs (VBScript & PowerShell) et de traçabilité de production.

## 1. Encodage et Flux de Données (CRITIQUE)
- **Standard Global** : Toujours utiliser le format **ANSI (Windows-1252)** pour la lecture et l'écriture.
- **En VBScript** :
    - Lecture : `objFSO.OpenTextFile(path, 1, False, 0)` (TristateFalse)
    - Écriture : `objFSO.CreateTextFile(path, True, False)`
- **En PowerShell (Recommandé)** :
    - Lecture : `$EncANSI = [System.Text.Encoding]::GetEncoding(1252)` puis `[System.IO.File]::ReadLines($path, $EncANSI)`
    - Écriture : **UTF-8 sans BOM** pour préserver les accents :
      ```powershell
      $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
      $sw = New-Object System.IO.StreamWriter($path, $false, $utf8NoBom)
      ```

**Problèmes d'encodage connus :**
- L'utilisation de l'Unicode (UTF-16) provoque des "caractères chinois" (mojibake).
- L'écriture en ANSI échoue si un caractère > 255 est présent.
- **Solution** : Utiliser impérativement une fonction **`Sanitize-Text`** (PS) ou `Sanitize` (VBS) pour filtrer les caractères hors plage ANSI avant écriture.

## 2. Logique d'Extraction Multi-Formats (CRITIQUE)
Le script doit supporter **simultanément** les deux formats de logs existants :
1.  **Ancien Format** : `... Datamatrix: #2025#13104 ...` (SN après `Datamatrix: #...#`)
2.  **Nouveau Format (2026+)** : `... Datamatrix: 043355 ===` (SN après `Datamatrix: `)

**Implémentation requise (Fonction `Extract-SN`)** :
- Utiliser une Regex pour détecter le format.
- **IMPORTANT** : N'accepter que des **chiffres** (`\d+`) pour le SN. Cela exclut automatiquement les URLs parasites (ex: `https://...`).
- Nettoyer le résultat retourné avec `Clean-Serial`.

## 3. Performance et Architecture
- **Algorithme** : $O(M)$ par fichier.
    - Charger les données en mémoire vive (Dictionnaires/HashSets).
    - Parcourir chaque fichier log **une seule fois** de haut en bas.
    - Écrire le rapport final en une seule opération disque.
- **PowerShell vs VBS** : Préférer l'utilisation des classes .NET (`System.Collections.Generic`, `System.IO`) dans PowerShell pour un gain de temps x50.

## 4. Nettoyage des Données
- **Octets Nuls** : Supprimer systématiquement `\0` (`Chr(0)`) lors de la lecture.
- **Normalisation** : Supprimer espaces, tabulations, `=`, retours chariot.
- **Trim** : Toujours faire un `.Trim()` final sur les numéros de série extraits.

## 5. Analyse de Statut
- **Succès** : Présence de `[PROD_OK]`.
- **Échec** : Présence de `[PROD_ERROR]`. Extraire le message d'erreur complet.
- **Priorité** : Si un bloc contient à la fois OK et ERROR (rare), ERROR prévaut généralement, ou l'ordre chronologique. Dans nos scripts : OK prévaut si présent, sauf si ERROR est explicite.
- **Incomplet** : Si aucun marqueur n'est trouvé avant la fin du bloc.

### Extraction du Défaut pour Nommage (Option 1 - Recherche KO)
- **Source** : Ligne contenant `[PROD_ERROR]: ... fail` dans chaque bloc.
- **Regex** : `\[PROD_ERROR\]:\s*(.+?)\s+fail\s*$` → capturer les mots entre `[PROD_ERROR]: ` et `fail`.
- **Fallback** : Si pas de `fail` → prendre tout après `[PROD_ERROR]: `. Si aucune ligne `[PROD_ERROR]:` → défaut = `Incomplet`.
- **Nommage fichier** : `SN_défaut1_défaut2_défaut3.txt` (un défaut par occurrence, dans l'ordre, non dédupliqué).
- **Exemples** : `043458_CP signal_CP signal_CP signal.txt`, `043491_LTE_LTE_LTE_LTE.txt`.

### Export des Manquants vers NumSerieKO.txt (Option 3 - Inventaire OK)
- Après la segmentation par lots, si des numéros manquants sont détectés, proposer via MessageBox de les ajouter dans `NumSerieKO.txt`.
- Le fichier est créé s'il n'existe pas, sinon les numéros sont ajoutés à la suite (un par ligne, UTF-8 sans BOM).

## 6. Historique des Bugs Résolus (Analyse)

| Erreur | Contexte | Solution |
| :--- | :--- | :--- |
| Caractères chinois (lecture) | VBS/PS | Forcer l'encodage ANSI-1252 à l'ouverture du fichier |
| Crash Write/Argument | VBS | Filtrer les caractères `AscW > 255` (Sanitize) |
| Espaces entre lettres | VBS | Désactiver le mode Unicode à l'écriture (`False`) |
| SN vide / non reconnu | Format 2026 | Adapter l'extraction pour supporter le format sans `#` |
| Bouton GUI caché | PowerShell | Désactiver le wrapping automatique (`WrapContents = $false`) |
| Accents corrompus (◇) | PS écriture ANSI | Passer à UTF-8 sans BOM pour l'écriture |
| "https" dans liste SN | URLs scannées | Regex `\d+` au lieu de `\w+` pour n'accepter que les chiffres |

## 7. Gestion et Automatisation (SFTP & RTC)

Pour le script de gestion (`_Gestion_Logs_Wattsy_Auto.ps1`), les règles suivantes s'appliquent :

### SFTP & WinSCP
- **WinSCP Portable** : Toujours utiliser `WinSCP.com` (CLI) avec `/ini=nul` pour éviter toute écriture dans le registre Windows.
- **Fingerprint** : Le fingerprint SSH doit être passé explicitement via le switch `-hostkey`.
- **Parsing des Sorties** : WinSCP injecte des lignes parasites (`batch abort`, etc.). Toujours utiliser des balises (ex: `echo __COUNT__`) pour isoler les données à parser en PowerShell.

### Intégrité des Données (Téléchargement)
- **Phase 1 (Sync)** : Utiliser `synchronize local` pour un transfert rapide des fichiers nouveaux ou modifiés.
- **Phase 2 (MD5)** : Vérifier systématiquement le contenu via `md5sum` sur le Pi et `Get-FileHash` sur le PC. Retélécharger si les hashs divergent.

### Gestion RTC (Module DS3231)
- **Lecture** : Exécuter le script existant sur le Pi (`test_rtc_2.py`).
- **Écriture** : Utiliser `python3 -c` avec une construction robuste pour injecter les octets BCD dans le bus I2C (adresse 0x68).
- **Reboot** : Un `sudo -n reboot` est recommandé après mise à jour du RTC.

## 8. Historique des Bugs Résolus (Gestion)

| Erreur | Contexte | Solution |
| :--- | :--- | :--- |
| `Unknown switch 'overwrite'` | WinSCP `get` | Supprimer le fichier local avant le `get` pour forcer l'écrasement. |
| Timeouts MD5 (>15s) | WinSCP `hash-list` | Utiliser la commande native Linux `md5sum` sur le Pi. |
| `SyntaxError` Python | RTC injection | Assurer que le code Python envoyé via `python -c` est mono-ligne. |
| Menu bloqué ou 0 ignoré | PowerShell Switch | Utiliser un label de boucle (ex: `:MainLoop`) pour le `break`. |
| Parsing "batch" en Int | Scan Logs | Utiliser des balises explicites (`__COUNT__`) pour le parsing. |

## 9. Segmentation par Lots / OF

Logique d'analyse de numéros de série validés (OK test) pour reconstituer des lots/OF et identifier les numéros manquants.
**Intégrée dans** `MasterLogTool.ps1` (option 3 — Inventaire Validés OK), et aussi disponible en standalone via `Liste_OF.ps1`.

### Flux de Données
```
Liste_OF.txt  -->  Liste_OF.ps1  -->  Liste_OF_traité.txt
   (liste brute)    (parse/tri/        (segments + manquants)
                     segmentation)
```

### Règles d'Implémentation (CRITIQUE)
- **Compatibilité PS 5.1** : Ne pas utiliser d'opérateurs non supportés (`??`, etc.). Toujours utiliser des conversions explicites (`[int]`).
- **Parsing** : Extraire les numéros via Regex `\d+` (numérique uniquement, cohérent avec §2).
- **Déduplication** : Utiliser une hashtable `$widthMap` pour dédupliquer tout en conservant la largeur d'affichage maximale observée.
- **Tri** : Tri numérique explicite (`Sort-Object { [int]$_ }`).
- **Segmentation** : Nouveau segment si l'écart entre deux numéros consécutifs triés dépasse `$GapThreshold` (défaut : 5). La comparaison doit être `$gap -gt $GapThreshold` sur des scalaires `[int]`.
- **Zéros initiaux** : Préserver l'affichage via une largeur par segment (max des largeurs observées dans le segment).
- **Exclusions** : Tableau `$Exclude` en tête de script (par forme texte).
- **Écriture** : UTF-8 via `Set-Content -Encoding UTF8`.
- **Fichiers** : `Liste_OF.txt` (entrée) et `Liste_OF_traité.txt` (sortie) dans le même dossier que le script.

### Format de Sortie Attendu
```
segment 1 : 043355–043544, present=188, missing=2 (043458, 043491)
segment 2 : 099001–099010, present=10, missing=0
```

## 10. Gestion des Ordres de Fabrication (OF)

Intégrée dans `MasterLogTool.ps1` (Option 5), cette fonctionnalité permet d'associer des numéros de série à des OF (7 chiffres).

### Stockage (`OF_Registry.json`)
- **Format** : JSON (Dictionnaire ordonné)
- **Clé** : Numéro d'OF (string, 7 chiffres)
- **Valeur** : Tableau de SN (strings)
- **Encodage** : UTF-8 sans BOM

**Exemple :**
```json
{
  "1234567": ["043355", "043356"],
  "9876543": ["044100", "044101"]
}
```

### Règles Métier
- **Unicité** : Un SN ne peut appartenir qu'à un seul OF.
- **Validation** : OF = 7 chiffres obligatoires. SN = chiffres uniquement.
- **Saisie** : Supporte l'ajout par plage (ex: `043590 - 043600`).
- **Persistance** : Sauvegarde automatique à chaque modification.

### Fonctions Utilitaires (API Interne)
- `Get-OFRegistry` : Charge le JSON en mémoire.
- `Save-OFRegistry($reg)` : Sauvegarde le JSON trié.
- `Find-OFBySN($reg, $sn)` : Retourne l'OF propriétaire d'un SN (ou `$null`).
- `Expand-SNRange($txt)` : Convertit une entrée (unique ou plage) en liste de SN.

## 11. Pièges PowerShell 5.1 Connus
- Un tableau non vide dans un `if` est évalué `True` même si son contenu est `$false`.
- `HashSet.ToArray()` peut échouer si la variable est écrasée en scalaire → préférer les hashtables.
- Toujours forcer `,` (opérateur virgule) pour ajouter un sous-tableau à un tableau : `$segments += ,$current`.

## 10. Historique des Bugs Résolus (Segmentation OF)

| Erreur | Contexte | Solution |
| :--- | :--- | :--- |
| `Jeton inattendu '??'` | PS 5.1 incompatible | Remplacer `($s ?? "")` par `if ($null -eq $s) { "" } else { $s }` |
| `System.Int32 ne contient pas ToArray()` | Variable écrasée en scalaire | Éviter `HashSet.ToArray()`, préférer hashtable + tri |
| 1 segment par numéro | Condition de segmentation toujours vraie | S'assurer que `$gap` et `$GapThreshold` sont des `[int]` scalaires, et utiliser `-gt` |
| Manquants non détectés | Effet secondaire du bug de segmentation | Corrigé par la correction de segmentation |
