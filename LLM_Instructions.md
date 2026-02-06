# Directives de Projet : Analyseur de Logs ELISA

Ce fichier définit les standards techniques et l'historique des résolutions pour garantir la performance et la stabilité des outils de traitement de logs (VBScript & PowerShell).

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

## 6. Historique des Bugs Résolus
| Erreur | Contexte | Solution |
| :--- | :--- | :--- |
| Caractères chinois (lecture) | VBS/PS | Forcer l'encodage ANSI-1252 à l'ouverture du fichier |
| Crash Write/Argument | VBS | Filtrer les caractères `AscW > 255` (Sanitize) |
| Espaces entre lettres | VBS | Désactiver le mode Unicode à l'écriture (`False`) |
| SN vide / non reconnu | Format 2026 | Adapter l'extraction pour supporter le format sans `#` |
| Bouton GUI caché | PowerShell | Désactiver le wrapping automatique (`WrapContents = $false`) |
| Accents corrompus (◇) | PS écriture ANSI | Passer à UTF-8 sans BOM pour l'écriture |
| "https" dans liste SN | URLs scannées | Regex `\d+` au lieu de `\w+` pour n'accepter que les chiffres | |
