# Directives de Projet : Analyseur de Logs ELISA (VBScript)

Ce fichier définit les standards techniques et l'historique des résolutions pour garantir la performance et la stabilité des outils de traitement de logs.

## 1. Encodage et Flux de Données (CRITIQUE)
- **Format de Fichier** : Toujours utiliser le format **ANSI** (Windows-1252 / TristateFalse) pour la lecture et l'écriture.
- **Lecture** : `objFSO.OpenTextFile(path, 1, False, 0)`
- **Écriture** : `objFSO.CreateTextFile(path, True, False)`
- **Problème rencontré** : L'utilisation de l'Unicode (UTF-16) provoque un affichage erroné de "caractères chinois" (mojibake) car les logs source sont en ANSI.
- **Crash "Argument Incorrect"** : Si une ligne de log contient des caractères non-ANSI (ex: symboles binaires ou spéciaux), l'écriture en mode ANSI échoue. 
- **Solution** : Utiliser la fonction **`Sanitize`** pour filtrer chaque caractère (garder uniquement `AscW(c) <= 255`) avant l'écriture finale en mode ANSI.

## 2. Performance et Architecture
- **Algorithme** : Proscrire les boucles imbriquées $O(N \times M)$.
- **Pattern Optimal** : 
    1. Charger les cibles (si présentes) dans un `Scripting.Dictionary`.
    2. Parcourir chaque fichier de log **une seule fois**.
    3. Stocker les résultats intermédiaires dans un dictionnaire global.
    4. Effectuer une écriture unique sur disque à la fin.
- **Mémoire** : Utiliser `Scripting.Dictionary` pour les recherches instantanées.

## 3. Logique d'Extraction (Blocs)
- **Délimiteur de Début** : Une ligne contenant le mot-clé `Datamatrix` (recherche insensible à la casse `vbTextCompare`).
- **Extraction du Numéro de Série (SN)** : 
    - **Ancien Format** (ex: `#2025#13104`) : Le SN est en 3ème position (`Split` sur `#` index 2).
    - **Nouveau Format** (ex: `Datamatrix: 043355`) : Le SN suit directement "Datamatrix:" (Regex ou Split sur `:`).
    - **Solution Unifiée** : Utiliser une fonction `Extract-SN` capable de détecter et traiter les deux cas.
- **Délimiteur de Fin** : Un bloc se termine à la prochaine ligne `Datamatrix` ou à la fin du fichier.
- **Gestion des Répétitions** : Si un même SN apparaît plusieurs fois, chaque occurrence doit être isolée ou comptabilisée séparément (ne pas fusionner aveuglément).

## 4. Nettoyage des Données (Fonction CleanSerial)
- **Octets Nuls** : Supprimer systématiquement `Chr(0)` lors de la lecture de chaque ligne.
- **Normalisation SN** : Supprimer les espaces, tabulations, retours à la ligne (`vbCr`, `vbLf`) et le signe `=`. 
- **Important** : Toujours appliquer un `Trim()` final sur le numéro de série pour éviter les clés de dictionnaire erronées.

## 5. Analyse de Statut (Logique de Test)
- **Succès** : Présence de la chaîne `[PROD_OK]` dans le bloc.
- **Échec** : Présence de `[PROD_ERROR]`. Pour plus de précision, extraire la ligne complète pour identifier la cause (ex: `LTE fail`).
- **Incomplet** : Cas par défaut si aucun des deux marqueurs n'est trouvé avant le prochain `Datamatrix`.

## 6. Interface et Sortie
- **Résultat Console** : Un seul `WScript.Echo` à la fin pour un résumé silencieux.
- **Tri des Résultats** : Pour les fichiers d'inventaire, proposer deux vues : 
    1. **Ordre d'apparition** (chronologique du log).
    2. **Ordre croissant** (facilitant la recherche manuelle).

## 7. Résolution des Bugs Historiques
| Erreur rencontrée | Cause | Solution |
| :--- | :--- | :--- |
| Caractères chinois | Lecture forcée en Unicode | Forcer TristateFalse (0) à l'ouverture du fichier |
| Crash lors de l'écriture | Caractères spéciaux > 255 | Fonction `Sanitize` (filtrage AscW) |
| Entêtes manquants | Vérification `If results <> ""` | Ajouter l'en-tête systématiquement dès l'occurrence #1 |
| Fichier vide | Casse de "Datamatrix" | Utiliser `InStr(..., vbTextCompare)` |
| Espaces entre les lettres | Mode Unicode activé à l'écriture | Revenir au mode ANSI dans `CreateTextFile` |
