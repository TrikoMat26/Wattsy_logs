# Boîte à Outils d'Analyse de Logs ELISA (VBScript)

Ce dossier contient un ensemble de scripts VBScript conçus pour automatiser l'analyse, l'extraction et l'inventaire des logs de production ELISA (`ELISA_Prod_Log_*.txt`).

---

## 1. RechercheSerie.vbs (Extracteur de Blocs)
**Utilité :** Extraire l'intégralité du contenu des logs pour une liste spécifique de produits.
- **Entrée :** Nécessite un fichier `NumSerieKO.txt` contenant un numéro de série par ligne.
- **Fonctionnement :** Parcourt tous les logs et recherche chaque numéro de série du fichier source.
- **Sortie :** Crée un fichier `.txt` individuel pour chaque numéro de série trouvé (ex: `13104.txt`).
- **Particularité :** Chaque passage d'un même numéro est séparé par un en-tête clair indiquant l'occurrence et le fichier d'origine.

## 2. InventaireSeries.vbs (Inventaire Global)
**Utilité :** Obtenir rapidement la liste de TOUS les produits passés sur le banc de test.
- **Fonctionnement :** Identifie chaque ligne `Datamatrix` dans les fichiers de logs.
- **Sortie :** Génère `Inventaire_Series.txt`.
- **Contenu :** Liste pour chaque fichier log le nombre de produits uniques et la liste de leurs numéros de série.

## 3. InventaireSeriesOK.vbs (Inventaire des Succès & Doublons)
**Utilité :** Identifier uniquement les produits ayant réussi les tests et détecter les anomalies de répétition.
- **Critère :** Ne retient que les produits dont le bloc de log contient le marqueur `[PROD_OK]`.
- **Sortie :** Génère `Inventaire_Series_OK.txt`.
- **Fonctionnalités avancées :** 
    - Affiche les SN par **ordre d'apparition** (chronologique).
    - Affiche les SN par **ordre croissant** (numérique).
    - **Analyse des doublons** : Liste les produits testés plusieurs fois dans le même log ou à travers différents fichiers.

## 4. HistoriqueTests.vbs (Traçabilité Complète)
**Utilité :** Suivre le parcours de chaque produit et comprendre pourquoi certains tests ont échoué.
- **Fonctionnement :** Analyse chaque bloc de test et catégorise le résultat.
- **Sortie :** Génère `Historique_Tests_Complet.txt`.
- **Catégories de résultats :**
    - **Test OK** : Le marqueur `[PROD_OK]` a été trouvé.
    - **[PROD_ERROR]...** : Affiche le message d'erreur précis (ex: `LTE fail`) si le test a échoué.
    - **Test Incomplet** : Aucun marqueur de fin n'a été trouvé (coupure de log, crash, etc.).

---

## Informations Techniques Communes
- **Encodage :** Tous les scripts sont optimisés pour lire et écrire au format **ANSI**, garantissant l'absence de "caractères chinois" (mojibake).
- **Performance :** Utilisation de dictionnaires mémoire (`Scripting.Dictionary`) pour traiter des milliers de lignes en un seul passage par fichier.
- **Nettoyage :** Suppression automatique des octets nuls (`Chr(0)`) et des caractères parasites (`=`, espaces, tabulations) pour des données propres.
- **Instructions LLM :** Le fichier `LLM_Instructions.md` contient les règles techniques pour permettre à une IA de maintenir ces scripts sans introduire de bugs d'encodage ou de performance.
