# Outil de cryptage/décryptage des NNSS

Ce projet fournit un script **PowerShell** permettant de chiffrer ou de déchiffrer des numéros NNSS (NAVS) contenus dans des fichiers CSV ou Excel. Il a pour but de faciliter l’échange sécurisé de ces données entre le SPC et l’Hospice Général.

## Sommaire
- [Fonctionnement](#fonctionnement)
- [Principes cryptographiques](#principes-cryptographiques)
- [Prérequis](#prérequis)
- [Installation](#installation)
- [Utilisation](#utilisation)
- [Bonnes pratiques de sécurité](#bonnes-pratiques-de-sécurité)
- [Contribuer](#contribuer)
- [Licence](#licence)

## Fonctionnement
- **Cryptage AES‑256‑CBC** : les fonctions `Protect-NNSSData` et `Unprotect-NNSSData` appliquent l’algorithme AES en mode CBC avec remplissage PKCS7 pour chaque valeur à traiter.
- **Clé et IV ajustés** : `ConvertTo-SecureAESKey` s’assure que la clé (32 octets) et le vecteur d’initialisation (16 octets) sont aux bonnes longueurs avant l’opération 【F:functions/crypt-functions.ps1†L1-L34】.
- **Traitement CSV/Excel** : selon l’extension du fichier, `Process-CSVFile` ou `Process-ExcelFile` charge les données, applique le cryptage ou le décryptage et sauvegarde un nouveau fichier 【F:functions/crypt-functions.ps1†L178-L229】【F:functions/crypt-functions.ps1†L233-L314】.
- **Interface graphique** : le script génère un formulaire Windows Forms permettant de choisir le fichier d’entrée, de sélectionner la colonne NNSS, de saisir la clé partagée et l’IV puis de lancer le traitement.
- **Détection automatique de colonne** : lors du choix du fichier, le script tente d’identifier la colonne contenant les numéros NNSS/NAVS pour simplifier la sélection.
- **Indicateur de progression** : une barre d’avancement et des messages d’état informent de la progression du traitement.

## Principes cryptographiques
L’outil met en œuvre la norme **AES** (Advanced Encryption Standard) dans sa
version 256 bits. AES est un chiffrement symétrique en bloc reposant sur un
réseau de substitutions et de permutations. Pour une clé de cette longueur,
quatorze tours successifs sont appliqués, chacun combinant les opérations
**SubBytes**, **ShiftRows**, **MixColumns** et **AddRoundKey** afin d’assurer la
diffusion et la confusion des données.

Le bloc traité mesure toujours 16 octets (128 bits). En mode **CBC**
(Cipher Block Chaining), chaque bloc est préalablement « XORé » avec le bloc
chiffré précédent ou avec l’IV pour le premier bloc, garantissant ainsi que deux
blocs identiques ne produisent pas le même résultat. Le vecteur d’initialisation
transmis est ajusté à exactement 16 octets par la fonction
`ConvertTo-SecureAESKey` 【F:functions/crypt-functions.ps1†L12-L28】, tandis que la
clé est tronquée ou complétée à 32 octets 【F:functions/crypt-functions.ps1†L14-L21】.

Le résultat du chiffrement est encodé en Base64 pour pouvoir être stocké dans un
fichier texte ou tableur. Ce script n’ajoute pas de mécanisme d’authentification
des données ; l’ajout d’une signature HMAC ou l’emploi d’un mode authentifié
(par exemple AES‑GCM) est recommandé si l’intégrité doit être garantie.

## Prérequis
- Windows avec **PowerShell 5** ou version ultérieure.
- Microsoft Excel installé pour traiter les fichiers `.xlsx`/`.xls` (utilisation d’`Microsoft.Office.Interop.Excel`).
- Le script charge plusieurs assemblies .NET via `Add-Type` 【F:nnss-crypt.ps1†L16-L24】.

## Installation
1. Télécharger ou cloner ce dépôt.
2. Ouvrir une console PowerShell dans le répertoire du projet.

## Utilisation
Exécuter le script avec la politique d’exécution levée :

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\nnss-crypt.ps1
```

Dans la fenêtre qui s’ouvre :
1. Parcourir et sélectionner le fichier CSV ou Excel contenant les numéros à traiter (les colonnes sont chargées automatiquement).
2. Choisir la colonne cible si elle n’a pas été détectée.
3. Renseigner la clé et le vecteur d’initialisation identiques pour l’émetteur et le destinataire.
4. Choisir le mode **Crypter** ou **Décrypter** puis l’emplacement du fichier de sortie.
5. Lancer l’opération et patienter jusqu’à l’affichage du message de résultat.

Le fichier résultant est enregistré à l’emplacement indiqué.

## Bonnes pratiques de sécurité
- Utiliser une clé et un IV suffisamment longs (le script vérifie une longueur minimale de 12 caractères pour la clé et 8 pour l’IV).
- Transmettre ces valeurs par un canal sécurisé et éviter de les conserver en clair.
- Exécuter le script dans un environnement de confiance uniquement.

## Contribuer
Les demandes d’amélioration ou de correction sont les bienvenues via un système de Pull Request. Merci de décrire brièvement la modification proposée et de respecter la structure existante du projet.

## Licence
Aucune licence open source n’est fournie dans ce dépôt. Veuillez contacter les mainteneurs pour toute question d’utilisation ou de redistribution.

