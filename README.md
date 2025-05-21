# Outil de cryptage/décryptage des NNSS

Ce dépôt contient un script PowerShell permettant de chiffrer ou de déchiffrer des numéros NNSS (ou NAVS) stockés dans un fichier CSV ou Excel. Le script a été conçu pour faciliter les échanges sécurisés entre le SPC et l'Hospice Général.

## Fonctionnement

- **Cryptage AES‑256‑CBC** : les fonctions `Protect-NNSSData` et `Unprotect-NNSSData` utilisent l'algorithme AES avec un mode CBC et un remplissage PKCS7 pour chiffrer ou déchiffrer la valeur de chaque cellule.
- **Clé et IV ajustés** : `ConvertTo-SecureAESKey` convertit les chaînes fournies en tableau d'octets et ajuste la longueur de la clé (32 octets) et du vecteur d'initialisation (16 octets) 【F:nnss-crypt.ps1†L17-L48】.
- **Traitement CSV/Excel** : selon l'extension du fichier, les fonctions `Process-CSVFile` ou `Process-ExcelFile` chargent les données, appliquent le cryptage/décryptage et enregistrent un nouveau fichier 【F:nnss-crypt.ps1†L193-L238】【F:nnss-crypt.ps1†L248-L318】.
- **Interface graphique** : le script crée un formulaire Windows Forms permettant de choisir le fichier d'entrée, la colonne NNSS, la clé partagée, l'IV et l'emplacement du fichier de sortie avant de lancer l'opération.

## Prérequis

- Windows avec **PowerShell 5**.
- Microsoft Excel doit être installé pour traiter les fichiers `.xlsx`/`.xls` (utilisation d'`Microsoft.Office.Interop.Excel`).
- Le script charge plusieurs assemblies .NET via `Add-Type` 【F:nnss-crypt.ps1†L9-L14】.

## Utilisation

1. Ouvrir une console PowerShell.
2. Exécuter le script en autorisant son lancement :

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\nnss-crypt.ps1
```

3. Dans la fenêtre qui s'ouvre :
   - Sélectionner le fichier CSV ou Excel contenant les NNSS.
   - Choisir la colonne à traiter.
   - Saisir la clé et le vecteur d'initialisation identiques entre les deux parties.
   - Choisir si l'on souhaite crypter ou décrypter puis indiquer le fichier de sortie.
   - Lancer le traitement.

Le fichier résultant sera enregistré à l'emplacement indiqué.

## Avertissement

Vérifiez que la clé et l'IV utilisés sont transmis de manière sécurisée et ne sont pas stockés en clair. Le script doit être exécuté dans un environnement de confiance.

