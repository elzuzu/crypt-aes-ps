# Fonctions de cryptage/décryptage

function Test-Base64String {
    param (
        [Parameter(Mandatory = $true)]
        [string]$InputString
    )

    if ([string]::IsNullOrEmpty($InputString)) {
        return $false
    }

    if ($InputString.Length % 4 -ne 0) {
        return $false
    }

    $base64Pattern = '^[A-Za-z0-9+/]*={0,2}$'
    if ($InputString -notmatch $base64Pattern) {
        return $false
    }

    try {
        [Convert]::FromBase64String($InputString) | Out-Null
        return $true
    }
    catch {
        return $false
    }
}
function ConvertTo-SecureAESKey {
    param (
        [Parameter(Mandatory = $true)]
        [string]$KeyString,
        [Parameter(Mandatory = $true)]
        [string]$InitVector
    )
    
    # S'assurer que la clé et l'IV sont de la bonne longueur pour AES-256-CBC
    $KeyBytes = [System.Text.Encoding]::UTF8.GetBytes($KeyString)
    $IVBytes = [System.Text.Encoding]::UTF8.GetBytes($InitVector)
    
    # Ajuster la taille de la clé à 32 octets (256 bits)
    if ($KeyBytes.Length -gt 32) {
        $KeyBytes = $KeyBytes[0..31]
    }
    elseif ($KeyBytes.Length -lt 32) {
        $KeyBytes = $KeyBytes + (New-Object byte[] (32 - $KeyBytes.Length))
    }
    
    # Ajuster la taille de l'IV à 16 octets (128 bits)
    if ($IVBytes.Length -gt 16) {
        $IVBytes = $IVBytes[0..15]
    }
    elseif ($IVBytes.Length -lt 16) {
        $IVBytes = $IVBytes + (New-Object byte[] (16 - $IVBytes.Length))
    }
    
    return @{
        Key = $KeyBytes
        IV  = $IVBytes
    }
}

function Protect-NNSSData {
    param (
        [Parameter(Mandatory = $true)]
        [string]$InputText,
        [Parameter(Mandatory = $true)]
        [byte[]]$Key,
        [Parameter(Mandatory = $true)]
        [byte[]]$InitVector
    )

    try {
        if (Test-Base64String -InputString $InputText) {
            Write-Warning "La valeur '$InputText' semble déjà être cryptée (format Base64). Valeur ignorée."
            return $InputText
        }

        # Convertir la chaîne en tableau d'octets
        $InputBytes = [System.Text.Encoding]::UTF8.GetBytes($InputText)
        
        # Créer un objet de cryptage AES
        $AES = [System.Security.Cryptography.Aes]::Create()
        $AES.Key = $Key
        $AES.IV = $InitVector
        $AES.Mode = [System.Security.Cryptography.CipherMode]::CBC
        $AES.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
        
        # Effectuer le cryptage
        $Encryptor = $AES.CreateEncryptor()
        $EncryptedBytes = $Encryptor.TransformFinalBlock($InputBytes, 0, $InputBytes.Length)
        
        # Convertir le résultat en chaîne Base64
        $EncryptedText = [Convert]::ToBase64String($EncryptedBytes)
        
        return $EncryptedText
    }
    catch {
        $errorMessage = "Erreur lors du cryptage de '$InputText': $($_.Exception.Message)"
        Write-Warning $errorMessage
        return "[ERREUR_CRYPT] $InputText"
    }
    finally {
        if ($AES) {
            $AES.Dispose()
        }
    }
}

function Unprotect-NNSSData {
    param (
        [Parameter(Mandatory = $true)]
        [string]$EncryptedText,
        [Parameter(Mandatory = $true)]
        [byte[]]$Key,
        [Parameter(Mandatory = $true)]
        [byte[]]$InitVector
    )

    try {
        if (-not (Test-Base64String -InputString $EncryptedText)) {
            Write-Warning "La valeur '$EncryptedText' ne semble pas être cryptée (format Base64 invalide). Valeur ignorée."
            return $EncryptedText
        }

        $EncryptedBytes = [Convert]::FromBase64String($EncryptedText)

        if ($EncryptedBytes.Length -lt 16) {
            Write-Warning "La valeur cryptée '$EncryptedText' est trop courte pour être valide. Valeur ignorée."
            return $EncryptedText
        }
        
        # Créer un objet de décryptage AES
        $AES = [System.Security.Cryptography.Aes]::Create()
        $AES.Key = $Key
        $AES.IV = $InitVector
        $AES.Mode = [System.Security.Cryptography.CipherMode]::CBC
        $AES.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
        
        # Effectuer le décryptage
        $Decryptor = $AES.CreateDecryptor()
        $DecryptedBytes = $Decryptor.TransformFinalBlock($EncryptedBytes, 0, $EncryptedBytes.Length)
        
        # Convertir le résultat en chaîne
        $DecryptedText = [System.Text.Encoding]::UTF8.GetString($DecryptedBytes)
        
        return $DecryptedText
    }
    catch {
        $errorMessage = "Erreur lors du décryptage de '$EncryptedText': $($_.Exception.Message)"
        Write-Warning $errorMessage
        return "[ERREUR_DECRYPT] $EncryptedText"
    }
    finally {
        if ($AES) {
            $AES.Dispose()
        }
    }
}

# Fonctions de traitement des fichiers
function Get-FileColumns {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()

    if ($extension -eq ".csv") {
        # Pour les fichiers CSV
        try {
            # Lecture plus robuste pour PowerShell 5
            $firstLine = Get-Content $FilePath -First 1 -Encoding UTF8
            $headers = $firstLine -split ',' | ForEach-Object { $_.Trim('"') }
            return $headers
        }
        catch {
            [System.Windows.MessageBox]::Show("Erreur lors de la lecture du fichier CSV: $_", "Erreur", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return $null
        }
    }
    elseif ($extension -eq ".xlsx" -or $extension -eq ".xls") {
        # Vérifier si Excel est disponible
        if (-not $global:ExcelAvailable) {
            [System.Windows.MessageBox]::Show("Microsoft Excel n'est pas installé. Veuillez utiliser un fichier CSV ou installer Microsoft Excel pour traiter les fichiers Excel.", "Excel non disponible", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return $null
        }

        # Pour les fichiers Excel (si disponible)
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false

            $workbook = $excel.Workbooks.Open($FilePath)
            $worksheet = $workbook.Sheets.Item(1)
            $range = $worksheet.UsedRange

            $columnCount = $range.Columns.Count
            $headers = @()

            for ($i = 1; $i -le $columnCount; $i++) {
                $cellValue = $worksheet.Cells.Item(1, $i).Text
                $headers += $cellValue
            }

            $workbook.Close($false)
            $excel.Quit()

            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            [System.GC]::Collect()

            return $headers
        }
        catch {
            [System.Windows.MessageBox]::Show("Erreur lors de la lecture du fichier Excel: $_", "Erreur", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return $null
        }
    }
    else {
        [System.Windows.MessageBox]::Show("Format de fichier non pris en charge. Veuillez utiliser un fichier CSV ou Excel.", "Erreur", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $null
    }
}

function Process-CSVFile {
    param (
        [Parameter(Mandatory = $true)]
        [string]$InputFilePath,
        [Parameter(Mandatory = $true)]
        [string]$OutputFilePath,
        [Parameter(Mandatory = $true)]
        [string]$ColumnName,
        [Parameter(Mandatory = $true)]
        [byte[]]$Key,
        [Parameter(Mandatory = $true)]
        [byte[]]$InitVector,
        [Parameter(Mandatory = $true)]
        [bool]$IsEncryption
    )
    
    try {
        if (!(Test-Path $InputFilePath)) {
            Write-Error "Le fichier d'entrée n'existe pas : $InputFilePath"
            return $false
        }

        # Charger les données CSV
        $csvData = Import-Csv -Path $InputFilePath

        # Traiter chaque ligne
        $processedCount = 0
        $errorCount = 0
        $skippedCount = 0

        foreach ($row in $csvData) {
            # S'assurer que la colonne existe
            if ($row.PSObject.Properties.Name -contains $ColumnName) {
                $originalValue = $row.$ColumnName

                # Ignorer les valeurs vides ou nulles
                if (![string]::IsNullOrEmpty($originalValue)) {
                    if ($IsEncryption) {
                        # Crypter
                        $newValue = Protect-NNSSData -InputText $originalValue -Key $Key -InitVector $InitVector
                    }
                    else {
                        # Décrypter
                        $newValue = Unprotect-NNSSData -EncryptedText $originalValue -Key $Key -InitVector $InitVector
                    }

                    if ($newValue -ne $null) {
                        if ($newValue.StartsWith("[ERREUR_")) {
                            $errorCount++
                            Write-Warning "Erreur de traitement pour la valeur: $originalValue"
                        } else {
                            $row.$ColumnName = $newValue
                            $processedCount++
                        }
                    }
                } else {
                    $skippedCount++
                }
            }
        }

        # Enregistrer le fichier modifié
        $csvData | Export-Csv -Path $OutputFilePath -NoTypeInformation

        Write-Host "Traitement terminé - Traités: $processedCount, Erreurs: $errorCount, Ignorés: $skippedCount" -ForegroundColor Green

        return $true
    }
    catch {
        Write-Error "Erreur dans Process-CSVFile: $($_.Exception.Message)"
        return $false
    }
}

function Process-ExcelFile {
    param (
        [Parameter(Mandatory = $true)]
        [string]$InputFilePath,
        [Parameter(Mandatory = $true)]
        [string]$OutputFilePath,
        [Parameter(Mandatory = $true)]
        [string]$ColumnName,
        [Parameter(Mandatory = $true)]
        [byte[]]$Key,
        [Parameter(Mandatory = $true)]
        [byte[]]$InitVector,
        [Parameter(Mandatory = $true)]
    [bool]$IsEncryption
    )

    if (-not $global:ExcelAvailable) {
        Write-Error "Microsoft Excel n'est pas disponible"
        return $false
    }

    try {
        if (!(Test-Path $InputFilePath)) {
            Write-Error "Le fichier d'entrée n'existe pas : $InputFilePath"
            return $false
        }

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        $workbook = $excel.Workbooks.Open($InputFilePath)
        $worksheet = $workbook.Sheets.Item(1)
        $range = $worksheet.UsedRange
        
        $rowCount = $range.Rows.Count
        $columnCount = $range.Columns.Count
        
        # Trouver l'index de la colonne
        $columnIndex = $null
        for ($i = 1; $i -le $columnCount; $i++) {
            if ($worksheet.Cells.Item(1, $i).Text -eq $ColumnName) {
                $columnIndex = $i
                break
            }
        }
        
        if ($columnIndex -eq $null) {
            [System.Windows.MessageBox]::Show("La colonne spécifiée n'a pas été trouvée dans le fichier Excel.", "Erreur", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            $workbook.Close($false)
            $excel.Quit()
            return $false
        }
        
        # Traiter chaque cellule de la colonne
        for ($i = 2; $i -le $rowCount; $i++) {
            $cellValue = $worksheet.Cells.Item($i, $columnIndex).Text
            
            # Ignorer les valeurs vides ou nulles
            if (![string]::IsNullOrEmpty($cellValue)) {
                if ($IsEncryption) {
                    # Crypter
                    $newValue = Protect-NNSSData -InputText $cellValue -Key $Key -InitVector $InitVector
                }
                else {
                    # Décrypter
                    $newValue = Unprotect-NNSSData -EncryptedText $cellValue -Key $Key -InitVector $InitVector
                }
                
                if ($newValue -ne $null) {
                    $worksheet.Cells.Item($i, $columnIndex) = $newValue
                }
            }
        }
        
        # Enregistrer le fichier modifié
        $workbook.SaveAs($OutputFilePath)
        $workbook.Close($true)
        $excel.Quit()
        
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        
        return $true
    }
    catch {
        Write-Error "Erreur dans Process-ExcelFile: $($_.Exception.Message)"
        return $false
    }
}

