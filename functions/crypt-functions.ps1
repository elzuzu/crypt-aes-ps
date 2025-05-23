# Fonctions de cryptage/décryptage
function ConvertTo-SecureAESKey {
    param (
        [Parameter(Mandatory = $true)]
        [string]$KeyString,
        [Parameter(Mandatory = $true)]
        [string]$IVString
    )
    
    # S'assurer que la clé et l'IV sont de la bonne longueur pour AES-256-CBC
    $KeyBytes = [System.Text.Encoding]::UTF8.GetBytes($KeyString)
    $IVBytes = [System.Text.Encoding]::UTF8.GetBytes($IVString)
    
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
        [byte[]]$IV
    )
    
    try {
        # Convertir la chaîne en tableau d'octets
        $InputBytes = [System.Text.Encoding]::UTF8.GetBytes($InputText)
        
        # Créer un objet de cryptage AES
        $AES = [System.Security.Cryptography.Aes]::Create()
        $AES.Key = $Key
        $AES.IV = $IV
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
        [System.Windows.MessageBox]::Show("Erreur lors du cryptage: $_", "Erreur", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $null
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
        [byte[]]$IV
    )
    
    try {
        # Convertir la chaîne Base64 en tableau d'octets
        $EncryptedBytes = [Convert]::FromBase64String($EncryptedText)
        
        # Créer un objet de décryptage AES
        $AES = [System.Security.Cryptography.Aes]::Create()
        $AES.Key = $Key
        $AES.IV = $IV
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
        [System.Windows.MessageBox]::Show("Erreur lors du décryptage: $_", "Erreur", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $null
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
            $headers = (Get-Content $FilePath -First 1) -split ','
            return $headers
        }
        catch {
            [System.Windows.MessageBox]::Show("Erreur lors de la lecture du fichier CSV: $_", "Erreur", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return $null
        }
    }
    elseif ($extension -eq ".xlsx" -or $extension -eq ".xls") {
        # Pour les fichiers Excel
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
        [byte[]]$IV,
        [Parameter(Mandatory = $true)]
        [bool]$IsEncryption
    )
    
    try {
        # Charger les données CSV
        $csvData = Import-Csv -Path $InputFilePath
        
        # Traiter chaque ligne
        foreach ($row in $csvData) {
            # S'assurer que la colonne existe
            if ($row.PSObject.Properties.Name -contains $ColumnName) {
                $originalValue = $row.$ColumnName
                
                # Ignorer les valeurs vides ou nulles
                if (![string]::IsNullOrEmpty($originalValue)) {
                    if ($IsEncryption) {
                        # Crypter
                        $newValue = Protect-NNSSData -InputText $originalValue -Key $Key -IV $IV
                    }
                    else {
                        # Décrypter
                        $newValue = Unprotect-NNSSData -EncryptedText $originalValue -Key $Key -IV $IV
                    }
                    
                    if ($newValue -ne $null) {
                        $row.$ColumnName = $newValue
                    }
                }
            }
        }
        
        # Enregistrer le fichier modifié
        $csvData | Export-Csv -Path $OutputFilePath -NoTypeInformation
        
        return $true
    }
    catch {
        [System.Windows.MessageBox]::Show("Erreur lors du traitement du fichier CSV: $_", "Erreur", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
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
        [byte[]]$IV,
        [Parameter(Mandatory = $true)]
        [bool]$IsEncryption
    )
    
    try {
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
                    $newValue = Protect-NNSSData -InputText $cellValue -Key $Key -IV $IV
                }
                else {
                    # Décrypter
                    $newValue = Unprotect-NNSSData -EncryptedText $cellValue -Key $Key -IV $IV
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
        [System.Windows.MessageBox]::Show("Erreur lors du traitement du fichier Excel: $_", "Erreur", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $false
    }
}

