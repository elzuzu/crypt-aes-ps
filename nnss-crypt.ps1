# Script de Cryptage/Décryptage des NNSS pour l'échange SPC-Hospice Général
# Auteur: Alexandre Wyer
# Date: Mai 2025
# Version: 1.0
#
# Ce script permet de crypter ou décrypter les numéros NNSS dans un fichier CSV ou Excel
# en utilisant l'algorithme AES-256-CBC avec une clé et un vecteur d'initialisation partagés.

# Fonction de vérification de compatibilité
function Test-PowerShellCompatibility {
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        Write-Error "Ce script nécessite PowerShell 5.0 ou supérieur. Version actuelle: $($PSVersionTable.PSVersion)" -Category InvalidArgument
        return $false
    }

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
    }
    catch {
        Write-Error "Impossible de charger les assemblies .NET requises: $_" -Category InvalidOperation
        return $false
    }

    return $true
}

if (-not (Test-PowerShellCompatibility)) {
    return
}

# Charger les assemblies nécessaires
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# Vérifier si Excel est disponible (optionnel)
$global:ExcelAvailable = $false
try {
    $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    $global:ExcelAvailable = $true
    Write-Host "Microsoft Excel détecté - Support des fichiers .xlsx/.xls activé" -ForegroundColor Green
}
catch {
    Write-Host "Microsoft Excel non détecté - Seuls les fichiers CSV seront supportés" -ForegroundColor Yellow
}
# Charger les fonctions de traitement
. (Join-Path $PSScriptRoot "functions/crypt-functions.ps1")
. (Join-Path $PSScriptRoot "functions/ui-styles.ps1")

# Fonction de redimensionnement intelligente pour les contrôles
function Update-ControlSizes {
    param($FormWidth)

    $availableWidth = $FormWidth - 60
    $textBoxWidth = [Math]::Max(350, $availableWidth - 150)
    $buttonX = $textBoxWidth + 30

    if($inputFileTextBox){
        $inputFileTextBox.Width = $textBoxWidth
        $inputFileBrowseButton.Left = $buttonX
    }
    if($outputFileTextBox){
        $outputFileTextBox.Width = $textBoxWidth
        $outputFileBrowseButton.Left = $buttonX
    }
    if($keyTextBox){ $keyTextBox.Width = [Math]::Min(450, $textBoxWidth) }
    if($ivTextBox){ $ivTextBox.Width = [Math]::Min(450, $textBoxWidth) }

    if($processButton){ $processButton.Left = [Math]::Max(380, $FormWidth - 320) }
    if($cancelButton){ $cancelButton.Left = [Math]::Max(510, $FormWidth - 190) }
}

# Créer l'interface utilisateur
$form = New-Object System.Windows.Forms.Form
$form.Text = "SPC - Cryptage/Décryptage des NNSS"
$form.Size = New-Object System.Drawing.Size(700, 580)
$form.MinimumSize = New-Object System.Drawing.Size(650, 550)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.MaximizeBox = $true
$form.AutoScroll = $false  # Désactiver pour éviter les décalages
$form.BackColor = $themeColors.Background
$form.ForeColor = $themeColors.TextDark
$form.Font = $fontRegular

$form.Padding = New-Object System.Windows.Forms.Padding(15)  # Padding réduit

# Utiliser un panel principal simple
$mainPanel = New-Object System.Windows.Forms.Panel
$mainPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$mainPanel.AutoScroll = $true
$mainPanel.BackColor = $themeColors.Background
$form.Controls.Add($mainPanel)

# Header - Position Y = 10
$lockIcon = New-Object System.Windows.Forms.Label
$lockIcon.Text = "🔒"
$lockIcon.Size = New-Object System.Drawing.Size(40, 40)
$lockIcon.Location = New-Object System.Drawing.Point(20, 10)
$lockIcon.Font = New-Object System.Drawing.Font("Segoe UI", 20)
$lockIcon.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$lockIcon.ForeColor = $themeColors.Primary
$mainPanel.Controls.Add($lockIcon)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(70, 10)
$titleLabel.Size = New-Object System.Drawing.Size(500, 25)
$titleLabel.Text = "Cryptage/Décryptage des NNSS"
$titleLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 12)
$titleLabel.ForeColor = $themeColors.TextDark
$mainPanel.Controls.Add($titleLabel)

$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Location = New-Object System.Drawing.Point(70, 35)
$subtitleLabel.Size = New-Object System.Drawing.Size(500, 15)
$subtitleLabel.Text = "Outil pour l'échange sécurisé de données SPC-Hospice Général"
$subtitleLabel.Font = $fontSmall
$subtitleLabel.ForeColor = [System.Drawing.Color]::FromArgb(96, 94, 92)
$mainPanel.Controls.Add($subtitleLabel)

# Séparateur 1 - Position Y = 60
$separator1 = New-Object System.Windows.Forms.Panel
$separator1.Location = New-Object System.Drawing.Point(20, 60)
$separator1.Size = New-Object System.Drawing.Size(630, 1)
$separator1.BackColor = $themeColors.Border
$mainPanel.Controls.Add($separator1)

# Section 1 - Position Y = 70
$fileSelectionLabel = New-Object System.Windows.Forms.Label
$fileSelectionLabel.Location = New-Object System.Drawing.Point(20, 70)
$fileSelectionLabel.Size = New-Object System.Drawing.Size(630, 20)
$fileSelectionLabel.Text = "1. Sélection du fichier"
$fileSelectionLabel.Font = $fontHeader
$fileSelectionLabel.ForeColor = $themeColors.Primary
$mainPanel.Controls.Add($fileSelectionLabel)

$inputFileLabel = New-Object System.Windows.Forms.Label
$inputFileLabel.Location = New-Object System.Drawing.Point(20, 95)
$inputFileLabel.Size = New-Object System.Drawing.Size(150, 20)
$inputFileLabel.Text = "Fichier d'entrée:"
$inputFileLabel.Font = $fontRegular
$mainPanel.Controls.Add($inputFileLabel)

$inputFileInfo = New-Object System.Windows.Forms.Label
$inputFileInfo.Location = New-Object System.Drawing.Point(170, 95)
$inputFileInfo.Size = New-Object System.Drawing.Size(400, 20)
$inputFileInfo.Text = "(CSV ou Excel contenant les NNSS à traiter)"
$inputFileInfo.Font = $fontSmall
$inputFileInfo.ForeColor = [System.Drawing.Color]::FromArgb(96, 94, 92)
$mainPanel.Controls.Add($inputFileInfo)

$inputFileTextBox = New-Object System.Windows.Forms.TextBox
$inputFileTextBox.Location = New-Object System.Drawing.Point(20, 118)
$inputFileTextBox.Size = New-Object System.Drawing.Size(480, 25)
$inputFileTextBox.ReadOnly = $true
Set-ModernTextBoxStyle -TextBox $inputFileTextBox
$mainPanel.Controls.Add($inputFileTextBox)

$inputFileBrowseButton = New-Object System.Windows.Forms.Button
$inputFileBrowseButton.Location = New-Object System.Drawing.Point(510, 118)
$inputFileBrowseButton.Size = New-Object System.Drawing.Size(100, 25)
$inputFileBrowseButton.Text = "Parcourir"
Set-ModernButtonStyle -Button $inputFileBrowseButton -BackColor $themeColors.Secondary -ForeColor $themeColors.TextDark
$mainPanel.Controls.Add($inputFileBrowseButton)

$columnLabel = New-Object System.Windows.Forms.Label
$columnLabel.Location = New-Object System.Drawing.Point(20, 150)
$columnLabel.Size = New-Object System.Drawing.Size(150, 20)
$columnLabel.Text = "Colonne à traiter:"
$columnLabel.Font = $fontRegular
$mainPanel.Controls.Add($columnLabel)

$columnInfo = New-Object System.Windows.Forms.Label
$columnInfo.Location = New-Object System.Drawing.Point(170, 150)
$columnInfo.Size = New-Object System.Drawing.Size(400, 20)
$columnInfo.Text = "(Sélectionnez la colonne contenant les valeurs à traiter)"
$columnInfo.Font = $fontSmall
$columnInfo.ForeColor = [System.Drawing.Color]::FromArgb(96, 94, 92)
$mainPanel.Controls.Add($columnInfo)

$columnComboBox = New-Object System.Windows.Forms.ComboBox
$columnComboBox.Location = New-Object System.Drawing.Point(20, 173)
$columnComboBox.Size = New-Object System.Drawing.Size(300, 25)
$columnComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
Set-ModernComboBoxStyle -ComboBox $columnComboBox
$mainPanel.Controls.Add($columnComboBox)
$columnComboBox.Add_SelectedIndexChanged({
    if ($columnComboBox.SelectedIndex -ge 0) {
        $statusLabel.Text = "Colonne sélectionnée : $($columnComboBox.SelectedItem). Configurez les paramètres de sécurité."
    }
})

# Séparateur 2 - Position Y = 210
$separator2 = New-Object System.Windows.Forms.Panel
$separator2.Location = New-Object System.Drawing.Point(20, 210)
$separator2.Size = New-Object System.Drawing.Size(630, 1)
$separator2.BackColor = $themeColors.Border
$mainPanel.Controls.Add($separator2)

# Section 2 - Position Y = 220
$cryptoLabel = New-Object System.Windows.Forms.Label
$cryptoLabel.Location = New-Object System.Drawing.Point(20, 220)
$cryptoLabel.Size = New-Object System.Drawing.Size(630, 20)
$cryptoLabel.Text = "2. Paramètres de sécurité"
$cryptoLabel.Font = $fontHeader
$cryptoLabel.ForeColor = $themeColors.Primary
$mainPanel.Controls.Add($cryptoLabel)

$keyLabel = New-Object System.Windows.Forms.Label
$keyLabel.Location = New-Object System.Drawing.Point(20, 245)
$keyLabel.Size = New-Object System.Drawing.Size(150, 20)
$keyLabel.Text = "Clé partagée:"
$keyLabel.Font = $fontRegular
$mainPanel.Controls.Add($keyLabel)

$keyInfo = New-Object System.Windows.Forms.Label
$keyInfo.Size = New-Object System.Drawing.Size(20, 20)
$keyInfo.Location = New-Object System.Drawing.Point(170, 245)
$keyInfo.Text = "ℹ"
$keyInfo.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
New-InfoTooltip -Control $keyInfo -Text "La clé doit être identique entre le SPC et l'Hospice Général"
$mainPanel.Controls.Add($keyInfo)

$keyTextBox = New-Object System.Windows.Forms.TextBox
$keyTextBox.Location = New-Object System.Drawing.Point(20, 268)
$keyTextBox.Size = New-Object System.Drawing.Size(400, 25)
$keyTextBox.PasswordChar = '•'
Set-ModernTextBoxStyle -TextBox $keyTextBox
Set-TextBoxPlaceholder -TextBox $keyTextBox -Text "min. 12 caractères"
$mainPanel.Controls.Add($keyTextBox)

$ivLabel = New-Object System.Windows.Forms.Label
$ivLabel.Location = New-Object System.Drawing.Point(20, 300)
$ivLabel.Size = New-Object System.Drawing.Size(200, 20)
$ivLabel.Text = "Vecteur d'initialisation (IV):"
$ivLabel.Font = $fontRegular
$mainPanel.Controls.Add($ivLabel)

$ivInfo = New-Object System.Windows.Forms.Label
$ivInfo.Size = New-Object System.Drawing.Size(20, 20)
$ivInfo.Location = New-Object System.Drawing.Point(220, 300)
$ivInfo.Text = "ℹ"
$ivInfo.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
New-InfoTooltip -Control $ivInfo -Text "Le vecteur d'initialisation doit être identique"
$mainPanel.Controls.Add($ivInfo)

$ivTextBox = New-Object System.Windows.Forms.TextBox
$ivTextBox.Location = New-Object System.Drawing.Point(20, 323)
$ivTextBox.Size = New-Object System.Drawing.Size(400, 25)
$ivTextBox.PasswordChar = '•'
Set-ModernTextBoxStyle -TextBox $ivTextBox
Set-TextBoxPlaceholder -TextBox $ivTextBox -Text "min. 8 caractères"
$mainPanel.Controls.Add($ivTextBox)

# Séparateur 3 - Position Y = 360
$separator3 = New-Object System.Windows.Forms.Panel
$separator3.Location = New-Object System.Drawing.Point(20, 360)
$separator3.Size = New-Object System.Drawing.Size(630, 1)
$separator3.BackColor = $themeColors.Border
$mainPanel.Controls.Add($separator3)

# Section 3 - Position Y = 370
$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Location = New-Object System.Drawing.Point(20, 370)
$outputLabel.Size = New-Object System.Drawing.Size(630, 20)
$outputLabel.Text = "3. Fichier de sortie"
$outputLabel.Font = $fontHeader
$outputLabel.ForeColor = $themeColors.Primary
$mainPanel.Controls.Add($outputLabel)

$outputFileLabel = New-Object System.Windows.Forms.Label
$outputFileLabel.Location = New-Object System.Drawing.Point(20, 395)
$outputFileLabel.Size = New-Object System.Drawing.Size(150, 20)
$outputFileLabel.Text = "Fichier de sortie:"
$outputFileLabel.Font = $fontRegular
$mainPanel.Controls.Add($outputFileLabel)

$outputFileTextBox = New-Object System.Windows.Forms.TextBox
$outputFileTextBox.Location = New-Object System.Drawing.Point(20, 418)
$outputFileTextBox.Size = New-Object System.Drawing.Size(480, 25)
$outputFileTextBox.ReadOnly = $true
Set-ModernTextBoxStyle -TextBox $outputFileTextBox
$mainPanel.Controls.Add($outputFileTextBox)

$outputFileBrowseButton = New-Object System.Windows.Forms.Button
$outputFileBrowseButton.Location = New-Object System.Drawing.Point(510, 418)
$outputFileBrowseButton.Size = New-Object System.Drawing.Size(100, 25)
$outputFileBrowseButton.Text = "Parcourir"
Set-ModernButtonStyle -Button $outputFileBrowseButton -BackColor $themeColors.Secondary -ForeColor $themeColors.TextDark
$mainPanel.Controls.Add($outputFileBrowseButton)

# Section 4 - Position Y = 455
$modeGroupBox = New-Object System.Windows.Forms.GroupBox
$modeGroupBox.Location = New-Object System.Drawing.Point(20, 455)
$modeGroupBox.Size = New-Object System.Drawing.Size(250, 60)
$modeGroupBox.Text = "Mode"
$modeGroupBox.Font = $fontRegular
$modeGroupBox.ForeColor = $themeColors.TextDark
$mainPanel.Controls.Add($modeGroupBox)

$encryptRadioButton = New-Object System.Windows.Forms.RadioButton
$encryptRadioButton.Location = New-Object System.Drawing.Point(15, 25)
$encryptRadioButton.Size = New-Object System.Drawing.Size(80, 25)
$encryptRadioButton.Text = "Crypter"
$encryptRadioButton.Checked = $true
$encryptRadioButton.Font = $fontRegular
$modeGroupBox.Controls.Add($encryptRadioButton)

$decryptRadioButton = New-Object System.Windows.Forms.RadioButton
$decryptRadioButton.Location = New-Object System.Drawing.Point(100, 25)
$decryptRadioButton.Size = New-Object System.Drawing.Size(100, 25)
$decryptRadioButton.Text = "Décrypter"
$decryptRadioButton.Font = $fontRegular
$modeGroupBox.Controls.Add($decryptRadioButton)

$processButton = New-Object System.Windows.Forms.Button
$processButton.Location = New-Object System.Drawing.Point(380, 470)
$processButton.Size = New-Object System.Drawing.Size(120, 35)
$processButton.Text = "Traiter"
Set-ModernButtonStyle -Button $processButton -BackColor $themeColors.Primary -ForeColor $themeColors.TextLight -IsPrimary
$mainPanel.Controls.Add($processButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(510, 470)
$cancelButton.Size = New-Object System.Drawing.Size(100, 35)
$cancelButton.Text = "Fermer"
Set-ModernButtonStyle -Button $cancelButton -BackColor $themeColors.Secondary -ForeColor $themeColors.TextDark
$mainPanel.Controls.Add($cancelButton)

# Indicateur de progression
$progressPanel = New-Object System.Windows.Forms.Panel
$progressPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
$progressPanel.Height = 45  # Plus compact
$progressPanel.Visible = $false
$form.Controls.Add($progressPanel)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Dock = [System.Windows.Forms.DockStyle]::Top
$progressBar.Height = 20    # Plus fin
$progressBar.Margin = New-Object System.Windows.Forms.Padding(10,5,10,5)
$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
$progressBar.MarqueeAnimationSpeed = 30
$progressPanel.Controls.Add($progressBar)

$progressLabel = New-Object System.Windows.Forms.Label
$progressLabel.Location = New-Object System.Drawing.Point(0, 25)  # Position ajustée
$progressLabel.Size = New-Object System.Drawing.Size(630, 20)
$progressLabel.Text = "Traitement en cours..."
$progressLabel.Font = $fontSmall
$progressLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$progressPanel.Controls.Add($progressLabel)

# Barre d'état
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Prêt"
$statusStrip.Items.Add($statusLabel)
$form.Controls.Add($statusStrip)

# Initialiser la taille des contrôles
Update-ControlSizes -FormWidth $form.ClientSize.Width

# Gestion du redimensionnement
$form.Add_Resize({
    Update-ControlSizes -FormWidth $form.ClientSize.Width
    $form.Refresh()
})

# Événements
$inputFileBrowseButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        if ($global:ExcelAvailable) {
            $openFileDialog.Filter = "Fichiers de données (*.csv;*.xlsx;*.xls)|*.csv;*.xlsx;*.xls|Fichiers CSV (*.csv)|*.csv|Fichiers Excel (*.xlsx;*.xls)|*.xlsx;*.xls|Tous les fichiers (*.*)|*.*"
        } else {
            $openFileDialog.Filter = "Fichiers CSV (*.csv)|*.csv|Tous les fichiers (*.*)|*.*"
        }
        $openFileDialog.Title = "Sélectionner le fichier d'entrée"
        
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $inputFileTextBox.Text = $openFileDialog.FileName
            $statusLabel.Text = "Analyse du fichier en cours..."
            
            # Désactiver les contrôles pendant l'analyse
            Set-ControlState -Enabled $false -Controls @($inputFileBrowseButton,$columnComboBox,$outputFileBrowseButton,$processButton)
            
            # Montrer le panneau de progression
            $progressPanel.Visible = $true
            $progressLabel.Text = "Analyse du fichier en cours..."
            $form.Refresh()
            
            try {
                # Récupérer et charger les colonnes du fichier
                $columns = Get-FileColumns -FilePath $openFileDialog.FileName
                
                if ($columns -ne $null) {
                    $columnComboBox.Items.Clear()
                    foreach ($column in $columns) {
                        $columnComboBox.Items.Add($column)
                    }
                    
                    # Essayer de détecter automatiquement la colonne NNSS/BNF
                    $nnssColumnIndex = -1

                    for ($i = 0; $i -lt $columns.Count; $i++) {
                        if ($columns[$i] -match "NNSS|NAVS|AVS|NSS|no_avs|numero_avs|BNF") {
                            $nnssColumnIndex = $i
                            break
                        }
                    }
                    
                    if ($nnssColumnIndex -ge 0) {
                        $columnComboBox.SelectedIndex = $nnssColumnIndex
                    } elseif ($columnComboBox.Items.Count -gt 0) {
                        $columnComboBox.SelectedIndex = 0
                    }

                    if ($columnComboBox.SelectedIndex -lt 0 -and $columnComboBox.Items.Count -gt 0) {
                        $columnComboBox.SelectedIndex = 0
                    }
                    
                    # Suggérer un nom de fichier de sortie
                    $outputPath = [System.IO.Path]::GetDirectoryName($openFileDialog.FileName)
                    $fileNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($openFileDialog.FileName)
                    $extension = [System.IO.Path]::GetExtension($openFileDialog.FileName)
                    
                    $operation = if ($encryptRadioButton.Checked) { "crypte" } else { "decrypte" }
                    $outputFileTextBox.Text = [System.IO.Path]::Combine($outputPath, "$fileNameWithoutExt`_$($operation)$extension")
                    
                    $statusLabel.Text = "Fichier chargé avec succès. Sélectionnez la colonne à traiter et configurez les paramètres de sécurité."
                }
            }
            catch {
                [System.Windows.MessageBox]::Show("Erreur lors de l'analyse du fichier: $_", "Erreur", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                $statusLabel.Text = "Erreur lors de l'analyse du fichier."
            }
            finally {
                # Réactiver les contrôles
                Set-ControlState -Enabled $true -Controls @($inputFileBrowseButton,$columnComboBox,$outputFileBrowseButton,$processButton)
                
                # Cacher le panneau de progression
                $progressPanel.Visible = $false
            }
        }
    })

$outputFileBrowseButton.Add_Click({
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $extension = [System.IO.Path]::GetExtension($inputFileTextBox.Text).ToLower()
        
        if ($extension -eq ".csv") {
            $saveFileDialog.Filter = "Fichier CSV (*.csv)|*.csv"
        }
        elseif ($extension -eq ".xlsx") {
            $saveFileDialog.Filter = "Fichier Excel (*.xlsx)|*.xlsx"
        }
        elseif ($extension -eq ".xls") {
            $saveFileDialog.Filter = "Ancien format Excel (*.xls)|*.xls"
        }
        else {
            $saveFileDialog.Filter = "Tous les fichiers (*.*)|*.*"
        }
        
        $saveFileDialog.Title = "Enregistrer le fichier de sortie"
        
        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $outputFileTextBox.Text = $saveFileDialog.FileName
        }
    })

$encryptRadioButton.Add_CheckedChanged({
        if ($encryptRadioButton.Checked -and ![string]::IsNullOrEmpty($inputFileTextBox.Text)) {
            # Mettre à jour le nom du fichier de sortie lors du changement de mode
            $outputPath = [System.IO.Path]::GetDirectoryName($inputFileTextBox.Text)
            $fileNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($inputFileTextBox.Text)
            $extension = [System.IO.Path]::GetExtension($inputFileTextBox.Text)
            $outputFileTextBox.Text = [System.IO.Path]::Combine($outputPath, "$fileNameWithoutExt`_crypte$extension")
        }
    })

$decryptRadioButton.Add_CheckedChanged({
        if ($decryptRadioButton.Checked -and ![string]::IsNullOrEmpty($inputFileTextBox.Text)) {
            # Mettre à jour le nom du fichier de sortie lors du changement de mode
            $outputPath = [System.IO.Path]::GetDirectoryName($inputFileTextBox.Text)
            $fileNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($inputFileTextBox.Text)
            $extension = [System.IO.Path]::GetExtension($inputFileTextBox.Text)
            $outputFileTextBox.Text = [System.IO.Path]::Combine($outputPath, "$fileNameWithoutExt`_decrypte$extension")
        }
    })

$processButton.Add_Click({
        # Valider les entrées
        if ([string]::IsNullOrEmpty($inputFileTextBox.Text)) {
            [System.Windows.MessageBox]::Show("Veuillez sélectionner un fichier d'entrée.", "Champ manquant", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if ([string]::IsNullOrEmpty($outputFileTextBox.Text)) {
            [System.Windows.MessageBox]::Show("Veuillez sélectionner un fichier de sortie.", "Champ manquant", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if ($columnComboBox.SelectedIndex -lt 0) {
            [System.Windows.MessageBox]::Show("Veuillez sélectionner une colonne à traiter.", "Champ manquant", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if ([string]::IsNullOrEmpty($keyTextBox.Text)) {
            [System.Windows.MessageBox]::Show("Veuillez saisir une clé de cryptage.", "Champ manquant", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if ([string]::IsNullOrEmpty($ivTextBox.Text)) {
            [System.Windows.MessageBox]::Show("Veuillez saisir un vecteur d'initialisation (IV).", "Champ manquant", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        # Vérification de force minimale
        if (($keyTextBox.Text).Length -lt 12) {
            [System.Windows.MessageBox]::Show("La clé de cryptage doit contenir au moins 12 caractères pour une sécurité suffisante.", "Clé insuffisante", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if (($ivTextBox.Text).Length -lt 8) {
            [System.Windows.MessageBox]::Show("Le vecteur d'initialisation (IV) doit contenir au moins 8 caractères pour une sécurité suffisante.", "IV insuffisant", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        # Obtenir les paramètres de cryptage
        $cryptoParams = ConvertTo-SecureAESKey -KeyString $keyTextBox.Text -IVString $ivTextBox.Text
        
        # Configurer le traitement
        $inputFilePath = $inputFileTextBox.Text
        $outputFilePath = $outputFileTextBox.Text
        $columnName = $columnComboBox.SelectedItem.ToString()
        $isEncryption = $encryptRadioButton.Checked
        
        # Vérifier si le fichier de sortie existe déjà
        if (Test-Path -Path $outputFilePath) {
            $confirmation = [System.Windows.MessageBox]::Show("Le fichier de sortie existe déjà. Voulez-vous le remplacer?", "Confirmation", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
            if ($confirmation -eq [System.Windows.Forms.DialogResult]::No) {
                return
            }
        }
        
        # Désactiver les contrôles pendant le traitement
        Set-ControlState -Enabled $false -Controls @($inputFileBrowseButton,$columnComboBox,$outputFileBrowseButton,$processButton,$cancelButton,$keyTextBox,$ivTextBox,$encryptRadioButton,$decryptRadioButton)
        
        # Montrer le panneau de progression
        $progressPanel.Visible = $true
        $operation = if ($isEncryption) { "cryptage" } else { "décryptage" }
        $progressLabel.Text = "$operation en cours..."
        $statusLabel.Text = "Traitement en cours..."
        $form.Refresh()
        
        # Traiter le fichier de façon synchrone
        $extension = [System.IO.Path]::GetExtension($inputFilePath).ToLower()

        try {
            if ($extension -eq ".csv") {
                $result = Process-CSVFile -InputFilePath $inputFilePath -OutputFilePath $outputFilePath -ColumnName $columnName -Key $cryptoParams.Key -IV $cryptoParams.IV -IsEncryption $isEncryption
            }
            elseif ($extension -eq ".xlsx" -or $extension -eq ".xls") {
                $result = Process-ExcelFile -InputFilePath $inputFilePath -OutputFilePath $outputFilePath -ColumnName $columnName -Key $cryptoParams.Key -IV $cryptoParams.IV -IsEncryption $isEncryption
            }
            else {
                throw "Format de fichier non pris en charge."
            }

            if ($result) {
                $operation = if ($isEncryption) { "cryptage" } else { "décryptage" }
                [System.Windows.MessageBox]::Show("Le $operation a été effectué avec succès.", "Succès", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                $statusLabel.Text = "Traitement terminé avec succès."
            }
        }
        catch {
            [System.Windows.MessageBox]::Show("Erreur lors du traitement: $_", "Erreur", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            $statusLabel.Text = "Erreur lors du traitement."
        }
        finally {
            # Réactiver les contrôles
            Set-ControlState -Enabled $true -Controls @($inputFileBrowseButton,$columnComboBox,$outputFileBrowseButton,$processButton,$cancelButton,$keyTextBox,$ivTextBox,$encryptRadioButton,$decryptRadioButton)
            $progressPanel.Visible = $false
        }
    })

$cancelButton.Add_Click({
        $form.Close()
    })

# Afficher le formulaire
$form.ShowDialog() | Out-Null
