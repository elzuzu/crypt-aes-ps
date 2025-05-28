# Script de Cryptage/Décryptage des NNSS pour l'échange SPC-Hospice Général
# Auteur: Alexandre Wyer
# Date: Mai 2025
# Version: 1.0
#
# Ce script permet de crypter ou décrypter les numéros NNSS dans un fichier CSV ou Excel
# en utilisant l'algorithme AES-256-CBC avec une clé et un vecteur d'initialisation partagés.

# Vérifier la version de PowerShell
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Error "Ce script nécessite PowerShell 5 ou supérieur." -Category InvalidArgument
    return
}

# Charger les assemblies nécessaires
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
Add-Type -AssemblyName System.Windows.Forms.DataVisualization
# Charger les fonctions de traitement
. (Join-Path $PSScriptRoot "functions/crypt-functions.ps1")
. (Join-Path $PSScriptRoot "functions/ui-styles.ps1")

# Fonction de redimensionnement intelligente pour les contrôles
function Update-ControlSizes {
    param($FormWidth)

    # Ajuster la largeur des textbox selon la taille de la fenêtre
    $textBoxWidth = [Math]::Max(300, $FormWidth - 200)
    $buttonX = $textBoxWidth + 10

    # Appliquer aux contrôles principaux
    if($inputFileTextBox){ $inputFileTextBox.Width = $textBoxWidth }
    if($outputFileTextBox){ $outputFileTextBox.Width = $textBoxWidth }
    if($keyTextBox){ $keyTextBox.Width = [Math]::Min(400, $textBoxWidth - 50) }
    if($ivTextBox){ $ivTextBox.Width = [Math]::Min(400, $textBoxWidth - 50) }

    # Repositionner les boutons Parcourir
    if($inputFileBrowseButton){ $inputFileBrowseButton.Left = $buttonX }
    if($outputFileBrowseButton){ $outputFileBrowseButton.Left = $buttonX }

    # Repositionner les boutons d'action (nouveau)
    if($processButton){ $processButton.Left = [Math]::Max(380, $FormWidth - 260) }
    if($cancelButton){ $cancelButton.Left = [Math]::Max(510, $FormWidth - 130) }
}

# Créer l'interface utilisateur
$form = New-Object System.Windows.Forms.Form
$form.Text = "SPC - Cryptage/Décryptage des NNSS"
$form.Size = New-Object System.Drawing.Size(700, 720)
$form.MinimumSize = New-Object System.Drawing.Size(650, 600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.MaximizeBox = $true
$form.AutoScroll = $true
$form.BackColor = $themeColors.Background
$form.ForeColor = $themeColors.TextDark
$form.Font = $fontRegular

$form.Padding = New-Object System.Windows.Forms.Padding(24)

# Créons un panel principal adaptatif
$mainPanel = New-Object System.Windows.Forms.TableLayoutPanel
$mainPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$mainPanel.ColumnCount = 1
$mainPanel.RowCount = 6
$mainPanel.Padding = New-Object System.Windows.Forms.Padding(20)
$mainPanel.AutoScroll = $true
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 80)))   # Header
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 140)))  # Section 1
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 160)))  # Section 2  
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 100)))  # Section 3
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 80)))   # Section 4
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))   # Espace restant
$form.Controls.Add($mainPanel)

# Logo et titre
$logoPanel = New-Object System.Windows.Forms.Panel
$logoPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$logoPanel.BackColor = [System.Drawing.Color]::Transparent
$mainPanel.Controls.Add($logoPanel,0,0)

$logoImage = New-Object System.Windows.Forms.PictureBox
$logoImage.Size = New-Object System.Drawing.Size(50, 50)
$logoImage.Location = New-Object System.Drawing.Point(24, 10)
$logoImage.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$logoImage.BackColor = [System.Drawing.Color]::Transparent

# Icône de cadenas (simulée par un label)
$lockIcon = New-Object System.Windows.Forms.Label
$lockIcon.Text = "🔒"  # Symbole cadenas Unicode
$lockIcon.Size = New-Object System.Drawing.Size(50, 50)
$lockIcon.Location = New-Object System.Drawing.Point(24, 10)
$lockIcon.Font = New-Object System.Drawing.Font("Segoe UI", 24)
$lockIcon.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$lockIcon.ForeColor = $themeColors.Primary
$logoPanel.Controls.Add($lockIcon)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(80, 10)
$titleLabel.Size = New-Object System.Drawing.Size(570, 35)
$titleLabel.Text = "Cryptage/Décryptage des NNSS"
$titleLabel.Font = $fontTitle
$titleLabel.ForeColor = $themeColors.TextDark
$logoPanel.Controls.Add($titleLabel)

$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Location = New-Object System.Drawing.Point(80, 45)
$subtitleLabel.Size = New-Object System.Drawing.Size(570, 20)
$subtitleLabel.Text = "Outil pour l'échange sécurisé de données SPC-Hospice Général"
$subtitleLabel.Font = $fontSmall
$subtitleLabel.ForeColor = [System.Drawing.Color]::FromArgb(96, 94, 92)
$logoPanel.Controls.Add($subtitleLabel)

# Section 1 - conteneur
$section1Panel = New-Object System.Windows.Forms.Panel
$section1Panel.Dock = [System.Windows.Forms.DockStyle]::Fill
$section1Panel.Margin = New-Object System.Windows.Forms.Padding(0,5,0,5)
$mainPanel.Controls.Add($section1Panel,0,1)

# Séparateur
$separator1 = New-Object System.Windows.Forms.Panel
$separator1.Location = New-Object System.Drawing.Point(0, 0)
$separator1.Size = New-Object System.Drawing.Size(630, 1)
$separator1.BackColor = $themeColors.Border
$section1Panel.Controls.Add($separator1)

# Section 1: Sélection du fichier
$fileSelectionLabel = New-Object System.Windows.Forms.Label
$fileSelectionLabel.Location = New-Object System.Drawing.Point(0, 5)
$fileSelectionLabel.Size = New-Object System.Drawing.Size(630, 25)
$fileSelectionLabel.Text = "1. Sélection du fichier"
$fileSelectionLabel.Font = $fontHeader
$fileSelectionLabel.ForeColor = $themeColors.Primary
$section1Panel.Controls.Add($fileSelectionLabel)

# Sélection du fichier d'entrée
$inputFilePanel = New-Object System.Windows.Forms.Panel
$inputFilePanel.Location = New-Object System.Drawing.Point(0, 34)
$inputFilePanel.Size = New-Object System.Drawing.Size(630, 60)
$inputFilePanel.BackColor = [System.Drawing.Color]::Transparent
$section1Panel.Controls.Add($inputFilePanel)

$inputFileLabel = New-Object System.Windows.Forms.Label
$inputFileLabel.Location = New-Object System.Drawing.Point(0, 0)
$inputFileLabel.Size = New-Object System.Drawing.Size(200, 25)
$inputFileLabel.Text = "Fichier d'entrée:"
$inputFileLabel.Font = $fontRegular
$inputFilePanel.Controls.Add($inputFileLabel)

$inputFileInfo = New-Object System.Windows.Forms.Label
$inputFileInfo.Location = New-Object System.Drawing.Point(200, 0)
$inputFileInfo.Size = New-Object System.Drawing.Size(430, 25)
$inputFileInfo.Text = "(CSV ou Excel contenant les NNSS à traiter)"
$inputFileInfo.Font = $fontSmall
$inputFileInfo.ForeColor = [System.Drawing.Color]::FromArgb(96, 94, 92)
$inputFilePanel.Controls.Add($inputFileInfo)

$inputFileTextBox = New-Object System.Windows.Forms.TextBox
$inputFileTextBox.Location = New-Object System.Drawing.Point(0, 30)
$inputFileTextBox.Size = New-Object System.Drawing.Size(480, 30)
$inputFileTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$inputFileTextBox.ReadOnly = $true
Set-ModernTextBoxStyle -TextBox $inputFileTextBox
$inputFilePanel.Controls.Add($inputFileTextBox)

$inputFileBrowseButton = New-Object System.Windows.Forms.Button
$inputFileBrowseButton.Location = New-Object System.Drawing.Point(490, 30)
$inputFileBrowseButton.Size = New-Object System.Drawing.Size(100, 30)
$inputFileBrowseButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$inputFileBrowseButton.Text = "Parcourir"
Set-ModernButtonStyle -Button $inputFileBrowseButton -BackColor $themeColors.Secondary -ForeColor $themeColors.TextDark
$inputFilePanel.Controls.Add($inputFileBrowseButton)

# Sélection de la colonne à traiter
$columnPanel = New-Object System.Windows.Forms.Panel
$columnPanel.Location = New-Object System.Drawing.Point(0, 98)
$columnPanel.Size = New-Object System.Drawing.Size(630, 60)
$columnPanel.BackColor = [System.Drawing.Color]::Transparent
$section1Panel.Controls.Add($columnPanel)

# Section 2 - conteneur
$section2Panel = New-Object System.Windows.Forms.Panel
$section2Panel.Dock = [System.Windows.Forms.DockStyle]::Fill
$section2Panel.Margin = New-Object System.Windows.Forms.Padding(0,5,0,5)
$mainPanel.Controls.Add($section2Panel,0,2)
$columnLabel = New-Object System.Windows.Forms.Label
$columnLabel.Location = New-Object System.Drawing.Point(0, 0)
$columnLabel.Size = New-Object System.Drawing.Size(250, 25)
$columnLabel.Text = "Colonne à traiter:"
$columnLabel.Font = $fontRegular
$columnPanel.Controls.Add($columnLabel)

$columnInfo = New-Object System.Windows.Forms.Label
$columnInfo.Location = New-Object System.Drawing.Point(250, 0)
$columnInfo.Size = New-Object System.Drawing.Size(380, 25)
$columnInfo.Text = "(Sélectionnez la colonne contenant les valeurs à traiter)"
$columnInfo.Font = $fontSmall
$columnInfo.ForeColor = [System.Drawing.Color]::FromArgb(96, 94, 92)
$columnPanel.Controls.Add($columnInfo)

$columnComboBox = New-Object System.Windows.Forms.ComboBox
$columnComboBox.Location = New-Object System.Drawing.Point(0, 30)
$columnComboBox.Size = New-Object System.Drawing.Size(300, 30)
$columnComboBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$columnComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
Set-ModernComboBoxStyle -ComboBox $columnComboBox
$columnPanel.Controls.Add($columnComboBox)
$columnComboBox.Add_SelectedIndexChanged({
    if ($columnComboBox.SelectedIndex -ge 0) {
        $statusLabel.Text = "Colonne sélectionnée : $($columnComboBox.SelectedItem). Configurez les paramètres de sécurité."
    }
})

# Séparateur
$separator2 = New-Object System.Windows.Forms.Panel
$separator2.Location = New-Object System.Drawing.Point(0, 0)
$separator2.Size = New-Object System.Drawing.Size(630, 1)
$separator2.BackColor = $themeColors.Border
$section2Panel.Controls.Add($separator2)

# Section 2: Paramètres de cryptage
$cryptoLabel = New-Object System.Windows.Forms.Label
$cryptoLabel.Location = New-Object System.Drawing.Point(0, 5)
$cryptoLabel.Size = New-Object System.Drawing.Size(630, 25)
$cryptoLabel.Text = "2. Paramètres de sécurité"
$cryptoLabel.Font = $fontHeader
$cryptoLabel.ForeColor = $themeColors.Primary
$section2Panel.Controls.Add($cryptoLabel)

# Groupbox pour les paramètres de cryptage
$cryptoPanel = New-Object System.Windows.Forms.Panel
$cryptoPanel.Location = New-Object System.Drawing.Point(0, 34)
$cryptoPanel.Size = New-Object System.Drawing.Size(630, 130)
$cryptoPanel.BackColor = [System.Drawing.Color]::Transparent
$section2Panel.Controls.Add($cryptoPanel)

# Clé de cryptage
$keyLabel = New-Object System.Windows.Forms.Label
$keyLabel.Location = New-Object System.Drawing.Point(0, 5)
$keyLabel.Size = New-Object System.Drawing.Size(200, 25)
$keyLabel.Text = "Clé partagée:"
$keyLabel.Font = $fontRegular
$cryptoPanel.Controls.Add($keyLabel)

$keyInfo = New-Object System.Windows.Forms.Label
$keyInfo.Size = New-Object System.Drawing.Size(20, 20)
$keyInfo.Location = New-Object System.Drawing.Point(200, 5)
$keyInfo.BackColor = [System.Drawing.Color]::Transparent
$keyInfo.Text = "ℹ"
$keyInfo.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
New-InfoTooltip -Control $keyInfo -Text "La clé doit être identique entre le SPC et l'Hospice Général pour assurer la compatibilité du cryptage/décryptage"
$cryptoPanel.Controls.Add($keyInfo)

$keyTextBox = New-Object System.Windows.Forms.TextBox
$keyTextBox.Location = New-Object System.Drawing.Point(0, 30)
$keyTextBox.Size = New-Object System.Drawing.Size(420, 30)
$keyTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$keyTextBox.PasswordChar = '•'
Set-ModernTextBoxStyle -TextBox $keyTextBox
Set-TextBoxPlaceholder -TextBox $keyTextBox -Text "min. 12 caractères"
$cryptoPanel.Controls.Add($keyTextBox)

# Vecteur d'initialisation (IV)
$ivLabel = New-Object System.Windows.Forms.Label
$ivLabel.Location = New-Object System.Drawing.Point(0, 70)
$ivLabel.Size = New-Object System.Drawing.Size(250, 25)
$ivLabel.Text = "Vecteur d'initialisation (IV):"
$ivLabel.Font = $fontRegular
$cryptoPanel.Controls.Add($ivLabel)

$ivInfo = New-Object System.Windows.Forms.Label
$ivInfo.Size = New-Object System.Drawing.Size(20, 20)
$ivInfo.Location = New-Object System.Drawing.Point(250, 70)
$ivInfo.BackColor = [System.Drawing.Color]::Transparent
$ivInfo.Text = "ℹ"
$ivInfo.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
New-InfoTooltip -Control $ivInfo -Text "Le vecteur d'initialisation doit également être identique entre le SPC et l'Hospice Général"
$cryptoPanel.Controls.Add($ivInfo)

$ivTextBox = New-Object System.Windows.Forms.TextBox
$ivTextBox.Location = New-Object System.Drawing.Point(0, 95)
$ivTextBox.Size = New-Object System.Drawing.Size(420, 30)
$ivTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$ivTextBox.PasswordChar = '•'
Set-ModernTextBoxStyle -TextBox $ivTextBox
Set-TextBoxPlaceholder -TextBox $ivTextBox -Text "min. 8 caractères"
$cryptoPanel.Controls.Add($ivTextBox)

# Section 3 - conteneur
$section3Panel = New-Object System.Windows.Forms.Panel
$section3Panel.Dock = [System.Windows.Forms.DockStyle]::Fill
$section3Panel.Margin = New-Object System.Windows.Forms.Padding(0,5,0,5)
$mainPanel.Controls.Add($section3Panel,0,3)
# Séparateur
$separator3 = New-Object System.Windows.Forms.Panel
$separator3.Location = New-Object System.Drawing.Point(0, 0)
$separator3.Size = New-Object System.Drawing.Size(630, 1)
$separator3.BackColor = $themeColors.Border
$section3Panel.Controls.Add($separator3)

# Section 3: Fichier de sortie
$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Location = New-Object System.Drawing.Point(0, 5)
$outputLabel.Size = New-Object System.Drawing.Size(630, 25)
$outputLabel.Text = "3. Fichier de sortie"
$outputLabel.Font = $fontHeader
$outputLabel.ForeColor = $themeColors.Primary
$section3Panel.Controls.Add($outputLabel)

# Sélection du fichier de sortie
$outputFilePanel = New-Object System.Windows.Forms.Panel
$outputFilePanel.Location = New-Object System.Drawing.Point(0, 34)
$outputFilePanel.Size = New-Object System.Drawing.Size(630, 60)
$outputFilePanel.BackColor = [System.Drawing.Color]::Transparent
$section3Panel.Controls.Add($outputFilePanel)

$outputFileLabel = New-Object System.Windows.Forms.Label
$outputFileLabel.Location = New-Object System.Drawing.Point(0, 0)
$outputFileLabel.Size = New-Object System.Drawing.Size(200, 25)
$outputFileLabel.Text = "Fichier de sortie:"
$outputFileLabel.Font = $fontRegular
$outputFilePanel.Controls.Add($outputFileLabel)

$outputFileTextBox = New-Object System.Windows.Forms.TextBox
$outputFileTextBox.Location = New-Object System.Drawing.Point(0, 30)
$outputFileTextBox.Size = New-Object System.Drawing.Size(480, 30)
$outputFileTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$outputFileTextBox.ReadOnly = $true
Set-ModernTextBoxStyle -TextBox $outputFileTextBox
$outputFilePanel.Controls.Add($outputFileTextBox)

$outputFileBrowseButton = New-Object System.Windows.Forms.Button
$outputFileBrowseButton.Location = New-Object System.Drawing.Point(490, 30)
$outputFileBrowseButton.Size = New-Object System.Drawing.Size(100, 30)
$outputFileBrowseButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$outputFileBrowseButton.Text = "Parcourir"
Set-ModernButtonStyle -Button $outputFileBrowseButton -BackColor $themeColors.Secondary -ForeColor $themeColors.TextDark
$outputFilePanel.Controls.Add($outputFileBrowseButton)

# Séparateur
$separator4 = New-Object System.Windows.Forms.Panel
$separator4.Location = New-Object System.Drawing.Point(0, 98)
$separator4.Size = New-Object System.Drawing.Size(630, 1)
$separator4.BackColor = $themeColors.Border
$section3Panel.Controls.Add($separator4)

# Section 4: Mode et actions
$section4Panel = New-Object System.Windows.Forms.Panel
$section4Panel.Dock = [System.Windows.Forms.DockStyle]::Fill
$section4Panel.Margin = New-Object System.Windows.Forms.Padding(0, 5, 0, 5)
$mainPanel.Controls.Add($section4Panel, 0, 4)

# Mode de traitement
$modeGroupBox = New-Object System.Windows.Forms.GroupBox
$modeGroupBox.Location = New-Object System.Drawing.Point(0, 10)
$modeGroupBox.Size = New-Object System.Drawing.Size(300, 70)
$modeGroupBox.Text = "Mode"
$modeGroupBox.Font = $fontRegular
$modeGroupBox.ForeColor = $themeColors.TextDark
$section4Panel.Controls.Add($modeGroupBox)

$encryptRadioButton = New-Object System.Windows.Forms.RadioButton
$encryptRadioButton.Location = New-Object System.Drawing.Point(20, 25)
$encryptRadioButton.Size = New-Object System.Drawing.Size(120, 30)
$encryptRadioButton.Text = "Crypter"
$encryptRadioButton.Checked = $true
$encryptRadioButton.Font = $fontRegular
$modeGroupBox.Controls.Add($encryptRadioButton)

$decryptRadioButton = New-Object System.Windows.Forms.RadioButton
$decryptRadioButton.Location = New-Object System.Drawing.Point(150, 25)
$decryptRadioButton.Size = New-Object System.Drawing.Size(120, 30)
$decryptRadioButton.Text = "Décrypter"
$decryptRadioButton.Font = $fontRegular
$modeGroupBox.Controls.Add($decryptRadioButton)

# Boutons d'action avec ancrage correct
$processButton = New-Object System.Windows.Forms.Button
$processButton.Location = New-Object System.Drawing.Point(380, 25)
$processButton.Size = New-Object System.Drawing.Size(120, 40)
$processButton.Text = "Traiter"
$processButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
Set-ModernButtonStyle -Button $processButton -BackColor $themeColors.Primary -ForeColor $themeColors.TextLight -IsPrimary
$section4Panel.Controls.Add($processButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(510, 25)
$cancelButton.Size = New-Object System.Drawing.Size(120, 40)
$cancelButton.Text = "Fermer"
$cancelButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
Set-ModernButtonStyle -Button $cancelButton -BackColor $themeColors.Secondary -ForeColor $themeColors.TextDark
$section4Panel.Controls.Add($cancelButton)

# Indicateur de progression
$progressPanel = New-Object System.Windows.Forms.Panel
$progressPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
$progressPanel.Height = 60
$progressPanel.Visible = $false
$form.Controls.Add($progressPanel)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Dock = [System.Windows.Forms.DockStyle]::Top
$progressBar.Height = 25
$progressBar.Margin = New-Object System.Windows.Forms.Padding(10,5,10,5)
$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
$progressBar.MarqueeAnimationSpeed = 30
$progressPanel.Controls.Add($progressBar)

$progressLabel = New-Object System.Windows.Forms.Label
$progressLabel.Location = New-Object System.Drawing.Point(0, 35)
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
        $openFileDialog.Filter = "Fichiers de données (*.csv;*.xlsx;*.xls)|*.csv;*.xlsx;*.xls|Tous les fichiers (*.*)|*.*"
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
