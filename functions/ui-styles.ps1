# Définition du thème et des styles pour l'interface graphique

$themeColors = @{
    Background = [System.Drawing.Color]::FromArgb(250,250,250)
    Primary    = [System.Drawing.Color]::FromArgb(0,120,215)
    Secondary  = [System.Drawing.Color]::FromArgb(240,240,240)
    TextDark   = [System.Drawing.Color]::FromArgb(32,32,32)
    TextLight  = [System.Drawing.Color]::White
    Border     = [System.Drawing.Color]::FromArgb(200,200,200)
}

$fontRegular = New-Object System.Drawing.Font('Segoe UI',9)
$fontHeader  = New-Object System.Drawing.Font('Segoe UI Semibold',10)
$fontTitle   = New-Object System.Drawing.Font('Segoe UI Semibold',14)
$fontSmall   = New-Object System.Drawing.Font('Segoe UI',8)

function Set-ModernButtonStyle {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.Button]$Button,
        [System.Drawing.Color]$BackColor,
        [System.Drawing.Color]$ForeColor,
        [switch]$IsPrimary
    )
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.BackColor = $BackColor
    $Button.ForeColor = $ForeColor
    $Button.FlatAppearance.BorderSize = 0
    if($IsPrimary){
        $Button.Font = New-Object System.Drawing.Font($fontRegular.FontFamily,9,[System.Drawing.FontStyle]::Bold)
    } else {
        $Button.Font = $fontRegular
    }
}

function Set-ModernTextBoxStyle {
    param([System.Windows.Forms.TextBox]$TextBox)
    $TextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $TextBox.BackColor   = $themeColors.Secondary
}

function Set-ModernComboBoxStyle {
    param([System.Windows.Forms.ComboBox]$ComboBox)
    $ComboBox.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $ComboBox.BackColor = $themeColors.Secondary
}

function Set-ControlState {
    param(
        [bool]$Enabled,
        [System.Windows.Forms.Control[]]$Controls
    )
    foreach($ctrl in $Controls){
        if($null -ne $ctrl){ $ctrl.Enabled = $Enabled }
    }
}

# Crée et associe une info-bulle à un contrôle
$_toolTips = @()
function New-InfoTooltip {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.Control]$Control,
        [Parameter(Mandatory)]
        [string]$Text
    )
    $tip = New-Object System.Windows.Forms.ToolTip
    $tip.SetToolTip($Control, $Text)
    $_toolTips += $tip
}

