<#
    Pactio · Bilan express — lanceur graphique (gratuit)
    Lancé par « Lancer le bilan.cmd ». Ouvre une fenêtre avec les options, puis lance
    Bilan-Express.ps1 (même dossier) et ouvre le rapport à la fin.
    Rien n'est modifié sur le PC : le blocage des scripts n'est levé que pour ce lancement.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$Ici     = Split-Path -Parent $PSCommandPath
$Script  = Join-Path $Ici 'Bilan-Express.ps1'
$Titre   = 'Pactio · Bilan express'
$Vert    = [System.Drawing.ColorTranslator]::FromHtml('#127A70')

function Show-Message([string]$Texte, [string]$Icone = 'Information') {
    [void][System.Windows.Forms.MessageBox]::Show($Texte, $Titre, 'OK', $Icone)
}

# --- Vérifications avant d'afficher la fenêtre -------------------------------------
if (-not (Test-Path -LiteralPath $Script)) {
    Show-Message "Le fichier « Bilan-Express.ps1 » est introuvable.`n`nIl doit se trouver dans le même dossier que « Lancer le bilan »." 'Error'
    exit 1
}
if ($Ici -match '\\AppData\\Local\\Temp\\') {
    Show-Message "Le bilan est lancé depuis l'intérieur du fichier ZIP.`n`nFaites d'abord un clic droit sur le ZIP > « Extraire tout… », puis lancez « Lancer le bilan » depuis le dossier extrait." 'Warning'
    exit 1
}

# --- Fenêtre -----------------------------------------------------------------------
$Police     = New-Object System.Drawing.Font('Segoe UI', 10)
$PoliceTitre = New-Object System.Drawing.Font('Segoe UI Semibold', 14)

$f = New-Object System.Windows.Forms.Form
$f.Text            = $Titre
$f.StartPosition   = 'CenterScreen'
$f.FormBorderStyle = 'FixedDialog'
$f.MaximizeBox     = $false
$f.MinimizeBox     = $false
$f.ClientSize      = New-Object System.Drawing.Size(440, 380)
$f.BackColor       = [System.Drawing.Color]::White
$f.Font            = $Police

$bandeau = New-Object System.Windows.Forms.Panel
$bandeau.Dock      = 'Top'
$bandeau.Height    = 64
$bandeau.BackColor = $Vert
$f.Controls.Add($bandeau)

$lblTitre = New-Object System.Windows.Forms.Label
$lblTitre.Text      = 'Bilan de santé du PC'
$lblTitre.Font      = $PoliceTitre
$lblTitre.ForeColor = [System.Drawing.Color]::White
$lblTitre.AutoSize  = $true
$lblTitre.Location  = New-Object System.Drawing.Point(20, 10)
$bandeau.Controls.Add($lblTitre)

$lblSous = New-Object System.Windows.Forms.Label
$lblSous.Text      = 'Lecture seule : rien n''est modifié sur ce PC'
$lblSous.ForeColor = [System.Drawing.Color]::White
$lblSous.AutoSize  = $true
$lblSous.Location  = New-Object System.Drawing.Point(22, 38)
$bandeau.Controls.Add($lblSous)

$lblClient = New-Object System.Windows.Forms.Label
$lblClient.Text     = 'Nom du client (affiché sur le rapport) :'
$lblClient.AutoSize = $true
$lblClient.Location = New-Object System.Drawing.Point(20, 84)
$f.Controls.Add($lblClient)

$txtClient = New-Object System.Windows.Forms.TextBox
$txtClient.Location = New-Object System.Drawing.Point(22, 108)
$txtClient.Size     = New-Object System.Drawing.Size(396, 26)
$txtClient.Text     = 'Client'
$f.Controls.Add($txtClient)

function New-Case([string]$Texte, [string]$Aide, [int]$Y, [bool]$Coche) {
    $c = New-Object System.Windows.Forms.CheckBox
    $c.Text     = $Texte
    $c.AutoSize = $true
    $c.Checked  = $Coche
    $c.Location = New-Object System.Drawing.Point(22, $Y)
    $f.Controls.Add($c)
    $a = New-Object System.Windows.Forms.Label
    $a.Text      = $Aide
    $a.ForeColor = [System.Drawing.Color]::DimGray
    $a.AutoSize  = $true
    $a.Location  = New-Object System.Drawing.Point(42, ($Y + 22))
    $f.Controls.Add($a)
    return $c
}

$cAdmin    = New-Case 'Lancer en administrateur (recommandé)' 'Bilan plus complet : BitLocker, SMBv1, santé des disques' 150 $true
$cDetaille = New-Case 'Rapport détaillé' 'Ajoute la liste de tous les contrôles, même ceux qui sont OK' 204 $false
$cDebit    = New-Case 'Ne pas mesurer le débit Internet' 'Environ 30 secondes de gagnées' 258 $false

$btnLancer = New-Object System.Windows.Forms.Button
$btnLancer.Text      = 'Lancer le bilan'
$btnLancer.Size      = New-Object System.Drawing.Size(180, 38)
$btnLancer.Location  = New-Object System.Drawing.Point(238, 322)
$btnLancer.BackColor = $Vert
$btnLancer.ForeColor = [System.Drawing.Color]::White
$btnLancer.FlatStyle = 'Flat'
$btnLancer.FlatAppearance.BorderSize = 0
$btnLancer.DialogResult = 'OK'
$f.Controls.Add($btnLancer)
$f.AcceptButton = $btnLancer

$btnAnnuler = New-Object System.Windows.Forms.Button
$btnAnnuler.Text     = 'Annuler'
$btnAnnuler.Size     = New-Object System.Drawing.Size(110, 38)
$btnAnnuler.Location = New-Object System.Drawing.Point(118, 322)
$btnAnnuler.DialogResult = 'Cancel'
$f.Controls.Add($btnAnnuler)
$f.CancelButton = $btnAnnuler

if ($f.ShowDialog() -ne 'OK') { exit 0 }

# --- Préparation du lancement ------------------------------------------------------
$client = ($txtClient.Text -replace '[\\/:*?"<>|]', '').Trim()
if (-not $client) { $client = 'Client' }
$dossier = Join-Path $Ici ('Bilan_' + $client)
$rapport = Join-Path $dossier 'Rapport_Bilan_Sante.html'

function Protect([string]$s) { "'" + ($s -replace "'", "''") + "'" }

$options = ''
if ($cDetaille.Checked) { $options += ' -Detaille' }
if ($cDebit.Checked)    { $options += ' -SansDebit' }

# Commande exécutée dans une fenêtre PowerShell visible (progression à l'écran),
# puis ouverture du rapport et attente d'Entrée avant de fermer.
$commande = @"
`$Host.UI.RawUI.WindowTitle = 'Pactio · Bilan express'
& $(Protect $Script) -Client $(Protect $client)$options
`$r = $(Protect $rapport)
if (Test-Path -LiteralPath `$r) { Invoke-Item -LiteralPath `$r }
else { Write-Host '  Le rapport n''a pas été créé : faites une capture de cette fenêtre.' -ForegroundColor Red }
Write-Host ''
Read-Host '  Appuyez sur Entrée pour fermer cette fenêtre' | Out-Null
"@
$encodee = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($commande))
$args0 = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', $encodee)

try {
    if ($cAdmin.Checked) {
        Start-Process -FilePath 'powershell.exe' -ArgumentList $args0 -Verb RunAs -WorkingDirectory $Ici
    } else {
        Start-Process -FilePath 'powershell.exe' -ArgumentList $args0 -WorkingDirectory $Ici
    }
} catch {
    # Fenêtre « Autoriser cette application à apporter des modifications ? » refusée
    $rep = [System.Windows.Forms.MessageBox]::Show("Les droits administrateur ont été refusés.`n`nLancer quand même le bilan sans ces droits ? (bilan un peu moins complet)", $Titre, 'YesNo', 'Question')
    if ($rep -eq 'Yes') { Start-Process -FilePath 'powershell.exe' -ArgumentList $args0 -WorkingDirectory $Ici }
}
