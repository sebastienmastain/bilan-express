<#
.SYNOPSIS
    Pactio · Bilan express (gratuit) : bilan de santé d'UN PC Windows, avec un rapport client noté sur 100.

.DESCRIPTION
    À lancer sur le PC à analyser. Rien n'est modifié.
        .\Bilan-Express.ps1 -Client "Boulangerie Martin"
        .\Bilan-Express.ps1 -Client "Boulangerie Martin" -Detaille

    Vérifie : version de Windows et fin de support, âge du PC, mémoire, disques, antivirus,
    pare-feu, comptes administrateurs, comptes sans mot de passe, BitLocker, mises à jour,
    historique de fiabilité (plantages), Wi-Fi et connexion Internet (latence, pertes, débits).

    Version gratuite, limitée à un PC. Le Pack TPE complet ajoute la tournée de tous les PC du
    client (un seul rapport commun), le scan du réseau (box, NAS, imprimantes, caméras, ports à
    risque), l'option serveur et domaine, la migration de PC, votre logo et vos couleurs sur le
    rapport, et les procédures PDF illustrées :
        https://pactiofr.gumroad.com/l/pack-tpe

.PARAMETER Client
    Nom du client, affiché sur le rapport.

.PARAMETER SansDebit
    Ne mesure pas le débit Internet (gagne environ 30 secondes).

.PARAMETER Detaille
    Ajoute au rapport une annexe « Détail des vérifications » : chaque contrôle effectué, avec la
    valeur relevée, le seuil appliqué et le résultat, y compris ce qui est en ordre. Utile pour
    montrer au client tout ce qui a été vérifié.

.NOTES
    Pactio · Bilan express — v1.1 (gratuit)
    Modifications :
      1.1 — option -Detaille (annexe « Détail des vérifications ») ; historique de fiabilité de Windows
            (indice de stabilité, plantages et erreurs Windows des 30 derniers jours).
    Prérequis : Windows 10 / 11, PowerShell 5.1. Aucun module à installer.
    Droits    : fonctionne en utilisateur standard ; en administrateur, le bilan est plus complet.
    Licence   : usage libre, y compris chez vos clients. Redistribution interdite.
#>

[CmdletBinding()]
param(
    [ValidateSet('Complet', 'Poste', 'Rapport')]
    [string]$Mode = 'Complet',
    [string]$Client = 'Client',
    [string]$Prestataire = '',
    [string]$Logo,
    [string]$Couleur = '#127A70',
    [switch]$SansDebit,
    [switch]$SansReseau,
    [switch]$Detaille,
    [string]$Dossier
)

# Version gratuite : un seul PC, sans scan du réseau, sans personnalisation du rapport.
$Mode = 'Complet'; $SansReseau = $true; $Logo = $null; $Prestataire = ''; $Couleur = '#127A70'
$PactioPack = 'https://pactiofr.gumroad.com/l/pack-tpe'
$Version = '1.1'

$ErrorActionPreference = 'Continue'
$Maintenant = Get-Date
$EstAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }

if (-not $Dossier) {
    $propre = ($Client -replace '[\\/:*?"<>|]', '').Trim()
    $Dossier = Join-Path (Split-Path -Parent $PSCommandPath) ("Bilan_" + $propre)
}
New-Item -ItemType Directory -Path $Dossier -Force | Out-Null

function Etape([string]$t) { Write-Host "  > $t" -ForegroundColor Cyan }

#region ===== Constats et notation =================================================
# Chaque constat : domaine, gravité (Critique / Important / Conseil), texte client, action proposée.
$script:Constats = New-Object System.Collections.Generic.List[object]
function Add-Constat {
    param([string]$Domaine, [ValidateSet('Critique', 'Important', 'Conseil')][string]$Gravite,
          [string]$Sujet, [string]$Constat, [string]$Action, [string]$Appareil = $env:COMPUTERNAME)
    $script:Constats.Add([pscustomobject]@{ Domaine = $Domaine; Gravite = $Gravite; Sujet = $Sujet; Constat = $Constat; Action = $Action; Appareil = $Appareil })
}

# Chaque contrôle effectué, pour l'annexe « Détail des vérifications » (-Detaille) : valeur relevée, seuil, résultat.
# Sans résultat imposé, il se déduit des constats du même sujet : Critique ou Important = à corriger, Conseil = à surveiller.
# Valeur vide = donnée inaccessible : « non mesuré », jamais d'erreur.
$script:Verifs = New-Object System.Collections.Generic.List[object]
function Add-Verif {
    param([string]$Theme, [string]$Controle, [string]$Valeur, [string]$Seuil, [string[]]$Sujets = @(), [string]$Resultat)
    if (-not $Valeur) { $Valeur = '—'; if (-not $Resultat) { $Resultat = 'non mesuré' } }
    if (-not $Resultat) {
        $g = @($script:Constats | Where-Object { $Sujets -contains $_.Sujet } | ForEach-Object { $_.Gravite })
        $Resultat = if ($g -contains 'Critique' -or $g -contains 'Important') { 'à corriger' } elseif ($g.Count) { 'à surveiller' } else { 'OK' }
    }
    $script:Verifs.Add([pscustomobject]@{ Theme = $Theme; Controle = $Controle; Valeur = $Valeur; Seuil = $Seuil; Resultat = $Resultat })
}
function Format-Nb($v) { ([double]$v).ToString('0.#', [Globalization.CultureInfo]::GetCultureInfo('fr-FR')) }
#endregion

#region ===== 1. Le poste ===========================================================
function Get-BilanPoste {
    # Versions de Windows : nom lisible et fin des mises à jour (éditions Famille / Pro).
    $VersionsWindows = @{
        '19045' = @('Windows 10 22H2', '2025-10-14'); '19044' = @('Windows 10 21H2', '2023-06-13')
        '22000' = @('Windows 11 21H2', '2023-10-10'); '22621' = @('Windows 11 22H2', '2024-10-08')
        '22631' = @('Windows 11 23H2', '2025-11-11'); '26100' = @('Windows 11 24H2', '2026-10-13')
        '26200' = @('Windows 11 25H2', '2027-10-12')
    }
    $c = $script:Constats.Count
    $script:Verifs.Clear()
    $p = [ordered]@{ Nom = $env:COMPUTERNAME; Date = $Maintenant.ToString('s'); Admin = $EstAdmin }

    # Système
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
    $bios = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue
    $cpu = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
    $build = [string]$os.BuildNumber
    $edition = [string]$os.Caption
    $p.Fabricant = [string]$cs.Manufacturer; $p.Modele = [string]$cs.Model
    $p.Processeur = ([string]$cpu.Name).Trim()
    $p.MemoireGo = [math]::Round([double]$cs.TotalPhysicalMemory / 1GB)
    $p.Edition = $edition
    $p.Utilisateur = [string]$cs.UserName
    $p.Domaine = if ($cs.PartOfDomain) { [string]$cs.Domain } else { '' }

    # Serveurs : table propre (les numéros de build se recoupent : 26100 = Windows 11 24H2 ET Windows Server 2025).
    $VersionsServeurPoste = @{
        '7601' = @('Windows Server 2008 R2', '2020-01-14'); '9200' = @('Windows Server 2012', '2023-10-10')
        '9600' = @('Windows Server 2012 R2', '2023-10-10'); '14393' = @('Windows Server 2016', '2027-01-12')
        '17763' = @('Windows Server 2019', '2029-01-09'); '20348' = @('Windows Server 2022', '2031-10-14')
        '26100' = @('Windows Server 2025', '2034-10-10')
    }
    $p.EstServeur = $edition -match 'Server'
    $p.EstVM = ($p.Modele -match 'Virtual Machine|VMware|VirtualBox|KVM|QEMU|HVM domU') -or ($p.Fabricant -match 'QEMU|Xen|innotek')
    $infoV = if ($p.EstServeur) { $VersionsServeurPoste[$build] } else { $VersionsWindows[$build] }
    $entreprise = $edition -match 'Enterprise|Entreprise|Education|Éducation'
    if ($p.EstServeur -and $infoV) {
        $fin = [datetime]::ParseExact($infoV[1], 'yyyy-MM-dd', $null)
        $p.Windows = $infoV[0]; $p.FinSupport = $fin.ToString('yyyy-MM-dd'); $p.Supporte = ($Maintenant -lt $fin)
    } elseif ($p.EstServeur) {
        $p.Windows = "Windows Server (build $build)"; $p.FinSupport = ''; $p.Supporte = ([int]$build -ge 20348)
    } elseif ($infoV) {
        $fin = [datetime]::ParseExact($infoV[1], 'yyyy-MM-dd', $null)
        if ($entreprise -and $infoV[0] -like 'Windows 11*') { $fin = $fin.AddMonths(12) }
        $p.Windows = $infoV[0]; $p.FinSupport = $fin.ToString('yyyy-MM-dd'); $p.Supporte = ($Maintenant -lt $fin)
    } else {
        $p.Windows = "Windows (build $build)"; $p.FinSupport = ''; $p.Supporte = ([int]$build -ge 26200)
    }
    if (-not $p.Supporte) {
        $quand = if ($p.FinSupport) { " depuis le " + ([datetime]$p.FinSupport).ToString('dd/MM/yyyy') } else { '' }
        if ($p.EstServeur) { Add-Constat 'Serveur' 'Critique' 'Serveur hors support' "$($p.Windows) ne reçoit plus de correctifs de sécurité$quand." 'Migrer vers Windows Server 2022 ou 2025' }
        else { Add-Constat 'Postes' 'Critique' 'Windows plus mis à jour' "$($p.Windows) ne reçoit plus de correctifs de sécurité$quand." 'Mettre à niveau vers une version supportée de Windows 11, ou remplacer le poste' }
    }
    Add-Verif 'Poste' 'Version de Windows' "$($p.Windows)$(if ($p.FinSupport) { ' · correctifs jusqu''au ' + ([datetime]$p.FinSupport).ToString('dd/MM/yyyy') })" 'version encore suivie par Microsoft' 'Windows plus mis à jour', 'Serveur hors support'

    # Âge estimé d'après la génération du processeur (la date du BIOS change à chaque mise à jour).
    $p.AgeAns = $null; $annee = $null; $nomCpu = [string]$cpu.Name
    if ($nomCpu -match 'Core\(TM\)\s+i[3579]-(\d{4,5})') {
        $num = $Matches[1]; $gen = if ($num.Length -eq 5) { [int]$num.Substring(0, 2) } else { [int]$num.Substring(0, 1) }
        $annee = @{ 2 = 2011; 3 = 2012; 4 = 2013; 5 = 2015; 6 = 2015; 7 = 2016; 8 = 2017; 9 = 2018; 10 = 2019; 11 = 2020; 12 = 2021; 13 = 2022; 14 = 2023 }[$gen]
    } elseif ($nomCpu -match 'Core\(TM\)\s+(Ultra|[3579])\s') { $annee = 2023 }
    elseif ($nomCpu -match 'Ryzen\s+\d\s+(?:PRO\s+)?(\d)\d{3}') {
        $annee = @{ 1 = 2017; 2 = 2018; 3 = 2019; 4 = 2020; 5 = 2021; 6 = 2022; 7 = 2023; 8 = 2024; 9 = 2025 }[[int]$Matches[1]]
    } elseif ($nomCpu -match 'Celeron|Pentium|Atom') { $p.EntreeDeGamme = $true }
    # Machine virtuelle : le processeur vu est celui de l'hôte, l'âge n'a pas de sens.
    if ($p.EstVM) { $annee = $null; $p.EntreeDeGamme = $false }
    if ($annee) { $p.AnneeEstimee = $annee; $p.AgeAns = [math]::Max(0, $Maintenant.Year - $annee) }
    # Windows 11 exige au minimum un Intel de 8e génération ou un AMD Ryzen 2000 (processeurs de 2017-2018).
    if ($annee -and $annee -lt 2017) { Add-Constat 'Postes' 'Important' 'PC incompatible Windows 11' "Processeur de $annee (environ $($p.AgeAns) ans), non pris en charge par Windows 11." 'Remplacer le poste' }
    elseif ($p.AgeAns -ge 8) { Add-Constat 'Postes' 'Important' 'PC ancien' "Processeur de $annee (environ $($p.AgeAns) ans) : fin de vie proche, pannes et lenteurs de plus en plus probables." 'Prévoir le renouvellement' }
    elseif ($p.AgeAns -ge 6) { Add-Constat 'Postes' 'Conseil' 'PC vieillissant' "Processeur de $annee (environ $($p.AgeAns) ans)." 'Anticiper le renouvellement dans les 12 à 24 mois' }
    if ($p.EntreeDeGamme) { Add-Constat 'Postes' 'Conseil' 'Processeur d''entrée de gamme' "Processeur $($nomCpu.Trim()) : lenteurs en usage bureautique moderne." 'À remplacer en priorité lors du prochain renouvellement' }
    if ($p.EstServeur) {
        if ($p.MemoireGo -lt 4) { Add-Constat 'Serveur' 'Important' 'Mémoire du serveur insuffisante' "$($p.MemoireGo) Go de mémoire sur le serveur $($p.Nom) : lenteurs des connexions et des partages." 'Passer à 8 Go minimum (16 Go si le serveur héberge des applications)' }
    } elseif ($p.MemoireGo -lt 8) { Add-Constat 'Postes' 'Important' 'Mémoire insuffisante' "$($p.MemoireGo) Go de mémoire : lenteurs probables avec Windows 11 et les navigateurs actuels." 'Ajouter de la mémoire (16 Go recommandés) ou prévoir le renouvellement' }
    if ($p.EstVM) { Add-Verif 'Poste' 'Âge du PC' 'machine virtuelle' 'moins de 6 ans ; processeur compatible Windows 11 (2017 ou plus récent)' -Resultat 'sans objet' }
    else { Add-Verif 'Poste' 'Âge du PC' $(if ($annee) { "environ $($p.AgeAns) ans (processeur de $annee)" }) 'moins de 6 ans ; processeur compatible Windows 11 (2017 ou plus récent)' 'PC incompatible Windows 11', 'PC ancien', 'PC vieillissant' }
    Add-Verif 'Poste' 'Processeur' $nomCpu.Trim() 'hors entrée de gamme (Celeron, Pentium, Atom)' 'Processeur d''entrée de gamme'
    Add-Verif 'Poste' 'Mémoire' "$($p.MemoireGo) Go" $(if ($p.EstServeur) { '4 Go au moins (serveur)' } else { '8 Go au moins' }) 'Mémoire insuffisante', 'Mémoire du serveur insuffisante'

    # Redémarrage
    if ($os.LastBootUpTime) {
        $p.JoursSansRedemarrage = [int]($Maintenant - $os.LastBootUpTime).TotalDays
        if ($p.JoursSansRedemarrage -ge 14) { Add-Constat 'Postes' 'Conseil' 'PC rarement redémarré' "Pas de redémarrage depuis $($p.JoursSansRedemarrage) jours : les mises à jour ne s'appliquent pas complètement." 'Utiliser « Redémarrer » (et non « Arrêter ») au moins une fois par semaine' }
    }
    Add-Verif 'Poste' 'Dernier redémarrage' $(if ($null -ne $p.JoursSansRedemarrage) { "il y a $($p.JoursSansRedemarrage) jour(s)" }) 'moins de 14 jours' 'PC rarement redémarré'

    # Mises à jour
    $derniere = $null
    try { $derniere = (New-Object -ComObject Microsoft.Update.AutoUpdate).Results.LastInstallationSuccessDate } catch { }
    if ($derniere -and $derniere.Year -gt 2000) {
        $p.JoursSansMiseAJour = [int]($Maintenant - $derniere.ToLocalTime()).TotalDays
        if ($p.JoursSansMiseAJour -gt 45) { Add-Constat 'Postes' 'Important' 'Mises à jour en retard' "Aucune mise à jour Windows installée depuis $($p.JoursSansMiseAJour) jours." 'Débloquer Windows Update (script 03 du pack)' }
    }
    Add-Verif 'Poste' 'Mises à jour Windows' $(if ($null -ne $p.JoursSansMiseAJour) { "dernière installation il y a $($p.JoursSansMiseAJour) jour(s)" }) '45 jours au plus' 'Mises à jour en retard'

    # Disques
    $disques = @()
    foreach ($v in (Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction SilentlyContinue)) {
        if (-not $v.Size) { continue }
        $pct = [math]::Round(100 * $v.FreeSpace / $v.Size)
        $disques += [ordered]@{ Lecteur = $v.DeviceID; TailleGo = [math]::Round($v.Size / 1GB); LibreGo = [math]::Round($v.FreeSpace / 1GB); LibrePct = $pct }
        if ($pct -lt 10) { Add-Constat 'Postes' 'Important' 'Disque presque plein' "Lecteur $($v.DeviceID) : $([math]::Round($v.FreeSpace / 1GB)) Go libres ($pct %)." 'Nettoyer le disque ou augmenter le stockage' }
        Add-Verif 'Poste' "Espace libre sur $($v.DeviceID)" "$([math]::Round($v.FreeSpace / 1GB)) Go libres sur $([math]::Round($v.Size / 1GB)) Go ($pct %)" '10 % libres au moins' -Resultat $(if ($pct -lt 10) { 'à corriger' } else { 'OK' })
    }
    $p.Disques = $disques
    $physiques = @()
    try {
        foreach ($d in (Get-PhysicalDisk -ErrorAction Stop)) {
            $physiques += [ordered]@{ Nom = [string]$d.FriendlyName; Type = [string]$d.MediaType; Sante = [string]$d.HealthStatus; TailleGo = [math]::Round($d.Size / 1GB) }
            if ([string]$d.HealthStatus -ne 'Healthy') { Add-Constat 'Postes' 'Critique' 'Disque en mauvais état' "Le disque « $($d.FriendlyName) » signale un état « $($d.HealthStatus) » : risque de perte de données." 'Sauvegarder immédiatement et remplacer le disque' }
            if ([string]$d.MediaType -eq 'HDD' -and $d.DeviceId -eq '0') { Add-Constat 'Postes' 'Conseil' 'Disque mécanique' "Le système est sur un disque mécanique (HDD) : c'est la première cause de lenteur." 'Remplacer par un SSD (gain de vitesse immédiat, coût modéré)' }
            $resD = if ([string]$d.HealthStatus -ne 'Healthy') { 'à corriger' } elseif ([string]$d.MediaType -eq 'HDD' -and $d.DeviceId -eq '0') { 'à surveiller' } else { 'OK' }
            $seuilD = if ($d.DeviceId -eq '0') { 'état « Healthy » ; SSD pour le disque système' } else { 'état « Healthy » (disque secondaire ou externe)' }
            Add-Verif 'Poste' "État du disque « $($d.FriendlyName) »" "$($d.HealthStatus) · $(if ([string]$d.MediaType -and [string]$d.MediaType -ne 'Unspecified') { $d.MediaType } else { 'type non précisé' }) · $([math]::Round($d.Size / 1GB)) Go" $seuilD -Resultat $resD
        }
    } catch { }
    if (-not $physiques.Count) { Add-Verif 'Poste' 'État des disques' '' 'état « Healthy » ; SSD pour le disque système' }
    $p.DisquesPhysiques = $physiques

    # Antivirus
    $av = @()
    try {
        foreach ($a in (Get-CimInstance -Namespace root\SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction Stop)) {
            $hex = '{0:X6}' -f [int]$a.productState
            $actif = $hex.Substring(2, 2) -in '10', '11'
            $aJour = $hex.Substring(4, 2) -eq '00'
            $av += [ordered]@{ Nom = [string]$a.displayName; Actif = $actif; AJour = $aJour }
        }
    } catch { }
    $p.Antivirus = $av
    $avActif = @($av | Where-Object { $_.Actif })
    if ($av.Count -gt 0 -and $avActif.Count -eq 0) { Add-Constat 'Sécurité' 'Critique' 'Antivirus désactivé' 'Aucun antivirus actif sur ce poste.' 'Réactiver Microsoft Defender ou l''antivirus installé' }
    elseif (@($avActif | Where-Object { -not $_.AJour }).Count -gt 0) { Add-Constat 'Sécurité' 'Important' 'Antivirus pas à jour' "L'antivirus « $($avActif[0].Nom) » n'est pas à jour." 'Forcer la mise à jour des signatures et vérifier l''abonnement' }
    try {
        $mp = Get-MpComputerStatus -ErrorAction Stop
        $p.DefenderTempsReel = [bool]$mp.RealTimeProtectionEnabled
        $p.DefenderSignaturesJours = [int]$mp.AntivirusSignatureAge
        if ($mp.AntivirusEnabled -and $mp.AntivirusSignatureAge -gt 7) { Add-Constat 'Sécurité' 'Important' 'Signatures antivirus anciennes' "Les définitions de Microsoft Defender ont $($mp.AntivirusSignatureAge) jours." 'Vérifier les mises à jour Windows / Defender' }
    } catch { }
    Add-Verif 'Sécurité' 'Antivirus' $(if ($av.Count) { ($av | ForEach-Object { "$($_.Nom) ($(if ($_.Actif) { 'actif' } else { 'inactif' }), $(if ($_.AJour) { 'à jour' } else { 'pas à jour' }))" }) -join ' ; ' }) 'au moins un antivirus actif et à jour' 'Antivirus désactivé', 'Antivirus pas à jour'
    Add-Verif 'Sécurité' 'Microsoft Defender' $(if ($null -ne $p.DefenderSignaturesJours) { "signatures de $($p.DefenderSignaturesJours) jour(s) ; protection en temps réel $(if ($p.DefenderTempsReel) { 'active' } else { 'inactive' })" }) 'signatures de 7 jours au plus (s''il est l''antivirus actif)' 'Signatures antivirus anciennes'

    # Pare-feu
    try {
        $off = @(Get-NetFirewallProfile -ErrorAction Stop | Where-Object { -not $_.Enabled } | ForEach-Object { $_.Name })
        $p.PareFeuDesactive = $off
        if ($off.Count -gt 0) { Add-Constat 'Sécurité' 'Critique' 'Pare-feu désactivé' "Le pare-feu Windows est désactivé (profil : $($off -join ', '))." 'Réactiver le pare-feu Windows' }
    } catch { }
    Add-Verif 'Sécurité' 'Pare-feu Windows' $(if ($null -ne $p.PareFeuDesactive) { if (@($p.PareFeuDesactive).Count) { 'désactivé : ' + ($p.PareFeuDesactive -join ', ') } else { 'activé sur tous les profils' } }) 'activé sur tous les profils' 'Pare-feu désactivé'

    # Comptes
    $p.UtilisateurEstAdmin = $null
    try {
        $admins = @(Get-LocalGroupMember -SID 'S-1-5-32-544' -ErrorAction Stop | ForEach-Object { [string]$_.Name })
        $p.AdminsLocaux = $admins
        $moi = [Security.Principal.WindowsIdentity]::GetCurrent().Name
        $p.UtilisateurEstAdmin = ($admins -contains $moi)
        if ($p.UtilisateurEstAdmin -and -not $p.Domaine) { Add-Constat 'Sécurité' 'Important' 'Utilisateur administrateur' "Le compte utilisé au quotidien ($moi) est administrateur du PC : un virus aurait tous les droits." 'Créer un compte administrateur séparé et passer l''utilisateur en compte standard' }
        # Les comptes créés à l'installation de Windows ont tous « mot de passe non obligatoire » :
        # on ne signale que ceux dont le mot de passe est réellement VIDE (test d'ouverture de session
        # locale avec un mot de passe vide : 1326 = mot de passe présent ; 1327, 1385, 1330, 1907, 1793 = mot de passe
        # vide accepté mais ouverture refusée pour une autre raison : le compte est bien sans mot de passe).
        $candidats = @(Get-LocalUser -ErrorAction Stop | Where-Object { $_.Enabled -and -not $_.PasswordRequired } | ForEach-Object { $_.Name })
        $sansMdp = @()
        if ($candidats.Count -gt 0) {
            if (-not ('Pactio.MdpVide' -as [type])) {
                Add-Type -Namespace Pactio -Name MdpVide -MemberDefinition '[DllImport("advapi32.dll", SetLastError=true, CharSet=CharSet.Unicode)] public static extern bool LogonUser(string u, string d, string p, int t, int pr, out IntPtr h); [DllImport("kernel32.dll")] public static extern bool CloseHandle(IntPtr h);'
            }
            foreach ($compteLocal in $candidats) {
                $jeton = [IntPtr]::Zero
                $ouvert = [Pactio.MdpVide]::LogonUser($compteLocal, '.', '', 2, 0, [ref]$jeton)
                $codeErr = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
                if ($ouvert) { [void][Pactio.MdpVide]::CloseHandle($jeton) }
                if ($ouvert -or @(1327, 1385, 1330, 1907, 1793) -contains $codeErr) { $sansMdp += $compteLocal }
            }
        }
        if ($sansMdp.Count -gt 0) { Add-Constat 'Sécurité' 'Critique' 'Compte sans mot de passe' "Comptes sans aucun mot de passe : $($sansMdp -join ', ')." 'Définir un mot de passe ou désactiver ces comptes' }
    } catch { }
    Add-Verif 'Sécurité' 'Compte utilisé au quotidien' $(if ($null -ne $p.UtilisateurEstAdmin) { if ($p.UtilisateurEstAdmin) { 'administrateur du PC' } else { 'pas administrateur direct du PC' } }) 'compte standard (hors domaine)' 'Utilisateur administrateur'
    Add-Verif 'Sécurité' 'Comptes sans mot de passe' $(if ($null -ne $sansMdp) { "$(@($sansMdp).Count) compte(s)" }) 'aucun' 'Compte sans mot de passe'

    # Bureau à distance, UAC
    $rdp = $null; try { $rdp = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server' -Name fDenyTSConnections -ErrorAction Stop).fDenyTSConnections } catch { }
    $p.BureauADistance = ($rdp -eq 0)
    if ($rdp -eq 0) { Add-Constat 'Sécurité' 'Conseil' 'Bureau à distance activé' 'Le bureau à distance (RDP) est activé sur ce poste.' 'Le désactiver s''il ne sert pas, ou le réserver au réseau local' }
    $uac = $null; try { $uac = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name EnableLUA -ErrorAction Stop).EnableLUA } catch { }
    if ($uac -eq 0) { Add-Constat 'Sécurité' 'Critique' 'Contrôle de compte (UAC) désactivé' 'Le contrôle de compte utilisateur est désactivé : tout programme s''exécute avec les pleins droits.' 'Réactiver l''UAC' }
    Add-Verif 'Sécurité' 'Bureau à distance' $(if ($null -ne $rdp) { if ($rdp -eq 0) { 'activé' } else { 'désactivé' } }) 'désactivé s''il ne sert pas' 'Bureau à distance activé'
    Add-Verif 'Sécurité' 'Contrôle de compte (UAC)' $(if ($null -ne $uac) { if ($uac -eq 0) { 'désactivé' } else { 'activé' } }) 'activé' 'Contrôle de compte (UAC) désactivé'

    # SMBv1 et BitLocker (administrateur)
    if ($EstAdmin) {
        try { $smb1 = (Get-SmbServerConfiguration -ErrorAction Stop).EnableSMB1Protocol; $p.SMB1 = [bool]$smb1
              if ($smb1) { Add-Constat 'Sécurité' 'Critique' 'Protocole SMBv1 actif' 'Le vieux protocole de partage SMBv1, utilisé par le rançongiciel WannaCry, est actif.' 'Désactiver SMBv1 (et vérifier les vieux NAS ou imprimantes qui en dépendent)' } } catch { }
        try { $bl = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop; $p.BitLocker = [string]$bl.ProtectionStatus
              if ([string]$bl.ProtectionStatus -ne 'On' -and $os.Caption -notmatch 'Famille|Home') {
                  $grav = if ($cs.PCSystemType -eq 2) { 'Important' } else { 'Conseil' }
                  Add-Constat 'Sécurité' $grav 'Disque non chiffré' 'Le disque n''est pas chiffré (BitLocker) : en cas de vol, toutes les données sont lisibles.' 'Activer BitLocker et sauvegarder la clé de récupération' } } catch { }
    } else { $p.BitLocker = 'non vérifié (lancer en administrateur)' }
    Add-Verif 'Sécurité' 'Protocole SMBv1' $(if ($null -ne $p.SMB1) { if ($p.SMB1) { 'actif' } else { 'désactivé' } }) 'désactivé' 'Protocole SMBv1 actif'
    Add-Verif 'Sécurité' 'Chiffrement du disque (BitLocker)' $(if ($EstAdmin -and $p.BitLocker) { if ($p.BitLocker -eq 'On') { 'actif' } else { 'inactif' } }) 'actif (éditions Pro et Entreprise)' 'Disque non chiffré'

    # Batterie
    try {
        $bat = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($bat) {
            $plein = (Get-CimInstance -Namespace root\wmi -ClassName BatteryFullChargedCapacity -ErrorAction SilentlyContinue | Select-Object -First 1).FullChargedCapacity
            $origine = (Get-CimInstance -Namespace root\wmi -ClassName BatteryStaticData -ErrorAction SilentlyContinue | Select-Object -First 1).DesignedCapacity
            if ($plein -and $origine) { $p.BatterieSantePct = [math]::Min(100, [math]::Round(100 * $plein / $origine)) }
        }
    } catch { }
    if ($null -eq $p.BatterieSantePct -and $cs.PCSystemType -eq 2) {
        # Seconde méthode : le rapport batterie de Windows (powercfg).
        try {
            $xmlBat = Join-Path $env:TEMP 'pactio_batterie.xml'
            & powercfg.exe /batteryreport /xml /output $xmlBat 2>$null | Out-Null
            if (Test-Path $xmlBat) {
                [xml]$rb = Get-Content $xmlBat -Raw
                $b0 = @($rb.BatteryReport.Batteries.Battery)[0]
                if ($b0 -and [double]$b0.DesignCapacity -gt 0) { $p.BatterieSantePct = [math]::Min(100, [math]::Round(100 * [double]$b0.FullChargeCapacity / [double]$b0.DesignCapacity)) }
                Remove-Item $xmlBat -Force -ErrorAction SilentlyContinue
            }
        } catch { }
    }
    if ($p.BatterieSantePct -and $p.BatterieSantePct -lt 60) { Add-Constat 'Postes' 'Conseil' 'Batterie usée' "La batterie ne tient plus que $($p.BatterieSantePct) % de sa capacité d'origine." 'Remplacer la batterie' }
    if ($null -ne $p.BatterieSantePct) { Add-Verif 'Poste' 'Batterie' "$($p.BatterieSantePct) % de la capacité d'origine" '60 % au moins' 'Batterie usée' }
    elseif ($cs.PCSystemType -eq 2) { Add-Verif 'Poste' 'Batterie' '' '60 % au moins' }

    # Démarrage
    try { $p.ProgrammesDemarrage = @(Get-CimInstance Win32_StartupCommand -ErrorAction Stop).Count
          if ($p.ProgrammesDemarrage -gt 12) { Add-Constat 'Postes' 'Conseil' 'Démarrage chargé' "$($p.ProgrammesDemarrage) programmes se lancent au démarrage : démarrage lent." 'Faire le tri dans les programmes au démarrage' } } catch { }
    Add-Verif 'Poste' 'Programmes au démarrage' $(if ($null -ne $p.ProgrammesDemarrage) { "$($p.ProgrammesDemarrage)" }) '12 au plus' 'Démarrage chargé'

    # Historique de fiabilité (Moniteur de fiabilité de Windows). Indice sous 5/10 : plusieurs plantages ces derniers
    # jours (Windows part de 10 et retire des points à chaque incident) ; plus de 5 en 30 jours : plus d'un par semaine.
    $fb = Get-Fiabilite
    $p.Fiabilite = $fb
    if (($null -ne $fb.Indice -and $fb.Indice -lt 5) -or ($null -ne $fb.Plantages -and $fb.Plantages -gt 5)) {
        Add-Constat 'Postes' 'Conseil' 'Plantages fréquents' "Historique de fiabilité de Windows : $(Get-TexteFiabilite $fb)." 'Identifier l''application ou le pilote en cause (Moniteur de fiabilité de Windows) et le mettre à jour ou le réinstaller'
    }
    Add-Verif 'Poste' 'Historique de fiabilité' (Get-TexteFiabilite $fb) 'indice de stabilité de 5/10 au moins ; 5 plantages au plus en 30 jours' 'Plantages fréquents'

    $p.Constats = @($script:Constats.ToArray() | Select-Object -Skip $c)
    $p.Verifications = $script:Verifs.ToArray()
    return [pscustomobject]$p
}

function Get-Fiabilite {
    # Lecture seule. Moniteur de fiabilité : indice de stabilité sur 10 (dernière valeur) et plantages des 30 derniers
    # jours ; à défaut, journaux Application et Système. Donnée inaccessible : « non mesuré », jamais d'erreur.
    $f = [ordered]@{ Indice = $null; Plantages = $null; Source = ''; Derniers = @() }
    $depuis = $Maintenant.AddDays(-30)
    try {
        $m = Get-CimInstance Win32_ReliabilityStabilityMetrics -ErrorAction Stop | Sort-Object TimeGenerated -Descending | Select-Object -First 1
        if ($m) { $f.Indice = [math]::Round([double]$m.SystemStabilityIndex, 1) }
    } catch { }
    $types = @{ 'Application Error' = 'Plantage d''application'; 'Application Hang' = 'Application bloquée'; 'Microsoft-Windows-WER-SystemErrorReporting' = 'Erreur grave de Windows (écran bleu)'; 'EventLog' = 'Arrêt inattendu de Windows' }
    $evts = $null
    try {
        $wql = "(SourceName='Application Error' OR SourceName='Application Hang' OR SourceName='Microsoft-Windows-WER-SystemErrorReporting' OR (SourceName='EventLog' AND EventIdentifier=6008)) AND TimeGenerated >= '$([Management.ManagementDateTimeConverter]::ToDmtfDateTime($depuis))'"
        $evts = @(Get-CimInstance Win32_ReliabilityRecords -Filter $wql -ErrorAction Stop | ForEach-Object {
            [pscustomobject]@{ Date = $_.TimeGenerated; Application = $(if ([string]$_.SourceName -in 'Application Error', 'Application Hang') { [string]$_.ProductName } else { 'Windows' }); Type = $types[[string]$_.SourceName] } })
        $f.Source = 'Moniteur de fiabilité'
    } catch { $evts = $null }
    if ($null -eq $evts) {
        # Classe absente ou refusée : journaux de Windows (lisibles en utilisateur standard).
        try {
            $lire = { param($Filtre) try { @(Get-WinEvent -FilterHashtable $Filtre -ErrorAction Stop) } catch { if ([string]$_.FullyQualifiedErrorId -like 'NoMatchingEventsFound*') { @() } else { throw } } }
            $evts = @(& $lire @{ LogName = 'Application'; ProviderName = 'Application Error', 'Application Hang'; StartTime = $depuis } | ForEach-Object {
                [pscustomobject]@{ Date = $_.TimeCreated; Application = [string]$_.Properties[0].Value; Type = $(if ($_.ProviderName -eq 'Application Hang') { 'Application bloquée' } else { 'Plantage d''application' }) } })
            $evts += @(& $lire @{ LogName = 'System'; Id = 6008, 1001; StartTime = $depuis } | Where-Object { $_.Id -eq 6008 -or $_.ProviderName -match 'BugCheck|WER-SystemErrorReporting' } | ForEach-Object {
                [pscustomobject]@{ Date = $_.TimeCreated; Application = 'Windows'; Type = $(if ($_.Id -eq 6008) { 'Arrêt inattendu de Windows' } else { 'Erreur grave de Windows (écran bleu)' }) } })
            $f.Source = 'journaux de Windows'
        } catch { $evts = $null }
    }
    if ($null -ne $evts) {
        $f.Plantages = @($evts).Count
        # Seulement la date, le nom du programme et le type : pas de chemin ni de message (données personnelles possibles).
        $f.Derniers = @($evts | Sort-Object Date -Descending | Select-Object -First 5 | ForEach-Object { [ordered]@{ Date = ([datetime]$_.Date).ToString('s'); Application = $_.Application; Type = $_.Type } })
    }
    return [pscustomobject]$f
}

function Get-TexteFiabilite($Fb) {
    $t = @()
    if ($null -ne $Fb.Indice) { $t += "indice de stabilité de $(Format-Nb $Fb.Indice)/10" }
    if ($null -ne $Fb.Plantages) { $t += "$($Fb.Plantages) plantage(s) ou erreur(s) Windows en 30 jours" }
    $t -join ' · '
}
#endregion

#region ===== 2. Le réseau local =====================================================
# Scan du réseau local (box, NAS, imprimantes, caméras, ports à risque) : réservé au Pack TPE complet.
#endregion

#region ===== 3. Internet ===========================================================
function Measure-Latence {
    param([string]$Cible, [int]$N = 10)
    $pg = New-Object System.Net.NetworkInformation.Ping
    $t = @(); $perdus = 0
    foreach ($i in 1..$N) {
        try { $r = $pg.Send($Cible, 1500); if ($r.Status -eq 'Success') { $t += [int]$r.RoundtripTime } else { $perdus++ } } catch { $perdus++ }
        Start-Sleep -Milliseconds 150
    }
    $pg.Dispose()
    $moy = if ($t.Count) { [math]::Round(($t | Measure-Object -Average).Average) } else { $null }
    $gigue = $null
    if ($t.Count -gt 1) { $ecarts = for ($j = 1; $j -lt $t.Count; $j++) { [math]::Abs($t[$j] - $t[$j - 1]) }; $gigue = [math]::Round(($ecarts | Measure-Object -Average).Average) }
    return [ordered]@{ Cible = $Cible; MoyenneMs = $moy; GigueMs = $gigue; PertePct = [math]::Round(100 * $perdus / $N) }
}

function Get-BilanInternet {
    $net = [ordered]@{}
    try {
        $info = Invoke-RestMethod -Uri 'https://ipinfo.io/json' -TimeoutSec 8 -UseBasicParsing
        $net.IPPublique = [string]$info.ip; $net.Operateur = ([string]$info.org -replace '^AS\d+\s+', ''); $net.Ville = [string]$info.city
    } catch { $net.IPPublique = ''; Add-Constat 'Internet' 'Critique' 'Pas d''accès Internet' 'Impossible de joindre Internet depuis ce poste.' 'Vérifier la box et la connexion' }
    Add-Verif 'Internet et Wi-Fi' 'Accès à Internet' $(if ($net.IPPublique) { 'oui' + $(if ($net.Operateur) { " ($($net.Operateur))" }) } else { 'non' }) 'Internet joignable' 'Pas d''accès Internet'

    Etape 'Mesure de la latence et de la stabilité…'
    $net.Latence = Measure-Latence '1.1.1.1'
    $l = $net.Latence
    if ($l.PertePct -ge 5) { Add-Constat 'Internet' 'Important' 'Connexion instable' "$($l.PertePct) % des paquets sont perdus : coupures en visio et en téléphonie." 'Vérifier la box, le câblage ou le Wi-Fi ; contacter l''opérateur si besoin' }
    if ($l.MoyenneMs -and $l.MoyenneMs -gt 80) { Add-Constat 'Internet' 'Conseil' 'Latence élevée' "Temps de réponse moyen de $($l.MoyenneMs) ms (normal : moins de 30 ms en fibre)." 'Vérifier le Wi-Fi et la box' }
    if ($l.GigueMs -and $l.GigueMs -gt 30) { Add-Constat 'Internet' 'Conseil' 'Gigue importante' "Variation du temps de réponse de $($l.GigueMs) ms : appels Teams ou téléphoniques hachés." 'Privilégier le câble pour les postes de visio' }
    Add-Verif 'Internet et Wi-Fi' 'Pertes de paquets' "$($l.PertePct) %" 'moins de 5 %' 'Connexion instable'
    Add-Verif 'Internet et Wi-Fi' 'Temps de réponse' $(if ($null -ne $l.MoyenneMs) { "$($l.MoyenneMs) ms" }) '80 ms au plus' 'Latence élevée'
    Add-Verif 'Internet et Wi-Fi' 'Gigue (variation du temps de réponse)' $(if ($null -ne $l.GigueMs) { "$($l.GigueMs) ms" }) '30 ms au plus' 'Gigue importante'

    # DNS
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try { [void][System.Net.Dns]::GetHostAddresses('www.' + [guid]::NewGuid().ToString('N').Substring(0, 6) + '.fr') } catch { }
    try { [void][System.Net.Dns]::GetHostAddresses('www.service-public.fr') } catch { }
    $sw.Stop(); $net.DnsMs = [int]$sw.ElapsedMilliseconds
    if ($net.DnsMs -gt 500) { Add-Constat 'Internet' 'Conseil' 'Résolution DNS lente' "Résolution des noms en $($net.DnsMs) ms : ouverture des sites ralentie." 'Utiliser les DNS de l''opérateur ou un DNS public rapide' }
    Add-Verif 'Internet et Wi-Fi' 'Résolution des noms (DNS)' "$($net.DnsMs) ms" '500 ms au plus' 'Résolution DNS lente'

    # Débit
    if (-not $SansDebit -and $net.IPPublique) {
        Etape 'Mesure du débit descendant…'
        # 4 flux en parallèle, comme les tests de débit du marché : un seul flux sous-estime la fibre.
        try {
            $clients = 1..4 | ForEach-Object { New-Object System.Net.WebClient }
            $sw = [Diagnostics.Stopwatch]::StartNew()
            $taches = @($clients | ForEach-Object { $_.DownloadDataTaskAsync("https://speed.cloudflare.com/__down?bytes=$(25MB)") })
            [void][System.Threading.Tasks.Task]::WaitAll($taches, 60000)
            $sw.Stop()
            $total = 0; foreach ($t in $taches) { if ($t.Status -eq 'RanToCompletion') { $total += $t.Result.Length } }
            $net.DebitDescendantMbps = if ($total) { [math]::Round(($total * 8 / 1MB) / $sw.Elapsed.TotalSeconds, 1) } else { $null }
        } catch { $net.DebitDescendantMbps = $null }
        Etape 'Mesure du débit montant…'
        try {
            $corps = New-Object byte[] (10MB)
            $clients = 1..4 | ForEach-Object { New-Object System.Net.WebClient }
            $sw = [Diagnostics.Stopwatch]::StartNew()
            $taches = @($clients | ForEach-Object { $_.UploadDataTaskAsync('https://speed.cloudflare.com/__up', 'POST', $corps) })
            [void][System.Threading.Tasks.Task]::WaitAll($taches, 60000)
            $sw.Stop()
            $ok = @($taches | Where-Object { $_.Status -eq 'RanToCompletion' }).Count
            $net.DebitMontantMbps = if ($ok) { [math]::Round(($ok * $corps.Length * 8 / 1MB) / $sw.Elapsed.TotalSeconds, 1) } else { $null }
        } catch { $net.DebitMontantMbps = $null }
        $d = $net.DebitDescendantMbps; $u = $net.DebitMontantMbps
        if ($d -ne $null -and $d -lt 8) { Add-Constat 'Internet' 'Important' 'Débit descendant faible' "$d Mb/s en réception : lent pour le cloud et la visio." 'Comparer à l''offre souscrite ; vérifier l''éligibilité à la fibre' }
        if ($u -ne $null -and $u -lt 2) { Add-Constat 'Internet' 'Important' 'Débit montant faible' "$u Mb/s en envoi : sauvegardes en ligne et envois de fichiers très lents." 'Vérifier l''éligibilité à la fibre' }
    }
    Add-Verif 'Internet et Wi-Fi' 'Débit descendant' $(if ($null -ne $net.DebitDescendantMbps) { "$(Format-Nb $net.DebitDescendantMbps) Mb/s" }) '8 Mb/s au moins' 'Débit descendant faible'
    Add-Verif 'Internet et Wi-Fi' 'Débit montant' $(if ($null -ne $net.DebitMontantMbps) { "$(Format-Nb $net.DebitMontantMbps) Mb/s" }) '2 Mb/s au moins' 'Débit montant faible'
    return [pscustomobject]$net
}
#endregion

#region ===== 4. Wi-Fi ==============================================================
function Get-BilanWifi {
    $w = [ordered]@{ Connecte = $false }
    # netsh écrit en UTF-8 sur les Windows récents : on lit dans ce format, puis on nettoie les espaces
    # insécables que la typographie française place avant les « : ».
    $encAvant = [Console]::OutputEncoding
    try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }
    $sortie = & netsh.exe wlan show interfaces 2>$null
    try { [Console]::OutputEncoding = $encAvant } catch { }
    if (-not $sortie) { Add-Verif 'Internet et Wi-Fi' 'Wi-Fi' 'pas de carte Wi-Fi active' 'signal, sécurité et bande, si le poste est en Wi-Fi' -Resultat 'sans objet'; return [pscustomobject]$w }
    foreach ($brut in $sortie) {
        $ligne = ([string]$brut).Replace([string][char]0x00A0, ' ').Replace([string][char]0x202F, ' ').Replace('┬á', ' ')
        if ($ligne -notmatch '^\s+([^:]+?)\s*:\s*(.+)$') { continue }
        $cle = ($Matches[1] -replace '[^\p{L}\p{N} ]', '').Trim(); $val = $Matches[2].Trim()
        switch -Regex ($cle) {
            '^SSID$'                         { $w.SSID = $val }
            '^Signal$'                       { $w.SignalPct = [int](($val -replace '\D', '') + '0') / 10 }
            '^(Canal|Channel)$'              { $w.Canal = $val }
            '^(Type de radio|Radio type)$'   { $w.Norme = $val }
            '^(Bande|Band)$'                 { $w.Bande = $val }
            '^(Authentification|Authentication)$' { $w.Securite = $val }
            '^(État|Etat|State)$'            { $w.Etat = $val }
            'ception|Receive rate'         { $w.DebitRecepMbps = $val }
        }
    }
    $w.Connecte = [bool]$w.SSID
    if ($w.Connecte) {
        if ($w.SignalPct -lt 50) { Add-Constat 'Réseau' 'Important' 'Wi-Fi faible' "Signal Wi-Fi de $($w.SignalPct) % sur ce poste : lenteurs et déconnexions." 'Rapprocher la borne, ajouter un point d''accès ou passer en câble' }
        if ($w.Securite -match 'WEP|Ouvert|Open') { Add-Constat 'Sécurité' 'Critique' 'Wi-Fi non protégé' "Le réseau Wi-Fi « $($w.SSID) » utilise une sécurité obsolète ou aucune ($($w.Securite))." 'Passer en WPA2 ou WPA3 avec un mot de passe solide' }
        if ($w.Bande -match '2[.,]4' -or ($w.Canal -and [int]($w.Canal -replace '\D', '') -le 14)) { Add-Constat 'Réseau' 'Conseil' 'Wi-Fi en 2,4 GHz' 'Le poste est connecté en 2,4 GHz, la bande la plus lente et la plus encombrée.' 'Activer et privilégier la bande 5 GHz sur la box' }
        # Le nom du réseau (SSID) n'est pas repris : inutile pour la preuve des vérifications.
        Add-Verif 'Internet et Wi-Fi' 'Signal Wi-Fi' "$($w.SignalPct) %" '50 % au moins' 'Wi-Fi faible'
        Add-Verif 'Internet et Wi-Fi' 'Sécurité du Wi-Fi' $w.Securite 'WPA2 ou WPA3' 'Wi-Fi non protégé'
        Add-Verif 'Internet et Wi-Fi' 'Bande Wi-Fi' $(if ($w.Bande) { $w.Bande } elseif ($w.Canal) { "canal $($w.Canal)" }) '5 GHz de préférence' 'Wi-Fi en 2,4 GHz'
    } else { Add-Verif 'Internet et Wi-Fi' 'Wi-Fi' 'poste non connecté en Wi-Fi' 'signal, sécurité et bande, si le poste est en Wi-Fi' -Resultat 'sans objet' }
    return [pscustomobject]$w
}
#endregion

#region ===== Notation ===============================================================
$Penalites = @{ Critique = 25; Important = 12; Conseil = 4 }
$PoidsSimple  = [ordered]@{ 'Postes' = 30; 'Sécurité' = 35; 'Réseau' = 15; 'Internet' = 20 }
$PoidsDomaine = [ordered]@{ 'Postes' = 25; 'Sécurité' = 30; 'Serveur' = 25; 'Réseau' = 10; 'Internet' = 10 }
$Poids = $PoidsSimple
$NonMesures = @()

function Get-Notes {
    param($Constats)
    if (@($Constats | Where-Object { $_.Domaine -eq 'Serveur' }).Count -gt 0 -or $script:AvecDomaine) { $script:Poids = $PoidsDomaine } else { $script:Poids = $PoidsSimple }
    $notes = [ordered]@{}
    foreach ($d in $Poids.Keys) {
        # Thème non mesuré (ex. réseau et Internet quand le bilan du site n'a pas été lancé) : pas de note, jamais de faux 100.
        if ($script:NonMesures -contains $d) { $notes[$d] = $null; continue }
        $pen = 0
        # Pour les postes, une même anomalie présente sur plusieurs PC compte une fois de plus par PC, plafonnée.
        foreach ($g in (@($Constats | Where-Object { $_.Domaine -eq $d }) | Group-Object Sujet)) {
            $base = $Penalites[$g.Group[0].Gravite]
            $pen += [math]::Min($base * 2, $base + ($g.Count - 1) * [math]::Ceiling($base / 3))
        }
        $notes[$d] = [math]::Max(0, 100 - $pen)
    }
    $total = 0; $poidsMesures = 0
    foreach ($d in $Poids.Keys) { if ($null -ne $notes[$d]) { $total += $notes[$d] * $Poids[$d]; $poidsMesures += $Poids[$d] } }
    $notes['Global'] = if ($poidsMesures -gt 0) { [math]::Round($total / $poidsMesures) } else { 0 }
    # Un point critique empêche une note « bonne » : le client doit le voir.
    $nbCritiques = @($Constats | Where-Object { $_.Gravite -eq 'Critique' }).Count
    if ($nbCritiques -gt 0) { $notes['Global'] = [math]::Min($notes['Global'], 74 - 5 * [math]::Min(3, $nbCritiques - 1)) }
    return $notes
}
#endregion

#region ===== Rapport HTML ===========================================================
function New-RapportBilan {
    param([string]$Dossier)
    $fr = [Globalization.CultureInfo]::GetCultureInfo('fr-FR')
    function Enc([string]$t) { [Net.WebUtility]::HtmlEncode($t) }
    if ($Couleur -notmatch '^#[0-9A-Fa-f]{6}$') { $script:Couleur = '#127A70' }

    $postes = @(Get-ChildItem $Dossier -Filter "Poste_$($env:COMPUTERNAME).json" -ErrorAction SilentlyContinue | ForEach-Object { Get-Content $_.FullName -Raw -Encoding UTF8 | ConvertFrom-Json })
    $reseau = $null; $internet = $null; $wifi = $null; $constatsSite = @(); $verifsSite = @()
    $f = Join-Path $Dossier 'Site.json'
    if (Test-Path $f) { $site = Get-Content $f -Raw -Encoding UTF8 | ConvertFrom-Json; $reseau = $site.Reseau; $internet = $site.Internet; $wifi = $site.Wifi; $constatsSite = @($site.Constats); $verifsSite = @($site.Verifications) }
    $script:NonMesures = if (Test-Path $f) { @() } else { @('Réseau', 'Internet') }
    # Bilan lancé avec -SansReseau : pas de note Réseau, jamais de faux 100.
    if ((Test-Path $f) -and -not $reseau) { $script:NonMesures += 'Réseau' }
    # Aucun PC (hors serveurs) analysé : la note Postes n'a pas de sens, jamais de faux 100.
    $pcClients = @($postes | Where-Object { [string]$_.Edition -notmatch 'Server' })
    if ($pcClients.Count -eq 0) { $script:NonMesures += 'Postes' }
    $domaine = $null; $constatsDomaine = @()
    $fd = Join-Path $Dossier 'Domaine.json'
    if (Test-Path $fd) { $domaine = Get-Content $fd -Raw -Encoding UTF8 | ConvertFrom-Json; $constatsDomaine = @($domaine.Constats); $script:AvecDomaine = $true } else { $script:AvecDomaine = $false }
    $tous = @($constatsSite) + @($constatsDomaine) + @($postes | ForEach-Object { $_.Constats })
    $tous = @($tous | Where-Object { $_ })
    $notes = Get-Notes $tous

    function Couleur-Note([int]$n) { if ($n -ge 80) { $Couleur } elseif ($n -ge 60) { '#E0A030' } else { '#C8553D' } }
    function Anneau($n, [int]$taille, [string]$titre, [string]$note = '') {
        if ($null -eq $n) { return "<div class=""g""><svg viewBox=""0 0 130 130"" style=""width:${taille}px;height:${taille}px""><circle cx=""65"" cy=""65"" r=""52"" fill=""none"" stroke=""#E8ECEC"" stroke-width=""12""/><text x=""65"" y=""74"" text-anchor=""middle"" class=""pc"">—</text></svg><b>$(Enc $titre)</b><span class=""muted"">non mesuré</span></div>" }
        $n = [int]$n
        $r = 52; $c = 2 * [math]::PI * $r; $plein = [math]::Round($c * $n / 100, 1).ToString([Globalization.CultureInfo]::InvariantCulture)
        "<div class=""g""><svg viewBox=""0 0 130 130"" style=""width:${taille}px;height:${taille}px""><circle cx=""65"" cy=""65"" r=""$r"" fill=""none"" stroke=""#E8ECEC"" stroke-width=""12""/><circle cx=""65"" cy=""65"" r=""$r"" fill=""none"" stroke=""$(Couleur-Note $n)"" stroke-width=""12"" stroke-linecap=""round"" stroke-dasharray=""$plein 999"" transform=""rotate(-90 65 65)""/><text x=""65"" y=""74"" text-anchor=""middle"" class=""pc"">$n</text></svg><b>$(Enc $titre)</b>$(if ($note) { '<span class=''muted''>' + (Enc $note) + '</span>' })</div>"
    }
    $niveau = if ($notes.Global -ge 80) { 'Bon état général' } elseif ($notes.Global -ge 60) { 'Des points à corriger' } else { 'Situation à risque' }
    function Nb($v) { ([double]$v).ToString('0.#', $fr) }
    # Débit non mesuré (-SansDebit) : la note Internet ne repose que sur la latence et les pertes, on le dit.
    $debitMesure = [bool]($internet -and ($internet.DebitDescendantMbps -or $internet.DebitMontantMbps))
    $anneaux = ''; foreach ($d in $script:Poids.Keys) { $anneaux += Anneau $notes[$d] 96 $(if ($d -eq 'Serveur') { 'Serveur & domaine' } else { $d }) $(if ($d -eq 'Internet' -and $internet -and -not $debitMesure) { 'débit non mesuré' } else { '' }) }

    # Tuiles Internet
    $tuiles = ''
    if ($internet) {
        $t = @()
        if ($internet.DebitDescendantMbps) { $t += @("$(Nb $internet.DebitDescendantMbps) Mb/s", 'Débit descendant') }
        if ($internet.DebitMontantMbps) { $t += @("$(Nb $internet.DebitMontantMbps) Mb/s", 'Débit montant') }
        if ($internet.Latence.MoyenneMs) { $t += @("$($internet.Latence.MoyenneMs) ms", 'Temps de réponse') }
        $t += @("$($internet.Latence.PertePct) %", 'Pertes de paquets')
        if ($wifi -and $wifi.Connecte -and $wifi.SignalPct) { $t += @("$($wifi.SignalPct) %", "Signal Wi-Fi$(if ($wifi.Bande) { ' (' + ($wifi.Bande -replace '[^\d,.]+', ' ').Trim() + ' GHz)' })") }
        for ($i = 0; $i -lt $t.Count; $i += 2) { $tuiles += "<div class=""t""><div class=""v"">$(Enc $t[$i])</div><div>$(Enc $t[$i+1])</div></div>" }
        $op = if ($internet.Operateur) { " · " + (Enc $internet.Operateur) } else { '' }
        $tuiles = "<section><h2>Votre connexion Internet$op</h2><div class=""tuiles"">$tuiles</div></section>"
    }

    # Appareils
    $tabApp = ''
    if ($reseau -and $reseau.Appareils) {
        $grp = @($reseau.Appareils | Group-Object Type | Sort-Object Count -Descending)
        $puces = ($grp | ForEach-Object { "<span class=""chip""><b>$($_.Count)</b> $(Enc $_.Name)</span>" }) -join ''
        $tabApp = "<section><h2>Votre réseau : $(@($reseau.Appareils).Count) appareils détectés</h2><div>$puces</div></section>"
    }

    # Postes
    $tabPostes = ''
    if ($postes.Count) {
        $lignes = ''
        foreach ($p in $postes) {
            $cp = @($p.Constats | Where-Object { $_ })
            $etat = if (@($cp | Where-Object { $_.Gravite -eq 'Critique' }).Count) { '<span class="b r">À traiter</span>' } elseif (@($cp | Where-Object { $_.Gravite -eq 'Important' }).Count) { '<span class="b o">À surveiller</span>' } else { '<span class="b v">Bon état</span>' }
            $age = if ($p.AgeAns) { "$(Nb $p.AgeAns) ans" } else { '—' }
            $lignes += "<tr><td><b>$(Enc $p.Nom)</b><br><span class=""muted"">$(Enc ($p.Fabricant + ' ' + $p.Modele))</span></td><td>$(Enc $p.Windows)</td><td class=""n"">$($p.MemoireGo) Go</td><td class=""n"">$age</td><td>$etat</td></tr>"
        }
        # Historique de fiabilité : une ligne sous le tableau (le détail des derniers événements est dans l'annexe -Detaille).
        $fiab = @($postes | ForEach-Object { $tx = Get-TexteFiabilite $_.Fiabilite; $(if ($postes.Count -gt 1) { (Enc $_.Nom) + ' : ' }) + $(if ($tx) { Enc $tx } else { 'non mesuré' }) })
        $ligneFiab = "<p class=""muted"" style=""margin:10px 0 0"">Historique de fiabilité de Windows : $($fiab -join ' · ').</p>"
        $tabPostes = "<section><h2>Vos ordinateurs</h2><table><thead><tr><th>Poste</th><th>Windows</th><th class=""n"">Mémoire</th><th class=""n"">Âge</th><th>État</th></tr></thead><tbody>$lignes</tbody></table>$ligneFiab</section>"
    }

    # Serveur et domaine
    $tabDomaine = ''
    if ($domaine) {
        $tuilesD = ''
        $vals = @(
            @("$($domaine.NbUtilisateursActifs)", 'Comptes utilisateurs actifs'),
            @("$($domaine.NbOrdinateurs)", 'Ordinateurs dans le domaine'),
            @("$($domaine.NbAdmins)", 'Administrateurs du domaine'),
            $(if ([string]$domaine.DerniereSauvegarde -like 'via *') { , @('À vérifier', "Sauvegarde du serveur ($(([string]$domaine.DerniereSauvegarde).Substring(4)))") }
              elseif ($domaine.DerniereSauvegarde) { , @([string]$domaine.DerniereSauvegarde, 'Dernière sauvegarde du serveur') }
              else { , @('Aucune', 'Sauvegarde du serveur') }))
        foreach ($v in $vals) { $tuilesD += "<div class=""t""><div class=""v"">$(Enc $v[0])</div><div>$(Enc $v[1])</div></div>" }
        $lignesS = ''
        foreach ($srv in @($domaine.Serveurs)) {
            $finProche = $srv.Fin -and ([datetime]$srv.Fin - $Maintenant).TotalDays -lt 365
            $badge = if (-not $srv.Supporte) { '<span class="b r">Hors support</span>' } elseif ($finProche) { "<span class=""b o"">Fin le $(([datetime]$srv.Fin).ToString('dd/MM/yyyy'))</span>" } else { '<span class="b v">Supporté</span>' }
            $lignesS += "<tr><td><b>$(Enc $srv.Nom)</b><br><span class=""muted"">$(Enc $srv.Role)</span></td><td>$(Enc $srv.Systeme)</td><td>$badge</td></tr>"
        }
        $tabS = if ($lignesS) { "<table><thead><tr><th>Serveur</th><th>Système</th><th>Support</th></tr></thead><tbody>$lignesS</tbody></table>" } else { '' }
        $tabDomaine = "<section><h2>Votre serveur et votre domaine · $(Enc $domaine.Nom)</h2><div class=""tuiles"">$tuilesD</div><div style=""margin-top:14px"">$tabS</div></section>"
    }

    # Actions prioritaires
    $ordre = @{ Critique = 0; Important = 1; Conseil = 2 }
    $actions = ''
    foreach ($g in ($tous | Group-Object Sujet | Sort-Object { $ordre[$_.Group[0].Gravite] }, { - $_.Count })) {
        $c0 = $g.Group[0]
        $cls = @{ Critique = 'r'; Important = 'o'; Conseil = 'g' }[$c0.Gravite]
        # Les appareils concernés sont nommés : le client et le technicien savent sur quoi intervenir.
        $noms = @($g.Group | ForEach-Object { [string]$_.Appareil } | Where-Object { $_ -and $_ -ne 'Domaine' } | Select-Object -Unique)
        $qui = if ($noms.Count -gt 1) { "<span class=""muted""> · $($noms.Count) appareils : $(Enc ($noms -join ', '))</span>" } elseif ($noms.Count -eq 1) { "<span class=""muted""> · $(Enc $noms[0])</span>" } else { '' }
        $actions += "<tr><td><span class=""b $cls"">$($c0.Gravite)</span></td><td><b>$(Enc $c0.Sujet)</b>$qui<br><span class=""muted"">$(Enc $c0.Constat)</span></td><td>$(Enc $c0.Action)</td></tr>"
    }
    if (-not $actions) { $actions = '<tr><td colspan="3"><b>Aucun point d''attention : votre informatique est en bon état.</b></td></tr>' }

    # Annexe « Détail des vérifications » (-Detaille) : tous les contrôles, y compris ceux qui sont en ordre.
    $annexe = ''
    if ($Detaille) {
        $verifs = @(@($postes | ForEach-Object { $_.Verifications }) + @($verifsSite) | Where-Object { $_ })
        $clsR = @{ 'OK' = 'v'; 'à surveiller' = 'o'; 'à corriger' = 'r' }
        $blocs = ''
        foreach ($th in 'Poste', 'Sécurité', 'Internet et Wi-Fi') {
            $vt = @($verifs | Where-Object { $_.Theme -eq $th })
            if (-not $vt.Count) { continue }
            $lignesV = ($vt | ForEach-Object {
                $res = [string]$_.Resultat; $cl = $clsR[$res]; if (-not $cl) { $cl = 'g' }
                "<tr><td>$(Enc $_.Controle)</td><td>$(Enc $_.Valeur)</td><td class=""muted"">$(Enc $_.Seuil)</td><td><span class=""b $cl"">$(Enc ($res.Substring(0, 1).ToUpper() + $res.Substring(1)))</span></td></tr>" }) -join ''
            $blocs += "<h3>$(Enc $th)</h3><table class=""fix""><thead><tr><th>Contrôle</th><th>Valeur relevée</th><th>Seuil appliqué</th><th>Résultat</th></tr></thead><tbody>$lignesV</tbody></table>"
            if ($th -eq 'Poste') {
                $derniers = @($postes | ForEach-Object { @($_.Fiabilite.Derniers) } | Where-Object { $_ })
                if ($derniers.Count) {
                    $lignesE = ($derniers | ForEach-Object { "<tr><td>$(([datetime]$_.Date).ToString('dd/MM/yyyy HH:mm'))</td><td>$(Enc $_.Application)</td><td>$(Enc $_.Type)</td></tr>" }) -join ''
                    $blocs += "<p class=""muted"" style=""margin:16px 0 6px"">Derniers événements de l'historique de fiabilité (30 jours) :</p><table><thead><tr><th>Date</th><th>Programme</th><th>Type</th></tr></thead><tbody>$lignesE</tbody></table>"
                }
            }
        }
        $nbOk = @($verifs | Where-Object { $_.Resultat -eq 'OK' }).Count
        $annexe = "<style>.annexe{break-before:page}.annexe h3{font-size:15px;margin:22px 0 6px;color:var(--c)}.annexe td{font-size:13px;padding:8px}.annexe table.fix{table-layout:fixed}.annexe .fix th:first-child{width:28%}.annexe .fix th:last-child{width:17%}</style><section class=""annexe""><h2>Détail des vérifications</h2><p class=""muted"">$($verifs.Count) contrôles effectués, dont $nbOk en ordre, avec la valeur relevée, le seuil appliqué et le résultat. Bilan express v$Version, en lecture seule.</p>$blocs</section>"
    }

    $logoHtml = ''
    if ($Logo -and (Test-Path -LiteralPath $Logo)) {
        $ext = [IO.Path]::GetExtension($Logo).ToLower()
        $mime = @{ '.png' = 'image/png'; '.jpg' = 'image/jpeg'; '.jpeg' = 'image/jpeg'; '.svg' = 'image/svg+xml' }[$ext]
        if ($mime) { $logoHtml = "<img class=""logo"" src=""data:$mime;base64,$([Convert]::ToBase64String([IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $Logo).Path)))"" alt="""">" }
    }
    $sign = if ($Prestataire) { ' — ' + (Enc $Prestataire) } else { '' }

    $html = @"
<!doctype html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Bilan de santé informatique — $(Enc $Client)</title>
<style>
:root{--c:$Couleur;--ink:#1D2A30;--muted:#5D6B70;--line:#E3E7E8}
*{box-sizing:border-box}body{margin:0;background:#F4F6F6;color:var(--ink);font-family:"Segoe UI",Roboto,Arial,sans-serif;font-size:15px;line-height:1.5}
.page{max-width:920px;margin:24px auto;background:#fff;padding:40px 48px;border-radius:12px;box-shadow:0 4px 24px rgba(0,0,0,.06)}
header{display:flex;justify-content:space-between;align-items:center;border-bottom:4px solid var(--c);padding-bottom:18px;margin-bottom:28px;gap:20px}
.logo{max-height:56px;max-width:200px}h1{margin:0;font-size:26px;color:var(--c)}.sub{color:var(--muted)}
.hero{display:flex;align-items:center;gap:28px;margin-bottom:26px}.hero .txt b{font-size:24px;display:block}
.g{text-align:center}.g b{display:block;font-size:14px;margin-top:4px}.pc{font-size:30px;font-weight:700;fill:var(--ink)}
.notes{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:12px;border:1px solid var(--line);border-radius:10px;padding:16px;margin-bottom:30px}
h2{font-size:18px;margin:0 0 12px}section{margin-bottom:30px}
.tuiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px}.t{border:1px solid var(--line);border-radius:10px;padding:14px}.t .v{font-size:24px;font-weight:700;color:var(--c)}
.chip{display:inline-block;border:1px solid var(--line);border-radius:99px;padding:5px 12px;margin:0 8px 8px 0;font-size:14px}
table{width:100%;border-collapse:collapse}td,th{padding:11px 8px;border-bottom:1px solid var(--line);vertical-align:top;text-align:left}
th{font-size:13px;color:var(--muted);font-weight:600}.n{text-align:right;white-space:nowrap}.muted{color:var(--muted);font-size:13px}
.b{display:inline-block;font-size:12px;font-weight:600;padding:2px 9px;border-radius:99px;white-space:nowrap}
.b.r{background:#FBE3DE;color:#9E3A27}.b.o{background:#FCEFD6;color:#8A5A00}.b.v{background:#E2F3EF;color:#0E5F55}.b.g{background:#EEF1F2;color:#46555B;display:inline-block;text-align:left}
footer{color:var(--muted);font-size:12px;border-top:1px solid var(--line);padding-top:14px}
@media (max-width:700px){.page{padding:24px 18px;margin:0;border-radius:0}.notes,.tuiles{grid-template-columns:repeat(2,1fr)}.hero{flex-direction:column;align-items:flex-start}header{flex-direction:column;align-items:flex-start}}
@media print{body{background:#fff}.page{box-shadow:none;margin:0;max-width:none;padding:0}section,tr{break-inside:avoid}}
</style></head><body><div class="page">
<header><div><h1>Bilan de santé de votre informatique</h1><div class="sub">$(Enc $Client) · $($Maintenant.ToString('d MMMM yyyy', $fr))</div></div>$logoHtml</header>
<div class="hero">$(Anneau $notes.Global 150 'Note globale')<div class="txt"><b>$niveau</b><span class="muted">$($postes.Count) ordinateur(s) analysé(s)$(if ($reseau) { ', ' + @($reseau.Appareils).Count + ' appareils sur le réseau' }). $(@($tous | Where-Object { $_.Gravite -eq 'Critique' }).Count) point(s) critique(s), $(@($tous | Where-Object { $_.Gravite -eq 'Important' }).Count) important(s).</span></div></div>
<div class="notes">$anneaux</div>
$tuiles
$tabDomaine
$tabPostes
$tabApp
<section><h2>Ce que nous vous recommandons</h2><table><thead><tr><th>Priorité</th><th>Constat</th><th>Action proposée</th></tr></thead><tbody>$actions</tbody></table></section>
$annexe
<footer>Bilan réalisé sur place, en lecture seule : aucune modification n'a été faite sur vos équipements. $(if ($debitMesure) { 'Mesures du débit indicatives, prises au moment du passage.' } elseif ($internet) { 'Débit non mesuré lors de ce passage : la note Internet porte sur la stabilité de la connexion (temps de réponse, pertes).' })$sign<br>Bilan express Pactio, version gratuite limitée à un PC. Tournée de tout le parc, scan du réseau et rapport à vos couleurs : <a href="$PactioPack">Pack TPE</a>.</footer>
</div></body></html>
"@
    $chemin = Join-Path $Dossier 'Rapport_Bilan_Sante.html'
    [IO.File]::WriteAllText($chemin, $html, (New-Object Text.UTF8Encoding $true))
    return [pscustomobject]@{ Chemin = $chemin; Notes = $notes; Constats = $tous }
}
#endregion

#region ===== Programme principal ===================================================
function Save-Json($Objet, [string]$Chemin) {
    $json = $Objet | ConvertTo-Json -Depth 8
    [IO.File]::WriteAllText($Chemin, $json, (New-Object Text.UTF8Encoding $true))
}

function Start-Bilan {
    Write-Host ""
    Write-Host "===== Pactio · Bilan express v$Version (gratuit) — $Client =====" -ForegroundColor Green
    if (-not $EstAdmin -and $Mode -ne 'Rapport') { Write-Host "  (Lancé sans droits administrateur : BitLocker, SMBv1 et santé des disques ne seront pas tous vérifiés.)" -ForegroundColor DarkYellow }

    if ($Mode -in 'Complet', 'Poste') {
        Etape "Analyse de ce poste ($env:COMPUTERNAME)…"
        $poste = Get-BilanPoste
        Save-Json $poste (Join-Path $Dossier "Poste_$($env:COMPUTERNAME).json")
    }
    if ($Mode -eq 'Complet') {
        $script:Constats.Clear(); $script:Verifs.Clear()
        $site = [ordered]@{ Date = $Maintenant.ToString('s'); DepuisPoste = $env:COMPUTERNAME }
        Etape 'Wi-Fi…'; $site.Wifi = Get-BilanWifi
        if (-not $SansReseau) { $site.Reseau = Get-BilanReseau }
        Etape 'Connexion Internet…'; $site.Internet = Get-BilanInternet
        $site.Constats = $script:Constats.ToArray()
        $site.Verifications = $script:Verifs.ToArray()
        Save-Json ([pscustomobject]$site) (Join-Path $Dossier 'Site.json')
    }

    $r = New-RapportBilan -Dossier $Dossier
    Write-Host ""
    Write-Host ("  Note globale : {0}/100" -f $r.Notes.Global) -ForegroundColor $(if ($r.Notes.Global -ge 80) { 'Green' } elseif ($r.Notes.Global -ge 60) { 'Yellow' } else { 'Red' })
    foreach ($d in $script:Poids.Keys) { if ($null -eq $r.Notes[$d]) { Write-Host ("    {0,-10}   — (non mesuré)" -f $d) -ForegroundColor DarkGray } else { Write-Host ("    {0,-10} {1,3}/100" -f $d, $r.Notes[$d]) } }
    Write-Host ""
    foreach ($g in ($r.Constats | Group-Object Gravite)) { Write-Host ("  {0,-10} {1}" -f $g.Name, $g.Count) }
    Write-Host ""
    Write-Host "  Dossier du bilan : $Dossier"
    Write-Host "  Rapport client   : Rapport_Bilan_Sante.html (ouvrir, puis Ctrl+P pour un PDF)" -ForegroundColor Green
    if ($Detaille) { Write-Host "  Annexe « Détail des vérifications » incluse dans le rapport (option -Detaille)." -ForegroundColor Green }
    Write-Host ""
    Write-Host "  Version gratuite : un seul PC, réseau non scanné." -ForegroundColor DarkGray
    Write-Host "  Tous les PC du client dans un seul rapport, scan du réseau, votre logo : Pack TPE" -ForegroundColor Cyan
    Write-Host "  $PactioPack" -ForegroundColor Cyan
}

if ($MyInvocation.InvocationName -ne '.') { Start-Bilan }
#endregion
