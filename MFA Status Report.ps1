<#
    MFA Status Report v1.1.0
    
    Dit script genereert een authenticatie statusrapport van alle accounts in de cloud.
    Het rapport wordt bewaard in de folder waar het script staat. 
    In een map met de naam "OutputFiles". Ieder HTML rapport heeft een unique naam.

    Deze code is geïnspireerd onderstaande website:
    https://practical365.com/mfa-status-user-accounts/

    Geschreven door: Jos Severijnse.
#>
cls
Write-Host "Momentje geduld......"

$totalAccounts = $PreferredAuthenticationMethodEnabled = $IsSsprRegistered = $IsSsprEnabled = $IsSsprCapable = $IsPasswordlessCapable = $IsMfaRegistered = $IsMfaCapable = $IsAdmin = [int]$i = 0
$users = $userAuthDetails = [psobject]$allUserDetails = @()

[string]$TenantId           = "xxxxxx-xxxxxx-xxxxxx-xxxxxx"
[string]$scriptFolder       = $PSScriptRoot
[string]$scriptSubFolder    = Join-Path $scriptFolder "OutputFiles"
[string]$html               = ""

# .Net code. Zorg er voor dat het script verbinding kan maken met het internet, middels het TLS protocol.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Installeer de MS Graph module, als hij niet geïnstalleerd is!
if (!(Get-Module -ListAvailable -name Microsoft.Graph)) {
    Install-Module Microsoft.Graph
}

# Maak stilletjes verbinding met de Graph API'S.
# Connect-MgGraph -TenantId $TenantId -Scopes "User.Read.All", "Directory.Read.All" # Voor deze methode heb je admin approval nodig.....
Connect-MgGraph -TenantId $TenantId -NoWelcome

# Haal alle accounts in de wolk op.
$users = Get-MgUser -All

# Haal alle authenticatie eigenschappen op.
$userAuthDetails = Get-MgReportAuthenticationMethodUserRegistrationDetail

# Functie om de twee objecten, $users en $userAuthDetails te combineren en daar de benodigde eigenschappen van op te halen.
function Get-AllUserAuthProperties {

    # Input waardes van de functie.
    param (
        $users,
        $userAuthDetails
    )

    # Het object om alle accountgegevens in op te slaan.
    $allUserDetails = @()

    # De hashtable (key, pair), waar de data van de huidige user in de loop in op wordt geslagen.
    $thisUserDetails = @{}

    # Loop door de accounteigenschappen van ieder account in de omgeving.
    foreach ($userDetails in $userAuthDetails) {

        $thisUserID = $userDetails.Id

        # Loop door alle users in de AAD tennent.
        foreach ($userPropertie in $users) {

            # Zodra het user-id uit de eerste foreach loop, gelijk is aan het ID in de huidige loop.
            # Bewaar dan de voor, achternaam en het e-mail adres van de gebruiker. Stop de loop, zodat het script verder kan gaan.
            if ($thisUserID -eq $userPropertie.Id) {
                $DisplayName = $userPropertie.DisplayName
                $mail = $userPropertie.UserPrincipalName

                break
            }
        }

        # Als het om een SomeCompany account gaat, bewaar dan de gegevens van het account in de loop.
        if ($mail -match "@SomeCompany.nl") {

            # De tijdelijke hashtable, met de account details van de huidige user in de loop.
            $thisUserDetails = @{
                UserName                                     = $DisplayName
                mail                                         = $mail
                Id                                           = $userDetails.Id
                IsAdmin                                      = $userDetails.IsAdmin
                IsMfaCapable                                 = $userDetails.IsMfaCapable
                IsMfaRegistered                              = $userDetails.IsMfaRegistered
                IsPasswordlessCapable                        = $userDetails.IsPasswordlessCapable
                IsSsprCapable                                = $userDetails.IsSsprCapable
                IsSsprEnabled                                = $userDetails.IsSsprEnabled
                IsSsprRegistered                             = $userDetails.IsSsprRegistered
                IsSystemPreferredAuthenticationMethodEnabled = $userDetails.IsSystemPreferredAuthenticationMethodEnabled
                LastUpdatedDateTime                          = $userDetails.LastUpdatedDateTime
                MethodsRegistered                            = $userDetails.MethodsRegistered
                SystemPreferredAuthenticationMet             = $userDetails.SystemPreferredAuthenticationMet
            }

            # Bewaar de account gegevens van de huide user in de loop in het object $allUserDetails
            $allUserDetails += New-Object PSObject -Property $thisUserDetails
        }
    }

    # return het object $allUserDetails met alle authenticatie eigenschappen van ieder SomeCompany account in de wolk.
    return $allUserDetails
}

# Start Chrome met het URL naar de brief.
function Start-Browser {

    param ($Lettter_URL)

    start-Process -FilePath "chrome.exe" -ArgumentList "$Lettter_URL" -PassThru
    Start-Sleep -Milliseconds 2000
    $Process = ($ThisProcess = (Get-process -name "chrome"))[($ThisProcess.length)-1]

    # Wacht totdat Chrome is gesloten.
    while ($Process.HasExited -eq $false) {
        
        # Exit de while loop zodra de browser gesloten wordt.
        if (!(Get-Process -Id $Process.Id -ErrorAction SilentlyContinue)) {
            break 
        }
        Start-Sleep -Milliseconds 50 
    }

    return $Process
}

# Roep de functie aan met zijn input waardes. En krijg alle account authenticatie eigenschappen in object $allUserDetails.
$allUserDetails = Get-AllUserAuthProperties $users $userAuthDetails

# Dit stukje is: De HTML header van de webpagina. Het bepaald de titel en opmaak van de pagina.
$html = @"
<!DOCTYPE html>
<html>
<head>
    <title>MFA Status Report</title>
    <style>
        table {
            width: 100%;
            border-collapse: collapse;
        }
        th, td {
            border: 1px solid black;
            padding: 8px;
            text-align: center; 
            vertical-align: middle;
        }
        th {
            background-color: #f2f2f2;
        }
        .yellow {
            background-color: yellow;
            color: black;
        }
        .red {
            background-color: red;
            color: white;
        }
        .orange {
            background-color: orange;
            color: black;   
        }
        .summery {
            width: 10%;
            margin-left: auto;
            margin-right: auto;
        }
    </style>
</head>
"@

# Haalt de cijfers op, voor het eerste tabel van de pagina.
foreach ($propertie in $allUserDetails) {
    $totalAccounts++

    # Tel alle false waardes, van iedere eigenschap op.                       Plus 1.
    if ($propertie.IsAdmin -eq $true)                                       { $IsAdmin++ }
    if ($propertie.IsMfaCapable -eq $false)                                 { $IsMfaCapable++ }
    if ($propertie.IsMfaRegistered -eq $false)                              { $IsMfaRegistered++ }
    if ($propertie.IsPasswordlessCapable -eq $false)                        { $IsPasswordlessCapable++ }
    if ($propertie.IsSsprCapable -eq $false)                                { $IsSsprCapable++ }
    if ($propertie.IsSsprRegistered -eq $false)                             { $IsSsprRegistered++ }
    if ($propertie.IsSsprEnabled -eq $false)                                { $IsSsprEnabled++ }
    if ($propertie.IsSystemPreferredAuthenticationMethodEnabled -eq $false) { $PreferredAuthenticationMethodEnabled++ }
}

# De body van de webpagina. Hier wordt het eerste tabel gemaakt.
# En als laatst wordt de eerste row/header, van het tweede tabel gemaakt.
$html += @"
<body>
    <h2 style="text-align: center;">MFA/SSO Status Report</h2>

    <table class="summery">
        <thead>
            <tr>
                <th></th>
                <th>Admin Accounts</th>
                <th>IsMfaCapable</th>
                <th>IsMfaRegistered</th>
                <th>IsPasswordlessCapable</th>
                <th>IsSsprCapable</th>
                <th>IsSsprEnabled</th>
                <th>IsSsprRegistered</th>
                <th>IsSystemPreferredAuthenticationMethodEnabled</th>
                <th></th>
            </tr>
        </thead>
        <tbody>
        <tr>
            <td style="background-color: #f2f2f2;"><b>True</b></td>
            <td>$($IsAdmin)</td>
            <td>$($totalAccounts - $IsMfaCapable)</td>
            <td>$($totalAccounts - $IsMfaRegistered)</td>
            <td>$($totalAccounts - $IsPasswordlessCapable)</td>
            <td>$($totalAccounts - $IsSsprCapable)</td>
            <td>$($totalAccounts - $IsSsprRegistered)</td>
            <td>$($totalAccounts - $IsSsprEnabled)</td>
            <td>$($totalAccounts - $PreferredAuthenticationMethodEnabled)</td>
            <td style="background-color: #f2f2f2;"><b>.......</b></td>
        </tr>
        <tr>
            <td style="background-color: #f2f2f2;"><b>False</b></td>
            <td>$($totalAccounts - $IsAdmin)</td>
            <td>$($IsMfaCapable)</td>
            <td>$($IsMfaRegistered)</td>
            <td>$($IsPasswordlessCapable)</td>
            <td>$($IsSsprCapable)</td>
            <td>$($IsSsprRegistered)</td>
            <td>$($IsSsprEnabled)</td>
            <td>$($PreferredAuthenticationMethodEnabled)</td>
            <td style="background-color: #f2f2f2;"><b>.......</b></td>
        </tr>
        </tbody>
    </table>

    <br><br>

    <table>
        <thead>
            <tr>
                <th>Nr</th>
                <th>DisplayName</th>
                <th>Email</th>
                <th>Id</th>
                <th>IsAdmin</th>
                <th>IsMfaCapable</th>
                <th>IsMfaRegistered</th>
                <th>IsPasswordlessCapable</th>
                <th>IsSsprCapable</th>
                <th>IsSsprEnabled</th>
                <th>IsSsprRegistered</th>
                <th>IsSystemPreferredAuthenticationMethodEnabled</th>
                <th>LastUpdatedDateTime</th>
                <th>MethodsRegistered</th>
            </tr>
        </thead>
        <tbody>
"@

# Haalt de data op, voor het tweede tabel van de pagine. En vult alle rijen op de pagina.
foreach ($propertie in $allUserDetails) {

    # Aantal accounts in de wolk.
    $i++

    # Bepaal de CSS klassen op basis van de eigenschappen true/false. Geeft de cell een andere kleur, indien false.
    $adminClass                                     = if ($propertie.IsAdmin -eq $true) { "red" } else { "" }
    $mfaCapableClass                                = if ($propertie.IsMfaCapable -eq $false) { "yellow" } else { "" }
    $mfaRegisteredClass                             = if ($propertie.IsMfaRegistered -eq $false) { "yellow" } else { "" }
    $IsPasswordlessCapableClass                     = if ($propertie.IsPasswordlessCapable -eq $false) { "orange" } else { "" }
    $IsSsprCapableClass                             = if ($propertie.IsSsprCapable -eq $false) { "orange" } else { "" }
    $IsSsprRegisteredClass                          = if ($propertie.IsSsprRegistered -eq $false) { "orange" } else { "" }
    $IsSsprEnabledClass                             = if ($propertie.IsSsprEnabled -eq $false) { "orange" } else { "" }
    $IsSystemPreferredAuthenticationMethodEnabled   = if ($propertie.IsSystemPreferredAuthenticationMethodEnabled -eq $false) { "red" } else { "" }

    # Alle date voor de huidige rij in het tabel.
    $html += "<tr>"
        $html += "<td>$($i)</td>"
        $html += "<td>$($propertie.UserName)</td>"
        $html += "<td>$($propertie.mail)</td>"
        $html += "<td>$($propertie.Id)</td>"
        $html += "<td class = '$adminClass'>$($propertie.IsAdmin)</td>"
        $html += "<td class = '$mfaCapableClass'>$($propertie.IsMfaCapable)</td>"
        $html += "<td class = '$mfaRegisteredClass'>$($propertie.IsMfaRegistered)</td>"
        $html += "<td class = '$IsPasswordlessCapableClass'>$($propertie.IsPasswordlessCapable)</td>"
        $html += "<td class = '$IsSsprCapableClass'>$($propertie.IsSsprCapable)</td>"
        $html += "<td class = '$IsSsprRegisteredClass'>$($propertie.IsSsprRegistered)</td>"
        $html += "<td class = '$IsSsprEnabledClass'>$($propertie.IsSsprEnabled)</td>"
        $html += "<td class = '$IsSystemPreferredAuthenticationMethodEnabled'>$($propertie.IsSystemPreferredAuthenticationMethodEnabled)</td>"
        $html += "<td>$($propertie.LastUpdatedDateTime)</td>"
        $html += "<td>$($propertie.MethodsRegistered)</td>"
    $html += "</tr>"
}

# HTML footer. Sluit de pagina.
$html += @"
        </tbody>
    </table>
</body>
</html>
"@

# Maak de map om de output van dit script in te bewaren, als hij niet bestaat.
if (!(Test-Path -Path $scriptSubFolder)) {
    New-Item -Path $scriptSubFolder -ItemType Directory
    Start-Sleep -Milliseconds 100
}

# Waar de html pagina bewaard moet worden.
$outputPath = "$scriptSubFolder\mfa_status_report_$(Get-Date -Format 'ddMMyyyyHHmmss').html"
# $outputPath = "$scriptSubFolder\mfa_status_report.html"

# Sla de html pagina op.
$html | Out-File -FilePath $outputPath -Encoding UTF8

# Laat de gebruiker weten waar het rapport is opgeslagen.
cls
Write-Host "Het HTML rapport, is in de volgende map opgeslagen:`n`n   $outputPath`n`n`nStarting browser......"

# Start de browser met het url naar de rapport.
Start-Browser $outputPath > $null
