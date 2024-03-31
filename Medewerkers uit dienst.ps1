<#
    Medewerkers uit dienst v1.2
    
    Haal een lijst met gebruikers op uit de Disabled OU in AD.
    Specificeer ook welk onderliggend OU's de lijst uit opgehaald moet worden.
    Je kan meerdere OU's opgeven om de data uit op te halen.

    Geschreven door Jos Severijnse.

#>
cls

# Leeg deze variabelen voor als het script nog een keer gestart wordt.
$Everything = $AddMore = $OU = ''

# Script variabelen.
$i = 0
$WriteToFile = @()
$OrganizationalUnit = [System.Collections.ArrayList]@()
$Properties = "Name, SamAccountName, AccountExpirationDate"
$RootOU = "OU=Disabled users,OU=Organisatie,DC=SomeDomain,DC=org"
$File = '.\' + (Get-Date -Format 'yy-MM-dd') + ' Disabled Accounts.txt'

# Vraag de gebruiker of alles opgehaald moet worden of uit verschillende containers.
"Met dit scriptje kan je accounts ophalen van mensen die uit dienst zijn.`nJe kan verschillende OU's opgeven of alle oude accounts ophalen.`n"
While ((!$AddMore) -or ($AddMore -eq "yes") -or ($AddMore -eq "y")) {

    # Vragen of alles of anders verschillende OU's gebruikt moeten worden.
    If (!$Everything) {
        $Everything = Read-Host "Type 'ja' om alle gebruikers op te halen.`nOf druk op enter om OU's toe te voegen"

        # Voegt het root OU toe als alles opgehaald moet worden en stopt de while loop.
        If ($Everything -eq "ja") {
            $OrganizationalUnit.Add($RootOU)
            break

        # Als niet alle accouts opgehaald moeten worden, dan is variabele everytning nee.
        } Else {
            $Everything = "nee"
        }
    }

    # Voeg verschillende OU's toe om accounts uit op te halen.
    $OU = Read-Host "`nGeef de naam van de OU op, waaruit je de accounts wilt ophalen"
    $OrganizationalUnit.Add($OU)
    $AddMore = Read-Host "`nWil je nog een OU toevoegen?`nDruk op enter of type y/yes om nog een OU toe te voegen`nOf type een willekeurig karakter om door te gaan"
}

# Verwijder het Disabled Accounts.txt bestand als het script op dezelfde dag nog een keer wordt uitgevoerd.
If (Test-Path $File) {
    Remove-Item $File
}

# Loop door alle opgegeven OU's.
Foreach ($OU in $OrganizationalUnit) {

    # Als niet alle accounts opgehaald moeten worden. Uit welk OU in de root OU, moeten de accounts opgehaald worden.
    If ($Everything -eq "nee") {
        $SearchPath = "OU=$OU, $RootOU"

    # Vraag alle uitgeschakelde accounts op.
    } ElseIf ($Everything -eq "ja") {
        $SearchPath = $OU
    }

    # Loop door alle opgegeven containers en vraag de account eigenschappen op in die containers.
    Get-ADUser -SearchBase $SearchPath -filter * -Properties $Properties | % {
        $i++
        $WriteToFile = @()

        # Laat de output in het console zien.
        "`n_____________________ $i _____________________"
        'Medewerker      : ' + $_.Name
        'Account naam    : ' + $_.SamAccountname
        'Laatste werkdag : ' + $_.AccountExpirationDate

        # Bewaar tijdelijk de data in een array.
        $WriteToFile += "`n`n_____________________ $i _____________________`n"
        $WriteToFile += 'Medewerker      : ' + $_.Name
        $WriteToFile += 'Account naam    : ' + $_.SamAccountname
        $WriteToFile += 'Laatste werkdag : ' + $_.AccountExpirationDate

        # Schrijf de output van de array naar een bestand in de map waar het script in uitgevoerd wordt.
        Out-File -FilePath $File -Append -InputObject $WriteToFile
    }
}
