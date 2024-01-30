<#
    Script om een nieuwe user in AD toe te voegen.
    AddAdUser Versie 1.0

    Gecodeerd door Jos Severijnse.
    Co auteur, Michiel Vreeken, Thanks voor de hulp hier en daar!

#>
# Leeg het scherm voor nieuw console script.
cls

# Dit script is afhankelijk van deze onpremisse servers.
$ServerExists = 'SomeServer'
$FederationServer = 'SomeFederationServer'

'Even checken of wij in de test of productie omgeving zitten.. . .   .'
$TestOmgeving = Test-Connection -ComputerName $ServerExists -count 1 -Delay 1 -Quiet    # Zitten wij in test of productie? Returns True (1) als wij in test zitten.
$Domain = If ($TestOmgeving) {"@stadgenoottest.nl"} Else {"@stadgenoot.nl"}             # Domein naam van de huidige omgeving. Alleen de nieuwe powershell ondersteund tennery operators. ? true : false
$NewUser = $ExampleUser = $ChangeNumber = $ExpirationDate = $NewAccount = ''            # maakt al deze variabele weer leeg, zodra je het script opnieuw start. Voorkomt problemen bij een restart en is nodig voor while loops.
$ExpirationTime = '23:00:00'                                                            # 20:00:00 8 uur in de avond. Tijd is opioneel en kan gebruikt worden.
$ProfilePath = '\\Some.Path\To\users\'                                      # Path naar alle gebruikers profielen (RUPs)
$HomeDrive = 'Z'                                                                        # De drive waar alle profielen op staan.
$Today = get-Date -Format 'yy-MM-dd'                                                    # De datum van vandaag. Wordt gebruikt als voorbeeld voor het invullen van de datum.
$SleepPeriod = 10

# Array met ongeldige groepen. Ongeldige groepen zijn groepen die appart moeten worden aangevraagd of niet meer van toepassing zijn.
# Deze groepen worden verwijderd van het nieuwe account als het voorbeeld account deze groepen wel toegewezen heeft.
$GroupToRemove = @(
                   "Domain Users"
                   )

# Vraag naar de volledige naam van de nieuwe gebruiker/medewerker.
''
'Geef de volledige naam van onze nieuwe medewerker op!'
While (!$NewUser) {                                                 # Zolang de variabele $Newuser leeg is.
    $NewUser = Read-Host 'Naam van de medewerker'

    # Ik had verwacht dat als de hele statement in de while loop false was. De body van de loop niet wordt uitgevoer.
    # lijkt toch anders te zijn met en tweede definitie zoals: while(!$NewUser -and $NewUser.lenght -lt 5)
    #
    # Daarom:
    # Als er minder dan 5 karakers in de naam zitten, If op deze plek om $NewUser leeg te maken.
    if ($NewUser.Length -lt 5) {
        $NewUser = ''
    }
}

# Maak variabele aan voor de voornaam, volledige achternaam, achternaam zonder toevoegsels, upn en login naam.
$Name = $NewUser.Substring(0, $NewUser.IndexOf(" "))
$FullSurname = $NewUser.Substring($NewUser.IndexOf(" ")+1)
$Surname = $FullSurname.split()[-1]
$Upn = $Name[0] + $FullSurname.replace(' ', '') + $Domain
$UpnExist = Get-ADUser -filter {UserPrincipalName -like $Upn}
$Account = (($Name[0] + $Name[1] + $Surname[0] + $Surname[1]).tolower() + "*")


# Als het e-mail adress van de nieuwe user al bestaat. Dus er was al een Jan Jansen en nu wordt er een Jaap Jansen gemaakt.
# Plak hier dan een 2 achter de naam, dus: JJansen01@stadgenoot.nl
If ($UpnExist) {
   $Upn = $Name[0] + $FullSurname.replace(' ', '') + '2' + $Domain 
}

# Haal de namen van alle vergelijkbare SAM accounts op en sorteer de lijst van boven naar beneden.
$SamAccount = (Get-ADUser -Filter {SamAccountName -like $Account} | Sort-Object SamAccountName -Descending | Select SamAccountName -First 1).SamAccountName

# Als de variabele SAMAccount leeg is, dan bestond het account nog helemaal niet. Defineer de nieuwe user hier.
if (!$SamAccount) {
    $SamAccount = $Account.Replace('*', '01')
    $NewAccount = 1
}

# Als het om een admin account gaat, negeer deze dan en gebruik het 2e account in de array.
# Dit kan een bug veroozaken als er meer dan 999 SAM accounts zijn voor een gewone user!
# maar ik betwijldat dat dit ooit zal gebeuren....
If ($SamAccount.Substring(4, 2) -eq '99') {
    $SamAccount = ((Get-ADUser -Filter {SamAccountName -like $Account} | Sort-Object SamAccountName -Descending | Select SamAccountName -First 2).SamAccountName)[1]
}

# Is het account opvolgnummer kleiner of groter dan 10? Er kunnen dus 99 accounts gemaakt worden waarvan de letters het zelfde zijn.
if ([int]$SamAccount.Substring(4, 2) + 1 -lt '10') {

    # Negeer een nieuw 01 account. Maakt alle SAM account namen van 02 tot en met 09 aan.
    # Misschien dat de !$NewAccount ook in de defenitie van de eerte if gezet kan worden. Zoiets als -and !$NewAccount.
    if (!$NewAccount) {
        $int =  [int]$SamAccount.Substring(4, 2) + 1
        $int = '0' + $int 
        $SamAccount = $SamAccount.Replace($SamAccount.Substring(4, 2), $int)
    }

} else {
    # Alle accounts vanaf 10 tot en met 99
    $int = [int]$SamAccount.Substring(4, 2) + 1
    $SamAccount = $SamAccount.Replace($SamAccount.Substring(4, 2), $int)
}

# Het Sam account is pas vanaf hier bepaald, dus daarom kan nu de var voor het profilepad gemaakt worden.
$ProfilePath = $ProfilePath + $SamAccount


# Vraag de voor + achternaam van het voorbeeld account op.
''
'Geef de volledige naam, van de voorbeeld gebruiker op!'
While (!$ExampleUser) {
    $ExampleUser = Read-Host 'Naam van het voorbeeld'

    if ($ExampleUser.Length -lt 5) {
        $ExampleUser = ''
    }
}

# Maak de UPN en SAM account naam van het voorbeeldaccount aan.
$ExampleUpn = ($ExampleUser.Substring(0, $ExampleUser.IndexOf(" ")))[0] + ($ExampleUser.Substring($ExampleUser.IndexOf(" ")+1)).replace(' ', '') + $Domain
$ExampleLSam = (Get-ADUser -Filter {UserPrincipalName -eq $ExampleUpn} | Select SamAccountName).SamAccountName

# Vraag de properties op van het voorbeeldaccount.
$Title = (Get-ADUser -identity $ExampleLSam -Properties * | select Title).Title
$Department  = (Get-ADUser -identity $ExampleLSam -Properties * | select Department).Department
$Company = (Get-ADUser -identity $ExampleLSam -Properties * | select Company).Company

# In welk OU moet het account worden aangemaakt.
$OrganizationalUnit = ((Get-ADUser $ExampleLSam | select DistinguishedName).DistinguishedName).Substring(((Get-ADUser $ExampleLSam | select DistinguishedName).DistinguishedName).IndexOf(',')+1)

if ($OrganizationalUnit.Contains("Disabled Accounts")) {
$OrganizationalUnit
    ''
    "Deze persoon $ExampleUser is inmiddels uit dienst."
    'Geef een ander voorbeeld account op.'
    exit
}

# Haal alle groepslidmaatschappen van het voorbeeld account op.
$ExampleUserGroups = (Get-ADPrincipalGroupMembership $ExampleLSam | select name).name

# Verwijder alle onnodige groepen van het account.
Foreach ($group in $ExampleUserGroups) {
    if ($GroupToRemove.Contains($group)) {

        # Wel een vreemde manier om waardes uit een array te halen. Ik ben pop of remove at (index) o.i.d gewend.
        $ExampleUserGroups = ($ExampleUserGroups -ne $group)
    } 
}

# Vraag op wanneer het accout moet verlopen.
''
''
'Geef de verloopdatum van het account op!'
'De datum moet op deze manier ingevuld worden.'
''
'jaar-maand-dag ' + $Today

# Zolang de verloopdatum leeg is of niet good is ingevulde. RETURN
While (!$ExpirationDate) {
    $Date = Read-Host 'Verloopdatum van het account'

    # Simpele conditie om te checken of de datum goed is ingevuld.
    if ($Date.Substring(0, 2) -ige (get-Date -Format "yy") -and $Date.Substring(2, 1) -eq '-' -and $Date.Substring(5, 1) -eq '-') {

        # Stop de waardes in aparte variabele omdat de omschrijving op een ander manier wordt ingevuld dan de verloopdatum.
        $Year =  $Date.Substring(0, 2)
        $Mounth = $Date.Substring(3, 2)
        $Day = $Date.Substring(6, 2)

        $ExpirationDate = $Day + '-' + $Mounth + '-' + $Year
        $ExpirationDate = (Get-Date (([datetime]::ParseExact($ExpirationDate, 'dd-MM-yy', $null)).AddDays(1)) -format "dd-MM-yy")

    # Als de datum niet goed is ingevuld. Display massage and return to while.
    # En leeg de variabele $Date voor hergebruik.
    } else {
        'Je hebt de datum niet goed ingevuld.'
        'Vul de datum op de volgende manier in.'
        ''
        ''
        'jaar-maand-Dag in twee digit nummers'
         $Today
        ''
        $Date = ''
    }
}

# De dag en tijd waarop het account verloopt.
$AccountExpirationDate = $ExpirationDate + ' ' + $ExpirationTime

''
'Geef het nummer van de Topdeskmelding op.'
While (!$ChangeNumber) {
    $ChangeNumber = Read-Host 'Nummer van de melding is'
    # ToDo, nog een check inbouwen of het nummer begint met WA?
}

# De omschrijving van het account, is de verloopdatum plus 1 dag. En het wijzigingsnummer.
$Description = (Get-Date (([datetime]::ParseExact($Date, 'yy-MM-dd', $null)).AddDays(1)) -format "yy-MM-dd")  + ' ' + $ChangeNumber.ToUpper()

# Maakt het account aan in AD en set alle account eigenschappen.
# Als het wachtwoord niet aan de eisen voldoet, dan wordt het account disabled bij het aanmaken.
New-ADUser `
    -Name $NewUser `
    -GivenName $Name `
    -Surname $FullSurname `
    -DisplayName $NewUser `
    -SamAccountName $SamAccount `
    -UserPrincipalName $Upn `
    -Title $Title `
    -Department  $Department `
    -Company $Company `
    -AccountPassword (Read-Host -AsSecureString "Geef het wachtwoord op!") `
    -ChangePasswordAtLogon $False `
    -Enabled $True `
    -Description $Description `
    -AccountExpirationDate $AccountExpirationDate `
    -HomeDirectory $ProfilePath `
    -HomeDrive $HomeDrive `
    -Path $OrganizationalUnit

# Maakt de user owner van zijn/haar homefolder
$Acl = Get-Acl $ProfilePath.FullName
$Acl.SetOwner([System.Security.Principal.NTAccount]"DomainName\$SamAccount")
Set-Acl $ProfilePath.FullName $Acl -Verbose

# Voeg de groepslidmaatschappen toe aan het account.
Foreach ($group in $ExampleUserGroups) {
    Add-ADGroupMember -Identity $group -Members $SamAccount
}

# Maakt de Exchange online mailbox aan in de productieomgeving.
# Deze stap wordt overgeslagen als er een account in de testomgeving wordt aangemaakt.
if ($TestOmgeving -eq 0) {
    ''
    'Momentje geduld, de Exchange Management Shell snap-in wordt nu geladen.'
    # Add-PSsnapin Microsoft.Exchange.Management.PowerShell.E2010
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

    # Maak de mailbox aan voor de nieuwe gebruiker.
    Enable-RemoteMailbox -Identity $SamAccount -RemoteRoutingAddress "$SamAccount@stadgenootweb.mail.onmicrosoft.com"

    # start-sleep -Seconds 10
    ''
    "Even 10 seconden wachten voordat wij AD kunnen synchroniseren"
    sleep $SleepPeriod

    # Ververs de federation server zodat SSO goed werkt.
    Invoke-Command -ComputerName $FederationServer -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}
}

''
''
'Er is een nieuw accout aangemaakt voor ' + $NewUser
'Het e-mail adress van de medewerker is: ' + $upn
''
''
'Check in Active directory of het account goed is aangemaakt!'
'Log in op het account van de gebruiker om de calender sharing in te te stellen.'
''
'Vergeet daarna niet het vinkje aan te zetten dat de gebruiker het wachtwoord MOET wijzigen.'
