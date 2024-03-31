<#
    Script om een nieuwe user in AD toe te voegen.
    AddAdUser Versie 1.0

    Gecodeerd door Jos Severijnse.
#>
# Leeg het scherm voor nieuw console script.
cls

# Dit script is afhankelijk van deze onpremisse servers.
$ServerExists = 'SomeServer'
$FederationServer = 'SomeFederationServer'

'Even checken of wij in de test of productie omgeving zitten.. . .   .'
$TestOmgeving = Test-Connection -ComputerName $ServerExists -count 1 -Delay 1 -Quiet    # Zitten wij in test of productie? Returns True (1) als wij in test zitten.
$Domain = If ($TestOmgeving) {"@TestSomeDomain.org"} Else {"@SomeDomain.org"}           # Domein naam van de huidige omgeving. Alleen de nieuwe powershell ondersteund tennery operators. ? true : false
$NewUser = $ExampleUser = $ChangeNumber = $ExpirationDate = $NewAccount = ''            # maakt al deze variabele weer leeg, zodra je het script opnieuw start. Voorkomt problemen bij een restart en is nodig voor while loops.
$ExpirationTime = '23:00:00'                                                            # 20:00:00 8 uur in de avond. Tijd is opioneel en kan gebruikt worden.
$ProfilePath = '\\Domain.local\Company\users\'                                          # Path naar alle gebruikers profielen (RUPs)
$HomeDrive = 'Z'                                                                        # De drive waar alle profielen op staan.
$Today = get-Date -Format 'yy-MM-dd'                                                    # De datum van vandaag. Wordt gebruikt als voorbeeld voor het invullen van de datum.
$SleepPeriod = 10

# Array met ongeldige groepen. Ongeldige groepen zijn groepen die appart moeten worden aangevraagd of niet meer van toepassing zijn.
# Deze groepen worden verwijderd van het nieuwe account als het voorbeeld account deze groepen wel toegewezen heeft.
$GroupToRemove = @(
                   "SomeSecurityGroup", `
                   "Domain Users"
                   )

# Functie om te kijken of een variabele een string of een integer is.
function IsNumber ($Value) {
    return $value -match "^[\d\.]+$"
}

<#
    Functie om de nieuwe gebruiker eigenaar te maken en volledig beheer over de homefolder te geven.

    Met onderstaande commando's kan je de opties opvragen voor de verschillende
    parameters voor de FileSystemAccessRule() method.
    
    FileSystemAccessRule($User, $FileSystemRights, $InheritanceFlags, $PropagationFlags, $AccessControlType) 

    FileSystemRights:
    [enum]::GetValues('System.Security.AccessControl.FileSystemRights')

    InheritanceFlags:
    [enum]::GetValues('System.Security.AccessControl.InheritanceFlags')
  
    PropagationFlags:
    [enum]::GetValues('System.Security.AccessControl.PropagationFlags')

    AccessControlType:
    [enum]::GetValues('System.Security.AccessControl.AccessControlType')

#>
Function Set-FolderPermission {
    Param($User, $Folder, $FolderNew)

    if ($FolderNew) {
        New-Item -ItemType Directory -Path $FolderNew
        $Folder = $FolderNew
    }

    # Maakt de user owener.
    $Acl = Get-Acl $Folder
    $Acl.SetOwner([System.Security.Principal.NTAccount]"SomeDomain\$User")

    # Geef fullControl rechten op de map.
    $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("SomeDomain\$User","FullControl","Allow")
    $Acl.addAccessRule($AccessRule)
    Set-Acl $Folder $Acl

    # Geef fullcontrol en stel inheritance in.
    $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("SomeDomain\$User","FullControl", "ContainerInherit,ObjectInherit", "InheritOnly", "Allow")  
    $acl.addAccessRule($AccessRule)
    $acl | Set-Acl $Folder
}

# Vraag naar de volledige naam van de nieuwe gebruiker/medewerker.
''
'Geef de volledige naam van onze nieuwe medewerker op!'
While (!$NewUser) {                                                 # Zolang de variabele $Newuser leeg is.
    $NewUser = Read-Host 'Naam van de medewerker'

    # Ik had verwacht dat als de hele statement in de while loop false was. De body van de loop niet wordt uitgevoerd.
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
$Account = (($Name[0] + $Name[1] + $Surname[0] + $Surname[1]).tolower() + "*")

# Is er al een UserPrincipalName met deze naam is?
# En als dat zo is, dan heeft de variabele $UpnExist een waarde.
$UpnExist = Get-ADUser -filter {UserPrincipalName -like $Upn}

# Als het e-mail adress van de nieuwe user al bestaat. Dus er was al een Jan Jansen (jaja)en nu wordt er een Jaap Jansen (jaja) gemaakt.
# Plak hier dan een 2 of een opvolgnummer achter de naam van het account.
If ($UpnExist) {

    # Stop de naam (eerste letter voornaam plus achternaam) en het opvolgnummer in apparte varabele.
    $UpnUser = $UpnExist.Substring(0, $UpnExist.IndexOf('@'))
    $UpnNumber = $UpnUser[-1]

    # Als functie IsNuber True is, dan bestond er al een tweede account met vergelijkbare upn.
    if (IsNumber $UpnNumber) {
        $Upn = $Name[0] + $FullSurname.replace(' ', '') + ([int]::Parse($UpnNumber) + 1) + $Domain
        
    } else {
        $Upn = $Name[0] + $FullSurname.replace(' ', '') + '2' + $Domain
    }
}

# Haal de namen van alle vergelijkbare SAM accounts op en sorteer de lijst van boven naar beneden.
$SamAccount = (Get-ADUser -Filter {SamAccountName -like $Account} | Sort-Object SamAccountName -Descending | Select SamAccountName -First 1).SamAccountName

# Als de variabele SAMAccount leeg is, dan bestond het account nog helemaal niet. Defineer de nieuwe user hier.
if (!$SamAccount) {

    $SamAccount = $Account.Replace('*', '01')
    $NewAccount = 1
}

# Als het om een admin account gaat, negeer deze dan en gebruik het 2e account in de array.
# Dit kan een bug veroozaken als er meer dan 99 SAM accounts zijn voor een gewone user!
# maar ik betwijldat dat dit ooit zal gebeuren....
If ($SamAccount.Substring(4, 2) -eq '22') {

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

# Haal de Sam naam op van het voorbeeldaccount.
$ExampleLSam = (get-aduser -filter {name -like $ExampleUser}).SamAccountName

# Vraag de properties op van het voorbeeldaccount.
$Title = (Get-ADUser -identity $ExampleLSam -Properties * | select Title).Title
$Department  = (Get-ADUser -identity $ExampleLSam -Properties * | select Department).Department
$Company = (Get-ADUser -identity $ExampleLSam -Properties * | select Company).Company

# In welk OU moet het account worden aangemaakt.
$OrganizationalUnit = ((Get-ADUser $ExampleLSam | select DistinguishedName).DistinguishedName).Substring(((Get-ADUser $ExampleLSam | select DistinguishedName).DistinguishedName).IndexOf(',')+1)

# Als het voorbeeldaccount inmiddels niet meer in dienst is. Exit!
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

# Zolang de verloopdatum leeg is of niet goed is ingevulde. RETURN
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

# Voeg de groepslidmaatschappen toe aan het account.
Foreach ($group in $ExampleUserGroups) {
    Add-ADGroupMember -Identity $group -Members $SamAccount
}

# Maak de home folder aan.
New-Item -ItemType Directory -Path $ProfilePath

# Maak de user eigenaar van de home folder en geef fullcontol  permission.
Set-FolderPermission $SamAccount $ProfilePath

# Maakt de Exchange online mailbox aan in de productieomgeving.
# Deze stap wordt overgeslagen als er een account in de testomgeving wordt aangemaakt.
if ($TestOmgeving -eq 0) {
    ''
    'Momentje geduld, de Exchange Management Shell snap-in wordt nu geladen.'
    # Add-PSsnapin Microsoft.Exchange.Management.PowerShell.E2010
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

    # Maak de mailbox aan voor de nieuwe gebruiker.
    Enable-RemoteMailbox -Identity $SamAccount -RemoteRoutingAddress "$SamAccount@Company web.mail.onmicrosoft.com"

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
