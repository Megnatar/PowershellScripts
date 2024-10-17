<#
    AddAdUser Versie 2.0.0.4

    Met dit tooltje kan je eenvoudig een nieuwe medewerker aanmaken in AD.
    Het doet dit op basis van vijf waardes die door de gebruiker worden ingevoerd.

    Deze waardes zijn:

        De volledige naam van de nieuwe medewerker.
        Een account waarvan de gegevens gekopieerd moeten worden.
        De datum waarop het contract eindigt. Dit is standaard een jaar.
        Het wachtwoord van het nieuwe account.
        Het wijzigingsnummer van de aanvraag voor het nieuwe account.

    Na het aanmaken van een account wordt de welkom brief gemaakt.
    De browser zal opstarten met de link naar de brief die uitgeprint moet worden.
    Het uitprinten, kan je doen door de sneltoeltoets ctrl+p te gebruiken.
    Zorg dat je print naar PDF, en zet het printen naar headers en footers uit!

    Het is belangrijk dat de brief wordt opgeslagen in de folder waar het script in uitgevoerd wordt!

    Zodra de brief is gemaakt, wordt er standaard een e-mail verzonden naar de helpdesk en topdesk.
    Na het zenden van de e-mail wordt die brief verplaatst naar de share met alle welkoms brieven van het huidige jaar.

    Geschreven door Jos Severijnse
#>


# ________________________________________________ Script variabelen en array's  ______________________________________________________

# Maakt al deze variabele weer leeg, voor als het script opnieuw wordt gestart.
$mailbox = $NewUser = $Date = $Password = $ExampleUser = $ChangeNumber = $ExpirationDate = $NewAccount = ''

# Dit script is afhankelijk van deze onpremisse servers.
$ServerExists       = 'SomeServer'                                                            # De naam van de server in een testomgeving.
$FederationServer   = 'SomeServer'                                                            # De naam van de Federation server.

# Globale variabelen.
$ExpirationTime     = '23:00:00'                                                               # 20:00:00 8 uur in de avond. Tijd is opioneel en kan gebruikt worden.
$Image              = 'SomeCompany.png'                                                        # SomeCompany image voor de welkoms brief.
$ProfilePath        = '\\connect.local\SomeCompany\users\'                                     # Path naar alle gebruikers profielen (RUPs)
$pathToletters      = "\\connect.local\SomeCompany\SomePath\$(get-date -format yyyy)"
$Today              = get-Date -Format 'dd-MM-yyyy'                                           # De datum van vandaag. Wordt niet gebruikt.
$HomeDrive          = 'SomeDriveLetter'                                                       # De drive/mapping waar alle profielen op staan.
$SleepPeriod        = 10                                                                      # Wacht periode is 10 seconde.
$AccountExist       = 0                                                                       # Is er al een nieuw account gemaakt?
$pdfFiles           = @()                                                                     # Een array om alle gevonde pdf bestanden in te bewaren.

# Array met ongeldige groepen. Ongeldige groepen zijn groepen die appart moeten worden aangevraagd of niet meer van toepassing zijn.
# Deze groepen worden verwijderd van het nieuwe account als het voorbeeld account deze groepen wel toegewezen heeft.
# Groepen waarvan de weergavenaam anders is dan de groepsnaam, kunnen een error geven.
$GroupToRemove = @(
                "SomeSecurityGroup", `
                "SomeSecurityGroup",
)

# Zorg dat het script gebruik kan maken van TLS protocol.
# En login op exchange online indien nodig
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$session = (Get-ConnectionInformation).name

if (!$session) {
    Connect-ExchangeOnline
}

cls
Write-Host          "Testen of wij in de test of productie omgevinge zitten.`n"
$TestOmgeving       = Test-Connection -ComputerName $ServerExists -count 1 -Delay 1 -Quiet    # Zitten wij in test of productie? Returns True (1) als wij in test zitten.
$Domain             = If ($TestOmgeving) {"@SomeCompanytest.nl"} Else {"@SomeCompany.nl"}       # Domein naam van de huidige omgeving. Alleen de nieuwe powershell ondersteund tennery operators. ? true : false
Write-Host          "Op dit moment wordt domein $domain gebruikt."

# ________________________________________________ Functions ______________________________________________________

# Functie om te kijken of een variabele een string of een integer is.
function IsNumber ($Value) {
    return $value -match "^[\d\.]+$"
}

# Start Chrome met het URL naar de brief.
function Start-Browser {

    param ($Lettter_URL)

    start-Process -FilePath "chrome.exe" -ArgumentList "$Lettter_URL" -PassThru
    Start-Sleep -Milliseconds 2000
    $Process = ($ThisProcess = (Get-process -name "chrome"))[($ThisProcess.length)-1]

    # Wacht totdat Chrome is gesloten. Druk ctrl+P om naar pdf te printen. Zet het printen van headers en footers uit!
    while ($Process.HasExited -eq $false) {
        
        if (!(Get-Process -Id $Process.Id -ErrorAction SilentlyContinue)) {
            break 
        }

        Start-Sleep -Milliseconds 50 
    }

    return $Process
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
# Stel de NTFS eigenschappen in voor de home folder. $FolderNew is optioneel.
Function Set-FolderPermission {

    Param(
        $User,
        $Folder,
        $FolderNew
    )

    if ($FolderNew) {
        New-Item -ItemType Directory -Path $FolderNew > $null
        $Folder = $FolderNew
    }

    # Maakt de user owner.
    $Acl = Get-Acl $Folder
    $Acl.SetOwner([System.Security.Principal.NTAccount]"CONNECT\$User")

    # Geef fullControl rechten op de map.
    $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("CONNECT\$User","FullControl","Allow")
    $Acl.addAccessRule($AccessRule)
    Set-Acl $Folder $Acl

    # Geef fullcontrol op de submappen en stel inheritance in.
    $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("CONNECT\$User","FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")  
    $acl.addAccessRule($AccessRule)
    Set-Acl $Folder $acl
}
 
# ________________________________________________ GUI voor het script ______________________________________________________

# Zorg dat powershell de .net bibliotheken voor het maken van een GUI kan gebruiken.
# Het form, de GUI, zal de visuele style van het OS gebruiken.
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# De GUI functie. Dit is de defenitie van de GUI. Het bepaald hoe de GUI er uit ziet.
# Welke controlls en elementen er gebruikt worden en wat de positie is van deze objecten.
# Alles wat met de GUI te maken heeft gebeurt hier!
function Show-GUI {

    param(
        $AccountExist,
        $NaamMedewerker,
        $VoorbeeldAccount,
        $EindDatum,
        $Wachtwoord,
        $Topdesknummer,
        $Manager
    )

    # ________________________________________________ Form eigenschappen ______________________________________________________

    $Form                            = New-Object system.Windows.Forms.Form
    $Form.ClientSize                 = New-Object System.Drawing.Point(310,280)
    $Form.text                       = "Add user to AD"
    $Form.StartPosition              = "CenterScreen"
    $Form.FormBorderStyle            = "FixedDialog"
    $Form.TopMost                    = $true
    $Form.MaximizeBox                = $false
    $Form.MinimizeBox                = $true

    # ________________________________________________ Nieuwe medewerker ______________________________________________________

    $Label1                          = New-Object system.Windows.Forms.Label
    $Label1.text                     = "Naam medewerker : "
    $Label1.AutoSize                 = $true
    $Label1.width                    = 25
    $Label1.height                   = 12
    $Label1.location                 = New-Object System.Drawing.Point(14,18)

    $TextBox1                        = New-Object system.Windows.Forms.TextBox
    $TextBox1.multiline              = $false
    
    If ($NaamMedewerker){
        $TextBox1.text               = $NaamMedewerker

    } Else {
        $TextBox1.text               = ""
    }

    $TextBox1.width                  = 150
    $TextBox1.height                 = 20
    $TextBox1.location               = New-Object System.Drawing.Point(145,16)

    # Zet de controll uit als er een nieuwe user is aangemaakt.
    If ($AccountExist -eq 1) {
        $TextBox1.Enabled            = $false
    }

    # ________________________________________________ Voorbeeld account ______________________________________________________

    $Label2                          = New-Object system.Windows.Forms.Label
    $Label2.text                     = "Voorbeeld account : "
    $Label2.AutoSize                 = $true
    $Label2.width                    = 25
    $Label2.height                   = 12
    $Label2.location                 = New-Object System.Drawing.Point(14,53)

    $TextBox2                        = New-Object system.Windows.Forms.TextBox
    $TextBox2.multiline              = $false

    If ($VoorbeeldAccount){
        $TextBox2.text               = $VoorbeeldAccount

    } Else {
        $TextBox2.text               = ""
    }

    $TextBox2.width                  = 150
    $TextBox2.height                 = 20
    $TextBox2.location               = New-Object System.Drawing.Point(145,51)

    If ($AccountExist -eq 1) {
        $TextBox2.Enabled            = $false
    }
    # ________________________________________________ Datum einde contract ______________________________________________________

    $Label3                          = New-Object system.Windows.Forms.Label
    $Label3.text                     = "Eind datum : "
    $Label3.AutoSize                 = $true
    $Label3.width                    = 25
    $Label3.height                   = 12
    $Label3.location                 = New-Object System.Drawing.Point(14,88)

    $TextBox3                        = New-Object system.Windows.Forms.TextBox
    $TextBox3.multiline              = $false

    If ($EindDatum) {
        $TextBox3.text               = $EindDatum

    } Else {
        $TextBox3.text               = ""
    }

    $TextBox3.width                  = 150
    $TextBox3.height                 = 20
    $TextBox3.location               = New-Object System.Drawing.Point(145,86)

    # $DateTimePicker1                 = New-Object System.Windows.Forms.DateTimePicker
    # $DateTimePicker1.location        = New-Object System.Drawing.Point(145,86)
    # $DateTimePicker1.width           = 150

    If ($AccountExist -eq 1) {
        $TextBox3.Enabled     = $false
    }

    # ________________________________________________ Wachtwoord ______________________________________________________

    $Label4                          = New-Object system.Windows.Forms.Label
    $Label4.text                     = "Wachtwoord : "
    $Label4.AutoSize                 = $true
    $Label4.width                    = 60
    $Label4.height                   = 12
    $Label4.location                 = New-Object System.Drawing.Point(14,123)

    $TextBox4                        = New-Object system.Windows.Forms.TextBox
    $TextBox4.multiline              = $false

    If ($Wachtwoord){
        $TextBox4.text               = $Wachtwoord

    } Else {
        $TextBox4.text               = ""
    }

    $TextBox4.width                  = 150
    $TextBox4.height                 = 20
    $TextBox4.location               = New-Object System.Drawing.Point(145,121)

    If ($AccountExist -eq 1) {
        $TextBox4.Enabled            = $false
    }

    # ________________________________________________ Nummer van de wijziging ______________________________________________________

    $Label5                          = New-Object system.Windows.Forms.Label
    $Label5.text                     = "Topdesknummer : "
    $Label5.AutoSize                 = $true
    $Label5.width                    = 90
    $Label5.height                   = 12
    $Label5.location                 = New-Object System.Drawing.Point(14,158)

    $TextBox5                        = New-Object system.Windows.Forms.TextBox
    $TextBox5.multiline              = $false

    If ($Topdesknummer){
        $TextBox5.text               = $Topdesknummer

    } Else {
        $TextBox5.text               = ""
    }

    $TextBox5.width                  = 150
    $TextBox5.height                 = 20
    $TextBox5.location               = New-Object System.Drawing.Point(145,156)

    If ($AccountExist -eq 1) {
        $TextBox5.Enabled            = $false
    }

    # ________________________________________________ manager  ______________________________________________________

    $Label6                          = New-Object system.Windows.Forms.Label
    $Label6.text                     = "Naam Manager : "
    $Label6.AutoSize                 = $true
    $Label6.width                    = 90
    $Label6.height                   = 12
    $Label6.location                 = New-Object System.Drawing.Point(14,193)

    $TextBox6                        = New-Object system.Windows.Forms.TextBox
    $TextBox6.multiline              = $false

    If ($Topdesknummer){
        $TextBox6.text               = $Manager

    } Else {
        $TextBox6.text               = ""
    }

    $TextBox6.width                  = 150
    $TextBox6.height                 = 20
    $TextBox6.location               = New-Object System.Drawing.Point(146,191)

    If ($AccountExist -eq 1) {
        $TextBox6.Enabled            = $false
    }

    # ________________________________________________ Button Submit ______________________________________________________
    
    $ButtonSubmit                   = New-Object system.Windows.Forms.Button
    $ButtonSubmit.text              = "Submit"
    $ButtonSubmit.width             = 80
    $ButtonSubmit.height            = 30
    $ButtonSubmit.location          = New-Object System.Drawing.Point(135,235)

    # Voegt de functie Add_Click aan object $ButtonSubmit toe.
    $ButtonSubmit.Add_Click({
        $script:returnValue = @{
            "NaamMedewerker"    = $TextBox1.text
            "VoorbeeldAccount"  = $TextBox2.text
            "EindDatum"         = $TextBox3.text    # DateTimePicker1.Value.ToShortDateString()
            "Wachtwoord"        = $TextBox4.text
            "Topdesknummer"     = $TextBox5.Text
            "Manager"           = $TextBox6.Text
        }

        $Form.Close()
    })

    # ________________________________________________ Button Exit ______________________________________________________

    $ButtonExit                     = New-Object system.Windows.Forms.Button
    $ButtonExit.text                = "Exit"
    $ButtonExit.width               = 80
    $ButtonExit.height              = 30
    $ButtonExit.location            = New-Object System.Drawing.Point(220,235)

    $ButtonExit.Add_Click({
        $script:returnValue = @{
            "Close" = 1
        }

        $Form.Close()
    })

    # ________________________________________________ Button SemdMail ______________________________________________________

    $ButtonNewMail                     = New-Object system.Windows.Forms.Button
    $ButtonNewMail.text                = "New mail"
    $ButtonNewMail.width               = 80
    $ButtonNewMail.height              = 30
    $ButtonNewMail.location            = New-Object System.Drawing.Point(50,235)
    $ButtonNewMail.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

    If ($AccountExist -eq 1) {

        $ButtonSubmit.Enabled = $false

        $ButtonNewMail.Add_Click({
            $script:returnValue = @{
                "Close" = 1
                "Sendmail" = 1
            }
            # $Form.refresh()
            $Form.Close()
        })

    } Else {
        $ButtonNewMail.Enabled = $false
    }
    
    # Voeg de controlls aan het form toe.
    $Form.controls.AddRange(@(
            $Label1,
            $TextBox1,
            $Label2,
            $TextBox2,
            $Label3,
            $Label4,
            $TextBox3,
            $Label5,
            $TextBox4,
            $TextBox5,
            $Label6,
            $TextBox6,
            $ButtonSubmit,
            $ButtonExit,
            $ButtonNewMail
    ))

    # Laat het form, bovenop alle andere windows zien.
    $Form.Add_Shown({$Form.Activate()})

    # Teken het form op het scherm. Void output naar console.
    [void] $Form.ShowDialog()

    if ($returnValue.Sendmail -eq 1){

        $script:returnEmail = Show-SendMail
        $Email = 1
    }

    if (!$AccountExist) {
        return $script:returnValue

    } elseif ($email -eq 1) {
        return $script:returnEmail
    }
}

# Gui voor het opstellen van een e-mail.
# Je kan hier de eigenschappen van de e-mail aanpassen.
function Show-SendMail {

    $Form                            = New-Object system.Windows.Forms.Form
    $Form.ClientSize                 = New-Object System.Drawing.Point(310,250)
    $Form.text                       = "Send Email"
    $Form.StartPosition              = "CenterScreen"
    $Form.FormBorderStyle            = "FixedDialog"
    $Form.TopMost                    = $true
    $Form.MaximizeBox                = $false
    $Form.MinimizeBox                = $true
    $Form.ControlBox                 = $false
    
    # ________________________________________________ Ontvanger ______________________________________________________

    $LabelFrom                      = New-Object system.Windows.Forms.Label
    $LabelFrom.text                 = "Van : "
    $LabelFrom.AutoSize             = $true
    $LabelFrom.width                = 25
    $LabelFrom.height               = 15
    $LabelFrom.location             = New-Object System.Drawing.Point(14,18)

    $TextBoxFrom                    = New-Object system.Windows.Forms.TextBox
    $TextBoxFrom.multiline          = $false
    $TextBoxFrom.text               = "BEHEER03Script@SomeCompany.nl"
    $TextBoxFrom.width              = 170
    $TextBoxFrom.height             = 15
    $TextBoxFrom.location           = New-Object System.Drawing.Point(125,16)

    # ________________________________________________ CC ______________________________________________________

    $LabelTo                        = New-Object system.Windows.Forms.Label
    $LabelTo.text                   = "Aan : "
    $LabelTo.AutoSize               = $true
    $LabelTo.width                  = 25
    $LabelTo.height                 = 15
    $LabelTo.location               = New-Object System.Drawing.Point(14,53)

    $TextBoxTo                      = New-Object system.Windows.Forms.TextBox
    $TextBoxTo.multiline            = $false
    $TextBoxTo.text                 = "ServiceDesk@SomeCompany.nl"
    $TextBoxTo.width                = 170
    $TextBoxTo.height               = 15
    $TextBoxTo.location             = New-Object System.Drawing.Point(125,51)

    # ________________________________________________ From ______________________________________________________

    $LabelCC                        = New-Object system.Windows.Forms.Label
    $LabelCC.text                   = "CC : "
    $LabelCC.AutoSize               = $true
    $LabelCC.width                  = 25
    $LabelCC.height                 = 15
    $LabelCC.location               = New-Object System.Drawing.Point(14,88)

    $TextBoxCC                      = New-Object system.Windows.Forms.TextBox
    $TextBoxCC.multiline            = $false
    $TextBoxCC.text                 = "TopdeskImport@SomeCompany.nl"
    $TextBoxCC.width                = 170
    $TextBoxCC.height               = 15
    $TextBoxCC.location             = New-Object System.Drawing.Point(125,90)
  
    # ________________________________________________ Text ______________________________________________________

    $LabelTxT                        = New-Object system.Windows.Forms.Label
    $LabelTxT.text                   = "Instructies:`nDruk op de 'New Letter' knop om de brief te maken`nNadat de browser is gestart kan je de brief printen.`nDruk Ctrl+P om naar pdf te printen.`nZet het printen van headers en footers uit!"
    $LabelTxT.AutoSize               = $true
    $LabelTxT.width                  = 25
    $LabelTxT.height                 = 15
    $LabelTxT.location               = New-Object System.Drawing.Point(14,123)

    # ________________________________________________ Send Email Button ______________________________________________________

    $ButtonSendEmail                = New-Object system.Windows.Forms.Button
    $ButtonSendEmail.text           = "New Letter"
    $ButtonSendEmail.width          = 100
    $ButtonSendEmail.height         = 30
    $ButtonSendEmail.location       = New-Object System.Drawing.Point(85,200)

    $ButtonSendEmail.Add_Click({

        $script:returnEmail = @{
            "continue" = 1
        }
        $Form.close()
    })

    # ________________________________________________ Button Exit______________________________________________________

    $ButtonExit                 = New-Object system.Windows.Forms.Button
    $ButtonExit.text            = "Exit"
    $ButtonExit.width           = 100
    $ButtonExit.height          = 30
    $ButtonExit.location        = New-Object System.Drawing.Point(195,200)

    $ButtonExit.Add_Click({
        $script:returnEmail = @{
            "Close" = 1
        }
        $Form.Close()
    })

    # ______________________________________________________________________________________________________

    $Form.controls.AddRange(@(
        $LabelFrom,
        $TextBoxFrom,
        $LabelTo,
        $TextBoxTo,
        $LabelCC,
        $TextBoxCC,
        $ButtonSendEmail,
        $ButtonExit,
        $LabelTxT
    ))

    $Form.Add_Shown({$Form.Activate()})
    [void] $Form.ShowDialog()

    if ($returnEmail.Close -eq 1) {
        Exit

    } elseif ($returnEmail.continue -eq 1) {

        $script:returnEmail = @{
            "From"          = $TextBoxFrom.text
            "To"            = $TextBoxTo.Text
            "CC"            = $TextBoxCC.text
            "Attachment"    = $Attachment.Text
        }
    }

    return $script:returnEmail
}

# Start de GUI en bewaar alle user input in variabele $returnedValues.
# User input wordt bewaard nadat de gui gesloten wordt.
$returnedValues = Show-GUI

# sluit het script af als button exit is ingedrukt.
If (($($returnedValues.Close)) -eq 1) {

    # Stop het script.
    exit
}

# ________________________________________________ Maak het account aan ______________________________________________________

$NewUser            = $returnedValues.NaamMedewerker
$ExampleUser        = $returnedValues.VoorbeeldAccount
$Date               = $returnedValues.EindDatum
$ChangeNumber       = $returnedValues.Topdesknummer
$Password           = $returnedValues.Wachtwoord
$Manager            = $returnedValues.Manager

$SecurePassword     = ConvertTo-SecureString -String $Password -AsPlainText -Force

$NewUserTrue        = ($NewUser.Length -gt 5)                                           # Check of de namen gelijk aan, of groter zijn dan vijf karakters.
$ExampleUserTrue    = ($ExampleUser.Length -gt 5)                                       # Ik kan geen namen bedenken met minder dan vijf karakters. Li Su?
$DateTrue           = ($Date -match '^\d{2}-\d{2}-\d{2}')                               # Datum is als volgd: yy-MM-dd. Er zit geen controlle op de volgorde van jaar, maand en dag        
$PasswordTrue       = ($Password -match '^(?=.*[A-Z])(?=.*\d)(?=.*[^\w\s]).{8,}$')      # Wachtwoord is 8 karakters of langer, het moet een hoofdletter en een cijfer bevatten.
$ChangeNumberTrue   = ($ChangeNumber -match '^(WA|W|M)\d{4}\s\d{3,4}$')                 # Wijzigingsnummer begint met WA, W of M en kan tussen de 7 en 8 karakters lang zijn.

# Als alle 5 waardes in de gui goed zijn ingevuld. Maak dan het account van de nieuwe medewerker aan.
If (($NewUser -and ($NewUserTrue -eq 'True')) -and ($ExampleUser -and ($ExampleUserTrue -eq 'True')) -and ($Date -and ($DateTrue -eq 'True')) -and ($Password -and ($PasswordTrue -eq 'True')) -and ($ChangeNumber -and ($ChangeNumberTrue -eq 'True'))) {

    cls
    Write-Host "De volgende gegevens zullen gebruikt worden om een nieuw account te maken.`n__________________________________________________________________________`nNieuwe medewerker   : $NewUser`nExample User        : $ExampleUser`nVerloopdatum        : $Date`nWachtwoord          : $Password`nChange Number       : $ChangeNumber"

    # Maak variabele aan voor de voornaam, volledige achternaam, achternaam zonder toevoegsels, upn en login naam.
    $Name           = $NewUser.Substring(0, $NewUser.IndexOf(" "))
    $FullSurname    = $NewUser.Substring($NewUser.IndexOf(" ")+1)
    $Surname        = $FullSurname.split()[-1]
    $Upn            = $Name[0] + $FullSurname.replace(' ', '') + $Domain
    $Account        = (($Name[0] + $Name[1] + $Surname[0] + $Surname[1]).tolower() + "*")

    if ($AccountAlreadyExist = Get-ADUser -Filter {Name -like $NewUser}) {
        cls
        write-host "Er bestaat al een account voor gebruiker $NewUser. Je moet het account weer aanzetten en handmatig instellen."
        $AccountAlreadyExist
        exit 
    }
 
    # Is er al een UserPrincipalName met deze naam is?
    # En als dat zo is, dan heeft de variabele $UpnExist een waarde.
    $UpnExist = (Get-ADUser -filter {UserPrincipalName -like $Upn}).UserPrincipalName

    # Als het e-mail adress van de nieuwe user al bestaat. Dus er was al een Jan Jansen (jaja)en nu wordt er een Jaap Jansen (jaja) gemaakt.
    # Plak hier dan een 2 of een opvolgnummer achter de naam van het account.
    If ($UpnExist) {

        # Stop de naam (eerste letter voornaam plus achternaam) en het opvolgnummer in apparte varabele
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

    # Haal de Sam naam op van het voorbeeldaccount.
    $ExampleLSam = (get-aduser -filter {name -like $ExampleUser}).SamAccountName

    # Vraag de properties op van het voorbeeldaccount.
    $UserProperties = Get-AdUser -Identity $ExampleLSam -Properties Title, Department, Company
    $Title          = $UserProperties.Title
    $Department     = $UserProperties.Department
    $Company        = $UserProperties.Company

    # In welk OU moet het account worden aangemaakt.
    $OrganizationalUnit = ((Get-ADUser $ExampleLSam | select DistinguishedName).DistinguishedName).Substring(((Get-ADUser $ExampleLSam | select DistinguishedName).DistinguishedName).IndexOf(',')+1)

    if ($OrganizationalUnit.Contains("Disabled Accounts")) {

        Write-Host $OrganizationalUnit
        Write-Host ''
        Write-Host "Deze persoon $ExampleUser is inmiddels uit dienst."
        Write-Host 'Geef een ander voorbeeld account op.'
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

    # Stop de waardes in aparte variabele omdat de omschrijving op een ander manier wordt ingevuld dan de verloopdatum.
    $Year =  $Date.Substring(0, 2)
    $Mounth = $Date.Substring(3, 2)
    $Day = $Date.Substring(6, 2)

    # Maak de variabelen voor de verloop datum. Tel er daarnaa 1 dag bij op.
    # $ExpirationDate = $Day + '-' + $Mounth + '-' + $Year
    $ExpirationDate = "$Day-$Mounth-$Year"
    $ExpirationDate = (Get-Date (([datetime]::ParseExact($ExpirationDate, 'dd-MM-yy', $null)).AddDays(1)) -format "dd-MM-yy")

    # De dag en tijd waarop het account verloopt.
    $AccountExpirationDate = $ExpirationDate + ' ' + $ExpirationTime

    # De omschrijving van het account, is de verloopdatum plus 1 dag. En het wijzigingsnummer.
    $Description = (Get-Date (([datetime]::ParseExact($Date, 'yy-MM-dd', $null)).AddDays(1)) -format "yy-MM-dd")  + ' ' + $ChangeNumber.ToUpper()

    # Haal de DistinguishedName van de manager op, op basis van voor en achternaam
    $ManagerDN = Get-ADUser -Filter {Name -like $Manager} -Properties DistinguishedName | Select-Object DistinguishedName

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
        -Department $Department `
        -Company $Company `
        -AccountPassword $securePassword `
        -ChangePasswordAtLogon $True `
        -Enabled $True `
        -Description $Description `
        -AccountExpirationDate $AccountExpirationDate `
        -HomeDirectory $ProfilePath `
        -HomeDrive $HomeDrive `
        -Path $OrganizationalUnit `
        -Manager $ManagerDN

    # Voeg de groepslidmaatschappen toe aan het account.
    Foreach ($group in $ExampleUserGroups) {
        Add-ADGroupMember -Identity $group -Members $SamAccount
    }

    # De nieuwe map is ook het pad naar naar het profiel.
    $NewFolder = $ProfilePath

    # Maak een nieuwe map aan en maak de user eigenaar van de home folder en geef fullcontol permission.
    Set-FolderPermission $SamAccount $ProfilePath $NewFolder

    # Maakt de Exchange online mailbox aan in de productieomgeving.
    # Deze stap wordt overgeslagen als er een account in de testomgeving wordt aangemaakt.
    if ($TestOmgeving -eq 0) {

        Write-Host ''
        Write-Host 'Momentje geduld, de Exchange Management Shell snap-in wordt nu geladen.'

        # Add-PSsnapin Microsoft.Exchange.Management.PowerShell.E2010
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

        # Maak de mailbox aan voor de nieuwe gebruiker.
        Enable-RemoteMailbox -Identity $SamAccount -RemoteRoutingAddress "$SamAccount@SomeCompanyweb.mail.onmicrosoft.com" > $null

        # Stel de taal in op Nederlands.
        Set-ADUser -Identity "$SamAccount" -Replace @{preferredLanguage="nl-NL"}

        # Ververs de federation server zodat SSO goed werkt.
        Invoke-Command -ComputerName $FederationServer -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}

        write-host "`nMomentje geduld, wij moeten hier wachten totdat de mailbox is aangemaakt!`nDit kan even duren......"

        # Wachten totdat de mailbox is aangemaakt in exchange online.
        while (-not $mailbox) {
            Start-Sleep -Seconds 30
            $mailbox = Get-Mailbox -Identity "$upn" -ErrorAction SilentlyContinue
        }

        # Stelt de default font in voor de EXO mailbox
        Set-MailboxMessageConfiguration -Identity "$upn" -DefaultFontName "Arial" -DefaultFontSize "10"

        # Stel het delen van de agenda/calender in.
        # Add-MailboxFolderPermission -Identity "$upn`:\agenda" -User All_Users_For_Calendar_Sharing -AccessRights reviewer
        Add-MailboxFolderPermission -Identity "$upn`:\calendar" -User All_Users_For_Calendar_Sharing -AccessRights reviewer
    }
    
    Show-GUI ($AccountExist = 1) $NewUser $ExampleUser $Date $Password $ChangeNumber > $null


} Else {

    cls    
    Write-Host "Je hebt niet alle velden goed ingevuld.`nExit script.`n`n"
    Exit
}

# De welkom brief.
$htmlContent = @"
<!DOCTYPE html>
<html lang="nl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Welkoms Brief</title>
    <style>
        @page {
            size: A4;
            margin: 15mm;
        }
        body {
            font-family: Arial, sans-serif;
            font-size: 15px;
            margin: 25;
            padding: 25;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 95vh;
            transform: scale(1.0);
        }
        .container {
            width: 100%;
            max-width: 210mm;
            height: 100%;
            max-height: 297mm;
            padding: 25px;
        }

        .logo {
            margin-left: 425px;
            margin-top: -10px;
            top: 1px;
            right: 30px;
            width: 250px;
        }
        .section {
            margin-bottom: 1em;
        }
        table {
            width: 55%;
            border-collapse: collapse;
        }
        table td {
            padding: 1px;
            border: 1px solid #ffffff;
        }
        p {
            width: 95%;
            margin-bottom: 20px;

        }
        table td:first-child {
            width: 45%;
        }
    </style>
</head>

<body>
    <div class="container">

        <img class="logo" src="SomeCompany.png" alt="Logo"> 
        <br><br><br><br><br><br><strong>Strikt persoonlijk voor:</strong> $NewUser</p><br>

        <div class="section">
            <p>$NewUser,</p>
            <p>Dit is jouw (nieuwe) inlog voor de omgeving van SomeCompany:</p>
        </div>

        <table>
            <tr>
                <td><strong>Je e-mailadres</strong></td>
                <td>$Upn</td>
            </tr>
            <tr>
                <td><strong>Login naam</strong></td>
                <td>$SamAccount</td>
            </tr>
            <tr>
                <td><strong>Wachtwoord</strong></td>
                <td>$Password</td>
            </tr>
        </table>
        <div class="section">
            <p>Als jij je voor de eerste keer aanmeldt!</p>
            <p>Wijzig dan meteen het wachtwoord. Een nieuw wachtwoord moet aan 3 eisen voldoen:</p>
            <ul>
                <li>het wachtwoord moet uit kleine letters, hoofdletters en cijfers bestaan;</li>
                <li>minimaal 8 posities lang zijn;</li>
                <li>en het wachtwoord mag ook niet op je naam of loginnaam lijken.</li>
            </ul>
        </div>
        <div class="section">
            <p>Houd je wachtwoord altijd geheim. Schrijf het niet op in een agenda of op geeltjes onder je toetsenbord. Want het is niet de bedoeling dat anderen met jouw code in het systeem kunnen.</p>
        </div>
        <div class="section">
            <p>Verlaat je jouw werkplek? Vergrendel dan je laptop. Dit doe je door gelijktijdig de WIN en L toetst in te drukken.</p>
            <p>Ga je aan het einde van de werkdag naar huis? Of verwacht je dezelfde dag niet meer terug te keren? Meld je dan altijd af van het systeem en schakel je laptop of werkstation uit. Zo blijven jouw gegevens veilig op het netwerk staan bij de dagelijkse back-up.</p>
        </div>
        <div class="section">
            <p>Heb je een vraag, opmerking of een probleem? Bel de Servicedesk via 020-5118329. Deze is geopend tussen 8:00 uur en 17:30 uur. Je kunt ook een e-mail sturen naar <a href="mailto:servicedesk@SomeCompany.nl">servicedesk@SomeCompany.nl</a> of stel je vraag aan een van je collega's.</p>
        </div>
        <div class="section">
            <p><br><br>Welkom bij SomeCompany!</p>
            <p><br>Met vriendelijke groet,<br>Systeembeheer</p>
        </div>
    </div>
</body>
</html>
"@

# Bewaar het pad naar de brief in variabele $letter.
# Bewaar de link naar het HTML bestand in variabele $Letter_URL.
# Schrijft de HTML brief naar een bestand.
$Letter         = "$PsScriptRoot\$ChangeNumber $NewUser.html"
$Letter_URL     = "file:///$Letter" -replace ' ', '%20'
$htmlContent    | Out-File -FilePath $Letter -Encoding utf8

# Start browser met het url naar de brief.
$Process = Start-Browser $Letter_URL

# Geef het script 1.5 seconde de tijd.
Start-Sleep -Milliseconds 1500

# Verzamel de namen van alle gevonden .pdf bestanden. De .pdf bestanden MOETEN in de zelfde map opgeslagen worden als waar dit script in uitgevoerd wordt.
$pathToletter = Join-Path -Path $PSScriptRoot -ChildPath $($extension = "*.pdf")

# Bewaar het volledige pad naar ieder gevonden .pdf bestand in een array.
foreach ($file in Get-ChildItem -Path $PSScriptRoot -Filter $extension -File) {
    $pdfFiles += $file.FullName
}

# Zijn er pdf bestanden in het opgegeven pad?
if (Test-Path $pathToletter) {

    # Als er meer dan 1 pdf bestand gevonden is.
    if ($PdfFiles.Length -gt 1) {

        # $nr = $PdfFiles.Length

        Write-Host "`nEr zijn $($PdfFiles.Length) pdf bestanden gevonden!`n"

        $i = 0

        # Laat het aantal gevonden bestanden aan de gebruiker zien.
        while ($PdfFiles.Length -gt $i) {
            
            $Number = ($i + 1)
            Write-Host "`t$Number`:" $PdfFiles[$i]
            $i++

            if ($PdfFiles.Length -eq $i) {
                break
            }
        }
        
        $i = 0
        
        # Welk bestand moet er gebruikt worden?
        $pdf = Read-Host "`nGeef het nummer op van het bestand dat je wilt gebruiken`n`nNummer"
        # Het nummer -1 want de array $pdfFiles begint te tellen vanaf 0.
        $pdf = ($pdf - 1)

    } elseif ($PdfFiles.Length -eq 1) {
        
        Write-Host "`nEr is 1 pdf bestand gevonden!`n"
        # Gebruik het eerste bestand in array $pdfFiles.
        $pdf = 0
        Write-Host $PdfFiles[$pdf]
    }

    # Bewaar de naam van de brief in een variabele.
    $letter_PDF = $pdfFiles[$pdf]

} else {

    Write-Host "Er zijn geen PDF bestanden gevonden in:`n$PSScriptRoot`n"
    # Sluit het script af om opnieuw te beginnen.
    Exit
}

# Variabelen voor de naamgeving van de brief.
$User_letter    = (" $NewUser") -replace ' ', '_'
$Date_Letter    = Get-date -Format yyyyMMdd
$New_Letter     = "$PSScriptRoot\$Date_Letter$User_letter.pdf"

# Hernoem het pdf bestand, naar een naam met het volgende formaat: yyyyMMdd_Voornaam_volledige_Acheternaam.pdf.
Rename-Item -Path $letter_PDF -NewName $New_Letter

# De waardes die gebruikt worden voor de opmaak van de e-mail.
$SendFrom       = $returnEmail.From
$SendTo         = $returnEmail.To
$SendCC         = $returnEmail.CC
$MailSubject    = "$ChangeNumber $NewUser"
$MailBody       = [string]$htmlContent
$MailServer     = "smtp.SomeCompany.nl"

# De e-mail om te verzenden.
$Email = @{
    From        = $SendFrom
    To          = $SendTo
    CC          = $SendCC
    Subject     = $MailSubject 
    Body        = $MailBody
    BodyAsHtml  = $true
    SmtpServer  = $MailServer
    Attachments = $New_Letter
}

# Stuur de e-mail!
Send-MailMessage @Email

# Geef het script 1.5 seconde de tijd.
Start-Sleep -Milliseconds 1500

# Verplaats de brief naar de share waar alle brieven worden bewaard.
Move-Item -path $New_Letter -Destination $pathToletters
