<#
    Accounts Opruimen V1.13

    Dit script ruimt de accounts en home folders op van medewerkers die een maand of langer uit dienst zijn.
    Het doet dit door de datum uitdienst van een account op te halen uit het omschijving veld in het account.
    Vul het daarom altijd goed in!

    Het maakt de 'Domain admins' groep eigenaar van de home folder en gooit hem daarna weg.
    Alle groepslidmaatschappen van de gebruiker worden verwijderd.
    Bij het verwijderen wordt ook de Office365 licentie verwijderd, daarmee wordt ook de mailbox van de bebruiker weggegooid.
    Vervolgens wordt het account uit de tussenfase OU gehaald en in de _Disabled Accounts OU geplaatst van het huidige jaar.
    
    Er wordt een e-mail verzonden naar ictinfo@stadgenoot.nl met alle aanpassingen.

    Geschreven door: Jos Severijnse™.
#>

# Globale Variabelen.
$AdminGroup = "DOMAIN\'Domain admins"
$OuWaitOneMonth = "OU=OneMothPause, OU=Disabled Accounts,OU=Business,DC=Domain,DC=SubDomain"
$OuThisYear = "OU=Disabled Accounts,OU=Disabled Accounts,OU=Business,DC=Domain,DC=SubDomain"
$MessageBody = @("Disabeled accounts and delteted home folders`n___________________________________________________`n`n")
$OneMonthAgo = Get-Date (([datetime]::ParseExact((Get-Date -Format "yy-MM-dd"), 'yy-MM-dd', $null)).AddDays(-30)) -format "yy-MM-dd"

# Gebruikers ophalen uit de OU en loop door alle gevonden accounts.
Get-ADUser -SearchBase $OuWaitOneMonth  -filter * -Properties HomeDirectory, mailNickName, Description | % {

    # Variabelen met betrekking tot het huidige account in de loop.
    $Homedir = $_.HomeDirectory
    $SamAccountname = $_.SamAccountname
    $Mailnickname = $_.MailNickName
    $DateEndOfContract = $_.Description.Substring(0, 8)

    # Is de gebruiker al een maand uit dienst?
    if ($DateEndOfContract -lt $OneMonthAgo) {

        # Testen of de home folder nog bestaat.
        if(Test-Path $homedir) {

            # Maakt de 'Domain admins' groep eigenaar van de home folder.
            $Acl = Get-Acl $Homedir.FullName
            $Acl.SetOwner([System.Security.Principal.NTAccount]"$AdminGroup")
            Set-Acl $Homedir.FullName $Acl -Verbose

            # Vraag alle mappen op in de home folder van de gebruiker.
            $Folders = Get-ChildItem $Homedir -Directory -Recurse

            # loop door alle mappen en maak de 'Domain admins' groep eigenaar.
            Foreach($Folder in $Folders) {

                # Maakt de 'Domain admins' groep eigenaar van alle folders in de home folder.
                $Acl = Get-Acl $Folder.FullName
                $Acl.SetOwner([System.Security.Principal.NTAccount]"$AdminGroup")
                Set-Acl $Folder.FullName $Acl -Verbose
            }

           # Vraag alle bestanden op in de home folder van de gebruiker.
            $Files = Get-ChildItem $Homedir -File -Recurse

            # loop door alle bestanden en maak de 'Domain admins' groep eigenaar.
            Foreach($File in $Files) {

                # Maakt de 'Domain admins' groep eigenaar van alle files in de home folder.
                $Acl = Get-Acl $File.FullName
                $Acl.SetOwner([System.Security.Principal.NTAccount]"$AdminGroup")
                Set-Acl $File.FullName $Acl -Verbose
            }

            # Verwijder de home folder van de user.
            Remove-Item $Homedir -Recurse -Force

        } else {

            # Als de home folder al verwijderd was, laat dit dan zien in de log.
            $Homedir = "De home folder van deze gebruiker is niet gevonden!"
        }

        # Haal alle namen van de groepslidmaatschappen op.
        $GroupMembership = (Get-ADPrincipalGroupMembership $SamAccountname | select name).name

        # Loop door alle gevonden groepen in het account.
        Foreach ($group in $GroupMembership) {
            
            # Verwijder alle groepen van het account, maar laat de 'Domain users' groep staan.
            # De 'Domain users' groep mag niet worden verijderd. De rest wel.
            if ($group -ne "'Domain users'") {

                # Vraag de groep op en verwijder de user. Geeft een error als de groep niet meer bestaat.
                Remove-ADGroupMember -Identity $group -Members $SamAccountname -Confirm:$false
            }
        }

        # Haal de user uit tussenfase en verplaats hem naar de OU voor accounts van users uit dienst.
        get-aduser -identity $SamAccountname | Move-ADObject -TargetPath $OuThisYear 

        # Hou bij welke accounts en home folders er zijn opgeschoond.
        $MessageBody += "$SamAccountname`n$Homedir`n`n"
    }
}

# e-mail opmaak naar Support@SomeDomain.org met een lijst van de wijzigingen.
$Email = @{
    From = "SomeServer@SomeDomain.org"
    To = "Support@SomeDomain.org"
    SmtpServer = "smtp.SomeDomain.org"
    Subject = "Disabled Accounts and removed home folders."
    Body = [string]$MessageBody
}

# Stuur de e-mail!
Send-MailMessage @Email

# Leeg het scherm en laat alle aanpassingen zien.
cls
$MessageBody
