<#
    Accounts Opruimen V1.2


    Dit script ruimt de accounts en home folders op van medewerkers die een maand of langer uit dienst zijn.
    Het Script wordt automatisch 1 keer per week uitgevoerd up de SomeServer via een schedule task.

    Het maakt de 'Domain admins' groep eigenaar van de home folder en gooit hem daarna weg, als deze bestaat.
    Alle groepslidmaatschappen van de gebruiker worden verwijderd door het script, behalve de Domain users groep.
    Bij het verwijderen wordt ook de Office365 licentie verwijderd, daarmee wordt ook de mailbox van de gebruiker weggegooid.
    Vervolgens wordt het account uit de MaandWachten OU gehaald en in de _Disabled Accounts OU geplaatst, van het huidige jaar.
    
    Er wordt als laatst een e-mail verzonden naar SomeEmail@SomeDomain.org met alle aanpassingen.
    Als er geen aanpassingen zijn, dan wordt er geen mail verzonden.

    Geschreven door: Jos Severijnse™.
#>

# Globale Variabelen.
$AdminGroup = "CONNECT\'Domain admins"
$OuMaandWachten = "OU=OneMothPause, OU=Disabled Accounts,OU=Business,DC=Domain,DC=local"
$Year = 2024
$OuThisYear = "OU=_Disabled Accounts $Year,OU=Disabled Accounts,OU=Business,DC=Domain,DC=local"
$OneMonthAgo = Get-Date (([datetime]::ParseExact((Get-Date -Format "yy-MM-dd"), 'yy-MM-dd', $null)).AddDays(-30)) -format "yy-MM-dd"
$MessageBodySend = $MessageBody = @("De volgende accounts en home folders zijn opgeruimt:`n___________________________________________________`n`n")

# Gebruikers ophalen uit de OU, en 'loop' door alle gevonden accounts. $mailNickName kan er uitgehaald worden.
# Maar omdat een bepaald sleutelwoord meerdere properties kan ophalen, moet dat nog even getest worden.
Get-ADUser -SearchBase $OuTussenfase  -filter * -Properties HomeDirectory, mailNickName, Description | % {

    # Variabelen met betrekking tot het huidige account in de 'loop'.
    $Homedir = $_.HomeDirectory
    $SamAccountname = $_.SamAccountname
    $DateEndOfContract = $_.Description.Substring(0, 8)
    # $Mailnickname = $_.MailNickName

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

            # 'loop' door alle mappen en maak de 'Domain admins' groep eigenaar.
            Foreach($Folder in $Folders) {

                # Maakt de 'Domain admins' groep eigenaar van de huidige map in de 'loop'.
                $Acl = Get-Acl $Folder.FullName
                $Acl.SetOwner([System.Security.Principal.NTAccount]"$AdminGroup")
                Set-Acl $Folder.FullName $Acl -Verbose
            }

           # Vraag alle bestanden op in de home folder van de gebruiker.
            $Files = Get-ChildItem $Homedir -File -Recurse

            # 'loop' door alle bestanden en maak de 'Domain admins' groep eigenaar.
            Foreach($File in $Files) {

                # Maakt de 'Domain admins' groep eigenaar van het huidige bestand in de 'loop'.
                $Acl = Get-Acl $File.FullName
                $Acl.SetOwner([System.Security.Principal.NTAccount]"$AdminGroup")
                Set-Acl $File.FullName $Acl -Verbose
            }
            # Nu is de domein admins groep eigenaar en kunnen wij de home folder Verwijder.
            Remove-Item $Homedir -Recurse -Force

        } else {

            # Als de home folder al verwijderd was, laat dit dan zien in de log.
            $Homedir = "De home folder van deze gebruiker is niet gevonden!"
        }
        # Haal alle namen van de groepslidmaatschappen op.
        $GroupMembership = (Get-ADPrincipalGroupMembership $SamAccountname | select name).name

        # 'loop' door alle gevonden groepen in het account.
        Foreach ($group in $GroupMembership) {
            
            # Verwijder alle groepen van het account, maar laat de 'Domain users' groep staan.
            # De 'Domain users' groep mag niet worden verijderd. De rest wel.
            if ($group -ne "'Domain users'") {

                # Vraag de groep op en verwijder de user. Geeft een error als de groep niet meer bestaat.
                Remove-ADGroupMember -Identity $group -Members $SamAccountname -Confirm:$false
            }
        }
        # Haal de user uit MaandWachten en verplaats hem naar de OU voor accounts van users uit dienst.
        get-aduser -identity $SamAccountname | Move-ADObject -TargetPath $OuThisYear 

        # Hou bij welke accounts en home folders er worden opgeschoond.
        $MessageBodySend += "`n___________________________________________________`n`nAccount $SamAccountname is verplaatst naar:`n$OuThisYear`n`nDe volgende Home folder is verwijderd:`n$Homedir`n"
    }
}

# Stuur de e-mail alleen als er wijzigen zijn!
If ([string]$MessageBody -ne [string]$MessageBodySend) {

    # e-mail opmaak naar SomeEmail@SomeDomain.org met een lijst van de wijzigingen.
    $Email = @{
        From = "SomeServer@SomeDomain.org"
        To = "SomeEmail@SomeDomain.org"
        SmtpServer = "smtp.someDomain.org"
        Subject = "Accounts die opgeruimd zijn"
        Body = [string]$MessageBodySend
    }
    # Stuur de e-mail!. This is not secure and is decrepit.
    Send-MailMessage @Email

    # Leeg het scherm en laat alle aanpassingen zien.
    cls
    $MessageBody

} else {

    # Leeg het scherm en laat bericht zien.
    cls
    "Er zijn geen wijzigen gedaan!"
}
