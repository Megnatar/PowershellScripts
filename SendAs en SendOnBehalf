<#
    Voer dit script niet uit!
    
    Het veranderd de manier hoe alle shared mailboxen omgaan met de Send Items folder in Outlook.
    Bij gebruik van SendAs en SendOnBehalf rechten, wordt er een kopie van de e-mail in de
    Send Items map van de shared mailbox geplaatst.

    Jammer dat er geen MessageMoveForSent.... parameter is.

    By: Jos Severijnse

#>

# Array met alle modules die al geladen zijn.
$AlreadyImportedModules = Get-Module

# Array met de modules om te laden.
$ModulesToLoad = @(“ExchangeOnlineManagement”)

# Loopt door alle te laden modules in de array, en als de module niet geladen is dan wordt deze geladen.
ForEach($module in $ModulesToLoad) {
    If($AlreadyImportedModules.Name -notcontains $module) {
        Import-Module $module
    }
}

# Variabele met het pad naar waar alle accounts in AD staan voor de shared mailboxen.
$OuPath = "OU=Groepsmailboxen,OU=Organisatie,DC=connect,DC=local"

# Array met alle namen van de mailboxen.
$a = Get-ADUser -Filter * -SearchBase $OUpath | Select-object Name

# Loopt door alle elementen in de array en past de settings aan voor alle aanwezige shared mailboxen.
for ($i=0; $i -lt $a.Length; $i++) {
    Get-Mailbox $a[$i].Name | Set-Mailbox -MessageCopyForSentAsEnabled $true -MessageCopyForSendOnBehalfEnabled $true  
}
