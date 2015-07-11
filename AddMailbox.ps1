param($client,$domain,$user,$password,[INT]$quota,[INT]$retention,[switch]$OWA,[switch]$ActiveSync,[switch]$POP,[switch]$IMAP)

add-pssnapin quest.activeroles.admanagement
#add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010
Import-Module C:\Users\Administrator\AppData\Roaming\Microsoft\Exchange\RemotePowerShell\mail2.home.local\mail2.home.local.psm1

$AL = "\$domain"
$CN = "Home.local/Hosting/$client/$domain"
$UPN = $user+"@"+$domain
$USERCN = $CN+"/"+$UPN
$CLIENTDN = "OU=$domain,OU=$client,OU=HOSTING,DC=HOME,DC=LOCAL"
$USERDN = "CN="+$upn+","+$CLIENTDN
$GROUPDN = "CN=$domain" + "_Exchange_Group"+","+$CLIENTDN
$MBSERVER = (get-mailboxserver).name
$database = (Get-MailboxDatabase).servername+"\"+(get-mailboxdatabase).storagegroupname+"\"+(get-mailboxdatabase).name

$SAM = $domain+$user
$quotainMB = "$quota"+"MB"
$warning = $quota * .05
$warninginMB = "$warning"+"MB"


if ($SAM.length -gt 14)
{
$SAM = $sam.Remove(0,($sam.length -15))
}

if ($OWA){[Boolean]$owa = $true}else{[Boolean]$owa = $false}
if ($ActiveSync){[Boolean]$ActiveSync = $true}else{[Boolean]$ActiveSync = $false}
if ($POP){[Boolean]$POP = $true}else{[Boolean]$POP = $false}
if ($OWA){[Boolean]$IMAP = $true}else{[Boolean]$IMAP = $false}


New-QADUser -Name $upn -UserPrincipalName $upn -ParentContainer $CLIENTDN -UserPassword $password -SamAccountName $sam -DisplayName $user -ObjectAttributes @{mail=$upn}

add-QadGroupMember -identity $GROUPDN -member $USERDN

Enable-mailbox -identity $USERCN -alias $user -Database $database

Update-emailaddresspolicy -identity $domain
Update-addresslist -identity $AL
Update-globaladdresslist -identity $domain

Get-User -Filter { userPrincipalName -eq $UPN } | Set-Mailbox -OfflineAddressBook $domain

Set-QADObject $USERDN -ObjectAttributes @{msExchQueryBaseDN=$CLIENTDN}

Get-Mailbox|where {$_.name -eq $UPN}|set-mailbox -UseDatabaseQuotaDefaults $false
Get-Mailbox|where {$_.name -eq $UPN}|set-mailbox -UseDatabaseRetentionDefaults $false
Get-Mailbox|where {$_.name -eq $UPN}|set-mailbox -Prohibitsendquota $quotainMB
Get-Mailbox|where {$_.name -eq $UPN}|set-mailbox -Prohibitsendreceivequota $quotainMB
Get-Mailbox|where {$_.name -eq $UPN}|set-mailbox -issuewarningquota $warninginMB
Get-Mailbox|where {$_.name -eq $UPN}|Set-Mailbox -RetainDeletedItemsFor $retention
Set-CASMailbox -Identity $UPN -OWAEnabled $OWA
Set-CASMailbox -Identity $UPN -Activesyncenabled $ActiveSync
Set-CASMailbox -Identity $UPN -Popenabled $POP
Set-CASMailbox -Identity $UPN -Imapenabled $IMAP
