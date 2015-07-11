param($client,$domain)

add-pssnapin quest.activeroles.admanagement
#add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010
Import-Module C:\Users\Administrator\AppData\Roaming\Microsoft\Exchange\RemotePowerShell\mail2.home.local\mail2.home.local.psm1

$CN = "Home.local/Hosting/$client/$domain"
$DN = "OU=$domain,OU=$client,OU=HOSTING,DC=HOME,DC=LOCAL"
$GROUPDN = "CN=$domain"+"_Exchange_Group,"+$DN
$SMTPTEMPLATE = "smtp:%m@"+$domain
$MBSERVER = (get-mailboxserver).name
$CASERVER = (get-clientaccessserver).name
$DGROUPNAME = $domain + "_Exchange_Group"
$AL = "\$domain"
$VD = $CASERVER+"\OAB (Default Web Site)"
$EXCHANGESERVICES = "CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=Home,DC=Local"
$ALC = "CN=Address Lists Container,CN=Home,$EXCHANGESERVICES"


New-QADObject -ParentContainer "OU=$client,OU=HOSTING,DC=HOME,DC=LOCAL" -name $domain -Type organizationalunit -Description "$domain Exchange OU"

Set-QADObject "OU=$domain,OU=$client,OU=HOSTING,DC=HOME,DC=LOCAL" -ObjectAttributes @{uPNSuffixes=$domain}

New-DistributionGroup -Name $DGROUPNAME -Type 'Security' -OrganizationalUnit $CN -SamAccountName $DGROUPNAME -Alias $DGROUPNAME

New-AcceptedDomain -Name $domain -DomainName $domain -DomainType 'Authoritative'

New-emailaddresspolicy $domain -RecipientFilter {((Memberofgroup -eq $GROUPDN) -and (recipienttype -eq 'usermailbox'))} -enabledprimarySMTPaddresstemplate $SMTPTEMPLATE

New-addresslist -name $domain -container "\" -RecipientFilter {((Memberofgroup -eq $GROUPDN) -and (recipienttype -eq 'usermailbox'))}

New-offlineaddressbook -name $domain -addresslists $AL -virtualdirectories $VD
#New-offlineaddressbook -name $domain -server $MBSERVER -addresslists $AL -publicfolderdistributionenabled $true -virtualdirectories $VD
#New-offlineaddressbook -name $domain -server $MBSERVER -addresslists $AL -publicfolderdistributionenabled $false -virtualdirectories $VD

New-globaladdresslist -name $domain -RecipientFilter {((Memberofgroup -eq $GROUPDN) -and (recipienttype -eq 'usermailbox'))}


$permissions = Get-QADPermission -schema -inherited "CN=$domain,CN=All Address Lists,$ALC"

$permissions | add-qadpermission "CN=$domain,CN=All Address Lists,$ALC"

Set-QADObjectSecurity -Lockinheritance -Remove "CN=$domain,CN=All Address Lists,$ALC"

Get-QADPermission "CN=$domain,CN=All Address Lists,$ALC" -schema |where {$_.accountname -eq "NT AUTHORITY\Authenticated Users"}|Remove-QADPermission


$permissions = Get-QADPermission -schema -inherited "CN=$domain,CN=All Global Address Lists,$ALC"

$permissions | add-qadpermission "CN=$domain,CN=All Global Address Lists,$ALC"

Set-QADObjectSecurity -Lockinheritance -Remove "CN=$domain,CN=All Global Address Lists,$ALC"

Get-QADPermission "CN=$domain,CN=All Global Address Lists,$ALC" -schema |where {$_.accountname -eq "NT AUTHORITY\Authenticated Users"}|Remove-QADPermission


$permissions = Get-QADPermission -schema -inherited "CN=$domain,CN=Offline Address Lists,$ALC"

$permissions | add-qadpermission "CN=$domain,CN=Offline Address Lists,$ALC"

Set-QADObjectSecurity -Lockinheritance -Remove "CN=$domain,CN=Offline Address Lists,$ALC"

Get-QADPermission "CN=$domain,CN=Offline Address Lists,$ALC" -schema |where {$_.accountname -eq "NT AUTHORITY\Authenticated Users"}|Remove-QADPermission


$permissions = Get-QADPermission -schema -inherited $DN

$permissions | add-qadpermission $DN

Set-QADObjectSecurity -Lockinheritance -Remove $DN

Get-QADPermission $DN -schema |where {$_.accountname -eq "NT AUTHORITY\Authenticated Users"}|Remove-QADPermission


$container = "CN=$domain,CN=All Address Lists,$ALC"

Add-ADPermission $container -User $DGROUPNAME -AccessRights GenericRead, ListChildren -ExtendedRights Open-Address-Book


$container = "CN=$domain,CN=All Global Address Lists,$ALC"

Add-ADPermission $container -User $DGROUPNAME -AccessRights GenericRead, ListChildren -ExtendedRights Open-Address-Book


$container = "CN=$domain,CN=Offline Address Lists,$ALC"

Add-ADPermission $container -User $DGROUPNAME -AccessRights GenericRead, ListChildren -ExtendedRights ms-Exch-Download-OAB

Set-QADObject $EXCHANGESERVICES -ObjectAttributes @{addressBookRoots=@{append=@("CN=$domain,CN=All Address Lists,$ALC")}}