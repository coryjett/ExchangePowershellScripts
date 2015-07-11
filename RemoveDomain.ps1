param($client=".",$domain=".")

add-pssnapin quest.activeroles.admanagement

$CN = "Home.com/Hosting/$client"
$DN = "OU=$client,OU=HOSTING,DC=HOME,DC=COM"
$GROUPDN = "CN=$domain,OU=$client,OU=HOSTING,DC=HOME,DC=COM"
$SMTPTEMPLATE = "smtp:%m@"+$domain
$MBSERVER = (get-mailboxserver).name
$CASERVER = (get-clientaccessserver).name
$AL = "\"+$domain
$VD = $CASERVER+"\OAB (Default Web Site)"


Set-QADObject "CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=Home,DC=Com" -ObjectAttributes @{addressBookRoots=@{delete=@("CN=$domain,CN=All Address Lists,CN=Address Lists Container,CN=Home,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=home,DC=com")}}