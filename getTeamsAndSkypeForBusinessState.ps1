Import-Module SkypeOnlineConnector

$strSecureString = "<SecureStringForAdminAccountUsedBelow>"
$strPassword = $strSecureString | ConvertTo-SecureString
$cred = new-object -typename PSCredential -argumentlist 'AdminAccount@example.onmicrosoft.com',$strPassword

$sfbSession = New-CsOnlineSession -Credential $cred
Import-PSSession $sfbSession
$strFile = ".\S4BUserInfo.txt"
remove-item $strFile

$objAllUsers = get-csonlineuser
ForEach ($objUser in $objAllUsers){
    $strLine = $objUser.userPrincipalName + "`t" + $objUser.InterpretedUserType + "`t" + $objUser.HostingProvider + "`t" + $objUser.TeamsUpgradeEffectiveMode
    write-output $strLine | out-file ".\S4BUserInfo.txt" -Append
    write-output $strLine 
}
